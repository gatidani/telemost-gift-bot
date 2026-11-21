import importlib.util
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from bot import API, APIError, Bot, State, certificate_row, read_registry


class FakeAPI:
    def __init__(self):
        self.calls = []
        self.fail = False

    def send(self, destination, text, attempt):
        self.calls.append((destination, text, attempt))
        if self.fail:
            raise APIError('Тайм-аут')
        return 123


class DeliveryTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp.name)
        self.state = State(self.folder / 'state.sqlite3')
        self.api = FakeAPI()
        self.bot = Bot(self.state, self.api, self.folder / 'input.xlsx', self.folder / 'report.xlsx')
        self.bot.report_status = lambda: None
        self.rows = [certificate_row('anna@example.org', 'Secret A'),
                     certificate_row('anna@example.org', 'Secret B')]
        self.state.sync(self.rows)
        self.bot.rows = lambda: self.rows

    def tearDown(self):
        self.state.db.close()
        self.temp.cleanup()

    def deliver(self, kind='automatic', event='auto-1'):
        self.bot.deliver('anna@example.org', self.rows, kind, event, {'login': 'anna@example.org'})

    def update(self, uid=1, mid=10, **changes):
        update = {'update_id': uid, 'message_id': mid, 'chat': {'id': 'private-chat', 'type': 'private'},
                  'from': {'login': 'anna@example.org', 'robot': False}, 'text': '/certificates'}
        update.update(changes)
        return update

    def test_sent_survives_restart(self):
        self.deliver()
        self.state.db.close()
        self.state = State(self.folder / 'state.sqlite3')
        self.bot.state = self.state
        self.state.recover()
        self.deliver(event='auto-2')
        self.assertEqual(len(self.api.calls), 1)

    def test_ambiguous_network_result_not_retried(self):
        self.api.fail = True
        self.deliver()
        self.api.fail = False
        self.deliver(event='auto-2')
        self.assertEqual(len(self.api.calls), 1)
        self.assertEqual(self.state.db.execute('SELECT status FROM attempts').fetchone()[0], 'unknown')

    def test_crash_before_confirmation_is_unknown(self):
        self.state.reserve('anna@example.org', self.rows, 'automatic', 'interrupted')
        self.state.recover()
        self.deliver()
        self.assertEqual(self.api.calls, [])
        self.assertEqual(self.state.db.execute('SELECT status FROM attempts').fetchone()[0], 'unknown')

    def test_new_explicit_request_allowed(self):
        self.deliver()
        self.bot.handle(self.update())
        self.bot.handle(self.update(2, 11))
        self.assertEqual(len(self.api.calls), 3)
        self.assertEqual(self.api.calls[-1][0], {'chat_id': 'private-chat'})
        self.assertIn('Secret A', self.api.calls[-1][1])
        self.assertIn('Secret B', self.api.calls[-1][1])

    def test_replayed_message_with_new_update_id_not_repeated(self):
        self.bot.handle(self.update())
        self.bot.handle(self.update(2))
        self.assertEqual(len(self.api.calls), 1)

    def test_explicit_delivery_prevents_automatic_duplicate(self):
        self.bot.handle(self.update())
        self.deliver()
        self.assertEqual(len(self.api.calls), 1)

    def test_automatic_only_sends_new_certificates(self):
        self.deliver()
        self.rows.append(certificate_row('anna@example.org', 'New'))
        self.state.sync(self.rows)
        self.deliver(event='auto-2')
        self.assertEqual(self.api.calls[-1][1], 'Ваши сертификаты:\n\nNew')

    def test_group_and_robot_ignored(self):
        self.bot.handle(self.update(chat={'id': 'group', 'type': 'group'}))
        self.bot.handle(self.update(**{'from': {'login': 'anna@example.org', 'robot': True}}))
        self.assertEqual(self.api.calls, [])

    def test_no_lookup_by_message_text(self):
        self.bot.handle(self.update(**{'from': {'login': 'stranger@example.org'}, 'text': '/certificates anna@example.org'}))
        self.assertEqual(len(self.api.calls), 1)
        self.assertNotIn('Secret', self.api.calls[0][1])

    def test_registry_change_rolls_back_all_rows(self):
        new = certificate_row('anna@example.org', 'New')
        changed = dict(self.rows[0], user_login='other@example.org')
        with self.assertRaises(ValueError):
            self.state.sync([new, changed])
        self.assertIsNone(self.state.db.execute("SELECT * FROM certificates WHERE id=?", (new["certificate_id"],)).fetchone())

    def test_two_connections_cannot_reserve_same_certificates(self):
        other = State(self.folder / 'state.sqlite3')
        try:
            self.assertIsNotNone(self.state.reserve('anna@example.org', self.rows, 'automatic', 'a'))
            self.assertIsNone(other.reserve('anna@example.org', self.rows, 'automatic', 'b'))
        finally:
            other.db.close()

    def test_report_failure_does_not_repeat_delivery(self):
        self.bot.report_status = Bot.report_status.__get__(self.bot)
        with patch.object(self.state, 'export', side_effect=PermissionError('Excel open')):
            self.deliver()
            self.deliver(event='auto-2')
        self.assertEqual(len(self.api.calls), 1)
        self.assertEqual(self.state.db.execute('SELECT status FROM attempts').fetchone()[0], 'sent')

    def test_legacy_history_survives_without_input_ids(self):
        import hashlib
        import json
        for i, status in enumerate(['sent', 'unknown', 'sending']):
            login = f'legacy{i}@example.org'
            row = certificate_row(login, 'Legacy text')
            old_id = f'customer-id-{i}'
            with self.state.db:
                self.state.db.execute('INSERT INTO certificates VALUES (?,?,?)',
                                      (old_id, login, hashlib.sha256(b'Legacy text').hexdigest()))
                self.state.db.execute('INSERT INTO attempts VALUES (?,?,?,?,?,?,?,?,?,?)',
                                      (old_id, old_id, 'automatic', login, json.dumps([old_id]), status,
                                       '2026-11-21', None, None, None))
        self.state.db.close()
        self.state = State(self.folder / 'state.sqlite3')
        self.state.recover()
        for i in range(3):
            login = f'legacy{i}@example.org'
            rows = [certificate_row(login, 'Legacy text')]
            self.state.sync(rows)
            self.assertIsNone(self.state.reserve(login, rows, 'automatic', f'new-{i}'))
        self.assertEqual(self.state.db.execute('SELECT COUNT(*) FROM attempts').fetchone()[0], 3)

    def test_text_or_recipient_change_is_new(self):
        self.deliver()
        for row in [certificate_row('anna@example.org', 'Changed text'),
                    certificate_row('another@example.org', 'Secret A')]:
            self.state.sync([row])
            self.assertIsNotNone(self.state.reserve(row['user_login'], [row], 'automatic', row['certificate_id']))


@unittest.skipUnless(importlib.util.find_spec('openpyxl'), 'openpyxl не установлен')
class ExcelTests(unittest.TestCase):
    setUp = DeliveryTests.setUp
    tearDown = DeliveryTests.tearDown
    deliver = DeliveryTests.deliver

    def make_file(self, rows):
        from openpyxl import Workbook
        book = Workbook()
        book.active.title = 'Certificates'
        book.active.append(['user_login', 'certificate_text'])
        for row in rows:
            book.active.append(row)
        path = self.folder / 'input.xlsx'
        book.save(path)
        book.close()
        return path

    def test_excel_validation(self):
        path = self.make_file([['anna@example.org', 'First'], ['anna@example.org', 'Second']])
        self.assertEqual(len(read_registry(path)), 2)
        for rows in [[['anna@example.org', '=1+1']], [['anna', 'Text']]]:
            with self.assertRaises(ValueError):
                read_registry(self.make_file(rows))

    def test_report_can_be_rebuilt(self):
        from openpyxl import load_workbook
        self.deliver()
        path = self.folder / 'report.xlsx'
        self.state.export(path)
        self.state.export(path)
        book = load_workbook(path)
        self.assertEqual(book.sheetnames, ['Certificates', 'Attempts'])
        self.assertEqual(book['Certificates']['C2'].value, 'sent')
        self.assertEqual(book['Attempts'].max_row, 2)
        book.close()

    def test_duplicate_rows_and_reordering_do_not_resend(self):
        rows = [['Anna@Example.org', 'Secret A'], ['anna@example.org', 'Secret B'],
                ['anna@example.org', 'Secret A']]
        first = read_registry(self.make_file(rows))
        self.assertEqual(len(first), 2)
        self.bot.deliver('anna@example.org', first, 'automatic', 'first', {'login': 'anna@example.org'})
        second = read_registry(self.make_file(list(reversed(rows))))
        self.bot.deliver('anna@example.org', second, 'automatic', 'second', {'login': 'anna@example.org'})
        self.assertEqual(len(self.api.calls), 1)

    def test_old_id_column_is_ignored(self):
        from openpyxl import load_workbook
        path = self.make_file([['anna@example.org', 'Secret A']])
        book = load_workbook(path)
        book.active['C1'] = 'certificate_id'
        book.active['C2'] = 'old-id'
        book.save(path)
        book.close()
        self.assertEqual(read_registry(path)[0], certificate_row('anna@example.org', 'Secret A'))


class APITests(unittest.TestCase):
    def setUp(self):
        from unittest.mock import Mock
        self.api = API.__new__(API)
        self.api.transport_error = ConnectionError
        self.api.session = Mock()
        self.response = self.api.session.request.return_value
        self.response.status_code = 200
        self.response.json.return_value = {'ok': True, 'message_id': 123}

    def test_payload_and_private_destination(self):
        self.assertEqual(self.api.send({'chat_id': 'private'}, 'Text', 'attempt-id'), 123)
        kwargs = self.api.session.request.call_args.kwargs
        self.assertEqual(kwargs['json']['payload_id'], 'attempt-id')
        self.assertEqual(kwargs['json']['chat_id'], 'private')
        self.assertFalse(kwargs['allow_redirects'])

    def test_network_failure_is_not_retried(self):
        self.api.session.request.side_effect = ConnectionError()
        with self.assertRaises(APIError):
            self.api.send({'login': 'anna@example.org'}, 'Text', 'attempt-id')
        self.assertEqual(self.api.session.request.call_count, 1)

    def test_unconfirmed_response_is_error(self):
        for data in [{'ok': False}, {'ok': True}, []]:
            self.response.json.return_value = data
            with self.assertRaises(APIError):
                self.api.send({'chat_id': 'private'}, 'Text', 'attempt-id')

    def test_http_failure_does_not_log_body(self):
        self.response.status_code = 403
        with self.assertRaises(APIError) as error:
            self.api.send({'chat_id': 'private'}, 'Text', 'attempt-id')
        self.assertIn('403', str(error.exception))
        self.response.json.assert_not_called()


@unittest.skipUnless(importlib.util.find_spec('openpyxl'), 'openpyxl не установлен')
class StartupTests(unittest.TestCase):
    def run_bot(self, folder, api, check=False, organization=7835550):
        from bot import main
        args = ['bot.py', '--organization-id', str(organization)]
        if check:
            args.append('--check')
        with patch('bot.BASE', folder), patch('bot.load_settings', return_value=('test-token', folder / 'input.xlsx')), \
                patch('bot.API', return_value=api), patch('sys.argv', args), patch('bot.time.sleep'):
            if check or organization != 7835550:
                main()
            else:
                with self.assertRaises(KeyboardInterrupt):
                    main()

    def prepare(self, folder):
        from openpyxl import Workbook
        book = Workbook()
        book.active.title = 'Certificates'
        book.active.append(['user_login', 'certificate_text'])
        book.active.append(['anna@example.org', 'Test certificate'])
        book.save(folder / 'input.xlsx')
        book.close()
        api = FakeAPI()

        def call(method, body=None):
            if method == 'self/get':
                return {'ok': True, 'id': 'test-bot', 'login': 'test-bot', 'organizations': [7835550]}
            raise KeyboardInterrupt()

        api.call = call
        return api

    def test_start_sends_then_restart_skips(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            api = self.prepare(folder)
            self.run_bot(folder, api)
            self.assertEqual(len(api.calls), 1)
            self.run_bot(folder, api)
            self.assertEqual(len(api.calls), 1)

    def test_check_never_sends(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            api = self.prepare(folder)
            self.run_bot(folder, api, check=True)
            self.assertEqual(api.calls, [])

    def test_wrong_organization_never_sends(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            api = self.prepare(folder)
            with self.assertRaises(ValueError):
                self.run_bot(folder, api, organization=123)
            self.assertEqual(api.calls, [])


if __name__ == '__main__':
    unittest.main()
