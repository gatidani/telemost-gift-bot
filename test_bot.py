import importlib.util
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from bot import API, APIError, Bot, State, TokenValidationError, certificate_row, read_gallery_image, read_greeting, read_registry, validate_bot_token


class FakeAPI:
    def __init__(self):
        self.calls = []
        self.fail = False
        self.gallery_calls = []

    def send(self, destination, text, attempt):
        self.calls.append((destination, text, attempt))
        if self.fail:
            raise APIError('Тайм-аут')
        return 123

    def send_gallery(self, destination, text, image):
        self.gallery_calls.append((destination, text, image))
        if self.fail:
            raise APIError('Тайм-аут галереи')
        return 456



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

    def test_arbitrary_sheet_name(self):
        from openpyxl import load_workbook
        path = self.make_file([['anna@example.org', 'Secret A']])
        book = load_workbook(path)
        book.active.title = 'Подарки сотрудникам'
        book.save(path)
        book.close()
        self.assertEqual(read_registry(path)[0], certificate_row('anna@example.org', 'Secret A'))

    def test_registry_need_not_be_first_or_active_sheet(self):
        from openpyxl import load_workbook
        path = self.make_file([['anna@example.org', 'Secret A']])
        book = load_workbook(path)
        book.active.title = 'Лист1'
        note = book.create_sheet('Инструкция', 0)
        note.append(['Как заполнить таблицу'])
        book.active = 0
        book.save(path)
        book.close()
        self.assertEqual(len(read_registry(path)), 1)

    def test_multiple_matching_sheets_are_rejected(self):
        from openpyxl import load_workbook
        path = self.make_file([['anna@example.org', 'Secret A']])
        book = load_workbook(path)
        book.copy_worksheet(book.active).title = 'Другая таблица'
        book.save(path)
        book.close()
        with self.assertRaisesRegex(ValueError, 'несколько листов'):
            read_registry(path)

    def test_missing_or_duplicate_headers_are_rejected(self):
        from openpyxl import load_workbook
        for headers in [['Имя', 'Подарок'], ['user_login', 'user_login', 'certificate_text']]:
            path = self.make_file([])
            book = load_workbook(path)
            for column, value in enumerate(headers, 1):
                book.active.cell(1, column, value)
            book.save(path)
            book.close()
            with self.assertRaisesRegex(ValueError, 'не найден лист'):
                read_registry(path)

    def test_names_loaded_and_merged_across_certificates(self):
        from openpyxl import load_workbook
        path = self.make_file([['anna@example.org', 'First'], ['anna@example.org', 'Second']])
        book = load_workbook(path)
        book.active['C1'] = 'full_name'
        book.active['C3'] = 'Иванова Анна Петровна'
        book.save(path)
        book.close()
        rows = read_registry(path, '{first_name}, поздравляем!')
        self.assertTrue(all(row['full_name'] == 'Иванова Анна Петровна' for row in rows))

    def test_conflicting_names_even_on_duplicate_rows_are_rejected(self):
        from openpyxl import load_workbook
        path = self.make_file([['anna@example.org', 'Same'], ['anna@example.org', 'Same']])
        book = load_workbook(path)
        book.active['C1'] = 'full_name'
        book.active['C2'] = 'Иванова Анна'
        book.active['C3'] = 'Иванова Елена'
        book.save(path)
        book.close()
        with self.assertRaises(ValueError):
            read_registry(path, '{first_name}')


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

    def test_single_image_and_text_use_one_multipart_request(self):
        self.response.json.return_value = {'ok': True, 'message_id': 456, 'images': [{'file_id': 'image-id'}]}
        image = ('greeting.png', b'png-data', 'image/png')
        self.assertEqual(self.api.send_gallery({'login': 'anna@example.org'}, 'Text', image), 456)
        self.api.session.request.assert_called_once()
        args, kwargs = self.api.session.request.call_args
        self.assertTrue(args[1].endswith('messages/sendGallery/'))
        self.assertEqual(kwargs['data'], {'login': 'anna@example.org', 'text': 'Text'})
        self.assertEqual(kwargs['files'], [('images', image)])
        self.assertNotIn('json', kwargs)
        self.assertFalse(kwargs['allow_redirects'])

    def test_gallery_timeout_is_not_retried(self):
        self.api.session.request.side_effect = ConnectionError()
        with self.assertRaises(APIError):
            self.api.send_gallery({'login': 'anna@example.org'}, 'Text', ('greeting.png', b'png', 'image/png'))
        self.assertEqual(self.api.session.request.call_count, 1)

    def test_gallery_requires_confirmation_of_image(self):
        for response in [{'ok': True, 'message_id': 456}, {'ok': True, 'message_id': 456, 'images': []}]:
            self.response.json.return_value = response
            with self.assertRaises(APIError):
                self.api.send_gallery({'login': 'anna@example.org'}, 'Text', ('greeting.png', b'png', 'image/png'))




    def test_token_http_errors_stop_after_one_check(self):
        for status in [401, 403, 404]:
            with self.subTest(status=status):
                self.api.session.request.reset_mock()
                self.response.status_code = status
                with self.assertRaises(TokenValidationError) as error:
                    validate_bot_token(self.api)
                self.assertIn(str(status), str(error.exception))
                self.api.session.request.assert_called_once()
                self.assertTrue(self.api.session.request.call_args.args[1].endswith('self/get'))

    def test_token_requires_confirmed_bot_identity(self):
        for data in [{'ok': False}, {'ok': True}, {'ok': True, 'id': '', 'login': 'bot'}]:
            self.response.json.return_value = data
            with self.assertRaises(TokenValidationError):
                validate_bot_token(self.api)

    def test_send_404_with_invalid_token_stops_fatally(self):
        self.response.status_code = 404
        with self.assertRaises(TokenValidationError):
            self.api.send({'login': 'anna@example.org'}, 'Text', 'attempt-id')
        self.assertEqual(self.api.session.request.call_count, 2)
        self.assertTrue(self.api.session.request.call_args.args[1].endswith('self/get'))

    def test_recipient_404_with_valid_token_remains_recipient_error(self):
        from unittest.mock import Mock
        failed = Mock(status_code=404)
        valid = Mock(status_code=200)
        valid.json.return_value = {'ok': True, 'id': 'bot-id', 'login': 'bot-login'}
        self.api.session.request.side_effect = [failed, valid]
        with self.assertRaises(APIError) as error:
            self.api.send({'login': 'missing@example.org'}, 'Text', 'attempt-id')
        self.assertNotIsInstance(error.exception, TokenValidationError)
        self.assertEqual(error.exception.status_code, 404)

    def test_failed_logical_send_checks_token(self):
        from unittest.mock import Mock
        failed = Mock(status_code=200)
        failed.json.return_value = {'ok': False}
        invalid = Mock(status_code=200)
        invalid.json.return_value = {'ok': False}
        self.api.session.request.side_effect = [failed, invalid]
        with self.assertRaises(TokenValidationError):
            self.api.send({'login': 'anna@example.org'}, 'Text', 'attempt-id')
        self.assertEqual(self.api.session.request.call_count, 2)


class GalleryDeliveryTests(unittest.TestCase):
    setUp = DeliveryTests.setUp
    tearDown = DeliveryTests.tearDown
    deliver = DeliveryTests.deliver

    def enable_gallery(self):
        self.bot.image_path = self.folder / 'greeting.png'
        self.bot.image_path.write_bytes(b'\x89PNG\r\n\x1a\nfixture')
        self.bot.greeting_path = self.folder / 'greeting.txt'
        self.bot.greeting_path.write_text('{full_name}, с Новым годом!', encoding='utf-8')
        for row in self.rows:
            row['full_name'] = 'Иванова Анна Петровна'

    def test_gallery_sends_greeting_and_certificates_without_extra_text(self):
        self.enable_gallery()
        self.deliver()
        self.assertEqual(self.api.calls, [])
        self.assertEqual(len(self.api.gallery_calls), 1)
        self.assertIn('Иванова Анна Петровна, с Новым годом!', self.api.gallery_calls[0][1])
        self.assertIn('Secret A', self.api.gallery_calls[0][1])
        self.assertIn('Secret B', self.api.gallery_calls[0][1])
        row = self.state.db.execute('SELECT status,message_id FROM attempts').fetchone()
        self.assertEqual(tuple(row), ('sent', '456'))

    def test_unknown_gallery_never_falls_back_to_text_or_retries(self):
        self.enable_gallery()
        self.api.fail = True
        self.deliver()
        self.api.fail = False
        self.deliver(event='reconnect')
        self.assertEqual(self.api.calls, [])
        self.assertEqual(len(self.api.gallery_calls), 1)
        self.assertEqual(self.state.db.execute('SELECT status FROM attempts').fetchone()[0], 'unknown')

    def test_restart_skips_confirmed_gallery(self):
        self.enable_gallery()
        self.deliver()
        self.state.db.close()
        self.state = State(self.folder / 'state.sqlite3')
        self.bot.state = self.state
        self.state.recover()
        self.deliver(event='restart')
        self.assertEqual(len(self.api.gallery_calls), 1)

    def test_operator_resend_sends_new_gallery(self):
        self.enable_gallery()
        self.deliver()
        self.deliver(kind='operator_resend', event='new-confirmed-request')
        self.assertEqual(len(self.api.gallery_calls), 2)

    def test_missing_or_invalid_image_stops_before_reserving(self):
        self.bot.image_path = self.folder / 'missing.png'
        for content in [None, b'not-png']:
            if content is not None:
                self.bot.image_path.write_bytes(content)
            with self.assertRaises(ValueError):
                self.deliver()
        self.assertEqual(self.state.db.execute('SELECT COUNT(*) FROM attempts').fetchone()[0], 0)
        self.assertEqual(self.api.gallery_calls, [])

    def test_service_replies_remain_text_only(self):
        self.enable_gallery()
        self.bot.deliver('anna@example.org', [{'certificate_id': 'service'}], 'user_request', 'service',
                         {'login': 'anna@example.org'}, text='Help')
        self.assertEqual(len(self.api.calls), 1)
        self.assertEqual(self.api.gallery_calls, [])


    def test_invalid_token_stops_batch_before_next_recipient(self):
        from unittest.mock import Mock
        self.rows.append(certificate_row('other@example.org', 'Other certificate'))
        self.state.sync(self.rows)
        self.api.send = Mock(side_effect=TokenValidationError('Проверка токена не пройдена'))
        with self.assertRaises(TokenValidationError):
            self.bot.automatic()
        self.api.send.assert_called_once()
        attempts = self.state.db.execute('SELECT login,status FROM attempts').fetchall()
        self.assertEqual([tuple(row) for row in attempts], [('anna@example.org', 'unknown')])


class GreetingTests(unittest.TestCase):
    setUp = DeliveryTests.setUp
    tearDown = DeliveryTests.tearDown
    deliver = DeliveryTests.deliver
    update = DeliveryTests.update

    def save_greeting(self, text):
        self.bot.greeting_path = self.folder / 'greeting.txt'
        self.bot.greeting_path.write_text(text, encoding='utf-8')

    def test_greeting_precedes_all_certificates(self):
        self.save_greeting('С праздником!\nОт вашей команды')
        self.deliver()
        self.assertEqual(self.api.calls[0][1], 'С праздником!\nОт вашей команды\n\nВаши сертификаты:\n\nSecret A\n\nSecret B')

    def test_first_name_is_extracted_from_surname_first(self):
        self.save_greeting('{first_name}, поздравляем! {full_name}')
        for row in self.rows:
            row['full_name'] = 'Иванова Анна Петровна'
        self.deliver()
        self.assertTrue(self.api.calls[0][1].startswith('Анна, поздравляем! Иванова Анна Петровна'))

    def test_explicit_first_name_takes_priority(self):
        self.save_greeting('{first_name}, поздравляем!')
        for row in self.rows:
            row.update(full_name='Иванова Анна Петровна', first_name='Аня')
        self.deliver()
        self.assertTrue(self.api.calls[0][1].startswith('Аня, поздравляем!'))

    def test_missing_name_prevents_attempt(self):
        self.save_greeting('{first_name}, поздравляем!')
        with self.assertRaises(ValueError):
            self.deliver()
        self.assertEqual(self.api.calls, [])
        self.assertEqual(self.state.db.execute('SELECT COUNT(*) FROM attempts').fetchone()[0], 0)

    def test_name_change_does_not_repeat_automatic_delivery(self):
        self.save_greeting('{first_name}, поздравляем!')
        for row in self.rows:
            row['full_name'] = 'Иванова Анна'
        self.deliver()
        for row in self.rows:
            row['full_name'] = 'Иванова Анна Петровна'
        self.deliver(event='new-start')
        self.assertEqual(len(self.api.calls), 1)

    def test_rendered_length_checked_before_reserving(self):
        self.save_greeting('{first_name}')
        for row in self.rows:
            row['first_name'] = 'A' * 6000
        with self.assertRaises(ValueError):
            self.deliver()
        self.assertEqual(self.state.db.execute('SELECT COUNT(*) FROM attempts').fetchone()[0], 0)

    def test_edit_is_loaded_for_next_explicit_request(self):
        self.save_greeting('Первый текст')
        self.bot.handle(self.update())
        self.save_greeting('Новый текст')
        self.bot.handle(self.update(2, 11))
        self.assertTrue(self.api.calls[0][1].startswith('Первый текст'))
        self.assertTrue(self.api.calls[1][1].startswith('Новый текст'))

    def test_edit_does_not_trigger_automatic_duplicate(self):
        self.save_greeting('Первый текст')
        self.deliver()
        self.save_greeting('Новый текст')
        self.deliver(event='second-start')
        self.assertEqual(len(self.api.calls), 1)

    def test_missing_file_reserves_nothing(self):
        self.bot.greeting_path = self.folder / 'missing.txt'
        with self.assertRaises(ValueError):
            self.deliver()
        self.assertEqual(self.api.calls, [])
        self.assertEqual(self.state.db.execute('SELECT COUNT(*) FROM attempts').fetchone()[0], 0)

    def test_combined_length_checked_before_reserving(self):
        self.save_greeting('A' * 5980)
        with self.assertRaises(ValueError):
            self.deliver()
        self.assertEqual(self.state.db.execute('SELECT COUNT(*) FROM attempts').fetchone()[0], 0)

    def test_utf8_bom_is_supported(self):
        self.save_greeting('\ufeffПоздравляем!')
        self.assertEqual(read_greeting(self.bot.greeting_path), 'Поздравляем!')

    def test_invalid_text_files_are_rejected(self):
        for content in [b' ', b'{\\rtf1 example}', b'\xff\xfeinvalid']:
            self.bot.greeting_path = self.folder / 'greeting.txt'
            self.bot.greeting_path.write_bytes(content)
            with self.assertRaises(ValueError):
                read_greeting(self.bot.greeting_path)




@unittest.skipUnless(importlib.util.find_spec('openpyxl'), 'openpyxl не установлен')
class StartupTests(unittest.TestCase):
    def test_invalid_token_is_checked_before_registry_and_state(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            api = self.prepare(folder)
            from unittest.mock import Mock
            api.call = Mock(side_effect=APIError('HTTP 404'))
            with patch('bot.read_registry', side_effect=AssertionError('Registry loaded before token check')):
                with self.assertRaises(APIError):
                    self.run_bot(folder, api, check=True)
            api.call.assert_called_once_with('self/get')
            self.assertFalse((folder / 'bot_state.sqlite3').exists())
            self.assertFalse((folder / 'delivery_status.xlsx').exists())
            self.assertEqual(api.calls, [])
            self.assertEqual(api.gallery_calls, [])

    def run_bot(self, folder, api, check=False, organization=7835550, only_users=None, resend=False, answer='ОТПРАВИТЬ', image=None):
        from bot import main
        args = ['bot.py', '--organization-id', str(organization)]
        if check:
            args.append('--check')
        if resend:
            args.append('--resend')
        if only_users:
            args.extend(['--only-users', *only_users])
        if image is not None:
            args.extend(['--image', image])
        with patch('bot.BASE', folder), patch('bot.load_settings', return_value=('test-token', folder / 'input.xlsx')), \
                patch('bot.API', return_value=api), patch('sys.argv', args), patch('bot.time.sleep'), \
                patch('builtins.input', return_value=answer), patch('bot.GREETING_IMAGE_PATH', ''):
            if check or resend or organization != 7835550:
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
        (folder / 'greeting.txt').write_text('С праздником!', encoding='utf-8')
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

    def test_resend_with_image_sends_single_gallery(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            api = self.prepare(folder)
            (folder / 'greeting.png').write_bytes(b'\x89PNG\r\n\x1a\nfixture')
            self.run_bot(folder, api, resend=True, only_users=['anna@example.org'], image='greeting.png')
            self.assertEqual(api.calls, [])
            self.assertEqual(len(api.gallery_calls), 1)

    def test_check_with_image_never_sends(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            api = self.prepare(folder)
            (folder / 'greeting.png').write_bytes(b'\x89PNG\r\n\x1a\nfixture')
            self.run_bot(folder, api, check=True, image='greeting.png')
            self.assertEqual(api.calls, [])
            self.assertEqual(api.gallery_calls, [])

    def test_wrong_organization_never_sends(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            api = self.prepare(folder)
            with self.assertRaises(ValueError):
                self.run_bot(folder, api, organization=123)
            self.assertEqual(api.calls, [])

    def test_only_selected_user_receives_startup_delivery(self):
        from openpyxl import load_workbook
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            api = self.prepare(folder)
            book = load_workbook(folder / 'input.xlsx')
            book.active.append(['other@example.org', 'Other certificate'])
            book.save(folder / 'input.xlsx')
            book.close()
            self.run_bot(folder, api, only_users=['ANNA@example.org'])
            self.assertEqual(len(api.calls), 1)
            self.assertEqual(api.calls[0][0], {'login': 'anna@example.org'})

    def test_missing_selected_user_prevents_all_delivery(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            api = self.prepare(folder)
            with self.assertRaises(ValueError):
                self.run_bot(folder, api, only_users=['anna@example.org', 'missing@example.org'])
            self.assertEqual(api.calls, [])

    def test_other_users_private_requests_are_ignored(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            state = State(folder / 'state.sqlite3')
            api = FakeAPI()
            try:
                bot = Bot(state, api, folder / 'missing.xlsx', folder / 'report.xlsx', only_users=['selected@example.org'])
                bot.handle({'update_id': 1, 'message_id': 1, 'chat': {'type': 'private', 'id': 'chat'},
                            'from': {'login': 'other@example.org', 'robot': False}, 'text': '/certificates'})
                self.assertEqual(api.calls, [])
            finally:
                state.db.close()

    def test_resend_requires_selected_users(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            api = self.prepare(folder)
            with self.assertRaises(SystemExit) as error:
                self.run_bot(folder, api, resend=True)
            self.assertEqual(error.exception.code, 2)
            self.assertEqual(api.calls, [])

    def test_resend_preserves_history_and_skips_other_recipients(self):
        from openpyxl import load_workbook
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            api = self.prepare(folder)
            self.run_bot(folder, api)
            first_payload = api.calls[0][2]
            book = load_workbook(folder / 'input.xlsx')
            book.active.append(['other@example.org', 'Not selected'])
            book.save(folder / 'input.xlsx')
            book.close()
            self.run_bot(folder, api, resend=True, only_users=['anna@example.org'])
            self.assertEqual(len(api.calls), 2)
            self.assertNotEqual(api.calls[-1][2], first_payload)
            self.assertEqual(api.calls[-1][0], {'login': 'anna@example.org'})
            self.assertIn('С праздником!', api.calls[-1][1])
            state = State(folder / 'bot_state.sqlite3')
            try:
                kinds = [row[0] for row in state.db.execute('SELECT kind FROM attempts ORDER BY rowid')]
                self.assertEqual(kinds, ['automatic', 'operator_resend'])
            finally:
                state.db.close()

    def test_resend_cancellation_sends_nothing(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            api = self.prepare(folder)
            for answer in ['', 'нет', 'yes']:
                self.run_bot(folder, api, resend=True, only_users=['anna@example.org'], answer=answer)
            self.assertEqual(api.calls, [])

    def test_resend_unknown_is_not_automatically_retried(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            api = self.prepare(folder)
            api.fail = True
            self.run_bot(folder, api, resend=True, only_users=['anna@example.org'])
            api.fail = False
            self.run_bot(folder, api)
            self.assertEqual(len(api.calls), 1)

    def test_resend_rejects_missing_recipient_before_confirmation(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            api = self.prepare(folder)
            with patch('builtins.input') as prompt:
                with self.assertRaises(ValueError):
                    self.run_bot(folder, api, resend=True, only_users=['missing@example.org'])
                prompt.assert_not_called()
            self.assertEqual(api.calls, [])

    def test_check_and_resend_are_mutually_exclusive(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            api = self.prepare(folder)
            with self.assertRaises(SystemExit):
                self.run_bot(folder, api, check=True, resend=True, only_users=['anna@example.org'])
            self.assertEqual(api.calls, [])


if __name__ == '__main__':
    unittest.main()
