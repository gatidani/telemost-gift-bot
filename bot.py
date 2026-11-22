"""Бот выдачи сертификатов. Запуск и настройка описаны в README.md."""
import argparse
import hashlib
import json
import logging
import os
from pathlib import Path
import sqlite3
import tempfile
import time
from datetime import datetime, timezone
from uuid import uuid4

from settings import GREETING_PATH, load_settings

BASE = Path(__file__).resolve().parent
LOG = logging.getLogger('giftbot')


def now():
    return datetime.now(timezone.utc).isoformat(timespec='seconds')


def certificate_key(login, digest):
    return 'sha256:' + hashlib.sha256(json.dumps(
        [login.strip().casefold(), digest], ensure_ascii=False).encode()).hexdigest()


def certificate_row(login, text):
    login, text = login.strip().casefold(), text.strip()
    digest = hashlib.sha256(text.encode()).hexdigest()
    return {'certificate_id': certificate_key(login, digest),
            'user_login': login, 'certificate_text': text}


def read_greeting(path):
    try:
        text = Path(path).read_text(encoding='utf-8-sig').strip()
    except UnicodeError:
        raise ValueError('Сохраните greeting.txt как обычный текст в кодировке UTF-8.') from None
    except OSError:
        raise ValueError('Не удалось прочитать поздравление. Положите greeting.txt рядом с bot.py или проверьте GREETING_PATH в settings.py.') from None
    if not text or text.startswith('{\\rtf'):
        raise ValueError('Заполните greeting.txt поздравлением и сохраните как обычный текст, не RTF.')
    if len(text) + len('Ваши сертификаты:\n\n') + 2 >= 6000:
        raise ValueError('Поздравление слишком длинное. Сократите greeting.txt: сообщение с сертификатами должно укладываться в 6000 символов.')
    return text


def read_registry(path, greeting=''):
    """Проверяем всю таблицу до отправки хотя бы одного сообщения."""
    from openpyxl import load_workbook
    try:
        book = load_workbook(path, read_only=True, data_only=False)
    except Exception:
        raise ValueError('Не удалось открыть Excel. Сохраните таблицу как .xlsx.') from None
    try:
        required = ['user_login', 'certificate_text']
        candidates = []
        for sheet in book.worksheets:
            header = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
            if all(header.count(name) == 1 for name in required):
                candidates.append(sheet)
        if not candidates:
            raise ValueError('В certificates.xlsx не найден лист с уникальными заголовками user_login и certificate_text в первой строке. Имя листа может быть любым.')
        if len(candidates) > 1:
            raise ValueError('В certificates.xlsx несколько листов с заголовками user_login и certificate_text. Оставьте одну таблицу сертификатов; имя её листа может быть любым.')
        rows = iter(candidates[0].iter_rows())
        header = next(rows, ())
        names = [c.value for c in header]
        result, seen = [], set()
        for number, cells in enumerate(rows, 2):
            if all(c.value is None for c in cells):
                continue
            values = {}
            for field in required:
                cell = cells[names.index(field)]
                if cell.data_type == 'f' or not isinstance(cell.value, str) or not cell.value.strip():
                    raise ValueError(f'Строка {number}: поле {field} должно содержать непустой текст, не формулу.')
                values[field] = cell.value.strip()
            values['user_login'] = values['user_login'].casefold()
            if '@' not in values['user_login'] or any(c.isspace() for c in values['user_login']):
                raise ValueError(f'Строка {number}: укажите полный корпоративный адрес в user_login.')
            values = certificate_row(values['user_login'], values['certificate_text'])
            if values['certificate_id'] in seen:
                continue
            seen.add(values['certificate_id'])
            result.append(values)
        for login in {r['user_login'] for r in result}:
            if len(message([r for r in result if r['user_login'] == login], greeting)) > 6000:
                raise ValueError('Поздравление и сертификаты одного пользователя превышают 6000 символов. Сократите текст.')
        return result
    finally:
        book.close()


def message(rows, greeting=''):
    body = 'Ваши сертификаты:\n\n' + '\n\n'.join(r['certificate_text'] for r in rows)
    return greeting + '\n\n' + body if greeting else body


class State:
    def __init__(self, path):
        self.db = sqlite3.connect(path, timeout=30)
        self.db.row_factory = sqlite3.Row
        self.db.execute('PRAGMA journal_mode=WAL')
        self.db.execute('PRAGMA synchronous=FULL')
        self.db.executescript('''
            CREATE TABLE IF NOT EXISTS certificates (
                id TEXT PRIMARY KEY, login TEXT NOT NULL, digest TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS attempts (
                id TEXT PRIMARY KEY, event TEXT UNIQUE NOT NULL, kind TEXT NOT NULL,
                login TEXT NOT NULL, ids TEXT NOT NULL, status TEXT NOT NULL,
                created_at TEXT NOT NULL, sent_at TEXT, message_id TEXT, error TEXT);
            CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS certificate_keys (id TEXT PRIMARY KEY, key TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS certificate_texts (key TEXT PRIMARY KEY, text TEXT NOT NULL);
        ''')
        # Старые ID и попытки остаются нетронутыми. Связь восстанавливается
        # по уже сохранённым адресу и хешу текста, без отправки сообщений.
        with self.db:
            for row in self.db.execute('SELECT * FROM certificates').fetchall():
                self.db.execute('INSERT OR IGNORE INTO certificate_keys VALUES (?,?)',
                                (row['id'], certificate_key(row['login'], row['digest'])))

    def recover(self):
        with self.db:
            self.db.execute("UPDATE attempts SET status='unknown', error='Запуск после незавершённой отправки' WHERE status='sending'")

    def sync(self, rows):
        with self.db:
            for row in rows:
                digest = hashlib.sha256(row['certificate_text'].encode()).hexdigest()
                old = self.db.execute('SELECT * FROM certificates WHERE id=?', (row['certificate_id'],)).fetchone()
                if old and (old['login'] != row['user_login'] or old['digest'] != digest):
                    raise ValueError('Конфликт внутреннего ключа сертификата. Отправка остановлена; обратитесь к администратору.')
                self.db.execute('INSERT OR IGNORE INTO certificates VALUES (?,?,?)', (row['certificate_id'], row['user_login'], digest))
                key = certificate_key(row['user_login'], digest)
                self.db.execute('INSERT OR IGNORE INTO certificate_keys VALUES (?,?)', (row['certificate_id'], key))
                self.db.execute('INSERT OR REPLACE INTO certificate_texts VALUES (?,?)', (key, row['certificate_text']))

    def get(self, key, default='0'):
        row = self.db.execute('SELECT value FROM meta WHERE key=?', (key,)).fetchone()
        return row[0] if row else default

    def put(self, key, value):
        with self.db:
            self.db.execute('INSERT OR REPLACE INTO meta VALUES (?,?)', (key, str(value)))

    def reserve(self, login, rows, kind, event):
        self.db.execute('BEGIN IMMEDIATE')
        try:
            if self.db.execute('SELECT 1 FROM attempts WHERE event=?', (event,)).fetchone():
                self.db.rollback()
                return None
            if kind == 'automatic':
                # Любая предыдущая попытка, включая явный запрос, исключает
                # последующую автоматическую выдачу этого сертификата.
                used = set()
                aliases = dict(self.db.execute('SELECT id,key FROM certificate_keys'))
                for a in self.db.execute('SELECT ids FROM attempts WHERE login=?', (login,)):
                    used.update(aliases.get(cid, cid) for cid in json.loads(a[0]))
                rows = [r for r in rows if r['certificate_id'] not in used]
            if not rows:
                self.db.rollback()
                return None
            attempt = str(uuid4())
            self.db.execute('INSERT INTO attempts VALUES (?,?,?,?,?,?,?,?,?,?)',
                            (attempt, event, kind, login, json.dumps([r['certificate_id'] for r in rows]), 'sending', now(), None, None, None))
            self.db.commit()
            return attempt, rows
        except BaseException:
            self.db.rollback()
            raise

    def finish(self, attempt, status, message_id=None, error=None):
        with self.db:
            self.db.execute('UPDATE attempts SET status=?,sent_at=?,message_id=?,error=? WHERE id=?',
                            (status, now() if status == 'sent' else None, str(message_id) if message_id is not None else None, error, attempt))

    def export(self, path):
        from openpyxl import Workbook
        book = Workbook()
        sheet = book.active
        sheet.title = 'Certificates'
        sheet.append(['user_login', 'certificate_text', 'status', 'sent_at', 'message_id', 'last_error', 'certificate_key'])
        attempts = list(self.db.execute('SELECT * FROM attempts ORDER BY rowid'))
        aliases = dict(self.db.execute('SELECT id,key FROM certificate_keys'))
        texts = dict(self.db.execute('SELECT key,text FROM certificate_texts'))
        latest = {}
        for a in attempts:
            for cid in json.loads(a['ids']):
                latest[aliases.get(cid, cid)] = a
        seen = set()
        for c in self.db.execute('SELECT * FROM certificates ORDER BY login,id'):
            key = aliases[c['id']]
            if key in seen:
                continue
            seen.add(key)
            a = latest.get(key)
            sheet.append([c['login'], texts.get(key, 'Текст отсутствует в текущей таблице'),
                          a['status'] if a else 'pending', a['sent_at'] if a else None,
                          a['message_id'] if a else None, a['error'] if a else None, key])
        sheet.column_dimensions['G'].hidden = True
        sheet = book.create_sheet('Attempts')
        columns = ['id', 'event', 'kind', 'login', 'ids', 'status', 'created_at', 'sent_at', 'message_id', 'error']
        sheet.append(columns)
        for a in attempts:
            sheet.append([a[c] for c in columns])
        # Все значения отчёта - текст, включая строки, начинающиеся с '='.
        for sheet in book:
            sheet.freeze_panes = 'A2'
            sheet.auto_filter.ref = sheet.dimensions
            for row in sheet:
                for cell in row:
                    if isinstance(cell.value, str):
                        cell.data_type = 's'
        fd, temp = tempfile.mkstemp(suffix='.xlsx', dir=path.parent)
        os.close(fd)
        try:
            book.save(temp)
            os.replace(temp, path)
        finally:
            book.close()
            if os.path.exists(temp):
                os.unlink(temp)


class APIError(Exception):
    pass


class API:
    URL = 'https://botapi.messenger.yandex.net/bot/v1/'

    def __init__(self, token):
        import requests
        self.transport_error = requests.RequestException
        self.session = requests.Session()
        self.session.headers['Authorization'] = 'OAuth ' + token

    def call(self, method, body=None):
        # Никаких автоматических повторов POST и переходов на другой адрес.
        try:
            response = self.session.request('GET' if body is None else 'POST', self.URL + method,
                                            json=body, timeout=(10, 35), allow_redirects=False)
        except self.transport_error:
            raise APIError('Сбой связи или TLS; результат запроса может быть неизвестен.') from None
        if response.status_code != 200:
            raise APIError(f'API вернул HTTP {response.status_code}; проверьте доступ и настройки.')
        try:
            data = response.json()
        except ValueError:
            raise APIError('Некорректный ответ API.') from None
        if not isinstance(data, dict) or data.get('ok') is not True:
            raise APIError('API не подтвердил выполнение запроса.')
        return data

    def send(self, destination, text, attempt):
        data = self.call('messages/sendText/', {**destination, 'text': text, 'payload_id': attempt,
                                                'disable_web_page_preview': True})
        if not isinstance(data.get('message_id'), int):
            raise APIError('В ответе нет идентификатора отправленного сообщения.')
        return data['message_id']


class Bot:
    def __init__(self, state, api, registry, report, greeting_path=None, only_users=None):
        self.state, self.api, self.registry, self.report = state, api, registry, report
        self.greeting_path = greeting_path
        self.only_users = {login.strip().casefold() for login in only_users} if only_users else None

    def greeting(self):
        return read_greeting(self.greeting_path) if self.greeting_path is not None else ''

    def report_status(self):
        try:
            self.state.export(self.report)
        except OSError:
            LOG.error('Не удалось обновить отчёт. Закройте delivery_status.xlsx в Excel и проверьте права на папку. История сохранена в базе.')

    def rows(self):
        rows = read_registry(self.registry, self.greeting())
        if self.only_users is not None:
            if self.only_users - {row['user_login'] for row in rows}:
                raise ValueError('Не все выбранные пользователи найдены в certificates.xlsx. Проверьте user_login; отправка не начата.')
            rows = [row for row in rows if row['user_login'] in self.only_users]
        self.state.sync(rows)
        return rows

    def deliver(self, login, rows, kind, event, destination, text=None):
        greeting = self.greeting() if text is None else ''
        if text is None and len(message(rows, greeting)) > 6000:
            raise ValueError('Поздравление и сертификаты превышают 6000 символов. Сократите текст.')
        reserved = self.state.reserve(login, rows, kind, event)
        if reserved is None:
            return
        attempt, selected = reserved
        self.report_status()
        try:
            mid = self.api.send(destination, text if text is not None else message(selected, greeting), attempt)
        except APIError as error:
            self.state.finish(attempt, 'unknown', error=str(error))
        else:
            self.state.finish(attempt, 'sent', message_id=mid)
        self.report_status()

    def handle(self, update):
        chat, sender = update.get('chat', {}), update.get('from', {})
        if chat.get('type') != 'private' or sender.get('robot', False):
            return
        # Только доверенные поля API; никаких поисков по тексту сообщения.
        login = sender.get('login', '').casefold()
        if self.only_users is not None and login not in self.only_users:
            return
        if not login or not chat.get('id') or not isinstance(update.get('message_id'), int):
            return
        if 'text' not in update or update.get('forwarded_messages'):
            return
        command = str(update.get('text', '')).strip().casefold()
        commands = {'/start', '/certificates', 'сертификат', 'сертификаты', 'пришли сертификат ещё раз', 'получить сертификаты'}
        rows = self.rows()
        selected = [r for r in rows if r['user_login'] == login] if command in commands else []
        text = None
        if not selected:
            text = ('Сертификаты для вашей учётной записи не найдены. Обратитесь к администратору.'
                    if command in commands else 'Чтобы получить свои сертификаты, отправьте /certificates. Повторный запрос выдаст их ещё раз.')
        # Служебный ответ также фиксируется, чтобы повтор события не создавал сообщения.
        if not selected:
            selected = [{'certificate_id': '__reply__:' + str(update['update_id'])}]
        self.deliver(login, selected, 'user_request', 'message:' + chat['id'] + ':' + str(update['message_id']),
                     {'chat_id': chat['id']}, text)

    def automatic(self):
        rows = self.rows()
        for login in sorted({r['user_login'] for r in rows}):
            self.deliver(login, [r for r in rows if r['user_login'] == login], 'automatic',
                         'auto:' + str(uuid4()), {'login': login})
            time.sleep(1)

    def operator_resend(self):
        if not self.only_users:
            raise ValueError('Для повторной отправки укажите адреса через --only-users.')
        rows = self.rows()
        print('ПРИНУДИТЕЛЬНАЯ ПОВТОРНАЯ ОТПРАВКА')
        print('Будут повторно отправлены поздравление и все сертификаты из таблицы, включая ранее отправленные и unknown:')
        for login in sorted(self.only_users):
            count = sum(row['user_login'] == login for row in rows)
            print(f'  {login}: сертификатов {count}')
        try:
            confirmed = input('Для отправки введите ОТПРАВИТЬ и нажмите Enter. Любой другой ответ отменяет действие: ')
        except EOFError:
            confirmed = ''
        if confirmed.strip() != 'ОТПРАВИТЬ':
            print('Отменено. Сообщения не отправлялись.')
            return
        run_id = str(uuid4())
        for login in sorted(self.only_users):
            self.deliver(login, [row for row in rows if row['user_login'] == login],
                         'operator_resend', f'operator-resend:{run_id}:{login}', {'login': login})
            time.sleep(1)
        print('Повторная отправка завершена. Проверьте результаты в delivery_status.xlsx, включая возможные unknown.')


class ProcessLock:
    """ОС снимает блокировку при аварии; файл истории не удаляется."""
    def __enter__(self):
        self.file = open(BASE / 'bot.lock', 'a+b')
        self.file.seek(0)
        self.file.write(b'0')
        self.file.flush()
        self.file.seek(0)
        try:
            if os.name == 'nt':
                import msvcrt
                msvcrt.locking(self.file.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(self.file, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            self.file.close()
            raise ValueError('Бот уже запущен. Сначала остановите предыдущее окно.') from None
        return self

    def __exit__(self, *args):
        self.file.close()


def main():
    parser = argparse.ArgumentParser(description='Бот сертификатов для Телемоста')
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument('--check', action='store_true', help='Проверить таблицу и API без отправки сообщений')
    mode.add_argument('--send-pending', action='store_true', help='Однократно отправить ещё не отправленные сертификаты')
    mode.add_argument('--resend', action='store_true', help='Принудительно повторить выдачу выбранным --only-users после подтверждения и завершиться')
    parser.add_argument('--organization-id', type=int, help='Проверить принадлежность бота указанной организации перед запуском')
    parser.add_argument('--only-users', nargs='+', help='Отправлять сертификаты и отвечать только указанным корпоративным адресам')
    args = parser.parse_args()
    if args.resend and not args.only_users:
        parser.error('--resend требует --only-users с адресами получателей.')
    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
    token, registry = load_settings()
    with ProcessLock():
        if not (BASE / 'bot_state.sqlite3').exists() and (BASE / 'delivery_status.xlsx').exists():
            raise ValueError('Найден прежний отчёт, но отсутствует история bot_state.sqlite3. Восстановите базу из резервной копии, чтобы не отправить дубли.')
        state = State(BASE / 'bot_state.sqlite3')
        try:
            state.recover()
            greeting_path = Path(GREETING_PATH).expanduser()
            if not greeting_path.is_absolute():
                greeting_path = BASE / greeting_path
            bot = Bot(state, API(token), registry, BASE / 'delivery_status.xlsx', greeting_path, args.only_users)
            bot.report_status()
            bot.rows()
            bot.report_status()
            info = bot.api.call('self/get')
            if args.organization_id is not None and args.organization_id not in info.get('organizations', []):
                raise ValueError('API не подтвердил принадлежность бота указанной организации. Отправка запрещена; проверьте токен и доступ бота.')
            if info.get('webhook_url'):
                raise ValueError('У бота настроен webhook. Попросите администратора отключить его для работы polling.')
            previous = state.get('bot_id', '')
            if previous and previous != info['id']:
                raise ValueError('Токен принадлежит другому боту. Используйте исходный токен и его историю.')
            state.put('bot_id', info['id'])
            if args.check:
                print('Проверка успешна. Таблица и API доступны. Сообщения не отправлялись.')
                return
            if args.resend:
                bot.operator_resend()
                return
            print('Проверки пройдены. Начинается автоматическая отправка новых сертификатов.')
            bot.automatic()
            if args.send_pending:
                return
            print('Бот запущен. Автоматическая отправка завершена; статусы в delivery_status.xlsx. Повторный запрос: /certificates. Остановка: Ctrl+C.')
            delay = 2
            while True:
                try:
                    data = bot.api.call('messages/getUpdates/', {'offset': int(state.get('offset')), 'limit': 100})
                    updates = data.get('updates')
                    if not isinstance(updates, list):
                        raise APIError('В ответе API нет списка обновлений.')
                    for update in sorted(updates, key=lambda u: u['update_id']):
                        if update.get('from', {}).get('login') != info.get('login'):
                            bot.handle(update)
                        state.put('offset', update['update_id'] + 1)
                    delay = 2
                except (APIError, ValueError) as error:
                    LOG.error('%s', error)
                    delay = min(delay * 2, 30)
                bot.report_status()
                time.sleep(delay)
        finally:
            state.db.close()


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print('\nБот остановлен. История отправок сохранена.')
    except ModuleNotFoundError:
        print('Не установлены библиотеки. Выполните: python -m pip install -r requirements.txt')
        raise SystemExit(1)
    except (ValueError, APIError, sqlite3.Error, OSError) as error:
        print(f'Бот остановлен: {error}')
        raise SystemExit(1)
