# =============================================================================
# КЛИЕНТ ВРЕМЕННОЙ ПОЧТЫ — mail.tm и maildrop.cc  ·  цветной консольный TUI
# =============================================================================
#
# Возможности:
#   • меню входа: список недавних ящиков + создание нового на выбранном провайдере;
#   • два провайдера: mail.tm (ящик с паролем) и maildrop.cc (без регистрации);
#   • экран настроек: провайдер, TTL писем, удаление после прочтения, интервал
#     опроса, лимит хранения, сортировка, цвет — меняются прямо в интерфейсе;
#   • таблица входящих с рамками, цветом и прокруткой;
#   • сортировка писем (дата / отправитель / тема, по возрастанию и убыванию);
#   • авто-удаление писем старше N минут (TTL) — настраивается;
#   • авто-удаление письма через N секунд после прочтения — настраивается;
#   • копирование адреса в буфер обмена одной клавишей (c);
#   • Esc / m из почты — возврат в меню выбора ящика, q — выход;
#   • ящики сохраняются между запусками (в tempmail_config.json рядом с .exe);
#   • цветной просмотр письма с прокруткой, HTML-письма читаются как текст.
#
# Все настройки лежат в файле tempmail_config.json рядом с программой,
# их можно менять и вручную, и прямо в интерфейсе (клавиша «,»).
#
# ИНСТРУКЦИЯ ПО СБОРКЕ В .EXE (Windows, PowerShell или cmd):
#
# Шаг 1. Убедитесь, что Python установлен:
#         python --version
#
# Шаг 2. Установите зависимости:
#         pip install requests pyinstaller
#
# Шаг 3. Перейдите в папку с этим файлом:
#         cd C:\путь\до\папки\с\файлом
#
# Шаг 4. Соберите .exe:
#         pyinstaller --onefile --console --name TempMail tempmail.py
#
# Шаг 5. Готовый файл: dist\TempMail.exe
#
# =============================================================================

import copy
import ctypes
import email
import email.policy
import email.utils
import html as html_lib
import json
import os
import random
import re
import shutil
import string
import subprocess
import sys
import textwrap
import time
from datetime import datetime

import requests

# Строки сразу попадают в файл/пайп при перенаправлении вывода
try:
    sys.stdout.reconfigure(line_buffering=True, encoding="utf-8")
except (AttributeError, ValueError):
    pass

# --- Константы ---
API_BASE = "https://api.mail.tm"

DEFAULT_CONFIG = {
    "provider": "mailtm",             # провайдер, на котором создаётся новый ящик
    "poll_interval": 5,               # как часто проверять почту, сек
    "message_ttl_minutes": 30,        # удалять письма старше N минут (0 = выкл)
    "delete_after_read_seconds": 0,   # удалять через N сек после прочтения (0 = выкл)
    "keep_last": 100,                 # хранить не более N писем (0 = без лимита)
    "sort_by": "date",                # date | sender | subject
    "sort_order": "desc",             # desc (новые сверху) | asc
    "color": True,                    # цветной вывод
    "mailbox": None,                  # текущий ящик {"provider", "address", "password"}
    "recent": [],                     # недавние ящики — список таких же записей
}


def app_dir() -> str:
    """Папка рядом с .exe (или со скриптом) — там хранится конфиг."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_PATH = os.path.join(app_dir(), "tempmail_config.json")


# -----------------------------------------------------------------------------
# Конфигурация
# -----------------------------------------------------------------------------

class Config(dict):
    """Словарь настроек с чтением/записью в tempmail_config.json."""

    def __init__(self) -> None:
        super().__init__()
        data = {}
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (OSError, ValueError):
            data = {}

        merged = copy.deepcopy(DEFAULT_CONFIG)
        if isinstance(data, dict):
            merged.update({k: v for k, v in data.items() if k in DEFAULT_CONFIG})
        if not isinstance(merged.get("recent"), list):
            merged["recent"] = []
        self.update(merged)

        if not os.path.exists(CONFIG_PATH):
            self.save()

    def save(self) -> None:
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(dict(self), f, ensure_ascii=False, indent=2)
        except OSError:
            pass


# -----------------------------------------------------------------------------
# Цвета и работа с ANSI
# -----------------------------------------------------------------------------

class Palette:
    """Набор ANSI-кодов; при выключенном цвете все коды пустые."""

    def __init__(self, enabled: bool = True) -> None:
        def code(num) -> str:
            return f"\033[{num}m" if enabled else ""

        self.reset = code(0)
        self.bold = code(1)
        self.dim = code(2)
        self.rev = code(7)
        self.red = code(31)
        self.green = code(32)
        self.yellow = code(33)
        self.blue = code(34)
        self.magenta = code(35)
        self.cyan = code(36)
        self.white = code(97)
        self.gray = code(90)


ANSI_RE = re.compile(r"\033\[[0-9;]*m")


def vlen(text: str) -> int:
    """Длина строки без ANSI-последовательностей."""
    return len(ANSI_RE.sub("", text))


def fit(text, width: int, align: str = "left") -> str:
    """Обрезает/дополняет строку ровно до width видимых символов."""
    text = "" if text is None else str(text)
    if width <= 0:
        return ""
    if len(text) > width:
        text = text[: width - 1] + "…" if width > 1 else text[:1]
    pad = width - len(text)
    if align == "right":
        return " " * pad + text
    if align == "center":
        left = pad // 2
        return " " * left + text + " " * (pad - left)
    return text + " " * pad


def enable_ansi() -> None:
    """Включает обработку ANSI-кодов в консоли Windows 10+."""
    if os.name != "nt":
        return
    try:
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.GetStdHandle(-11)  # STD_OUTPUT_HANDLE
        mode = ctypes.c_uint32()
        if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
            kernel32.SetConsoleMode(handle, mode.value | 0x0004)
    except Exception:
        try:
            os.system("")
        except Exception:
            pass


# -----------------------------------------------------------------------------
# Работа с API
# -----------------------------------------------------------------------------

def _request(method: str, path: str, base: str = API_BASE, **kwargs) -> requests.Response:
    """
    Единая точка запросов: таймауты и повторы при 429/5xx.
    """
    url = base + path
    kwargs.setdefault("timeout", 15)

    last_error = None
    for attempt in range(3):
        try:
            response = requests.request(method, url, **kwargs)
        except requests.exceptions.RequestException as e:
            last_error = e
            time.sleep(1 + attempt)
            continue

        if response.status_code == 429:
            try:
                wait = float(response.headers.get("Retry-After", 2))
            except ValueError:
                wait = 2.0
            time.sleep(min(wait, 10))
            last_error = requests.HTTPError(f"429 Too Many Requests ({url})")
            continue

        if response.status_code >= 500:
            last_error = requests.HTTPError(f"{response.status_code} Server Error ({url})")
            time.sleep(1 + attempt)
            continue

        return response

    raise last_error if last_error else requests.RequestException(f"Запрос не удался: {url}")


def _auth_headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}", "Accept": "application/json"}


def _as_list(payload) -> list:
    """API отдаёт то Hydra-коллекцию, то обычный массив — приводим к списку."""
    if isinstance(payload, list):
        return payload
    if isinstance(payload, dict):
        return payload.get("hydra:member", []) or []
    return []


class ProviderError(RuntimeError):
    """Ошибка на стороне провайдера (не сетевая)."""


class ProviderAuthError(ProviderError):
    """Сессия истекла — нужен повторный вход."""


class MailTmProvider:
    """Провайдер mail.tm: регистрация ящика + JWT, удаление писем поддержано."""

    id = "mailtm"
    title = "mail.tm"
    base = API_BASE
    supports_delete = True
    supports_mailbox_delete = True
    needs_password = True
    note = "Ящик с паролем, живёт, пока вы его не удалите"

    def create_mailbox(self) -> dict:
        response = _request("GET", "/domains", base=self.base)
        response.raise_for_status()
        domains = _as_list(response.json())
        if not domains:
            raise ProviderError("сервис не вернул ни одного домена")

        address = f"{_random_string(10)}@{domains[0]['domain']}"
        password = _random_string(16) + "Aa1!"

        response = _request(
            "POST", "/accounts", base=self.base,
            json={"address": address, "password": password},
            headers={"Content-Type": "application/json", "Accept": "application/json"},
        )
        if response.status_code not in (200, 201):
            raise ProviderError(
                f"не удалось создать ящик (HTTP {response.status_code}): {response.text[:160]}"
            )
        return {"provider": self.id, "address": address, "password": password}

    def login(self, mailbox: dict) -> dict:
        response = _request(
            "POST", "/token", base=self.base,
            json={"address": mailbox["address"], "password": mailbox.get("password", "")},
            headers={"Content-Type": "application/json", "Accept": "application/json"},
        )
        if response.status_code != 200:
            raise ProviderError(f"не удалось авторизоваться (HTTP {response.status_code})")
        return {"token": response.json()["token"], "mailbox": mailbox}

    def list_messages(self, ctx: dict) -> list:
        response = _request("GET", "/messages", base=self.base,
                            headers=_auth_headers(ctx["token"]))
        if response.status_code == 401:
            raise ProviderAuthError("сессия истекла")
        response.raise_for_status()
        return _as_list(response.json())

    def read_message(self, ctx: dict, msg_id: str) -> dict:
        response = _request("GET", f"/messages/{msg_id}", base=self.base,
                            headers=_auth_headers(ctx["token"]))
        if response.status_code == 401:
            raise ProviderAuthError("сессия истекла")
        response.raise_for_status()
        return response.json()

    def delete_message(self, ctx: dict, msg_id: str) -> bool:
        response = _request("DELETE", f"/messages/{msg_id}", base=self.base,
                            headers=_auth_headers(ctx["token"]))
        if response.status_code == 401:
            raise ProviderAuthError("сессия истекла")
        return response.status_code in (200, 204, 404)

    def delete_mailbox(self, ctx: dict) -> bool:
        """Удаляет аккаунт на сервере: GET /me -> DELETE /accounts/{id}."""
        response = _request("GET", "/me", base=self.base,
                            headers=_auth_headers(ctx["token"]))
        if response.status_code == 401:
            raise ProviderAuthError("сессия истекла")
        response.raise_for_status()
        account_id = (response.json() or {}).get("id")
        if not account_id:
            raise ProviderError("сервис не вернул идентификатор ящика")

        response = _request("DELETE", f"/accounts/{account_id}", base=self.base,
                            headers=_auth_headers(ctx["token"]))
        if response.status_code == 401:
            raise ProviderAuthError("сессия истекла")
        return response.status_code in (200, 204, 404)


MAILDROP_QUERY_INBOX = """
query Inbox($m: String!) {
  inbox(mailbox: $m) {
    id
    subject
    date
    headerfrom
  }
}
"""

MAILDROP_QUERY_MESSAGE = """
query Message($m: String!, $id: String!) {
  message(mailbox: $m, id: $id) {
    id
    headerfrom
    subject
    date
    data
  }
}
"""


class MaildropProvider:
    """
    Провайдер maildrop.cc: публичный ящик без регистрации.
    Адрес выбирается свободно (<имя>@maildrop.cc), письма живут ~24 часа.
    Удаление письма через API недоступно — такие письма скрываются локально.
    """

    id = "maildrop"
    title = "maildrop.cc"
    base = "https://api.maildrop.cc"
    supports_delete = False
    supports_mailbox_delete = False
    needs_password = False
    note = "Без регистрации, письма живут около 24 часов, удаление недоступно"

    def _graphql(self, query: str, variables: dict) -> dict:
        response = requests.post(
            f"{self.base}/graphql",
            json={"query": query, "variables": variables},
            timeout=20,
            headers={
                "Content-Type": "application/json",
                "Origin": "https://maildrop.cc",
                "Referer": "https://maildrop.cc/",
                "User-Agent": "Mozilla/5.0 TempMailClient",
            },
        )
        response.raise_for_status()
        payload = response.json()
        if payload.get("errors"):
            message = payload["errors"][0].get("message", "неизвестная ошибка API")
            raise ProviderError(f"maildrop: {message}")
        return payload.get("data") or {}

    def create_mailbox(self) -> dict:
        name = _random_string(10)
        return {"provider": self.id, "address": f"{name}@maildrop.cc"}

    def login(self, mailbox: dict) -> dict:
        name, _, domain = mailbox["address"].partition("@")
        if domain.lower() != "maildrop.cc" or not name:
            raise ProviderError("адрес maildrop.cc должен быть вида имя@maildrop.cc")
        return {"name": name, "mailbox": mailbox}

    @staticmethod
    def _sender(headerfrom: str) -> dict:
        name, address = email.utils.parseaddr(headerfrom or "")
        return {"address": address or (headerfrom or "неизвестно"), "name": name}

    def list_messages(self, ctx: dict) -> list:
        data = self._graphql(MAILDROP_QUERY_INBOX, {"m": ctx["name"]})
        items = data.get("inbox") or []
        return [{
            "id": item.get("id"),
            "from": self._sender(item.get("headerfrom", "")),
            "subject": item.get("subject") or "(без темы)",
            "createdAt": to_iso(item.get("date")) or "",
            "seen": False,
        } for item in items]

    def read_message(self, ctx: dict, msg_id: str) -> dict:
        data = self._graphql(MAILDROP_QUERY_MESSAGE, {"m": ctx["name"], "id": msg_id})
        message = data.get("message") or {}
        raw = message.get("data") or ""
        return parse_raw_email(
            raw,
            msg_id=msg_id,
            fallback_from=message.get("headerfrom", ""),
            fallback_subject=message.get("subject", ""),
            fallback_date=message.get("date", ""),
        )

    def delete_message(self, ctx: dict, msg_id: str) -> bool:
        return False

    def delete_mailbox(self, ctx: dict) -> bool:
        """У maildrop.cc нет аккаунтов — ящик просто убирается из списка."""
        return True


PROVIDERS = {
    MailTmProvider.id: MailTmProvider(),
    MaildropProvider.id: MaildropProvider(),
}
DEFAULT_PROVIDER = MailTmProvider.id


def get_provider(provider_id) -> object:
    """Провайдер по id; неизвестный/старый конфиг -> mail.tm."""
    return PROVIDERS.get(provider_id) or PROVIDERS[DEFAULT_PROVIDER]


def _random_string(length: int) -> str:
    return "".join(random.choices(string.ascii_lowercase + string.digits, k=length))



# -----------------------------------------------------------------------------
# Текстовые утилиты
# -----------------------------------------------------------------------------

def html_to_text(raw_html: str) -> str:
    """Грубое преобразование HTML-письма в читаемый текст."""
    text = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", raw_html)
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"(?i)</(p|div|tr|li|h[1-6]|blockquote)>", "\n", text)
    text = re.sub(r"<[^>]+>", "", text)
    text = html_lib.unescape(text)
    text = text.replace("\r\n", "\n").replace("\xa0", " ")
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def message_body(msg: dict) -> str:
    """Текст письма: сначала text, если пусто — конвертируем HTML."""
    body = (msg.get("text") or "").strip()
    if body:
        return body
    parts = msg.get("html") or []
    if isinstance(parts, str):
        parts = [parts]
    body = html_to_text("\n".join(parts))
    return body or "(письмо не содержит текстового содержимого)"


def sender_address(msg: dict) -> str:
    sender = msg.get("from") or {}
    if isinstance(sender, str):
        return sender
    return sender.get("address") or "неизвестно"


def sender_display(msg: dict) -> str:
    sender = msg.get("from") or {}
    if isinstance(sender, str):
        return sender
    address = sender.get("address") or "неизвестно"
    name = (sender.get("name") or "").strip()
    return f"{name} <{address}>" if name else address


def parse_time(value) -> float:
    """ISO-дата -> epoch. 0.0, если разобрать не удалось."""
    if not value:
        return 0.0
    try:
        return datetime.fromisoformat(str(value).replace("Z", "+00:00")).timestamp()
    except ValueError:
        pass
    try:
        return email.utils.parsedate_to_datetime(str(value)).timestamp()
    except (TypeError, ValueError):
        return 0.0


def to_iso(value) -> str:
    """Любая дата письма -> ISO-строка (для единообразного отображения)."""
    if not value:
        return ""
    text = str(value)
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).isoformat()
    except ValueError:
        pass
    try:
        return email.utils.parsedate_to_datetime(text).isoformat()
    except (TypeError, ValueError):
        return text


def extract_bodies(parsed) -> tuple:
    """Достаёт из MIME-письма текст и HTML."""
    text, html = "", ""
    try:
        if parsed.is_multipart():
            part = parsed.get_body(preferencelist=("plain",))
            if part is not None:
                text = part.get_content()
            else:
                part = parsed.get_body(preferencelist=("html",))
                if part is not None:
                    html = part.get_content()
        else:
            content = parsed.get_content()
            if parsed.get_content_type() == "text/html":
                html = content
            else:
                text = content
    except Exception:
        try:
            payload = parsed.get_payload(decode=True)
            if isinstance(payload, bytes):
                text = payload.decode("utf-8", "replace")
        except Exception:
            text = ""

    return (text or "").strip(), ([html] if html else [])


def parse_raw_email(raw: str, msg_id: str = "", fallback_from: str = "",
                    fallback_subject: str = "", fallback_date: str = "") -> dict:
    """Разбирает «сырое» MIME-письмо (так письма отдаёт maildrop.cc)."""
    result = {
        "id": msg_id,
        "from": {"address": "неизвестно", "name": ""},
        "subject": fallback_subject or "(без темы)",
        "createdAt": fallback_date or "",
        "text": "",
        "html": [],
    }

    parsed = None
    if raw:
        try:
            parsed = email.message_from_string(raw, policy=email.policy.default)
        except Exception:
            parsed = None

    if parsed is None:
        name, address = email.utils.parseaddr(fallback_from or "")
        result["from"] = {"address": address or (fallback_from or "неизвестно"), "name": name}
        result["text"] = (raw or "").strip()
        return result

    name, address = email.utils.parseaddr(str(parsed.get("From") or fallback_from or ""))
    result["from"] = {"address": address or (fallback_from or "неизвестно"), "name": name}
    result["subject"] = str(parsed.get("Subject") or fallback_subject or "(без темы)")
    result["createdAt"] = to_iso(parsed.get("Date")) or fallback_date
    result["text"], result["html"] = extract_bodies(parsed)
    return result


def fmt_time(value) -> str:
    ts = parse_time(value)
    if not ts:
        return str(value or "неизвестно")
    return time.strftime("%d.%m.%Y %H:%M:%S", time.localtime(ts))


def human_age(seconds: float) -> str:
    seconds = int(max(0, seconds))
    if seconds < 60:
        return f"{seconds}с"
    if seconds < 3600:
        return f"{seconds // 60}м"
    if seconds < 86400:
        return f"{seconds // 3600}ч"
    return f"{seconds // 86400}д"


def wrap_block(text: str, width: int) -> list:
    """Разбивает текст на строки шириной не больше width."""
    width = max(10, width)
    lines = []
    for paragraph in text.split("\n"):
        if not paragraph.strip():
            lines.append("")
            continue
        lines.extend(textwrap.wrap(paragraph, width=width) or [""])
    return lines


# Токены с цифрами (коды подтверждения) — по ним можно кликать мышью
CODE_RE = re.compile(
    r"(?<![A-Za-z0-9])(?=[A-Za-z0-9-]*\d)[A-Za-z0-9][A-Za-z0-9_-]{2,14}[A-Za-z0-9](?![A-Za-z0-9])"
)


def code_spans(text: str) -> list:
    """Позиции кодов в строке: [(начало, конец, токен), ...]."""
    spans = []
    for match in CODE_RE.finditer(text or ""):
        token = match.group(0)
        if any(ch.isdigit() for ch in token):
            spans.append((match.start(), match.end(), token))
    return spans


# -----------------------------------------------------------------------------
# Чтение клавиш и мыши
# -----------------------------------------------------------------------------

SPECIAL_KEYS = {
    "H": "up", "P": "down", "K": "left", "M": "right",
    "S": "del", "I": "pgup", "Q": "pgdn",
    "G": "home", "O": "end",
}

VK_KEYS = {
    0x26: "up", 0x28: "down", 0x25: "left", 0x27: "right",
    0x2E: "del", 0x21: "pgup", 0x22: "pgdn", 0x24: "home", 0x23: "end",
}


class _COORD(ctypes.Structure):
    _fields_ = [("X", ctypes.c_short), ("Y", ctypes.c_short)]


class _SMALL_RECT(ctypes.Structure):
    _fields_ = [("Left", ctypes.c_short), ("Top", ctypes.c_short),
                ("Right", ctypes.c_short), ("Bottom", ctypes.c_short)]


class _CONSOLE_SCREEN_BUFFER_INFO(ctypes.Structure):
    _fields_ = [("dwSize", _COORD), ("dwCursorPosition", _COORD),
                ("wAttributes", ctypes.c_ushort), ("srWindow", _SMALL_RECT),
                ("dwMaximumWindowSize", _COORD)]


class _KEY_EVENT_RECORD(ctypes.Structure):
    _fields_ = [
        ("bKeyDown", ctypes.c_int),
        ("wRepeatCount", ctypes.c_ushort),
        ("wVirtualKeyCode", ctypes.c_ushort),
        ("wVirtualScanCode", ctypes.c_ushort),
        ("uChar", ctypes.c_ushort),
        ("dwControlKeyState", ctypes.c_ulong),
    ]


class _MOUSE_EVENT_RECORD(ctypes.Structure):
    _fields_ = [
        ("dwMousePosition", _COORD),
        ("dwButtonState", ctypes.c_ulong),
        ("dwControlKeyState", ctypes.c_ulong),
        ("dwEventFlags", ctypes.c_ulong),
    ]


class _EVENT_UNION(ctypes.Union):
    _fields_ = [("KeyEvent", _KEY_EVENT_RECORD), ("MouseEvent", _MOUSE_EVENT_RECORD)]


class _INPUT_RECORD(ctypes.Structure):
    _fields_ = [("EventType", ctypes.c_ushort), ("Event", _EVENT_UNION)]


class KeyReader:
    """
    Неблокирующее чтение клавиш и мыши.
    Windows: ReadConsoleInputW (мышь, колесо, стрелки), иначе — stdin/select.
    """

    ENABLE_MOUSE_INPUT = 0x0010
    ENABLE_EXTENDED_FLAGS = 0x0080
    ENABLE_QUICK_EDIT_MODE = 0x0040

    def __init__(self) -> None:
        self.windows = (os.name == "nt" and sys.stdin.isatty()
                        and getattr(sys.stdout, "isatty", lambda: False)())
        self.kernel32 = None
        self.handle = None
        self.out_handle = None
        self.original_mode = None
        self.queue = []
        if self.windows:
            self._setup()

    def _setup(self) -> None:
        try:
            self.kernel32 = ctypes.windll.kernel32
            self.handle = self.kernel32.GetStdHandle(-10)
            self.out_handle = self.kernel32.GetStdHandle(-11)
            mode = ctypes.c_uint32()
            if not self.kernel32.GetConsoleMode(self.handle, ctypes.byref(mode)):
                self.windows = False
                return
            self.original_mode = mode.value
            new_mode = ((mode.value | self.ENABLE_MOUSE_INPUT | self.ENABLE_EXTENDED_FLAGS)
                        & ~self.ENABLE_QUICK_EDIT_MODE)
            self.kernel32.SetConsoleMode(self.handle, new_mode)
        except Exception:
            self.windows = False

    def restore(self) -> None:
        """Возвращает консоли обычный режим (в т.ч. быстрое выделение мышью)."""
        if self.original_mode is not None and self.handle and self.kernel32:
            try:
                self.kernel32.SetConsoleMode(self.handle, self.original_mode)
            except Exception:
                pass
            self.original_mode = None

    def _next_record(self):
        if not self.windows:
            return None
        count = ctypes.c_ulong(0)
        if not self.kernel32.GetNumberOfConsoleInputEvents(self.handle, ctypes.byref(count)):
            return None
        if count.value == 0:
            return None
        record = _INPUT_RECORD()
        read = ctypes.c_ulong(0)
        if not self.kernel32.ReadConsoleInputW(self.handle, ctypes.byref(record), 1,
                                               ctypes.byref(read)):
            return None
        return record if read.value else None

    def _window_origin(self):
        info = _CONSOLE_SCREEN_BUFFER_INFO()
        if self.kernel32.GetConsoleScreenBufferInfo(self.out_handle, ctypes.byref(info)):
            return int(info.srWindow.Left), int(info.srWindow.Top)
        return 0, 0

    @staticmethod
    def _translate_key(key):
        if not key.bKeyDown:
            return None
        char = chr(key.uChar) if key.uChar else ""
        if char:
            mapped = {
                "\r": "enter", "\n": "enter", "\x1b": "esc",
                "\x03": "ctrl-c", "\x08": "backspace", "\x7f": "backspace",
                "\t": "tab", " ": "space",
            }.get(char)
            if mapped:
                return mapped
            return char if char.isprintable() else None
        return VK_KEYS.get(key.wVirtualKeyCode)

    def _translate(self, record):
        if record.EventType == 0x0001:            # KEY_EVENT
            return self._translate_key(record.Event.KeyEvent)
        if record.EventType == 0x0002:            # MOUSE_EVENT
            mouse = record.Event.MouseEvent
            flags = mouse.dwEventFlags
            if flags == 0x0004:                   # MOUSE_WHEELED
                delta = ctypes.c_short(mouse.dwButtonState >> 16).value
                return ("mouse_wheel", 1 if delta > 0 else -1) if delta else None
            if flags == 0x0008:                   # горизонтальное колесо
                return None
            if flags == 0 and (mouse.dwButtonState & 0x0001):
                left, top = self._window_origin()
                return ("mouse_click",
                        int(mouse.dwMousePosition.X) - left,
                        int(mouse.dwMousePosition.Y) - top)
        return None

    def read(self):
        if self.queue:
            return self.queue.pop(0)

        if self.windows:
            while True:
                record = self._next_record()
                if record is None:
                    return None
                token = self._translate(record)
                if token is None:
                    continue
                if record.EventType == 0x0001 and record.Event.KeyEvent.wRepeatCount > 1:
                    self.queue.extend([token] * (record.Event.KeyEvent.wRepeatCount - 1))
                return token

        if os.name == "nt":
            import msvcrt
            if not msvcrt.kbhit():
                return None
            ch = msvcrt.getwch()
            if ch in ("\x00", "\xe0"):
                return SPECIAL_KEYS.get(msvcrt.getwch())
            return {
                "\r": "enter", "\n": "enter", "\x1b": "esc",
                "\x03": "ctrl-c", "\x08": "backspace", "\t": "tab", " ": "space",
            }.get(ch, ch)

        import select
        if select.select([sys.stdin], [], [], 0)[0]:
            ch = sys.stdin.read(1)
            return {"\n": "enter", "\x1b": "esc", "\x03": "ctrl-c"}.get(ch, ch)
        return None


# -----------------------------------------------------------------------------
# Приложение
# -----------------------------------------------------------------------------

SORT_NAMES = {"date": "дата", "sender": "отправитель", "subject": "тема"}
SORT_FIELDS = ["date", "sender", "subject"]

# Описание пунктов экрана настроек: подпись, тип, границы, подсказка
SETTINGS_FIELDS = [
    {
        "key": "provider", "label": "Провайдер новых ящиков",
        "kind": "choice",
        "options": [(provider.id, provider.title) for provider in PROVIDERS.values()],
        "hint": "Сервис, на котором создаются новые ящики (можно выбрать и при создании)",
    },
    {
        "key": "message_ttl_minutes", "label": "Время жизни письма (TTL)",
        "kind": "int", "min": 0, "max": 10080, "step": 5, "unit": "мин", "zero": "выкл",
        "hint": "Письма старше этого возраста удаляются (на maildrop — скрываются локально)",
    },
    {
        "key": "delete_after_read_seconds", "label": "Удалять после прочтения",
        "kind": "int", "min": 0, "max": 86400, "step": 5, "unit": "с", "zero": "выкл",
        "hint": "Письмо удаляется через N секунд после закрытия (на maildrop — скрывается)",
    },
    {
        "key": "poll_interval", "label": "Интервал проверки почты",
        "kind": "int", "min": 2, "max": 300, "step": 1, "unit": "с", "zero": "2 с",
        "hint": "Как часто программа спрашивает провайдера о новых письмах",
    },
    {
        "key": "keep_last", "label": "Хранить писем не более",
        "kind": "int", "min": 0, "max": 1000, "step": 10, "unit": "", "zero": "без лимита",
        "hint": "Лишние письма удаляются, начиная с самых старых",
    },
    {
        "key": "sort_by", "label": "Сортировка писем",
        "kind": "choice",
        "options": [("date", "дата"), ("sender", "отправитель"), ("subject", "тема")],
        "hint": "По какому полю сортируется список входящих",
    },
    {
        "key": "sort_order", "label": "Порядок сортировки",
        "kind": "choice",
        "options": [("desc", "новые сверху"), ("asc", "старые сверху")],
        "hint": "Направление сортировки списка писем",
    },
    {
        "key": "color", "label": "Цветной вывод",
        "kind": "bool",
        "hint": "ANSI-цвета в интерфейсе (Enter — переключить)",
    },
    {
        "key": "reset", "label": "Сбросить настройки",
        "kind": "action",
        "hint": "Вернуть значения по умолчанию; ящики и провайдер сохранятся",
    },
]


class App:
    def __init__(self, cfg: Config) -> None:
        self.cfg = cfg
        self.p = Palette(bool(cfg.get("color", True)) and sys.stdout.isatty())
        self.keys = KeyReader()
        self.interactive = sys.stdout.isatty()

        self.mailbox = None
        self.provider = None
        self.ctx = None
        self.messages = []

        self.selected = 0
        self.offset = 0
        self.detail = None
        self.detail_scroll = 0
        self.pending_delete = {}     # id -> epoch, когда удалить
        self.hidden = set()          # (адрес, id) — скрытые локально письма

        self.event = ""
        self.event_color = "green"
        self.event_time = 0.0
        self.error = ""
        self.error_time = 0.0

        self.help_visible = False
        self.next_poll = 0.0
        self.running = True

        self.screen = "start"          # start | provider | inbox | settings
        self.settings_return = "start"  # куда вернуться из настроек
        self.provider_return = "start"  # куда вернуться с выбора провайдера
        self.settings_selection = 0
        self.provider_selection = 0
        self.start_selection = 0
        self.start_offset = 0
        self.loading_text = ""
        self.has_saved = self._saved_mailbox() is not None

        # области экрана под мышь: (строка кадра, индекс)
        self.start_hits = []
        self.provider_hits = []
        self.settings_hits = []
        self.inbox_hits = []
        self.detail_hits = []

    # ---------------- сессия ----------------

    @staticmethod
    def _mailbox_entry(mailbox) -> dict:
        """Нормализует запись ящика: провайдер + адрес (+ пароль, если нужен)."""
        entry = {
            "provider": get_provider((mailbox or {}).get("provider")).id,
            "address": str((mailbox or {}).get("address") or "").strip(),
        }
        if (mailbox or {}).get("password"):
            entry["password"] = mailbox["password"]
        return entry

    @staticmethod
    def _mailbox_key(entry: dict) -> tuple:
        return (entry.get("provider"), entry.get("address", "").lower())

    def _saved_mailbox(self):
        saved = self.cfg.get("mailbox")
        if not isinstance(saved, dict) or not saved.get("address"):
            return None
        entry = self._mailbox_entry(saved)
        if get_provider(entry["provider"]).needs_password and not entry.get("password"):
            return None
        return entry

    def recent_mailboxes(self) -> list:
        """Недавние ящики: текущий и сохранённые в конфиге, без повторов."""
        result, seen = [], set()
        for raw in [self.cfg.get("mailbox")] + list(self.cfg.get("recent") or []):
            if not isinstance(raw, dict) or not raw.get("address"):
                continue
            entry = self._mailbox_entry(raw)
            if get_provider(entry["provider"]).needs_password and not entry.get("password"):
                continue
            key = self._mailbox_key(entry)
            if key in seen:
                continue
            seen.add(key)
            result.append(entry)
            if len(result) >= 4:
                break
        return result

    def remember_mailbox(self, entry: dict) -> None:
        entry = self._mailbox_entry(entry)
        recent, key = [entry], self._mailbox_key(entry)
        for old in self.cfg.get("recent") or []:
            if not isinstance(old, dict) or not old.get("address"):
                continue
            old_entry = self._mailbox_entry(old)
            if self._mailbox_key(old_entry) == key:
                continue
            recent.append(old_entry)
            if len(recent) >= 8:
                break
        self.cfg["mailbox"] = entry
        self.cfg["recent"] = recent
        self.cfg["provider"] = entry["provider"]
        self.cfg.save()
        self.has_saved = True

    def reset_session_state(self) -> None:
        self.messages = []
        self.selected = 0
        self.offset = 0
        self.detail = None
        self.detail_scroll = 0
        self.pending_delete.clear()
        self.help_visible = False

    def open_mailbox(self, mailbox: dict, loading: str = "") -> bool:
        """Вход в ящик любым провайдером и переход в почту."""
        entry = self._mailbox_entry(mailbox)
        provider = get_provider(entry["provider"])
        self.draw_loading(loading or f"Подключаемся к {provider.title}…")
        try:
            ctx = provider.login(entry)
        except Exception as e:
            self.set_error(f"не удалось открыть ящик: {e}")
            return False

        self.provider = provider
        self.ctx = ctx
        self.mailbox = entry
        self.remember_mailbox(entry)
        self.reset_session_state()
        self.token_ready()
        return True

    def start(self) -> None:
        """Автозапуск для линейного режима: продолжить или создать ящик."""
        saved = self._saved_mailbox()
        if saved and self.resume_mailbox(saved):
            return
        if saved:
            self.set_event("сохранённый ящик недоступен, создаём новый", "yellow")
        self.new_mailbox()

    def resume_mailbox(self, mailbox=None) -> bool:
        """Открыть сохранённый (или выбранный в меню) ящик."""
        entry = self._mailbox_entry(mailbox) if mailbox else self._saved_mailbox()
        if not entry or not entry.get("address"):
            self.set_error("сохранённого адреса нет")
            return False

        if not self.open_mailbox(entry, "Входим в сохранённый ящик…"):
            return False
        self.set_event(f"ящик открыт: {entry['address']}", "cyan")
        return True

    def token_ready(self) -> None:
        """Общий переход в почту после успешного входа/создания ящика."""
        self.screen = "inbox"
        self.detail = None
        self.help_visible = False
        self.next_poll = 0.0
        self.poll()

    def new_mailbox(self, provider_id=None) -> bool:
        provider = get_provider(provider_id or self.cfg.get("provider"))
        self.draw_loading(f"Создаём новый ящик на {provider.title}…")
        try:
            mb = provider.create_mailbox()
        except Exception as e:
            self.set_error(f"не удалось создать ящик: {e}")
            return False

        if not self.open_mailbox(mb, f"Открываем новый ящик на {provider.title}…"):
            return False
        self.set_event(f"новый ящик на {provider.title}", "green")
        return True

    def relogin(self) -> bool:
        if not self.mailbox or not self.provider:
            return False
        try:
            self.ctx = self.provider.login(self.mailbox)
            return True
        except Exception as e:
            self.set_error(f"не удалось перелогиниться: {e}")
            return False

    def provider_call(self, action, *args):
        """Вызов метода провайдера с повторным входом при истёкшей сессии."""
        try:
            return action(*args)
        except ProviderAuthError:
            if not self.relogin():
                raise
            return action(*args)

    # ---------------- события / ошибки ----------------

    def set_event(self, text: str, color: str = "green") -> None:
        self.event = text
        self.event_color = color
        self.event_time = time.time()
        if not self.interactive:
            print(f"  [i] {text}")

    def set_error(self, text: str) -> None:
        self.error = text
        self.error_time = time.time()
        if not self.interactive:
            print(f"  [!] {text}")

    def draw_loading(self, text: str) -> None:
        """Короткая надпись по центру во время сетевых операций."""
        self.loading_text = text
        if not self.interactive:
            print(f"  {text}")
            return
        frame = self.render_loading()
        lines = frame.split("\n")
        buffer = ["\033[H"]
        for index, line in enumerate(lines):
            buffer.append(line + "\033[K")
            if index != len(lines) - 1:
                buffer.append("\n")
        sys.stdout.write("".join(buffer))
        sys.stdout.flush()

    # ---------------- данные ----------------

    def view(self) -> list:
        """Письма, отсортированные согласно настройкам."""
        by = self.cfg.get("sort_by", "date")
        reverse = self.cfg.get("sort_order", "desc") == "desc"

        if by == "sender":
            key = lambda m: sender_address(m).lower()
        elif by == "subject":
            key = lambda m: (m.get("subject") or "").lower()
        else:
            key = lambda m: parse_time(m.get("createdAt"))

        return sorted(self.messages, key=key, reverse=reverse)

    def current_id(self):
        v = self.view()
        if 0 <= self.selected < len(v):
            return v[self.selected].get("id")
        return None

    def reselect(self, msg_id) -> None:
        v = self.view()
        if not v:
            self.selected = 0
            return
        for i, m in enumerate(v):
            if m.get("id") == msg_id:
                self.selected = i
                return
        self.selected = max(0, min(self.selected, len(v) - 1))

    def unread_count(self) -> int:
        return sum(1 for m in self.messages if not m.get("seen"))

    def provider_title(self) -> str:
        return self.provider.title if self.provider else "временная почта"

    def poll(self) -> list:
        """Забирает список писем с сервера. Возвращает новые письма."""
        if not self.provider or not self.ctx:
            return []
        try:
            server = self.provider_call(self.provider.list_messages, self.ctx)
        except requests.exceptions.RequestException as e:
            self.set_error(f"сеть: {e}")
            return []
        except ProviderError as e:
            self.set_error(str(e))
            return []

        self.error = ""
        known = {m.get("id"): m for m in self.messages if m.get("id")}
        address = (self.mailbox or {}).get("address", "")
        merged, new = [], []
        for item in server:
            mid = item.get("id")
            if (address, mid) in self.hidden:
                continue
            if mid in known:
                seen_before = bool(known[mid].get("seen"))
                known[mid].update(item)
                if seen_before:
                    known[mid]["seen"] = True
                merged.append(known[mid])
            else:
                merged.append(item)
                new.append(item)

        current = self.current_id()
        self.messages = merged
        self.reselect(current)
        return new

    def enforce_retention(self) -> None:
        """Удаляет письма по TTL, после прочтения и по лимиту хранения."""
        if not self.messages:
            return

        now = time.time()
        victims = {}

        ttl_minutes = int(self.cfg.get("message_ttl_minutes") or 0)
        if ttl_minutes > 0:
            for m in self.messages:
                created = parse_time(m.get("createdAt"))
                if created and now - created >= ttl_minutes * 60:
                    victims[m.get("id")] = f"старше {ttl_minutes} мин"

        for mid, deadline in list(self.pending_delete.items()):
            if now >= deadline:
                victims[mid] = "после прочтения"

        keep = int(self.cfg.get("keep_last") or 0)
        if keep > 0 and len(self.messages) > keep:
            oldest = sorted(self.messages, key=lambda m: parse_time(m.get("createdAt")))
            for m in oldest[: len(oldest) - keep]:
                victims.setdefault(m.get("id"), f"лимит {keep}")

        for mid, reason in victims.items():
            self.delete_message(mid, reason)

    def delete_message(self, msg_id, reason: str) -> None:
        if not msg_id:
            return
        subject = next((m.get("subject") for m in self.messages if m.get("id") == msg_id), "")
        current = self.current_id()

        server_delete = bool(self.provider and self.provider.supports_delete)
        if server_delete:
            try:
                ok = bool(self.provider_call(self.provider.delete_message, self.ctx, msg_id))
            except requests.exceptions.RequestException as e:
                self.set_error(f"не удалось удалить письмо: {e}")
                self.pending_delete.pop(msg_id, None)
                return
            except ProviderError as e:
                self.set_error(f"не удалось удалить письмо: {e}")
                self.pending_delete.pop(msg_id, None)
                return

            if not ok:
                self.set_error("сервер отказался удалять письмо")
                self.pending_delete.pop(msg_id, None)
                return
        else:
            # провайдер не умеет удалять письма — скрываем локально
            self.hidden.add(((self.mailbox or {}).get("address", ""), msg_id))

        self.messages = [m for m in self.messages if m.get("id") != msg_id]
        self.pending_delete.pop(msg_id, None)
        if self.detail and self.detail.get("id") == msg_id:
            self.detail = None
        self.reselect(current)

        label = subject or str(msg_id)
        suffix = "" if server_delete else " · скрыто локально"
        self.set_event(f"удалено «{fit(label, 24).strip()}» — {reason}{suffix}", "yellow")
        if server_delete:
            time.sleep(0.15)  # не упираться в rate-limit

    # ---------------- действия ----------------

    def delete_selected(self) -> None:
        mid = self.current_id()
        if not mid:
            return
        self.delete_message(mid, "вручную")

    def open_selected(self) -> None:
        if not self.provider:
            return
        msg = next((m for m in self.view() if m.get("id") == self.current_id()), None)
        if not msg:
            return
        try:
            full = self.provider_call(self.provider.read_message, self.ctx, msg["id"])
        except requests.exceptions.RequestException as e:
            self.set_error(f"не удалось открыть письмо: {e}")
            return
        except ProviderError as e:
            self.set_error(f"не удалось открыть письмо: {e}")
            return

        if full.get("id"):
            msg["seen"] = True
            self.detail = full
            self.detail_scroll = 0

    def close_detail(self) -> None:
        msg, self.detail = self.detail, None
        seconds = int(self.cfg.get("delete_after_read_seconds") or 0)
        if msg and seconds > 0 and msg.get("id"):
            self.pending_delete[msg["id"]] = time.time() + seconds
            self.set_event(f"письмо удалится через {seconds} с", "yellow")

    def cycle_sort(self) -> None:
        current = self.cfg.get("sort_by", "date")
        nxt = SORT_FIELDS[(SORT_FIELDS.index(current) + 1) % len(SORT_FIELDS)] \
            if current in SORT_FIELDS else "date"
        self.cfg["sort_by"] = nxt
        self.cfg.save()
        self.reselect(self.current_id())
        self.set_event(f"сортировка: {SORT_NAMES[nxt]}", "cyan")

    def toggle_order(self) -> None:
        self.cfg["sort_order"] = "asc" if self.cfg.get("sort_order") == "desc" else "desc"
        self.cfg.save()
        self.reselect(self.current_id())
        self.set_event(f"порядок: {'по возрастанию' if self.cfg['sort_order'] == 'asc' else 'по убыванию'}", "cyan")

    def ask_int(self, label: str, current: int, lo: int, hi: int):
        raw = self.prompt_input(f"{label} [{current}]: ")
        if not raw:
            return None
        try:
            value = int(raw)
        except ValueError:
            self.set_error("нужно ввести целое число")
            return None
        if not lo <= value <= hi:
            self.set_error(f"допустимый диапазон: {lo}..{hi}")
            return None
        return value

    def configure_ttl(self) -> None:
        value = self.ask_int("Срок жизни писем, мин (0 — не удалять)",
                             int(self.cfg.get("message_ttl_minutes") or 0), 0, 10080)
        if value is not None:
            self.cfg["message_ttl_minutes"] = value
            self.cfg.save()
            self.set_event(f"TTL: {'выключен' if value == 0 else str(value) + ' мин'}", "cyan")

    def configure_after_read(self) -> None:
        value = self.ask_int("Удалять после прочтения, сек (0 — не удалять)",
                             int(self.cfg.get("delete_after_read_seconds") or 0), 0, 86400)
        if value is not None:
            self.cfg["delete_after_read_seconds"] = value
            self.cfg.save()
            self.set_event(f"после прочтения: {'выключено' if value == 0 else str(value) + ' с'}", "cyan")

    def configure_interval(self) -> None:
        value = self.ask_int("Интервал проверки почты, сек",
                             int(self.cfg.get("poll_interval") or 5), 2, 300)
        if value is not None:
            self.cfg["poll_interval"] = value
            self.cfg.save()
            self.next_poll = time.time() + value
            self.set_event(f"опрос каждые {value} с", "cyan")

    # ---------------- настройки ----------------

    def setting_text(self, field) -> str:
        """Как значение параметра выглядит в списке настроек."""
        value = self.cfg.get(field["key"])
        kind = field["kind"]
        if kind == "bool":
            return "да" if value else "нет"
        if kind == "choice":
            return dict(field["options"]).get(value, str(value))
        if kind == "action":
            return "Enter"
        if field.get("zero") and not value:
            return field["zero"]
        unit = field.get("unit") or ""
        return f"{value} {unit}".strip()

    def after_setting_change(self, field) -> None:
        self.cfg.save()
        if field["key"] == "color":
            self.p = Palette(bool(self.cfg.get("color")) and self.interactive)
        if field["key"] == "poll_interval":
            self.next_poll = time.time() + int(self.cfg.get("poll_interval") or 5)
        self.set_event(f"{field['label']}: {self.setting_text(field)}", "cyan")
        if self.screen == "inbox":
            self.enforce_retention()

    def change_setting(self, field, delta: int = 1) -> None:
        kind = field["kind"]
        key = field["key"]

        if kind == "action":
            self.reset_settings()
            return

        if kind == "bool":
            self.cfg[key] = not bool(self.cfg.get(key))
            self.after_setting_change(field)
            return

        if kind == "choice":
            options = [value for value, _ in field["options"]]
            current = self.cfg.get(key)
            selected_id = self.current_id()
            index = options.index(current) if current in options else 0
            self.cfg[key] = options[(index + delta) % len(options)]
            self.after_setting_change(field)
            if key in ("sort_by", "sort_order"):
                self.reselect(selected_id)
            return

        current = int(self.cfg.get(key) or 0)
        step = field.get("step", 1)
        self.cfg[key] = max(field["min"], min(field["max"], current + delta * step))
        self.after_setting_change(field)

    def ask_setting(self, field) -> None:
        key = field["key"]
        current = int(self.cfg.get(key) or 0)
        raw = self.prompt_input(f" {field['label']} [{current}]: ")
        if not raw:
            return
        try:
            value = int(raw)
        except ValueError:
            self.set_error("нужно ввести целое число")
            return
        if not field["min"] <= value <= field["max"]:
            self.set_error(f"допустимый диапазон: {field['min']}..{field['max']}")
            return
        self.cfg[key] = value
        self.after_setting_change(field)

    def reset_settings(self) -> None:
        keep = {key: self.cfg.get(key) for key in ("provider", "mailbox", "recent")}
        self.cfg.clear()
        self.cfg.update(copy.deepcopy(DEFAULT_CONFIG))
        self.cfg.update(keep)
        self.cfg.save()
        self.p = Palette(bool(self.cfg.get("color")) and self.interactive)
        self.set_event("настройки сброшены к значениям по умолчанию", "cyan")

    def open_settings(self, back_screen: str = "inbox") -> None:
        self.screen = "settings"
        self.settings_return = back_screen
        self.settings_selection = 0

    def close_settings(self) -> None:
        self.screen = self.settings_return
        if self.screen == "inbox":
            self.next_poll = 0.0

    def handle_settings_key(self, key) -> None:
        items = SETTINGS_FIELDS

        if key == "m":
            self.go_to_menu()
            return
        if key in ("q", "esc", "ctrl-c", "backspace"):
            self.close_settings()
            return
        if key in ("up", "k"):
            self.settings_selection = (self.settings_selection - 1) % len(items)
            return
        if key in ("down", "j"):
            self.settings_selection = (self.settings_selection + 1) % len(items)
            return

        field = items[self.settings_selection]
        if key == "enter":
            self.activate_setting(field)
        elif key in ("left", "right"):
            if field["kind"] == "int":
                self.change_setting(field, -1 if key == "left" else 1)
            elif field["kind"] == "choice":
                self.change_setting(field, -1 if key == "left" else 1)
            elif field["kind"] == "bool":
                self.change_setting(field)

    def activate_setting(self, field) -> None:
        if field["kind"] == "int":
            self.ask_setting(field)
        else:
            self.change_setting(field)

    # ---------------- экран входа ----------------

    def start_options(self) -> list:
        """Пункты меню: недавние ящики, «Создать почту», настройки, выход."""
        options = []
        current = self._saved_mailbox()
        current_key = self._mailbox_key(current) if current else None

        for entry in self.recent_mailboxes():
            provider = get_provider(entry["provider"])
            extra = provider.title
            if current_key == self._mailbox_key(entry):
                extra += " · текущий"
            options.append({
                "kind": "open",
                "mailbox": entry,
                "label": entry["address"],
                "extra": extra,
            })

        options.append({
            "kind": "new",
            "label": "Создать почту",
            "extra": "выбор провайдера",
        })

        options.append({"kind": "settings", "label": "Настройки", "extra": ""})
        options.append({"kind": "quit", "label": "Выход", "extra": ""})
        return options

    def activate_start_option(self, option: dict) -> None:
        kind = option.get("kind")
        if kind == "open":
            self.resume_mailbox(option.get("mailbox"))
        elif kind == "new":
            self.open_provider_choice()
        elif kind == "settings":
            self.open_settings("start")
        elif kind == "quit":
            self.running = False

    def handle_start_key(self, key) -> None:
        options = self.start_options()
        if not options:
            return

        if key in ("q", "esc", "ctrl-c"):
            self.running = False
            return
        if key in ("up", "k"):
            self.start_selection = (self.start_selection - 1) % len(options)
            return
        if key in ("down", "j"):
            self.start_selection = (self.start_selection + 1) % len(options)
            return
        if isinstance(key, str) and len(key) == 1 and key.isdigit():
            index = int(key) - 1
            if 0 <= index < len(options):
                self.start_selection = index
                self.activate_start_option(options[index])
            return
        if key == "enter":
            index = max(0, min(self.start_selection, len(options) - 1))
            self.activate_start_option(options[index])
            return
        if key in ("d", "del"):
            index = max(0, min(self.start_selection, len(options) - 1))
            option = options[index]
            if option.get("kind") == "open":
                self.delete_mailbox_entry(option.get("mailbox"))
            else:
                self.set_event("удалить можно только сохранённый ящик", "yellow")

    # ---------------- выбор провайдера ----------------

    def provider_list(self) -> list:
        return list(PROVIDERS.values())

    def open_provider_choice(self, back_screen: str = "start") -> None:
        self.screen = "provider"
        self.provider_return = back_screen if back_screen in ("start", "inbox") else "start"
        ids = list(PROVIDERS)
        current = self.cfg.get("provider")
        self.provider_selection = ids.index(current) if current in ids else 0
        self.error = ""
        self.error_time = 0.0

    def close_provider_choice(self) -> None:
        self.screen = self.provider_return
        if self.screen == "start":
            for index, option in enumerate(self.start_options()):
                if option.get("kind") == "new":
                    self.start_selection = index
                    break

    def create_mailbox_at(self, index: int) -> None:
        providers = self.provider_list()
        if not providers:
            return
        provider = providers[max(0, min(index, len(providers) - 1))]
        self.cfg["provider"] = provider.id
        self.cfg.save()
        self.new_mailbox(provider.id)

    def handle_provider_key(self, key) -> None:
        providers = self.provider_list()
        if not providers:
            return

        if key in ("q", "esc", "ctrl-c", "backspace"):
            self.close_provider_choice()
            return
        if key in ("up", "k"):
            self.provider_selection = (self.provider_selection - 1) % len(providers)
            return
        if key in ("down", "j"):
            self.provider_selection = (self.provider_selection + 1) % len(providers)
            return
        if isinstance(key, str) and len(key) == 1 and key.isdigit():
            index = int(key) - 1
            if 0 <= index < len(providers):
                self.provider_selection = index
                self.create_mailbox_at(index)
            return
        if key == "enter":
            self.create_mailbox_at(self.provider_selection)

    def go_to_menu(self) -> None:
        """Выход из почты/настроек в меню выбора ящика."""
        self.screen = "start"
        self.detail = None
        self.help_visible = False
        self.start_selection = 0
        self.start_offset = 0
        self.error = ""
        self.error_time = 0.0
        self.set_event("выберите ящик или создайте новый", "cyan")

    def copy_text(self, text: str) -> bool:
        try:
            process = subprocess.Popen("clip", stdin=subprocess.PIPE, shell=True)
            process.communicate(str(text).encode("utf-16-le"))
            return True
        except Exception as e:
            self.set_error(f"не удалось скопировать: {e}")
            return False

    def copy_address(self) -> None:
        address = self.mailbox["address"] if self.mailbox else ""
        if not address:
            return
        if self.copy_text(address):
            self.set_event("адрес скопирован в буфер обмена", "cyan")

    def copy_token(self, token: str) -> None:
        if token and self.copy_text(token):
            self.set_event(f"скопировано: {token}", "green")

    # ---------------- удаление ящика ----------------

    def confirm(self, question: str) -> bool:
        """Подтверждение действия по одной клавише (y — да, остальное — нет)."""
        if not self.interactive:
            return True
        while self.running:
            self.draw(prompt_line=f" {question} (y/N): ")
            key = self.keys.read()
            if key is None:
                time.sleep(0.04)
                continue
            if isinstance(key, tuple):
                continue
            return key in ("y", "Y", "н", "Н")
        return False

    def forget_mailbox(self, entry) -> None:
        """Убирает ящик из списка недавних (и из текущего, если совпал)."""
        key = self._mailbox_key(self._mailbox_entry(entry))
        recent = [
            old for old in (self.cfg.get("recent") or [])
            if isinstance(old, dict) and self._mailbox_key(self._mailbox_entry(old)) != key
        ]
        self.cfg["recent"] = recent
        saved = self.cfg.get("mailbox")
        if isinstance(saved, dict) and self._mailbox_key(self._mailbox_entry(saved)) == key:
            self.cfg["mailbox"] = None
        self.cfg.save()
        self.has_saved = self._saved_mailbox() is not None

    def delete_mailbox_entry(self, entry) -> bool:
        """Удаляет почту: на mail.tm — аккаунт на сервере, maildrop — из списка."""
        entry = self._mailbox_entry(entry or {})
        address = entry.get("address")
        if not address:
            return False

        provider = get_provider(entry["provider"])
        is_current = bool(self.mailbox) and self._mailbox_key(self.mailbox) == self._mailbox_key(entry)

        if provider.supports_mailbox_delete:
            question = f"Удалить почту {address} на {provider.title}?"
        else:
            question = f"Убрать {address} из списка? Письма на сервере останутся."
        if not self.confirm(question):
            return False

        if provider.supports_mailbox_delete:
            self.draw_loading(f"Удаляем ящик на {provider.title}…")
            reason = "сервис отказался удалять ящик"
            try:
                ctx = self.ctx if is_current and self.ctx else provider.login(entry)
                try:
                    ok = bool(provider.delete_mailbox(ctx))
                except ProviderAuthError:
                    ctx = provider.login(entry)
                    ok = bool(provider.delete_mailbox(ctx))
            except Exception as e:
                ok = False
                reason = str(e)

            if not ok:
                if not self.confirm("Не удалось удалить на сервере. Убрать из списка?"):
                    self.set_error(f"ящик не удалён: {reason}")
                    return False

        self.forget_mailbox(entry)
        if is_current:
            self.mailbox = None
            self.provider = None
            self.ctx = None
            self.reset_session_state()

        self.screen = "start"
        self.start_selection = 0
        self.start_offset = 0
        label = "почта удалена" if provider.supports_mailbox_delete else "убрано из списка"
        self.set_event(f"{label}: {address}", "yellow")
        return True

    # ---------------- мышь ----------------

    @staticmethod
    def _hit(hits, row: int, column: int, width: int):
        """Индекс элемента по координате клика (строка кадра + колонка)."""
        if row < 0 or not 1 <= column <= width - 2:
            return None
        for hit_row, index in hits:
            if hit_row == row:
                return index
        return None

    def handle_mouse_event(self, token) -> None:
        if not isinstance(token, tuple) or not token:
            return
        if token[0] == "mouse_wheel":
            self.handle_wheel(int(token[1]))
        elif token[0] == "mouse_click":
            self.handle_click(int(token[1]), int(token[2]))

    def handle_wheel(self, delta: int) -> None:
        # delta > 0 — колесо вверх (к началу списка), delta < 0 — вниз
        step = -1 if delta > 0 else 1

        if self.screen == "start":
            options = self.start_options()
            if options:
                self.start_selection = (self.start_selection + step) % len(options)
            return
        if self.screen == "provider":
            providers = self.provider_list()
            if providers:
                self.provider_selection = (self.provider_selection + step) % len(providers)
            return
        if self.screen == "settings":
            self.settings_selection = (self.settings_selection + step) % len(SETTINGS_FIELDS)
            return
        if self.detail is not None:
            self.detail_scroll = max(0, self.detail_scroll - delta * 3)
            return

        view = self.view()
        if view:
            self.selected = max(0, min(len(view) - 1, self.selected + step))

    def handle_click(self, column: int, row: int) -> None:
        width, _ = self.size()

        if self.screen == "start":
            index = self._hit(self.start_hits, row, column, width)
            if index is not None:
                options = self.start_options()
                if 0 <= index < len(options):
                    self.start_selection = index
                    self.activate_start_option(options[index])
            return

        if self.screen == "provider":
            index = self._hit(self.provider_hits, row, column, width)
            if index is not None and index < len(self.provider_list()):
                self.provider_selection = index
                self.create_mailbox_at(index)
            return

        if self.screen == "settings":
            index = self._hit(self.settings_hits, row, column, width)
            if index is not None and 0 <= index < len(SETTINGS_FIELDS):
                self.settings_selection = index
                self.activate_setting(SETTINGS_FIELDS[index])
            return

        if self.help_visible:
            self.help_visible = False
            return

        if self.detail is not None:
            # клик по коду подтверждения — копируем его в буфер обмена
            for hit_row, start, end, token in self.detail_hits:
                if hit_row == row and start <= column < end:
                    self.copy_token(token)
                    return
            return

        if row == 1 and 1 <= column <= width - 2:
            self.copy_address()
            return

        index = self._hit(self.inbox_hits, row, column, width)
        if index is not None:
            view = self.view()
            if 0 <= index < len(view):
                if index == self.selected:
                    self.open_selected()
                else:
                    self.selected = index

    def handle_key(self, key) -> None:
        if isinstance(key, tuple):
            self.handle_mouse_event(key)
            return

        if self.screen == "start":
            self.handle_start_key(key)
            return
        if self.screen == "settings":
            self.handle_settings_key(key)
            return
        if self.screen == "provider":
            self.handle_provider_key(key)
            return

        if self.help_visible:
            self.help_visible = False
            return

        if self.detail is not None:
            if key in ("esc", "q", "backspace"):
                self.close_detail()
            elif key == "m":
                self.detail = None
                self.go_to_menu()
            elif key in ("up", "k"):
                self.detail_scroll -= 1
            elif key in ("down", "j"):
                self.detail_scroll += 1
            elif key == "pgup":
                self.detail_scroll -= 10
            elif key == "pgdn":
                self.detail_scroll += 10
            elif key in ("d", "del"):
                msg_id = self.detail.get("id")
                self.detail = None
                self.delete_message(msg_id, "вручную")
            elif key == "D":
                entry = self.mailbox or self._saved_mailbox()
                if entry and self.delete_mailbox_entry(entry):
                    self.detail = None
            return

        count = len(self.view())
        if key in ("q", "ctrl-c"):
            self.running = False
        elif key in ("esc", "m"):
            self.go_to_menu()
        elif key in ("up", "k") and count:
            self.selected = max(0, self.selected - 1)
        elif key in ("down", "j") and count:
            self.selected = min(count - 1, self.selected + 1)
        elif key == "home":
            self.selected = 0
        elif key == "end":
            self.selected = max(0, count - 1)
        elif key == "enter":
            self.open_selected()
        elif key in ("d", "del"):
            self.delete_selected()
        elif key == "D":
            if self.mailbox:
                self.delete_mailbox_entry(self.mailbox)
        elif key == "s":
            self.cycle_sort()
        elif key == "o":
            self.toggle_order()
        elif key == "t":
            self.configure_ttl()
        elif key == "p":
            self.configure_after_read()
        elif key == "i":
            self.configure_interval()
        elif key == "n":
            self.open_provider_choice("inbox")
        elif key == ",":
            self.open_settings("inbox")
        elif key == "c":
            self.copy_address()
        elif key == "r":
            self.poll()
            self.set_event("обновлено", "cyan")
        elif key == "?":
            self.help_visible = True

    # ---------------- ввод строки ----------------

    def prompt_input(self, label: str, initial: str = ""):
        """Строка ввода в нижней части экрана. Возвращает str или None (Esc)."""
        buffer = str(initial or "")
        while self.running:
            self.draw(prompt_line=f" {label}{buffer}▌")
            key = self.keys.read()
            if key is None:
                time.sleep(0.04)
                continue
            if key == "enter":
                return buffer.strip()
            if key in ("esc", "ctrl-c"):
                return None
            if key == "backspace":
                buffer = buffer[:-1]
            elif isinstance(key, str) and len(key) == 1 and key.isprintable():
                buffer += key
        return None

    # ---------------- отрисовка ----------------

    def size(self):
        ts = shutil.get_terminal_size((100, 32))
        width = min(max(30, ts.columns - 1), 100)
        height = max(16, ts.lines)
        return width, height

    @staticmethod
    def _columns(inner: int):
        """Ширины колонок таблицы: (отправитель, возраст, тема)."""
        age_w = 5
        sender_w = 18 if inner >= 62 else 0
        subject_w = max(8, inner - 14 - (sender_w + 1 if sender_w else 0))
        return sender_w, age_w, subject_w

    def _top(self, title: str, width: int, right: str = "") -> str:
        p = self.p
        inner = width - 2
        title = f" {title} "
        right = f" {right} " if right else ""
        if len(title) + len(right) > inner - 1:
            right = ""
        if len(title) > inner - 1:
            title = fit(title, max(0, inner - 1))
        fill = inner - 1 - len(title) - len(right)
        return (f"{p.gray}┌─{p.reset}{p.bold}{p.cyan}{title}{p.reset}"
                f"{p.gray}{'─' * max(0, fill)}{p.reset}{p.dim}{right}{p.reset}{p.gray}┐{p.reset}")

    def _bottom(self, width: int) -> str:
        p = self.p
        return f"{p.gray}└{'─' * (width - 2)}┘{p.reset}"

    def _sep(self, width: int) -> str:
        p = self.p
        return f"{p.gray}├{'─' * (width - 2)}┤{p.reset}"

    def _row(self, content: str, width: int) -> str:
        p = self.p
        return f"{p.gray}│{p.reset}{content}{p.gray}│{p.reset}"

    def colorize_codes(self, text: str, base: str = "") -> str:
        """Подсвечивает коды подтверждения жёлтым (видимая длина не меняется)."""
        p = self.p
        spans = code_spans(text)
        if not spans:
            return f"{base}{text}{p.reset}" if base else text
        parts, last = [], 0
        for start, end, token in spans:
            if start > last:
                chunk = text[last:start]
                parts.append(f"{base}{chunk}{p.reset}" if base else chunk)
            parts.append(f"{p.bold}{p.yellow}{token}{p.reset}")
            last = end
        if last < len(text):
            chunk = text[last:]
            parts.append(f"{base}{chunk}{p.reset}" if base else chunk)
        return "".join(parts)

    def _mailbox_row(self, inner: int) -> str:
        p = self.p
        address = self.mailbox["address"] if self.mailbox else "…"
        prefix = " Ящик: "
        full_right = f"писем {len(self.messages)} · новых {self.unread_count()} "
        short_right = f"{len(self.messages)}/{self.unread_count()} "

        if inner - len(prefix) - len(full_right) >= 18:
            right = full_right
        elif inner - len(prefix) - len(short_right) >= 8:
            right = short_right
        else:
            right = ""

        room = max(4, inner - len(prefix) - len(right) - 1)
        address = fit(address, room)
        pad = max(0, inner - len(prefix) - len(address) - len(right))
        return (f"{p.dim}{prefix}{p.reset}{p.bold}{p.white}{address}{p.reset}"
                f"{' ' * pad}{p.dim}{right}{p.reset}")

    def _head_row(self, inner: int) -> str:
        p = self.p
        sender_w, age_w, subject_w = self._columns(inner)
        if sender_w:
            text = (f" {fit('№', 3, 'right')}   {fit('ОТ КОГО', sender_w)} "
                    f"{fit('ТЕМА', subject_w)} {fit('ВОЗР.', age_w, 'right')} ")
        else:
            text = (f" {fit('№', 3, 'right')}   {fit('ТЕМА', subject_w)} "
                    f"{fit('ВОЗР.', age_w, 'right')} ")
        return f"{p.dim}{fit(text, inner)}{p.reset}"

    def _message_row(self, msg: dict, number: int, selected: bool, inner: int) -> str:
        p = self.p
        sender_w, age_w, subject_w = self._columns(inner)

        created = parse_time(msg.get("createdAt"))
        age = human_age(time.time() - created) if created else "—"
        unread = "●" if not msg.get("seen") else " "

        num = fit(str(number), 3, "right")
        subject = fit(msg.get("subject") or "(без темы)", subject_w)
        age = fit(age, age_w, "right")
        marker_color = p.green if unread == "●" else ""
        sender = fit(sender_address(msg), sender_w) if sender_w else ""

        if selected:
            plain = f" {num} {unread} " + (f"{sender} " if sender_w else "") + f"{subject} {age} "
            return p.rev + fit(plain, inner) + p.reset

        return (f" {p.dim}{num}{p.reset} "
                f"{marker_color}{unread}{p.reset} "
                + (f"{p.cyan}{sender}{p.reset} " if sender_w else "")
                + f"{p.white}{subject}{p.reset} "
                f"{p.dim}{age}{p.reset} ")

    def _status_row(self, inner: int) -> str:
        p = self.p
        ttl = int(self.cfg.get("message_ttl_minutes") or 0)
        after = int(self.cfg.get("delete_after_read_seconds") or 0)
        interval = int(self.cfg.get("poll_interval") or 5)
        by = SORT_NAMES.get(self.cfg.get("sort_by", "date"), "дата")
        arrow = "↓" if self.cfg.get("sort_order", "desc") == "desc" else "↑"

        if inner >= 62:
            left = (f" TTL {str(ttl) + ' мин' if ttl else 'выкл'} · "
                    f"после чтения {str(after) + ' с' if after else 'выкл'} · "
                    f"сорт {by} {arrow} · опрос {interval} с ")
        else:
            left = (f" TTL {str(ttl) + 'м' if ttl else '-'} · "
                    f"чтение {str(after) + 'с' if after else '-'} · "
                    f"{by} {arrow} · {interval}с ")
        if self.pending_delete:
            left += f"· удаление {max(0, int(min(self.pending_delete.values()) - time.time()))} с "

        if self.error and time.time() - self.error_time < 12:
            right, color = self.error, p.red
        elif self.event:
            right, color = self.event, getattr(p, self.event_color, p.dim)
        else:
            right, color = "", p.dim

        right_w = min(len(right) + 2, inner // 2)
        left_w = max(0, inner - right_w)
        return (f"{p.dim}{fit(left, left_w)}{p.reset}"
                f"{color}{fit(right + ' ', right_w, 'right')}{p.reset}")

    def _hint_row(self, inner: int, prompt_line=None) -> str:
        p = self.p
        if prompt_line is not None:
            return f"{p.yellow}{fit(prompt_line, inner)}{p.reset}"
        hint = (" ↑↓ Enter · d письмо · D почту · n новая · "
                ", настройки · ? · q выход")
        return f"{p.dim}{fit(hint, inner)}{p.reset}"

    def _panel_rows(self, body: list, inner: int, available: int) -> list:
        """Дополняет набор строк пустыми до available и оборачивает в рамку."""
        clipped = body[:available]
        rows = []
        for line in clipped:
            extra = inner - vlen(line)
            if extra > 0:
                line = line + " " * extra
            rows.append(self._row(line, inner + 2))
        for _ in range(available - len(clipped)):
            rows.append(self._row(" " * inner, inner + 2))
        return rows

    def option_line(self, number, label, extra, selected: bool, inner: int) -> str:
        """Строка пункта меню: номер, подпись слева, провайдер справа."""
        p = self.p
        prefix = f" {number:>2}. "
        room = max(0, inner - len(prefix))
        extra_w = min(len(extra) + 1, room // 2) if extra and room >= 26 else 0
        label_w = room - extra_w

        plain = (prefix + fit(label, label_w)
                 + (fit(extra, extra_w, "right") if extra_w else ""))
        if selected:
            return f"{p.rev}{fit(plain, inner)}{p.reset}"
        return (f"{p.dim}{prefix}{p.reset}"
                f"{p.white}{fit(label, label_w)}{p.reset}"
                f"{p.dim}{fit(extra, extra_w, 'right') if extra_w else ''}{p.reset}")

    def render_start(self, prompt_line=None) -> str:
        p = self.p
        width, height = self.size()
        inner = width - 2
        available = max(3, height - 4)

        options = self.start_options()
        self.start_selection = max(0, min(self.start_selection, len(options) - 1)) if options else 0

        ttl = int(self.cfg.get("message_ttl_minutes") or 0)
        after = int(self.cfg.get("delete_after_read_seconds") or 0)
        summary = (f"  Новые ящики: {get_provider(self.cfg.get('provider')).title} · "
                   f"TTL {str(ttl) + ' мин' if ttl else 'выкл'} · "
                   f"после чтения {str(after) + ' с' if after else 'выкл'}")

        tail = ["", f"{p.dim}{fit(summary, inner)}{p.reset}"]
        if self.error and time.time() - self.error_time < 15:
            tail += ["", f"{p.red}{fit('  ' + self.error, inner)}{p.reset}"]
        elif self.event:
            tail += ["", f"{p.dim}{fit('  ' + self.event, inner)}{p.reset}"]

        head = [
            "",
            f"{p.bold}{p.white}{fit('  Временная почта', inner)}{p.reset}",
            f"{p.dim}{fit('  Выберите ящик или создайте новый — регистрация не нужна', inner)}{p.reset}",
            "",
        ]

        room = max(1, available - len(head) - len(tail))
        visible = room
        if len(options) > room and room > 2:
            visible = room - 1

        if self.start_selection < self.start_offset:
            self.start_offset = self.start_selection
        if self.start_selection >= self.start_offset + visible:
            self.start_offset = self.start_selection - visible + 1
        self.start_offset = max(0, min(self.start_offset, max(0, len(options) - visible)))

        body = head[:]
        option_lines = []
        for offset, option in enumerate(options[self.start_offset:self.start_offset + visible]):
            index = self.start_offset + offset
            option_lines.append((index, self.option_line(
                index + 1, option["label"], option["extra"],
                index == self.start_selection, inner)))
        more = None
        if len(options) > visible:
            more = f"{p.dim}{fit(f'  … ещё {len(options) - visible}', inner)}{p.reset}"

        pad_top = max(0, (available - len(head) - len(option_lines)
                          - (1 if more else 0) - len(tail)) // 2)
        body = [""] * pad_top + head
        hits = []
        for index, line in option_lines:
            hits.append((1 + len(body), index))
            body.append(line)
        if more:
            body.append(more)
        body += tail
        self.start_hits = hits

        out = [self._top("ВРЕМЕННАЯ ПОЧТА", width, "меню")]
        out += self._panel_rows(body, inner, available)
        out.append(self._sep(width))
        hint = " ↑↓ Enter · d убрать · мышь: клик по пункту · Esc/q выход"
        out.append(self._row(f"{p.dim}{fit(hint, inner)}{p.reset}", width))
        out.append(self._bottom(width))
        return "\n".join(out)

    def render_provider(self, prompt_line=None) -> str:
        p = self.p
        width, height = self.size()
        inner = width - 2
        available = max(3, height - 4)

        providers = self.provider_list()
        if providers:
            self.provider_selection = max(0, min(self.provider_selection, len(providers) - 1))

        head = [
            "",
            f"{p.bold}{p.white}{fit('  Куда создать почту', inner)}{p.reset}",
            f"{p.dim}{fit('  Сервис определяет адрес и срок жизни ящика', inner)}{p.reset}",
            "",
        ]

        body = head[:]
        rows = []
        current = self.cfg.get("provider")
        for index, provider in enumerate(providers):
            if index:
                body.append("")
            selected = index == self.provider_selection
            label = f" {index + 1}. {provider.title}"
            if selected:
                title = f"{p.rev}{fit(label, inner)}{p.reset}"
            else:
                title = (f"{p.dim}{fit(f' {index + 1}.', 4)}{p.reset}"
                         f"{p.bold}{p.white}{fit(provider.title, inner - 4)}{p.reset}")
            rows.append((len(body), index))
            body.append(title)

            note_text = "  " + provider.note
            if provider.id == current:
                note_text += " · по умолчанию"
            body.append(f"{p.dim}{fit(note_text, inner)}{p.reset}")

        pad_top = max(0, (available - len(body)) // 2)
        body = [""] * pad_top + body
        self.provider_hits = [
            (1 + pad_top + row, index) for row, index in rows if row < available
        ]

        out = [self._top("ВРЕМЕННАЯ ПОЧТА", width, "выбор провайдера")]
        out += self._panel_rows(body, inner, available)
        out.append(self._sep(width))
        if self.error and time.time() - self.error_time < 15:
            out.append(self._row(f"{p.red}{fit(' ' + self.error, inner)}{p.reset}", width))
        else:
            hint = " ↑↓ выбор · Enter создать · мышь: клик по сервису · Esc назад"
            out.append(self._row(f"{p.dim}{fit(hint, inner)}{p.reset}", width))
        out.append(self._bottom(width))
        return "\n".join(out)

    def render_settings(self, prompt_line=None) -> str:
        p = self.p
        width, height = self.size()
        inner = width - 2
        available = max(3, height - 5)

        body = [""]
        for index, field in enumerate(SETTINGS_FIELDS):
            selected = index == self.settings_selection

            if field["kind"] == "action":
                text = f"  {field['label']}"
                if selected:
                    body.append(f"{p.rev}{fit(text, inner)}{p.reset}")
                else:
                    body.append(f"{p.magenta}{fit(text, inner)}{p.reset}")
                continue

            value_w = min(24, max(8, inner // 3))
            label_w = max(4, inner - value_w)
            label = f"  {field['label']}"
            value = self.setting_text(field)

            if selected:
                plain = fit(label, label_w) + fit(value, value_w, "right")
                body.append(f"{p.rev}{fit(plain, inner)}{p.reset}")
            else:
                body.append(f"{p.white}{fit(label, label_w)}{p.reset}"
                            f"{p.yellow}{fit(value, value_w, 'right')}{p.reset}")

        self.settings_hits = [
            (2 + index, index)
            for index in range(min(len(SETTINGS_FIELDS), max(0, available - 1)))
        ]

        field = SETTINGS_FIELDS[self.settings_selection]
        out = [self._top("НАСТРОЙКИ", width, "Esc — назад")]
        out += self._panel_rows(body, inner, available)
        out.append(self._sep(width))

        if prompt_line is not None:
            out.append(self._row(self._hint_row(inner, prompt_line), width))
        elif self.error and time.time() - self.error_time < 15:
            out.append(self._row(f"{p.red}{fit(' ' + self.error, inner)}{p.reset}", width))
        else:
            out.append(self._row(f"{p.dim}{fit(' ' + field['hint'], inner)}{p.reset}", width))

        hint = " ↑↓ выбор · Enter изменить · ←→ быстро · q/Esc назад"
        out.append(self._row(f"{p.dim}{fit(hint, inner)}{p.reset}", width))
        out.append(self._bottom(width))
        return "\n".join(out)

    def render_loading(self) -> str:
        p = self.p
        width, height = self.size()
        inner = width - 2
        area = max(1, height - 5)

        out = [self._top(f"ВРЕМЕННАЯ ПОЧТА · {self.provider_title()}", width), self._sep(width)]
        out.append(self._row(f"{p.bold}{p.white}{fit(' ' + self.loading_text, inner, 'center')}{p.reset}", width))
        for _ in range(area - 1):
            out.append(self._row(" " * inner, width))
        out.append(self._sep(width))
        out.append(self._row(f"{p.dim}{fit(' Подождите…', inner)}{p.reset}", width))
        out.append(self._bottom(width))
        return "\n".join(out)

    def render_inbox(self, prompt_line=None) -> str:
        p = self.p
        width, height = self.size()
        inner = width - 2
        rows = max(3, height - 9)
        self.inbox_hits = []

        out = [self._top(f"ВРЕМЕННАЯ ПОЧТА · {self.provider_title()}", width,
                         f"новых {self.unread_count()}"),
               self._row(self._mailbox_row(inner), width),
               self._sep(width)]

        if self.help_visible:
            out.append(self._row(f"{p.dim}{fit(' ГОРЯЧИЕ КЛАВИШИ', inner)}{p.reset}", width))
            out.append(self._sep(width))
            help_lines = [
                ("↑ / ↓", "выбрать письмо (работает колесо мыши)"),
                ("Enter / клик", "открыть письмо"),
                ("d / Del", "удалить выбранное письмо"),
                ("D", "удалить почту целиком"),
                ("c", "скопировать адрес ящика"),
                ("мышь", "клик по коду в письме — копировать"),
                ("Esc / m", "выйти в меню выбора ящика"),
                (",", "настройки: провайдер, время удаления, опрос, сортировка"),
                ("s / o", "быстро сменить сортировку / порядок"),
                ("t / p / i", "быстро изменить TTL / удаление после чтения / опрос"),
                ("n", "создать почту (с выбором сервиса)"),
                ("r", "обновить сейчас"),
                ("?", "эта справка"),
                ("q", "выход из программы"),
            ]
            def help_row(key, text):
                return (f"  {p.bold}{p.cyan}{fit(key, 18)}{p.reset}"
                        f"{fit(text, inner - 20)}")

            for key, text in help_lines[:rows]:
                out.append(self._row(help_row(key, text), width))
            for _ in range(rows - min(len(help_lines), rows)):
                out.append(self._row(" " * inner, width))
            out.append(self._sep(width))
        else:
            out.append(self._row(self._head_row(inner), width))
            out.append(self._sep(width))

            view = self.view()
            total = len(view)
            if self.selected < self.offset:
                self.offset = self.selected
            if self.selected >= self.offset + rows:
                self.offset = self.selected - rows + 1
            self.offset = max(0, min(self.offset, max(0, total - rows)))

            for i in range(rows):
                index = self.offset + i
                if total == 0 and i == 0:
                    empty = "Писем пока нет — адрес выше можно использовать для регистрации"
                    out.append(self._row(f"{p.dim}{fit(empty, inner, 'center')}{p.reset}", width))
                    continue
                if index >= total:
                    out.append(self._row(" " * inner, width))
                    continue
                msg = view[index]
                out.append(self._row(self._message_row(msg, index + 1, index == self.selected, inner), width))
            self.inbox_hits = [
                (5 + i, self.offset + i)
                for i in range(rows)
                if self.offset + i < total
            ]
            out.append(self._sep(width))

        out.append(self._row(self._status_row(inner), width))
        out.append(self._row(self._hint_row(inner, prompt_line), width))
        out.append(self._bottom(width))
        return "\n".join(out)

    def render_detail(self, prompt_line=None) -> str:
        p = self.p
        width, height = self.size()
        inner = width - 2
        msg = self.detail or {}

        out = [self._top("ПИСЬМО", width),
               self._sep(width)]
        self.detail_hits = []

        for label, value in (("От кого", sender_display(msg)),
                             ("Тема", msg.get("subject") or "(без темы)"),
                             ("Дата", fmt_time(msg.get("createdAt")))):
            shown = fit(value, inner - 11)
            content = (f"{p.dim} {label:<8}: {p.reset}"
                       f"{self.colorize_codes(shown, p.white)}")
            out.append(self._row(content, width))
            if label == "Тема":
                for start, end, token in code_spans(shown):
                    self.detail_hits.append((len(out) - 1, 12 + start, 12 + end, token))

        out.append(self._sep(width))

        body = wrap_block(message_body(msg), inner - 2)
        area = max(3, height - 9)
        max_scroll = max(0, len(body) - area)
        self.detail_scroll = max(0, min(self.detail_scroll, max_scroll))

        for i in range(area):
            index = self.detail_scroll + i
            text = body[index] if index < len(body) else ""
            for start, end, token in code_spans(text):
                self.detail_hits.append((len(out), 2 + start, 2 + end, token))
            out.append(self._row(f" {self.colorize_codes(fit(text, inner - 2))} ", width))

        out.append(self._sep(width))
        if prompt_line is not None:
            out.append(self._row(self._hint_row(inner, prompt_line), width))
        elif self.event and time.time() - self.event_time < 6:
            color = getattr(p, self.event_color, p.dim)
            out.append(self._row(f"{color}{fit(' ' + self.event, inner)}{p.reset}", width))
        else:
            hint = (f" ↑↓ прокрутка · клик по коду копирует · d удалить · "
                    f"D удалить почту · Esc назад · m меню · "
                    f"строка {self.detail_scroll + 1} из {max(1, len(body))}")
            out.append(self._row(f"{p.dim}{fit(hint, inner)}{p.reset}", width))
        out.append(self._bottom(width))
        return "\n".join(out)

    def render(self, prompt_line=None) -> str:
        if self.screen == "start":
            return self.render_start(prompt_line)
        if self.screen == "provider":
            return self.render_provider(prompt_line)
        if self.screen == "settings":
            return self.render_settings(prompt_line)
        if self.detail is not None:
            return self.render_detail(prompt_line)
        return self.render_inbox(prompt_line)

    def draw(self, prompt_line=None) -> None:
        frame = self.render(prompt_line)
        lines = frame.split("\n")
        buffer = ["\033[H"]
        for index, line in enumerate(lines):
            buffer.append(line + "\033[K")
            if index != len(lines) - 1:
                buffer.append("\n")
        sys.stdout.write("".join(buffer))
        sys.stdout.flush()

    # ---------------- режимы запуска ----------------

    def run_plain(self) -> None:
        """Линейный вывод для перенаправления в файл / пайп."""
        print("=" * 60)
        print("   ВРЕМЕННАЯ ПОЧТА")
        print("=" * 60)
        print("  Готовим ящик...")
        self.start()
        if not self.mailbox:
            print("  [!] не удалось подготовить ящик — выход")
            return

        ttl = int(self.cfg.get("message_ttl_minutes") or 0)
        after = int(self.cfg.get("delete_after_read_seconds") or 0)
        interval = int(self.cfg.get("poll_interval") or 5)

        print(f"  Провайдер: {self.provider_title()}")
        print(f"  Адрес: {self.mailbox['address']}")
        print(f"  TTL: {ttl} мин · после прочтения: {after} с · опрос: {interval} с")
        print("=" * 60)

        self.next_poll = 0.0
        while True:
            if time.time() >= self.next_poll:
                new = self.poll()
                self.next_poll = time.time() + interval
                for msg in reversed(new):
                    try:
                        full = self.provider_call(self.provider.read_message, self.ctx, msg["id"])
                    except (requests.exceptions.RequestException, ProviderError) as e:
                        print(f"  [!]  не удалось загрузить письмо: {e}")
                        continue
                    print_plain_message(full)
                    if after > 0:
                        self.pending_delete[msg["id"]] = time.time() + after
            self.enforce_retention()
            time.sleep(0.2)

    def run(self) -> None:
        if not self.interactive:
            self.run_plain()
            return

        enable_ansi()
        sys.stdout.write("\033[?1049h\033[?25l\033[2J\033[H")
        sys.stdout.flush()
        self.screen = "start"
        self.draw()
        try:
            last_draw = 0.0

            while self.running:
                now = time.time()
                if self.screen == "inbox" and now >= self.next_poll:
                    self.poll()
                    self.next_poll = now + max(2, int(self.cfg.get("poll_interval") or 5))

                if self.screen == "inbox":
                    self.enforce_retention()

                key = self.keys.read()
                if key is not None:
                    self.handle_key(key)

                if key is not None or time.time() - last_draw >= 0.3:
                    self.draw()
                    last_draw = time.time()
                time.sleep(0.04)
        finally:
            self.keys.restore()
            sys.stdout.write("\033[?25h\033[?1049l")
            sys.stdout.flush()

        print("\n  Программа завершена. До свидания!\n")


# -----------------------------------------------------------------------------
# Линейный вывод письма (режим без TTY)
# -----------------------------------------------------------------------------

def print_plain_message(msg: dict) -> None:
    print()
    print("=" * 60)
    print("  [!]  НОВОЕ ПИСЬМО!")
    print("=" * 60)
    print(f"  От кого : {sender_display(msg)}")
    print(f"  Тема    : {msg.get('subject') or '(без темы)'}")
    print(f"  Дата    : {fmt_time(msg.get('createdAt'))}")
    print("-" * 60)
    print(message_body(msg))
    print("=" * 60)
    print()


# -----------------------------------------------------------------------------
# Точка входа
# -----------------------------------------------------------------------------

def main() -> None:
    cfg = Config()
    app = App(cfg)
    try:
        app.run()
    except KeyboardInterrupt:
        pass
    except Exception as e:
        print(f"\n  [ОШИБКА] {e}\n")
        sys.exit(1)
    finally:
        app.keys.restore()


if __name__ == "__main__":
    main()
