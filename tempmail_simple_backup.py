# =============================================================================
# КЛИЕНТ ВРЕМЕННОЙ ПОЧТЫ — mail.tm
# =============================================================================
#
# ВАЖНО: прежний сервис 1secmail.com закрылся — его API отвечает 404,
#        а TLS-сертификат домена больше не соответствует имени хоста.
#        Поэтому клиент переведён на mail.tm (бесплатно, без API-ключа).
#
# ИНСТРУКЦИЯ ПО СБОРКЕ В .EXE (Windows, PowerShell или cmd):
#
# Шаг 1. Убедитесь, что Python установлен. Проверьте версию:
#         python --version
#
# Шаг 2. Установите необходимые библиотеки:
#         pip install requests pyinstaller
#
# Шаг 3. Перейдите в папку, где лежит этот файл (tempmail.py):
#         cd C:\путь\до\папки\с\файлом
#
# Шаг 4. Соберите .exe одной командой:
#         pyinstaller --onefile --console --name TempMail tempmail.py
#
# Шаг 5. Готовый файл появится в папке dist\TempMail.exe
#         Запустите его двойным кликом или через консоль:
#         dist\TempMail.exe
#
# Примечания:
#   --onefile  — всё в одном исполняемом файле
#   --console  — оставить консольное окно (нужно для этой программы)
#   --name     — имя выходного .exe файла
#
# =============================================================================

import html as html_lib
import random
import re
import string
import sys
import time

import requests

# Чтобы строки сразу попадали в файл/пайп при перенаправлении вывода
try:
    sys.stdout.reconfigure(line_buffering=True, encoding="utf-8")
except (AttributeError, ValueError):
    pass

# --- Константы ---
API_BASE   = "https://api.mail.tm"   # Базовый URL API mail.tm
POLL_DELAY = 5                       # Интервал проверки почты (в секундах)

# Символы для декоративного оформления консоли
LINE_THICK = "=" * 60
LINE_THIN  = "-" * 60


# -----------------------------------------------------------------------------
# Вспомогательные функции
# -----------------------------------------------------------------------------

def _random_string(length: int) -> str:
    """Случайная строка из строчных латинских букв и цифр."""
    alphabet = string.ascii_lowercase + string.digits
    return "".join(random.choices(alphabet, k=length))


def _request(method: str, path: str, **kwargs) -> requests.Response:
    """
    Единая точка выполнения запросов: таймауты, повторы при 429/5xx.
    Повышает устойчивость к rate-limit сервиса (лимит около 8 запросов/сек).
    """
    url = API_BASE + path
    kwargs.setdefault("timeout", 15)

    last_error = None
    for attempt in range(3):
        try:
            response = requests.request(method, url, **kwargs)
        except requests.exceptions.RequestException as e:
            last_error = e
            time.sleep(1 + attempt)
            continue

        # 429 — слишком много запросов, сервер сам сообщает паузу
        if response.status_code == 429:
            try:
                wait = float(response.headers.get("Retry-After", 2))
            except ValueError:
                wait = 2.0
            time.sleep(min(wait, 10))
            last_error = requests.HTTPError(f"429 Too Many Requests ({url})")
            continue

        # 5xx — временная проблема на стороне сервиса
        if response.status_code >= 500:
            last_error = requests.HTTPError(f"{response.status_code} Server Error ({url})")
            time.sleep(1 + attempt)
            continue

        return response

    raise last_error if last_error else requests.RequestException(f"Запрос не удался: {url}")


def _auth_headers(token: str) -> dict:
    """Заголовки авторизации по токену."""
    return {"Authorization": f"Bearer {token}", "Accept": "application/json"}


def _as_list(payload) -> list:
    """
    API умеет отвечать как Hydra-коллекцией ({"hydra:member": [...]}),
    так и обычным JSON-массивом — приводим оба варианта к списку.
    """
    if isinstance(payload, list):
        return payload
    if isinstance(payload, dict):
        return payload.get("hydra:member", []) or []
    return []


# -----------------------------------------------------------------------------
# Работа с API
# -----------------------------------------------------------------------------

def create_mailbox() -> dict:
    """
    Создаёт новый временный ящик и возвращает словарь:
        {"address": ..., "password": ..., "token": ...}
    mail.tm не выдаёт адрес «просто так» — нужно зарегистрировать аккаунт
    и получить JWT-токен для чтения почты.
    """
    # 1. Получаем список доступных доменов и берём первый активный
    response = _request("GET", "/domains")
    response.raise_for_status()
    domains = _as_list(response.json())
    if not domains:
        raise RuntimeError("Сервис не вернул ни одного доступного домена")

    domain = domains[0]["domain"]
    address = f"{_random_string(10)}@{domain}"
    password = _random_string(16) + "Aa1!"

    # 2. Регистрируем аккаунт
    response = _request(
        "POST",
        "/accounts",
        json={"address": address, "password": password},
        headers={"Content-Type": "application/json", "Accept": "application/json"},
    )
    if response.status_code not in (200, 201):
        raise RuntimeError(
            f"Не удалось создать ящик (HTTP {response.status_code}): {response.text[:200]}"
        )

    # 3. Получаем токен доступа
    token = login(address, password)
    return {"address": address, "password": password, "token": token}


def login(address: str, password: str) -> str:
    """Возвращает новый JWT-токен для указанного ящика."""
    response = _request(
        "POST",
        "/token",
        json={"address": address, "password": password},
        headers={"Content-Type": "application/json", "Accept": "application/json"},
    )
    if response.status_code != 200:
        raise RuntimeError(
            f"Не удалось авторизоваться (HTTP {response.status_code}): {response.text[:200]}"
        )
    return response.json()["token"]


def get_messages(token: str) -> list:
    """
    Возвращает список входящих писем (краткие карточки).
    Поля: id, from{address,name}, subject, intro, seen, createdAt.
    """
    response = _request("GET", "/messages", headers=_auth_headers(token))
    response.raise_for_status()
    return _as_list(response.json())


def read_message(token: str, msg_id: str) -> dict:
    """
    Загружает полное содержимое письма по его ID.
    Поля: from, to, subject, text, html[], createdAt и др.
    """
    response = _request("GET", f"/messages/{msg_id}", headers=_auth_headers(token))
    response.raise_for_status()
    return response.json()


# -----------------------------------------------------------------------------
# Вывод в консоль
# -----------------------------------------------------------------------------

def _html_to_text(raw_html: str) -> str:
    """
    Грубое, но достаточное преобразование HTML-письма в читаемый текст:
    убираем скрипты/стили, теги, раскрываем HTML-сущности.
    """
    text = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", raw_html)
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"(?i)</(p|div|tr|li|h[1-6])>", "\n", text)
    text = re.sub(r"<[^>]+>", "", text)
    text = html_lib.unescape(text)
    text = text.replace("\r\n", "\n").replace("\xa0", " ")
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _sender_name(msg: dict) -> str:
    """Возвращает 'Имя <адрес>' либо просто адрес отправителя."""
    sender = msg.get("from") or {}
    if isinstance(sender, str):
        return sender
    address = sender.get("address", "неизвестно")
    name = (sender.get("name") or "").strip()
    return f"{name} <{address}>" if name else address


def print_message(msg: dict) -> None:
    """
    Красиво форматирует и выводит письмо в консоль.
    """
    print()
    print(LINE_THICK)
    print("  [!]  НОВОЕ ПИСЬМО!")
    print(LINE_THICK)
    print(f"  От кого : {_sender_name(msg)}")
    print(f"  Тема    : {msg.get('subject') or '(без темы)'}")
    print(f"  Дата    : {msg.get('createdAt', 'неизвестно')}")
    print(LINE_THIN)

    # Предпочитаем текстовое тело; если его нет — конвертируем HTML
    body = (msg.get("text") or "").strip()
    if not body:
        html_parts = msg.get("html") or []
        if isinstance(html_parts, str):
            html_parts = [html_parts]
        body = _html_to_text("\n".join(html_parts))

    if body:
        print(body)
    else:
        print("  (Письмо не содержит текстового содержимого)")

    print(LINE_THICK)
    print()


# -----------------------------------------------------------------------------
# Главный цикл
# -----------------------------------------------------------------------------

def main() -> None:
    """
    Главная функция: создаёт ящик, затем в бесконечном цикле
    проверяет входящую почту каждые POLL_DELAY секунд.
    """
    print()
    print(LINE_THICK)
    print("   ВРЕМЕННАЯ ПОЧТА — mail.tm")
    print(LINE_THICK)
    print("  Генерируем почтовый адрес...")

    # --- Создание ящика ---
    try:
        mailbox = create_mailbox()
    except Exception as e:
        print(f"\n  [ОШИБКА] Не удалось получить адрес: {e}")
        sys.exit(1)

    email = mailbox["address"]
    password = mailbox["password"]
    token = mailbox["token"]

    print()
    print("  [OK]  Ваш временный адрес:")
    print()
    print(f"      >>> {email} <<<")
    print()
    print("  Скопируйте адрес выше и используйте его для регистрации.")
    print(f"  Входящие письма будут проверяться каждые {POLL_DELAY} сек.")
    print("  Пароль ящика (на случай ручного входа на mail.tm):")
    print(f"      {password}")
    print("  Для выхода нажмите Ctrl+C.")
    print(LINE_THICK)

    # Множество уже показанных ID писем — защита от дублей
    seen_ids: set = set()

    # --- Бесконечный цикл проверки почты ---
    while True:
        time.sleep(POLL_DELAY)

        try:
            messages = get_messages(token)
        except requests.HTTPError as e:
            # 401 — истёк токен: перелогиниваемся и продолжаем работу
            if "401" in str(e):
                try:
                    token = login(email, password)
                    print("  [i]  Токен обновлён, продолжаем проверку.")
                except Exception as relogin_error:
                    print(f"  [!]  Не удалось обновить токен: {relogin_error}")
            else:
                print(f"  [!]  Ошибка сервера: {e}. Повтор через {POLL_DELAY} сек...")
            continue
        except requests.exceptions.RequestException as e:
            # Сетевые ошибки не останавливают программу
            print(f"  [!]  Ошибка соединения: {e}. Повтор через {POLL_DELAY} сек...")
            continue

        # Перебираем каждое письмо из списка (API отдаёт новые первыми)
        for msg_summary in reversed(messages):
            msg_id = msg_summary.get("id")
            if msg_id in seen_ids:
                continue

            # Запоминаем ID как обработанное
            seen_ids.add(msg_id)

            # Загружаем полное содержимое письма
            try:
                full_msg = read_message(token, msg_id)
                print_message(full_msg)
            except requests.exceptions.RequestException as e:
                print(f"  [!]  Не удалось загрузить письмо {msg_id}: {e}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        # Перехватываем Ctrl+C и выводим аккуратное прощальное сообщение
        print()
        print(LINE_THIN)
        print("  Программа завершена. До свидания!")
        print(LINE_THIN)
        print()
        sys.exit(0)
