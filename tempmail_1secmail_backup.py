# =============================================================================
# КЛИЕНТ ВРЕМЕННОЙ ПОЧТЫ — 1secmail.com
# =============================================================================
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

import time
import sys
import requests

# --- Константы ---
API_BASE   = "https://www.1secmail.com/api/v1/"  # Базовый URL API
POLL_DELAY = 5                                    # Интервал проверки почты (в секундах)

# Символы для декоративного оформления консоли
LINE_THICK = "=" * 60
LINE_THIN  = "-" * 60


def generate_email() -> str:
    """
    Запрашивает у API случайный почтовый адрес.
    Возвращает строку вида 'user@domain.com'.
    """
    params = {"action": "genRandomMailbox", "count": 1}
    response = requests.get(API_BASE, params=params, timeout=10)
    response.raise_for_status()          # Выбросит исключение при HTTP-ошибке
    email = response.json()[0]           # API возвращает список из N адресов
    return email


def get_messages(login: str, domain: str) -> list:
    """
    Возвращает список входящих писем для указанного ящика.
    Каждый элемент списка — словарь с полями: id, from, subject, date.
    """
    params = {"action": "getMessages", "login": login, "domain": domain}
    response = requests.get(API_BASE, params=params, timeout=10)
    response.raise_for_status()
    return response.json()


def read_message(login: str, domain: str, msg_id: int) -> dict:
    """
    Загружает полное содержимое конкретного письма по его ID.
    Возвращает словарь с полями: from, subject, date, textBody, htmlBody и др.
    """
    params = {
        "action": "readMessage",
        "login": login,
        "domain": domain,
        "id": msg_id
    }
    response = requests.get(API_BASE, params=params, timeout=10)
    response.raise_for_status()
    return response.json()


def print_message(msg: dict) -> None:
    """
    Красиво форматирует и выводит письмо в консоль.
    """
    print()
    print(LINE_THICK)
    print("  [!]  НОВОЕ ПИСЬМО!")
    print(LINE_THICK)
    print(f"  От кого : {msg.get('from', 'неизвестно')}")
    print(f"  Тема    : {msg.get('subject', '(без темы)')}")
    print(f"  Дата    : {msg.get('date', 'неизвестно')}")
    print(LINE_THIN)

    # Предпочитаем текстовое тело; если его нет — показываем заглушку
    body = msg.get("textBody", "").strip()
    if body:
        print(body)
    else:
        print("  (Письмо не содержит текстового содержимого)")

    print(LINE_THICK)
    print()


def main() -> None:
    """
    Главная функция: генерирует ящик, затем в бесконечном цикле
    проверяет входящую почту каждые POLL_DELAY секунд.
    """
    print()
    print(LINE_THICK)
    print("   ВРЕМЕННАЯ ПОЧТА — 1secmail.com")
    print(LINE_THICK)
    print("  Генерируем почтовый адрес...")

    # --- Генерация адреса ---
    try:
        email = generate_email()
    except Exception as e:
        print(f"\n  [ОШИБКА] Не удалось получить адрес: {e}")
        sys.exit(1)

    # Разбиваем адрес на login и domain для последующих запросов
    login, domain = email.split("@")

    print()
    print(f"  [OK]  Ваш временный адрес:")
    print()
    print(f"      >>> {email} <<<")
    print()
    print("  Скопируйте адрес выше и используйте его для регистрации.")
    print(f"  Входящие письма будут проверяться каждые {POLL_DELAY} сек.")
    print("  Для выхода нажмите Ctrl+C.")
    print(LINE_THICK)

    # Множество уже показанных ID писем — защита от дублей
    seen_ids: set = set()

    # --- Бесконечный цикл проверки почты ---
    while True:
        time.sleep(POLL_DELAY)

        try:
            messages = get_messages(login, domain)
        except requests.exceptions.RequestException as e:
            # Сетевые ошибки не останавливают программу — просто выводим предупреждение
            print(f"  [!]  Ошибка соединения: {e}. Повтор через {POLL_DELAY} сек...")
            continue

        # Перебираем каждое письмо из списка
        for msg_summary in messages:
            msg_id = msg_summary.get("id")

            # Пропускаем уже прочитанные письма
            if msg_id in seen_ids:
                continue

            # Запоминаем ID как обработанное
            seen_ids.add(msg_id)

            # Загружаем полное содержимое письма
            try:
                full_msg = read_message(login, domain, msg_id)
                print_message(full_msg)
            except requests.exceptions.RequestException as e:
                print(f"  [!]  Не удалось загрузить письмо #{msg_id}: {e}")


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
