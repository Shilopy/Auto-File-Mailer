import win32com.client
import pythoncom
import os
import json
import schedule # Можно удалить, если не используем schedule для основного цикла
import time
from datetime import datetime, timedelta
import logging

# Создание папки для логов если не существует
if not os.path.exists("logs"):
    os.makedirs("logs")

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/service.log', encoding='utf-8'),
        logging.StreamHandler() # Убрать, если не хотите вывод в консоль сервиса
    ]
)

# --- ФУНКЦИИ ДЛЯ РАБОТЫ С ЖУРНАЛОМ ОТПРАВЛЕННЫХ ФАЙЛОВ ---
def get_sent_files_log_path():
    """Возвращает путь к файлу журнала отправленных файлов."""
    return os.path.join("logs", "sent_files.json")

def load_sent_files():
    """Загружает множество имен уже отправленных файлов."""
    log_path = get_sent_files_log_path()
    if not os.path.exists(log_path):
        return set()
    try:
        with open(log_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            if isinstance(data, list):
                return set(data)
            else:
                logging.warning(f"Неверный формат sent_files.json. Создается новый.")
                return set()
    except Exception as e:
        logging.error(f"Ошибка загрузки журнала отправленных файлов {log_path}: {e}")
        return set()

def save_sent_files(sent_files_set):
    """Сохраняет множество имен отправленных файлов."""
    log_path = get_sent_files_log_path()
    os.makedirs(os.path.dirname(log_path), exist_ok=True)
    try:
        with open(log_path, 'w', encoding='utf-8') as f:
            json.dump(list(sent_files_set), f, indent=2, ensure_ascii=False)
    except Exception as e:
        logging.error(f"Ошибка сохранения журнала отправленных файлов {log_path}: {e}")
# --- КОНЕЦ ФУНКЦИЙ ДЛЯ ЖУРНАЛА ---

def load_config():
    """Загружает конфигурацию из файла config.json."""
    default_config = {
        "folder_path": "C:\\Files\\Reports\\",
        "schedule_times": ["16:00"], # Может не использоваться в новой логике, но оставлено для совместимости
        "email_config": {
            "7210": "ibukhtoyarov@soudal.ru",
            "7220": "ibukhtoyarov@soudal.ru",
            "7230": "ibukhtoyarov@soudal.ru"
        },
        "sender_email": "ashilo@soudal.ru",
        "date_config": {
            "7210": {"days_offset": 1, "send_on_friday": 3},
            "7220": {"days_offset": 2, "send_on_friday": 4},
            "7230": {"days_offset": 2, "send_on_friday": 4}
        }
    }
    
    if os.path.exists("config.json"):
        try:
            with open("config.json", 'r', encoding='utf-8') as f:
                config = json.load(f)
                for key, value in default_config.items():
                    if key not in config:
                        config[key] = value
                return config
        except Exception as e:
            logging.error(f"Ошибка загрузки конфигурации: {str(e)}. Используются настройки по умолчанию.")
            return default_config
    else:
        logging.warning("Файл config.json не найден. Используются настройки по умолчанию.")
        return default_config

def connect_outlook():
    """Подключается к Outlook через COM."""
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        logging.info("Успешное подключение к Outlook")
        return outlook
    except Exception as e:
        logging.error(f"Ошибка подключения к Outlook: {str(e)}")
        return None

def get_files_for_sending(folder_path):
    """
    Ищет файлы для отправки на основе настроек date_config.
    Для каждого склада определяет целевую дату файла (сегодня + days_offset или send_on_friday).
    Ищет файлы с датой в формате YYYYMMDD и кодом склада в начале имени.
    Возвращает словарь файлов, подходящих для отправки на основе date_config.
    """
    try:
        today_dt_obj = datetime.now()
        config = load_config() # Всегда загружаем свежую конфигурацию
        date_config = config.get('date_config', {})
        email_config = config.get('email_config', {})

        # Словарь для хранения файлов, подходящих для отправки по date_config
        files_ready_to_send = {}

        # 1. Сначала определим для каждого склада, файлы с какой датой мы ищем
        target_dates_per_warehouse = {}
        for warehouse_code, date_info in date_config.items():
            email = email_config.get(warehouse_code)
            if not email:
                logging.debug(f"Пропущен склад {warehouse_code}: нет email в конфигурации.")
                continue

            days_offset = date_info.get('days_offset', 0)
            if today_dt_obj.weekday() == 4:  # Пятница
                days_offset = date_info.get('send_on_friday', days_offset)

            target_date = today_dt_obj + timedelta(days=days_offset)
            target_date_str = target_date.strftime('%Y%m%d')
            
            target_dates_per_warehouse[warehouse_code] = {
                'target_date_str': target_date_str,
                'email': email
            }
            logging.debug(f"Склад {warehouse_code}: ищу файлы с датой {target_date_str}")

        # 2. Теперь просканируем папку
        if not os.path.exists(folder_path):
            logging.error(f"Папка {folder_path} не существует")
            return {}

        try:
            all_files = os.listdir(folder_path)
        except OSError as e:
            logging.error(f"Ошибка доступа к папке {folder_path}: {e}")
            return {}

        logging.info(f"Найдено {len(all_files)} файлов в папке {folder_path}")

        # 3. Проверим каждый файл
        for file in all_files:
            if file.lower().endswith('.xlsx'):
                # Проверим, соответствует ли файл критериям какого-либо склада
                for warehouse_code, target_info in target_dates_per_warehouse.items():
                    target_date_str = target_info['target_date_str']
                    email = target_info['email']
                    
                    # Условие соответствия: имя начинается с кода склада и содержит целевую дату
                    if file.startswith(f"{warehouse_code}_") and target_date_str in file:
                        if warehouse_code not in files_ready_to_send:
                            files_ready_to_send[warehouse_code] = {
                                'email': email,
                                'files': []
                            }
                        files_ready_to_send[warehouse_code]['files'].append(file)
                        logging.debug(f"Файл {file} подходит для склада {warehouse_code}")
                        # Файл может подходить только одному складу из-за уникального префикса, выходим
                        break 

        return files_ready_to_send

    except Exception as e:
        logging.error(f"Ошибка поиска файлов для отправки: {str(e)}")
        return {}

def send_email(outlook, to_email, subject, body, attachments, folder_path):
    """Отправляет email через Outlook с вложениями."""
    try:
        mail = outlook.CreateItem(0)
        mail.To = to_email
        mail.Subject = subject
        mail.Body = body
        
        for attachment in attachments:
            full_path = os.path.join(folder_path, attachment)
            if os.path.exists(full_path):
                mail.Attachments.Add(full_path)
            else:
                logging.warning(f"Файл не найден: {full_path}")
        
        mail.Send()
        logging.info(f"Письмо отправлено на {to_email}")
        return True
    except Exception as e:
        logging.error(f"Ошибка отправки письма на {to_email}: {str(e)}")
        return False

def monitor_and_send():
    """
    Основная функция мониторинга папки и отправки новых файлов.
    Эта функция будет вызываться регулярно в цикле.
    """
    try:
        logging.info("--- Начало цикла мониторинга ---")
        
        # Проверяем выходной день
        today = datetime.now()
        if today.weekday() in [5, 6]:  # Суббота=5, Воскресенье=6
            logging.info("Сегодня выходной день. Мониторинг приостановлен.")
            return

        # Загружаем конфигурацию
        config = load_config()
        folder_path = config['folder_path']
        
        # 1. Получаем список файлов, которые нужно отправить (по date_config)
        files_ready_to_send = get_files_for_sending(folder_path)

        if not files_ready_to_send:
             logging.info("Нет файлов, подходящих для отправки по критериям date_config.")
             return

        # 2. Загружаем список уже отправленных файлов
        previously_sent_files = load_sent_files()
        logging.debug(f"Загружено {len(previously_sent_files)} ранее отправленных файлов.")

        # 3. Определяем новые файлы для отправки
        new_files_to_send = {}
        total_files_found = 0
        total_new_files = 0
        for warehouse_code, data in files_ready_to_send.items():
            total_files_found += len(data['files'])
            # Фильтруем список файлов для этого склада, исключая уже отправленные
            new_files_for_warehouse = [f for f in data['files'] if f not in previously_sent_files]
            total_new_files += len(new_files_for_warehouse)
            
            if new_files_for_warehouse:
                new_files_to_send[warehouse_code] = {
                    'email': data['email'],
                    'files': new_files_for_warehouse
                }
        
        logging.info(f"Файлов по критериям: {total_files_found}. Новых файлов: {total_new_files}.")

        if not new_files_to_send:
            logging.info("Нет новых файлов для отправки.")
            return

        # 4. Подключаемся к Outlook
        outlook = connect_outlook()
        if not outlook:
            logging.error("Не удалось подключиться к Outlook. Повторная попытка через интервал.")
            return

        # 5. Отправляем письма
        success_count = 0
        newly_sent_files = set() # Собираем файлы, отправленные в этом цикле
        
        for warehouse_code, data in new_files_to_send.items():
            if not data['files']: # На всякий случай
                 continue

            subject = f"Отчеты склада {warehouse_code} за {today.strftime('%d.%m.%Y')}"
            body = f"Во вложении отчеты склада {warehouse_code} за {today.strftime('%d.%m.%Y')}"

            success = send_email(
                outlook,
                data['email'],
                subject,
                body,
                data['files'],
                folder_path
            )
            
            if success:
                success_count += 1
                logging.info(f"Письмо для склада {warehouse_code} успешно отправлено на {data['email']} ({len(data['files'])} файлов)")
                newly_sent_files.update(data['files']) # Добавляем в список отправленных
            else:
                logging.error(f"Ошибка отправки письма для склада {warehouse_code} на {data['email']}")

        logging.info(f"Цикл мониторинга завершен. Успешно отправлено: {success_count}/{len(new_files_to_send)} складов.")

        # 6. Обновляем журнал отправленных файлов
        if newly_sent_files:
            updated_sent_files = previously_sent_files.union(newly_sent_files)
            save_sent_files(updated_sent_files)
            logging.info(f"Журнал отправленных файлов обновлен. Добавлено {len(newly_sent_files)} файлов.")
            
    except Exception as e:
        logging.error(f"Критическая ошибка в цикле мониторинга: {str(e)}")


def main():
    """Главная функция сервиса - запуск цикла мониторинга."""
    logging.info("Сервис автоматической отправки (режим мониторинга) запущен")
    
    # Можно загрузить начальную конфигурацию для лога
    config = load_config()
    logging.info(f"Мониторинг папки: {config.get('folder_path', 'Не указана')}")
    logging.info("Проверка новых файлов будет выполняться каждую минуту.")

    # Бесконечный цикл мониторинга
    while True:
        try:
            monitor_and_send()
        except Exception as e:
            logging.error(f"Неожиданная ошибка в основном цикле: {e}")
        
        # Ждем 60 секунд перед следующей проверкой
        # Это интервал мониторинга. Можно сделать короче (30 сек) или длиннее (5 мин)
        time.sleep(60) 

if __name__ == "__main__":
    # Убедимся, что папка логов существует при запуске
    if not os.path.exists("logs"):
        os.makedirs("logs")
    main()
