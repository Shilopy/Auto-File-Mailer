import re
import win32com.client
import pythoncom
import os
import json
import schedule
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
        logging.StreamHandler()
    ]
)

def load_config():
    """Загружает конфигурацию из файла config.json."""
    default_config = {
        "folder_path": "C:\\Files\\Reports\\",
        "schedule_times": ["16:00"], # Множество времени для отправки
        "email_config": {
            "7210": "ivanov@ya.ru",
            "7220": "pupkin@ya.ru",
            "7230": "gorohov@ya.ru"
        },
        "sender_email": "ashilo@soudal.ru",
        "date_config": { # Новая конфигурация дат
            "7210": {"days_offset": 1, "send_on_friday": 3},
            "7220": {"days_offset": 2, "send_on_friday": 4},
            "7230": {"days_offset": 2, "send_on_friday": 4}
        }
    }
    
    if os.path.exists("config.json"):
        try:
            with open("config.json", 'r', encoding='utf-8') as f:
                config = json.load(f)
                # Убедиться, что все ключи из default_config присутствуют
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
        pythoncom.CoInitialize() # Инициализация COM библиотеки
        outlook = win32com.client.Dispatch("Outlook.Application")
        logging.info("Успешное подключение к Outlook")
        return outlook
    except Exception as e:
        logging.error(f"Ошибка подключения к Outlook: {str(e)}")
        return None

def get_files_for_today(folder_path):
    """
    Ищет файлы для отправки на основе настроек date_config.
    Для каждого склада определяет целевую дату файла (сегодня + days_offset или send_on_friday).
    Ищет файлы с датой в формате YYYYMMDD и кодом склада в начале имени.
    """
    try:
        today_dt_obj = datetime.now()
        config = load_config()
        date_config = config.get('date_config', {})
        email_config = config.get('email_config', {})

        grouped_files = {}

        # Используем date_config для определения дат файлов для каждого склада
        for warehouse_code, date_info in date_config.items():
            email = email_config.get(warehouse_code)
            if not email:
                logging.debug(f"Пропущен склад {warehouse_code}: нет email в конфигурации.")
                continue

            # --- ЛОГИКА ОПРЕДЕЛЕНИЯ ЦЕЛЕВОЙ ДАТЫ ---
            # Определяем смещение в зависимости от дня недели
            days_offset = date_info.get('days_offset', 0)
            # Проверяем, является ли сегодня пятницей (weekday() == 4)
            if today_dt_obj.weekday() == 4:  # Пятница
                days_offset = date_info.get('send_on_friday', days_offset)
            # --- КОНЕЦ ЛОГИКИ ОПРЕДЕЛЕНИЯ ЦЕЛЕВОЙ ДАТЫ ---

            # Рассчитываем ЦЕЛЕВУЮ дату для поиска файлов
            target_date = today_dt_obj + timedelta(days=days_offset)
            # Формат даты в файле: YYYYMMDD (без точек)
            target_date_str = target_date.strftime('%Y%m%d')

            logging.debug(f"Склад {warehouse_code}: поиск файлов с датой {target_date_str} (offset: {days_offset})")

            if not os.path.exists(folder_path):
                logging.error(f"Папка {folder_path} не существует")
                return {} # Если папка не существует, возвращаем пустой результат

            try:
                files = os.listdir(folder_path)
            except OSError as e:
                logging.error(f"Ошибка доступа к папке {folder_path}: {e}")
                return {}

            # --- ЛОГИКА ГРУППИРОВКИ ПО СКЛАДУ ---
            # Ищем файлы, которые НАЧИНАЮТСЯ с кода склада и содержат целевую дату
            for file in files:
                if file.lower().endswith('.xlsx'):
                    # Проверяем, НАЧИНАЕТСЯ ли имя файла с кода склада и содержит ли ЦЕЛЕВУЮ дату
                    # Пример: файл "7210_20250815_..." должен подходить для склада "7210"
                    if file.startswith(f"{warehouse_code}_") and target_date_str in file:
                        if warehouse_code not in grouped_files:
                            grouped_files[warehouse_code] = {
                                'email': email,
                                'files': []
                            }
                        grouped_files[warehouse_code]['files'].append(file)
                        logging.debug(f"Найден файл для склада {warehouse_code}: {file}")
            # --- КОНЕЦ ЛОГИКИ ГРУППИРОВКИ ПО СКЛАДУ ---

        return grouped_files
    except Exception as e:
        logging.error(f"Ошибка поиска файлов: {str(e)}")
        return {}

def send_email(outlook, to_email, subject, body, attachments, folder_path):
    """Отправляет email через Outlook с вложениями."""
    try:
        mail = outlook.CreateItem(0) # 0 = olMailItem
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

def send_files_job():
    """Основная функция отправки файлов, запускаемая по расписанию."""
    try:
        logging.info("Начало автоматической отправки файлов")
        
        # Проверяем выходной день
        today = datetime.now()
        if today.weekday() in [5, 6]:  # Суббота=5, Воскресенье=6
            logging.info("Сегодня выходной день. Отправка не производится.")
            return
        
        config = load_config()
        grouped_files = get_files_for_today(config['folder_path'])
        
        if not grouped_files:
            logging.info("Нет файлов для отправки сегодня")
            return
        
        outlook = connect_outlook()
        if not outlook:
            return # Если не удалось подключиться, выходим
        
        success_count = 0
        for warehouse_code, data in grouped_files.items():
            subject = f"Отчеты склада {warehouse_code} за {today.strftime('%d.%m.%Y')}"
            body = f"Во вложении отчеты склада {warehouse_code} за {today.strftime('%d.%m.%Y')}"
            
            success = send_email(
                outlook,
                data['email'],
                subject,
                body,
                data['files'],
                config['folder_path']
            )
            
            if success:
                success_count += 1
                logging.info(f"Файлы склада {warehouse_code} успешно отправлены на {data['email']}")
            else:
                logging.error(f"Ошибка отправки файлов склада {warehouse_code} на {data['email']}")
        
        logging.info(f"Автоматическая отправка завершена. Успешно отправлено: {success_count}/{len(grouped_files)}")
        
    except Exception as e:
        logging.error(f"Критическая ошибка автоматической отправки: {str(e)}")

def main():
    """Главная функция сервиса."""
    logging.info("Сервис автоматической отправки запущен")
    
    # Загружаем конфигурацию
    config = load_config()
    schedule_times = config.get('schedule_times', ["16:00"])
    
    # Проверка формата времени
    valid_times = []
    for t in schedule_times:
        try:
            datetime.strptime(t, '%H:%M')
            valid_times.append(t)
        except ValueError:
            logging.warning(f"Неверный формат времени '{t}'. Пропущено. Используйте ЧЧ:ММ.")
    
    if not valid_times:
        logging.error("Не указано корректное время отправки. Используется время по умолчанию 16:00.")
        valid_times = ["16:00"]

    # Планируем задачу для каждого времени из конфигурации
    for schedule_time in valid_times:
        schedule.every().day.at(schedule_time).do(send_files_job)
        logging.info(f"Задача отправки запланирована на {schedule_time}")
    
    logging.info(f"Отправка запланирована на {', '.join(valid_times)} (по будням)")
    
    # Бесконечный цикл проверки задач
    while True:
        schedule.run_pending()
        time.sleep(60) # Проверка каждую минуту

if __name__ == "__main__":
    main()
