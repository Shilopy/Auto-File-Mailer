import streamlit as st
import win32com.client
import pythoncom
import os
import json
import time
from datetime import datetime, timedelta
import logging
import pandas as pd
import psutil
import subprocess
import sys

# Настройка страницы Streamlit
st.set_page_config(
    page_title="Auto Sender Outlook",
    page_icon="📧",
    layout="wide"
)

# Создание папки для логов если не существует
if not os.path.exists("logs"):
    os.makedirs("logs")

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/sender.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# Инициализация сессии (убран scheduler_running)
if 'last_log_entries' not in st.session_state:
    st.session_state.last_log_entries = []

# --- ПЕРЕМЕЩЕННЫЕ ФУНКЦИИ УПРАВЛЕНИЯ СЕРВИСОМ НАЧАЛО ---
# Эти функции определены здесь, чтобы быть доступны для вызова в основном потоке выполнения
def is_service_running():
    """Проверяет, запущен ли сервис отправки"""
    for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
        try:
            # Проверяем, содержит ли командная строка процесса имя файла сервиса
            if proc.info['cmdline'] and 'sender_service.py' in ' '.join(proc.info['cmdline']):
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            # Игнорируем процессы, к которым нет доступа или которые уже завершились
            pass
    return False

def start_service():
    """Запускает сервис отправки в новом окне консоли"""
    try:
        # subprocess.CREATE_NEW_CONSOLE - создает новое окно консоли для сервиса
        # Для скрытого запуска можно использовать creationflags=subprocess.CREATE_NO_WINDOW (Python 3.7+)
        # или startupinfo (см. предыдущие ответы)
        subprocess.Popen([sys.executable, 'sender_service.py'],
                        creationflags=subprocess.CREATE_NEW_CONSOLE)
        st.sidebar.success("Сервис отправки запущен!")
        # Небольшая задержка, чтобы процесс успел запуститься
        time.sleep(0.5)
        # Перезагружаем страницу, чтобы обновить статус
        st.rerun()
    except Exception as e:
        st.sidebar.error(f"Ошибка запуска сервиса отправки: {e}")

def stop_service():
    """Останавливает сервис отправки"""
    try:
        stopped_any = False
        # Итерируемся по всем запущенным процессам
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                # Проверяем, содержит ли командная строка процесса имя файла сервиса
                if proc.info['cmdline'] and 'sender_service.py' in ' '.join(proc.info['cmdline']):
                    # Принудительно завершаем процесс
                    proc.terminate()
                    # Ждем завершения процесса до 3 секунд
                    proc.wait(timeout=3)
                    stopped_any = True
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess, subprocess.TimeoutExpired):
                # Игнорируем ошибки или процессы, которые не удалось завершить за таймаут
                pass
        if stopped_any:
            st.sidebar.success("Сервис отправки остановлен!")
        else:
            st.sidebar.info("Сервис отправки не найден или уже остановлен.")
        # Небольшая задержка
        time.sleep(0.5)
        # Перезагружаем страницу, чтобы обновить статус
        st.rerun()
    except Exception as e:
        st.sidebar.error(f"Ошибка остановки сервиса: {e}")
# --- ПЕРЕМЕЩЕННЫЕ ФУНКЦИИ УПРАВЛЕНИЯ СЕРВИСОМ КОНЕЦ ---

# Заголовок приложения
st.title("📧 Автоматическая рассылка файлов через Outlook Win32")
st.markdown("---")

# Боковая панель для навигации
page = st.sidebar.selectbox("Навигация", ["Конфигурация", "Отправка файлов", "Логи", "Инструкция"])

# Функция для загрузки конфигурации
def load_config():
    default_config = {
        "folder_path": "C:\\Files\\Reports\\",
        "schedule_times": ["16:00"], # Может не использоваться в новой логике сервиса, но оставлено
        "email_config": {
            "7210": "ivanov@ya.ru",
            "7220": "pupkin@ya.ru",
            "7230": "gorohov@ya.ru"
        },
        "sender_email": "your@email.ru",
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
                # Объединяем с дефолтной конфигурацией
                for key in default_config:
                    if key not in config:
                        config[key] = default_config[key]
                return config
        except Exception as e:
            st.error(f"Ошибка загрузки конфигурации: {str(e)}")
            logging.error(f"Ошибка загрузки конфигурации: {str(e)}")
            return default_config
    else:
        # Сохраняем дефолтную конфигурацию
        with open("config.json", 'w', encoding='utf-8') as f:
            json.dump(default_config, f, indent=2, ensure_ascii=False)
        return default_config

# Функция для сохранения конфигурации
def save_config(config):
    try:
        with open("config.json", 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        st.success("Конфигурация сохранена!")
        logging.info("Конфигурация сохранена")
    except Exception as e:
        st.error(f"Ошибка сохранения конфигурации: {str(e)}")
        logging.error(f"Ошибка сохранения конфигурации: {str(e)}")

# Функция для подключения к Outlook
def connect_outlook():
    try:
        # Инициализация COM библиотеки
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        st.success("✅ Успешное подключение к Outlook")
        logging.info("Успешное подключение к Outlook")
        return outlook
    except Exception as e:
        st.error(f"❌ Ошибка подключения к Outlook: {str(e)}")
        logging.error(f"Ошибка подключения к Outlook: {str(e)}")
        return None

# --- ОБНОВЛЕННАЯ ФУНКЦИЯ get_files_for_today для ПРЕДВАРИТЕЛЬНОГО ПРОСМОТРА ---
# (Используется ТОЛЬКО на странице "Отправка файлов" для показа файлов,
#  которые подходят по date_config НА ДАННЫЙ МОМЕНТ)
def get_files_for_today(folder_path):
    """
    Ищет файлы для отправки на основе настроек date_config (предварительный просмотр).
    Для каждого склада определяет целевую дату файла (сегодня + days_offset или send_on_friday).
    Ищет файлы с датой в формате YYYYMMDD и кодом склада в начале имени.
    """
    try:
        today_dt_obj = datetime.now() # Получаем объект datetime для проверки дня недели
        config = load_config() # Загружаем конфигурацию
        date_config = config.get('date_config', {}) # Получаем date_config
        email_config = config.get('email_config', {}) # Получаем email_config

        grouped_files = {}

        # Используем date_config для определения дат файлов для каждого склада
        for warehouse_code, date_info in date_config.items():
            # Получаем email для склада
            email = email_config.get(warehouse_code)
            # Пропускаем склад, если нет email
            if not email:
                logging.debug(f"Пропущен склад {warehouse_code}: нет email в конфигурации.")
                continue

            # Определяем смещение в зависимости от дня недели
            days_offset = date_info.get('days_offset', 0)
            # Проверяем, является ли сегодня пятницей (weekday() == 4)
            if today_dt_obj.weekday() == 4:  # Пятница
                days_offset = date_info.get('send_on_friday', days_offset)

            # Рассчитываем ЦЕЛЕВУЮ дату для поиска файлов
            target_date = today_dt_obj + timedelta(days=days_offset)
            # Формат даты в файле: YYYYMMDD (без точек)
            target_date_str = target_date.strftime('%Y%m%d')

            logging.debug(f"Предпросмотр для склада {warehouse_code}: поиск файлов с датой {target_date_str} (offset: {days_offset})")

            if not os.path.exists(folder_path):
                st.error(f"❌ Папка {folder_path} не существует")
                logging.error(f"Папка {folder_path} не существует")
                return {} # Если папка не существует, возвращаем пустой результат

            try:
                files = os.listdir(folder_path)
            except OSError as e:
                st.error(f"❌ Ошибка доступа к папке {folder_path}: {e}")
                logging.error(f"Ошибка доступа к папке {folder_path}: {e}")
                return {}

            # --- ЛОГИКА ГРУППИРОВКИ ПО СКЛАДУ ---
            # Ищем файлы, которые НАЧИНАЮТСЯ с кода склада и содержат целевую дату
            for file in files:
                if file.lower().endswith('.xlsx'):
                    # Проверяем, НАЧИНАЕТСЯ ли имя файла с кода склада и содержит ли ЦЕЛЕВУЮ дату
                    # Пример: файл "7210_20250815_..." должен подходить для склада "7210"
                    if file.startswith(f"{warehouse_code}_") and target_date_str in file:
                        # Если группа для склада еще не создана, создаем её
                        if warehouse_code not in grouped_files:
                            grouped_files[warehouse_code] = {
                                'email': email,
                                'files': []
                            }
                        # Добавляем файл в группу склада
                        grouped_files[warehouse_code]['files'].append(file)
                        logging.debug(f"Найден файл для склада {warehouse_code}: {file}")

        return grouped_files
    except Exception as e:
        st.error(f"❌ Ошибка поиска файлов: {str(e)}")
        logging.error(f"Ошибка поиска файлов: {str(e)}")
        return {}

# Функция для отправки email
def send_email(outlook, to_email, subject, body, attachments, folder_path):
    try:
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = to_email
        mail.Subject = subject
        mail.Body = body
        # Добавляем вложения
        for attachment in attachments:
            full_path = os.path.join(folder_path, attachment)
            if os.path.exists(full_path):
                mail.Attachments.Add(full_path)
            else:
                st.warning(f"⚠️ Файл не найден: {full_path}")
                logging.warning(f"Файл не найден: {full_path}")
        mail.Send()
        st.success(f"✅ Письмо отправлено на {to_email}")
        logging.info(f"Письмо отправлено на {to_email}")
        return True
    except Exception as e:
        st.error(f"❌ Ошибка отправки письма на {to_email}: {str(e)}")
        logging.error(f"Ошибка отправки письма на {to_email}: {str(e)}")
        return False

# Функция для отправки файлов (ручная отправка)
def send_files_now():
    try:
        st.info("🚀 Начало отправки файлов")
        logging.info("Начало отправки файлов")

        # Проверяем выходной день
        today = datetime.now()
        if today.weekday() in [5, 6]:  # Суббота=5, Воскресенье=6
            st.info("ℹ️ Сегодня выходной день. Отправка не производится.")
            logging.info("Сегодня выходной день. Отправка не производится.")
            return False

        # Загружаем конфигурацию
        config = load_config()

        # Получаем файлы для отправки (используем логику предварительного просмотра)
        grouped_files = get_files_for_today(config['folder_path'])

        if not grouped_files:
            st.info("ℹ️ Нет файлов для отправки сегодня")
            logging.info("Нет файлов для отправки сегодня")
            return False

        # Подключаемся к Outlook
        outlook = connect_outlook()
        if not outlook:
            return False

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
                st.success(f"✅ Файлы склада {warehouse_code} успешно отправлены на {data['email']}")
                logging.info(f"Файлы склада {warehouse_code} успешно отправлены на {data['email']}")
            else:
                st.error(f"❌ Ошибка отправки файлов склада {warehouse_code} на {data['email']}")
                logging.error(f"Ошибка отправки файлов склада {warehouse_code} на {data['email']}")

        st.success(f"🏁 Отправка завершена. Успешно отправлено: {success_count}/{len(grouped_files)}")
        logging.info(f"Отправка завершена. Успешно отправлено: {success_count}/{len(grouped_files)}")
        return True
    except Exception as e:
        st.error(f"❌ Критическая ошибка отправки: {str(e)}")
        logging.error(f"Критическая ошибка отправки: {str(e)}")
        return False

# Страница конфигурации
if page == "Конфигурация":
    st.header("⚙️ Конфигурация приложения")
    # Загружаем текущую конфигурацию
    config = load_config()
    # Форма конфигурации
    with st.form("config_form"):
        st.subheader("Основные настройки")
        folder_path = st.text_input("Путь к папке с файлами", value=config.get('folder_path', ''))
        # Ввод множества времени отправки (для совместимости/информации)
        st.subheader("Время отправки (информационно)")
        st.info("Сервис теперь работает в режиме мониторинга папки и отправляет файлы сразу при их появлении и соответствии критериям.")
        schedule_times_str = st.text_input(
            "Время отправки (через запятую, формат ЧЧ:ММ) - используется только для информации",
            value=', '.join(config.get('schedule_times', ['16:00'])),
            disabled=True # Сделаем поле неактивным, так как оно больше не используется сервисом
        )
        sender_email = st.text_input("Email отправителя", value=config.get('sender_email', ''))
        st.subheader("Конфигурация email адресов")
        st.write("Введите код склада и соответствующий email адрес:")
        # Создаем DataFrame для редактирования email конфигурации
        email_data = []
        for code, email in config.get('email_config', {}).items():
            email_data.append({"Код склада": code, "Email": email})
        email_df = pd.DataFrame(email_data)
        edited_email_df = st.data_editor(email_df, num_rows="dynamic", key="email_editor")
        st.subheader("Конфигурация дат файлов")
        st.write("Укажите количество дней к сегодняшней дате для каждого склада:")
        # Создаем DataFrame для редактирования date_config
        date_data = []
        for code, date_info in config.get('date_config', {}).items():
            date_data.append({
                "Код склада": code,
                "Дней к сегодняшней дате": date_info.get('days_offset', 0),
                "Отправка в пятницу": date_info.get('send_on_friday', 0)
            })
        date_df = pd.DataFrame(date_data)
        edited_date_df = st.data_editor(date_df, num_rows="dynamic", key="date_editor")
        # Кнопка сохранения
        submitted = st.form_submit_button("💾 Сохранить конфигурацию")
        if submitted:
            # Преобразуем DataFrame обратно в словарь для email
            email_config = {}
            for index, row in edited_email_df.iterrows():
                if pd.notna(row["Код склада"]) and pd.notna(row["Email"]):
                    email_config[str(row["Код склада"]).strip()] = str(row["Email"]).strip()
            # Преобразуем DataFrame обратно в словарь для date_config
            date_config = {}
            for index, row in edited_date_df.iterrows():
                if pd.notna(row["Код склада"]):
                    code = str(row["Код склада"]).strip()
                    try:
                        days_offset = int(row["Дней к сегодняшней дате"]) if pd.notna(row["Дней к сегодняшней дате"]) else 0
                        send_on_friday = int(row["Отправка в пятницу"]) if pd.notna(row["Отправка в пятницу"]) else 0
                    except ValueError:
                        days_offset = 0
                        send_on_friday = 0
                    date_config[code] = {"days_offset": days_offset, "send_on_friday": send_on_friday}
            # Обновляем конфигурацию
            config['folder_path'] = folder_path
            # config['schedule_times'] = schedule_times # Не обновляем, так как не используется сервисом
            config['sender_email'] = sender_email
            config['email_config'] = email_config
            config['date_config'] = date_config
            # Сохраняем конфигурацию
            save_config(config)

# Страница отправки файлов (обновлена)
elif page == "Отправка файлов":
    st.header("📤 Отправка файлов")
    # Текущее состояние
    st.subheader("Текущее состояние")
    col1, col2 = st.columns(2)
    with col1:
        # Убрана проверка st.session_state.scheduler_running
        # Отображаем статус внешнего сервиса
        if is_service_running(): # Теперь функция определена выше
             st.success("🟢 Сервис отправки (мониторинг) запущен")
        else:
             st.error("🔴 Сервис отправки (мониторинг) остановлен")

    with col2:
        config = load_config()
        st.info(f"📁 Папка для мониторинга: {config.get('folder_path', 'Не указана')}")
        st.info("⏱️ Интервал проверки: ~1 минута")
    # Разделитель
    st.markdown("---")
    # Кнопки управления
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🚀 Отправить сейчас", type="primary", use_container_width=True):
            send_files_now()
    # Проверка файлов для отправки (предварительный просмотр)
    st.subheader("📋 Файлы для отправки сегодня (предварительный просмотр)")
    st.info("Этот список показывает файлы, которые подходят по критериям `date_config` на текущий момент.")
    config = load_config()
    grouped_files = get_files_for_today(config['folder_path'])
    if grouped_files:
        for warehouse_code, data in grouped_files.items():
            with st.expander(f"📦 Склад {warehouse_code} → {data['email']}", expanded=True):
                st.write(f"**Email:** {data['email']}")
                st.write(f"**Файлы ({len(data['files'])}):**")
                for file in data['files']:
                    st.code(file)
    else:
        st.info("ℹ️ Нет файлов для отправки сегодня (предварительный просмотр)")

# Страница логов
elif page == "Логи":
    st.header("📋 Логи приложения")
    # Кнопки управления логами
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🔄 Обновить логи"):
            st.rerun()
    with col2:
        if st.button("🗑️ Очистить логи"):
            try:
                with open('logs/sender.log', 'w') as f:
                    f.write('')
                st.success("Логи очищены!")
                st.rerun()
            except Exception as e:
                st.error(f"Ошибка очистки логов: {str(e)}")
    # Отображение логов
    try:
        if os.path.exists('logs/sender.log'):
            with open('logs/sender.log', 'r', encoding='utf-8') as f:
                log_content = f.read()
            if log_content:
                # Отображаем последние 100 строк
                lines = log_content.split('\n')
                last_lines = lines[-100:] if len(lines) > 100 else lines
                log_text = '\n'.join(last_lines)
                st.text_area("Логи приложения (sender.log)", value=log_text, height=500, key="log_display")
            else:
                st.info("Логи пусты")
        else:
            st.info("Файл логов не найден")
    except Exception as e:
        st.error(f"Ошибка чтения логов: {str(e)}")

    # Отображение логов сервиса
    st.subheader("📝 Логи сервиса (service.log)")
    try:
        if os.path.exists('logs/service.log'):
            with open('logs/service.log', 'r', encoding='utf-8') as f:
                log_content_service = f.read()
            if log_content_service:
                # Отображаем последние 100 строк
                lines_service = log_content_service.split('\n')
                last_lines_service = lines_service[-100:] if len(lines_service) > 100 else lines_service
                log_text_service = '\n'.join(last_lines_service)
                st.text_area("Логи сервиса (service.log)", value=log_text_service, height=500, key="log_display_service")
            else:
                st.info("Логи сервиса пусты")
        else:
            st.info("Файл логов сервиса не найден")
    except Exception as e:
        st.error(f"Ошибка чтения логов сервиса: {str(e)}")


# Страница инструкции
elif page == "Инструкция":
    st.header("📖 Инструкция по использованию")
    st.subheader("1. Установка приложения")
    st.markdown("""
    1. Установите Python 3.8 или выше с [python.org](https://python.org)
    2. Установите необходимые библиотеки:
    ```bash
    pip install streamlit pywin32 schedule pandas pythoncom psutil
    ```
    3. Убедитесь, что Outlook установлен на компьютере
    """)
    st.subheader("2. Настройка конфигурации")
    st.markdown("""
    1. Перейдите на страницу "Конфигурация"
    2. Укажите путь к папке с файлами
    3. Укажите email адреса для каждого склада
    4. Настройте таблицу дат файлов: количество дней к сегодняшней дате для каждого склада
    5. Сохраните конфигурацию
    """)
    st.subheader("3. Запуск приложения")
    st.markdown("""
    1. Запустите сервис отправки в отдельном терминале:
    ```bash
    python sender_service.py
    ```
    2. Запустите интерфейс Streamlit в другом терминале:
    ```bash
    streamlit run auto_sender.py
    ```
    3. Откройте браузер по адресу, указанному в консоли
    4. На странице "Отправка файлов" нажмите "Отправить сейчас" для тестовой отправки
    """)
    st.subheader("4. Автоматическая отправка")
    st.markdown("""
    1. Сервис `sender_service.py` работает в **режиме мониторинга**.
    2. Он **постоянно** (примерно раз в минуту) проверяет папку на наличие новых файлов.
    3. Если находятся **новые** файлы, соответствующие правилам `date_config`, они **немедленно** отправляются.
    4. Отправка не производится в субботу и воскресенье.
    5. Сервис отслеживает уже отправленные файлы и **не отправляет их повторно**.
    """)
    st.subheader("5. Формат файлов")
    st.markdown("""
    Файлы должны иметь формат имени:
    ```
    [КодСклада]_[ГГГГММДД]_[ЛюбоеДругоеИмя].xlsx
    Пример: 7210_20250814_СД00-014490_ЗаданиеНаОтгрузку.XLSX
    ```
    *   `[КодСклада]`: Код склада (например, `7210`, `7220`).
    *   `[ГГГГММДД]`: Дата в формате `YYYYMMDD` (например, `20250813`).
    *   `[ЛюбоеДругоеИмя]`: Любое другое имя файла.
    *   `.xlsx`: Расширение файла Excel.
    """)
    st.subheader("6. Решение проблем")
    st.markdown("""
    **Outlook не подключается:**
    - Убедитесь, что Outlook запущен
    - Проверьте права доступа к Outlook
    - Перезапустите приложение
    **Файлы не находятся:**
    - Проверьте путь к папке в конфигурации
    - Убедитесь, что файлы имеют правильный формат имени и дату
    - Проверьте настройки `date_config`
    **Отправка не работает:**
    - Проверьте логи на странице "Логи"
    - Убедитесь, что сервис `sender_service.py` запущен
    - Убедитесь, что антивирус не блокирует отправку
    """)

# Фоновая задача для управления внешним сервисом (в боковой панели)
st.sidebar.markdown("---")
st.sidebar.subheader("⚙️ Автоматическая отправка")

# Статус сервиса отправки (использует перемещенные функции)
# Проверяем статус при каждой загрузке страницы
service_running = is_service_running()
if service_running:
    st.sidebar.success("🟢 Сервис отправки (мониторинг) запущен")
    if st.sidebar.button("⏹️ Остановить сервис отправки"):
        stop_service()
        # st.rerun() вызывается внутри stop_service
else:
    st.sidebar.error("🔴 Сервис отправки (мониторинг) остановлен")
    if st.sidebar.button("▶️ Запустить сервис отправки"):
        start_service()
        # st.rerun() вызывается внутри start_service

# Футер
st.markdown("---")
st.markdown("📧 Auto Sender Outlook - Автоматическая рассылка файлов через Win32 API")
