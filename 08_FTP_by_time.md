Для загрузки файлов на FTP-сервер по расписанию можно использовать Python и модуль `ftplib` в сочетании с планировщиком, например:  

- `cron` (для Linux)  
- `systemd` (Linux)  
- `Task Scheduler` (Windows)  

### **1. Установка зависимостей**
Python стандартный, дополнительные библиотеки не нужны.

### **2. Скрипт для загрузки файлов**
```python
from ftplib import FTP
import os

# Настройки FTP
FTP_HOST = "ftp.example.com"
FTP_USER = "your_username"
FTP_PASS = "your_password"
REMOTE_DIR = "/path/on/server"
LOCAL_DIR = "/path/to/local/files"

def upload_files():
    try:
        ftp = FTP(FTP_HOST)
        ftp.login(FTP_USER, FTP_PASS)
        ftp.cwd(REMOTE_DIR)

        for filename in os.listdir(LOCAL_DIR):
            local_path = os.path.join(LOCAL_DIR, filename)
            if os.path.isfile(local_path):
                with open(local_path, "rb") as file:
                    ftp.storbinary(f"STOR {filename}", file)
                print(f"Uploaded: {filename}")

        ftp.quit()
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    upload_files()
```

---

### **3. Запуск по расписанию**
#### **Linux (cron)**
1. Открываем crontab:  
   ```
   crontab -e
   ```
2. Добавляем задачу, например, запуск каждый день в 02:00:  
   ```
   0 2 * * * /usr/bin/python3 /path/to/ftp_upload.py
   ```

#### **Linux (systemd)**
Создаём два файла в `/etc/systemd/system/`:

**`ftp-upload.service`**
```ini
[Unit]
Description=FTP Upload Service
After=network.target

[Service]
ExecStart=/usr/bin/python3 /path/to/ftp_upload.py
Restart=always
User=your_user

[Install]
WantedBy=multi-user.target
```

**`ftp-upload.timer`**
```ini
[Unit]
Description=Run FTP upload daily

[Timer]
OnCalendar=*-*-* 02:00:00
Persistent=true

[Install]
WantedBy=timers.target
```

Затем выполняем:
```sh
sudo systemctl enable --now ftp-upload.timer
```

#### **Windows (Task Scheduler)**
1. Открыть "Планировщик заданий".  
2. Создать новую задачу → "Действие" → "Запуск программы".  
3. В поле "Программа/скрипт" указать путь к `python.exe`, а в "Аргументы" — путь к `ftp_upload.py`.  
4. Настроить расписание (например, ежедневно в 02:00).  

---

Такой способ позволит автоматизировать загрузку файлов на сервер по расписанию. Если нужны доработки (например, логирование, фильтрация файлов), можно дополнить скрипт.
