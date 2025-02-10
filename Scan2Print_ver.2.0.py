import tkinter as tk
from tkinter import ttk
import threading
import win32com.client
import os
import datetime
from PIL import Image
import time
import socket
import pythoncom

LOG_TO_FILE = True
LOG_FILE = "scan2print.log"

if LOG_TO_FILE:
    with open(LOG_FILE, 'w', encoding='utf-8') as f:
        f.write("")

def log(message):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_line = f"{timestamp} - {message}"
    text_log.config(state=tk.NORMAL)
    text_log.insert(tk.END, log_line + "\n")
    text_log.see(tk.END)
    text_log.config(state=tk.DISABLED)
    if LOG_TO_FILE:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(log_line + "\n")

def scan_document():
    try:
        log("[INFO] Открываю диалог выбора сканера...")
        common_dialog = win32com.client.Dispatch("WIA.CommonDialog")
        device = common_dialog.ShowSelectDevice(1, False, True)
        if not device:
            log("[ERROR] Сканер не выбран!")
            return None

        try:
            device_name = device.Properties("Name").Value
        except Exception:
            device_name = "Неизвестно"
        log(f"[INFO] Выбран сканер: {device_name}")

        item = device.Items[1]

        log("[INFO] Настраиваю сканер на A4 и максимальное качество...")
        resolution = 600  # DPI – максимальное качество
        log(f"[INFO] Устанавливаю разрешение: {resolution} DPI")
        item.Properties["6147"].Value = resolution
        item.Properties["6148"].Value = resolution

        width_pixels = int(8.27 * resolution)
        height_pixels = int(11.69 * resolution)
        log(f"[INFO] Задаю область сканирования: {width_pixels}x{height_pixels} пикселей")
        item.Properties["6149"].Value = 0
        item.Properties["6150"].Value = 0
        item.Properties["6151"].Value = width_pixels
        item.Properties["6152"].Value = height_pixels

        try:
            log("[INFO] Устанавливаю цветной режим...")
            item.Properties["6146"].Value = 1  # 1 = цветной режим
        except Exception as e:
            log("[WARNING] Не удалось установить цветной режим: " + str(e))

        log("[INFO] Сканирую...")
        image = item.Transfer("{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}")

        timestamp = datetime.datetime.now().strftime("%d-%m-%y_%H.%M")
        filename = f"scan{timestamp}.jpg"
        image.SaveFile(filename)
        log(f"[SUCCESS] Скан готов. Файл сохранен: {filename}")
        return filename
    except Exception as e:
        log("[ERROR] Ошибка при сканировании: " + str(e))
        return None

def compress_image(input_file, output_file, quality=85):
    try:
        log(f"[INFO] Сжимаю изображение {input_file} с качеством {quality}...")
        img = Image.open(input_file)
        img.save(output_file, "JPEG", quality=quality)
        log(f"[SUCCESS] Сжатие завершено. Файл: {output_file}")
        return output_file
    except Exception as e:
        log("[ERROR] Ошибка при сжатии: " + str(e))
        return input_file

def raw_print(file_path, printer_ip, port=9100):
    try:
        log(f"[INFO] Отправляю файл {file_path} на печать через raw printing на {printer_ip}:{port}...")
        with open(file_path, 'rb') as f:
            data = f.read()
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.connect((printer_ip, port))
        s.sendall(data)
        s.close()
        log("[SUCCESS] Файл отправлен на печать!")
    except Exception as e:
        log("[ERROR] Ошибка при отправке на печать: " + str(e))

def process_copy():
    pythoncom.CoInitialize()  # Инициализация COM в этом потоке
    try:
        scanned_file = scan_document()
        if scanned_file:
            compressed_file = scanned_file.replace(".jpg", "_compressed.jpg")
            compressed_file = compress_image(scanned_file, compressed_file, quality=85)
            printer_ip = "192.168.1.199"  # IP принтера
            log(f"[INFO] Использую принтер с IP: {printer_ip}")
            raw_print(compressed_file, printer_ip)
    finally:
        pythoncom.CoUninitialize()  # Освобождение COM

def on_copy_button_click():
    copy_button.config(state=tk.DISABLED)
    threading.Thread(target=lambda: [process_copy(), copy_button.config(state=tk.NORMAL)]).start()

def copy_log():
    log_text = text_log.get("1.0", tk.END)
    root.clipboard_clear()
    root.clipboard_append(log_text)
    log("[INFO] Лог скопирован в буфер обмена.")

# Создаем графическое окно
root = tk.Tk()
root.title("Scan2Print")
root.geometry("600x400")

frame = ttk.Frame(root, padding=10)
frame.pack(fill=tk.BOTH, expand=True)

copy_button = ttk.Button(frame, text="Сделать серокопию", command=on_copy_button_click)
copy_button.pack(pady=10)

copy_log_button = ttk.Button(frame, text="Копировать лог", command=copy_log)
copy_log_button.pack(pady=5)

text_log = tk.Text(frame, wrap=tk.WORD, height=15, state=tk.DISABLED)
text_log.pack(fill=tk.BOTH, expand=True)

root.mainloop()


# КОД ПОСВЕЩАЮ РАБОТЕ В ИИП И МОЕМУ ЛЮБИМОМУ СМЕТНОМУ ОТДЕЛУ
# САНЯ КАК ВСЕГДА ПРИВЕТ Я НАДЕЮСЬ ТЫ СМОТРИШЬ КАЖДЫЙ МОЙ ЕБАНЫЙ КОМИТ НА ГИТЕ

"""
                           +═════════════════════════════════════════════════════════════════════════+
                           ║      ███▄ ▄███▓ ██ ▄█▀ █    ██  ██▓    ▄▄▄█████▓ ██▀███   ▄▄▄           ║
                           ║     ▓██▒▀█▀ ██▒ ██▄█▒  ██  ▓██▒▓██▒    ▓  ██▒ ▓▒▓██ ▒ ██▒▒████▄         ║
                           ║     ▓██    ▓██░▓███▄░ ▓██  ▒██░▒██░    ▒ ▓██░ ▒░▓██ ░▄█ ▒▒██  ▀█▄       ║
                           ║     ▒██    ▒██ ▓██ █▄ ▓▓█  ░██░▒██░    ░ ▓██▓ ░ ▒██▀▀█▄  ░██▄▄▄▄██      ║
                           ║     ▒██▒   ░██▒▒██▒ █▄▒▒█████▓ ░██████▒  ▒██▒ ░ ░██▓ ▒██▒ ▓█   ▓██▒     ║
                           ║     ░ ▒░   ░  ░▒ ▒▒ ▓▒░▒▓▒ ▒ ▒ ░ ▒░▓  ░  ▒ ░░   ░ ▒▓ ░▒▓░ ▒▒   ▓▒█░     ║
                           ║     ░  ░      ░░ ░▒ ▒░░░▒░ ░ ░ ░ ░ ▒  ░    ░      ░▒ ░ ▒░  ▒   ▒▒ ░     ║
                           ║     ░      ░   ░░ ░  ░░░ ░ ░   ░ ░     ░        ░░   ░   ░   ▒          ║
                           ║            ░   ░  ░      ░         ░  ░            ░           ░  ░     ║ 
                           ║                                                                         ║
                           +═════════════════════════════════════════════════════════════════════════+
                           ║                               By MKultra69                              ║
                           +═════════════════════════════════════════════════════════════════════════+
"""