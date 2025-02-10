import win32com.client
import os
import datetime
from PIL import Image
import time
import socket

def print_banner():
    banner = """
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
                           ║                               MKultra69                                 ║
                           +═════════════════════════════════════════════════════════════════════════+
    """
    print(banner)

def scan_document():
    try:
        print("[INFO] Окно выбора...")
        common_dialog = win32com.client.Dispatch("WIA.CommonDialog")
        device = common_dialog.ShowSelectDevice(1, False, True)
        if not device:
            print("[ERROR] Ты сканер не выбрал долбаебина.")
            return None

        try:
            device_name = device.Properties("Name").Value
        except Exception:
            device_name = "Неизвестно"
        print(f"[INFO] Выбран сканер: {device_name}")

        item = device.Items[1]

        print("[INFO] Сетапим сканер на a4 и максимальное качество...")
        resolution = 600  # DPI – максимальное качество
        print(f"[INFO] Устанавливаем разрешение: {resolution} DPI (горизонтальное и вертикальное)")
        item.Properties["6147"].Value = resolution
        item.Properties["6148"].Value = resolution

        width_pixels = int(8.27 * resolution)
        height_pixels = int(11.69 * resolution)
        print(f"[INFO] Задаем область сканирования для A4: {width_pixels}x{height_pixels} пикселей")
        item.Properties["6149"].Value = 0  # X-начало
        item.Properties["6150"].Value = 0  # Y-начало
        item.Properties["6151"].Value = width_pixels
        item.Properties["6152"].Value = height_pixels

        try:
            print("[INFO] Принудительно устанавливаем цветной режим...")
            item.Properties["6146"].Value = 1  # 1 = цветной режим
        except Exception as e:
            print("[WARNING] НЕ СМОГ ПОСТАВИТЬ ЦВЕТНОЙ РЕЖИМ WTF??:", e)

        print("[INFO] Сканим суку.....")
        image = item.Transfer("{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}")

        timestamp = datetime.datetime.now().strftime("%d-%m-%y_%H.%M")
        filename = f"scan{timestamp}.jpg"
        image.SaveFile(filename)
        print(f"[SUCCESS] Скан готов. Файл сохранен: {filename}")
        return filename
    except Exception as e:
        print("[ERROR] НЕ СМОГ ОТСКАНИТЬ=(:", e)
        return None

def compress_image(input_file, output_file, quality=85):
    try:
        print(f"[INFO] Сжимаем пикчу чтоб не было 100мб.jpg {input_file} с качеством {quality}...")
        img = Image.open(input_file)
        img.save(output_file, "JPEG", quality=quality)
        print(f"[SUCCESS] Сжал эту суку как папа. Сжатый файл: {output_file}")
        return output_file
    except Exception as e:
        print("[ERROR] НЕ ОСИЛИЛ СЖАТЬ(:", e)
        return input_file

def raw_print(file_path, printer_ip, port=9100):
    try:
        print(f"[INFO] Отправляем эту суку: {file_path} на печать через raw printing на {printer_ip}:{port}...")
        with open(file_path, 'rb') as f:
            data = f.read()
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.connect((printer_ip, port))
        s.sendall(data)
        s.close()
        print("[SUCCESS] Эта сука ушла на печать через raw printing.")
    except Exception as e:
        print("[ERROR] НЕ СМОГ ОТПРАВИТЬ СУКУ НА ПЕЧАТЬ:", e)

if __name__ == "__main__":
    print_banner()
    scanned_file = scan_document()
    if scanned_file:
        compressed_file = scanned_file.replace(".jpg", "_compressed.jpg")
        compressed_file = compress_image(scanned_file, compressed_file, quality=85)

        # ВАЖНО ЧТОБ ЭТА ПОМОЙКА БЫЛА СО СТАТИКОЙ ИНАЧЕ КАКОЙ СМЫСЛ
        printer_ip = "192.168.1.199"  # указываем IP принтера
        print(f"[INFO] Используем мфу с IP: {printer_ip}")

        raw_print(compressed_file, printer_ip)


# ДЛЯ МОЕГО ЛЮБИМОГО СЫПАЮЩЕГОСЯ СМЕТНОГО ОТДЕЛА 90+ WOMENS HAHAHAHAHA
# By MKultra69 with hate