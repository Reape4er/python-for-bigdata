import os
import sys
import glob
from pathlib import Path
from pdf2docx import Converter
from docx2pdf import convert
from PIL import Image

MENU_OPTIONS = [
    "0. Сменить рабочий каталог",
    "1. Преобразовать PDF в Docx",
    "2. Преобразовать Docx в PDF",
    "3. Произвести сжатие изображений",
    "4. Удалить группу файлов",
    "5. Выход"
]

def display_menu():
    print("\nВыберите действие:")
    for option in MENU_OPTIONS:
        print(f"  {option}")

def get_user_choice():
    try:
        choice = int(input("Введите номер действия: "))
        return choice if 0 <= choice <= 5 else None
    except ValueError:
        return None

def change_working_directory():
    path = input("Введите путь к новому рабочему каталогу: ")
    os.chdir(path)
    print(f"Текущий каталог: {os.getcwd()}")

def convert_pdf_to_docx():
    pdf = input("Путь к PDF: ")
    docx = str(Path(pdf).with_suffix('.docx'))
    cv = Converter(pdf)
    cv.convert(docx)
    cv.close()
    print(f"Создан файл: {docx}")

def convert_docx_to_pdf():
    docx = input("Путь к DOCX: ")
    pdf = str(Path(docx).with_suffix('.pdf'))
    convert(docx, pdf)
    print(f"Создан файл: {pdf}")

def compress_images():
    quality = int(input("Качество (1-95, по умолчанию 85): ") or 85)
    for ext in ("*.jpg", "*.jpeg", "*.png"):
        for f in glob.glob(ext):
            img = Image.open(f)
            img.save(f"compressed_{f}", optimize=True, quality=quality)

def delete_files_group():
    pattern = input("Маска файлов для удаления (например, *.tmp): ")
    files = glob.glob(pattern)
    for f in files:
        os.remove(f)

def exit_program():
    sys.exit(0)

ACTION_MAP = {
    0: change_working_directory,
    1: convert_pdf_to_docx,
    2: convert_docx_to_pdf,
    3: compress_images,
    4: delete_files_group,
    5: exit_program
}

def run_menu_loop():
    while True:
        display_menu()
        choice = get_user_choice()
        if choice is not None and choice in ACTION_MAP:
            ACTION_MAP[choice]()

if __name__ == "__main__":
    run_menu_loop()
    