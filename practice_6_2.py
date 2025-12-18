import os
import sys
import glob
import argparse
from pathlib import Path
from pdf2docx import Converter
from docx2pdf import convert
from PIL import Image

def pdf2docx_single(pdf_path):
    docx_path = str(Path(pdf_path).with_suffix('.docx'))
    cv = Converter(pdf_path)
    cv.convert(docx_path)
    cv.close()

def pdf2docx_all(workdir):
    workdir = Path(workdir)
    for pdf in workdir.glob("*.pdf"):
        pdf2docx_single(str(pdf))

def docx2pdf_single(docx_path):
    pdf_path = str(Path(docx_path).with_suffix('.pdf'))
    convert(docx_path, pdf_path)

def docx2pdf_all(workdir):
    workdir = Path(workdir)
    for docx in workdir.glob("*.docx"):
        docx2pdf_single(str(docx))

def compress_image(path, quality):
    img = Image.open(path)
    out_path = f"compressed_{Path(path).name}"
    img.save(out_path, optimize=True, quality=quality)

def compress_images_single(path, quality):
    compress_image(path, quality)

def compress_images_all(workdir, quality):
    for ext in ("*.jpg", "*.jpeg", "*.png"):
        for f in glob.glob(ext, root_dir=workdir):
            full_path = os.path.join(workdir, f)
            compress_image(full_path, quality)

def delete_files(delete_dir, mode, pattern):
    files = []
    if mode == "extension":
        files = glob.glob(f"*.{pattern}", root_dir=delete_dir)
    elif mode == "startswith":
        files = [f for f in os.listdir(delete_dir) if f.startswith(pattern)]
    elif mode == "endswith":
        files = [f for f in os.listdir(delete_dir) if f.endswith(pattern)]
    elif mode == "contains":
        files = [f for f in os.listdir(delete_dir) if pattern in f]

    for f in files:
        os.remove(os.path.join(delete_dir, f))

# --- Интерактивный режим (без argparse) ---
def change_working_directory():
    path = input("Введите путь к новому рабочему каталогу: ")
    os.chdir(path)
    print(f"Текущий каталог: {os.getcwd()}")

def interactive_pdf2docx():
    pdf = input("Путь к PDF: ")
    pdf2docx_single(pdf)

def interactive_docx2pdf():
    docx = input("Путь к DOCX: ")
    docx2pdf_single(docx)

def interactive_compress():
    q = int(input("Качество (1-100): ") or 75)
    mode = input("Один файл (1) или все (2)? ")
    if mode == "1":
        f = input("Путь к изображению: ")
        compress_images_single(f, q)
    else:
        d = input("Папка: ")
        compress_images_all(d, q)

def interactive_delete():
    d = input("Папка для удаления: ")
    m = input("Режим (extension/startswith/endswith/contains): ")
    p = input("Паттерн: ")
    delete_files(d, m, p)

def exit_program():
    print("До свидания!")
    sys.exit(0)

def run_interactive():
    actions = {
        0: change_working_directory,
        1: interactive_pdf2docx,
        2: interactive_docx2pdf,
        3: interactive_compress,
        4: interactive_delete,
        5: exit_program,
    }
    menu = [
        "0. Сменить рабочий каталог",
        "1. Преобразовать PDF в Docx",
        "2. Преобразовать Docx в PDF",
        "3. Произвести сжатие изображений",
        "4. Удалить группу файлов",
        "5. Выход"
    ]
    while True:
        print("\nВыберите действие:")
        for line in menu:
            print(f"  {line}")
        try:
            choice = int(input("Введите номер действия: "))
            if choice in actions:
                actions[choice]()
            else:
                print("Неверный выбор.")
        except ValueError:
            print("Введите число.")

# --- Основной запуск с argparse ---
def main():
    parser = argparse.ArgumentParser(description="Утилита для работы с офисными файлами и изображениями.")
    parser.add_argument("--pdf2docx", metavar="PATH", help="Путь к PDF-файлу или 'all'")
    parser.add_argument("--docx2pdf", metavar="PATH", help="Путь к DOCX-файлу или 'all'")
    parser.add_argument("--compress-images", metavar="PATH", help="Путь к изображению или 'all'")
    parser.add_argument("--quality", type=int, default=75, help="Качество сжатия изображений (1-100)")
    parser.add_argument("--workdir", help="Рабочая директория (для обработки всех файлов)")
    parser.add_argument("--delete", action="store_true", help="Включить режим удаления")
    parser.add_argument("--delete-mode", choices=["startswith", "endswith", "contains", "extension"], help="Режим удаления")
    parser.add_argument("--delete-pattern", help="Паттерн для удаления")
    parser.add_argument("--delete-dir", help="Директория для удаления")
    parser.add_argument("-i", "--interactive", action="store_true", help="Запустить интерактивный режим")

    args = parser.parse_args()

    # Определяем, был ли задан хоть один рабочий аргумент
    has_action = any([
        args.pdf2docx,
        args.docx2pdf,
        args.compress_images,
        args.delete,
        args.interactive
    ])

    if not has_action:
        args.interactive = True

    if args.interactive:
        run_interactive()
        return

    if args.pdf2docx:
        if args.pdf2docx == "all":
            if not args.workdir:
                parser.error("--workdir обязателен при --pdf2docx all")
            pdf2docx_all(args.workdir)
        else:
            pdf2docx_single(args.pdf2docx)

    if args.docx2pdf:
        if args.docx2pdf == "all":
            if not args.workdir:
                parser.error("--workdir обязателен при --docx2pdf all")
            docx2pdf_all(args.workdir)
        else:
            docx2pdf_single(args.docx2pdf)

    if args.compress_images:
        if args.compress_images == "all":
            if not args.workdir:
                parser.error("--workdir обязателен при --compress-images all")
            compress_images_all(args.workdir, args.quality)
        else:
            compress_images_single(args.compress_images, args.quality)

    if args.delete:
        if not (args.delete_mode and args.delete_pattern and args.delete_dir):
            parser.error("Для --delete требуются --delete-mode, --delete-pattern и --delete-dir")
        delete_files(args.delete_dir, args.delete_mode, args.delete_pattern)

if __name__ == "__main__":
    main()