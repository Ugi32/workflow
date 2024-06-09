import os
import re
import datetime
from PyPDF2 import PdfReader, PdfMerger

def get_next_tuesday():
    today = datetime.date.today()
    days_until_tuesday = (1 - today.weekday() + 7) % 7
    days_until_tuesday = days_until_tuesday if days_until_tuesday != 0 else 7
    next_tuesday = today + datetime.timedelta(days=days_until_tuesday)
    return next_tuesday.strftime("%d.%m.%Y")

def get_top_number(filename):
    match = re.search(r"TOP(\d+)", filename)
    return int(match.group(1)) if match else None

def sort_key(filename):
    top_number = get_top_number(filename)
    top_number = top_number if top_number is not None else 0
    is_template = 'Vorlage' in filename

    if is_template:
        return (top_number, not is_template, 0)
    else:
        template_name = filename[:filename.rfind('Anlage')]
        attachment_number = int(filename.split()[-1][:-4]) if filename.endswith('Anlage') else 0
        return (top_number, not is_template, template_name, attachment_number)

def sorted_files(files):
    return sorted(files, key=sort_key)

def is_pdf_valid(filepath):
    try:
        with open(filepath, "rb") as file:
            reader = PdfReader(file)
            return True
    except Exception:  # Catching all exceptions that might occur during reading
        return False

def merge_pdfs(folder_path):
    merger = PdfMerger()
    all_pdfs = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
    valid_pdfs = [pdf for pdf in all_pdfs if is_pdf_valid(os.path.join(folder_path, pdf))]
    sorted_pdfs = sorted_files(valid_pdfs)

    next_tuesday = get_next_tuesday()
    output_filename = f"TO Unterlagen Vosi IBB {next_tuesday}.pdf"
    bookmarks = {}
    total_pages = 0

    # Hauptlesezeichen erstellen
    main_bookmark_title = output_filename[:-4]
    main_bookmark = merger.add_outline_item(title=main_bookmark_title, pagenum=0)

    for pdf in sorted_pdfs:
        pdf_path = os.path.join(folder_path, pdf)
        top_number = get_top_number(filename=pdf)
        is_template = 'Vorlage' in pdf

        with open(pdf_path, "rb") as file:
            reader = PdfReader(file)
            num_pages = len(reader.pages)

        merger.append(pdf_path)

        if is_template:
            bookmark = merger.add_outline_item(title=pdf[:-4], pagenum=total_pages, parent=main_bookmark)
            bookmarks[top_number] = bookmark
        else:
            if top_number in bookmarks:
                bookmark = bookmarks[top_number]
                merger.add_outline_item(title=pdf[:-4], pagenum=total_pages, parent=bookmark)
            else:
                print(f"Warning: No matching TOP number found for attachment: {pdf}")

        total_pages += num_pages

    output_path = os.path.join(folder_path, output_filename)
    merger.write(output_path)
    merger.close()
    print(f'Gesamtdokument gespeichert unter: {output_path}')

folder_path = 'C:/Users/ugur-/Desktop/PDF'
merge_pdfs(folder_path)
