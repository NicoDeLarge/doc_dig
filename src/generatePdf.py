from PyPDF2 import PdfMerger, PdfReader
from pathlib import Path
import datetime
try:
    from PIL import Image
except ImportError:
    import Image
import pytesseract
import re
import os
import subprocess

from prompt_toolkit import prompt
from prompt_toolkit.completion import WordCompleter, FuzzyWordCompleter
from prompt_toolkit.shortcuts import ProgressBar


inbox_dir = Path('F:\\Scans\\inbox')
processed_dir = Path('F:\\Scans')
archive_dir = Path('F:\\Scans\\archive_raw')


document_split_time = 2 # in seconds


class File:
    def __set__(self, file_path, parent_file = None):
        self.file_path = file_path
        self.parent_file = parent_file

    def get_file_type(self):
        self.file_path.exists()

    def exists(self):
        return self.file_path.exists()


class Doc:

    def __init__(self, pages):
        self.creator = ''
        self.recipient = ''
        self.date_creation = ''
        self.topic = ''
        self.pages = pages

    def add_page(self, page):
        self.pages.append(page)

    def delete_all_pages(self):
        for idx in range(len(self.pages)-1, -1, -1):
            # archive images
            source = Path(self.pages[idx]).parent / (Path(self.pages[idx]).stem + ".jpg")
            dest = Path(archive_dir) / (Path(self.pages[idx]).stem + ".jpg")

            if not source.exists():
                print(source, " not found. Skip moving to archive.")
            else:
                if not dest.exists():
                    source.replace(dest)
                else:
                    print(dest, " already exists. Skip moving to archive.")

            # delete pdf pages
            Path(self.pages[idx]).unlink()

            if Path(self.pages[idx]).exists():
                raise FileExistsError
            else:
                self.pages.remove(self.pages[idx])

    def get_file_name(self):
        date = datetime.datetime.strftime(self.date_creation, '%Y-%m-%d')

        return date + '_-_' + self.creator.replace(' ', '_') + '_-_' + self.recipient.replace(' ', '_') + '_-_' + self.topic.replace(' ', '_')

    def save_as_pdf(self, path):
        merger = PdfMerger()
        input_page_sizes = []

        for page in self.pages:
            if page.suffix == '.pdf':
                merger.append(PdfReader(str(page), strict=False))
                input_page_sizes.append(os.stat(page).st_size)

        path_merged_file = Path(path) / Path(self.get_file_name() + '.pdf')

        merger.write(str(path_merged_file))
        merger.close()

        if not path_merged_file.exists():
            return None
        else:
            if os.stat(path_merged_file).st_size < sum(input_page_sizes) - min(input_page_sizes) :
                return None
            else:
                return path_merged_file
            
def extract_attributes_from_file_name(file_name):
    parts = file_name.split('_-_')
    
    date = datetime.datetime.now()
    creator = "James Bond"
    receiver = "Dr. No"
    topic = "Toppling"
    
    if len(parts) > 0:
        date = datetime.datetime.strptime(parts[0], '%Y-%m-%d')
     
    if len(parts) > 1:
        creator = parts[1].replace("_", " ").strip()
        
    if len(parts) > 2:
        receiver = parts[2].replace("_", " ").strip()
        
    if len(parts) > 3:
        topic = parts[3].replace("_", " ").strip()
    
    return date, creator, receiver, topic

def build_completers(processed_dir):
    creator_receiver_list = []
    topic_list = []
    
    if processed_dir.exists():
        file_list = sorted(processed_dir.glob('./*.pdf'))
        
        for f in file_list:
            date, creator, receiver, topic = extract_attributes_from_file_name(f.stem)
            
            creator_receiver_list.append(creator)
            creator_receiver_list.append(receiver)
            topic_list.append(topic)
    
    creator_receiver_completer = FuzzyWordCompleter(list(set(creator_receiver_list)))
    topic_completer = FuzzyWordCompleter(list(set(topic_list)))
    
    return creator_receiver_completer, topic_completer

def get_scan_date_time_from_file_name(file):
    date = datetime.datetime.now()
     
    try: 
        match1 = re.search('\d{8}_\d{6}', file.name)
        date = datetime.datetime.strptime(match1.group(0), '%Y%m%d_%H%M%S')
    except IndexError:
        None
    
    try:
        match2 = re.search('_\d{8}_\d{6}_', file.name)
        datetime.datetime.strptime(match2.group(0), '_%Y%m%d_%H%M%S_')
    except:
        None 
    
    return date

if inbox_dir.exists():
    creator_receiver_completer, topic_completer = build_completers(processed_dir)
    
    print("Step 1: Read new images")
    img_list = sorted(inbox_dir.glob('*.jpg'))

    print("Step 2: Generate OCRed PDFs for " + str(len(img_list)) + " images")

    do_ocr = input("Do OCR [Y/n]?")

    if do_ocr == "Y" or do_ocr == "":
        # convert image to ocr'ed pdf
        
        with ProgressBar() as pb:
            for img in pb(img_list):
                pdf = pytesseract.pytesseract.image_to_pdf_or_hocr(str(img), lang='deu', extension='pdf', timeout=30)

                with open(img.with_suffix('.pdf'), 'w+b') as f:
                    f.write(pdf)

    # merge single page pdfs to one multi-page pdf
    print("Step 3: Merge single page PDFs")
    pdf_list = sorted(inbox_dir.glob('*.pdf'))

    doc_list = []

    if pdf_list:
        doc_list = [Doc([pdf_list[0]])]

        for i in range(1, len(img_list)):
            delta_time = get_scan_date_time_from_file_name(pdf_list[i]) - get_scan_date_time_from_file_name(pdf_list[i-1])

            if delta_time.total_seconds() < document_split_time:
                doc_list[-1].add_page(pdf_list[i])
            else:
                doc_list.append(Doc([pdf_list[i]]))

    # assign tags
    print("Step 4: Assign document Tags")

    for doc in doc_list:
        print("\n#################\nStep 1: Merge single page PDFs to document based on the following PDFs")

        for page in doc.pages:
            print("+ " + str(page))

        doc.topic = str(doc.pages[0].stem)
        doc.date_creation = get_scan_date_time_from_file_name(doc.pages[0])
        doc.creator = ""
        doc.recipient = ""

        saved_file_dir = doc.save_as_pdf(processed_dir)

        if saved_file_dir is not None:
            print("Created merged PDF ", saved_file_dir)

            subprocess.run(["start", str(saved_file_dir)], check=True, shell=True)

            delete_requested = input("Delete single page PDFs [Y/n]?")

            if delete_requested == "Y" or delete_requested == "":
                doc.delete_all_pages()
                doc.add_page(saved_file_dir)

        # assign tags
        default_date = doc.date_creation.strftime("%Y-%m-%d")
        requested_date = prompt("Date: ", default=default_date)

        if requested_date != "":
            doc.date_creation = datetime.datetime.strptime(requested_date, "%Y-%m-%d")

        doc.creator = prompt("Creator: ", completer=creator_receiver_completer)
        doc.recipient = prompt("Recipient: ", completer=creator_receiver_completer)
        doc.topic = prompt("Topic: ", completer=topic_completer)

        new_file_name = Path((doc.get_file_name() + ".pdf"))

        request_rename = input("Rename from '" + str(Path(saved_file_dir).name) + "'\nto '" + str(new_file_name) + "'? [Y/n]")

        if request_rename == "" or request_rename == "Y":
            new_file_path = Path(saved_file_dir)
            Path(saved_file_dir).rename(Path(saved_file_dir).parent / new_file_name)
