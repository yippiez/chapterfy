
import PySimpleGUI as sg
from pypdf import PdfReader, PdfWriter

from dataclasses import dataclass
from typing import Tuple
import re
import sys
import os


################ CONSTANTS
TARGET_PDF_PATH = ''
CHAPTERS = []
CHAPTER_TABLE_VALUE = None
OUTPUT_FOLDER = ''


################ DATA CLASSES

@dataclass
class Chapter:
    chapter_name: str
    begining_page: int
    end_page: int

def add_chapter(chapter_name: str, begining_page: int, end_page: int):
    to_add = Chapter(chapter_name, begining_page, end_page)

    if to_add not in CHAPTERS:
        CHAPTERS.append(to_add)

def remove_chapter(chapter_name: str):
    for index, chapter in enumerate(CHAPTERS):
        if chapter.chapter_name == chapter_name:
            CHAPTERS.pop(index)
            return

################ FUNCTIONS

def update_table_info() -> None:
    chapter_count = len(CHAPTER_TABLE_VALUE)
    selected_chapter_count = len([o for o in CHAPTER_TABLE_VALUE if o[3] == True])
    window['-CHAPTERS_INFO-'].update(f"{selected_chapter_count}/{chapter_count} chapters selected")

def read_nested_bookmarsk(reader: PdfReader, bookmarks: list, o) -> list:
    if isinstance(o, list):
        # for o2 in o:
            # read_nested_bookmarsk(reader, bookmarks, o2)
        pass
    else:
        bookmarks.append((o.title.strip().rstrip(), reader.get_destination_page_number(o) + 1))

def update_chapters_data(reader: PdfReader, keyword: str = 'chapter') -> None:
    global CHAPTER_TABLE_VALUE

    # bookmarks = [(o.title.strip().rstrip(), reader.get_destination_page_number(o) + 1) for o in reader.outline if not isinstance(o, list)]
    bookmarks = []
    for i, o in enumerate(reader.outline):
        read_nested_bookmarsk(reader, bookmarks, o)

    result = []

    for i in range(len(bookmarks) - 1):
        chapter_name, begining_page = bookmarks[i]
        _, end_page = bookmarks[i + 1]
        result.append([chapter_name, begining_page, end_page - 1, False])

    # filter result where o.title.lower() contains keyword
    result = list(filter(lambda o: keyword in o[0].lower(), result))

    # update table
    window['-CHAPTERS-'].update(values=result)
    CHAPTER_TABLE_VALUE = result

    # update table info
    update_table_info()


def compress_pdf(writer: PdfWriter):
    for page in writer.pages:
        page.compress_content_streams()

def chapterfy(pdf_path: str, chapters: list, output_folder: str):
    print('Chapterfying...')
    print(f"PDF path: {pdf_path}")
    print(f"Chapters: {chapters}")
    print(f"Output folder: {output_folder}")
    
    # Check if output folder exists
    if not os.path.exists(output_folder):
        sg.popup(f'Output folder does not exist: {output_folder}')
        return

    # Check if pdf exists
    if not os.path.exists(pdf_path):
        sg.popup(f'PDF does not exist: {pdf_path}')
        return

    reader = PdfReader(pdf_path)
    number_of_pages = len(reader.pages)    

    for index, chapter in enumerate(chapters):
        print(f'\nDoing chapter {index + 1} of {len(chapters)}')

        begining_page, end_page = chapter.begining_page, chapter.end_page

        match (begining_page > number_of_pages, end_page > number_of_pages):
            case (True, True):
                sg.popup(f'Both page ranges is bigger than number of pages: {chapter} skiping...')
                continue
            case (True, False):
                sg.popup(f'Begining page is bigger than number of pages: {begining_page} skiping...')
                continue
            case (False, True):
                sg.popup(f'End page is bigger than number of pages: {end_page} skiping...')
                continue
            case _:
                pass
        
        # make name safe for windows
        output_file_name = chapter.chapter_name
        output_file_name = output_file_name.replace('/', '_')
        output_file_name = re.sub(r'[^\w]', ' ', output_file_name)
        output_file_name = f'{output_file_name.strip()}.pdf'

        # make output file path
        output_file_path = os.path.join(output_folder, output_file_name)
        
        print(f"Chapter {index + 1} of {len(chapters)}")
        print(f"Begining page: {begining_page}")
        print(f"End page: {end_page}")

        writer = PdfWriter()
        
        pages = reader.pages[begining_page - 1:end_page]
        
        for page in pages:
            writer.add_page(page)

        compress_pdf(writer)

        writer.write(output_file_path)

    print('Done!')


def check_inputs_present() -> Tuple[bool, bool, bool]:
    is_pdf_path_present = TARGET_PDF_PATH != ''
    is_chapters_present = len(CHAPTERS) > 0
    is_output_folder_present = OUTPUT_FOLDER != ''

    return (is_pdf_path_present, is_chapters_present, is_output_folder_present)


def check_valid_range(_range: str) -> Tuple[bool, bool]:
    match = re.match(r'\d+-\d+', _range)
    
    is_match_present = match is not None
    is_range_valid = False

    if is_match_present:
        range_start, range_end = _range.split('-')
        is_range_valid = int(range_start) <= int(range_end)

    return (is_match_present, is_range_valid)


################ GUI
layout = [
    [sg.Column([
        [sg.Text('Chapterfy', size=(40, 1), font=("Helvetica", 25), justification='center')]
        ], element_justification='center')],

    [sg.Text('Select PDF file', size=(15, 1), auto_size_text=False, justification='right'),
        sg.InputText(key='-SELECTED_INPUT_PATH-', readonly=True, change_submits=True), sg.FileBrowse()],

    [sg.Text("Known chapters", size=(15, 1), auto_size_text=False, justification='right'),
        sg.Table(values=CHAPTERS, headings=['Chapter name', 'Begining page', 'End page', 'Included'], 
        key='-CHAPTERS-', enable_events=True, justification='center', num_rows=5, col_widths=[10, 10, 10, 5]),],

    [sg.Text("0/0 chapters selected", size=(30, 1), auto_size_text=False, justification='right', key='-CHAPTERS_INFO-')],

    [sg.Text('Denote page ranges', size=(15, 1), auto_size_text=False, justification='right'),
        sg.InputText(key='-PAGE_RANGE-'), sg.Button('Apply', key='-PAGE_RANGE_APPLY-'), sg.Text('e.g. 1-5, 7, 9-11')],

    [sg.Text('Select output folder', size=(15, 1), auto_size_text=False, justification='right'),
        sg.InputText(key='-SELECTED_OUTPUT_PATH-', readonly=True, change_submits=True), sg.FolderBrowse()],

    [sg.Column([
        [sg.Submit(), sg.Cancel()]
    ], element_justification='right')],
]

window = sg.Window('Chapterfy', layout, default_element_size=(70, 1), grab_anywhere=False)

while True:
    event, values = window.read()
    
    if event == sg.WIN_CLOSED or event == 'Cancel':
        break

    if event == '-SELECTED_INPUT_PATH-':
        TARGET_PDF_PATH = values['-SELECTED_INPUT_PATH-']
        update_chapters_data(PdfReader(TARGET_PDF_PATH), '')
        CHAPTERS = []
        print(f"PDF path: {TARGET_PDF_PATH}")      

    if event == '-PAGE_RANGE_APPLY-':
        _chapters = [s.strip() for s in values['-PAGE_RANGE-'].split(',')]

        for chapter in _chapters:
            match check_valid_range(chapter):
                case (True, True):
                    begining_page, end_page = [int(r) for r in chapter.split('-')]
                    chapter_name = f'{begining_page}_{end_page}'
                    add_chapter(chapter_name, begining_page, end_page)
                case (True, False):
                    sg.popup(f'Invalid range: {chapter}')
                case (False, False):
                    sg.popup(f'Invalid chapter syntax: {chapter}')
                case _:
                    raise Exception('Invalid match')

        print(f"Chapters: {CHAPTERS}")

    if event == '-SELECTED_OUTPUT_PATH-':
        OUTPUT_FOLDER = values['-SELECTED_OUTPUT_PATH-']
        print(f"Output folder: {OUTPUT_FOLDER}")

    if event == 'Submit':
        match check_inputs_present():
            case (True, True, True):
                print('All inputs present')
                chapterfy(TARGET_PDF_PATH, CHAPTERS, OUTPUT_FOLDER)
            case (False, True, True):
                sg.popup('PDF path not present')
            case (True, False, True):
                sg.popup('Chapters not present')
            case (True, True, False):
                sg.popup('Output folder not present')
            case (False, False, True):
                sg.popup('PDF path and chapters not present')
            case (False, True, False):
                sg.popup('PDF path and output folder not present')
            case (True, False, False):
                sg.popup('Chapters and output folder not present')
            case (False, False, False):
                sg.popup('No inputs present')

    if event == '-CHAPTERS-':
        if values['-CHAPTERS-'] == []:
            continue

        print(f"Selected row: {values['-CHAPTERS-']}")

        selected_row = values['-CHAPTERS-'][0]

        selected_chapter = CHAPTER_TABLE_VALUE[selected_row][0]
        begining_page = CHAPTER_TABLE_VALUE[selected_row][1]
        end_page = CHAPTER_TABLE_VALUE[selected_row][2]

        # Check if chapter is already in CHAPTERS
        should_add_chapter = not (Chapter(selected_chapter, begining_page, end_page) in CHAPTERS) 
        if should_add_chapter:
            print(f"Adding chapter: {selected_chapter}")
            add_chapter(selected_chapter, begining_page, end_page)
        else:
            print(f"Removing chapter: {selected_chapter}")
            remove_chapter(selected_chapter)

        # Update table
        CHAPTER_TABLE_VALUE[selected_row][3] = should_add_chapter
        window['-CHAPTERS-'].update(values=CHAPTER_TABLE_VALUE)

        # Update table info
        update_table_info() 

        print(f"Chapters: {CHAPTERS}")

window.close()