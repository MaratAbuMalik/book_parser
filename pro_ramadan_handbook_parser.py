# -*- coding: utf-8 -*-

import json
import xlrd
from xml.sax.saxutils import unescape

text_lst = ['']

header_lst = ['']

bookmarks_header_lst = ['']

rb = xlrd.open_workbook('text.xlsx')
sheet = rb.sheet_by_index(0)
book_table = sheet.row_values


def parce_text(file):
    global rb
    global sheet
    global book_table
    global text_lst
    global header_lst
    global bookmarks_header_lst

    rb = xlrd.open_workbook(file)
    sheet = rb.sheet_by_index(0)
    book_table = sheet.row_values
    text_lst = ['']
    header_lst = ['']
    bookmarks_header_lst = ['']

    return parse_section(1, 0, '')


def parse_section(row, col, header, prefix=''):
    section_dict = dict()

    section_dict['name'] = prefix + book_table(row)[col]
    if book_table(row)[col + 1] and not book_table(row)[col + 2]:
        section_dict['textIndex'] = row
        text_lst.append(book_table(row)[col + 1])

        bookmarks_header_lst.append(book_table(row)[col])
        header_lst.append(header + '\n' + book_table(row)[col])
    else:
        section_dict['textIndex'] = 0

    section_dict['sections'] = []

    if book_table(row)[col + 2]:
        header = header + '\n' + book_table(row)[col] if header != '' else book_table(row)[col]
        section_dict['sections'].append(
            parse_section(row, col + 1, header + '\n' if 'ProRamadan' not in header else '',
                          prefix=(prefix.strip() + str(len(section_dict['sections']) + 1) + '. ')))
        while(True):
            row += 1
            if row == sheet.nrows or col >= get_level(row):
                break

            if book_table(row)[col + 1]:
                section_dict['sections'].append(
                    parse_section(row, col + 1, header + '\n' if 'ProRamadan' not in header else '',
                                  prefix=(prefix.strip() + str(len(section_dict['sections']) + 1)
                                          + '. ')))

    return section_dict


def get_level(row):
    level = 0
    while level < len(book_table(row)) and not book_table(row)[level]:
        level += 1
    return level


def get_structure(data):
    sections = [get_structure(i) + ", " for i in data["sections"]]
    section_string = f"""Section(name: '''{data["name"]}''', textIndex: {data["textIndex"]}, sections: """

    section_string += '['

    for i in sections:
        section_string += i

    section_string += ']'

    section_string += ')'

    return section_string


def get_text():
    text = 'List<String> handbookText = ['

    for i in text_lst:
        text += f"'''{i}''',\n"

    text += '];'

    return text


def get_text_headers():
    text = 'List<String> handbookTextHeaders = ['

    for i in header_lst:
        text += f"'''{i}''',\n"

    text += '];'

    return text


def get_bookmarks_headers():
    text = 'List<String> bookmarksTextHeaders = ['

    for i in bookmarks_header_lst:
        text += f"'''{i}''',\n"

    text += '];'

    return text


def get_share_text():
    text = 'List<String> handbookShareText = ['

    for i in range(len(header_lst)):
        text += f"'''{header_lst[i]}\n\n{text_lst[i]}''',\n"

    text += '];'

    return text


book_data = parce_text(file='text.xlsx')

print(len(header_lst))
print(len(text_lst))
with open('structure.dart', 'w', encoding='utf-8') as f:
    f.write(f"import '../util/util_classes.dart'; Section handbookStructure = {get_structure(book_data)};")
with open('text.dart', 'w', encoding='utf-8') as f:
    f.write(get_text_headers())
    f.write(get_bookmarks_headers())
    f.write(get_text())


share_book_data = parce_text(file='share_text.xlsx')
with open('share_text.dart', 'w', encoding='utf-8') as f:
    f.write(get_share_text())
