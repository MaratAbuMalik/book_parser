# -*- coding: utf-8 -*-

import json
import xlrd
from xml.sax.saxutils import unescape

text_lst = ['']

header_lst = ['']

bookmarks_header_lst = ['']

rb = xlrd.open_workbook('book.xlsx')
sheet = rb.sheet_by_index(0)
book_table = sheet.row_values


def parse_section(row, col, header):
    section_dict = dict()

    section_dict['name'] = book_table(row)[col]
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
            parse_section(row, col + 1, header + '.' if 'ProRamadan' not in header else ''))
        while(True):
            row += 1
            if row == sheet.nrows or col >= get_level(row):
                break

            if book_table(row)[col + 1]:
                section_dict['sections'].append(
                    parse_section(row, col + 1, header + '.' if 'ProRamadan' not in header else ''))

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


book_data = parse_section(1, 0, '')

print(len(header_lst))
print(len(text_lst))
with open('pro_ramadan_handbook_structure.dart', 'w', encoding='utf-8') as f:
    f.write(f"import '../util/section.dart'; Section handbookStructure = {get_structure(book_data)};")
with open('pro_ramadan_handbook_text.dart', 'w', encoding='utf-8') as f:
    f.write(get_text_headers())
    f.write(get_bookmarks_headers())
    f.write(get_text())
