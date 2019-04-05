# -*- coding: utf-8 -*-

import json
import xlrd
from xml.sax.saxutils import unescape

text_lst = ['']

header_lst = ['']

rb = xlrd.open_workbook('book.xlsx')
sheet = rb.sheet_by_index(0)
book_table = sheet.row_values


def parse_section(row, col):
    section_dict = dict()

    section_dict['name'] = book_table(row)[col]
    if book_table(row)[col + 1] and not book_table(row)[col + 2]:
        section_dict['textIndex'] = row
        text_lst.append(book_table(row)[col + 1])
        header_lst.append(book_table(row)[col])
    else:
        section_dict['textIndex'] = 0

    section_dict['sections'] = []

    if book_table(row)[col + 2]:
        section_dict['sections'].append(parse_section(row, col + 1))
        while(True):
            row += 1
            if row == sheet.nrows or col >= get_level(row):
                break

            if book_table(row)[col + 1]:
                section_dict['sections'].append(parse_section(row, col + 1))

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



def generate_text(lst):
    xml_begin = '<?xml version="1.0" encoding="utf-8"?>'
    root = etree.Element('resources')
    string_array = etree.Element('string-array', name='textArray', formatted='false')
    root.append(string_array)

    item = etree.Element('item')
    item.text = "<![CDATA[" + "empty element" + "]]>"
    string_array.append(item)

    for text in lst:
        item = etree.Element('item')
        item.text = "<![CDATA[" + text + "]]>"
        string_array.append(item)

    s = str(etree.tostring(root, pretty_print=True, encoding=str))
    open('text.xml', 'w', encoding='UTF-8').write(xml_begin + '\n' + unescape(s))


book_data = parse_section(1, 0)

print(len(header_lst))
print(len(text_lst))
with open('pro_ramadan_handbook_structure.dart', 'w', encoding='utf-8') as f:
    f.write(f"import '../util/section.dart'; Section handbookStructure = {get_structure(book_data)};")
with open('pro_ramadan_handbook_text.dart', 'w', encoding='utf-8') as f:
    f.write(get_text_headers())
    f.write(get_text())
