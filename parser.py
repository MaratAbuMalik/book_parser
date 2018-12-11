import xlrd

russian_titles = []
arabic_titles = []
russian_matns = []
arabic_matns = []


def parse():
    rb = xlrd.open_workbook('structure.xlsx')
    sheet = rb.sheet_by_index(0)
    for rownum in range(sheet.nrows):
        row = sheet.row_values(rownum)
        russian_titles.append(row[0])
        arabic_titles.append(row[1])
        russian_matns.append(row[2])
        arabic_matns.append(row[3])


parse()
with open('chapters.dart', 'w', encoding='utf-8') as f:
    f.write("import 'chapter_class.dart';\n")
    f.write("List<Chapter> chapters = [\n")

    for i in range(len(russian_titles)):
        f.write(f'Chapter(\n')
        f.write(f'"""{russian_titles[i]}""",\n')
        f.write(f'"""{arabic_titles[i]}""",\n')
        f.write(f'"""{russian_matns[i]}""",\n')
        f.write(f'"""{arabic_matns[i]}"""\n')
        f.write('),\n')

    f.write("];")

if __name__ == '__main__':
    for i in range(len(russian_titles)):
        print(russian_titles[i])
        print(arabic_titles[i])
        print(russian_matns[i])
        print(arabic_matns[i])
        print()
