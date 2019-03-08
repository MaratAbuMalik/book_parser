import xlrd

russian_titles = []
arabic_titles = []
russian_matns = []
arabic_matns = []
sharkhs = []
questions = []
lecturers = []


def parse():
    rb = xlrd.open_workbook('structure.xlsx')
    sheet = rb.sheet_by_index(0)
    for rownum in range(1, sheet.nrows):
        row = sheet.row_values(rownum)
        russian_titles.append(row[0])
        arabic_titles.append(row[1])
        russian_matns.append(row[2])
        arabic_matns.append(row[3])
        sharkhs.append(row[4])
        questions.append(row[5])
        lecturer = []
        for lecturers_rownum in range(6, 10):
            lecturer.append(str(row[lecturers_rownum]).split())
        lecturers.append(lecturer)


parse()

with open('chapters.dart', 'w', encoding='utf-8') as f:
    f.write("import '../util/chapter_class.dart';\n\n")
    f.write("List<Chapter> chapters = [\n")

    for i in range(len(russian_titles)):
        f.write(f'Chapter(\n')
        f.write(f'russianHeader: """{russian_titles[i]}""",\n')
        f.write(f'arabicHeader: """{arabic_titles[i]}""",\n')
        f.write(f'tabList: [\n')
        f.write(f'TabDescription(text: """{russian_matns[i]}"""),\n')
        f.write(f'TabDescription(text: """{arabic_matns[i]}""", isArabic: true),\n')
        f.write(f'TabDescription(text: """{sharkhs[i]}"""),\n')
        f.write(f'TabDescription(text: """{questions[i]}"""),\n')
        f.write(f'TabDescription(\n')
        f.write(f'lecturerList: [\n')
        for j in lecturers[i]:
            f.write(f'LecturerDescription(\n')
            f.write(f'audioList: [\n')
            for k in j:
                f.write(f'AudioDescription(\n')
                f.write(f'address: "{k}",\n')
                f.write('),\n')
            f.write('],\n')
            f.write('),\n')
        f.write('],\n')
        f.write('),\n')
        f.write('],\n')
        f.write('),\n')

    f.write("];")

if __name__ == '__main__':
    for i in range(len(russian_titles)):
        print(russian_titles[i])
        print(arabic_titles[i])
        print(russian_matns[i])
        print(arabic_matns[i])
        print(sharkhs[i])
        print(questions[i])
        for j in lecturers[i]:
            print(j)
        print()
