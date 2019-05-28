import xlrd
from requests import head

russian_titles = []
arabic_titles = []
russian_matns = []
arabic_matns = []
sharkhs = []
questions = []
audios = []
lecturers = []
chapters = []


def parse_book():
    rb = xlrd.open_workbook('book.xlsx')
    sheet = rb.sheet_by_index(0)
    for rownum in range(1, sheet.nrows):
        row = sheet.row_values(rownum)
        chapters.append(row[0])
    for rownum in range(1, sheet.nrows):
        row = sheet.row_values(rownum)
        russian_titles.append(row[0])
        arabic_titles.append(row[1])
        russian_matns.append(row[2])
        arabic_matns.append('<bdo dir="rtl">' + row[3] + '</bdo>')
        sharkhs.append(row[4])
        questions.append(row[5])


def parse_audio():
    rb = xlrd.open_workbook('audio.xlsx')
    sheet = rb.sheet_by_index(0)
    for rownum in range(0, 1):
        row = sheet.row_values(rownum)
        for lecturers_rownum in range(1, 5):
            lecturers.append(row[lecturers_rownum])
    for rownum in range(1, sheet.nrows):
        row = sheet.row_values(rownum)
        lecturer_audios = []
        for lecturers_rownum in range(1, 5):
            lecturer_audios.append(str(row[lecturers_rownum]).split())
        audios.append(lecturer_audios)


def write_book():
    with open('book.dart', 'w', encoding='utf-8') as book:
        book.write("import '../util/chapter_class.dart';\n\n")
        book.write("List<Chapter> chapters = [\n")

        for i in range(len(russian_titles)):
            book.write(f'Chapter(\n')
            book.write(f'russianHeader: """{russian_titles[i]}""",\n')
            book.write(f'arabicHeader: """{arabic_titles[i]}""",\n')
            book.write(f'tabList: [\n')
            book.write(f'TabDescription(text: """{russian_matns[i]}"""),\n')
            book.write(f'TabDescription(text: """{arabic_matns[i]}""", isArabic: true),\n')
            book.write(f'TabDescription(text: """{sharkhs[i]}"""),\n')
            book.write(f'TabDescription(text: """{questions[i]}"""),\n')
            book.write(f'TabDescription(\n')
            book.write(f'isAudio: true\n')
            book.write('),\n')
            book.write('],\n')
            book.write('),\n')
        book.write("];")


def write_audio():
    with open('audio.dart', 'w', encoding='utf-8') as lectures:
        lectures.write("import 'package:educational_audioplayer/player.dart';\n\n")

        lectures.write('List<String> authorNames = [')
        for i in lecturers:
            lectures.write(f'"{i}",')
        lectures.write("];\n\n")

        lectures.write('List<String> chapterNames = [')
        for i in chapters:
            lectures.write(f'"{i}",')
        lectures.write("];\n\n")

        lectures.write("List audios = [\n")
        for i in range(len(chapters)):
            lectures.write(f'[\n')
            for j in range(len(audios[i])):
                lectures.write(f'[\n')
                for k in range(len(audios[i][j])):
                    lectures.write('Audio(')
                    lectures.write(f'url: "{audios[i][j][k]}", ')
                    lectures.write(f'audioName: "Лекция № {k + 1}", ')
                    lectures.write(f'audioSize: {round(int(head(audios[i][j][k], allow_redirects=True).headers["Content-Length"]) / 1024 / 1024)}, ')
                    lectures.write(f'audioDescription: "", ')
                    lectures.write(f'chapterName: chapterNames[{i}], ')
                    lectures.write(f'authorName: authorNames[{j}], ')
                    lectures.write('), \n')
                lectures.write('],\n')
            lectures.write('],\n')
        lectures.write("];")


if __name__ == '__main__':
    parse_book()
    write_book()

    parse_audio()
    write_audio()
    # for i in range(len(russian_titles)):
        # print(russian_titles[i])
        # print(arabic_titles[i])
        # print(russian_matns[i])
        # print(arabic_matns[i])
        # print(sharkhs[i])
        # print(questions[i])
        # for j in audios[i]:
        #     print(j)
        # print()
    print(lecturers)
    print(chapters)
