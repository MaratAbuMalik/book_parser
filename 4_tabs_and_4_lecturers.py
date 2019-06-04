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


share_russian_titles = []
share_arabic_titles = []
share_russian_matns = []
share_arabic_matns = []
share_sharkhs = []
share_questions = []


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


def parse_share_book():
    rb = xlrd.open_workbook('share_book.xlsx')
    sheet = rb.sheet_by_index(0)
    for rownum in range(1, sheet.nrows):
        row = sheet.row_values(rownum)
        share_russian_titles.append(row[0])
        share_russian_matns.append(row[0] + '\n\n' + row[2])
        share_arabic_matns.append(row[1] + '\n\n' + row[3])
        share_sharkhs.append(row[0] + '\n\n' + row[4])
        share_questions.append(row[0] + '\n\n' + row[5])


def write_book():
    with open('book.dart', 'w', encoding='utf-8') as book:
        book.write("import 'share_book.dart';\n")
        book.write("import '../util/chapter_class.dart';\n\n")
        book.write("List<Chapter> chapters = [\n")

        for i in range(len(russian_titles)):
            book.write(f'Chapter(\n')
            book.write(f'russianHeader: """{russian_titles[i]}""",\n')
            book.write(f'arabicHeader: """{arabic_titles[i]}""",\n')
            book.write(f'tabList: [\n')
            book.write(f'TabDescription(text: """{russian_matns[i]}""", shareText: share_russian_matns[{i}]),\n')
            book.write(f'TabDescription(text: """{arabic_matns[i]}""", isArabic: true, shareText: share_arabic_matns[{i}]),\n')
            book.write(f'TabDescription(text: """{sharkhs[i]}""", shareText: share_sharkhs[{i}]),\n')
            book.write(f'TabDescription(text: """{questions[i]}""", shareText: share_questions[{i}]),\n')
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


def write_share_book():
    with open('share_book.dart', 'w', encoding='utf-8') as share_book:
        share_book.write("List<String> share_russian_matns = [\n")
        for i in range(len(share_russian_matns)):
            share_book.write(f'"""{share_russian_matns[i]}""",\n')
        share_book.write("];\n\n")

        share_book.write("List<String> share_arabic_matns = [\n")
        for i in range(len(share_arabic_matns)):
            share_book.write(f'"""{share_arabic_matns[i]}""",\n')
        share_book.write("];\n\n")

        share_book.write("List<String> share_sharkhs = [\n")
        for i in range(len(share_sharkhs)):
            share_book.write(f'"""{share_sharkhs[i]}""",\n')
        share_book.write("];\n\n")

        share_book.write("List<String> share_questions = [\n")
        for i in range(len(share_questions)):
            share_book.write(f'"""{share_questions[i]}""",\n')
        share_book.write("];\n\n")


if __name__ == '__main__':
    parse_share_book()
    write_share_book()

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
