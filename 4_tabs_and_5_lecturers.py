import xlrd
from requests import head

russian_headers = []
arabic_headers = []
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
        russian_headers.append(row[0])
        arabic_headers.append(row[1])
        russian_matns.append(row[2])
        arabic_matns.append(row[3])
        sharkhs.append(row[4])
        questions.append(row[5])


def parse_audio():
    rb = xlrd.open_workbook('audio.xlsx')
    sheet = rb.sheet_by_index(0)
    for rownum in range(0, 1):
        row = sheet.row_values(rownum)
        for lecturers_rownum in range(1, 6):
            lecturers.append(row[lecturers_rownum])
    for rownum in range(1, sheet.nrows):
        row = sheet.row_values(rownum)
        lecturer_audios = []
        for lecturers_rownum in range(1, 6):
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
        book.write("List<String> russianHeaders = [\n")
        for i in range(len(russian_headers)):
            book.write(f'"""{russian_headers[i]}""",\n')
        book.write("];\n\n")

        book.write("List<String> arabicHeaders = [\n")
        for i in range(len(arabic_headers)):
            book.write(f'"""{arabic_headers[i]}""",\n')
        book.write("];\n\n")

        book.write("List<String> russianMatns = [\n")
        for i in range(len(russian_matns)):
            book.write(f'"""{russian_matns[i]}""",\n')
        book.write("];\n\n")

        book.write("List<String> arabicMatns = [\n")
        for i in range(len(arabic_matns)):
            book.write(f'"""{arabic_matns[i]}""",\n')
        book.write("];\n\n")

        book.write("List<String> sharkhs = [\n")
        for i in range(len(sharkhs)):
            book.write(f'"""{sharkhs[i]}""",\n')
        book.write("];\n\n")

        book.write("List<String> questions = [\n")
        for i in range(len(questions)):
            book.write(f'"""{questions[i]}""",\n')
        book.write("];\n\n")


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
                    if i == 0:
                        lectures.write(f'audioName: "Введение, лекция {k + 1}", ')
                    else:
                        lectures.write(f'audioName: "Глава {i}, лекция {k + 1}", ')
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
        share_book.write("List<String> shareRussianMatns = [\n")
        for i in range(len(share_russian_matns)):
            share_book.write(f'"""{share_russian_matns[i]}""",\n')
        share_book.write("];\n\n")

        share_book.write("List<String> shareArabicMatns = [\n")
        for i in range(len(share_arabic_matns)):
            share_book.write(f'"""{share_arabic_matns[i]}""",\n')
        share_book.write("];\n\n")

        share_book.write("List<String> shareSharkhs = [\n")
        for i in range(len(share_sharkhs)):
            share_book.write(f'"""{share_sharkhs[i]}""",\n')
        share_book.write("];\n\n")

        share_book.write("List<String> shareQuestions = [\n")
        for i in range(len(share_questions)):
            share_book.write(f'"""{share_questions[i]}""",\n')
        share_book.write("];\n\n")


def write_structure():
    with open('structure.dart', 'w', encoding='utf-8') as structure:
        structure.write("import 'book.dart';\n")
        structure.write("import 'share_book.dart';\n")
        structure.write("import '../util/util_classes.dart';\n\n")
        structure.write("List<Chapter> chapters = [\n")

        for i in range(len(russian_headers)):
            structure.write(f'Chapter(\n')
            structure.write(f'russianHeader: russianHeaders[{i}],\n')
            structure.write(f'arabicHeader: arabicHeaders[{i}],\n')
            structure.write(f'tabList: [\n')
            structure.write(f'TabDescription(text: russianMatns[{i}], '
                            f'shareText: shareRussianMatns[{i}]),\n')
            structure.write(f'TabDescription(text: arabicMatns[{i}], isArabic: true, '
                            f'shareText: shareArabicMatns[{i}]),\n')
            structure.write(f'TabDescription(text: sharkhs[{i}], '
                            f'shareText: shareSharkhs[{i}]),\n')
            structure.write(f'TabDescription(text: questions[{i}], '
                            f'shareText: shareQuestions[{i}]),\n')
            structure.write(f'TabDescription(isAudio: true\n')
            structure.write('),\n')
            structure.write('],\n')
            structure.write('),\n')
        structure.write("];")


if __name__ == '__main__':
    parse_share_book()
    write_share_book()

    parse_book()
    write_book()

    parse_audio()
    write_audio()

    write_structure()

    # for i in range(len(russian_headers)):
        # print(russian_headers[i])
        # print(arabic_headers[i])
        # print(russian_matns[i])
        # print(arabic_matns[i])
        # print(sharkhs[i])
        # print(questions[i])
        # for j in audios[i]:
        #     print(j)
        # print()
    print(lecturers)
    print(chapters)
