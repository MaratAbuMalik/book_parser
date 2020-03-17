import xlrd
from requests import head

audios = []
audio_names = []
lecturers = []
lectures = []


def parse_audio():
    rb = xlrd.open_workbook('audio.xlsx')
    sheet = rb.sheet_by_index(0)

    for colnum in range(0, sheet.ncols):
        col = sheet.col_values(colnum)
        print(col)
        if colnum % 2:
            lectures.append(col[0])
            audio_names.append(col[1:])
        else:
            lecturers.append(col[0])
            audios.append(col[1:])


def write_audio():
    with open('audio.dart', 'w', encoding='utf-8') as audios_dart:
        audios_dart.write("import 'package:educational_audioplayer/player.dart';\n\n")

        audios_dart.write('List<String> lecturerNames = [')
        for i in lecturers:
            audios_dart.write(f'"{i}",')
        audios_dart.write("];\n\n")

        audios_dart.write('List<String> lectureSources = [')
        for i in lectures:
            audios_dart.write(f'"{i}",')
        audios_dart.write("];\n\n")

        audios_dart.write("List audios = [\n")
        for i in range(len(audios)):
            audios_dart.write(f'[\n')
            for j in range(len(audios[i])):
                if audios[i][j]:
                    audios_dart.write('Audio(')
                    audios_dart.write(f'url: "{audios[i][j]}", ')
                    audios_dart.write(f'audioName: "{audio_names[i][j]}", ')
                    audios_dart.write(f'audioSize: {round(int(head(audios[i][j], allow_redirects=True).headers["Content-Length"]) / 1024 / 1024)}, ')
                    audios_dart.write(f'audioDescription: "", ')
                    audios_dart.write(f'chapterName: lectureSources[{i}], ')
                    audios_dart.write(f'authorName: lecturerNames[{i}], ')
                    audios_dart.write('), \n')
            audios_dart.write('],\n')
        audios_dart.write("];")


if __name__ == '__main__':
    parse_audio()
    write_audio()
    print(lecturers)
    print()
    print(lectures)
    print()
    print(audios)
    print()
    print(audio_names)
