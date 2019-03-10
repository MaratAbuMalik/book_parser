import xlsxwriter
from requests import get
import re

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('kitab_at_tauhid_audio_daura.xlsx')
worksheet = workbook.add_worksheet()

# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

lectures = []

for i in range(1, 68):
    page = get(f'https://daura.com/kifayatul_mustazid_glava_{i}/').content.decode('utf-8')

    podster_audios = re.findall('https://toislam.podster.{10,50}ap=0', page)

    audios = []

    for audio in podster_audios:
        audio_page = get(audio).content.decode('utf-8')
        audio_uris = re.findall('https://toislam.podster.{10,50}mp3', audio_page)
        audios.append(audio_uris[0])

    print(i)
    print(audios)
    print(len(audios))
    print('\n')

    lectures.append(audios)

    # Iterate over the data and write it out row by row.
    worksheet.write_string(row, col, '\n'.join(audios))
    row += 1

workbook.close()
