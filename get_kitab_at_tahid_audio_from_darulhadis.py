import xlsxwriter
from requests import get
import re

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('kitab_at_tauhid_audio_darulhadis.xlsx')
worksheet = workbook.add_worksheet()

# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

page = get(f'http://darulhadis.ws/03-fath-ul-mazhid-sharh-kitab-ut-tauhid').content.decode('utf-8')

chapter_start = 2
while chapter_start >= 2:
    chapter_end = page.find('глава ', chapter_start)
    chapter = page[chapter_start:chapter_end]

    audios = re.findall('http://darulhadis.ws.{5,50}mp3<', chapter)

    print(row)
    print(chapter_start)
    print(chapter_end)
    print(audios)
    print(len(audios))
    print('\n')

    worksheet.write_string(row, col, '\n'.join(audio[0:-1] for audio in audios))
    row += 1
    chapter_start = chapter_end + 2

workbook.close()
