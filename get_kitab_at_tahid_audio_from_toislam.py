import xlsxwriter
from requests import get
import re

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('kitab_at_tauhid_audio_toislam.xlsx')
worksheet = workbook.add_worksheet()

# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

page = get(f'https://toislam.ws/audio-aqidah/sharh_sheikha_fauzana_na_knigu_edinobojiya').content.decode('utf-8')

audios = re.findall('http://static.toislam.ws/files/audio.{10,100}mp3<', page)

for i in audios:
    worksheet.write_string(row, col, i[0:-1])
    row += 1

workbook.close()
