import xlsxwriter
from requests import get

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('kitab_at_tauhid_sharkh.xlsx')
worksheet = workbook.add_worksheet()

# Some data we want to write to the worksheet.
chapters = []

for i in range(1, 68):
    sharkh_page = get(f'https://daura.com/kifayatul_mustazid_glava_{i}/').content.decode('utf-8')
    sharkh_page = sharkh_page.replace('﴿', '<bdo dir="rtl">﴿')
    sharkh_page = sharkh_page.replace('﴾', '﴾</bdo>')
    sharkh_start = sharkh_page.find('</blockquote>') + len('</blockquote>')
    if sharkh_start < len('</blockquote>'):
        sharkh_start = sharkh_page.find('<p><strong>Шейх')
    if sharkh_start < len('</blockquote>'):
        print(f'sharkh_start: {i}')
    sharkh_end = sharkh_page.find('РАССМАТРИВАЕМЫЕ ВОПРОСЫ:', sharkh_start)
    if sharkh_end < 0:
        sharkh_end = sharkh_page.find('</div></div></div></div></div></div></div></div>', sharkh_start)
    if sharkh_end < 0:
        print(f'sharkh_end_1: {i}')
    sharkh_end = sharkh_page.rfind('</p>', sharkh_start, sharkh_end) + len('</p>')
    if sharkh_end < len('</p>'):
        print(f'sharkh_end_2: {i}')
    chapters.append(sharkh_page[sharkh_start:sharkh_end])


# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

# Iterate over the data and write it out row by row.
for item in chapters:
    worksheet.write(row, col,     item)
    row += 1

workbook.close()