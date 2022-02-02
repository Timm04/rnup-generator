# Importing libraries
import string
import random
import xlsxwriter

# Excel file name
workbook = xlsxwriter.Workbook("Mappe1.xlsx")

# Defining worksheet name in Excel
worksheet = workbook.add_worksheet('Daten')

# Defining colum size
worksheet.set_column_pixels(0, 100, 93)

# Defining row size
worksheet.set_default_row(13)
worksheet.set_row(4, 14)
workbook.formats[0].set_font_size(10)

# Adding formating for data headings
data_heading_format = workbook.add_format({'bold': True, 'align': 'center', 'font_size': 11})
data_heading_format.set_font_name('Arial')
data_heading_format.set_align('vcenter')

# Adding formating for number, username and password
nup_format = workbook.add_format({'align': 'center', 'font_size': 10})
nup_format.set_font_name('Arial')
nup_format.set_align('vcenter')

# Worksheet layout
worksheet.set_page_view()
worksheet.set_landscape()
worksheet.set_paper(9)

# Adding header and footer
worksheet.set_header('&C &26 &"bold" KURS TG11')
worksheet.set_footer('&C&D')

# Adding first data row headings
worksheet.write('B3', 'TG11/1', data_heading_format)
worksheet.write('A5', 'Number',  data_heading_format)
worksheet.write('B5', 'Username',  data_heading_format)
worksheet.write('C5', 'Password',  data_heading_format)

# Adding second data row headings
worksheet.write('F3', 'TG11/2', data_heading_format)
worksheet.write('E5', 'Number',  data_heading_format)
worksheet.write('F5', 'Username',  data_heading_format)
worksheet.write('G5', 'Password',  data_heading_format)

# Adding third data row headings
worksheet.write('J3', 'TG11/3', data_heading_format)
worksheet.write('I5', 'Number',  data_heading_format)
worksheet.write('J5', 'Username',  data_heading_format)
worksheet.write('K5', 'Password',  data_heading_format)

# Defining rowIndex for moving down in rows
rowIndex = 6

# Defining numbering for visibility
numbering = 1

#Setting up loop of 30
# Runs through password generation (9 random numbers; 650362561)

# Loop start
for row in range(30):
    password = f'{random.randint(100, 999)}{random.randint(100, 999)}{random.randint(100, 999)}'

    # Runs through username generation (9 random letters and numbers; Ote6wEm2x)
    def random_string(length=9, uppercase=True, lowercase=True, numbers=True):
        character_set = ''

        if uppercase:
            character_set += string.ascii_uppercase
        if lowercase:
            character_set += string.ascii_lowercase
        if numbers:
            character_set += string.digits

        return ''.join(random.choice(character_set) for i in range(length))

    # We start at row 6, because of rowIndex we defined earlier, which gets increased by +1 after each go through, just like numbering. Lastly we add formating with nup_format
    worksheet.write('A' + str(rowIndex), numbering, nup_format)
    worksheet.write('B' + str(rowIndex), random_string(9), nup_format)
    worksheet.write('C' + str(rowIndex), password, nup_format)

    worksheet.write('E' + str(rowIndex), numbering, nup_format)
    worksheet.write('F' + str(rowIndex), random_string(9), nup_format)
    worksheet.write('G' + str(rowIndex), password, nup_format)

    worksheet.write('I' + str(rowIndex), numbering, nup_format)
    worksheet.write('J' + str(rowIndex), random_string(9), nup_format)
    worksheet.write('K' + str(rowIndex), password, nup_format)

    # Defining numbering and rowIndex
    numbering +=1
    rowIndex +=1

# Loop end
print([f'{random_string(9)}: {random.randint(100, 999)}{random.randint(100, 999)}{random.randint(100, 999)}' for _ in range(30)])

workbook.close()
