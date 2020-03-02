from xlsxwriter import Workbook
from datetime import datetime
from xlrd import open_workbook
from tkinter.filedialog import askopenfilename, askdirectory
from tkinter import Tk, Button

VISATYPES = {'ОБЫКНОВЕННАЯ ТУРИСТИЧЕСКАЯ': 'TOURIST',
             'ОБЫКНОВЕННАЯ ДЕЛОВАЯ': 'BUSINESS',
             'ОБЫКНОВЕННАЯ ЧАСТНАЯ': 'PRIVATE',
             'ОБЫКНОВЕННАЯ РАБОЧАЯ': 'WORK',
             'ОБЫКНОВЕННАЯ УЧЕБНАЯ': 'STUDY',
             'ОБЫКНОВЕННАЯ ГУМАНИТАРНАЯ': 'HUMANITARIAN',
             'ТРАНЗИТНАЯ ТР2': 'TRANSIT'
             }
NUMSOFENTRIES = {'ОДНОКРАТНАЯ': 'SINGLE',
                 'ДВУКРАТНАЯ': 'DOUBLE',
                 'МНОГОКРАТНАЯ': 'MULTI'
                 }
SERVICETYPES = {'Срочная 1 день': 'RUSH',
                'Обыкновенная 5 дней': 'REGULAR',
                'Обыкновенная 15 дней': 'USA REGULAR'
                }


class Visa:
    def __init__(self, visatype, entries, citizenship, servicetype, price, quantity, applications):
        self.visaType = visatype
        self.entries = entries
        self.citizenship = citizenship
        self.serviceType = servicetype
        self.price = int(price)
        self.quantity = int(quantity)
        self.applications = applications.replace(' ', '').split(',')
        for app in self.applications:
            if '-' in app:
                self.applications.remove(app)
                limits = app.split('-')
                start = int(limits[0])
                end = int(limits[1])
                for n in range(end - start + 1):
                    self.applications.append(str(n + start))
        self.applications = self.divide_in_pages(self.applications)

    @staticmethod
    def divide_in_pages(applications):
        for app in range(0, len(applications), 10):
            yield applications[app:app + 10]


def open_file():
    file = askopenfilename()
    return file


def is_valid_input(file):
    if file.endswith('.xls') or file.endswith('.xlsx'):
        return True
    else:
        return False


def process_file():
    while True:
        file = open_file()
        if not is_valid_input(file):
            return
        else:
            break
    wb = open_workbook(file)
    sheet = wb.sheet_by_index(0)
    data = []
    for row in sheet.get_rows():
        if row[0].value in VISATYPES.keys():
            current_row = Visa(
                row[0].value,
                row[3].value,
                row[6].value,
                row[7].value,
                row[10].value,
                row[11].value,
                row[13].value
            )
            data.append(current_row)

    file_directory = askdirectory()
    workbook = Workbook(file_directory + '/Cover Papers ' + datetime.now().strftime('%d.%m.%Y %H-%M-%S') + '.xlsx')
    format_header = workbook.add_format({'font_size': 32, 'bold': True, 'align': 'center'})
    format_applications = workbook.add_format(
        {'font_size': 24, 'bold': True, 'align': 'center', 'border': 1, 'bg_color': '#fafafa'})
    format_footer = workbook.add_format({'font_size': 28, 'bold': True, 'align': 'center'})
    worksheet = workbook.add_worksheet()
    col = 3

    for d in data:
        for page in d.applications:
            worksheet.write(0, col, d.citizenship, format_header)
            worksheet.write(1, col, VISATYPES.get(d.visaType), format_header)
            worksheet.write(2, col, NUMSOFENTRIES.get(d.entries), format_header)
            worksheet.write(3, col, SERVICETYPES.get(d.serviceType), format_header)
            row = 7
            i = 0
            for application in page:
                worksheet.write(row, col, application, format_applications)
                row += 1
                i += 1
            for j in range(10 - i):
                worksheet.write(row, col, "", format_applications)
                row += 1
            worksheet.write(20, col, str(i) + ' / ' + str(d.quantity), format_footer)
            worksheet.write(21, col, 'TOTAL', format_footer)
            worksheet.write(22, col, i * d.price, format_footer)
            worksheet.set_column(col, col, 35)
            col += 6

    workbook.close()


def main():
    window = Tk()
    window.title('Cover Papers')
    window.geometry('250x50')
    button = Button(window, text='OPEN FILE', command=process_file)
    button.pack(fill='both', expand=True)
    window.mainloop()


if __name__ == '__main__':
    main()
