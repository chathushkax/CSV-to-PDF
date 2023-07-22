import math
import os
import sys
import tkinter as tk
from tkinter import filedialog, END
import csv

import numpy as np
from PIL import Image
from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QFont, QPixmap
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QComboBox, QVBoxLayout, QWidget
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
import tkinter.messagebox

pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))


def window():
    global win
    app = QApplication(sys.argv)
    win = QMainWindow()
    win.setGeometry(700, 300, 700, 500)
    win.setWindowTitle("Excel to PDF")
    win.setWindowIcon(QIcon("logo2.png"))
    win.setToolTip("Excel to pdf")

    # add image to the main window
    label = QLabel(win)
    pixmap = QPixmap('logo2_win.png')
    pixmap.scaled(32, 32, Qt.KeepAspectRatio, Qt.FastTransformation)
    label.setPixmap(pixmap)
    label.resize(300, 200)
    label.move(10, -30)

    label2 = QLabel(win)
    pixmap = QPixmap('logo_win.png')
    pixmap.scaled(32, 32, Qt.KeepAspectRatio, Qt.FastTransformation)
    label2.setPixmap(pixmap)
    label2.resize(300, 200)
    label2.move(200, -40)

    # page size
    page_width = 3508
    page_height = 2480
    margin = 75
    # company_name = 'rettrack'

    # import company's logo
    im = Image.open('logo.png')
    width, height = im.size
    ratio = width / height
    image_width = 400
    image_height = int(image_width / ratio)

    # import company's logo2
    im = Image.open('logo2.png')
    width2, height2 = im.size
    ratio2 = width2 / height2
    image_width2 = 200
    image_height2 = int(image_width2 / ratio2)

    # file xlsx choose
    def choose_file():

        global file_path
        global file_path_download
        # global file_folder

        root = tk.Tk()
        root.withdraw()

        file_path = filedialog.askopenfilename()
        # specific path add
        # file_path_download = file_path.split(':')
        # file_path_download2 = file_path_download[0]
        # file_path_download1 = file_path_download[1].split('/')
        # file_path_download = file_path_download2 + ':' + '\\' + file_path_download1[1] + '\\' + file_path_download1[2]
        # print(file_path_download, 'inside')

        if file_path != '':

            # show message
            tkinter.messagebox.showinfo("Intertek", "File selected" + ":" + " " + file_path)
        else:

            # show message
            tkinter.messagebox.showinfo("Intertek", "Please choose csv file")

        return file_path

    # import excel/csv file

    # file choose button
    btn_choose = QtWidgets.QPushButton(win)
    btn_choose.setText('choose .csv file')
    btn_choose.clicked.connect(choose_file)
    btn_choose.move(250, 250)
    btn_choose.resize(200, 50)

    # txt_voy_no_values = QtWidgets.QLineEdit(win)
    # txt_voy_no_values.move(250, 180)
    # txt_voy_no_values.resize(400, 25)

    # def dropdown():
    #     # print("value is:" + variable.get())
    #     # dropdown menu
    #     OPTIONS = [
    #         "Jan",
    #         "Feb",
    #         "Mar"
    #     ]  # etc
    #
    #     master = tk.Tk()
    #
    #     variable = tk.StringVar(master)
    #     variable.set(OPTIONS[0])  # default value
    #
    #     w = tk.OptionMenu(master, variable, *OPTIONS)
    #     w.pack()
    #
    #     tk.mainloop()

    combobox1 = QComboBox()
    combobox1.addItem('WWL')
    combobox1.addItem('HOEGH')

    layout = QVBoxLayout()
    layout.addWidget(combobox1)

    container = QWidget(win)
    container.setLayout(layout)
    container.move(250, 300)
    container.resize(200, 75)

    content = combobox1.currentText()
    print(content)

    # combobox1.currentIndexChanged.connect(current_text_changed)
    #
    # def current_text_changed( s):
    #     print("Current text: ", s)

    # win.setCentralWidget(container)
    # btn_dropdown = QtWidgets.QPushButton(win)
    # btn_dropdown.setText('Dropdown')
    # btn_dropdown.clicked.connect(dropdown)
    # btn_dropdown.move(250, 300)
    # btn_dropdown.resize(200, 50)

    # btn_dropdown.pack()

    def create_invoice():
        print('file in', file_path)

        content_type = combobox1.currentText()
        print(content_type)

        vessel = []
        voy_no = []
        date = []
        survey_location = []
        vin_id = []
        damage_location = []
        damage_details = []
        cause = []
        with open(file_path, 'r') as f:
            mycsv = csv.DictReader(f)
            for col in mycsv:

                vessel.append(col['Title Page_Vessel '])
                voy_no.append(col['Title Page_Voyage Number'])
                date.append(col['Title Page_Inspection Date'])
                survey_location.append(col['Title Page_Location'])
                vin_id.append(col['Inspection report_Damaged car _VIN ID '])
                damage_location.append(col['Inspection report_Damaged car _Damaged location'])
                damage_details.append(col['Inspection report_Damaged car _Damage details - Severity'])
                cause.append(col['Inspection report_Damaged car _Cause'])
                print(vessel[0])

                # split into an arrays
                vin_id = vin_id[0].replace('"', '')
                vin_id = vin_id.split(',')
                damage_location = damage_location[0].replace('"', '')
                damage_location = damage_location.split(',')
                damage_details = damage_details[0].replace('"', '')
                damage_details = damage_details.split(',')
                cause = cause[0].replace('"', '')
                cause = cause.split(',')
                print(vin_id)

                # check the data availability in the Excel sheet
                chk_vessel = vessel[0]
                if chk_vessel != '':

                    # check the remaining data
                    chk_remain_data = vessel[0]
                    if chk_remain_data != '':

                        vessel = vessel[0]
                        voy_no = voy_no[0]
                        date = date[0]
                        survey_location = survey_location[0]
                        vin_id = vin_id
                        damage_location = damage_location
                        damage_details = damage_details
                        cause = cause
                        # split date
                        date = date.split(':')
                        date = date[0]
                        date = date.split('-')
                        date1 = list(date[2])
                        year = date[0]
                        month = date[1]
                        date = date1[0] + date1[1]
                        date = year + '-' + month + '-' + date

                        # page numbers checking
                        arr_page_chk = vin_id

                        if len(arr_page_chk) <= 12:
                            print('data below 12 rows')
                            global file_folder
                            page_no = math.ceil(len(arr_page_chk) / 12)
                            page_no = math.ceil(page_no)
                            # creating pdf
                            pdf_name = str(vessel) + '_' + str(page_no) + '_' + date + '.pdf'
                            # save_name = os.path.join(file_path_download + '\\Downloads', pdf_name)

                            # choose directory
                            root = tk.Tk()
                            root.withdraw()
                            file_folder = filedialog.askdirectory()

                            if file_folder != '':
                                save_name = os.path.join(file_folder, pdf_name)
                                print(save_name, 'not none')
                                c = canvas.Canvas(save_name)
                                c.setPageSize((page_width, page_height))

                                # user ask to remove
                                # c.setFont('Arial', 35)
                                # c.drawString(2750, 2400, 'Intertek Testing Service ')
                                # c.drawString(2750, 2350, 'Australia (Pty) Ltd ')
                                # c.drawString(2750, 2300, 'Christy Drive ')
                                # c.drawString(2750, 2250, 'Port Kembla Gateway ')
                                # c.drawString(2750, 2200, '(South Wing) ')
                                # c.drawString(2750, 2150, 'Port Kembla	 ')
                                # c.drawString(2750, 2100, 'NSW 2505 Australia')
                                #
                                # c.drawString(3170, 2400, 'Tel +61 2 4274 5335 ')
                                # c.drawString(3170, 2350, 'Fax +81 2 4274 3775 ')
                                # y -= margin

                                ROOT_DIR = os.path.abspath(os.curdir)
                                # Drawing the image2
                                c.drawInlineImage(ROOT_DIR + "\\logo2.png",
                                                  3250,
                                                  50,
                                                  image_width2, image_height2)

                                # Drawing the image1
                                c.drawInlineImage(ROOT_DIR + "\\logo.png",
                                                  100,
                                                  2200,
                                                  image_width, image_height)

                                c.setFont('Arial', 80)
                                c.setFillColorRGB(0.5, 0.5, 0.5)  # choose fill colour
                                c.rect(12 * inch, 30.29 * inch, 25 * inch, 2 * inch, fill=1)  # draw rectangle
                                c.setFillColorRGB(0, 0, 0)
                                text = 'CHRYSLER VEHICLES INSPECTION REPORT'
                                text_width = stringWidth(text, 'Arial', 80)
                                c.drawString((page_width - text_width) / 2, page_height - image_height - margin, text)
                                y = page_height - image_height - margin * 4
                                x = 2 * margin
                                x2 = x + 550
                                x3 = x + 350
                                x4 = x3 + 430
                                x5 = x4 + 480
                                x6 = x5 + 620
                                x7 = x6 + 540

                                c.setFillColorRGB(0, 0, 0)
                                c.setFont('Arial', 40)
                                c.drawString(x, y, 'Vessel ')
                                c.drawString(x2, y, ': ' + vessel)
                                y -= margin

                                c.drawString(x, y, 'Voy. No ')
                                c.drawString(x2, y, ': ' + voy_no)
                                y -= margin

                                c.drawString(x, y, 'Date ')
                                c.drawString(x2, y, ': ' + str(date))
                                y -= margin

                                c.drawString(x, y, 'Survey Location: ')
                                c.drawString(x2, y, ': ' + survey_location)
                                y -= 200

                                c.setFont('Arial', 40)
                                c.setFillColorRGB(0, 0, 0)
                                c.drawString(170, y, 'No')
                                c.drawString(x3, y, 'VIN ID')
                                c.drawString(x4, y, 'Damage Location')
                                c.drawString(x5, y, 'Damage Details-Severity')
                                c.drawString(x6, y, 'Cause')
                                c.drawString(x7, y, 'Remark by '+content_type+' Surveyor')
                                y -= margin

                                arr_length = vin_id

                                # split into array
                                arr_vin_id = vin_id
                                arr_damage_location = damage_location
                                arr_damage_details = damage_details
                                arr_cause = cause
                                listToStr = ','.join(map(str, [x for x in np.arange(1, len(arr_length) + 1)]))
                                listToStr = listToStr.split(',')
                                # print(listToStr, arr_vin_id)
                                # add array values to pdf
                                c.setFillColorRGB(0.5, 0.5, 0.5)
                                c.setFont('Arial', 30)

                                data = []

                                # constant data for table

                                constant_data = ['',
                                                 # '',
                                                 '',
                                                 '',
                                                 '',
                                                 '',
                                                 '',
                                                 '',
                                                 '',
                                                 '',
                                                 '',
                                                 '',
                                                 '',
                                                 '',
                                                 '',
                                                 '',
                                                 '',
                                                 '',
                                                 '']
                                data.append(constant_data)

                                print('range below 12')
                                for j in range(len(arr_vin_id) % 12):
                                    # c.drawString(x, y, listToStr[12 * (page_no - k - 1) + j])
                                    # # c.drawString(x, y, listToStr[j+1])
                                    # c.drawString(x3, y, arr_vin_id[12 * (page_no - k - 1) + j])
                                    # c.drawString(x4, y, arr_damage_location[12 * (page_no - k - 1) + j])
                                    # c.drawString(x5, y, arr_damage_details[12 * (page_no - k - 1) + j])
                                    # c.drawString(x6, y, arr_cause[12 * (page_no - k - 1) + j])
                                    # y -= margin

                                    table_row_data = [listToStr[j],
                                                      # '',
                                                      arr_vin_id[j],
                                                      '',
                                                      '',
                                                      arr_damage_location[j],
                                                      '',
                                                      '',
                                                      arr_damage_details[j],
                                                      '',
                                                      '',
                                                      arr_cause[j],
                                                      '',
                                                      '',
                                                      '',
                                                      '',
                                                      '',
                                                      '',
                                                      '']
                                    # print(table_row_data)

                                    data.append(table_row_data)

                                # data append for empty cells
                                for j in range(12 - (len(arr_vin_id) % 12)):
                                    table_row_data = ['',
                                                      # '',
                                                      '',
                                                      '',
                                                      '',
                                                      '',
                                                      '',
                                                      '',
                                                      '',
                                                      '',
                                                      '',
                                                      '',
                                                      '',
                                                      '',
                                                      '',
                                                      '',
                                                      '',
                                                      '',
                                                      '']
                                    data.append(table_row_data)

                                print(data, 'data')
                                t = Table(data, 18 * [2.5 * inch], 13 * [1 * inch])
                                t.setStyle(TableStyle([('ALIGN', (0, 0), (13, 18), 'CENTER'),
                                                       ('VALIGN', (0, 0), (13, 18), 'MIDDLE'),
                                                       ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
                                                       ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                                                       ('FONTSIZE', (0, 0), (0, 18), 40, colors.skyblue),
                                                       ('TEXTCOLOR', (0, 0), (0, 18), colors.black),
                                                       ('FONTSIZE', (0, 1), (13, 18), 30, colors.skyblue),
                                                       ('TEXTCOLOR', (0, 1), (13, 18), colors.dimgrey),
                                                       ('SPAN', (1, 0), (3, 0)), ('SPAN', (4, 0), (6, 0)),
                                                       ('SPAN', (7, 0), (9, 0)), ('SPAN', (10, 0), (11, 0)),
                                                       ('SPAN', (12, 0), (17, 0)),
                                                       ('SPAN', (1, 1), (3, 1)), ('SPAN', (4, 1), (6, 1)),
                                                       ('SPAN', (7, 1), (9, 1)), ('SPAN', (10, 1), (11, 1)),
                                                       ('SPAN', (12, 1), (17, 1)),
                                                       ('SPAN', (1, 2), (3, 2)), ('SPAN', (4, 2), (6, 2)),
                                                       ('SPAN', (7, 2), (9, 2)), ('SPAN', (10, 2), (11, 2)),
                                                       ('SPAN', (12, 2), (17, 2)),
                                                       ('SPAN', (1, 3), (3, 3)), ('SPAN', (4, 3), (6, 3)),
                                                       ('SPAN', (7, 3), (9, 3)), ('SPAN', (10, 3), (11, 3)),
                                                       ('SPAN', (12, 3), (17, 3)),
                                                       ('SPAN', (1, 4), (3, 4)), ('SPAN', (4, 4), (6, 4)),
                                                       ('SPAN', (7, 4), (9, 4)), ('SPAN', (10, 4), (11, 4)),
                                                       ('SPAN', (12, 4), (17, 4)),
                                                       ('SPAN', (1, 5), (3, 5)), ('SPAN', (4, 5), (6, 5)),
                                                       ('SPAN', (7, 5), (9, 5)), ('SPAN', (10, 5), (11, 5)),
                                                       ('SPAN', (12, 5), (17, 5)),
                                                       ('SPAN', (1, 6), (3, 6)), ('SPAN', (4, 6), (6, 6)),
                                                       ('SPAN', (7, 6), (9, 6)), ('SPAN', (10, 6), (11, 6)),
                                                       ('SPAN', (12, 6), (17, 6)),
                                                       ('SPAN', (1, 7), (3, 7)), ('SPAN', (4, 7), (6, 7)),
                                                       ('SPAN', (7, 7), (9, 7)), ('SPAN', (10, 7), (11, 7)),
                                                       ('SPAN', (12, 7), (17, 7)),
                                                       ('SPAN', (1, 8), (3, 8)), ('SPAN', (4, 8), (6, 8)),
                                                       ('SPAN', (7, 8), (9, 8)), ('SPAN', (10, 8), (11, 8)),
                                                       ('SPAN', (12, 8), (17, 8)),
                                                       ('SPAN', (1, 9), (3, 9)), ('SPAN', (4, 9), (6, 9)),
                                                       ('SPAN', (7, 9), (9, 9)), ('SPAN', (10, 9), (11, 9)),
                                                       ('SPAN', (12, 9), (17, 9)),
                                                       ('SPAN', (1, 10), (3, 10)),
                                                       ('SPAN', (4, 10), (6, 10)),
                                                       ('SPAN', (7, 10), (9, 10)),
                                                       ('SPAN', (10, 10), (11, 10)),
                                                       ('SPAN', (12, 10), (17, 10)),
                                                       ('SPAN', (1, 11), (3, 11)),
                                                       ('SPAN', (4, 11), (6, 11)),
                                                       ('SPAN', (7, 11), (9, 11)),
                                                       ('SPAN', (10, 11), (11, 11)),
                                                       ('SPAN', (12, 11), (17, 11)),
                                                       ('SPAN', (1, 12), (3, 12)),
                                                       ('SPAN', (4, 12), (6, 12)),
                                                       ('SPAN', (7, 12), (9, 12)),
                                                       ('SPAN', (10, 12), (11, 12)),
                                                       ('SPAN', (12, 12), (17, 12)),
                                                       ]))

                                width = 800
                                height = 100
                                x = 100
                                y = 700
                                t.wrapOn(c, width, height)
                                t.drawOn(c, x, y)
                                c.setFillColorRGB(0, 0, 0)
                                c.setFont('Arial', 35)
                                c.drawString(x, 3 * y / 4, 'NOTE ')
                                c.drawString(x2, 3 * y / 4,
                                             ': The vehicle were generally noted to be covered by travel dust and '
                                             'with dirt /water stains/ sand coat')
                                y -= margin
                                c.drawString(x, 2 * y / 3, 'METHOD ')
                                c.drawString(x2, 2 * y / 3, ': CHRYSLER-FCA, VEHICLE SHIPPING MANUAL 2019 PROCEDURES')
                                y -= margin
                                c.drawString(x, 2 * y / 3, 'SST ')
                                c.drawString(x2, 2 * y / 3,
                                             ': Represents damage that could have occurred at some stage of Transportation'
                                             ' between factory and the first point of rest after discharge from the vessel.')
                                y -= margin

                                c.setFont('Arial', 35)
                                c.drawString(x, y / 4,
                                             'Surveyor of Intertek ….........................................')
                                c.setFont('Arial', 35)
                                c.drawString(15 * x, y / 4,
                                             'Surveyor of '+content_type+' ….........................................')

                                c.setFillColorRGB(0.5, 0.5, 0.5)
                                c.setFont('Arial', 25)
                                c.drawString(x, y / 10, 'Intertek Testing Services Australia (Pty) Ltd')
                                c.setFont('Arial', 25)
                                c.drawString(7.5 * x, y / 10, 'ABN: 56 001 722 854')

                                # page numbers
                                c.setFillColorRGB(0.5, 0.5, 0.5)
                                c.setFont('Arial', 25)
                                c.drawString(3000, y / 10, 'page 1 of')
                                c.setFont('Arial', 25)
                                c.drawString(3110, y / 10, '1')

                                c.save()

                                # import messagebox from tkinter module
                                import tkinter.messagebox

                                # show message
                                tkinter.messagebox.showinfo("Intertek",
                                                            "Files are successfully downloaded" + " to :" + file_folder)
                                os.startfile(file_folder)

                            else:
                                # show message
                                tkinter.messagebox.showinfo("Intertek", "Please select a download folder")

                        else:
                            print('data above 12 rows')
                            page_no = len(arr_page_chk) / 12
                            page_no = math.ceil(page_no)

                            # choose directory
                            root = tk.Tk()
                            root.withdraw()
                            file_folder = filedialog.askdirectory()

                            if file_folder != '':
                                print('page no', page_no)
                                for k in range(page_no):
                                    # creating pdf
                                    pdf_name = str(vessel) + '_' + str(page_no - k) + '_' + date + '.pdf'
                                    # save_name = os.path.join(file_path_download + '\\Downloads', pdf_name)

                                    save_name = os.path.join(file_folder, pdf_name)
                                    print(save_name)
                                    c = canvas.Canvas(save_name)
                                    c.setPageSize((page_width, page_height))

                                    # user ask to remove
                                    # c.setFont('Arial', 35)
                                    # c.drawString(2750, 2400, 'Intertek Testing Service ')
                                    # c.drawString(2750, 2350, 'Australia (Pty) Ltd ')
                                    # c.drawString(2750, 2300, 'Christy Drive ')
                                    # c.drawString(2750, 2250, 'Port Kembla Gateway ')
                                    # c.drawString(2750, 2200, '(South Wing) ')
                                    # c.drawString(2750, 2150, 'Port Kembla	 ')
                                    # c.drawString(2750, 2100, 'NSW 2505 Australia')
                                    #
                                    # c.drawString(3170, 2400, 'Tel +61 2 4274 5335 ')
                                    # c.drawString(3170, 2350, 'Fax +81 2 4274 3775 ')
                                    # y -= margin

                                    ROOT_DIR = os.path.abspath(os.curdir)
                                    # print('root', ROOT_DIR)
                                    # Drawing the image2
                                    c.drawInlineImage(ROOT_DIR + "\\logo2.png",
                                                      3250,
                                                      50,
                                                      image_width2, image_height2)

                                    # Drawing the image1
                                    c.drawInlineImage(ROOT_DIR + "\\logo.png",
                                                      100,
                                                      2200,
                                                      image_width, image_height)

                                    c.setFont('Arial', 80)

                                    c.setFillColorRGB(0.5, 0.5, 0.5)  # choose fill colour
                                    c.rect(12 * inch, 30.29 * inch, 25 * inch, 2 * inch, fill=1)  # draw rectangle
                                    c.setFillColorRGB(0, 0, 0)
                                    text = 'CHRYSLER VEHICLES INSPECTION REPORT'
                                    text_width = stringWidth(text, 'Arial', 80)
                                    c.drawString((page_width - text_width) / 2, page_height - image_height - margin,
                                                 text)
                                    y = page_height - image_height - margin * 4
                                    x = 2 * margin
                                    x2 = x + 550
                                    x3 = x + 350
                                    x4 = x3 + 430
                                    x5 = x4 + 480
                                    x6 = x5 + 620
                                    x7 = x6 + 540

                                    c.setFillColorRGB(0, 0, 0)
                                    c.setFont('Arial', 40)
                                    c.drawString(x, y, 'Vessel ')
                                    c.drawString(x2, y, ': ' + vessel)
                                    y -= margin

                                    c.drawString(x, y, 'Voy. No ')
                                    c.drawString(x2, y, ': ' + voy_no)
                                    y -= margin

                                    c.drawString(x, y, 'Date ')
                                    c.drawString(x2, y, ': ' + str(date))
                                    y -= margin

                                    c.drawString(x, y, 'Survey Location: ')
                                    c.drawString(x2, y, ': ' + survey_location)
                                    y -= 200

                                    c.setFont('Arial', 40)
                                    c.setFillColorRGB(0, 0, 0)
                                    c.drawString(170, y, 'No')
                                    c.drawString(x3, y, 'VIN ID')
                                    c.drawString(x4, y, 'Damage Location')
                                    c.drawString(x5, y, 'Damage Details-Severity')
                                    c.drawString(x6, y, 'Cause')
                                    c.drawString(x7, y, 'Remark by '+content_type+' Surveyor')
                                    y -= margin

                                    arr_length = vin_id

                                    # split into array
                                    arr_vin_id = vin_id
                                    arr_damage_location = damage_location
                                    arr_damage_details = damage_details
                                    arr_cause = cause
                                    listToStr = ','.join(map(str, [x for x in np.arange(1, len(arr_length) + 1)]))
                                    listToStr = listToStr.split(',')
                                    # add array values to pdf
                                    c.setFillColorRGB(0.5, 0.5, 0.5)
                                    c.setFont('Arial', 30)

                                    if k == 0 and (len(arr_vin_id)) % 12 != 0:
                                        print('range 12', len(arr_vin_id) % 12, page_no)
                                        data = []

                                        # constant data for table

                                        constant_data = ['',
                                                         # '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '']
                                        data.append(constant_data)

                                        for j in range(len(arr_vin_id) % 12):
                                            # c.drawString(x, y, listToStr[12 * (page_no - 1) + j])
                                            # # c.drawString(x, y, listToStr[j+1])
                                            # c.drawString(x3, y, arr_vin_id[12 * (page_no - 1) + j])
                                            # c.drawString(x4, y, arr_damage_location[12 * (page_no - 1) + j])
                                            # c.drawString(x5, y, arr_damage_details[12 * (page_no - 1) + j])
                                            # c.drawString(x6, y, arr_cause[12 * (page_no - 1) + j])
                                            # y -= margin

                                            table_row_data = [listToStr[12 * (page_no - 1) + j],
                                                              # '',
                                                              arr_vin_id[12 * (page_no - 1) + j],
                                                              '',
                                                              '',
                                                              arr_damage_location[12 * (page_no - 1) + j],
                                                              '',
                                                              '',
                                                              arr_damage_details[12 * (page_no - 1) + j],
                                                              '',
                                                              '',
                                                              arr_cause[12 * (page_no - 1) + j],
                                                              '',
                                                              '',
                                                              '',
                                                              '',
                                                              '',
                                                              '',
                                                              '']
                                            # print(table_row_data)
                                            data.append(table_row_data)

                                            # data append for empty cells
                                        for j in range(12 - (len(arr_vin_id) % 12)):
                                            table_row_data = ['',
                                                              # '',
                                                              '',
                                                              '',
                                                              '',
                                                              '',
                                                              '',
                                                              '',
                                                              '',
                                                              '',
                                                              '',
                                                              '',
                                                              '',
                                                              '',
                                                              '',
                                                              '',
                                                              '',
                                                              '',
                                                              '']
                                            data.append(table_row_data)
                                            print(data)

                                        t = Table(data, 18 * [2.5 * inch], 13 * [1 * inch])
                                        t.setStyle(TableStyle([('ALIGN', (0, 0), (13, 18), 'CENTER'),
                                                               ('VALIGN', (0, 0), (13, 18), 'MIDDLE'),
                                                               ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
                                                               ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                                                               ('FONTSIZE', (0, 0), (0, 18), 40, colors.skyblue),
                                                               ('TEXTCOLOR', (0, 0), (0, 18), colors.black),
                                                               ('FONTSIZE', (0, 1), (13, 18), 30, colors.skyblue),
                                                               ('TEXTCOLOR', (0, 1), (13, 18), colors.dimgrey),
                                                               ('SPAN', (1, 0), (3, 0)), ('SPAN', (4, 0), (6, 0)),
                                                               ('SPAN', (7, 0), (9, 0)), ('SPAN', (10, 0), (11, 0)),
                                                               ('SPAN', (12, 0), (17, 0)),
                                                               ('SPAN', (1, 1), (3, 1)), ('SPAN', (4, 1), (6, 1)),
                                                               ('SPAN', (7, 1), (9, 1)), ('SPAN', (10, 1), (11, 1)),
                                                               ('SPAN', (12, 1), (17, 1)),
                                                               ('SPAN', (1, 2), (3, 2)), ('SPAN', (4, 2), (6, 2)),
                                                               ('SPAN', (7, 2), (9, 2)), ('SPAN', (10, 2), (11, 2)),
                                                               ('SPAN', (12, 2), (17, 2)),
                                                               ('SPAN', (1, 3), (3, 3)), ('SPAN', (4, 3), (6, 3)),
                                                               ('SPAN', (7, 3), (9, 3)), ('SPAN', (10, 3), (11, 3)),
                                                               ('SPAN', (12, 3), (17, 3)),
                                                               ('SPAN', (1, 4), (3, 4)), ('SPAN', (4, 4), (6, 4)),
                                                               ('SPAN', (7, 4), (9, 4)), ('SPAN', (10, 4), (11, 4)),
                                                               ('SPAN', (12, 4), (17, 4)),
                                                               ('SPAN', (1, 5), (3, 5)), ('SPAN', (4, 5), (6, 5)),
                                                               ('SPAN', (7, 5), (9, 5)), ('SPAN', (10, 5), (11, 5)),
                                                               ('SPAN', (12, 5), (17, 5)),
                                                               ('SPAN', (1, 6), (3, 6)), ('SPAN', (4, 6), (6, 6)),
                                                               ('SPAN', (7, 6), (9, 6)), ('SPAN', (10, 6), (11, 6)),
                                                               ('SPAN', (12, 6), (17, 6)),
                                                               ('SPAN', (1, 7), (3, 7)), ('SPAN', (4, 7), (6, 7)),
                                                               ('SPAN', (7, 7), (9, 7)), ('SPAN', (10, 7), (11, 7)),
                                                               ('SPAN', (12, 7), (17, 7)),
                                                               ('SPAN', (1, 8), (3, 8)), ('SPAN', (4, 8), (6, 8)),
                                                               ('SPAN', (7, 8), (9, 8)), ('SPAN', (10, 8), (11, 8)),
                                                               ('SPAN', (12, 8), (17, 8)),
                                                               ('SPAN', (1, 9), (3, 9)), ('SPAN', (4, 9), (6, 9)),
                                                               ('SPAN', (7, 9), (9, 9)), ('SPAN', (10, 9), (11, 9)),
                                                               ('SPAN', (12, 9), (17, 9)),
                                                               ('SPAN', (1, 10), (3, 10)),
                                                               ('SPAN', (4, 10), (6, 10)),
                                                               ('SPAN', (7, 10), (9, 10)),
                                                               ('SPAN', (10, 10), (11, 10)),
                                                               ('SPAN', (12, 10), (17, 10)),
                                                               ('SPAN', (1, 11), (3, 11)),
                                                               ('SPAN', (4, 11), (6, 11)),
                                                               ('SPAN', (7, 11), (9, 11)),
                                                               ('SPAN', (10, 11), (11, 11)),
                                                               ('SPAN', (12, 11), (17, 11)),
                                                               ('SPAN', (1, 12), (3, 12)),
                                                               ('SPAN', (4, 12), (6, 12)),
                                                               ('SPAN', (7, 12), (9, 12)),
                                                               ('SPAN', (10, 12), (11, 12)),
                                                               ('SPAN', (12, 12), (17, 12)),
                                                               ]))
                                        print('came here')
                                        width = 800
                                        height = 100
                                        x = 100
                                        # y = 1580 - (len(arr_vin_id) % 12) * 76 + 2.3 * (len(arr_vin_id) % 12)
                                        y = 700
                                        t.wrapOn(c, width, height)
                                        t.drawOn(c, x, y)
                                    else:

                                        data = []
                                        # constant data for table

                                        constant_data = ['',
                                                         # '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '',
                                                         '']
                                        data.append(constant_data)

                                        print('range below 12')
                                        for j in range(12):
                                            # c.drawString(x, y, listToStr[12 * (page_no - k - 1) + j])
                                            # # c.drawString(x, y, listToStr[j+1])
                                            # c.drawString(x3, y, arr_vin_id[12 * (page_no - k - 1) + j])
                                            # c.drawString(x4, y, arr_damage_location[12 * (page_no - k - 1) + j])
                                            # c.drawString(x5, y, arr_damage_details[12 * (page_no - k - 1) + j])
                                            # c.drawString(x6, y, arr_cause[12 * (page_no - k - 1) + j])
                                            # y -= margin

                                            table_row_data = [listToStr[12 * (page_no - k - 1) + j],
                                                              # '',
                                                              arr_vin_id[12 * (page_no - k - 1) + j],
                                                              '',
                                                              '',
                                                              arr_damage_location[12 * (page_no - k - 1) + j],
                                                              '',
                                                              '',
                                                              arr_damage_details[12 * (page_no - k - 1) + j],
                                                              '',
                                                              '',
                                                              arr_cause[12 * (page_no - k - 1) + j],
                                                              '',
                                                              '',
                                                              '',
                                                              '',
                                                              '',
                                                              '',
                                                              '']
                                            # print(table_row_data)

                                            data.append(table_row_data)

                                        print(data)
                                        t = Table(data, 18 * [2.5 * inch], 13 * [1 * inch])
                                        t.setStyle(TableStyle([('ALIGN', (0, 0), (13, 18), 'CENTER'),
                                                               ('VALIGN', (0, 0), (13, 18), 'MIDDLE'),
                                                               ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
                                                               ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                                                               ('FONTSIZE', (0, 0), (0, 18), 40, colors.skyblue),
                                                               ('TEXTCOLOR', (0, 0), (0, 18), colors.black),
                                                               ('FONTSIZE', (0, 1), (13, 18), 30, colors.skyblue),
                                                               ('TEXTCOLOR', (0, 1), (13, 18), colors.dimgrey),
                                                               ('SPAN', (1, 0), (3, 0)), ('SPAN', (4, 0), (6, 0)),
                                                               ('SPAN', (7, 0), (9, 0)), ('SPAN', (10, 0), (11, 0)),
                                                               ('SPAN', (12, 0), (17, 0)),
                                                               ('SPAN', (1, 1), (3, 1)), ('SPAN', (4, 1), (6, 1)),
                                                               ('SPAN', (7, 1), (9, 1)), ('SPAN', (10, 1), (11, 1)),
                                                               ('SPAN', (12, 1), (17, 1)),
                                                               ('SPAN', (1, 2), (3, 2)), ('SPAN', (4, 2), (6, 2)),
                                                               ('SPAN', (7, 2), (9, 2)), ('SPAN', (10, 2), (11, 2)),
                                                               ('SPAN', (12, 2), (17, 2)),
                                                               ('SPAN', (1, 3), (3, 3)), ('SPAN', (4, 3), (6, 3)),
                                                               ('SPAN', (7, 3), (9, 3)), ('SPAN', (10, 3), (11, 3)),
                                                               ('SPAN', (12, 3), (17, 3)),
                                                               ('SPAN', (1, 4), (3, 4)), ('SPAN', (4, 4), (6, 4)),
                                                               ('SPAN', (7, 4), (9, 4)), ('SPAN', (10, 4), (11, 4)),
                                                               ('SPAN', (12, 4), (17, 4)),
                                                               ('SPAN', (1, 5), (3, 5)), ('SPAN', (4, 5), (6, 5)),
                                                               ('SPAN', (7, 5), (9, 5)), ('SPAN', (10, 5), (11, 5)),
                                                               ('SPAN', (12, 5), (17, 5)),
                                                               ('SPAN', (1, 6), (3, 6)), ('SPAN', (4, 6), (6, 6)),
                                                               ('SPAN', (7, 6), (9, 6)), ('SPAN', (10, 6), (11, 6)),
                                                               ('SPAN', (12, 6), (17, 6)),
                                                               ('SPAN', (1, 7), (3, 7)), ('SPAN', (4, 7), (6, 7)),
                                                               ('SPAN', (7, 7), (9, 7)), ('SPAN', (10, 7), (11, 7)),
                                                               ('SPAN', (12, 7), (17, 7)),
                                                               ('SPAN', (1, 8), (3, 8)), ('SPAN', (4, 8), (6, 8)),
                                                               ('SPAN', (7, 8), (9, 8)), ('SPAN', (10, 8), (11, 8)),
                                                               ('SPAN', (12, 8), (17, 8)),
                                                               ('SPAN', (1, 9), (3, 9)), ('SPAN', (4, 9), (6, 9)),
                                                               ('SPAN', (7, 9), (9, 9)), ('SPAN', (10, 9), (11, 9)),
                                                               ('SPAN', (12, 9), (17, 9)),
                                                               ('SPAN', (1, 10), (3, 10)),
                                                               ('SPAN', (4, 10), (6, 10)),
                                                               ('SPAN', (7, 10), (9, 10)),
                                                               ('SPAN', (10, 10), (11, 10)),
                                                               ('SPAN', (12, 10), (17, 10)),
                                                               ('SPAN', (1, 11), (3, 11)),
                                                               ('SPAN', (4, 11), (6, 11)),
                                                               ('SPAN', (7, 11), (9, 11)),
                                                               ('SPAN', (10, 11), (11, 11)),
                                                               ('SPAN', (12, 11), (17, 11)),
                                                               ('SPAN', (1, 12), (3, 12)),
                                                               ('SPAN', (4, 12), (6, 12)),
                                                               ('SPAN', (7, 12), (9, 12)),
                                                               ('SPAN', (10, 12), (11, 12)),
                                                               ('SPAN', (12, 12), (17, 12)),
                                                               ]))

                                        width = 800
                                        height = 100
                                        x = 100
                                        y = 700
                                        t.wrapOn(c, width, height)
                                        t.drawOn(c, x, y)

                                    c.setFillColorRGB(0, 0, 0)
                                    c.setFont('Arial', 35)
                                    c.drawString(x, 3 * y / 4, 'NOTE ')
                                    c.drawString(x2, 3 * y / 4,
                                                 ': The vehicle were generally noted to be covered by travel dust and '
                                                 'with dirt /water stains/ sand coat')
                                    y -= margin
                                    c.drawString(x, 2 * y / 3, 'METHOD ')
                                    c.drawString(x2, 2 * y / 3,
                                                 ': CHRYSLER-FCA, VEHICLE SHIPPING MANUAL 2019 PROCEDURES')
                                    y -= margin
                                    c.drawString(x, 2 * y / 3, 'SST ')
                                    c.drawString(x2, 2 * y / 3,
                                                 ': Represents damage that could have occurred at some stage of Transportation'
                                                 ' between factory and the first point of rest after discharge from the vessel.')
                                    y -= margin

                                    c.setFont('Arial', 35)
                                    c.drawString(x, y / 4,
                                                 'Surveyor of Intertek ….........................................')
                                    c.setFont('Arial', 35)
                                    c.drawString(15 * x, y / 4,
                                                 'Surveyor of '+content_type+' ….........................................')

                                    c.setFillColorRGB(0.5, 0.5, 0.5)
                                    c.setFont('Arial', 25)
                                    c.drawString(x, y / 10, 'Intertek Testing Services Australia (Pty) Ltd')
                                    c.setFont('Arial', 25)
                                    c.drawString(7.5 * x, y / 10, 'ABN: 56 001 722 854')

                                    # page numbers
                                    c.setFillColorRGB(0.5, 0.5, 0.5)
                                    c.setFont('Arial', 25)
                                    c.drawString(3000, y / 10, 'page ' + str(page_no - k) + ' of')
                                    c.setFont('Arial', 25)
                                    c.drawString(3110, y / 10, str(page_no))

                                    # success message
                                    lbl_vessel = QtWidgets.QLabel(win)
                                    lbl_vessel.setText('PDF successfully downloaded...!!')
                                    lbl_vessel.move(150, 300)
                                    lbl_vessel.resize(450, 40)
                                    lbl_vessel.setFont(QFont('Arial', 20))
                                    lbl_vessel.setStyleSheet("color: green;")

                                    # root = tk.Tk()
                                    # root.withdraw()
                                    # filePath, _ = QFileDialog.getSaveFileUrl()
                                    # print(filePath)
                                    # if file path is blank return back
                                    # if filePath == "":
                                    #     return

                                    # saving canvas at desired path
                                    c.save()

                                    if k == page_no - 1:
                                        # import messagebox from tkinter module
                                        import tkinter.messagebox

                                        # show message
                                        tkinter.messagebox.showinfo("Intertek",
                                                                    "Files are successfully downloaded" + " to :" + file_folder)
                                        os.startfile(file_folder)


                            else:
                                # import messagebox from tkinter module
                                import tkinter.messagebox

                                # show message
                                tkinter.messagebox.showinfo("Intertek", "Please select a download folder")

                    else:
                        # import messagebox from tkinter module
                        import tkinter.messagebox

                        # show message
                        tkinter.messagebox.showinfo("Intertek",
                                                    "Files are successfully downloaded" + " to :" + file_folder)
                        os.startfile(file_folder)

                else:
                    lbl_vessel = QtWidgets.QLabel(win)
                    lbl_vessel.setText('Excel sheet is empty')
                    lbl_vessel.move(200, 600)
                    lbl_vessel.resize(120, 40)

    # open file downloaded directory
    # def Openfolder():
    #     open(file_folder)

    # btn_open_folder = QtWidgets.QPushButton(win)
    # btn_open_folder.setText('Download PDF')
    # # btn_save.triggered.connect(create_invoice)
    # btn_open_folder.clicked.connect(Openfolder)
    # btn_open_folder.move(10, 400)
    # btn_open_folder.resize(50, 50)

    btn_save = QtWidgets.QPushButton(win)
    btn_save.setText('Download PDF')
    # btn_save.triggered.connect(create_invoice)
    btn_save.clicked.connect(create_invoice)
    btn_save.move(250, 400)
    btn_save.resize(200, 50)

    win.show()
    sys.exit(app.exec_())


window()
