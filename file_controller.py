import os.path
import xlsxwriter
import datetime



workbook_one = xlsxwriter.Workbook("report/img_one.xlsx")
worksheet_one = workbook_one.add_worksheet()
img_one_list = []

workbook_two = xlsxwriter.Workbook("report/img_two.xlsx")
worksheet_two = workbook_two.add_worksheet()
img_two_list = []

workbook_three = xlsxwriter.Workbook("report/img_three.xlsx")
worksheet_three = workbook_three.add_worksheet()
img_three_list = []

workbook_plaque_img_name = xlsxwriter.Workbook("report/plaque_img_name.xlsx")
worksheet_plaque_img_name = workbook_plaque_img_name.add_worksheet()
plaque_img_name = []


for i in range(6000):

    path = "images_one/image_{}.png".format(i)
    if os.path.isfile(path) and os.access(path, os.R_OK):
        img_one_list.append(path)
    else:
        print("Dosyamız mevcut değil yada okunabilir durumda değildir.")

for i in range(6000):

    path = "images_two/image_{}.png".format(i)
    if os.path.isfile(path) and os.access(path, os.R_OK):
        img_two_list.append(path)
    else:
        print("Dosyamız mevcut değil yada okunabilir durumda değildir.")

for i in range(6000):

    path = "images_three/image_{}.png".format(i)
    if os.path.isfile(path) and os.access(path, os.R_OK):
        img_three_list.append(path)
    else:
        print("Dosyamız mevcut değil yada okunabilir durumda değildir.")

for i in range(6000):

    path_one = "plaque/plaque_one_images_{}.png".format(i)
    if os.path.isfile(path_one) and os.access(path_one, os.R_OK):
        plaque_img_name.append(path_one)
    else:
        print("Dosyamız mevcut değil yada okunabilir durumda değildir.")

    path_two = "plaque/plaque_two_images_{}.png".format(i)
    if os.path.isfile(path_two) and os.access(path_two, os.R_OK):
        plaque_img_name.append(path_two)
    else:
        print("Dosyamız mevcut değil yada okunabilir durumda değildir.")

    path_three = "plaque/plaque_three_images_{}.png".format(i)
    if os.path.isfile(path_three) and os.access(path_three, os.R_OK):
        plaque_img_name.append(path_three)
    else:
        print("Dosyamız mevcut değil yada okunabilir durumda değildir.")



row = 0
column = 0
for item in img_one_list:
    worksheet_one.write(row, column, item)
    row += 1
    print("Dosya ismi=" + item)
workbook_one.close()
row = 0
column = 0
for item in img_two_list:
    worksheet_two.write(row, column, item)
    row += 1
    print("Dosya ismi=" + item)
workbook_two.close()
row = 0
column = 0
for item in img_three_list:
    worksheet_three.write(row, column, item)
    row += 1
    print("Dosya ismi=" + item)
workbook_three.close()
row = 0
column = 0
for item in plaque_img_name:
    worksheet_plaque_img_name.write(row, column, item)
    row += 1
    print("Dosya ismi=" + item)
workbook_plaque_img_name.close()
