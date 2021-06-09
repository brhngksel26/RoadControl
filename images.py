import cv2
import xlrd

workbook_one = xlrd.open_workbook("report/img_one.xlsx")
sheet_one = workbook_one.sheet_by_index(0)

workbook_two = xlrd.open_workbook("report/img_two.xlsx")
sheet_two = workbook_two.sheet_by_index(0)

workbook_three = xlrd.open_workbook("report/img_three.xlsx")
sheet_three = workbook_three.sheet_by_index(0)

row = 0
column = 0

for row in range(127):
    img_name = sheet_one.cell_value(row,column)
    image = cv2.imread(img_name)
    cropped = image[300:500, 150:500]
    plaque_images = "plaque/plaque_one_images_{}.png".format(row)
    print(plaque_images)
    cv2.imwrite(plaque_images, cropped)

for row in range(2054):
    img_name = sheet_two.cell_value(row,column)
    image = cv2.imread(img_name)
    cropped = image[300:500, 150:500]
    plaque_images = "plaque/plaque_two_images_{}.png".format(row)
    print(plaque_images)
    cv2.imwrite(plaque_images, cropped)

for row in range(3817):
    img_name = sheet_three.cell_value(row,column)
    image = cv2.imread(img_name)
    cropped = image[300:500, 150:500]
    plaque_images = "plaque/plaque_three_images_{}.png".format(row)
    print(plaque_images)
    cv2.imwrite(plaque_images, cropped)






# write the cropped image to disk in PNG format
