import cv2
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

backround = cv2.createBackgroundSubtractorMOG2()
roads = cv2.VideoCapture('videos/roadsvideo.mkv')

total_car = 0
minArea = 2600

while True:
    row = 0
    column = 0
    _, frame = roads.read()
    fgmask = backround.apply(frame, None, 0.1)
    erode = cv2.erode(fgmask, None, iterations=4)
    moments = cv2.moments(erode, True)
    #gelen yol çizgi
    # x ekseni
    cv2.line(frame, (0, 300), (10000, 300), (0, 255, 0), 2)
    cv2.line(frame, (0, 500), (10000, 500), (0, 255, 0), 2)
    #y ekseni
    cv2.line(frame, (150, 0), (150, 10000), (255, 255, 0), 2)
    cv2.line(frame, (500, 0), (500, 10000), (255, 255, 0), 2)
    cv2.line(frame, (770, 0), (770, 10000), (255, 0, 0), 2)
    cv2.line(frame, (1150, 0), (1150, 10000), (255, 0, 0), 2)
    cv2.line(frame, (1300, 0), (1300, 10000), (0, 0, 0), 2)
    cv2.line(frame, (1700, 0), (1700, 10000), (0, 0, 0), 2)


    if moments['m00'] >= minArea:
        x = int(moments['m10']/moments['m00'])
        y = int(moments['m01'] / moments['m00'])

        #print('moment :' + str(moments['m00'])+" x : " + str(x) + " y :" + str(y))
        if (y > 150 and y < 500):
            total_car = total_car + 1
            img_name = "images_one/image_{}.png".format(total_car)
            time = str(datetime.datetime.now())
            worksheet_one.write(row, column, img_name)
            worksheet_one.write(row,column+1,time)
            row += 1
            column += 1
            cv2.imwrite(img_name, frame)
            print(img_name)
            print(time)

        elif ( y > 770 and y < 1150):
            total_car = total_car + 1
            img_name = "images_two/image_{}.png".format(total_car)
            time = str(datetime.datetime.now())
            worksheet_two.write(row, column, img_name)
            worksheet_two.write(row, column + 1, time)
            row += 1
            column += 1
            cv2.imwrite(img_name, frame)
            print('toplam = ' + str(total_car))
            print(img_name)
            print(time)




        elif (y > 1300 and y < 1700):
            total_car = total_car + 1
            img_name = "images_three/image_{}.png".format(total_car)
            time = str(datetime.datetime.now())
            worksheet_two.write(row, column, img_name)
            worksheet_two.write(row, column + 1, time)
            row += 1
            column += 1
            cv2.imwrite(img_name, frame)
            print('toplam = ' + str(total_car))
            print(img_name)
            print(time)

    workbook_one.close()
    workbook_two.close()
    workbook_three.close()
    cv2.putText(frame, 'Sayi: %r' % total_car, (200, 30), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)
    cv2.imshow('video',frame)
    key = cv2.waitKey(25)
    if key==ord('q'):
        break

roads.release()

cv2.destroyAllWindows()



