import serial
import xlwt
from xlwt import Workbook

wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')

arduinoSerialData = serial.Serial('com3', 115200)

rowCnt = 0

while (1 == 1):
    if (arduinoSerialData.inWaiting()>0):
        myData = arduinoSerialData.readline()
        myData = str(myData, 'utf-8')
        splitData = myData.split(' ')
        x = float(splitData[1])
        y = float(splitData[2])
        z = float(splitData[3])
        gx = float(splitData[4])
        gy = float(splitData[5])
        gz = float(splitData[6])
        sheet1.write(rowCnt, 0, x)
        sheet1.write(rowCnt, 1, y)
        sheet1.write(rowCnt, 2, z)
        sheet1.write(rowCnt, 3, gx)
        sheet1.write(rowCnt, 4, gy)
        sheet1.write(rowCnt, 5, gz)
        rowCnt = rowCnt+1
        print ('x = '+str(x)+' y = '+str(y)+' z = '+str(z)+' gx = '+str(gx)+' gy = '+str(gy)+' gz = '+str(gz))
        wb.save('Data value.xls')
        #print(splitData)