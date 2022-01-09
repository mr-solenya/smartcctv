import sys
import numpy as np
import cv2
import xlsxwriter
import time
from datetime import datetime

#initialise the xlsx book and sheet
book= xlsxwriter.Workbook(r'.\\report1.xlsx')  
sheet=book.add_worksheet()
row=0
col=0
sheet.write(row,col,"INSTRUSION DATA LOG")
row+=1
#path=input("Enter the path of the database: ")
#path=".\\haarcascade_frontalface_default.xml"
path=".\\haarcascade_upperbody.xml"
#path=".\\haarcascade_profileface.xml"
face_cascade = cv2.CascadeClassifier(path) #training the ml classifier

cap=cv2.VideoCapture(0) #initialising camera0
prevtime=datetime.now()

while(True):
    ret,img=cap.read() #reading each frame into a variable called 'img'
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY) #converting into a grayscale frame and storing into variable 'gray'
    faces = face_cascade.detectMultiScale(gray, 1.3, 5) #using cascaded classifiers to detect faces in 'gray'
    for (x,y,w,h) in faces:
        if x<600 and y<400 : #defining the ROI
            print('INTRUSION DETECTED')
            now=datetime.now()
            currtime=now.strftime("%H:%M")
            if((now-prevtime).total_seconds()>10):
                sheet.write(row,col,currtime)
                sheet.write(row,col+1,'INSTRUSION')
                row+=1
                prevtime=now
        else:
            print('NO INTRUSION DETECTED')
        cv2.rectangle(img,(x,y),(x+w,y+h),(255,0,0),2) #rectangle(image, start_point, end_point, color, thickness)
        #print(x,y)
    cv2.imshow('test window',img)
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break
    #time.sleep(10)
cap.release()
book.close()
cv2.destroyAllWindows()
