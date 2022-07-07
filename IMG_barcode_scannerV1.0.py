# -*- coding: utf-8 -*-
"""
Created on Sat Mar  5 15:37:14 2022

@author: Derkk
"""
import os
import cv2 as cv
from pyzbar import pyzbar
import xlsxwriter
from datetime import datetime

DateTime = datetime.now().strftime(" %m-%d-%Y %H.%M.%S")

#create workbook to check
workbooktitle = "img_barcode_extractor" + DateTime + ".xlsx"
workbook = xlsxwriter.Workbook(workbooktitle)
sheet_name = str("barcodes")
worksheet_analysis = workbook.add_worksheet(sheet_name)
worksheet_analysis.write(0,0,"img")
worksheet_analysis.write(0,1,"barcode")

#point to location of images to be scanned
directory = r"C:\Directory"

row_count = 0                      

for file in os.listdir(directory):
    row_count += 1
    
    try:
        filename = os.fsdecode(file)

        image = cv.imread(filename)
        
        #decode image
        barcodes = pyzbar.decode(image)
        
        for barcode in barcodes:
            x,y,w,h = barcode.rect
        
            #draw rectange over the code
            cv.rectangle(image, (x,y), (x+w, y+h), (255,0,0), 4)
        
            #convert into string
            ean = barcode.data.decode("utf-8")
            print(ean)
        worksheet_analysis.write(row_count,0,filename)
        worksheet_analysis.write(row_count,1,str(ean))
        ean = "no barcode"
        print(ean)
    except:
        ean = "no barcode"
        print(ean)
        worksheet_analysis.write(row_count,0,filename)
        worksheet_analysis.write(row_count,1,str(ean))
    

workbook.close()