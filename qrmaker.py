import os
from turtle import width
import qrcode
import docx
from docx.shared import Inches
from pickle import TRUE

saveQR = TRUE # Set FALSE if you don't want to save qr images

doc = docx.Document()
list = ["AEA 8", "AEB 4", "AEC 3", "AED 5", "AEE 8",
 "AEF 8", "AEG 9", "AEH 8", "AEI 8", "AEJ 5", "AEK 5",
  "AEL 5", "AEM 0", "AEN 5", "AEO 5", "AEP 5", "AEQ 5",
   "AER 5", "AES 5", "AET 5", "AEU 5"]
# totNum = 8

def printQR():
    for loc in list:
        loc = loc.split(" ")
        totNum = loc[1]
        for num in range (int(totNum)):

            run = doc.add_paragraph().add_run()
            run.font.size = docx.shared.Pt(33)
            run.bold = True
            run.add_picture("./" + str(loc[0]) + "/" + str(loc[0]) + str(num) + ".png", width=Inches(0.6), height=Inches(0.6))
            run.add_text("  " + str(loc[0]) + str(num))

            doc.save('./qr.docx')

def qrmake():
    for loc in list:
        loc = loc.split(" ")
        totNum = loc[1]
        for num in range (int(totNum) + 1):
            qr = qrcode.QRCode(
                border = 0
            )
            qr.add_data("LK"+ str(loc[0]) + str(num))
            qr.make(fit=True)
            img = qr.make_image()
            if (os.path.exists("./" + str(loc[0]))) and saveQR:
                img.save("./" + str(loc[0]) + "/" + str(loc[0]) + str(num) + ".png")
                print("saved QR: " + str(loc[0]) + str(num) + ".png")
            elif saveQR:
                os.makedirs("./" + str(loc[0]))
                img.save("./" + str(loc[0]) + "/" + str(loc[0]) + str(num) + ".png")
                print("saved QR: " + str(loc[0]) + str(num) + ".png")

qrmake()
printQR()