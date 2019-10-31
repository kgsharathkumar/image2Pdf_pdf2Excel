import PIL.ImageGrab
import datetime
import os.path
import time
import os
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import pyautogui
from fpdf import FPDF
from PIL import Image
import img2pdf
from array import array
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from Cython.Compiler.Errors import listing_file
import selenium.webdriver.common.keys
from selenium.webdriver.common.by import By
#excel
#import jpype
#from WorkingWithDocumentConversion import PdfToExcel


#root = "C:\\Users\\sgangadh.BLR\\Desktop\\SandBox\\py_image_stub\\" 
save_path = 'C:/Users/sgangadh.BLR/Desktop/SandBox/py_image_stub/'
d1=datetime.datetime.now().strftime("%Y%M%d %H%M%S")
print("#########################################################################")
print("########### Welcome to PathTech Automation Tool for QAC #####################")
print("########### Developed By PathTech Bengaluru #################################")
print("########### Thanks for using this Tool ##################################")
print("########### NOTE: This Tool is used for only Internal Purpose ###########")
print("######## For Improvement Contact: mailtokgsharath@gmail.com  ##############")
print("#########################################################################")

print(d1)
completeName = os.path.join(save_path, d1+'_Test_py'+".png")


#delay in 3 second
time.sleep(5)

img= PIL.ImageGrab.grab()
#img.save(str(str(d1)+'.png'),'png')
img.save(completeName)

last_height= 12

for i in range(last_height):
    time.sleep(1)
    d1=datetime.datetime.now().strftime("%Y%M%d %H%M%S")
    print(d1)
    completeName = os.path.join(save_path, d1+'_Test_py'+".png")
    img= PIL.ImageGrab.grab()
#img.save(str(str(d1)+'.png'),'png')
    img.save(completeName)
    #### Page Scroll Down
    pyautogui.press('pagedown')

'''   
#### Converting to PDF### WAY2
pdf = FPDF()
# imagelist is the list with all image filenames
dirpath = "C:\\Users\\sgangadh.BLR\\Desktop\\SandBox\\py_image_stub\\"
imagelist = os.path.basename(dirpath) + ".png"
for image in imagelist:
    pdf.add_page()
    #pdf.image(image,x,y,w,h)
    pdf.image(image,100,100,100,100)
pdf.output("C:\\Users\\sgangadh.BLR\\Desktop\\SandBox\\py_image_stub\\yourfile.pdf", "F")

'''

'''        
#### Converting to PDF ### WAY1
   
pdf_bytes = img2pdf.convert(['C:/Users/sgangadh.BLR/Desktop/SandBox/py_image_stub/20171510 161559_Test_py.png'])
print("Pdf started")
file = open("C:/Users/sgangadh.BLR/Desktop/SandBox/py_image_stub/name.pdf","wb")
file.write(pdf_bytes)   
print("Pdf end")
'''
########### PDF Converter ######## Best Way 
try:
    n = 0
    for dirpath, dirnames, filenames in os.walk(save_path):
        PdfOutputFileName = os.path.basename(dirpath) + save_path +"MyPDF.pdf" 
        c = canvas.Canvas(PdfOutputFileName)
        if n >= 0 :
            for filename in filenames:
                    LowerCaseFileName = filename.lower()
                    if LowerCaseFileName.endswith(".png"):
                        print(filename)
                        filepath    = os.path.join(dirpath, filename)
                        print(filepath)
                        im          = ImageReader(filepath)
                        imagesize   = im.getSize()
                        c.setPageSize(imagesize)
                        c.drawImage(filepath,0,0)
                        c.showPage()
        n = n + 1
        c.save()
    #print "PDF of Image directory created" + filepath.PdfOutputFileName
    print("PDF of Image directory created")
except:
    print("Failed creating PDF")
finally:
    print("PDF Conversion Successfully Done by EASi") 

################################## END OF PDF Converter Method #########################
    
########### Excel Converter ######## 
'''
doc=self.Document()
pdf = self.Document()
pdf=self.dataDir + save_path +"MyPDF.pdf"

# Instantiate ExcelSave Option object
excelsave=self.ExcelSaveOptions();
 
# Save the output to XLS format
doc.save(self.dataDir + save_path +"Converted_Excel.xls", excelsave);
 
print "Document has been converted successfully"

'''   
    
