import cv2, pytesseract
import xlwt, os

print('\n..which camera wanna use ?\n')
print('1). IP WebCam')
print('2). Laptop')
camera = input('\nWhich Camera wanna use : ')

if camera == '1':
    url = 'http://192.168.43.1:8080/video'
else:
    url = 0

photo = r'C:\Users\Vicky Kumar\Pictures\Screenshots\excel_ocr.jpg'
video = cv2.VideoCapture(url)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'

while True:
    check, frame = video.read()
    cv2.imshow('Press SPACE to click...', frame)
    key = cv2.waitKey(1)
    if key == ord(' '):
        break

video.release()
cv2.destroyAllWindows()
cv2.imwrite(photo, frame)

image = cv2.imread(photo, cv2.IMREAD_GRAYSCALE)
text = pytesseract.image_to_string(image)

wb = xlwt.Workbook() 
path = r'C:\Users\Vicky Kumar\Pictures\Screenshots\excel_ocr.csv'
sheet1 = wb.add_sheet('ocr') 

lst = text.split('\n')
for r, row in enumerate(lst):
    sheet1.write(int(r), 0, row)
        
wb.save(path) 
os.startfile(path)
