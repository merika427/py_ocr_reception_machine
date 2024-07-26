import cv2
import pyocr
from PIL import Image
import pyocr.builders
import tkinter as tk
import tkinter.messagebox as messagebox
import time

from playsound import playsound

# ggoleスプレッドシート 対応
from google.oauth2.service_account import Credentials
import gspread
import queue

# 認証鍵
key_file = "XXXXX.json"

pyocr.tesseract.TESSERACT_CMD = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

tools = pyocr.get_available_tools()
print(tools)
tool = tools[0]

cap = cv2.VideoCapture(0)                                                                      
time.sleep(2)                                                                    

#希望のセッティングにしてみる
cap.set(cv2.CAP_PROP_FPS,60); 
cap.set(cv2.CAP_PROP_FRAME_WIDTH,1280);
cap.set(cv2.CAP_PROP_FRAME_HEIGHT,720);


# スプレッドシートの設定
scopes = [
'https://www.googleapis.com/auth/spreadsheets',
'https://www.googleapis.com/auth/drive'
]

credentials = Credentials.from_service_account_file(
key_file,
scopes=scopes
)
gc = gspread.authorize(credentials)

# 対象のスプレッドシート
spreadsheet_url = "https://docs.google.com/spreadsheets/d/XXXXXX"
spreadsheet = gc.open_by_url(spreadsheet_url)



# 読み取り領域（枠）を定義
# これらの値を必要に応じて調整してください
x, y, w, h = 250, 150, 800, 400

def OcrReception():
    while True:
        ret, frame = cap.read()
        Height, Width = frame.shape[:2]


         # 読み取り領域を示す枠を描画
        cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)

        # 読み取り領域を取得
        roi = frame[y:y + h, x:x + w]
        img = cv2.resize(frame, (int(Width), int(Height)))

        ocr(img, Width, Height)

        cv2.imshow('Ocr Test', img)

        if cv2.waitKey(100) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()

def ocr(img, Width, Height):
    
    dst = img[100:Height-200, 100:Width-200]  # OCRで読みたい領域を切り出す
    PIL_Image = Image.fromarray(dst)
    text = tool.image_to_string(
        PIL_Image,
        lang='jpn',
        builder=pyocr.builders.TextBuilder()
    )

    if text != "":
        list = text.split()

        if list:
            googleSheet(list)


def googleSheet(text):
    worksheet = spreadsheet.sheet1
    all_vals = spreadsheet.sheet1.get_all_values()
    
    terget_row = 8
    match_row = []
    cal = 1
    for cols in all_vals:
        row = 0
        for val in cols:
            for txt in text:
                if(val == txt):
                    print(txt)
                    match_row = [cal,row]
                    break


            row += 1
            
        cal += 1
    
    #文字検索
    if(match_row):
        print(match_row)
        worksheet.update_cell(match_row[0], terget_row, '出席')
        worksheet.update_cell(match_row[0], 1, '読み取り')
        showMessage("受付完了","info",3000)

        

def showMessage(message, type='info', timeout=2500):
    import tkinter as tk
    from tkinter import messagebox as msgb

    root = tk.Tk()
    root.withdraw()
    try:
        if type == 'info':

            playsound("SAMPLE.mp3")

            root.after(timeout, root.destroy)
            msgb.showinfo('Info', message, master=root)
        elif type == 'warning':
            msgb.showwarning('Warning', message, master=root)
        elif type == 'error':
            msgb.showerror('Error', message, master=root)
    except:
        pass




OcrReception()
