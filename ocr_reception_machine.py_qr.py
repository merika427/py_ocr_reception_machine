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

if not cap.isOpened():
    print("カメラが検出されませんでした。接続を確認してください。")
    return

print("QRコードをカメラにかざしてください。 'q' を押して終了します。")

# QRコードのデコーダーを準備
qr_detector = cv2.QRCodeDetector()

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




def OcrReception():
    # カメラを起動
    cap = cv2.VideoCapture(0)
    if not cap.isOpened():
        print("カメラが検出されませんでした。接続を確認してください。")
        return

    print("QRコードをカメラにかざしてください。 'q' を押して終了します。")

    # QRコードのデコーダーを準備
    qr_detector = cv2.QRCodeDetector()

    while True:
        ret, frame = cap.read()
        if not ret:
            print("カメラからの映像が取得できませんでした。")
            break

        # QRコードをデコード
        data, bbox, _ = qr_detector.detectAndDecode(frame)
        if data:
            print(f"読み取ったデータ: {data}")

            # URLの場合、デフォルトのブラウザで開く
            if data.startswith("http://") or data.startswith("https://"):
                print(f"URLを開いています: {data}")

                object_url = url = data
                
                match = re.search(r'/([^/]+)/$', url)
                if match:
                    result = match.group(1)
                    
                    # スプレッドシートと照合
                    googleSheet([result])


                    # スプレッドシートと照合後、そのまま読み取ったURLでブラウザを開く場合
                    # if(googleSheet([result])):
                    #     webbrowser.open(data)  
                        
                
                
            print("再度QRコードを読み取る準備をしています...")
            continue  # 再度カメラを継続して起動

        # 映像を表示
        cv2.imshow("QRコードリーダー", frame)

        # 'q'を押すと終了
        if cv2.waitKey(1) & 0xFF == ord('q'):
            print("終了します。")
            break

    cap.release()
    cv2.destroyAllWindows()



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