import win32gui , win32con ,win32api ,win32ui

from ctypes import windll
from PIL import Image, ImageQt
import numpy as np
import cv2
import time 
import threading
import os

#---------------------------------------------------------------------------            
  
#遊戲視窗起始位置
def getGame_X_Y():
    hwnd = win32gui.FindWindow(None,'MapleStory')
    # hwnd =await get_jb_id()
    x1,y1,x2,y2 = win32gui.GetWindowRect(hwnd)

    #x軸+3  軸+26 去除標題的位移
    return x1+3,y1+26,x2,y2

#---------------------------------------------------------------------------            
# 滑鼠左鍵點擊
def mouseclick(x,y):
    win32api.SetCursorPos((x+( getGame_X_Y())[0], y+( getGame_X_Y())[1]))
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0) 
#---------------------------------------------------------------------------
#擷取楓谷畫面 
def get_game_pic():
    #擷取楓谷畫面 
    try:
        hwnd = win32gui.FindWindow(None,'MapleStory')
        # hwnd = get_jb_id()
        # 如果使用高 DPI 显示器（或 > 100% 缩放尺寸），添加下面一行，否则注释掉
        windll.user32.SetProcessDPIAware()

        # Change the line below depending on whether you want the whole window
        # or just the client area.
        # 根据您是想要整个窗口还是只需要 client area 来更改下面的行。
        left, top, right, bot = win32gui.GetClientRect(hwnd)
        # left, top, right, bot = win32gui.GetWindowRect(hwnd)
        w = right - left
        h = bot - top  

        hwndDC = win32gui.GetWindowDC(hwnd)  # 根据窗口句柄获取窗口的设备上下文DC（Divice Context）
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)  # 根据窗口的DC获取mfcDC
        saveDC = mfcDC.CreateCompatibleDC()  # mfcDC创建可兼容的DC

        saveBitMap = win32ui.CreateBitmap()  # 创建bitmap准备保存图片
        saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)  # 为bitmap开辟空间

        saveDC.SelectObject(saveBitMap)  # 高度saveDC，将截图保存到saveBitmap中

        # 选择合适的 window number，如0，1，2，3，直到截图从黑色变为正常画面
        result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 3)

        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)

        im = Image.frombuffer(
            'RGB',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRX', 0, 1)

        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwndDC)

        if result == 1:
            # PrintWindow Succeeded
            qimage = ImageQt.toqimage(im)
            
            return qimage  # 返回图片
        else:
            print("獲取遊戲圖片失敗")
    except Exception as e:  
        print (e)  
        print('抓取遊戲畫面失敗')    
#---------------------------------------------------------------------------            
#img檔 轉cv2 圖檔 為了使用get_1()
def qimgtocv2(qimg=get_game_pic()):
    try:
        # qimg=(maple_function.get_game_pic(hwnd))
        temp_shape = (qimg.height(), qimg.bytesPerLine() * 8 // qimg.depth())
        temp_shape += (4,)
        ptr = qimg.bits()
        ptr.setsize(qimg.byteCount())
        result = np.array(ptr, dtype=np.uint8).reshape(temp_shape)

        return result
    except Exception as e:  
        print (e)  
        print('抓取遊戲畫面失敗')  

#---------------------------------------------------------------------------            
#比較圖片(遊戲,尋找的圖)
def get_1(pic,pic2):
    

    pic1 = cv2.cvtColor(pic, cv2.COLOR_BGR2GRAY)

    pic2 = cv2.cvtColor(pic2, cv2.COLOR_BGR2GRAY)



    res = cv2.matchTemplate(pic2,pic1,cv2.TM_CCOEFF_NORMED)

    threshold = 0.96

    loc = np.where(res >= threshold)
    x=loc[1]
    y=loc[0]

    if len(x) and len(y):

        return True
    else :
        return False         

#---------------------------------------------------------------------------            
#比較圖片(遊戲,尋找的圖)
def get_2(pic,pic2):
    

    pic1 = cv2.cvtColor(pic, cv2.COLOR_BGR2GRAY)

    pic2 = cv2.cvtColor(pic2, cv2.COLOR_BGR2GRAY)



    res = cv2.matchTemplate(pic2,pic1,cv2.TM_CCOEFF_NORMED)

    threshold = 0.96

    loc = np.where(res >= threshold)
    x=loc[1]
    y=loc[0]

    if len(x) and len(y):

        return [x,y]
    else :
        return False  


pic1 = cv2.imread(os.path.dirname(os.path.abspath("minigame.py"))+"/result/1.png")
pic2 = cv2.imread(os.path.dirname(os.path.abspath("minigame.py"))+"/result/2.png")
pic3 = cv2.imread(os.path.dirname(os.path.abspath("minigame.py"))+"/result/3.png")
pic4 = cv2.imread(os.path.dirname(os.path.abspath("minigame.py"))+"/result/4.png")


def main():
    while True:
        print('等待抓取目標,請務必將楓谷視窗置頂')
        gamepic = qimgtocv2(get_game_pic())

        time.sleep(0.1)
        if  get_1(pic1,gamepic) or get_1(pic2,gamepic) or get_1(pic3,gamepic) or get_1(pic4,gamepic):
            print('遊戲開始,找到目標')

            try:
                aa = get_2(pic1,gamepic)
                time.sleep(0.1)
                if aa != False:
                    alist_x = list(aa[0])
                    alist_y = list(aa[1])
                else:
                    alist_x = []
                    alist_y = []
            except:
                print('沒有1號猴子')

            try:
                bb = get_2(pic2,gamepic)
                if bb != False:
                    blist_x = list(bb[0])
                    blist_y = list(bb[1])
                else:
                    blist_x = []
                    blist_y = []

            except:
                print('沒有2號猴子')

            try:
                cc = get_2(pic3,gamepic)
                if cc != False:
                    clist_x = list(cc[0])
                    clist_y = list(cc[1])
                else:
                    clist_x = []
                    clist_y = []                
            except:
                print('沒有3號猴子')

            try:
                dd = get_2(pic4,gamepic)
                if dd != False:
                    dlist_x = list(dd[0])
                    dlist_y = list(dd[1])
                else:
                    dlist_x = []
                    dlist_y = []                   
            except:
                print('沒有4號猴子')               

            alllist_x = alist_x + blist_x + clist_x + dlist_x
            alllist_y = alist_y + blist_y + clist_y + dlist_y
            print('共有 ' +str(len(alllist_x)) + ' 隻猴子')
            print('猴子X座標')
            print(alllist_x)
            print('猴子Y座標')
            print(alllist_y)


            for i in range(0,len(alllist_x)):
                mouseclick(alllist_x[i]+10,alllist_y[i]+10)
                time.sleep(0.4)
            
            mouseclick(1220,470)
            time.sleep(0.5)

        time.sleep(0.2)        
            







main1 = threading.Thread(target=main)

main1.start()




