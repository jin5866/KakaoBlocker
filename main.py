# 출처 https://airfox1.tistory.com/2
# -*- coding: cp949 -*-
import time, win32con, win32api, win32gui

# # 카톡창 이름, (활성화 상태의 열려있는 창)
kakao_opentalk_name = 'test'


def SendReturnNoRelease(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)

# # 엔터
def SendReturn(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    time.sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

def SendRclick(hwnd):
    tmppose = (34 << 16) + 145
    a = win32gui.GetWindowPlacement(hwnd)
    pose = ((a[0]+34) << 16) + a[1]+145
    win32api.PostMessage(hwnd, win32con.WM_RBUTTONDOWN, win32con.MK_RBUTTON ,pose)
    time.sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_RBUTTONUP,win32con.MK_RBUTTON, pose)
    #win32con.WM_RBUTTONUP

def ClickYes(hwnd):
    #162,79
    a = win32gui.GetWindowPlacement(hwnd)
    pose = ((a[0]+162) << 16) + a[1]+79
    win32api.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, pose)
    time.sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, pose)

def SendDownKey(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
    time.sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_DOWN, 0)

def check_chatroom(chatroom_name):
    desktop = win32gui.GetDesktopWindow()
    hwndkakao = win32gui.FindWindow(None, "카카오톡")
    hwndkakao_edit1 = win32gui.FindWindowEx( hwndkakao, None, "EVA_ChildWindow", None)
    hwndkakao_edit2_1 = win32gui.FindWindowEx( hwndkakao_edit1, None, "EVA_Window", None)
    #hwndkakao_edit2_2 = win32gui.FindWindowEx( hwndkakao_edit1, hwndkakao_edit2_1, "EVA_Window", None)
    hwndkakao_edit3 = win32gui.FindWindowEx( hwndkakao_edit2_1, None, "Edit", None)
    hwndkakao_search = win32gui.FindWindowEx( hwndkakao_edit2_1, None, "EVA_VH_ListControl_Dblclk", None)
    hwndkakao_search2 = win32gui.FindWindowEx( hwndkakao_search, None, "_EVA_CustomScrollCtrl", None)

    win32con.GW_OWNER

    # 방 위치
    hwndkakao_contact = win32gui.FindWindowEx( hwndkakao_edit2_1, hwndkakao_search, "EVA_VH_ListControl_Dblclk", None)

    target = hwndkakao_contact
    print(target)

    d = win32gui.GetWindow(hwndkakao_edit2_1,win32con.GW_OWNER )
    print(d)

    # # Edit에 검색 _ 입력되어있는 텍스트가 있어도 덮어쓰기됨
    win32api.SendMessage(hwndkakao_edit3, win32con.WM_SETTEXT, 0, chatroom_name)
    time.sleep(1)   # 안정성 위해 필요
    SendRclick(target)
    #SendReturn(hwndkakao_edit3)
    time.sleep(1)
    # 차단까지
    for i in range(0,7):
        SendDownKey(target)
        time.sleep(0.1)
    SendReturn(target)
    time.sleep(0.2)
    # 확인버튼
    tg = win32gui.FindWindow("EVA_Window_Dblclk",None)
    while(win32gui.GetWindow(tg,win32con.GW_OWNER )!=hwndkakao):
        tg = win32gui.FindWindowEx(desktop,tg,"EVA_Window_Dblclk",None)
    print(tg)
    # 확인
    SendReturnNoRelease(tg)
    #ClickYes(KakaoTalkShadowWnd)

def main():
    check_chatroom(kakao_opentalk_name)

if __name__ == '__main__':
    main()
