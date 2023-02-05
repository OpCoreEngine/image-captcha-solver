import time
import win32gui
import win32ui
from ctypes import windll
from PIL import Image

def get_screenshot(hwnd, request_type):
    time.sleep(0.5)
    left, top, right, bot = win32gui.GetWindowRect(hwnd)
    w, h = right - left, bot - top

    hwnd_dc = win32gui.GetWindowDC(hwnd)
    mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
    save_dc = mfc_dc.CreateCompatibleDC()

    save_bit_map = win32ui.CreateBitmap()
    save_bit_map.CreateCompatibleBitmap(mfc_dc, w, h)
    save_dc.SelectObject(save_bit_map)

    windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 2)
    bmp_info = save_bit_map.GetInfo()
    bmp_str = save_bit_map.GetBitmapBits(True)
    im = Image.frombuffer('RGB', (bmp_info['bmWidth'], bmp_info['bmHeight']), bmp_str, 'raw', 'BGRX', 0, 1)

    win32gui.DeleteObject(save_bit_map.GetHandle())
    save_dc.DeleteDC()
    mfc_dc.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwnd_dc)

    item_image = im.crop((310, 259, 510, 285)).convert('L')
    captcha_image = im.crop((260, 370, 550, 420))

    return captcha_image if request_type == "get_captcha" else item_image