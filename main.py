import ctypes
import tkinter as tk
import sys
from ctypes import windll
import jellyfish
import keyboard
import pyautogui
from python_imagesearch.imagesearch import imagesearch  # import the imagesearch module
import time
import win32gui
from screenshot import get_screenshot
import pytesseract as pt
from ItemDB import items
from threading import Thread

pt.pytesseract.tesseract_cmd = './Tesseract-OCR/tesseract'

def detect_captcha(client):
    time.sleep(1)
    print("Captcha tespit ediliyor")
    captcha_args = ['Bot', 'olmadığını', 'doğrulamak', 'için', 'lütfen', 'seçin', 'Copyright']
    bot_control_ss = get_screenshot(hwnd=client, request_type="get_captcha")
    bot_control_text = pt.image_to_string(bot_control_ss, lang='tur')
    for arg in captcha_args:
        bot_control_text = ''.join(bot_control_text)
        if arg in bot_control_text:
            return True
    return False


def analyse_captcha(client):
    compare_captcha = {}
    captcha_ss = get_screenshot(hwnd=client, request_type="get_item_name")
    item_text = pt.image_to_string(captcha_ss, lang='tur')
    modified_text = item_text.replace('4', 'a')
    item_text_list = ''.join(modified_text)
    for item in items:
        similarity = jellyfish.jaro_winkler_similarity(item_text_list.upper(), item)
        if similarity > 0.50:
            compare_captcha[item] = similarity
    if not compare_captcha:
        return False
    possible_item = max(compare_captcha, key=compare_captcha.get)
    print(f"{possible_item} istenilen eşya olabilir")
    return possible_item


def solve_captcha(possible_item):
    pos = imagesearch(f"captha_images/{possible_item}.png")
    if pos[0] == -1:
        return False
    captcha_pos_x = pos[0] + 3
    captcha_pos_y = pos[1] + 14
    dd_dll.DD_mov(captcha_pos_x, captcha_pos_y)
    time.sleep(0.05)
    dd_dll.DD_mov(captcha_pos_x + 1, captcha_pos_y + 1)
    time.sleep(0.05)
    dd_dll.DD_btn(1)
    dd_dll.DD_btn(2)


def window_title_update():
    windows = pyautogui.getAllWindows()
    for window in windows:
        if len(window.title) > 0:
            window_list.append(window.title)



def load_driver():
    global dd_dll
    driver_name = driver_option.get()
    if driver_name == 'Simple':
        dd_dll = windll.LoadLibrary('./drivers/DD94687.64.dll')
    elif driver_name == 'General':
        dd_dll = windll.LoadLibrary('./drivers/DD64.dll')
    elif driver_name == 'HID':
        dd_dll = windll.LoadLibrary('./drivers/DDHID64.dll')
    else:
        ctypes.windll.user32.MessageBoxW(0, "Driver Seçmeyi Unuttun", "HATA", 0)
        sys.exit()



client_amount_list = ["1", "2", "3", "4", "5", "6"]
driver_list = ["Simple", "General", "HID"]
client_names = []
window_list = []


stop_flag = 0
def scanning():
    ape_value = 0
    while True:
        for val in client_names:
            try:
                client = win32gui.FindWindow(None, str(val.get()))
                left, top, right, bot = win32gui.GetWindowRect(client)
                if detect_captcha(client):
                    print("Captcha detected")
                    time.sleep(0.2)
                    captcha_item_name = analyse_captcha(client)
                    if not captcha_item_name:
                        handle_no_captcha_item_name(ape_value, left, top)
                    else:
                        coz_bunu_esteban = solve_captcha(captcha_item_name)
                        if not coz_bunu_esteban:
                            handle_no_captcha_item_name(ape_value, left, top)
                else:
                    time.sleep(0.3)
            except Exception as e:
                client_names.remove(val)
                print(e)
        if stop_flag:
            break

def handle_no_captcha_item_name(ape_value, left, top):
    print("Captcha is not valid")
    time.sleep(0.1)
    ape_value += 1
    if ape_value > 5:
        print("Stuck in captcha")
        ape_value = 0
        time.sleep(0.1)
        captcha_pos_x = left + 240
        captcha_pos_y = top + 310
        dd_dll.DD_mov(captcha_pos_x, captcha_pos_y)
        time.sleep(0.05)
        dd_dll.DD_mov(captcha_pos_x + 1, captcha_pos_y + 1)
        dd_dll.DD_btn(1)
        dd_dll.DD_btn(2)
    time.sleep(0.05)

def start_thread():
    global stop_flag
    stop_flag = 0
    t = Thread(target=scanning)
    t.start()
    t2 = Thread(target=new_stop)
    t2.start()
    start_button.config(font=('Helvetica', 12), state=tk.DISABLED)

def new_stop():
    global stop_flag
    while True:
        keyboard.wait("F8")
        stop_flag = 1
        start_button.config(font=('Helvetica', 12), state=tk.NORMAL)



def create_option_menu(root, variable, value, x, y):
    option_menu = tk.OptionMenu(root, variable, *value)
    option_menu.config(bg='#252321', fg='#ffffff', activebackground='#4D4D4E', activeforeground='#ffffff',
                        font=('Helvetica', 11))
    option_menu["menu"].config(bg="#252321", fg="white", activebackground="#4D4D4E", activeforeground="BLACK",
                                font=("Helvetica", 12))
    option_menu.place(relx=0.5, rely=0.5, height=50, anchor=tk.CENTER, width=120, x=x, y=y)
    return option_menu

root = tk.Tk()
root.title("Captcha Solver Beta V3")
root.geometry("500x500")
root.config(bg='#252321')

client_amount = tk.StringVar(root)
client_amount.set("Client Miktari")
create_option_menu(root, client_amount, client_amount_list, 0, -200)

driver_option = tk.StringVar(root)
driver_option.set("Driver Seç")
create_option_menu(root, driver_option, driver_list, -150, -200)

def add_client():
    window_title_update()
    client_dropdown_location_y = -170
    for i in range(0, int(client_amount.get()), 1):
        value_inside = tk.StringVar(root)
        value_inside.set("Client" + str(i + 1))

        client_dropdown = tk.OptionMenu(root, value_inside, *window_list)
        client_dropdown.config(bg='#252321', fg='#ffffff', activebackground='#4D4D4E', activeforeground='#ffffff',
                               font=('Helvetica', 11))
        client_dropdown_location_y += 65
        client_dropdown.place(relx=0.5, rely=0.5, anchor=tk.CENTER, height=30, width=120, x=0,
                              y=client_dropdown_location_y)
        client_names.append(value_inside)
    load_driver()
    root.after(2000, disable_button)

client_button = tk.Button(root, text="ONAYLA", fg="white", bg="#252321", command=add_client)
client_button.place(relx=0.5, rely=0.5, anchor=tk.CENTER, x=0, y=-150)

start_button = tk.Button(root, text="START", fg="white", bg="#252321", command=start_thread)
start_button.config(font=('Helvetica', 12))
start_button.place(relx=0.5, rely=0.5, anchor=tk.CENTER, x=-150, y=0, height=50, width=100)


def disable_button():
    client_button.config(state=tk.DISABLED)
    return None


app = tk.Frame(root)
app.grid()
root.mainloop()
