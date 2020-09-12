import win32api
import win32gui
import win32con
import tkinter as tk
from tkinter import ttk
import ctypes.windll
import ctypes.wintypes
import base64
from PIL import Image, ImageTk
import os
from collections import OrderedDict
import time

# constants
VK_CODE = {
	'backspace': 0x08,
	'tab': 0x09,
	'clear': 0x0C,
	'enter': 0x0D,
	'shift': 0x10,
	'ctrl': 0x11,
	'alt': 0x12,
	'pause': 0x13,
	'caps_lock': 0x14,
	'esc': 0x1B,
	'spacebar': 0x20,
	'page_up': 0x21,
	'page_down': 0x22,
	'end': 0x23,
	'home': 0x24,
	'left_arrow': 0x25,
	'up_arrow': 0x26,
	'right_arrow': 0x27,
	'down_arrow': 0x28,
	'select': 0x29,
	'print': 0x2A,
	'execute': 0x2B,
	'print_screen': 0x2C,
	'ins': 0x2D,
	'del': 0x2E,
	'help': 0x2F,
	'0': 0x30,
	'1': 0x31,
	'2': 0x32,
	'3': 0x33,
	'4': 0x34,
	'5': 0x35,
	'6': 0x36,
	'7': 0x37,
	'8': 0x38,
	'9': 0x39,
	'a': 0x41,
	'b': 0x42,
	'c': 0x43,
	'd': 0x44,
	'e': 0x45,
	'f': 0x46,
	'g': 0x47,
	'h': 0x48,
	'i': 0x49,
	'j': 0x4A,
	'k': 0x4B,
	'l': 0x4C,
	'm': 0x4D,
	'n': 0x4E,
	'o': 0x4F,
	'p': 0x50,
	'q': 0x51,
	'r': 0x52,
	's': 0x53,
	't': 0x54,
	'u': 0x55,
	'v': 0x56,
	'w': 0x57,
	'x': 0x58,
	'y': 0x59,
	'z': 0x5A,
	'numpad_0': 0x60,
	'numpad_1': 0x61,
	'numpad_2': 0x62,
	'numpad_3': 0x63,
	'numpad_4': 0x64,
	'numpad_5': 0x65,
	'numpad_6': 0x66,
	'numpad_7': 0x67,
	'numpad_8': 0x68,
	'numpad_9': 0x69,
	'multiply_key': 0x6A,
	'add_key': 0x6B,
	'separator_key': 0x6C,
	'subtract_key': 0x6D,
	'decimal_key': 0x6E,
	'divide_key': 0x6F,
	'F1': 0x70,
	'F2': 0x71,
	'F3': 0x72,
	'F4': 0x73,
	'F5': 0x74,
	'F6': 0x75,
	'F7': 0x76,
	'F8': 0x77,
	'F9': 0x78,
	'F10': 0x79,
	'F11': 0x7A,
	'F12': 0x7B,
	'F13': 0x7C,
	'F14': 0x7D,
	'F15': 0x7E,
	'F16': 0x7F,
	'F17': 0x80,
	'F18': 0x81,
	'F19': 0x82,
	'F20': 0x83,
	'F21': 0x84,
	'F22': 0x85,
	'F23': 0x86,
	'F24': 0x87,
	'num_lock': 0x90,
	'scroll_lock': 0x91,
	'left_shift': 0xA0,
	'right_shift ': 0xA1,
	'left_control': 0xA2,
	'right_control': 0xA3,
	'left_menu': 0xA4,
	'right_menu': 0xA5,
	'browser_back': 0xA6,
	'browser_forward': 0xA7,
	'browser_refresh': 0xA8,
	'browser_stop': 0xA9,
	'browser_search': 0xAA,
	'browser_favorites': 0xAB,
	'browser_start_and_home': 0xAC,
	'volume_mute': 0xAD,
	'volume_Down': 0xAE,
	'volume_up': 0xAF,
	'next_track': 0xB0,
	'previous_track': 0xB1,
	'stop_media': 0xB2,
	'play/pause_media': 0xB3,
	'start_mail': 0xB4,
	'select_media': 0xB5,
	'start_application_1': 0xB6,
	'start_application_2': 0xB7,
	'attn_key': 0xF6,
	'crsel_key': 0xF7,
	'exsel_key': 0xF8,
	'play_key': 0xFA,
	'zoom_key': 0xFB,
	'clear_key': 0xFE,
	'+': 0xBB,
	',': 0xBC,
	'-': 0xBD,
	'.': 0xBE,
	'/': 0xBF,
	'`': 0xC0,
	';': 0xBA,
	'[': 0xDB,
	'\\': 0xDC,
	']': 0xDD,
	"'": 0xDE,
	'`': 0xC0
}
TIME_AXIS_ERROR_TEXT = '时间轴格式错误'
TIME_AXIS_VALID_TEXT = '成功加载时间轴'

# global variables
list_windows = OrderedDict()
list_script = []
time_axis = OrderedDict()
script_status = False
global listen_window
global script_status_desc

# get window size
def get_current_size(hWnd):
	try:
		f = ctypes.windll.dwmapi.DwmGetWindowAttribute
	except WindowsError:
		f = None
	if f:
		DWMWA_CLOAKED = 14
		cloaked = ctypes.wintypes.DWORD()
		f(hWnd, ctypes.wintypes.DWORD(DWMWA_CLOAKED),ctypes.byref(cloaked),ctypes.sizeof(cloaked))
		return cloaked.value

# get window handle in windows system
def get_all_hwnd(hwnd, extra):
	IsWindowVisible = ctypes.windll.user32.IsWindowVisible
	GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW

	if IsWindowVisible(hwnd) and GetWindowTextLength(hwnd) != 0 and not get_current_size(hwnd):
			windows = extra
			windows[hwnd] = win32gui.GetWindowText(hwnd)

# refresh window when click combo box
def click_window_combo_box(event):
	# get all visible windows
	list_windows.clear()
	win32gui.EnumWindows(get_all_hwnd, list_windows)
	show_list = []
	for (k,v) in list_windows.items():
		show_list.append(v)
	cb1['values'] = show_list

# set jx as window icon
def set_icon(window):
	tmp = open("tmp.ico", "wb+")
	tmp.write(base64.b64decode('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'))
	tmp.close()
	im = Image.open("tmp.ico")
	img = ImageTk.PhotoImage(im)
	window.tk.call('wm', 'iconphoto', window._w, img)
	os.remove("tmp.ico")

# add choose window event
def choose_window_in_combo_box(event):
	index = cb1.current()
	print(index)
	global listen_window
	listen_window = list(list_windows.keys())[index]

# kai shi da zhou
def start_battle(event):
	start = 0
	for t in time_axis:
		time.sleep((float(t) - start)/ 1000 - 0.001)
		key_input(time_axis[t])
		start = int(t)

# key input func
def key_input(input_words = ''):
	win32gui.SetForegroundWindow(listen_window)
	for word in input_words:
		win32api.keybd_event(VK_CODE[word], 0, 0, 0)
		time.sleep(0.001)
		win32api.keybd_event(VK_CODE[word], 0, win32con.KEYEVENTF_KEYUP, 0)

# add choose script event
def choose_script_in_combo_box(event):
	script_status = False
	index = cb2.current()
	f = open(list_script[index], 'r')
	while 1:
		line = f.readline()
		if line:
			axis = line.split(',')
			if len(axis) != 2:
				script_status_desc.set(TIME_AXIS_ERROR_TEXT)
				cb2_description.configure(bg='red')
				break
			key = axis[1].strip()
			if len(key) != 1:
				script_status_desc.set(TIME_AXIS_ERROR_TEXT)
				cb2_description.configure(bg='red')
				break
			try:
				time_ms = int(axis[0])
				time_axis[axis[0]] = key
			except:
				script_status_desc.set(TIME_AXIS_ERROR_TEXT)
				cb2_description.configure(bg='red')
				break
		else:
			script_status = True
			script_status_desc.set(TIME_AXIS_VALID_TEXT)
			cb2_description.configure(bg='green')
			break
	f.close()

# refresh window when click combo box
def click_script_combo_box(event):
	list_file = os.listdir()
	list_script.clear()
	for file_name in list_file:
		if file_name.endswith('.tho'):
			list_script.append(file_name)
	cb2['values'] = list_script

# create window
window = tk.Tk()
window.title('PCR打轴器 - Thomas')
window.geometry('300x200')
window.resizable(width=False, height=False)
set_icon(window)

# create window chosen combo box
cb1_title = tk.Label(window, text = '窗口选择，请选择对应的模拟器')
cb1_title.pack()
cb1 = ttk.Combobox(window, width = 25)
cb1.bind('<Button-1>', click_window_combo_box)
cb1.bind('<<ComboboxSelected>>', choose_window_in_combo_box)
cb1.pack()

# create script chosen combo box
cb2_title = tk.Label(window, text = '时间轴选择')
cb2_title.pack()
cb2 = ttk.Combobox(window, width = 25)
cb2.bind('<Button-1>', click_script_combo_box)
cb2.bind('<<ComboboxSelected>>', choose_script_in_combo_box)
cb2.pack()
script_status_desc = tk.StringVar()
cb2_description = tk.Label(window, textvariable = script_status_desc, bg = 'yellow')
script_status_desc.set("尚未选择时间轴")
cb2_description.pack()

# create send button
start_button = tk.Button(window, text = '开始打轴', width = 10, height = 3)
start_button.bind('<Button-1>', start_battle)
start_button.pack()

window.mainloop()