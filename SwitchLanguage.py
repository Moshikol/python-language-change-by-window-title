import win32gui
import win32api
import py_win_keyboard_layout
import threading


def set_interval(func, time):
    e = threading.Event()
    while not e.wait(time):
        func()


win32gui.window


def set_keyboard_layout():
    w = win32gui
    window_title = w.GetWindowText(w.GetForegroundWindow())
    print(window_title)
    if 'whatsapp' in window_title.lower():
        py_win_keyboard_layout.change_foreground_window_keyboard_layout(-264436723)  # to switch to Heb
        print("Keyboard changed" + str(win32api.GetKeyboardLayout()))
    elif 'visual studio code' in window_title.lower():
        py_win_keyboard_layout.change_foreground_window_keyboard_layout(67699721)  # to switch to Heb
        print("Keyboard changed" + str(win32api.GetKeyboardLayout()))
    elif 'chrome' in window_title.lower():
        py_win_keyboard_layout.change_foreground_window_keyboard_layout(67699721)  # to switch to Heb
        print("Keyboard changed" + str(win32api.GetKeyboardLayout()))


set_interval(set_keyboard_layout, 0.1)
# py_win_keyboard_layout.change_foreground_window_keyboard_layout(67699721)  # to switch to english us


# print(py_win_keyboard_layout.get_keyboard_layout_list())
