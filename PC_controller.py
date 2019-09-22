import pyautogui
import ctypes
import time
import socket
import win32api
import win32con
import win32gui
import sys

SCREEN_WIDTH = ctypes.windll.user32.GetSystemMetrics(0)
SCREEN_HEIGHT = ctypes.windll.user32.GetSystemMetrics(1)


#key simulation code from: https://gist.github.com/silmang/9c352b39fc494588dd03896cdecb806e
SendInput = ctypes.windll.user32.SendInput

# C struct redefinitions
PUL = ctypes.POINTER(ctypes.c_ulong)


class KeyBdInput(ctypes.Structure):
    _fields_ = [("wVk", ctypes.c_ushort),
                ("wScan", ctypes.c_ushort),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", PUL)]


class HardwareInput(ctypes.Structure):
    _fields_ = [("uMsg", ctypes.c_ulong),
                ("wParamL", ctypes.c_short),
                ("wParamH", ctypes.c_ushort)]


class MouseInput(ctypes.Structure):
    _fields_ = [("dx", ctypes.c_long),
                ("dy", ctypes.c_long),
                ("mouseData", ctypes.c_ulong),
                ("dwFlags", ctypes.c_ulong),
                ("time",ctypes.c_ulong),
                ("dwExtraInfo", PUL)]


class Input_I(ctypes.Union):
    _fields_ = [("ki", KeyBdInput),
                 ("mi", MouseInput),
                 ("hi", HardwareInput)]

class Input(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong),
                ("ii", Input_I)]


# Actual Functions
def PressKey(hexKeyCode):
    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput(0, hexKeyCode, 0x0008, 0, ctypes.pointer(extra))
    x = Input( ctypes.c_ulong(1), ii_ )
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))

def ReleaseKey(hexKeyCode):
    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput(0, hexKeyCode, 0x0008 | 0x0002, 0, ctypes.pointer(extra))
    x = Input(ctypes.c_ulong(1), ii_)
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))

def alt_tab():
    PressKey(56)
    PressKey(15)
    ReleaseKey(15)
    time.sleep(0.1)
    ReleaseKey(56)

def move_cursor_to(x, y):
    temp = (x/SCREEN_WIDTH*65535.0)
    temp2 = (y/SCREEN_WIDTH*65535.0)

    print("movement before rounding: ")
    print("x: ")
    print(temp)
    print("y: ")
    print(temp2)

    temp = int(x/SCREEN_WIDTH*65535.0)
    temp2 = int(y/SCREEN_WIDTH*65535.0)

    print("actual movement: ")
    print("x: ")
    print(temp)
    print("y: ")
    print(temp2)
    win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, int(x/SCREEN_WIDTH*65535.0), int(y/SCREEN_HEIGHT*65535.0))

def move_cursor(x_movement, y_movement):
    flags, hcursor, (original_x, original_y) = win32gui.GetCursorInfo()
    move_cursor_to(original_x + 1 * x_movement, original_y + 1 * y_movement)

def press_left_click(original_x, original_y):
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, original_x, original_y)
    print("Left click pressed")

def release_left_click(original_x, original_y):
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, original_x, original_y)
    print("Left click released")

def press_right_click(original_x, original_y):
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, original_x, original_y)
    print("Right click pressed")

def release_right_click(original_x, original_y):
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, original_x, original_y)
    print("Right click released")

def press_middle_click():
    pyautogui.click(button='middle')
    # win32api.mouse_event(win32con.MOUSEEVENTF_MIDDLEDOWN, original_x, original_y)
    print("Middle click pressed")

def release_middle_click(original_x, original_y):
    win32api.mouse_event(win32con.MOUSEEVENTF_MIDDLEUP, original_x, original_y)
    print("Middle click released")

def swipe(x_movement, y_movement):
    move_cursor(x_movement, y_movement)
    print("Swiping " + str(x_movement) + ", " + str(y_movement))


def handle_client_connection(client_socket):
    request_byte = client_socket.recv(21)
    print("reqeust_byte size: " + str(sys.getsizeof(request_byte)))
    request = request_byte.decode()
    print("Received " + request)

    if request[0] == "+":
        movement = request[1:]
        try:
            x_movement, y_movement = movement.split("_")
        except:
            print("skipping")
            return
        if len(x_movement) > 4:
            x_movement = x_movement[0:len(x_movement) - 3]
        if len(y_movement) > 4:
            y_movement = y_movement[0:len(y_movement) - 3]
        x_movement_float = ""
        y_movement_float = ""
        clear_to_proceed = True
        try:
            x_movement_float = float(x_movement)
            y_movement_float = float(y_movement)
        except:
            print("skipping")
            clear_to_proceed = False

        print("x_movement: " + x_movement + ", y_movement: " + y_movement)
        if clear_to_proceed and abs(x_movement_float) > 0.001 and abs(y_movement_float) > 0.001:
            swipe(x_movement_float, y_movement_float)

    elif request == "p_win":
        pyautogui.keyDown("win")
    elif request == "r_win":
        pyautogui.keyUp("win")

    else:
        try:
            scancode = int(request[2:5])
            print("scancode: ")
            print(scancode)
            is_mouse_event = scancode == 256 or scancode == 256 or scancode == 257
            if is_mouse_event:
                flags, hcursor, (original_x, original_y) = win32gui.GetCursorInfo()
            if request[0] == 'p':
                if scancode == 256:
                    press_left_click(original_x, original_y)
                elif scancode == 257:
                    press_right_click(original_x, original_y)
                elif scancode == 258:
                    press_middle_click()
                else:
                    PressKey(scancode)
            else:
                if scancode == 256:
                    release_left_click(original_x, original_y)
                elif scancode == 257:
                    release_right_click(original_x, original_y)
                #elif scancode == 258:
                    #release_middle_click(original_x, original_y)
                else:
                    ReleaseKey(scancode)
        except Exception as e:
            print(e)
            print("skipping")


port_num = 15200
listening_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
listening_socket.bind(('', port_num))
listening_socket.listen(50)
print("Listening on {}:{}".format(' ', port_num))
client_socket, address = listening_socket.accept()
print("Accepted connection from {}:{}".format(address[0], address[1]))

while True:
    handle_client_connection(client_socket)