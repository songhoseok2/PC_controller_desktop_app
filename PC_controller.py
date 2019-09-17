import ctypes
import time
import socket
import win32api
import win32con
import win32gui
import sys

SCREEN_WIDTH = ctypes.windll.user32.GetSystemMetrics(0)
SCREEN_HEIGHT = ctypes.windll.user32.GetSystemMetrics(1)

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


def press_W():
    PressKey(17)
    print("W pressed")

def release_W():
    ReleaseKey(17)
    print("W released")

def press_A():
    PressKey(30)
    print("A pressed")

def release_A():
    ReleaseKey(30)
    print("A released")

def press_S():
    PressKey(31)
    print("S pressed")

def release_S():
    ReleaseKey(31)
    print("S released")

def press_D():
    PressKey(32)
    print("D pressed")

def release_D():
    ReleaseKey(32)
    print("D released")

def press_Q():
    PressKey(16)
    print("Q pressed")

def release_Q():
    ReleaseKey(16)
    print("Q released")

def press_E():
    PressKey(18)
    print("E pressed")

def release_E():
    ReleaseKey(18)
    print("E released")

def press_Ctrl():
    PressKey(29)
    print("Ctrl pressed")

def release_Ctrl():
    ReleaseKey(29)
    print("Ctrl released")

def press_C():
    PressKey(46)
    print("C pressed")

def release_C():
    ReleaseKey(46)
    print("C released")

def press_R():
    PressKey(19)
    print("R pressed")

def release_R():
    ReleaseKey(19)
    print("R released")

def move_cursor_to(x, y):
    win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, int(x/SCREEN_WIDTH*65535.0), int(y/SCREEN_HEIGHT*65535.0))

def move_cursor(x_movement, y_movement):
    flags, hcursor, (original_x, original_y) = win32gui.GetCursorInfo()
    move_cursor_to(original_x + x_movement, original_y + y_movement)

def press_left_click():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, int(SCREEN_WIDTH/2), int(SCREEN_HEIGHT/2))
    print("Left click pressed")

def release_left_click():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, int(SCREEN_WIDTH/2), int(SCREEN_HEIGHT/2))
    print("Left click released")

def press_right_click():
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, int(SCREEN_WIDTH/2), int(SCREEN_HEIGHT/2))
    print("Right click pressed")

def release_right_click():
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, int(SCREEN_WIDTH/2), int(SCREEN_HEIGHT/2))
    print("Right click released")

def swipe(x_movement, y_movement):
    move_cursor(x_movement, y_movement)
    print("Swiping" + str(x_movement) + str(y_movement))


def handle_client_connection(client_socket):
    request_byte = client_socket.recv(1024)
    request = request_byte.decode()
    print("Received " + request)

    if request == "press_W":
        press_W()
    elif request == "release_W":
        release_W()
    elif request == "press_A":
        press_A()
    elif request == "release_A":
        release_A()
    elif request == "press_S":
        press_S()
    elif request == "release_S":
        release_S()
    elif request == "press_D":
        press_D()
    elif request == "release_D":
        release_D()
    elif request == "press_R":
        press_R()
    elif request == "release_R":
        release_R()
    elif request == "press_Q":
        press_Q()
    elif request == "release_Q":
        release_Q()
    elif request == "press_E":
        press_E()
    elif request == "release_E":
        release_E()
    elif request == "press_C":
        press_C()
    elif request == "release_C":
        release_C()
    elif request == "press_Ctrl":
        press_Ctrl()
    elif request == "release_Ctrl":
        release_Ctrl()
    elif request == "press_left_click":
        press_left_click()
    elif request == "release_left_click":
        release_left_click()
    elif request == "press_right_click":
        press_right_click()
    elif request == "release_right_click":
        release_right_click()

    elif request == "swipe":
        movement = client_socket.recv(1024)
        movement_str = movement.decode()
        x_movement, y_movement = movement_str.split("_")
        print("x: " + x_movement)
        print("y: " + y_movement)
        swipe(float(x_movement), float(y_movement))

port_num = 15200
listening_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
listening_socket.bind(('', port_num))
listening_socket.listen(50)
print("Listening on {}:{}".format(' ', port_num))
client_socket, address = listening_socket.accept()
print("Accepted connection from {}:{}".format(address[0], address[1]))

while True:
    handle_client_connection(client_socket)

