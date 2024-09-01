import win32gui
import win32ui
import win32con
import win32api
import time
import pyautogui
from pywinauto import Application
from pywinauto.keyboard import send_keys
from PIL import Image


score_screenshot = [1 , "lobby"]

def main(name_window):

    hwnd = win32gui.FindWindow(None, name_window)

    if hwnd:
        print("Все кайф робит")
    else:
        print("Окно не найдено!")

    screenshot_window(hwnd)




def click_in_window_at_coords(window_title, x, y):
    """
    Кликает в указанную точку (x, y) в окне с заданным заголовком.

    :param window_title: Заголовок окна, в котором нужно кликнуть.
    :param x: Координата x для клика.
    :param y: Координата y для клика.
    """
    try:
        app = Application().connect(title=window_title)
        
        window = app.window(title=window_title)
        
        window.set_focus()
        
        window.click_input(coords=(x, y))
        
        print(f"Кликнули в ({x}, {y}) в окне '{window_title}'.")

    except Exception as e:
        print(f"Не удалось кликнуть в окне '{window_title}': {e}")


def send_arrow_keys_to_window(window_title, direction, hits=1):
    """
    Устанавливает фокус на указанное окно и отправляет нажатия клавиш стрелок.

    :param window_title: Заголовок окна, на которое нужно установить фокус.
    :param direction: Направление стрелки (up, down, left, right).
    :param hits: Количество нажатий стрелки (по умолчанию 1).
    """
    try:
        app = Application().connect(title=window_title)

        window = app.window(title=window_title)
        
        window.set_focus()

        for _ in range(hits):
            match direction:
                case 'up':
                    send_keys('{UP}')
                case 'down':
                    send_keys('{DOWN}')
                case 'left':
                    send_keys('{LEFT}')
                case 'right':
                    send_keys('{RIGHT}')
                case _:
                    print(f"Неверное направление: {direction}. Пропускаем.")
                    return
        
        print(f"Отправлены стрелки в окно '{window_title}': {direction} {hits} раз.")
        
    except Exception as e:
        print(f"Не удалось отправить стрелки в окно '{window_title}': {e}")


def move_mouse_to_window(window_title, x_offset=0, y_offset=0):
    """
    Перемещает курсор мыши к заданным координатам внутри указанного окна и устанавливает на него фокус.

    :param window_title: Заголовок окна, к которому нужно переместить курсор.
    :param x_offset: Смещение по оси X относительно верхнего левого угла окна.
    :param y_offset: Смещение по оси Y относительно верхнего левого угла окна.
    """
    try:
        app = Application().connect(title=window_title)

        window = app.window(title=window_title)

        window.set_focus()

        rect = window.rectangle()
        x, y = rect.left, rect.top
        
        target_x = x + x_offset
        target_y = y + y_offset

        win32api.SetCursorPos((target_x, target_y))

        print(f"Курсор перемещен к окну '{window_title}' на координаты ({target_x}, {target_y}) и фокус установлен.")
    except Exception as e:
        print(f"Не удалось переместить курсор в окно '{window_title}': {e}")


def scroll_window(window_title, scroll_amount=1, direction='down'):
    """
    Прокручивает указанное окно вверх или вниз.

    :param window_title: Заголовок окна, которое нужно прокрутить.
    :param direction: Направление прокрутки ('up' или 'down').
    :param scroll_amount: Количество прокруток.
    """
    try:
        app = Application().connect(title=window_title)

        window = app.window(title=window_title)
        hwnd = window.handle

        window.set_focus()
      
        scroll_delta = -scroll_amount if direction == 'up' else scroll_amount

        for _ in range(abs(scroll_amount)):
            win32gui.SendMessage(hwnd, win32con.WM_VSCROLL, win32con.SB_LINEUP if direction == 'up' else win32con.SB_LINEDOWN, 0)

        print(f"Прокрутка {direction} на {scroll_amount} раз(а) в окне: '{window_title}'")
    except Exception as e:
        print(f"Не удалось прокрутить окно '{window_title}': {e}")


def press_escape_in_window(window_title):
    try:
        app = Application(backend="uia").connect(title=window_title)
        
        window = app.window(title=window_title)
        
        window.set_focus()
        
        send_keys('{ESC}')
        print(f"Клавиша Esc была отправлена в окно: '{window_title}'")
    except Exception as e:
        print(f"Не удалось отправить клавишу Esc в окно '{window_title}': {e}")


def send_text_to_window(window_title, text):
    try:
        app = Application(backend="uia").connect(title=window_title)
      
        window = app.window(title=window_title)
        
        window.set_focus()
        
        send_keys(text)
        print(f"Текст '{text}' был отправлен в окно: '{window_title}'")
    except Exception as e:
        print(f"Не удалось отправить текст в окно '{window_title}': {e}")


def print_windows(hwnd):
    def enum_windows_callback(hwnd, results):
        if win32gui.IsWindowVisible(hwnd):
            window_text = win32gui.GetWindowText(hwnd)
            if window_text:
                results.append((hwnd, window_text))

    windows = []
    win32gui.EnumWindows(enum_windows_callback, windows)

    for hwnd, window_text in windows:
        print(f"HWND: {hwnd}, Title: {window_text}")


def get_inner_windows(whndl):
    def callback(hwnd, hwnds):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            hwnds[win32gui.GetClassName(hwnd)] = hwnd
        return True
    hwnds = {}
    win32gui.EnumChildWindows(whndl, callback, hwnds)
    return hwnds


def screenshot_window(hwnd):
    left, top, right, bot = win32gui.GetWindowRect(hwnd)
    width = right - left
    height = bot - top

    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()

    saveBitMap = win32ui.CreateBitmap()
    saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
    saveDC.SelectObject(saveBitMap)

    saveDC.BitBlt((0, 0), (width, height), mfcDC, (0, 0), win32con.SRCCOPY)

    bmpinfo = saveBitMap.GetInfo()
    bmpstr = saveBitMap.GetBitmapBits(True)
    im = Image.frombuffer(
        'RGB',
        (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
        bmpstr, 'raw', 'BGRX', 0, 1
    )

    win32gui.DeleteObject(saveBitMap.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)

    im.save("C:\\Users\\SoVa\\Desktop\\screenshot\\{}{}.png".format(str(score_screenshot[0]), score_screenshot[1]))
    score_screenshot[0] = score_screenshot[0] + 1

if __name__ == '__main__':
    main("Название окна")
