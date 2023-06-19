import pyappauto.utils.errors

from threading import Thread
import pygetwindow as gw
import numpy as np
import subprocess
import win32con
import win32gui
import win32api
import win32ui
import time
import cv2
import os

__version__ = '1.0.1'


class Instance:
    def __init__(self, hwnd: int | None = None, debug: bool = False):
        """This is the constructor method for the class "Instance"

        Args:
            hwnd (int | None, optional): represents the handle to a window. Defaults to None.
            debug (bool, optional): if set to `True`, prints debug information to the console.
        """
        
        self._hwnd = hwnd
        self.__process = None
        self.__kill = False
        self._debug = debug

        if hwnd:
            Thread(target=self._close_handler).start()
    
    def __debug(self, title: str, content: str) -> None:
        if not self._debug:
            return
        
        print(f'[DEBUG] {title}: {content}')
    
    def _close_handler(self):
        self.__debug('Info', 'Created close handler.')

        while not self.__kill:
            if not win32gui.IsWindow(self._hwnd):
                break
            time.sleep(1)

        self.__debug('Info', 'Detected application close in close handler.')
        self.close()
    
    def open(self, exe_path: str, app_title: str, delay_seconds: float = 2) -> None:
        """ This function opens a process by path

        Args:
            exe_path (str): The path to the executable file that needs to be opened.
            app_title (str): The title of the application that will be created
            delay_seconds (float, optional): The delay between starting the application and getting the hwnd. Defaults to 2.

        Raises:
            utils.errors.ProcessAlreadyOpened: Raises if a process is already open.
            utils.errors.InvalidPath: Raises when the specified path could not be found.
            utils.errors.NoHwnd: Raises when the window could not be found.
        """
        if self._hwnd:
            raise utils.errors.ProcessAlreadyOpened()
        
        if not os.path.exists(exe_path):
            raise utils.errors.InvalidPath()

        self.__debug('Processing', 'Creating an application.')
        self.__process = subprocess.Popen(exe_path)

        time.sleep(delay_seconds)

        try:
            self._hwnd: int = gw.getWindowsWithTitle(app_title)[0]._hWnd
        except Exception:
            raise utils.errors.NoHwnd()
        
        if not self._hwnd:
            raise utils.errors.NoHwnd()
        
        self.__debug('Success', 'Created an application.')

        Thread(target=self._close_handler).start()

    def close(self) -> None:
        """ Closes the opened process

        Raises:
            utils.errors.NoOpenedProcess: Raises if a process is not opened.
        """
        if self.__kill:
            return
        
        if not self._hwnd:
            raise utils.errors.NoOpenedProcess()
        
        self.__kill = True
        self.__debug('Processing', 'Closing application.')

        if self.__process:
            self.__debug('Processing', 'Closing subprocess.')
            self.__process.terminate()
        else:
            win32gui.PostMessage(self._hwnd, win32con.WM_CLOSE, 0, 0)

        
        self.__process = None
        self._hwnd = None
        self.__debug('Success', 'Closed application.')

    def set_monitor(self, monitor_number: int = 0) -> None:
        """This function sets the position of a window to a specified monitor.

        Args:
            monitor_number (int, optional):  An integer representing the index of the monitor to set the window
            position to. Defaults to 0.

        Raises:
            utils.errors.NoOpenedProcess: Raises if no process is open.
            utils.errors.InvalidScreen: Raises if the screen index was not found, or is invalid.
        """
        if not self._hwnd:
            raise utils.errors.NoOpenedProcess()

        monitor_info = win32api.EnumDisplayMonitors(None, None)

        if monitor_number < len(monitor_info) and monitor_number > -(len(monitor_info)):
            monitor_rect = monitor_info[monitor_number][2]
            screen_left = monitor_rect[0]
            screen_top = monitor_rect[1]
            
            win32gui.SetWindowPos(self._hwnd, win32con.HWND_TOP, screen_left, screen_top, 0, 0, win32con.SWP_NOSIZE)

        else:
            raise utils.errors.InvalidScreen()

    def enter_full_screen(self) -> None:
        """ This function maximizes the window of a specified process.

        Raises:
            utils.errors.NoOpenedProcess: Raises if a process is not opened.
        """

        if not self._hwnd:
            raise utils.errors.NoOpenedProcess()

        self.__debug('Info', 'Entering full screen.')
        win32gui.ShowWindow(self._hwnd, win32con.SW_MAXIMIZE)

    def exit_full_screen(self) -> None:
        """ This function exits full screen mode by showing the window in normal mode.

        Raises:
            utils.errors.NoOpenedProcess: Raises if a process is not opened.
        """
        if not self._hwnd:
            raise utils.errors.NoOpenedProcess()
        
        self.__debug('Info', 'Exiting full screen.')
        win32gui.ShowWindow(self._hwnd, win32con.SW_NORMAL)

    def get_window_size(self) -> tuple[int, int]:
        """ This function returns the width and height of a window based on its handle.

        Raises:
            utils.errors.NoOpenedProcess: Raises if a process is not opened.

        Returns:
            tuple[int, int]: A tuple of two integers representing the width and height of the window associated with the `_hwnd` handle.
        """
        if not self._hwnd:
            raise utils.errors.NoOpenedProcess()
        
        rect = win32gui.GetWindowRect(self._hwnd)
        width = rect[2] - rect[0]
        height = rect[3] - rect[1]
        return width, height

    def get_window_position(self) -> tuple[int, int]:

        """ This function returns the position of a window as a tuple of integers.

        Raises:
            utils.errors.NoOpenedProcess: Raises if a process is not opened.

        Returns:
            tuple[int, int]: A tuple of two integers representing the x and y coordinates of the window position.
        """
        if not self._hwnd:
            raise utils.errors.NoOpenedProcess()
        
        rect = win32gui.GetWindowRect(self._hwnd)
        x = rect[0] if rect[0] >= 0 else 0
        y = rect[1] if rect[1] >= 0 else 0
        return x, y
    
    def set_window_dimensions(self, width: int, height: int) -> None:
        """ This function sets the dimensions of the window

        Args:
            width (int): represents the desired width of the window in pixels.
            height (int): represents the desired height of the window in pixels.

        Raises:
            utils.errors.NoOpenedProcess: Raises if a process is not opened.
        """
        if not self._hwnd:
            raise utils.errors.NoOpenedProcess()
        
        self.__debug('Info', f'Setting window dimensions to {width}x{height}.')

        x, y = self.get_window_position()
        win32gui.SetWindowPos(self._hwnd, 0, x, y, width, height, win32con.SWP_NOMOVE | win32con.SWP_NOZORDER)

    def set_window_position(self, x: int, y: int) -> None:
        """ This function sets the position of a window using the x and y coordinates provided as arguments.

        Args:
            x (int): The x-coordinate of the new position of the window
            y (int): The y-coordinate of the new position of the window

        Raises:
            utils.errors.NoOpenedProcess: Raises if a process is not opened.
        """
        if not self._hwnd:
            raise utils.errors.NoOpenedProcess()
        
        self.__debug('Info', f'Setting window position to ({x}, {y}).')
        win32gui.SetWindowPos(self._hwnd, win32con.HWND_TOP, x, y, 0, 0, win32con.SWP_NOSIZE)
    
    def left_click(self, x: int, y: int) -> None:
        """ This function simulates a left-click at a specified location on the window.

        Args:
            x (int): The x-coordinate of the point where the left click should be performed.
            y (int):  The y-coordinate of the point where the left click will be performed.

        Raises:
            utils.errors.NoOpenedProcess: Raises when no process is open.
        """
        if not self._hwnd:
            raise utils.errors.NoOpenedProcess()
        
        self.__debug('Info', f'Left clicking at ({x}, {y}).')
        win32gui.SendMessage(self._hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, win32api.MAKELONG(x, y))
        time.sleep(0.1)
        win32gui.SendMessage(self._hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, win32api.MAKELONG(x, y))

    def right_click(self, x: int, y: int) -> None:
        """ This function simulates a right-click at a specified location on the window.

        Args:
            x (int): The x-coordinate of the point where the right click should be performed.
            y (int):  The y-coordinate of the point where the right click will be performed.

        Raises:
            utils.errors.NoOpenedProcess: Raises when no process is open.
        """
        if not self._hwnd:
            raise utils.errors.NoOpenedProcess()

        self.__debug('Info', f'Right clicking at ({x}, {y}).')
        win32gui.SendMessage(self._hwnd, win32con.WM_RBUTTONDOWN, win32con.MK_RBUTTON, win32api.MAKELONG(x, y))
        time.sleep(0.1)
        win32gui.SendMessage(self._hwnd, win32con.WM_RBUTTONUP, win32con.MK_RBUTTON, win32api.MAKELONG(x, y))

    def right_double_click(self, x: int, y: int) -> None:
        if not self._hwnd:
            raise utils.errors.NoOpenedProcess()

        for _ in range(2):
            self.right_click(x, y)
    
    def left_double_click(self, x: int, y: int) -> None:
        if not self._hwnd:
            raise utils.errors.NoOpenedProcess()

        for _ in range(2):
            self.left_click(x, y)

    def press_key(self, key: int | str) -> None:
        """ The function simulates pressing a key on the keyboard for a specified window handle.

        Args:
            key (int | str): he key parameter can be either an integer or a string. If it is a string, it can be
                             a single character or one of the special keys "enter" or "tab". If it is an integer, it
                             represents the virtual key code of the key to be pressed.

        Raises:
            utils.errors.NoOpenedProcess: Raises when no process is open.
        """
        if not self._hwnd:
            raise utils.errors.NoOpenedProcess()
        
        self.__debug('Info', f'Pressing key {str(key)}.')
        if isinstance(key, str) and len(key) == 1:
            win32gui.SendMessage(self._hwnd, win32con.WM_CHAR, ord(key), 0)

        elif isinstance(key, str):
            special_keys = {
                "enter": win32con.VK_RETURN,
                "tab": win32con.VK_TAB
            }
            if key.lower() in special_keys:
                win32gui.SendMessage(self._hwnd, win32con.WM_KEYDOWN, special_keys[key.lower()], 0)
                time.sleep(0.1)
                win32gui.SendMessage(self._hwnd, win32con.WM_KEYUP, special_keys[key.lower()], 0)

        elif key:
            win32gui.SendMessage(self._hwnd, win32con.WM_KEYUP, key, 0)
            time.sleep(0.1)
            win32gui.SendMessage(self._hwnd, win32con.WM_KEYUP, key, 0)
            

    def typewrite(self, text: str, delay: float = 0.01) -> None:
        """ This function types out a given string of text character by character with a specified delay
            between each character.

        Args:
            text (str): The text that will be typed.
            delay (float, optional): The delay between each character press. Defaults to 0.01.

        Raises:
            utils.errors.NoOpenedProcess: Raises when no process is open.
        """
        if not self._hwnd:
            raise utils.errors.NoOpenedProcess()
        
        self.__debug('Info', f'Typing "{str(text)}".')
        for char in text:
            win32gui.SendMessage(self._hwnd, win32con.WM_CHAR, ord(char), 0)
            time.sleep(delay)
    
    def find_image_on_screen(self, image_path: str, confidence_threshhold: float = 0) -> tuple[int, int] | None:
        """ This function finds the center coordinates of an image within a window on the screen using
            template matching.

        Args:
            image_path (str): The path of the image that's being located
            confidence_threshhold (float, optional): The minimum confidence level that is required. Defaults to 0.

        Raises:
            utils.errors.NoOpenedProcess: Raises when no process has been opened.
            utils.errors.InvalidPath: Raises when the path provided is invalid.

        Returns:
            tuple[int, int] | None: The (x, y) coordinates of the image, or None if no image was found.
        """
        if not self._hwnd:
            raise utils.errors.NoOpenedProcess()
        
        if not os.path.exists(image_path):
            raise utils.errors.InvalidPath()
        
        target_image = cv2.imread(image_path)

        # Get the window dimensions
        rect = win32gui.GetWindowRect(self._hwnd)
        width = rect[2] - rect[0]
        height = rect[3] - rect[1]

        # Get the window screenshot
        window_dc = win32gui.GetWindowDC(self._hwnd)
        dc_obj = win32ui.CreateDCFromHandle(window_dc)
        mem_dc = dc_obj.CreateCompatibleDC()
        screenshot = win32ui.CreateBitmap()
        screenshot.CreateCompatibleBitmap(dc_obj, width, height)
        mem_dc.SelectObject(screenshot)
        mem_dc.BitBlt((0, 0), (width, height), dc_obj, (0, 0), win32con.SRCCOPY)

        # Convert the screenshot to a numpy array
        screenshot_info = screenshot.GetInfo()
        screenshot_data = screenshot.GetBitmapBits(True)
        img_np = np.frombuffer(screenshot_data, dtype=np.uint8).reshape(screenshot_info['bmHeight'], screenshot_info['bmWidth'], 4)

        # Convert BGR to RGB
        img_rgb = cv2.cvtColor(img_np, cv2.COLOR_BGR2RGB)

        # Perform template matching
        result = cv2.matchTemplate(img_rgb, target_image, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
        match_loc = max_loc
        confidence = max_val

        # Calculate the position within the window
        center_x = match_loc[0] + rect[0] + target_image.shape[1] // 2
        center_y = match_loc[1] + rect[1] + target_image.shape[0] // 2
        if confidence < confidence_threshhold:
            return None
        return (center_x, center_y)
    
    @property
    def hwnd(self) -> int:
        return self._hwnd
