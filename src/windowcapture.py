import numpy as np
import win32gui, win32ui, win32con
from threading import Thread, Lock

class WindowCapture:
   def __init__(self, window_name=None):
      self.lock = Lock()
      self.stopped = True
      self.screenshot = None

      self.w = 0
      self.h = 0
      self.hwnd = None
      self.cropped_x = 0
      self.cropped_y = 0
      self.offset_x = 0
      self.offset_y = 0

      if window_name is None:
         self.hwnd = win32gui.GetDesktopWindow()
      else:
         self.hwnd = win32gui.FindWindow(None, window_name)
         if not self.hwnd:
            raise Exception(f'Window not found: {window_name}')
      
      window_rect = win32gui.GetWindowRect(self.hwnd)
      self.w = window_rect[2] - window_rect[0]
      self.h = window_rect[3] - window_rect[1]

      border_pixels = 8
      titlebar_pixels = 30
      self.w = self.w - (border_pixels * 2)
      self.h = self.h - titlebar_pixels - border_pixels
      self.cropped_x = border_pixels
      self.cropped_y = titlebar_pixels

      self.offset_x = window_rect[0] + self.cropped_x
      self.offset_y = window_rect[1] + self.cropped_y
      
   def get_screenshot(self):
      wDC = win32gui.GetWindowDC(self.hwnd)
      dcObj = win32ui.CreateDCFromHandle(wDC)
      cDC = dcObj.CreateCompatibleDC()
      dataBitMap = win32ui.CreateBitmap()
      dataBitMap.CreateCompatibleBitmap(dcObj, self.w, self.h)
      cDC.SelectObject(dataBitMap)
      cDC.BitBlt((0, 0), (self.w, self.h), dcObj, (self.cropped_x, self.cropped_y), win32con.SRCCOPY)
      
      signedIntsArray = dataBitMap.GetBitmapBits(True)
      img = np.fromstring(signedIntsArray, dtype='uint8')
      img.shape = (self.h, self.w, 4)

      dcObj.DeleteDC()
      cDC.DeleteDC()
      win32gui.ReleaseDC(self.hwnd, wDC)
      win32gui.DeleteObject(dataBitMap.GetHandle())

      img = img[..., :3]

      img = np.ascontiguousarray(img)   

      return img
   
   def start(self):
      self.stopped = False
      t = Thread(target=self.run)
      t.start()

   def stop(self):
      self.stopped = True

   def run(self):
      while not self.stopped:
         screenshot = self.get_screenshot()

         with self.lock:
            self.screenshot = screenshot