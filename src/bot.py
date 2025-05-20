import logging
import keyboard
import pyautogui
import cv2 as cv
import webbrowser
import numpy as np
from enum import Enum
from math import sqrt
from time import sleep
from threading import Thread, Lock

class BotState(Enum):
   INICIO = 0
   ESTAGIO1 = 1
   ESTAGIO2 = 2
   ESTAGIO3 = 3
   RESULTADO = 4
   RETORNAR = 5

class Bot:
   IGNORE_RADIUS = 130
   TOOLTIP_MATCH_THRESHOLD = 0.75

   def __init__(self, save_path, window_offset, window_size):
      self.lock = Lock()
      self.stopped = True

      self.cnpj = None
      self.save_path = save_path
      self.url = "https://solucoes.receita.fazenda.gov.br/servicos/certidaointernet/pj/emitir"

      self.logger = logging.getLogger('bot_logger')

      self.window_offset = window_offset
      self.window_w = window_size[0]
      self.window_h = window_size[1]

      self.images = [
         cv.imread('img/estagio1.jpg', cv.IMREAD_UNCHANGED),   # 0
         cv.imread('img/estagio2.jpg', cv.IMREAD_UNCHANGED),   # 1
         cv.imread('img/estagio3.jpg', cv.IMREAD_UNCHANGED),   # 2
         cv.imread('img/resultado.jpg', cv.IMREAD_UNCHANGED),  # 3
         cv.imread('img/reiniciar.jpg', cv.IMREAD_UNCHANGED),  # 4
      ]

      self.erro = False
      self.parar = False
      self.rectangles = None
      self.screenshot = None
      self.state = BotState.INICIO

   def targets_ordered_by_distance(self, targets):
      my_pos = (int(self.window_w / 2), int(self.window_h / 2))

      def pythagorean_distance(pos):
         return sqrt((pos[0] - my_pos[0])**2 + (pos[1] - my_pos[1])**2)
      target = sorted(targets, key=pythagorean_distance)

      return target
  
   def get_screen_position(self, pos):
      return (pos[0] + self.window_offset[0], pos[1] + self.window_offset[1])
   
   def confirm_tooltip(self, image, method = cv.TM_CCOEFF_NORMED):
      result = cv.matchTemplate(self.screenshot, image, method)
      min_val, max_val, min_loc, max_loc = cv.minMaxLoc(result)

      if max_val >= self.TOOLTIP_MATCH_THRESHOLD:
         locations = np.where(result >= self.TOOLTIP_MATCH_THRESHOLD)
         locations = list(zip(*locations[::-1]))

         needle_w = image.shape[1]
         needle_h = image.shape[0]

         rectangles = []
         for loc in locations:
            rect = [int(loc[0]), int(loc[1]), needle_w, needle_h]

            rectangles.append(rect)
            rectangles.append(rect)

         self.rectangles, _ = cv.groupRectangles(rectangles, groupThreshold=1, eps=0.5)

         return True

      return False
   
   def get_click_points(self, rectangles):
      points = []

      for (x, y, w, h) in rectangles:
         center_x = x + int(w/2)
         center_y = y + int(h/2)
         points.append((center_x, center_y))

      return points

   def estagio(self, x=0, y=0):
      point = self.get_click_points(self.rectangles)
      target_pos = self.targets_ordered_by_distance(point)
      screen_x, screen_y = self.get_screen_position(target_pos[0])
      pyautogui.moveTo(x=screen_x+x, y=screen_y+y)
      sleep(1.250)
      pyautogui.click()

   def update_CNPJ(self, cnpj):
      with self.lock:
         self.cnpj = cnpj
         
   def update_screenshot(self, screenshot):
      with self.lock:
         self.screenshot = screenshot
   
   def start(self):
      self.stopped = False
      t = Thread(target=self.run)
      t.start()

   def stop(self):
      self.stopped = True
   
   def run(self):
      while not self.stopped:
         if self.state == BotState.INICIO:
            with self.lock:
               webbrowser.open(self.url)

               self.state = BotState.ESTAGIO1

         if self.screenshot is not None:
            if self.state == BotState.ESTAGIO1:
               if self.confirm_tooltip(self.images[0]):
                  self.estagio()

                  pyautogui.hotkey('ctrl', 'shift', 'j')

                  with self.lock:
                     self.state = BotState.ESTAGIO2

            elif self.state == BotState.ESTAGIO2:
               if self.confirm_tooltip(self.images[1]):
                  self.estagio(0, 50)

                  keyboard.write(f'document.querySelector("#NI").value = "{self.cnpj}"')
                  pyautogui.press('enter')
                  sleep(1)
                  keyboard.write('document.querySelector("#validar").click()')
                  pyautogui.press('enter')

                  with self.lock:
                     self.state = BotState.ESTAGIO3

            elif self.state == BotState.ESTAGIO3:
               sleep(3)
               
               if self.confirm_tooltip(self.images[4]):
                  # Erro 1

                  with self.lock:
                     self.erro = True

                     self.state = BotState.RETORNAR

               elif self.confirm_tooltip(self.images[3]):
                  with self.lock:
                     self.state = BotState.RESULTADO

               elif self.confirm_tooltip(self.images[2]):
                  if self.confirm_tooltip(self.images[1]):
                     self.estagio(0, 50)

                     keyboard.write('document.querySelector("#FrmSelecao > a:nth-child(6)").click()')
                     pyautogui.press('enter')

                     with self.lock:
                        self.state = BotState.RESULTADO

            elif self.state == BotState.RESULTADO:
               if self.confirm_tooltip(self.images[3]):
                  self.estagio(5, 0)
                  
                  # Seleciona caminho
                  pyautogui.hotkey('alt', 'e')
                  sleep(1)
                  keyboard.write(self.save_path)
                  pyautogui.press('enter')
                  sleep(1)
                  
                  # Renomear arquivo
                  # pyautogui.hotkey('alt', 'n')
                  # keyboard.write(self.cnpj)
                  # sleep(1)
                  
                  # Confirmar salvamento
                  pyautogui.hotkey('alt', 'l')

                  with self.lock:
                     self.state = BotState.RETORNAR
            
            elif self.state == BotState.RETORNAR:
               if self.confirm_tooltip(self.images[4]):
                  self.estagio()

                  self.parar = True
                  pyautogui.hotkey('ctrl', 'w')
                  sleep(1)

                  with self.lock:
                     self.state = BotState.INICIO