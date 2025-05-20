import os
import json
import logging
import pandas as pd
from time import sleep
from src import WindowCapture
from src import Bot, BotState

# Configuração de logging
def configure_logging():
   if not os.path.exists('logs'):
      os.makedirs('logs')
   
   logging.basicConfig(
      filename='logs/.log', 
      level=logging.INFO, 
      format='%(asctime)s - %(levelname)s - %(message)s'
   )

   return logging.getLogger('bot_logger')

# Carregar configurações gerais do JSON
def load_config():
   if not os.path.exists('config'):
      os.makedirs('config')
      
   config_path = 'config/.json'
   
   # Ler o arquivo de configuração
   with open(config_path, encoding='utf8') as file:
      return json.load(file)

# 
def configure_input_output(save_path, excel_path):
   # Criar pasta se não existir
   if not os.path.exists(save_path):
      os.makedirs(save_path)

   # Verificar se o arquivo Excel existe
   if not os.path.exists(excel_path):
      df = pd.DataFrame({'CNPJ': pd.Series(dtype='str'), 'Status': pd.Series(dtype='str')})
      df.to_excel(excel_path, index=False)
      print(f"Arquivo Excel criado em {excel_path}")
   
   # Ler o arquivo Excel
   return pd.read_excel(excel_path)

def main():
   while(True):
      if bot.state == BotState.INICIO:
         continue

      elif wincap.stopped is True:
         wincap.start()

      bot.update_screenshot(wincap.screenshot)

      if wincap.screenshot is None or bot.screenshot is None:
         continue

      if bot.state == BotState.RETORNAR and bot.parar is not False:
         if bot.erro is not True:
            return True
         else:
            return False

if __name__ == "__main__":
   # Carrega configurações
   logger = configure_logging()
   config = load_config()

   save_path = config.get('save_path', 'Save')
   excel_path = config.get('excel_path', 'cnpjs.xlsx')

   df = configure_input_output(save_path, excel_path)

   # Inicializa componentes
   wincap = WindowCapture(None)
   bot = Bot(save_path,(wincap.offset_x, wincap.offset_y), (wincap.w, wincap.h))

   bot.start()

   try:
      for index, row in df.iterrows():
         if row.get('Status') == 'Baixado':
            logger.info(f"CNPJ {row['CNPJ']} já foi baixado. Pulando.")
            continue
               
         try:
            cnpj = str(row['CNPJ']).zfill(14)
            bot.update_CNPJ(cnpj)
            logger.info(f"Processando CNPJ: {cnpj}")

            bot.erro = False
            bot.parar = False

            # Atualiza o status no Excel
            if main():
               df.at[index, 'Status'] = 'Baixado'
            else:
               df.at[index, 'Status'] = 'Erro'
               
            # Salva as alterações no Excel após cada CNPJ
            df.to_excel(excel_path, index=False)
            
            sleep(2)
               
         except Exception as e:
            logger.error(f"Erro ao processar CNPJ {row['CNPJ']}: {e}")
            df.at[index, 'Status'] = 'Erro'
            df.to_excel(excel_path, index=False)
               
   except KeyboardInterrupt:
      logger.info("Processo interrompido pelo usuário")

   except Exception as e:
      logger.error(f"Erro geral: {e}")

   finally:
      df.to_excel(excel_path, index=False)

      wincap.stop()
      bot.stop()