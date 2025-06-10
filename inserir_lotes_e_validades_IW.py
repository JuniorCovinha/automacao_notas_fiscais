import pyautogui as pa
import time
from pyscreeze import locateCenterOnScreen as lcos
from pyperclip import copy
import os
import pandas as pd

# Caminho para o arquivo
desktop = os.path.join(os.path.expanduser("~"), "OneDrive", "Área de Trabalho")
planilha_padrao = "NFe_Planilha_Padrão"
excel_saida = os.path.join(desktop, f"{planilha_padrao}.xlsx")
IMAGENS_DIR = r'C:\Users\solar\PycharmProjects\pythonProject1\.venv\Solar\Imagens'

def localizar_clicar(imagem, click_type='left', retries=3, wait_time=3, confidence=0.9, teste=False, verificar_deposito=False):
    imagem_path = os.path.join(IMAGENS_DIR, imagem)
    while True:
        for _ in range(retries):
            try:
                ximagem, yimagem = lcos(imagem_path, confidence=confidence, grayscale=False)
                if click_type == 'left':
                    pa.click(ximagem, yimagem)
                elif click_type == 'right':
                    pa.rightClick(ximagem, yimagem)
                elif click_type == 'move':
                    pa.moveTo(ximagem, yimagem)
                return
            except Exception as e:
                time.sleep(wait_time)
        if teste:
            return
        if not teste:
            print(
                "\33[31m"+f"Falha ao localizar a imagem {imagem} após {retries} tentativas, ajuste a tela e aperte qualquer tecla"+"\33[0;0m"
            )
            passar_ou_nao = input().capitalize()
            if passar_ou_nao == "P":
                localizar_clicar("IW_icone.png")
                return
            else:
                localizar_clicar("IW_icone.png")

        if verificar_deposito == True:
            for _ in range(retries):
                try:
                    imagem_path = os.path.join(IMAGENS_DIR, "Deposito_Destino.png")
                    ximagem, yimagem = lcos(imagem_path, confidence=confidence, grayscale=False)
                    pa.click(ximagem, yimagem)
                    pa.press("F")
                    return
                except Exception as e:
                    print(f"Erro ao localizar a imagem {imagem} {e}")
                    time.sleep(wait_time)
        else:
            pass

def copiar_colar(texto):
    copy(texto)
    time.sleep(0.3)
    pa.hotkey("ctrl", "v")

def ajustar_dia(data):
    if pd.isna(data):
        return None
    if data.month == 2:
        return data.replace(day=28)
    return data.replace(day=30)

# Leitura da planilha
df = pd.read_excel(excel_saida)

pa.PAUSE = 0.3
# Verifica se as colunas necessárias existem (case insensitive)
colunas = df.columns.str.capitalize()
if 'Lote' in colunas and 'Validade' in colunas:
    # Acessa as colunas com o nome real
    col_lote = df.columns[colunas.get_loc('Lote')]
    col_validade = df.columns[colunas.get_loc('Validade')]
    df[col_validade] = pd.to_datetime(df[col_validade], errors='coerce')

    # Filtra apenas as colunas desejadas e inverte a ordem
    df_filtrado = df[[col_lote, col_validade]].iloc[::-1]

    # Ajusta os dias para o padrão de dia 30
    df_filtrado[col_validade] = df_filtrado[col_validade].apply(ajustar_dia)

    # Formata a coluna validade para dd/mm/aaaa
    df_filtrado[col_validade] = pd.to_datetime(df_filtrado[col_validade], errors='coerce').dt.strftime('%d/%m/%Y')

    localizar_clicar("IW_icone.png")
    pa.press("tab", presses=3)
    pa.hotkey("ctrl", "home")
    pa.press("home")
    pa.hotkey("ctrl", "end")
    time.sleep(0.3)
    pa.press("right", presses=14)
    time.sleep(0.3)

    for index, row in df_filtrado.iterrows():
        lote = str(row[col_lote])
        validade = str(row[col_validade])

        pa.press("insert")
        copiar_colar(lote)
        pa.press("tab")
        pa.press("backspace", presses=3)
        time.sleep(0.3)
        copiar_colar(validade)
        time.sleep(0.1)
        pa.press("up")
        pa.press("left")

else:
    print("As colunas 'lote' e/ou 'validade' não foram encontradas na planilha.")