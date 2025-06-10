import pyautogui as pa
import time
from pyscreeze import locateCenterOnScreen as lcos
from pyperclip import copy
import os
import pandas as pd

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

def funcao_input_int(mensagem):
    while True:
        numero = (input(mensagem))
        verificar_numero = numero.isnumeric()
        if verificar_numero is True:
            numero = int(numero)
            return numero

        else:
            print("\33[31m"+"Campo obrigatório, só é aceito números."+"\33[0;0m")
            time.sleep(0.3)

def extrair_pedido(num_nf):
    localizar_clicar("IW_icone.png")
    pa.press("tab", presses=5)  # clicar na primeira linha
    pa.press("Home")  # ir para a primeira coluna
    pa.hotkey("ctrl", "end")  # descer para a última linha da tela de pedido
    pa.press("right", presses=19)  # ira para coluna valor item pedido
    time.sleep(0.5)
    localizar_clicar("Vlr_Item_Pedido.png", click_type='right')  # clicar no campo número de embalagens
    time.sleep(0.5)
    pa.press("down", presses=8)
    pa.press("ENTER")
    time.sleep(2)
    localizar_clicar("Vlr_Item_Pedido.png", click_type='right')  # clicar no campo número de embalagens
    time.sleep(0.5)
    pa.press("down", presses=6)
    pa.press("ENTER")
    localizar_clicar("Inicio.png")
    localizar_clicar("Nome_Arquivo.png")
    copiar_colar(num_nf)
    localizar_clicar("Salvar.png")
    localizar_clicar("Exportacao.png", click_type="move")
    pa.press("ENTER")
    return


numero_nf = funcao_input_int("Digite o número da NF: ")
extrair_pedido(numero_nf)

desktop = os.path.join(os.path.expanduser("~"), "OneDrive", "Área de Trabalho")
arquivo = os.path.join(desktop, f"{numero_nf}.csv")
df = pd.read_csv(arquivo, encoding="ISO-8859-1", sep=";")
colunas = df.columns.tolist()
print("Lista de colunas:", colunas)
input()
colunas_desejadas = [
    "Material",
    "Nome do Material",
    "U.M.",
    "Nº Embal. Entregues",
    "Vlr Item Pedido"
]
df_filtrado = df[colunas_desejadas].copy()

# Remove linhas onde "Nº Embal. Entregues" está vazio ou nulo
df_filtrado = df_filtrado[df_filtrado["Nº Embal. Entregues"].notna() & (df_filtrado["Nº Embal. Entregues"] != '')]

arquivo_saida = os.path.join(f"pedido_NF_{numero_nf}.xlsx")
while True:
    try:
        # Tenta salvar o Excel
        df_filtrado.to_excel(arquivo_saida, index=False)
        break

    except PermissionError:
        print("\n❌ Não foi possível salvar o arquivo.")
        input("Feche o arquivo e pressione ENTER para continuar...")

# Abrir o arquivo
try:
    os.startfile(arquivo_saida)
except Exception as e:
    print(f"Não foi possível abrir o arquivo automaticamente: {e}")