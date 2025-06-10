import pyautogui as pa
import time
from pyscreeze import locateCenterOnScreen as lcos
from random import randint
from pyperclip import copy
import os
from datetime import date


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

def preencher_todos(texto):
    if texto not in "0":
        c = 0
        while True:
            if c < linhas:
                pa.press("backspace", presses=3, interval=0.1)
                time.sleep(0.1)
                copiar_colar(f"{texto}")
                c = c + 1
            if c >= linhas:
                break
            pa.press("up")

def informacoes_boleto(linha_digitavel):
    def calcular_data_vencimento(fator_vencimento):
        from datetime import datetime, timedelta
        data_base = datetime(2022, 5, 29) if fator_vencimento < 9000 else datetime(1997, 10, 7)
        data_vencimento = data_base + timedelta(days=fator_vencimento)
        return data_vencimento.strftime('%d/%m/%Y')

    def identificar_digitavel(linha_digitavel):
        """
        Identifica se o boleto foi bipado ou digitado a partir da linha digitável.
        """
        qnt_digitos = len(str(linha_digitavel))
        if qnt_digitos == 44:
            return qnt_digitos

        elif qnt_digitos == 47:
            return qnt_digitos

    def calcular_valor_boleto(linha_digitavel, qnt_digitos):
        """
        Calcula o valor do boleto a partir da linha digitável.h
        """
        # Valor do boleto está nas posições 37 a 46 para boletos de cobrança
        valor_str = linha_digitavel[37:47] if qnt_digitos == 47 else linha_digitavel[10:19]
        valor = int(valor_str) / 100  # Converte centavos para reais
        valor_formatado = f"{valor}".replace('.', ',')
        return valor_formatado

    def extrair_informacoes_boleto(linha_digitavel, qnt_digitos):
        """
        Extrai o valor e a data de vencimento a partir da linha digitável de um boleto.
        """
        # O fator de vencimento está na posição 33 a 36 para boletos de cobrança
        fator_vencimento = int(linha_digitavel[33:37]) if qnt_digitos == 47 else int(linha_digitavel[5:9])
        data_vencimento = calcular_data_vencimento(fator_vencimento)
        # Calcula o valor do boleto
        valor_boleto = calcular_valor_boleto(linha_digitavel, qnt_digitos)
        print(f"\033[1;32mQuantidade de digitos do boleto:{qnt_digitos}\nData de vencimento calculada:{data_vencimento}")
        print(f"Valor calculado:{valor_boleto}\033[0m")
        return valor_boleto, data_vencimento

    qnt_digitos = identificar_digitavel(linha_digitavel)
    valor_boleto, data_vencimento = extrair_informacoes_boleto(linha_digitavel, qnt_digitos)
    return valor_boleto, data_vencimento

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

while True:

    linhas = funcao_input_int("Digite quantas linhas tem o pedido: ")

    if linhas > 0:
        lote = str(input("Digite o lote(Se o lote não for padrão, aperte enter): "))
        validade = str(input("Digite a validade(Se a validade não for padrão, aperte enter): "))
        pa.PAUSE = 0.1
        localizar_clicar("IW_icone.png")  # abrir iw
        pa.press("tab", presses=5)  # clicar na primeira linha
        pa.press("Home")  # ir para a primeira coluna
        pa.hotkey("ctrl", "end")  # descer para a última linha da tela de pedido
        pa.press("right", presses=15)  # ira para coluna validade
        time.sleep(0.5)
        localizar_clicar("NEmbalag.png", click_type='right')  # clicar no campo número de embalagens
        localizar_clicar("OrdCresc.png")  # clicar ordenar crescente
        time.sleep(0.5) if linhas <= 5 else time.sleep(linhas/2)
        pa.hotkey("ctrl", "end")  # descer para a última linha da tela de pedido
        pa.press("right", presses=15)  # ira para coluna validade
        c = 0
        if linhas > 5:
            c = -2
        while True:  # laço para selecionar o fabricante de todos os itens
            if c < linhas:
                localizar_clicar("Alterar_Fabricante.png")  # alterar fabricante
                time.sleep(0.5)
                localizar_clicar("Fabricante.png", click_type='right', confidence=0.9, retries=6)
                localizar_clicar("OrdCresc.png")
                time.sleep(0.2)
                pa.hotkey("tab")
                pa.press("down", presses=randint(3,10))
                time.sleep(0.5)
                localizar_clicar("Selecionar.png")  # selecionar
                time.sleep(0.5)
                pa.PAUSE = 0.1
                pa.hotkey("tab")
                pa.hotkey("tab")
            c = c + 1
            if c >= linhas:
                break
            pa.hotkey("up")


        pa.press("Home")  # ir para a primeira coluna
        pa.hotkey("ctrl", "end")  # descer para a última linha da tela de pedido
        pa.press("right", presses=14)  # ira para coluna validade

        preencher_todos(lote)
        pa.press("tab")
        pa.hotkey("ctrl", "end")
        data_material = "30/12/2030" if validade == "0" else validade
        preencher_todos(data_material)
        pa.press("up")
        time.sleep(0.5)
        localizar_clicar("OK.png")
        pa.hotkey("alt", "tab")  # volta para o pycharm

    dianf = funcao_input_int("Digite a data da NF(Se foi HOJE, digite 0): ") # data emissão da nf
    if dianf > 0:
        mesnf = funcao_input_int("Digite o mês: ")  # mês emissão

    diareceb = funcao_input_int("Digite a data do recebimento: ")  # data recebimento da nf
    if diareceb > 0:
        mesreceb = funcao_input_int("Digite o mês: ")  # mês recebimento

    frete = str(input("Digite o valor do frete: "))  # campo frete
    #   Laço inserção dados do boleto
    while True:
        boleto = str(input("Insira o código de barras do boleto (Se não tiver boleto, aperte Enter): "))  # campo boleto
        if len(str(boleto)) == 0:
            vt = str(input("Digite o valor total da nf: "))  # valor total nf
            prazo = funcao_input_int("Digite a data de vencimento: ")  # dia vencimento da nf
            mesvenc = funcao_input_int("Digite o mês: ")  # mês do vencimento
            if dianf == 0:
                mesnf = int(str(date.today())[5:7])
            datavenc = f"{prazo}/{mesvenc}/2025"
            break

        elif len(str(boleto)) == 44 or len(str(boleto)) == 47:
            vt, datavenc = informacoes_boleto(boleto)
            break

        else:
            print(f"\33[31m"+f"Números inseridos incorretamente. Digitados {len(str(boleto))}"+"\33[0;0m")
            time.sleep(0.3)

    localizar_clicar("IW_icone.png")  # abrir iw
    localizar_clicar("Deposito_Destino.png")# clica deposito destino
    pa.PAUSE = 0.3
    pa.write("F")  # F para farmácia
    pa.moveTo(50, 50)
    localizar_clicar("Deposito_Farmacia.png", verificar_deposito=True, click_type="move")

    pa.PAUSE = 0.2
    pa.click(x=296, y=390)
    time.sleep(0.3)
    pa.press("right")
    pa.hotkey("ctrl", "tab")
    pa.press("tab", presses=3)
    pa.press("home")
    pa.hotkey("ctrl", "a")
    localizar_clicar("Item.png", click_type='right')
    pa.press("up", presses=2)
    pa.press("ENTER")
    pa.press("down", presses=4)
    pa.press("right")
    pa.press("backspace", presses=2)
    pa.write("30")
    pa.press("down")
    pa.hotkey("ctrl", "tab")
    pa.press("ENTER")

    if dianf == 0:
        localizar_clicar("Data_NF.png", click_type='right')

    elif dianf > 0:
        datanf = f"{dianf}/{mesnf}/2024" if mesnf > 10 else f"{dianf}/{mesnf}/2025"
        localizar_clicar("Data_NF.png")
        copiar_colar(datanf)

    if diareceb == 0:
        localizar_clicar("Data_Recebimento.png", click_type='right')

    elif diareceb > 0:
        datareceb = f"{diareceb}/{mesreceb}/2024" if mesreceb > 10 else f"{diareceb}/{mesreceb}/2025"
        localizar_clicar("Data_Recebimento.png")
        copiar_colar(datareceb)

    pa.click(x=58, y=391)
    pa.write("NN")
    pa.click(x=58, y=258)
    copiar_colar(frete)
    pa.click(x=147, y=259)
    copiar_colar(vt)

    if len(str(boleto)) == 0:
        pa.doubleClick(x=191, y=392)
        pa.press("S")
    elif len(str(boleto)) > 0:
        pa.doubleClick(x=191, y=392)
        pa.press("C", presses=3, interval=0.3)
        pa.click(x=482, y=388)
        copiar_colar(boleto)

    if frete in "0":
        pa.doubleClick(x=296, y=390)
        pa.press("N")
    elif frete not in "0":
        pa.doubleClick(x=296, y=390)
        pa.press("S")

    pa.PAUSE = 0.1
    pa.doubleClick(x=855, y=174)
    copiar_colar(datavenc)
    pa.doubleClick(x=951, y=174)
    copiar_colar(vt)
    pa.hotkey("tab")
    localizar_clicar("verifi.nota.png")
    pa.press("tab")
    time.sleep(0.5)
    localizar_clicar("Alerta_Nao.png", teste=True)
    # imprimir = str(input("Imprimir?(Digite N se não) ")).capitalize()
    # if imprimir == "N":
    #     pass
    # elif imprimir != "N":
    #     localizar_clicar("IW_icone.png")
    #     localizar_clicar("Prazo_Validade_Minimo.png", click_type="right")
    #     pa.press("down", presses=5)
    #     pa.press("right")
    #     pa.press("enter")
    #     localizar_clicar("Imprimir.png")
    #     localizar_clicar("Copias.png", click_type='move')
    #     time.sleep(0.3)
    #     pa.press("enter")
    #     time.sleep(0.8)
    #     localizar_clicar("X.png")

    cod_barras_NF = funcao_input_int("Bipe a próxima NF ou digite apenas um número se quiser passar: ")
    while True:
        tamanho_cod_barras_NF = len(str(cod_barras_NF))
        if tamanho_cod_barras_NF == 1:
            break
        if tamanho_cod_barras_NF != 44:
            cod_barras_NF = funcao_input_int(f"Código de barras inválido - você digitou \033[1;34m{tamanho_cod_barras_NF}\33[0;0m/44 números\nbipe novamente: ")

        else:
            localizar_clicar("IW_icone.png")
            localizar_clicar("Nova_NF.png")
            pa.press("up", presses=4)
            pa.press("ENTER")
            localizar_clicar("NFE.png")
            copiar_colar(cod_barras_NF)
            pa.press("tab")
            localizar_clicar("Pedido_Fornec.png")
            pa.press("tab")
            pa.press("right", presses=13)
            localizar_clicar("Preco_Embalagem.png", click_type="right", wait_time=5)
            pa.press("down", presses=8)
            pa.press("ENTER")
            pa.press("tab", presses=17)
            pa.hotkey("ctrl", "end")
            pa.press("left", presses=3)
            break

    print()
    print("-"*90)
    print()

