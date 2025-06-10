import xml.etree.ElementTree as ET
import pandas as pd
import re
import os

def extrair_itens_nfe(xml_path, excel_output_path, nome_planilha):
    # Parse do XML
    tree = ET.parse(xml_path)
    root = tree.getroot()

    # Namespace
    ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

    # Lista para armazenar os dados extraídos
    data = []

    # Regex para extrair lote e validade do xProd
    regex_lote = re.compile(r'LOTE[:\s\-]*([A-Z0-9]+)', re.IGNORECASE)
    regex_validade = re.compile(r'VALIDADE\s*[:\.\-]?\s*(\d{2})/(\d{2})/(\d{2,4})', re.IGNORECASE)

    # Percorrer cada item da nota fiscal
    for det in root.findall('.//nfe:det', ns):
        prod = det.find('nfe:prod', ns)
        rastro = prod.find('nfe:rastro', ns)
        xProd = prod.findtext('nfe:xProd', default='', namespaces=ns)
        infAdProd = det.findtext('nfe:infAdProd', default='', namespaces=ns)

        lote = ''
        validade = ''

        # Tentar extrair lote e validade da tag <rastro>
        if rastro is not None:
            lote = rastro.findtext('nfe:nLote', default='', namespaces=ns)
            validade = rastro.findtext('nfe:dVal', default='', namespaces=ns)

        if not lote or not validade:
            match_lote = regex_lote.search(infAdProd)
            if match_lote:
                lote = match_lote.group(1)

            match_validade = regex_validade.search(infAdProd)
            if match_validade:
                dia, mes, ano = match_validade.groups()
                if len(ano) == 2:
                    ano = '20' + ano
                validade = f'{ano}-{mes}-{dia}'

        # Se não encontrou em <rastro>, tentar extrair do xProd
        if not lote:
            lote_match = regex_lote.search(xProd)
            lote = lote_match.group(1) if lote_match else ''

        if not validade:
            validade_match = regex_validade.search(xProd)
            if validade_match:
                dia, mes, ano = validade_match.groups()
                if len(ano) == 2:
                    ano = '20' + ano
                validade = f"{ano}-{mes}-{dia}"

        item = {
            "Produto": xProd,
            "Lote": lote,
            "Validade": validade,
            "Quantidade": float(prod.findtext('nfe:qCom', default='0', namespaces=ns)),
            "Valor Unitário": float(prod.findtext('nfe:vUnCom', default='0', namespaces=ns)),
            "Valor Total": float(prod.findtext('nfe:vProd', default='0', namespaces=ns))
        }
        data.append(item)

    # Criar DataFrame e exportar para Excel
    df = pd.DataFrame(data)
    df.sort_values(by="Valor Total", inplace=True)
    df.to_excel(excel_output_path, index=False)
    # Abrir o arquivo
    try:
        os.startfile(excel_output_path)
    except PermissionError:
        print(f"Não foi possível abrir o arquivo automaticamente: {e}")
    print(f"Planilha gerada: {nome_planilha}")

def montar_caminho_xml(chave, pasta_base="C:/Users/solar/Downloads"):
    caminho = os.path.join(pasta_base, f"{chave}.xml")
    return caminho

# Exemplo de uso
if __name__ == "__main__":
    chave_nfe = input("Bipe a NF: ").strip()
    xml_path = montar_caminho_xml(chave_nfe)

    print("Caminho completo: ", xml_path)

    desktop = os.path.join(os.path.expanduser("~"), "OneDrive", "Área de Trabalho")
    planilha_padrao = "NFe_Planilha_Padrão"
    xml_entrada = f"C:/Users/solar/Downloads/{chave_nfe}.xml"  # Substitua com o caminho do seu arquivo XML
    excel_saida = os.path.join(desktop, f"{planilha_padrao}.xlsx")
    extrair_itens_nfe(xml_path, excel_saida, planilha_padrao)
