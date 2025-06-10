Descrição
Este repositório contém um conjunto de scripts Python para automação de processos relacionados ao processamento de notas fiscais, incluindo extração de dados de XMLs, interação com sistemas ERP e manipulação de planilhas.

Scripts Disponíveis
1. extrair_itens_nfe.py
Função: Extrai informações de itens de notas fiscais eletrônicas (NF-e) a partir de arquivos XML e gera uma planilha Excel padronizada.

Recursos:

Extrai produto, lote, validade, quantidade e valores

Busca informações em tags XML e campos de texto

Formata dados para planilha Excel

Abre automaticamente o arquivo gerado

2. extrair_pedido_iw.py
Função: Automatiza a extração de dados de pedidos de um sistema ERP (IW).

Recursos:

Navegação automatizada via interface gráfica (pyautogui)

Exportação de dados para CSV

Filtragem e formatação de colunas específicas

Geração de arquivo Excel com informações relevantes

3. fabricante.py
Função: Automatiza o preenchimento de informações de fabricante, lote e validade em pedidos do sistema ERP.

Recursos:

Preenchimento em massa de campos

Cálculo automático de datas de vencimento

Processamento de informações de boletos

Integração com interface gráfica do sistema

4. inserir_lotes_e_validades_IW.py
Função: Insere informações de lotes e validades no sistema ERP a partir de uma planilha padrão.

Recursos:

Leitura de planilha Excel

Ajuste automático de formatos de data

Inserção sequencial de dados no sistema

Tratamento de erros e validações

Pré-requisitos
Python 3.x

Bibliotecas necessárias:

xml.etree.ElementTree

pandas

pyautogui

pyscreeze

pyperclip

os

re

time

datetime

Configuração
Clone o repositório

Instale as dependências: pip install -r requirements.txt

Configure os caminhos das imagens no arquivo IMAGENS_DIR conforme sua estrutura de pastas

Ajuste os caminhos de arquivos conforme necessário (Downloads, Desktop, etc.)

Uso
Cada script pode ser executado individualmente. Alguns possuem interfaces de entrada via prompt de comando para inserção de dados como número de NF, lotes e validades.

Observações
Os scripts foram desenvolvidos para um ambiente específico e podem requerer ajustes para funcionar em outros sistemas

Algumas funcionalidades dependem da estrutura de interface gráfica do sistema alvo

Recomenda-se testar em ambiente controlado antes de uso em produção

Contribuição
Contribuições são bem-vindas. Sinta-se à vontade para abrir issues ou pull requests com melhorias e correções.
