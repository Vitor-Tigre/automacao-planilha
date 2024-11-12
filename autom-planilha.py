import re 
from openpyxl import Workbook

# INSIRA AQUI OS CAMINHOS DOS ARQUIVOS .TXT E EXCEL!!
caminho_txt = 'documento-teste.txt'
caminho_excel = 'planilha-teste.xlsx'


# função para ler dados do arquivo .txt
def ler_dados_do_txt(caminho_txt):
    with open(caminho_txt, 'r', encoding='UTF-8') as arquivo:
        conteudo = arquivo.read()

    # expressões para capturar as informações
    padrao = r"Responsável:\s*(.*?)\nTask:\s*(.*?)\nDescrição:\s*(.*?)\nObservações:\s*(.*?)\n"
    dados = re.findall(padrao, conteudo, re.DOTALL)

    return dados

def criar_planilha(dados, caminho_excel):
    # criando uma nova planilha
    wb = Workbook()
    ws = wb.active

    # nomes das colunas
    ws.append(["Responsável", "Task", "Descrição", "Observações"])

    # inserindo os dados na planilha
    for dado in dados:
        ws.append(dado)

    # salvando no seu arquivo Excel ( que você inseriu na linha 6 )
    wb.save(caminho_excel)

# lendo os dados do .txt
dados = ler_dados_do_txt(caminho_txt)

print(dados)

if dados:
    criar_planilha(dados, caminho_excel)
    print(f'a planilha {caminho_excel} foi criada com sucesso!')
else:
    print(f"nenhum dado encontrado no arquivo {caminho_txt}")

