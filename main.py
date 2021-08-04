import openpyxl as opx

# Bibliotecas pessoais
from manipulaPlanilha.criaPlanilha import *


# Carrega a pasta de trabalho do Forms
pastaTrabalhoForms = input("Informe o nome do arquivo Excel: ")
pastaTrabalhoForms = "EQUIPE TESTE - DAP.xlsx"
pastaTrabalhoForms = opx.load_workbook(pastaTrabalhoForms)
planilhaForms = pastaTrabalhoForms.active

# Inicia a geração da pasta de trabalho com as planilhas organizadas
retorno = criaPlanilhaPessoal(planilhaForms)
novaPastaTrabalho = retorno[0]
novaPlanilha = retorno[1]
novaPastaTrabalho.save("teste.xlsx")