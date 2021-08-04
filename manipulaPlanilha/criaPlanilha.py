from openpyxl.workbook.workbook import Workbook

# Bibliiotecas pessoais
from manipulaPlanilha.utilidades import retornaInformacoes, contaUsuarios, geraCelulas


def criaPlanilhaPessoal(planilhaForms):

    """
    Cria uma planilha com as informações pessoais dos usuários
    """

    # Captura as informações citadas no intervalo passado pelo segundo parâmetro
    informacoesUsuarios = retornaInformacoes(planilhaForms, "B:H:7")


    # Cria uma nova pasta de trabalho
    novaPastaTrabalho = Workbook()
    novaPlanilha = novaPastaTrabalho.active
    novaPlanilha.title = "Informações Pessoais"



    # Usa o laço de repetição para destrinchar o dicionário e inserir
    # todas as informações armazenadas na nova planilha
    quantUsuarios = contaUsuarios(planilhaForms) + 1
    celulas = geraCelulas("A:G:7", quantUsuarios, linhaInicial=1)

    contadores = {
        'contadorCelula': 0,
        'contadorLaco': 1
    }

    for grupo in informacoesUsuarios:

        # Escreve o nome do grupo na primeira linha de cada coluna
        temp = celulas[ contadores["contadorCelula"] ]
        novaPlanilha[ temp[0] ] = grupo

        # Escreve as outras informações nas linhas remanescentes
        for informacao in informacoesUsuarios[grupo]:
            novaPlanilha[ temp[ contadores["contadorLaco"] ] ] = informacao
            contadores["contadorLaco"] += 1

        contadores["contadorLaco"] = 1
        contadores["contadorCelula"] += 1


    retorno = [novaPastaTrabalho, novaPlanilha]
    return retorno

