
def geraAlfabeto():

    """
    Retorna uma lista com as letras do alfabeto
    """

    alfabeto = [
        'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
        'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'
    ]

    return alfabeto


def contaUsuarios(planilha):

    """
    Recebe como parâmetro uma planilha do Forms e idêntifica quantas pessoas responderam o mesmo.
    """

    contadores = {
        'nCelula': 2,
        'quantUsuarios': 0
    }

    while(True):

        if contadores['quantUsuarios'] == 0:
            celula = "A2"

        if planilha[celula].value == None:
            break

        else:
            contadores['nCelula'] += 1
            contadores['quantUsuarios'] += 1
            celula = f"A{contadores['nCelula']}"


    return contadores['quantUsuarios']


def geraDicionarios(tipoPlanilha):

    """
    Gera os dicionários que irão armazenar as informações
    """

    if tipoPlanilha == "Informações Pessoais":
        # Chaves que serão usadas no dicionário abaixo
        chaves = [
            'Nome',
            'Email',
            'Cargo',
            'Matricula',
            'Local de Atuacao',
            'Nome Equipe',
            'Gestor'
        ]

        # Gera um dicionário vazio com uma lista para cada chave.
        # Essas listas serão usadas posteriormente
        dicionario = {}

        for chave in chaves:
            dicionario[chave] = []

        retorno = [dicionario, chaves]


    return retorno


def geraCelulas(intervaloCelulas, quantUsuarios, **kws):

    """
    Gera as células que terão o conteúdo lido
    """

    alfabeto = geraAlfabeto()

    # Armazena o intervalo de colunas no dicionário "intervalo" para ser usado posteriormente
    temp = intervaloCelulas.split(":")
    intervalo = {
        "Inicio": temp[0],
        "Fim": temp[1],
        "tamIntervalo": int(temp[2])
    }


    # Cria uma lista com o intervalo passado pelo usuário
    alfabetoPersonalizado = []

    pegaLetra = False
    for letra in alfabeto:

        if letra == intervalo["Inicio"]:
            pegaLetra = True

        if pegaLetra == True:
            alfabetoPersonalizado.append(letra)

        if letra == intervalo["Fim"]:
            break


    # Cria uma lista com as células que serão lidas
    linhaInicial = kws.get("linhaInicial")
    if linhaInicial == None: linhaInicial = 2

    celulas = []
    contadores = {
        "i": 0,
        "j": 0,
        "linhaAtual": linhaInicial
    }

    while(contadores["i"] < intervalo["tamIntervalo"]):

        # Gera uma lista com um grupo pequeno de células que serão adicionadas ao grupo principal
        while(contadores["j"] < quantUsuarios):

            if contadores["j"] == 0:
                temp = []

            colunaAtual = alfabetoPersonalizado[ contadores["i"] ]
            linhaAtual = contadores["linhaAtual"]
            celula = f"{colunaAtual}{linhaAtual}"

            temp.append(celula)
            contadores["j"] += 1
            contadores["linhaAtual"] += 1


        # Armazena aa célula gerada na lista geral de células
        celulas.append(temp)
        contadores["j"] = 0
        contadores["linhaAtual"] = linhaInicial

        contadores["i"] += 1


    return celulas


def retornaInformacoes(planilhaForms, intervalo):

    """
    Utiliza o intervalo informado por parâmetro para buscar as informações
    na planilha informada.
    """

    # Armazena numa variável a quantidade de usuários que responderam ao Forms
    quantUsuarios = contaUsuarios(planilhaForms)

    # Gera um dicionário com listas vazias
    dicionarioChaves = geraDicionarios("Informações Pessoais")
    dicionario = dicionarioChaves[0]
    chaves = dicionarioChaves[1]

    # Gera uma lista com as células que serão varridas
    celulas = geraCelulas(intervalo, quantUsuarios)

    # Faz uma varredura nas células selecionadas e armazena as informações no dicionário
    indiceChave = 0
    for grupo in celulas:

        for celula in grupo:
            temp = chaves[indiceChave]
            dicionario[temp].append(planilhaForms[celula].value)

        indiceChave += 1


    return dicionario
