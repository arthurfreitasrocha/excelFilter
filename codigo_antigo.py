from pandas import read_excel, DataFrame, ExcelWriter
from time import sleep


def contador_caracteres(sequencia_coluna):
    
    """
    Essa função avança em 1 casa a coluna do Excel passada como parâmetro
    """

    # Faz a validação dos parâmetros
    try:
        if sequencia_coluna.isalpha() == False:
            return "ERRO [01]: Você precisa passar como parâmetro uma sequência de LETRAS"

        # Informações essenciais
        tam_sequencia = len(sequencia_coluna)

    except:
        print("ERRO [02]: Você precisa passar como parâmetro uma sequência de caracteres")



    # Vetor usado para armazenar o alfabeto
    vetor_alfab = [
        'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
        'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'
    ]


    # Se a sequência for composta por apenas 1 letra, entra no if abaixo
    if tam_sequencia == 1:

        # Atualiza o valor da coluna para o próximo valor
        if sequencia_coluna != 'Z': sequencia_coluna = vetor_alfab[ vetor_alfab.index(sequencia_coluna)+1 ]

        # Atualiza o valor da coluna para o próximo valor
        else: sequencia_coluna = 'AA'


    # Se a sequência for composta por 2 letras ou mais, entra no elif abaixo
    elif tam_sequencia > 1:

        sequencia_separada = []
        for letra in sequencia_coluna: sequencia_separada.append(letra)


        if sequencia_coluna[-1] != 'Z':
            # Atualiza o valor da coluna para o próximo valor
            sequencia_separada[-1] = vetor_alfab[ vetor_alfab.index(sequencia_coluna[-1])+1 ]

        else:
            # Atualiza o valor da coluna para o próximo valor
            sequencia_separada[-1] = 'A'

            # Atualiza o valor da coluna para o próximo valor
            sequencia_separada[-2] = vetor_alfab[ vetor_alfab.index(sequencia_coluna[-2])+1 ]


        sequencia_coluna = ''
        for letra in sequencia_separada: sequencia_coluna += letra


    return sequencia_coluna



def info_pessoal(nome_documento):

    """
    Captura as informações pessoais de cada usuário
    """

    # Leitura dos dados
    info_pessoal = read_excel(nome_documento, usecols="B:H")
    dictionary = {}

    for index,row in info_pessoal.iterrows():
        dictionary[row[0]]={row[1]:[]}
        dictionary[row[0]][row[1]].append(row[2])
        dictionary[row[0]][row[1]].append(row[3])
        dictionary[row[0]][row[1]].append(row[4])
        dictionary[row[0]][row[1]].append(row[5])
        dictionary[row[0]][row[1]].append(row[6])


    # Filtragem dos dados
    nome = []
    email = []
    cargo = []
    matricula = []
    local_atuacao = []
    nome_equipe = []
    gestor = []

    dados_formatados = {}

    i = 0
    for item in dictionary:
        # Captura o nome do usuário
        nome.append( item )

        # Captura o email do usuário que também é a chave para o sub dicionario
        temp = dictionary[item]
        for info in temp:
            email.append( info )

        # Armazena numa lista as informações do sub dicionairo
        temp = dictionary[ item ]
        temp = temp[ email[i] ]

        # Filtra essas informações em listas separadas
        cargo.append( temp[0] )
        matricula.append( temp[1] )
        local_atuacao.append( temp[2] )
        nome_equipe.append( temp[3] )
        gestor.append( temp[4] )

        i += 1


    # Armazena os dados formatados no dicionário abaixo
    dados_formatados['Nome'] = nome
    dados_formatados['Email'] = email
    dados_formatados['Cargo'] = cargo
    dados_formatados['Matrícula'] = matricula
    dados_formatados['Local de Atuação'] = local_atuacao
    dados_formatados['Nome da Equipe'] = nome_equipe
    dados_formatados['Gestor'] = gestor

    return dados_formatados



def info_auto_avaliacao(nome_documento):

    """
    Captura as informações de auto avaliação
    """

    # Leitura dos dados
    info_auto_avaliacao = read_excel(nome_documento, usecols="B,I:W")
    dictionary = {}

    for index,row in info_auto_avaliacao.iterrows():
        dictionary[row[0]]={row[1]:[]}
        dictionary[row[0]][row[1]].append(row[2])
        dictionary[row[0]][row[1]].append(row[3])
        dictionary[row[0]][row[1]].append(row[4])
        dictionary[row[0]][row[1]].append(row[5])
        dictionary[row[0]][row[1]].append(row[6])
        dictionary[row[0]][row[1]].append(row[7])
        dictionary[row[0]][row[1]].append(row[8])
        dictionary[row[0]][row[1]].append(row[9])
        dictionary[row[0]][row[1]].append(row[10])
        dictionary[row[0]][row[1]].append(row[11])
        dictionary[row[0]][row[1]].append(row[12])
        dictionary[row[0]][row[1]].append(row[13])
        dictionary[row[0]][row[1]].append(row[14])
        dictionary[row[0]][row[1]].append(row[15])


    # Filtragem dos dados
    nome = []
    participacao_planejamento = []
    clareza_metas = []
    iniciativa = []
    criatividade = []
    eficiencia = []
    eficacia = []
    assiduidade = []
    compromisso = []
    zelo_material = []
    conduta = []
    espirito_equipe = []
    responsabilidade = []
    comunicacao = []
    auto_desenvolvimento = []
    competencia_tecnica = []

    dados_formatados = {}

    i = 0
    for item in dictionary:
        # Captura o nome do usuário
        nome.append( item )

        # Captura o email do usuário que também é a chave para o sub dicionario
        temp = dictionary[item]
        for info in temp:
            participacao_planejamento.append( info )

        # Armazena numa lista as informações do sub dicionairo
        temp = dictionary[ item ]
        temp = temp[ participacao_planejamento[i] ]

        # Filtra essas informações em listas separadas
        clareza_metas.append( temp[0] )
        iniciativa.append( temp[1] )
        criatividade.append( temp[2] )
        eficiencia.append( temp[3] )
        eficacia.append( temp[4] )
        assiduidade.append( temp[5] )
        compromisso.append( temp[6] )
        zelo_material.append( temp[7] )
        conduta.append( temp[8] )
        espirito_equipe.append( temp[9] )
        responsabilidade.append( temp[10] )
        comunicacao.append( temp[11] )
        auto_desenvolvimento.append( temp[12] )
        competencia_tecnica.append( temp[13] )

        i += 1


    # Armazena os dados formatados no dicionário abaixo
    dados_formatados['Nome'] = nome
    dados_formatados['Participação no Planejamento'] = participacao_planejamento
    dados_formatados['Clareza das Metas'] = clareza_metas
    dados_formatados['Iniciativa'] = iniciativa
    dados_formatados['Criatividade'] = criatividade
    dados_formatados['Eficiência'] = eficiencia
    dados_formatados['Eficacia'] = eficacia
    dados_formatados['Assiduidade'] = assiduidade
    dados_formatados['Compromisso'] = compromisso
    dados_formatados['Zelo de Material'] = zelo_material
    dados_formatados['Conduta'] = conduta
    dados_formatados['Espírito de Equipe'] = espirito_equipe
    dados_formatados['Responsabilidade'] = responsabilidade
    dados_formatados['Comunicação'] = comunicacao
    dados_formatados['Auto Desenvolvimento'] = auto_desenvolvimento
    dados_formatados['Competência Ténica'] = competencia_tecnica

    return dados_formatados



def info_equipe(nome_documento):

    """
    Captura as informações da avaliacao da equipe
    """

    # Leitura dos dados
    info_equipe = read_excel(nome_documento, usecols="B,X:AL")
    dictionary = {}

    for index,row in info_equipe.iterrows():
        dictionary[row[0]]={row[1]:[]}
        dictionary[row[0]][row[1]].append(row[2])
        dictionary[row[0]][row[1]].append(row[3])
        dictionary[row[0]][row[1]].append(row[4])
        dictionary[row[0]][row[1]].append(row[5])
        dictionary[row[0]][row[1]].append(row[6])
        dictionary[row[0]][row[1]].append(row[7])
        dictionary[row[0]][row[1]].append(row[8])
        dictionary[row[0]][row[1]].append(row[9])
        dictionary[row[0]][row[1]].append(row[10])
        dictionary[row[0]][row[1]].append(row[11])
        dictionary[row[0]][row[1]].append(row[12])
        dictionary[row[0]][row[1]].append(row[13])
        dictionary[row[0]][row[1]].append(row[14])
        dictionary[row[0]][row[1]].append(row[15])


    # Filtragem dos dados
    nome = []
    participacao_planejamento = []
    clareza_metas = []
    iniciativa = []
    criatividade = []
    eficiencia = []
    eficacia = []
    assiduidade = []
    compromisso = []
    zelo_material = []
    conduta = []
    espirito_equipe = []
    responsabilidade = []
    comunicacao = []
    auto_desenvolvimento = []
    competencia_tecnica = []

    dados_formatados = {}

    i = 0
    for item in dictionary:
        # Captura o nome do usuário
        nome.append( item )

        # Captura o email do usuário que também é a chave para o sub dicionario
        temp = dictionary[item]
        for info in temp:
            participacao_planejamento.append( info )

        # Armazena numa lista as informações do sub dicionairo
        temp = dictionary[ item ]
        temp = temp[ participacao_planejamento[i] ]

        # Filtra essas informações em listas separadas
        clareza_metas.append( temp[0] )
        iniciativa.append( temp[1] )
        criatividade.append( temp[2] )
        eficiencia.append( temp[3] )
        eficacia.append( temp[4] )
        assiduidade.append( temp[5] )
        compromisso.append( temp[6] )
        zelo_material.append( temp[7] )
        conduta.append( temp[8] )
        espirito_equipe.append( temp[9] )
        responsabilidade.append( temp[10] )
        comunicacao.append( temp[11] )
        auto_desenvolvimento.append( temp[12] )
        competencia_tecnica.append( temp[13] )

        i += 1


    # Armazena os dados formatados no dicionário abaixo
    dados_formatados['Nome'] = nome
    dados_formatados['Participação no Planejamento'] = participacao_planejamento
    dados_formatados['Clareza das Metas'] = clareza_metas
    dados_formatados['Iniciativa'] = iniciativa
    dados_formatados['Criatividade'] = criatividade
    dados_formatados['Eficiência'] = eficiencia
    dados_formatados['Eficacia'] = eficacia
    dados_formatados['Assiduidade'] = assiduidade
    dados_formatados['Compromisso'] = compromisso
    dados_formatados['Zelo de Material'] = zelo_material
    dados_formatados['Conduta'] = conduta
    dados_formatados['Espírito de Equipe'] = espirito_equipe
    dados_formatados['Responsabilidade'] = responsabilidade
    dados_formatados['Comunicação'] = comunicacao
    dados_formatados['Auto Desenvolvimento'] = auto_desenvolvimento
    dados_formatados['Competência Ténica'] = competencia_tecnica

    return dados_formatados



def info_avaliacao_individual(nome_documento, sequencia):

    """
    Captura as informações de avaliacao individual
    """

    # Leitura dos dados
    info_avaliacao_individual = read_excel(nome_documento, usecols=f'B,{sequencia}')
    dictionary = {}

    for index,row in info_avaliacao_individual.iterrows():
        dictionary[row[0]]={row[1]:[]}
        dictionary[row[0]][row[1]].append(row[2])
        dictionary[row[0]][row[1]].append(row[3])
        dictionary[row[0]][row[1]].append(row[4])
        dictionary[row[0]][row[1]].append(row[5])
        dictionary[row[0]][row[1]].append(row[6])
        dictionary[row[0]][row[1]].append(row[7])
        dictionary[row[0]][row[1]].append(row[8])
        dictionary[row[0]][row[1]].append(row[9])
        dictionary[row[0]][row[1]].append(row[10])
        dictionary[row[0]][row[1]].append(row[11])
        dictionary[row[0]][row[1]].append(row[12])
        dictionary[row[0]][row[1]].append(row[13])
        dictionary[row[0]][row[1]].append(row[14])
        dictionary[row[0]][row[1]].append(row[15])
        dictionary[row[0]][row[1]].append(row[16])


    # Filtragem dos dados
    nome = []
    pessoa_avaliada = []
    participacao_planejamento = []
    clareza_metas = []
    iniciativa = []
    criatividade = []
    eficiencia = []
    eficacia = []
    assiduidade = []
    compromisso = []
    zelo_material = []
    conduta = []
    espirito_equipe = []
    responsabilidade = []
    comunicacao = []
    auto_desenvolvimento = []
    competencia_tecnica = []

    dados_formatados = {}

    i = 0
    for item in dictionary:
        # Captura o nome do usuário
        nome.append( item )

        # Captura o email do usuário que também é a chave para o sub dicionario
        temp = dictionary[item]
        for info in temp:
            pessoa_avaliada.append( info )

        # Armazena numa lista as informações do sub dicionairo
        temp = dictionary[ item ]
        temp = temp[ pessoa_avaliada[i] ]

        # Filtra essas informações em listas separadas
        participacao_planejamento.append( temp[0] )
        clareza_metas.append( temp[1] )
        iniciativa.append( temp[2] )
        criatividade.append( temp[3] )
        eficiencia.append( temp[4] )
        eficacia.append( temp[5] )
        assiduidade.append( temp[6] )
        compromisso.append( temp[7] )
        zelo_material.append( temp[8] )
        conduta.append( temp[9] )
        espirito_equipe.append( temp[10] )
        responsabilidade.append( temp[11] )
        comunicacao.append( temp[12] )
        auto_desenvolvimento.append( temp[13] )
        competencia_tecnica.append( temp[14] )

        i += 1


    # Armazena os dados formatados no dicionário abaixo
    dados_formatados['Nome'] = nome
    dados_formatados['Pessoa Avaliada'] = pessoa_avaliada
    dados_formatados['Participação no Planejamento'] = participacao_planejamento
    dados_formatados['Clareza das Metas'] = clareza_metas
    dados_formatados['Iniciativa'] = iniciativa
    dados_formatados['Criatividade'] = criatividade
    dados_formatados['Eficiência'] = eficiencia
    dados_formatados['Eficacia'] = eficacia
    dados_formatados['Assiduidade'] = assiduidade
    dados_formatados['Compromisso'] = compromisso
    dados_formatados['Zelo de Material'] = zelo_material
    dados_formatados['Conduta'] = conduta
    dados_formatados['Espírito de Equipe'] = espirito_equipe
    dados_formatados['Responsabilidade'] = responsabilidade
    dados_formatados['Comunicação'] = comunicacao
    dados_formatados['Auto Desenvolvimento'] = auto_desenvolvimento
    dados_formatados['Competência Ténica'] = competencia_tecnica

    return dados_formatados




def info_tela(mensagem):

    i = 0
    while(i < 50):
        print('')
        i += 1

    print('\n===== MENSAGEM DO SISTEMA =====\n')
    print(mensagem)

    sleep(0.5)


i = 0
while(i < 50):
    print('')
    i += 1
print('\n===== MENSAGEM DO SISTEMA =====\n')
nome_arquivo = str(input("\nInforme o nome completo do arquivo: "))

info_p = None
linhas_arq = None
achou_arquivo = False

try:
    info_p = read_excel(nome_arquivo, usecols="A")
    info_p = str(info_p)
    linhas_arq = len(info_p.split('\n'))-1
    achou_arquivo = True

except:
    info_tela("ERRO: Não foi encontrado nenhum arquivo Excel com esse nome")



if achou_arquivo == True:

    # Retorna as informações pessoais dos usuários
    dict_info_pessoal = info_pessoal(nome_arquivo)
    info_tela("Lendo planilha - 25%")

    # Cria o DataFrame para ser armazenado no Excel
    dtf1 = DataFrame(dict_info_pessoal)


    # Retorna a nota da auto avaliacao
    dict_auto_avaliacao = info_auto_avaliacao(nome_arquivo)
    info_tela("Lendo planilha - 50%")

    # Cria o DataFrame para ser armazenado no Excel
    dtf2 = DataFrame(dict_auto_avaliacao)


    # Retorna a nota da equipe
    dict_equipe = info_equipe(nome_arquivo)
    info_tela("Lendo planilha - 75%")

    # Cria o DataFrame para ser armazenado no Excel
    dtf3 = DataFrame(dict_equipe)


    # Retorna a nota da avaliacao individual
    inicio_intervalo = fim_intervalo = 'AM'
    sequencia = ''
    list_avaliacoes = []

    i = 0
    while(i < linhas_arq):

        j = 0
        while(j <= 15):
            fim_intervalo = contador_caracteres(fim_intervalo)
            j += 1

        sequencia = f'B,{inicio_intervalo}:{fim_intervalo}'
        list_avaliacoes.append( info_avaliacao_individual(nome_arquivo, sequencia) )
        info_tela("Lendo planilha - 99%")

        inicio_intervalo = fim_intervalo
        i += 1

    info_tela("Lendo planilha - 100%")


    # Cria a pasta de trabalho do Excel com as planilhas
    nome_novo_arquivo = f'[ESTILIZADO] {nome_arquivo}'
    info_tela("Gerando nova planilha - 50%")

    with ExcelWriter(nome_novo_arquivo) as writer:  
        dtf1.to_excel(writer, index=False, sheet_name='Informações Pessoais')
        dtf2.to_excel(writer, index=False, sheet_name='Auto Avaliação')
        dtf3.to_excel(writer, index=False, sheet_name='Avaliação da Equipe')

        i = 0
        planilhas_individuais = []
        while(i < len(list_avaliacoes)):

            # Cria o nome da planilha
            nome_planilha = f'Avaliação Individual {i+1}'


            # Gera o DataFrame da planilha
            temp = list_avaliacoes[i]

            dtf4 = DataFrame(temp)
            dtf4.to_excel(writer, index=False, sheet_name=nome_planilha)
            info_tela("Gerando nova planilha - 99%")

            i += 1
            planilhas_individuais.append(dtf4)


    # Chama a função que vai criar a planilha com as médias
    planilha_final(dtf1, dtf2, dtf3, planilhas_individuais)

    info_tela("Gerando nova planilha - 100%")

    info_tela(f"Nova planilha com dados filtrados gerada com sucesso!\n\nO nome da planilha é: {nome_novo_arquivo}")


