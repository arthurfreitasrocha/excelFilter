from openpyxl import load_workbook

# Bibliotecas pessoais
from manipulaPastaTrabalho.manipulaPlanilha import Celulas, Planilha


# Carrega a pasta de trabalho do Forms
pastaTrabalhoForms = input("Informe o nome do arquivo Excel: ")
pastaTrabalhoForms = "EQUIPE TESTE - DAP.xlsx"
pastaTrabalhoForms = load_workbook(pastaTrabalhoForms)
planilhaForms = pastaTrabalhoForms.active


nomePlanilha = ["Informações Pessoais", "Auto Avaliação", "Avaliação da Equipe"]
intervaloLeitura = ["B:H", "B:I:W", "B:X:AL"]
intervaloEscrita = ["A:G", "A:P"]
sequenciaChaves = [
    "Nome:Email:Cargo:Matrícula:Local de Atuação:Nome Equipe:Gestor",
    "Nome:Participação no Planejamento:Clareza das Metas:Iniciativa:Criatividade:Eficiência:Eficácia:Assiduidade:Compromisso:Zelo com Material:Conduta:Espírito de Equipe:Responsabilidade:Comunicação:Auto Desenvolvimento:Competência Técnica"]

obj = Planilha(planilhaForms)
obj.criaPlanilhaDependente(nomePlanilha[0], intervaloLeitura[0], intervaloEscrita[0], sequenciaChaves[0])
obj.criaPlanilhaDependente(nomePlanilha[1], intervaloLeitura[1], intervaloEscrita[1], sequenciaChaves[1])
obj.criaPlanilhaDependente(nomePlanilha[2], intervaloLeitura[2], intervaloEscrita[1], sequenciaChaves[1])


# Cria as planilhas de avaliação individual
nomePlanilha = "Avaliação Individual"
intervaloLeitura = "B:AM:BB"
inicioIntervalo = intervaloLeitura[2] + intervaloLeitura[3]
fimIntervalo = intervaloLeitura[5] + intervaloLeitura[6]
intervaloEscrita = "A:Q"
sequenciaChaves = "Nome:Nome Avaliado:Participação no Planejamento:Clareza das Metas:Iniciativa:Criatividade:Eficiência:Eficácia:Assiduidade:Compromisso:Zelo com Material:Conduta:Espírito de Equipe:Responsabilidade:Comunicação:Auto Desenvolvimento:Competência Técnica"
somaCelula = Celulas(None, None, None)

repeticoes = 3
tamIntervalo = 16

for contaRepeticoes in range(repeticoes):
    novaPastaTrabalho = obj.criaPlanilhaDependente(nomePlanilha, intervaloLeitura, intervaloEscrita, sequenciaChaves)

    for contaTamIntervalo in range(tamIntervalo):
        inicioIntervalo = somaCelula.getSomaColuna(inicioIntervalo)
        fimIntervalo = somaCelula.getSomaColuna(fimIntervalo)

    intervaloLeitura = f"B:{inicioIntervalo}:{fimIntervalo}"



# Cria a planilha de avaliação geral
manipulaCelulas = Celulas(planilhaForms, "B", "Nome")

nomesUsuarios = manipulaCelulas.getInformacoes()
nomesUsuarios = nomesUsuarios["Nome"]

temp = nomesUsuarios[1].split(" ")
temp = temp[0]

temp2 = nomesUsuarios[2].split(" ")
temp2 = temp2[0]

temp3 = nomesUsuarios[3].split(" ")
temp3 = temp3[0]

nomesUsuarios = [temp, temp2, temp3]
dados = {
	"Perguntas": [
			"Perguntas",
            "", 
			"01 - PARTICIPAÇÃO OU CONHECIMENTO DO PLANEJAMENTO",
			"02 - CLAREZA DAS METAS/TAREFAS",
			"03 - INICIATIVA",
			"04 - CRIATIVIDADE",
			"05 - EFICIÊNCIA",
			"06 - EFICÁCIA",
			"07 - ASSIDUIDADE",
			"08 - COMPROMISSO",
			"09 - ZELO COM MATERIAIS E EQUIPAMENTOS",
			"10 - CONDUTA DISCIPLINAR",
			"11 - ESPÍRITO DE EQUIPE",
			"12 - RESPONSABILIDADE COM INFORMAÇÕES",
			"13 - COMUNICAÇÃO",
			"14 - AUTO-DESENVOLVIMENTO",
			"15 - COMPETÊNCIA TÉCNICA"],
	"|-": ["", "I", 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10],
	" -": ["", "II", 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9],
	"Formulários": ["Formulários", "III", 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8],
	"- ": ["", "IV", 7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7],
	"-|": ["", "V", 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6]
}
intervaloEscrita = "A:F"
sequenciaChaves = "Perguntas:|-: -:Formulários:- :-|"

for contaRepeticoes in range(repeticoes):

    nomePlanilha = f"Avaliação Geral - {nomesUsuarios[contaRepeticoes]}"

    if contaRepeticoes != repeticoes:
        obj.criaPlanilhaIndependente(nomePlanilha, dados, intervaloEscrita, sequenciaChaves, 17)
    else:
        novaPastaTrabalho = obj.criaPlanilhaIndependente(nomePlanilha, dados, intervaloEscrita, sequenciaChaves, 17)

novaPastaTrabalho.save("teste.xlsx")