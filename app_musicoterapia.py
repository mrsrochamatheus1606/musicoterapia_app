import os
import math
from io import BytesIO

import streamlit as st
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH


st.set_page_config(page_title="Avaliação Musicoterapêutica", layout="wide")
st.title("🎵 Avaliação Musicoterapêutica e Relatório Clínico Avançado")


st.header("Dados do Profissional")

col1, col2 = st.columns(2)
with col1:
    terapeuta = st.text_input("Nome completo do terapeuta")
with col2:
    registro = st.text_input("Registro profissional")


st.header("Identificação do Paciente")

col1, col2, col3 = st.columns(3)
with col1:
    nome = st.text_input("Nome do paciente")
with col2:
    idade = st.number_input("Idade", min_value=0, max_value=120, value=0, step=1)
with col3:
    diagnostico = st.text_input("Diagnóstico")

data_nascimento = st.text_input("Data de nascimento")
escolaridade = st.text_input("Escolaridade")
responsaveis = st.text_input("Nome dos pais ou responsáveis")

metodo_intensivo = st.selectbox(
    "Métodos Intensivos / Abordagem Terapêutica",
    ["", "MIG", "TREINI", "ABA", "Particular"]
)

data_inicio = st.text_input("Data de início da intervenção")
data_atual = st.text_input("Data atual")
historia_clinica = st.text_area("História clínica")
queixa = st.text_area("Queixa principal / motivo do encaminhamento")
observacoes = st.text_area("Observações clínicas")
preferencias = st.text_area("Preferências e recusas sonoro-musicais")
interacoes = st.text_area("Interações sonoro-musicais")


st.header("Escalas a utilizar no relatório")

escalas_escolhidas = st.multiselect(
    "Selecione as avaliações que deseja aplicar",
    ["Nordoff-Robbins", "IAPS", "DEMUCA"],
    default=["Nordoff-Robbins", "IAPS"]
)


def campo(label, key):
    return st.number_input(label, min_value=0, max_value=5, value=0, step=1, key=key)


nordoff_dominios = {}
iaps_improvisacao = {}
iaps_recriacao = {}
iaps_composicao = {}
iaps_escuta = {}

if "Nordoff-Robbins" in escalas_escolhidas:
    st.header("Escala de Comunicabilidade Musical Nordoff-Robbins")
    st.caption("Pontuação de 0 a 5 para cada domínio.")

    nordoff_dominios = {
        "Expressão emocional": campo("Expressão emocional", "n_exp_emocional"),
        "Exploração sonora": campo("Exploração sonora", "n_exploracao_sonora"),
        "Interação musical": campo("Interação musical", "n_interacao_musical"),
        "Engajamento": campo("Engajamento", "n_engajamento"),
        "Responsividade musical": campo("Responsividade musical", "n_responsividade"),
        "Iniciativa musical": campo("Iniciativa musical", "n_iniciativa"),
        "Sustentação da atividade musical": campo("Sustentação da atividade musical", "n_sustentacao"),
        "Comunicação não verbal": campo("Comunicação não verbal", "n_comunicacao_nao_verbal"),
        "Reciprocidade musical": campo("Reciprocidade musical", "n_reciprocidade"),
        "Organização musical": campo("Organização musical", "n_organizacao"),
    }


if "IAPS" in escalas_escolhidas:
    st.header("IAPS")

    st.subheader("IAPS - Improvisação")
    iaps_improvisacao = {
        "Iniciativa sonora": campo("Iniciativa sonora", "i_iniciativa"),
        "Resposta musical": campo("Resposta musical", "i_resposta"),
        "Organização sonora": campo("Organização sonora", "i_organizacao"),
        "Interação musical": campo("Interação musical", "i_interacao"),
    }

    st.subheader("IAPS - Recriação")
    iaps_recriacao = {
        "Memória musical": campo("Memória musical", "r_memoria"),
        "Coordenação motora": campo("Coordenação motora", "r_coordenacao"),
        "Seguimento musical": campo("Seguimento musical", "r_seguimento"),
        "Participação": campo("Participação", "r_participacao"),
    }

    st.subheader("IAPS - Composição")
    iaps_composicao = {
        "Criatividade": campo("Criatividade", "c_criatividade"),
        "Organização de ideias": campo("Organização de ideias", "c_organizacao"),
        "Expressão simbólica": campo("Expressão simbólica", "c_expressao"),
        "Autoria": campo("Autoria", "c_autoria"),
    }

    st.subheader("IAPS - Escuta Musical")
    iaps_escuta = {
        "Atenção auditiva": campo("Atenção auditiva", "e_atencao"),
        "Resposta emocional": campo("Resposta emocional", "e_resposta"),
        "Reflexão": campo("Reflexão", "e_reflexao"),
        "Integração da experiência sonora": campo("Integração da experiência sonora", "e_integracao"),
    }


DEMUCA_DOMINIOS = {
    "Comportamentos Restritivos": {
        "tipo": "restritivo",
        "maximo": 14,
        "itens": [
            ("Estereotipias", 1),
            ("Agressividade", 1),
            ("Desinteresse", 1),
            ("Passividade", 1),
            ("Reclusão (Isolamento)", 1),
            ("Resistência", 1),
            ("Pirraça", 1),
        ],
    },
    "Interação Social - Cognição": {
        "tipo": "positivo",
        "maximo": 18,
        "itens": [
            ("Contato visual", 1),
            ("Comunicação verbal", 1),
            ("Interação com outros objetos", 1),
            ("Interação com instrumentos musicais", 1),
            ("Interação com educador/musicoterapeuta", 1),
            ("Interação com os pais (se aplicável)", 1),
            ("Interação com os pares (se aplicável)", 1),
            ("Atenção", 1),
            ("Imitação", 1),
        ],
    },
    "Percepção - Exploração Rítmica": {
        "tipo": "positivo",
        "maximo": 16,
        "itens": [
            ("Pulso Interno", 1),
            ("Regulação Temporal", 1),
            ("Ritmo Real", 2),
            ("Apoio", 2),
            ("Contrastes de andamento", 2),
        ],
    },
    "Percepção - Exploração Sonora": {
        "tipo": "positivo",
        "maximo": 14,
        "itens": [
            ("Som / Silêncio", 1),
            ("Timbre", 1),
            ("Planos de altura", 1),
            ("Movimento sonoro", 1),
            ("Contrastes de intensidade", 1),
            ("Repetição de ideias rítmicas e/ou melódicas", 1),
            ("Senso de conclusão", 1),
        ],
    },
    "Exploração Vocal": {
        "tipo": "positivo",
        "maximo": 14,
        "itens": [
            ("Vocalizações", 1),
            ("Balbucios", 1),
            ("Sílabas canônicas", 1),
            ("Imitação de canções", 2),
            ("Criação vocal", 2),
        ],
    },
    "Movimentação corporal com a música": {
        "tipo": "positivo",
        "maximo": 14,
        "itens": [
            ("Andar", 1),
            ("Correr", 1),
            ("Parar", 1),
            ("Gesticular", 1),
            ("Dançar", 1),
            ("Movimentar-se no lugar", 1),
            ("Pular", 1),
        ],
    },
}


demuca_respostas = {}

if "DEMUCA" in escalas_escolhidas:
    st.header("DEMUCA")
    st.caption("Escala de classificação: N = Não | P = Pouco | M = Muito")

    for dominio, info in DEMUCA_DOMINIOS.items():
        st.subheader(f"DEMUCA - {dominio}")
        demuca_respostas[dominio] = {}

        for item, peso in info["itens"]:
            label = f"{item}" + ("  (x2)" if peso == 2 else "")
            resposta = st.radio(
                label,
                ["N", "P", "M"],
                horizontal=True,
                key=f"demuca_{dominio}_{item}"
            )
            demuca_respostas[dominio][item] = {
                "resposta": resposta,
                "peso": peso
            }


st.header("GAS / SMART")

meta_gas = st.text_input("META GAS")
gas_menos_2 = st.text_area("-2 Estado atual", value="")
gas_menos_1 = st.text_area("-1 Abaixo do esperado", value="")
gas_zero = st.text_area("0 Desfecho esperado após intervenção", value="")
gas_mais_1 = st.text_area("+1 Acima do esperado", value="")
gas_mais_2 = st.text_area("+2 Muito acima do esperado", value="")

curto_prazo = st.text_area("Objetivo SMART - Curto prazo (3 meses)")
longo_prazo = st.text_area("Objetivo SMART - Longo prazo (6 meses)")
prescricao = st.text_area("Plano terapêutico e prescrição")

st.info("O resumo geral, o parecer técnico e a conduta serão gerados automaticamente com base nos dados das escalas.")


def classificar(valor, maximo):
    percentual = (valor / maximo) * 100 if maximo else 0
    if percentual < 40:
        return "baixo"
    elif percentual < 70:
        return "moderado"
    return "adequado"


def percentual(valor, maximo):
    return round((valor / maximo) * 100, 1) if maximo else 0


def totals_to_order(totais):
    ordem = ["Improvisação", "Recriação", "Composição", "Escuta Musical"]
    return {k: totais[k] for k in ordem if k in totais}


def calcular_demuca(demuca_respostas):
    totais = {}

    for dominio, itens in demuca_respostas.items():
        tipo = DEMUCA_DOMINIOS[dominio]["tipo"]
        total = 0

        for item, dados in itens.items():
            resposta = dados["resposta"]
            peso = dados["peso"]

            if tipo == "restritivo":
                mapa = {"N": 2, "P": 1, "M": 0}
                total += mapa[resposta] * peso
            else:
                mapa = {"N": 0, "P": 1, "M": 2}
                total += mapa[resposta] * peso

        totais[dominio] = total

    return totais


def identificar_areas(totais_iaps):
    areas_baixas = []
    areas_moderadas = []
    areas_adequadas = []

    for area, valor in totals_to_order(totais_iaps).items():
        nivel = classificar(valor, 20)
        if nivel == "baixo":
            areas_baixas.append(area)
        elif nivel == "moderado":
            areas_moderadas.append(area)
        else:
            areas_adequadas.append(area)

    return areas_baixas, areas_moderadas, areas_adequadas


def identificar_areas_demuca(totais_demuca):
    baixas = []
    moderadas = []
    adequadas = []

    for dominio, valor in totais_demuca.items():
        maximo = DEMUCA_DOMINIOS[dominio]["maximo"]
        nivel = classificar(valor, maximo)

        if nivel == "baixo":
            baixas.append(dominio)
        elif nivel == "moderado":
            moderadas.append(dominio)
        else:
            adequadas.append(dominio)

    return baixas, moderadas, adequadas


def gerar_grafico_barras(titulo, dados, maximos=None):
    categorias = list(dados.keys())
    valores = list(dados.values())

    if maximos:
        limite = max(maximos.values())
    else:
        limite = max(valores) if valores else 5

    fig, ax = plt.subplots(figsize=(11, 5.5))
    ax.bar(categorias, valores)
    ax.set_title(titulo)
    ax.set_ylabel("Pontuação")
    ax.set_ylim(0, limite)
    ax.tick_params(axis="x", rotation=35)
    plt.tight_layout()

    buffer = BytesIO()
    fig.savefig(buffer, format="png", bbox_inches="tight", dpi=180)
    plt.close(fig)
    buffer.seek(0)
    return buffer


def gerar_grafico_iaps(totais_iaps):
    return gerar_grafico_barras(
        "IAPS - Pontuação por área",
        totais_iaps,
        {"Improvisação": 20, "Recriação": 20, "Composição": 20, "Escuta Musical": 20}
    )


def gerar_grafico_demuca(totais_demuca):
    maximos = {dominio: DEMUCA_DOMINIOS[dominio]["maximo"] for dominio in totais_demuca}
    return gerar_grafico_barras("DEMUCA - Pontuação por domínio", totais_demuca, maximos)


def gerar_grafico_radar(totais_iaps, nordoff_total=None, totais_demuca=None):
    categorias = []
    valores_percentuais = []

    if nordoff_total is not None:
        categorias.append("Nordoff")
        valores_percentuais.append(percentual(nordoff_total, 50))

    if totais_iaps:
        categorias.extend(["Improvisação", "Recriação", "Composição", "Escuta"])
        valores_percentuais.extend([
            percentual(totais_iaps["Improvisação"], 20),
            percentual(totais_iaps["Recriação"], 20),
            percentual(totais_iaps["Composição"], 20),
            percentual(totais_iaps["Escuta Musical"], 20),
        ])

    if totais_demuca:
        for dominio, valor in totais_demuca.items():
            categorias.append(dominio[:18])
            valores_percentuais.append(percentual(valor, DEMUCA_DOMINIOS[dominio]["maximo"]))

    if len(categorias) < 3:
        categorias = categorias + [""] * (3 - len(categorias))
        valores_percentuais = valores_percentuais + [0] * (3 - len(valores_percentuais))

    angles = [n / float(len(categorias)) * 2 * math.pi for n in range(len(categorias))]
    valores = valores_percentuais + valores_percentuais[:1]
    angles += angles[:1]

    fig = plt.figure(figsize=(8, 8))
    ax = plt.subplot(111, polar=True)
    ax.plot(angles, valores, linewidth=2)
    ax.fill(angles, valores, alpha=0.25)
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categorias)
    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_ylim(0, 100)
    ax.set_title("Perfil Musicoterapêutico Geral (%)", pad=20)

    buffer = BytesIO()
    fig.savefig(buffer, format="png", bbox_inches="tight", dpi=180)
    plt.close(fig)
    buffer.seek(0)
    return buffer


def explicar_escalas(incluir_demuca=False):
    texto = (
        "As Escalas Nordoff-Robbins são instrumentos de avaliação desenvolvidos no contexto da Musicoterapia Criativa, "
        "voltados à observação da comunicabilidade musical, da responsividade, da iniciativa, do engajamento, da reciprocidade "
        "e da organização musical do paciente no setting terapêutico. Sua aplicação permite compreender como o paciente se "
        "comunica musicalmente, como sustenta interações sonoras e como utiliza a música como recurso expressivo, relacional "
        "e regulatório.\n\n"
        "Os IAPS, Improvisation Assessment Profiles, são instrumentos voltados à análise da improvisação clínica em musicoterapia. "
        "Eles auxiliam na observação de aspectos expressivos, interacionais, criativos, motores, cognitivos e perceptivos presentes "
        "na produção musical do paciente. Neste relatório, os IAPS foram organizados nas áreas de Improvisação, Recriação, "
        "Composição e Escuta Musical.\n\n"
        "A escala GAS, Goal Attainment Scaling, é utilizada para estruturar metas terapêuticas individualizadas. Ela descreve níveis "
        "esperados de evolução funcional, de -2 a +2, permitindo mensurar o progresso do paciente em relação a objetivos clínicos "
        "específicos. Associada à formulação SMART, favorece a criação de metas claras, mensuráveis, alcançáveis, relevantes e "
        "temporalmente definidas."
    )

    if incluir_demuca:
        texto += (
            "\n\nA DEMUCA é uma escala musicoterapêutica organizada em domínios funcionais voltados à observação de comportamentos "
            "restritivos, interação social/cognição, percepção e exploração rítmica, percepção e exploração sonora, exploração vocal "
            "e movimentação corporal com a música. Sua estrutura utiliza a classificação N = Não, P = Pouco e M = Muito, permitindo "
            "quantificar indicadores clínicos observáveis no contexto musical. Nos comportamentos restritivos, a pontuação é invertida, "
            "pois a ausência do comportamento restritivo indica melhor funcionamento clínico. Nos demais domínios, maior pontuação indica "
            "maior presença da habilidade observada. Alguns itens possuem peso x2, respeitando sua relevância clínica dentro do domínio."
        )

    return texto


def interpretar_nordoff(dados):
    if not dados:
        return ""

    total = sum(dados.values())
    maximo = len(dados) * 5
    nivel = classificar(total, maximo)

    if nivel == "baixo":
        return (
            "Os resultados indicam baixa disponibilidade comunicativa musical, com necessidade de maior suporte para "
            "responsividade, iniciativa, reciprocidade e sustentação da interação musical. Observa-se que a organização da resposta "
            "musical ainda depende de mediação terapêutica consistente, previsibilidade no setting e propostas sonoro-musicais que "
            "favoreçam vínculo, regulação e engajamento progressivo."
        )

    if nivel == "moderado":
        return (
            "Os resultados indicam presença de recursos comunicativos musicais em desenvolvimento, com respostas funcionais, porém "
            "ainda oscilantes em termos de iniciativa, sustentação, reciprocidade e organização da interação musical. O paciente "
            "demonstra possibilidades de engajamento musical, mas ainda necessita de estruturação clínica para ampliar autonomia, "
            "estabilidade e intenção comunicativa no fazer musical."
        )

    return (
        "Os resultados indicam boa comunicabilidade musical, com presença de iniciativa, responsividade, engajamento e organização "
        "relacional no fazer musical. Observa-se disponibilidade para experiências musicais mais complexas, com possibilidade de "
        "ampliação da autonomia expressiva, da elaboração simbólica e da flexibilidade interacional no setting musicoterapêutico."
    )


def interpretar_iaps(totais):
    texto = ""

    for area, valor in totals_to_order(totais).items():
        nivel = classificar(valor, 20)

        if nivel == "baixo":
            texto += (
                f"{area}: desempenho baixo, sugerindo necessidade de intervenções estruturadas para ampliar recursos musicais, "
                f"expressivos, perceptivos e relacionais. Esse resultado indica que o domínio avaliado ainda não se apresenta "
                f"suficientemente consolidado e demanda mediação terapêutica frequente, previsibilidade e adaptação das propostas "
                f"ao perfil funcional do paciente.\n"
            )
        elif nivel == "moderado":
            texto += (
                f"{area}: desempenho moderado, indicando recursos presentes, porém ainda em processo de consolidação clínica. "
                f"Observa-se potencial terapêutico a ser ampliado por meio de propostas graduadas, repetição estruturada, aumento "
                f"progressivo da complexidade musical e fortalecimento da autonomia funcional.\n"
            )
        else:
            texto += (
                f"{area}: desempenho adequado, indicando boa disponibilidade funcional para esse domínio musical. Esse resultado "
                f"pode ser utilizado como ponto de apoio clínico para favorecer engajamento, motivação, expressividade e transferência "
                f"de habilidades para áreas com maior necessidade terapêutica.\n"
            )

    return texto


def interpretar_demuca(totais_demuca):
    texto = ""

    for dominio, valor in totais_demuca.items():
        maximo = DEMUCA_DOMINIOS[dominio]["maximo"]
        nivel = classificar(valor, maximo)
        porcentagem = percentual(valor, maximo)

        if dominio == "Comportamentos Restritivos":
            if nivel == "baixo":
                texto += (
                    f"{dominio}: pontuação {valor}/{maximo} ({porcentagem}%), classificada como baixa. "
                    f"Considerando a lógica invertida deste domínio, o resultado sugere presença relevante de comportamentos restritivos, "
                    f"como estereotipias, resistência, passividade, isolamento, agressividade, desinteresse ou pirraça, demandando manejo "
                    f"musicoterapêutico estruturado, previsibilidade, co-regulação e adaptação do setting.\n"
                )
            elif nivel == "moderado":
                texto += (
                    f"{dominio}: pontuação {valor}/{maximo} ({porcentagem}%), classificada como moderada. "
                    f"O resultado sugere presença parcial ou oscilante de comportamentos restritivos, indicando necessidade de monitoramento "
                    f"e intervenções voltadas à regulação, previsibilidade e ampliação do engajamento funcional.\n"
                )
            else:
                texto += (
                    f"{dominio}: pontuação {valor}/{maximo} ({porcentagem}%), classificada como adequada. "
                    f"O resultado sugere menor presença de comportamentos restritivos, favorecendo maior disponibilidade para participação, "
                    f"interação e responsividade nas experiências musicoterapêuticas.\n"
                )
        else:
            if nivel == "baixo":
                texto += (
                    f"{dominio}: pontuação {valor}/{maximo} ({porcentagem}%), classificada como baixa. "
                    f"Esse resultado indica que as habilidades avaliadas neste domínio ainda se encontram pouco disponíveis ou pouco organizadas "
                    f"no contexto sonoro-musical, demandando intervenção estruturada, repetição, mediação e progressão gradual.\n"
                )
            elif nivel == "moderado":
                texto += (
                    f"{dominio}: pontuação {valor}/{maximo} ({porcentagem}%), classificada como moderada. "
                    f"O domínio apresenta recursos em desenvolvimento, com respostas parciais ou inconsistentes, sugerindo necessidade de "
                    f"fortalecimento da autonomia, estabilidade e funcionalidade musical.\n"
                )
            else:
                texto += (
                    f"{dominio}: pontuação {valor}/{maximo} ({porcentagem}%), classificada como adequada. "
                    f"O domínio demonstra recursos funcionais relevantes, podendo ser utilizado como via terapêutica para ampliar comunicação, "
                    f"engajamento, organização, expressão e integração musical.\n"
                )

    return texto


def gerar_resumo_geral(nordoff_total, totais_iaps, totais_demuca):
    texto = (
        f"{nome if nome else 'O paciente'}, {idade} anos, com diagnóstico de "
        f"{diagnostico if diagnostico else 'não informado'}, foi avaliado em musicoterapia por meio de instrumentos clínicos voltados "
        f"à compreensão da comunicação musical, da interação, da expressividade, da organização sonora, da escuta, do comportamento "
        f"e da participação funcional no setting terapêutico.\n\n"
    )

    if nordoff_dominios:
        nivel_nordoff = classificar(nordoff_total, 50)
        texto += (
            f"Na Escala de Comunicabilidade Musical Nordoff-Robbins, o desempenho geral foi classificado como {nivel_nordoff}. "
            f"{interpretar_nordoff(nordoff_dominios)}\n\n"
        )

    if totais_iaps:
        areas_baixas, areas_moderadas, areas_adequadas = identificar_areas(totais_iaps)

        if areas_baixas:
            texto += (
                f"Nos IAPS, as áreas com menor desempenho foram: {', '.join(areas_baixas)}. Esses achados indicam que tais domínios devem "
                f"ser priorizados no plano terapêutico, pois podem interferir diretamente na qualidade da expressão musical, da organização "
                f"da resposta, da interação e da integração da experiência sonora.\n\n"
            )

        if areas_moderadas:
            texto += (
                f"As áreas classificadas como moderadas foram: {', '.join(areas_moderadas)}. Esses domínios demonstram potencial de "
                f"desenvolvimento, necessitando de propostas graduadas, repetição estruturada e aumento progressivo da complexidade "
                f"terapêutica.\n\n"
            )

        if areas_adequadas:
            texto += (
                f"As áreas com desempenho adequado foram: {', '.join(areas_adequadas)}. Esses aspectos representam recursos clínicos "
                f"importantes e podem ser utilizados como vias de acesso para favorecer engajamento, motivação, comunicação, regulação "
                f"e ampliação de habilidades em áreas de maior necessidade.\n\n"
            )

    if totais_demuca:
        demuca_baixas, demuca_moderadas, demuca_adequadas = identificar_areas_demuca(totais_demuca)

        texto += (
            "Na DEMUCA, a avaliação contemplou domínios relacionados aos comportamentos restritivos, interação social/cognição, "
            "percepção e exploração rítmica, percepção e exploração sonora, exploração vocal e movimentação corporal com a música. "
        )

        if demuca_baixas:
            texto += (
                f"Os domínios com menor desempenho foram: {', '.join(demuca_baixas)}. Esses resultados sugerem necessidade de intervenção "
                f"musicoterapêutica direcionada, com adaptação do setting, mediação ativa, previsibilidade e organização progressiva das "
                f"respostas musicais, corporais, vocais e interacionais.\n\n"
            )

        if demuca_moderadas:
            texto += (
                f"Os domínios classificados como moderados foram: {', '.join(demuca_moderadas)}. Tais indicadores mostram habilidades "
                f"em desenvolvimento, com necessidade de fortalecimento clínico, repetição estruturada e ampliação da funcionalidade no "
                f"contexto musical.\n\n"
            )

        if demuca_adequadas:
            texto += (
                f"Os domínios classificados como adequados foram: {', '.join(demuca_adequadas)}. Esses domínios representam pontos de apoio "
                f"para o planejamento terapêutico, podendo ser usados para promover engajamento, comunicação, regulação e participação "
                f"mais ativa.\n\n"
            )

    texto += (
        "De modo geral, o perfil musicoterapêutico observado sugere que o planejamento deve ser individualizado, integrando estratégias "
        "de improvisação clínica, escuta ativa, recriação musical, composição guiada, exploração sonoro-corporal, exploração vocal, "
        "experiências rítmicas e atividades musicais significativas. A intervenção deverá respeitar o ritmo do paciente, suas preferências, "
        "suas recusas, suas possibilidades sensoriais e seu nível atual de organização emocional, comunicativa e relacional."
    )

    return texto


def gerar_objetivos(nordoff_total, totais_iaps, totais_demuca):
    objetivos = []

    if nordoff_dominios and classificar(nordoff_total, 50) in ["baixo", "moderado"]:
        objetivos.append("Ampliar a comunicabilidade musical, favorecendo iniciativa, responsividade, reciprocidade e sustentação da interação terapêutica.")

    if totais_iaps:
        if classificar(totais_iaps["Improvisação"], 20) in ["baixo", "moderado"]:
            objetivos.append("Estimular a espontaneidade sonora, a exploração criativa e a construção de respostas musicais intencionais.")
        if classificar(totais_iaps["Recriação"], 20) in ["baixo", "moderado"]:
            objetivos.append("Fortalecer memória musical, coordenação motora, seguimento de modelos e participação em atividades musicais estruturadas.")
        if classificar(totais_iaps["Composição"], 20) in ["baixo", "moderado"]:
            objetivos.append("Favorecer criatividade, autoria, organização simbólica e expressão de conteúdos internos por meio da produção musical.")
        if classificar(totais_iaps["Escuta Musical"], 20) in ["baixo", "moderado"]:
            objetivos.append("Promover escuta ativa, atenção auditiva, resposta emocional e integração subjetiva da experiência sonora.")

    if totais_demuca:
        for dominio, valor in totais_demuca.items():
            maximo = DEMUCA_DOMINIOS[dominio]["maximo"]
            if classificar(valor, maximo) in ["baixo", "moderado"]:
                if dominio == "Comportamentos Restritivos":
                    objetivos.append("Reduzir a interferência de comportamentos restritivos no setting musicoterapêutico por meio de previsibilidade, co-regulação, organização sensorial e engajamento musical funcional.")
                elif dominio == "Interação Social - Cognição":
                    objetivos.append("Ampliar contato visual, atenção, imitação e interação social mediada pela música.")
                elif dominio == "Percepção - Exploração Rítmica":
                    objetivos.append("Desenvolver pulso interno, regulação temporal, apoio rítmico e percepção de contrastes de andamento.")
                elif dominio == "Percepção - Exploração Sonora":
                    objetivos.append("Estimular discriminação sonora, percepção de timbre, planos de altura, intensidade, movimento sonoro e senso de conclusão.")
                elif dominio == "Exploração Vocal":
                    objetivos.append("Favorecer vocalizações, balbucios, sílabas canônicas, imitação de canções e criação vocal em contexto terapêutico.")
                elif dominio == "Movimentação corporal com a música":
                    objetivos.append("Ampliar organização motora, expressão corporal e movimentação funcional associada à experiência musical.")

    if not objetivos:
        objetivos.append("Aprofundar os recursos musicais já estabelecidos, ampliando complexidade expressiva, flexibilidade interacional e autonomia musical.")

    return objetivos


def gerar_estrategias():
    return [
        "Utilizar improvisação clínica estruturada para favorecer diálogo sonoro, turn-taking e responsividade musical.",
        "Empregar canções estruturadas para previsibilidade, organização temporal e sustentação da atenção.",
        "Realizar atividades de escuta ativa com contrastes sonoros, pausas e mediação terapêutica.",
        "Propor experiências de recriação musical com repertório significativo para o paciente.",
        "Desenvolver propostas de composição guiada para estimular autoria, simbolização e expressão emocional.",
        "Organizar o setting musicoterapêutico de forma previsível, acessível e ajustada às necessidades sensoriais e relacionais do paciente.",
        "Utilizar propostas rítmicas, vocais, corporais e instrumentais de forma progressiva, respeitando o perfil funcional e os domínios avaliados."
    ]


def gerar_conduta_automatizada(nordoff_total, totais_iaps, totais_demuca):
    conduta = (
        "A partir da análise integrada dos dados avaliativos, recomenda-se a continuidade do acompanhamento musicoterapêutico com "
        "planejamento individualizado, considerando a relação entre comunicabilidade musical, expressão sonora, organização temporal, "
        "responsividade relacional, escuta, criatividade, recursos sensório-motores, exploração vocal, percepção rítmica, percepção sonora, "
        "comportamentos restritivos e possibilidades de autorregulação do paciente.\n\n"
    )

    if nordoff_dominios:
        nivel_nordoff = classificar(nordoff_total, 50)

        if nivel_nordoff == "baixo":
            conduta += (
                "A baixa pontuação geral na Escala de Comunicabilidade Musical Nordoff-Robbins indica que a conduta inicial deve priorizar "
                "a construção de vínculo terapêutico, a organização do setting, a previsibilidade das propostas e a ampliação gradual da "
                "responsividade musical. Recomenda-se utilizar intervenções de baixa complexidade, com estímulos sonoro-musicais claros, "
                "repetitivos e responsivos, favorecendo co-regulação, contato relacional, imitação sonora, alternância de turnos e aumento "
                "progressivo da iniciativa comunicativa.\n\n"
            )
        elif nivel_nordoff == "moderado":
            conduta += (
                "A classificação moderada na Escala de Comunicabilidade Musical Nordoff-Robbins indica que o paciente apresenta recursos "
                "comunicativos musicais em desenvolvimento. A conduta deverá favorecer maior consistência, autonomia e sustentação das "
                "respostas musicais, ampliando a reciprocidade, a intenção comunicativa, o engajamento e a organização musical em experiências "
                "compartilhadas.\n\n"
            )
        else:
            conduta += (
                "A classificação adequada na Escala de Comunicabilidade Musical Nordoff-Robbins indica boa disponibilidade comunicativa musical. "
                "A conduta poderá avançar para propostas de maior complexidade expressiva, simbólica e relacional, favorecendo autonomia, "
                "criatividade, elaboração musical, flexibilidade interacional e integração de recursos musicais com objetivos terapêuticos "
                "mais amplos.\n\n"
            )

    if totais_iaps:
        areas_baixas, areas_moderadas, areas_adequadas = identificar_areas(totais_iaps)

        if areas_baixas:
            conduta += (
                f"As áreas com desempenho mais baixo nos IAPS foram: {', '.join(areas_baixas)}. Recomenda-se que esses domínios sejam "
                f"priorizados no planejamento terapêutico, com intervenções graduadas, mediação constante, uso de repetição estruturada, "
                f"adaptações sensoriais e apoio clínico para favorecer a organização da resposta musical e a ampliação da funcionalidade "
                f"terapêutica.\n\n"
            )

        if areas_moderadas:
            conduta += (
                f"As áreas com desempenho moderado foram: {', '.join(areas_moderadas)}. Para esses domínios, recomenda-se progressão gradual "
                f"da complexidade, com propostas que favoreçam estabilidade, independência, maior elaboração musical e generalização das "
                f"habilidades observadas no setting terapêutico.\n\n"
            )

        if areas_adequadas:
            conduta += (
                f"As áreas com desempenho adequado foram: {', '.join(areas_adequadas)}. Tais recursos devem ser utilizados como pontos de apoio "
                f"clínico, contribuindo para engajamento, motivação, vínculo terapêutico e ampliação de habilidades em áreas que apresentaram "
                f"maior necessidade de suporte.\n\n"
            )

    if totais_demuca:
        demuca_baixas, demuca_moderadas, demuca_adequadas = identificar_areas_demuca(totais_demuca)

        if demuca_baixas:
            conduta += (
                f"Na DEMUCA, os domínios com menor desempenho foram: {', '.join(demuca_baixas)}. Recomenda-se direcionar o plano terapêutico "
                f"para o fortalecimento desses domínios, utilizando experiências musicais progressivas, atividades de exploração sonora e rítmica, "
                f"propostas vocais e corporais, intervenção relacional e estratégias de regulação. Quando o domínio de comportamentos restritivos "
                f"apresentar baixo desempenho, deve-se priorizar previsibilidade, estrutura, co-regulação e redução da interferência desses "
                f"comportamentos na participação musical.\n\n"
            )

        if demuca_moderadas:
            conduta += (
                f"Os domínios DEMUCA com desempenho moderado foram: {', '.join(demuca_moderadas)}. Recomenda-se fortalecer esses repertórios "
                f"por meio de repetição estruturada, variação gradual, ampliação da autonomia e integração entre corpo, voz, instrumento, escuta "
                f"e interação social.\n\n"
            )

        if demuca_adequadas:
            conduta += (
                f"Os domínios DEMUCA classificados como adequados foram: {', '.join(demuca_adequadas)}. Esses domínios devem ser utilizados "
                f"como vias terapêuticas de acesso para sustentar engajamento, ampliar comunicação e favorecer evolução em áreas de maior "
                f"necessidade.\n\n"
            )

    conduta += (
        "A conduta musicoterapêutica deverá integrar improvisação clínica, recriação musical, escuta ativa, composição guiada, exploração "
        "sonoro-corporal, exploração vocal, experiências rítmicas e atividades musicais significativas, sempre considerando o perfil sensorial, "
        "afetivo, cognitivo, motor, comunicativo e relacional do paciente. Recomenda-se reavaliação periódica dos objetivos terapêuticos, com "
        "ajustes conforme a evolução clínica, o engajamento nas sessões e a qualidade das respostas musicais observadas.\n\n"
        "Ao final deste parecer técnico, ressalta-se a importância de manter sessões de musicoterapia de forma regular, pois a continuidade "
        "do processo favorece a consolidação de habilidades, a ampliação da comunicação musical, a estabilidade regulatória e o desenvolvimento "
        "progressivo dos objetivos terapêuticos estabelecidos."
    )

    return conduta


def referencias():
    return [
        "Nordoff, P., & Robbins, C. (2007). Creative Music Therapy. Gilsum: Barcelona Publishers.",
        "Bruscia, K. (1998). Defining Music Therapy. Barcelona Publishers.",
        "Wigram, T., Pedersen, I. N., & Bonde, L. O. (2002). A Comprehensive Guide to Music Therapy. Jessica Kingsley Publishers.",
        "Bruscia, K. E. (1987). Improvisational Models of Music Therapy. Springfield: Charles C Thomas.",
        "Kiresuk, T. J., & Sherman, R. E. (1968). Goal Attainment Scaling: A general method for evaluating comprehensive community mental health programs. Community Mental Health Journal.",
        "DEMUCA. Escala de avaliação musicoterapêutica organizada por domínios funcionais: comportamentos restritivos, interação social/cognição, percepção rítmica, percepção sonora, exploração vocal e movimentação corporal com a música."
    ]


def limpar_documento(doc):
    body = doc._body._element
    for child in list(body):
        if child.tag.endswith("sectPr"):
            continue
        body.remove(child)


def add_title(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(0, 0, 0)


def add_section(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(12)


def add_subsection(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(11)


def add_text(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(text)
    run.font.size = Pt(11)


def add_bullet(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(f"• {text}")
    run.font.size = Pt(11)


def add_thumbnail_metodos(doc):
    for imagem in ["thumbnail_metodos_intensivos.png", "tumbnail_metodos_intensivos.png"]:
        if os.path.exists(imagem):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run()
            run.add_picture(imagem, width=Inches(2.2))
            return


def criar_word_modelo(grafico_nordoff, grafico_iaps, grafico_demuca, grafico_radar, totais_iaps, totais_demuca, nordoff_total, resumo_geral, conduta_final):
    try:
        doc = Document("modelo_relatorio.docx")
    except Exception:
        doc = Document()

    limpar_documento(doc)

    add_thumbnail_metodos(doc)
    add_title(doc, "MUSICOTERAPIA")

    add_section(doc, "IDENTIFICAÇÃO")

    tabela_id = doc.add_table(rows=0, cols=2)
    tabela_id.style = "Table Grid"

    dados_id = [
        ("Nome do paciente", nome),
        ("Data de nascimento", data_nascimento),
        ("Idade", str(idade)),
        ("Escolaridade", escolaridade),
        ("Nome dos pais ou responsáveis", responsaveis),
        ("Métodos Intensivos / Abordagem Terapêutica", metodo_intensivo),
        ("Diagnóstico", diagnostico),
        ("Data de início da intervenção", data_inicio),
        ("Data atual", data_atual),
        ("Nome do profissional", terapeuta),
        ("Registro profissional", registro),
    ]

    for campo_nome, campo_valor in dados_id:
        row = tabela_id.add_row().cells
        row[0].text = campo_nome
        row[1].text = campo_valor if campo_valor else ""

    add_section(doc, "HISTÓRIA CLÍNICA")
    add_text(doc, historia_clinica if historia_clinica else "Não informado.")

    add_section(doc, "AVALIAÇÕES MUSICOTERAPÊUTICAS")
    add_text(
        doc,
        f"Em {data_atual if data_atual else 'data não informada'}, foram realizadas avaliações musicoterapêuticas com base nos instrumentos selecionados, com o objetivo de coletar dados funcionais para elaboração do plano musicoterapêutico especializado."
    )

    add_subsection(doc, "INSTRUMENTOS AVALIATIVOS")
    add_text(doc, explicar_escalas("DEMUCA" in escalas_escolhidas))

    add_section(doc, "Preferências e Recusas sonoro-musicais")
    add_text(doc, preferencias if preferencias else "Não informado.")

    add_section(doc, "Interações sonoro-musicais")
    add_text(doc, interacoes if interacoes else "Não informado.")

    add_section(doc, "RESUMO GERAL")
    add_text(doc, resumo_geral)

    if grafico_nordoff:
        add_section(doc, "Gráfico - Escala de Comunicabilidade Musical Nordoff-Robbins")
        doc.add_picture(grafico_nordoff, width=Inches(6.2))

    if totais_iaps:
        add_section(doc, "Resultados IAPS")
        tabela_iaps = doc.add_table(rows=1, cols=3)
        tabela_iaps.style = "Table Grid"

        hdr = tabela_iaps.rows[0].cells
        hdr[0].text = "Área"
        hdr[1].text = "Pontuação"
        hdr[2].text = "Classificação"

        for area, valor in totals_to_order(totais_iaps).items():
            row = tabela_iaps.add_row().cells
            row[0].text = area
            row[1].text = f"{valor}/20"
            row[2].text = classificar(valor, 20).capitalize()

        add_section(doc, "Gráfico - IAPS")
        doc.add_picture(grafico_iaps, width=Inches(6.2))

    if totais_demuca:
        add_section(doc, "Resultados DEMUCA")
        tabela_demuca = doc.add_table(rows=1, cols=3)
        tabela_demuca.style = "Table Grid"

        hdr = tabela_demuca.rows[0].cells
        hdr[0].text = "Domínio"
        hdr[1].text = "Pontuação"
        hdr[2].text = "Classificação"

        for dominio, valor in totais_demuca.items():
            maximo = DEMUCA_DOMINIOS[dominio]["maximo"]
            row = tabela_demuca.add_row().cells
            row[0].text = dominio
            row[1].text = f"{valor}/{maximo}"
            row[2].text = classificar(valor, maximo).capitalize()

        add_section(doc, "Interpretação DEMUCA")
        add_text(doc, interpretar_demuca(totais_demuca))

        add_section(doc, "Gráfico - DEMUCA")
        doc.add_picture(grafico_demuca, width=Inches(6.2))

    add_section(doc, "Gráfico - Perfil Musicoterapêutico Geral")
    doc.add_picture(grafico_radar, width=Inches(5.7))

    add_section(doc, "PLANO TERAPÊUTICO")
    add_section(doc, "METAS CENTRADAS NA FAMÍLIA E NA CRIANÇA")
    add_text(
        doc,
        f"{nome if nome else 'O paciente'}, {idade} anos, com diagnóstico de "
        f"{diagnostico if diagnostico else 'não informado'}, apresenta um perfil musicoterapêutico que demanda planejamento "
        f"individualizado, considerando suas respostas musicais, sua forma de engajamento, suas preferências sonoro-musicais "
        f"e os domínios avaliados pelas escalas utilizadas."
    )

    add_section(doc, "Objetivos e metas terapêuticas (SMART)")
    add_text(doc, f"Curto Prazo (3 meses): {curto_prazo if curto_prazo else 'A definir conforme evolução clínica.'}")
    add_text(doc, f"Longo Prazo (6 meses): {longo_prazo if longo_prazo else 'A definir conforme evolução clínica.'}")

    add_section(doc, "Objetivos terapêuticos")
    for obj in gerar_objetivos(nordoff_total, totais_iaps, totais_demuca):
        add_bullet(doc, obj)

    add_section(doc, "Estratégias terapêuticas sugeridas")
    for est in gerar_estrategias():
        add_bullet(doc, est)

    add_section(doc, "GAS/SMART")
    add_text(doc, f"META 01: {meta_gas}")

    tabela_gas = doc.add_table(rows=1, cols=3)
    tabela_gas.style = "Table Grid"

    hdr = tabela_gas.rows[0].cells
    hdr[0].text = "Escore"
    hdr[1].text = "Descrição"
    hdr[2].text = "Objetivo / Meta prevista"

    dados_gas = [
        ("-2", "Estado atual", gas_menos_2),
        ("-1", "Abaixo do esperado", gas_menos_1),
        ("0", "Desfecho esperado após intervenção", gas_zero),
        ("+1", "Acima do esperado", gas_mais_1),
        ("+2", "Muito acima do esperado", gas_mais_2),
    ]

    for escore, descricao, meta in dados_gas:
        row = tabela_gas.add_row().cells
        row[0].text = escore
        row[1].text = descricao
        row[2].text = meta

    add_section(doc, "Plano terapêutico e prescrição")
    add_text(doc, prescricao if prescricao else "Incluindo intensidade, frequência, duração e participantes da equipe interdisciplinar.")

    add_section(doc, "PARECER TÉCNICO / CONDUTA")
    add_text(doc, conduta_final)

    add_section(doc, "Referências")
    for ref in referencias():
        add_bullet(doc, ref)

    doc.add_paragraph("\n\n__________________________________")
    doc.add_paragraph(terapeuta)
    doc.add_paragraph("MUSICOTERAPEUTA")
    doc.add_paragraph(registro)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


if st.button("Gerar relatório clínico em Word"):
    nordoff_total = sum(nordoff_dominios.values()) if nordoff_dominios else None

    totais_iaps = {}
    if iaps_improvisacao:
        totais_iaps = {
            "Improvisação": sum(iaps_improvisacao.values()),
            "Recriação": sum(iaps_recriacao.values()),
            "Composição": sum(iaps_composicao.values()),
            "Escuta Musical": sum(iaps_escuta.values()),
        }

    totais_demuca = {}
    if demuca_respostas:
        totais_demuca = calcular_demuca(demuca_respostas)

    resumo_geral = gerar_resumo_geral(nordoff_total, totais_iaps, totais_demuca)
    conduta_final = gerar_conduta_automatizada(nordoff_total, totais_iaps, totais_demuca)

    grafico_nordoff = None
    if nordoff_dominios:
        grafico_nordoff = gerar_grafico_barras(
            "Escala de Comunicabilidade Musical Nordoff-Robbins",
            nordoff_dominios,
            {k: 5 for k in nordoff_dominios}
        )

    grafico_iaps = None
    if totais_iaps:
        grafico_iaps = gerar_grafico_iaps(totais_iaps)

    grafico_demuca = None
    if totais_demuca:
        grafico_demuca = gerar_grafico_demuca(totais_demuca)

    grafico_radar = gerar_grafico_radar(totais_iaps, nordoff_total, totais_demuca)

    if grafico_nordoff:
        st.subheader("Gráfico Nordoff-Robbins")
        st.image(grafico_nordoff)

    if grafico_iaps:
        st.subheader("Gráfico IAPS")
        st.image(grafico_iaps)

    if grafico_demuca:
        st.subheader("Gráfico DEMUCA")
        st.image(grafico_demuca)

    st.subheader("Gráfico Radar - Perfil Geral")
    st.image(grafico_radar)

    st.subheader("Resumo geral automatizado")
    st.text_area("Resumo geral", resumo_geral, height=350)

    st.subheader("Parecer técnico / Conduta automatizada")
    st.text_area("Parecer técnico / Conduta", conduta_final, height=400)

    word = criar_word_modelo(
        grafico_nordoff=grafico_nordoff,
        grafico_iaps=grafico_iaps,
        grafico_demuca=grafico_demuca,
        grafico_radar=grafico_radar,
        totais_iaps=totais_iaps,
        totais_demuca=totais_demuca,
        nordoff_total=nordoff_total,
        resumo_geral=resumo_geral,
        conduta_final=conduta_final,
    )

    nome_arquivo = f"relatorio_{nome.replace(' ', '_') if nome else 'paciente'}.docx"

    st.download_button(
        "📄 Baixar relatório em Word",
        data=word,
        file_name=nome_arquivo,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )