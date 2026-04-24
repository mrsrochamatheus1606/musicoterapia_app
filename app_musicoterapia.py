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


# ==========================================================
# DADOS DO PROFISSIONAL
# ==========================================================
st.header("Dados do Profissional")

col1, col2 = st.columns(2)
with col1:
    terapeuta = st.text_input("Nome completo do terapeuta")
with col2:
    registro = st.text_input("Registro profissional")


# ==========================================================
# IDENTIFICAÇÃO DO PACIENTE
# ==========================================================
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


# ==========================================================
# CAMPOS NUMÉRICOS
# ==========================================================
def campo(label, key):
    return st.number_input(label, min_value=0, max_value=5, value=0, step=1, key=key)


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


# ==========================================================
# GAS / SMART
# ==========================================================
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


# ==========================================================
# FUNÇÕES DE CLASSIFICAÇÃO
# ==========================================================
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


# ==========================================================
# GRÁFICOS
# ==========================================================
def gerar_grafico_barras(titulo, dados, maximo_por_item=5):
    categorias = list(dados.keys())
    valores = list(dados.values())

    fig, ax = plt.subplots(figsize=(11, 5.5))
    ax.bar(categorias, valores)
    ax.set_title(titulo)
    ax.set_ylabel("Pontuação")
    ax.set_ylim(0, maximo_por_item)
    ax.tick_params(axis="x", rotation=35)
    plt.tight_layout()

    buffer = BytesIO()
    fig.savefig(buffer, format="png", bbox_inches="tight", dpi=180)
    plt.close(fig)
    buffer.seek(0)
    return buffer


def gerar_grafico_iaps(totais_iaps):
    categorias = list(totais_iaps.keys())
    valores = list(totais_iaps.values())

    fig, ax = plt.subplots(figsize=(9, 5))
    ax.bar(categorias, valores)
    ax.set_title("IAPS - Pontuação por área")
    ax.set_ylabel("Pontuação total")
    ax.set_ylim(0, 20)
    ax.tick_params(axis="x", rotation=20)
    plt.tight_layout()

    buffer = BytesIO()
    fig.savefig(buffer, format="png", bbox_inches="tight", dpi=180)
    plt.close(fig)
    buffer.seek(0)
    return buffer


def gerar_grafico_radar(totais_iaps, nordoff_total):
    categorias = ["Nordoff", "Improvisação", "Recriação", "Composição", "Escuta Musical"]

    valores_percentuais = [
        percentual(nordoff_total, 50),
        percentual(totais_iaps["Improvisação"], 20),
        percentual(totais_iaps["Recriação"], 20),
        percentual(totais_iaps["Composição"], 20),
        percentual(totais_iaps["Escuta Musical"], 20),
    ]

    angles = [n / float(len(categorias)) * 2 * math.pi for n in range(len(categorias))]
    valores = valores_percentuais + valores_percentuais[:1]
    angles += angles[:1]

    fig = plt.figure(figsize=(7, 7))
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


# ==========================================================
# TEXTOS TÉCNICOS
# ==========================================================
def explicar_escalas():
    return (
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


def interpretar_nordoff(dados):
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


def gerar_resumo_geral(nordoff_total, totais_iaps):
    nivel_nordoff = classificar(nordoff_total, 50)
    areas_baixas, areas_moderadas, areas_adequadas = identificar_areas(totais_iaps)

    texto = (
        f"{nome if nome else 'O paciente'}, {idade} anos, com diagnóstico de "
        f"{diagnostico if diagnostico else 'não informado'}, foi avaliado em musicoterapia por meio da Escala de "
        f"Comunicabilidade Musical Nordoff-Robbins, dos IAPS e da estrutura GAS/SMART. A avaliação contemplou indicadores "
        f"relacionados à comunicabilidade musical, responsividade, iniciativa, reciprocidade, engajamento, organização musical, "
        f"improvisação, recriação, composição e escuta musical.\n\n"
    )

    if nivel_nordoff == "baixo":
        texto += (
            "Na Escala de Comunicabilidade Musical Nordoff-Robbins, o desempenho geral foi classificado como baixo. Esse resultado "
            "sugere que a comunicação musical ainda se apresenta em estágio inicial, com necessidade de maior suporte terapêutico "
            "para que respostas musicais, corporais, afetivas e relacionais possam ser organizadas de forma mais funcional. "
            "Observa-se a necessidade de intervenções que favoreçam vínculo, previsibilidade, co-regulação, turn-taking, imitação "
            "sonora e ampliação da intencionalidade comunicativa.\n\n"
        )
    elif nivel_nordoff == "moderado":
        texto += (
            "Na Escala de Comunicabilidade Musical Nordoff-Robbins, o desempenho geral foi classificado como moderado. Esse resultado "
            "indica que há recursos comunicativos musicais presentes, porém ainda em processo de consolidação. O paciente demonstra "
            "possibilidades de engajamento e responsividade, mas necessita de continuidade terapêutica para ampliar estabilidade, "
            "iniciativa, reciprocidade e autonomia no fazer musical.\n\n"
        )
    else:
        texto += (
            "Na Escala de Comunicabilidade Musical Nordoff-Robbins, o desempenho geral foi classificado como adequado. Esse resultado "
            "indica boa disponibilidade para a comunicação musical, com presença de recursos expressivos, relacionais e organizacionais "
            "que podem ser aprofundados em propostas terapêuticas de maior complexidade.\n\n"
        )

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

    texto += (
        "De modo geral, o perfil musicoterapêutico observado sugere que o planejamento deve ser individualizado, integrando estratégias "
        "de improvisação clínica, escuta ativa, recriação musical, composição guiada e uso de experiências sonoro-musicais significativas. "
        "A intervenção deverá respeitar o ritmo do paciente, suas preferências, suas recusas, suas possibilidades sensoriais e seu nível "
        "atual de organização emocional, comunicativa e relacional."
    )

    return texto


def gerar_objetivos(nordoff_total, totais_iaps):
    objetivos = []

    if classificar(nordoff_total, 50) in ["baixo", "moderado"]:
        objetivos.append("Ampliar a comunicabilidade musical, favorecendo iniciativa, responsividade, reciprocidade e sustentação da interação terapêutica.")

    if classificar(totais_iaps["Improvisação"], 20) in ["baixo", "moderado"]:
        objetivos.append("Estimular a espontaneidade sonora, a exploração criativa e a construção de respostas musicais intencionais.")

    if classificar(totais_iaps["Recriação"], 20) in ["baixo", "moderado"]:
        objetivos.append("Fortalecer memória musical, coordenação motora, seguimento de modelos e participação em atividades musicais estruturadas.")

    if classificar(totais_iaps["Composição"], 20) in ["baixo", "moderado"]:
        objetivos.append("Favorecer criatividade, autoria, organização simbólica e expressão de conteúdos internos por meio da produção musical.")

    if classificar(totais_iaps["Escuta Musical"], 20) in ["baixo", "moderado"]:
        objetivos.append("Promover escuta ativa, atenção auditiva, resposta emocional e integração subjetiva da experiência sonora.")

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
    ]


def gerar_conduta_automatizada(nordoff_total, totais_iaps):
    nivel_nordoff = classificar(nordoff_total, 50)
    areas_baixas, areas_moderadas, areas_adequadas = identificar_areas(totais_iaps)

    conduta = (
        "A partir da análise integrada dos dados avaliativos, recomenda-se a continuidade do acompanhamento musicoterapêutico com "
        "planejamento individualizado, considerando a relação entre comunicabilidade musical, expressão sonora, organização temporal, "
        "responsividade relacional, escuta, criatividade, recursos sensório-motores e possibilidades de autorregulação do paciente.\n\n"
    )

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

    conduta += (
        "A conduta musicoterapêutica deverá integrar improvisação clínica, recriação musical, escuta ativa, composição guiada, exploração "
        "sonoro-corporal e experiências musicais significativas, sempre considerando o perfil sensorial, afetivo, cognitivo, motor, "
        "comunicativo e relacional do paciente. Recomenda-se reavaliação periódica dos objetivos terapêuticos, com ajustes conforme a "
        "evolução clínica, o engajamento nas sessões e a qualidade das respostas musicais observadas.\n\n"
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
    ]


# ==========================================================
# WORD - FORMATAÇÃO
# ==========================================================
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
    return p


def add_section(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(12)
    return p


def add_subsection(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(11)
    return p


def add_text(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(text)
    run.font.size = Pt(11)
    return p


def add_bullet(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(f"• {text}")
    run.font.size = Pt(11)
    return p


def add_thumbnail_metodos(doc):
    imagem = "thumbnail_metodos_intensivos.png"
    if os.path.exists(imagem):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(imagem, width=Inches(2.2))


def criar_word_modelo(grafico_nordoff, grafico_iaps, grafico_radar, totais_iaps, nordoff_total, resumo_geral, conduta_final):
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
        f"Em {data_atual if data_atual else 'data não informada'}, foram realizadas avaliações musicoterapêuticas com base na "
        f"Escala de Comunicabilidade Musical Nordoff-Robbins, nos IAPS e na estrutura GAS/SMART, com o objetivo de coletar dados "
        f"funcionais para elaboração do plano musicoterapêutico especializado."
    )

    add_subsection(doc, "ESCALAS NORDOFF-ROBBINS")
    add_text(
        doc,
        "As Escalas Nordoff-Robbins são ferramentas de avaliação desenvolvidas no contexto da Musicoterapia Criativa. Têm como objetivo "
        "observar qualitativamente a comunicabilidade musical, o engajamento, a responsividade, a iniciativa interpessoal, a reciprocidade "
        "e a organização musical do paciente em sessão. A análise permite compreender a forma como o paciente utiliza a música como meio "
        "de expressão, comunicação, regulação e relação terapêutica."
    )

    add_subsection(doc, "IAPS")
    add_text(
        doc,
        "Os IAPS, Improvisation Assessment Profiles, são instrumentos voltados à análise da improvisação clínica em musicoterapia. Sua "
        "utilização permite observar indicadores de iniciativa, resposta musical, organização sonora, interação, memória, coordenação, "
        "criatividade, autoria, escuta e integração da experiência sonora, oferecendo suporte ao raciocínio clínico e ao planejamento "
        "terapêutico baseado em dados observáveis."
    )

    add_subsection(doc, "GAS/SMART")
    add_text(
        doc,
        "A escala GAS, Goal Attainment Scaling, é utilizada para estruturar metas terapêuticas individualizadas. Ela descreve níveis "
        "esperados de evolução funcional, de -2 a +2, permitindo mensurar o progresso do paciente em relação a objetivos clínicos "
        "específicos. Associada à formulação SMART, favorece a criação de metas claras, mensuráveis, alcançáveis, relevantes e "
        "temporalmente definidas."
    )

    add_section(doc, "Preferências e Recusas sonoro-musicais")
    add_text(doc, preferencias if preferencias else "Não informado.")

    add_section(doc, "Interações sonoro-musicais")
    add_text(doc, interacoes if interacoes else "Não informado.")

    add_section(doc, "RESUMO GERAL")
    add_text(doc, resumo_geral)

    add_section(doc, "Gráfico - Escala de Comunicabilidade Musical Nordoff-Robbins")
    doc.add_picture(grafico_nordoff, width=Inches(6.2))

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
    for obj in gerar_objetivos(nordoff_total, totais_iaps):
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


# ==========================================================
# GERAR RELATÓRIO
# ==========================================================
if st.button("Gerar relatório clínico em Word"):
    nordoff_total = sum(nordoff_dominios.values())

    totais_iaps = {
        "Improvisação": sum(iaps_improvisacao.values()),
        "Recriação": sum(iaps_recriacao.values()),
        "Composição": sum(iaps_composicao.values()),
        "Escuta Musical": sum(iaps_escuta.values()),
    }

    resumo_geral = gerar_resumo_geral(nordoff_total, totais_iaps)
    conduta_final = gerar_conduta_automatizada(nordoff_total, totais_iaps)

    grafico_nordoff = gerar_grafico_barras(
        "Escala de Comunicabilidade Musical Nordoff-Robbins",
        nordoff_dominios,
        maximo_por_item=5
    )

    grafico_iaps = gerar_grafico_iaps(totais_iaps)
    grafico_radar = gerar_grafico_radar(totais_iaps, nordoff_total)

    st.subheader("Gráfico Nordoff-Robbins")
    st.image(grafico_nordoff)

    st.subheader("Gráfico IAPS")
    st.image(grafico_iaps)

    st.subheader("Gráfico Radar - Perfil Geral")
    st.image(grafico_radar)

    st.subheader("Resumo geral automatizado")
    st.text_area("Resumo geral", resumo_geral, height=350)

    st.subheader("Parecer técnico / Conduta automatizada")
    st.text_area("Parecer técnico / Conduta", conduta_final, height=400)

    word = criar_word_modelo(
        grafico_nordoff=grafico_nordoff,
        grafico_iaps=grafico_iaps,
        grafico_radar=grafico_radar,
        totais_iaps=totais_iaps,
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