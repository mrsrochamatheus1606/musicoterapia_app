import os
import math
from io import BytesIO

import streamlit as st
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt


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
# ESCALAS SELECIONADAS
# ==========================================================
st.header("Escalas a utilizar no relatório")

escalas_escolhidas = st.multiselect(
    "Selecione as avaliações que deseja aplicar",
    ["Nordoff-Robbins", "IAPS", "DEMUCA"],
    default=["Nordoff-Robbins", "IAPS"]
)


def campo(label, key):
    return st.number_input(label, min_value=0, max_value=5, value=0, step=1, key=key)


# ==========================================================
# NORDOFF-ROBBINS
# ==========================================================
nordoff_dominios = {}

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


# ==========================================================
# IAPS
# ==========================================================
iaps_improvisacao = {}
iaps_recriacao = {}
iaps_composicao = {}
iaps_escuta = {}

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


# ==========================================================
# DEMUCA
# ==========================================================
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


st.info("A GAS, o plano terapêutico, a prescrição e a conduta serão gerados automaticamente a partir dos scores das avaliações.")


# ==========================================================
# FUNÇÕES DE CÁLCULO
# ==========================================================
def classificar(valor, maximo):
    percentual_calc = (valor / maximo) * 100 if maximo else 0

    if percentual_calc < 40:
        return "baixo"
    elif percentual_calc < 70:
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
            else:
                mapa = {"N": 0, "P": 1, "M": 2}

            total += mapa[resposta] * peso

        totais[dominio] = total

    return totais


def identificar_prejuizos(nordoff_total, totais_iaps, totais_demuca):
    prejuizos = []

    if nordoff_total is not None:
        if classificar(nordoff_total, 50) in ["baixo", "moderado"]:
            prejuizos.append({
                "area": "Comunicabilidade musical",
                "origem": "Nordoff-Robbins",
                "nivel": classificar(nordoff_total, 50),
                "objetivo": "Ampliar a comunicação musical funcional, favorecendo iniciativa, responsividade, reciprocidade e sustentação da interação sonoro-musical.",
                "habilidade": "comunicação musical, responsividade e vínculo terapêutico"
            })

    if totais_iaps:
        mapa_iaps = {
            "Improvisação": {
                "objetivo": "Estimular a iniciativa sonora, a espontaneidade musical e a construção de respostas musicais intencionais em contexto improvisacional.",
                "habilidade": "iniciativa sonora e expressão espontânea"
            },
            "Recriação": {
                "objetivo": "Fortalecer memória musical, coordenação motora, imitação, seguimento de modelos e participação em atividades musicais estruturadas.",
                "habilidade": "memória musical, coordenação e participação estruturada"
            },
            "Composição": {
                "objetivo": "Favorecer criatividade, autoria, simbolização e organização de ideias musicais por meio da produção musical guiada.",
                "habilidade": "criatividade, autoria e elaboração simbólica"
            },
            "Escuta Musical": {
                "objetivo": "Promover atenção auditiva, escuta ativa, resposta emocional e integração subjetiva da experiência sonora.",
                "habilidade": "escuta ativa, atenção e integração sonora"
            },
        }

        for area, valor in totals_to_order(totais_iaps).items():
            nivel = classificar(valor, 20)
            if nivel in ["baixo", "moderado"]:
                prejuizos.append({
                    "area": area,
                    "origem": "IAPS",
                    "nivel": nivel,
                    "objetivo": mapa_iaps[area]["objetivo"],
                    "habilidade": mapa_iaps[area]["habilidade"]
                })

    if totais_demuca:
        mapa_demuca = {
            "Comportamentos Restritivos": {
                "objetivo": "Reduzir a interferência de comportamentos restritivos no setting musicoterapêutico, favorecendo co-regulação, previsibilidade, disponibilidade relacional e engajamento musical funcional.",
                "habilidade": "regulação, disponibilidade relacional e redução de interferências comportamentais"
            },
            "Interação Social - Cognição": {
                "objetivo": "Ampliar contato visual, atenção compartilhada, imitação, comunicação e interação social mediada pela música.",
                "habilidade": "interação social, atenção compartilhada e cognição musical"
            },
            "Percepção - Exploração Rítmica": {
                "objetivo": "Desenvolver pulso interno, regulação temporal, apoio rítmico, ritmo real e percepção de contrastes de andamento.",
                "habilidade": "percepção rítmica, organização temporal e coordenação musical"
            },
            "Percepção - Exploração Sonora": {
                "objetivo": "Estimular discriminação sonora, percepção de timbre, planos de altura, movimento sonoro, intensidade, repetição de ideias musicais e senso de conclusão.",
                "habilidade": "percepção sonora, discriminação auditiva e organização musical"
            },
            "Exploração Vocal": {
                "objetivo": "Favorecer vocalizações, balbucios, sílabas canônicas, imitação de canções e criação vocal em contexto terapêutico.",
                "habilidade": "expressão vocal, comunicação vocal e musicalidade da voz"
            },
            "Movimentação corporal com a música": {
                "objetivo": "Ampliar organização motora, expressão corporal, coordenação, deslocamento e movimentação funcional associada à experiência musical.",
                "habilidade": "expressão corporal, coordenação motora e integração música-movimento"
            },
        }

        for dominio, valor in totais_demuca.items():
            nivel = classificar(valor, DEMUCA_DOMINIOS[dominio]["maximo"])
            if nivel in ["baixo", "moderado"]:
                prejuizos.append({
                    "area": dominio,
                    "origem": "DEMUCA",
                    "nivel": nivel,
                    "objetivo": mapa_demuca[dominio]["objetivo"],
                    "habilidade": mapa_demuca[dominio]["habilidade"]
                })

    if not prejuizos:
        prejuizos.append({
            "area": "Manutenção e ampliação de repertório musicoterapêutico",
            "origem": "Síntese clínica",
            "nivel": "adequado",
            "objetivo": "Aprofundar os recursos musicais já estabelecidos, ampliando complexidade expressiva, autonomia, flexibilidade interacional e elaboração simbólica no setting musicoterapêutico.",
            "habilidade": "autonomia musical, flexibilidade e elaboração expressiva"
        })

    return prejuizos


# ==========================================================
# GRÁFICOS
# ==========================================================
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
        categorias += [""] * (3 - len(categorias))
        valores_percentuais += [0] * (3 - len(valores_percentuais))

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


# ==========================================================
# TEXTOS TÉCNICOS
# ==========================================================
def explicar_escalas(escalas_selecionadas):
    textos = []

    if "Nordoff-Robbins" in escalas_selecionadas:
        textos.append(
            "As Escalas Nordoff-Robbins são instrumentos de avaliação desenvolvidos no contexto da Musicoterapia Criativa, "
            "voltados à observação da comunicabilidade musical, da responsividade, da iniciativa, do engajamento, da reciprocidade "
            "e da organização musical do paciente no setting terapêutico. Sua aplicação permite compreender como o paciente se comunica "
            "musicalmente, como sustenta interações sonoras e como utiliza a música como recurso expressivo, relacional e regulatório."
        )

    if "IAPS" in escalas_selecionadas:
        textos.append(
            "Os IAPS, Improvisation Assessment Profiles, são instrumentos voltados à análise da improvisação clínica em musicoterapia. "
            "Eles auxiliam na observação de aspectos expressivos, interacionais, criativos, motores, cognitivos e perceptivos presentes "
            "na produção musical do paciente."
        )

    if "DEMUCA" in escalas_selecionadas:
        textos.append(
            "A DEMUCA é uma escala musicoterapêutica organizada em domínios funcionais voltados à observação de comportamentos "
            "restritivos, interação social/cognição, percepção e exploração rítmica, percepção e exploração sonora, exploração vocal "
            "e movimentação corporal com a música. Sua estrutura utiliza a classificação N = Não, P = Pouco e M = Muito. Nos comportamentos "
            "restritivos, a pontuação é invertida, pois a ausência do comportamento restritivo indica melhor funcionamento clínico. "
            "Nos demais domínios, maior pontuação indica maior presença da habilidade observada."
        )

    textos.append(
        "A escala GAS, Goal Attainment Scaling, é utilizada para estruturar metas terapêuticas individualizadas. Ela descreve níveis "
        "esperados de evolução funcional, de -2 a +2, permitindo mensurar o progresso do paciente em relação a objetivos clínicos "
        "específicos."
    )

    return "\n\n".join(textos)


def interpretar_nordoff(dados):
    if not dados:
        return ""

    total = sum(dados.values())
    nivel = classificar(total, 50)

    if nivel == "baixo":
        return (
            "Os resultados indicam baixa disponibilidade comunicativa musical, com necessidade de maior suporte para responsividade, "
            "iniciativa, reciprocidade e sustentação da interação musical."
        )

    if nivel == "moderado":
        return (
            "Os resultados indicam presença de recursos comunicativos musicais em desenvolvimento, com respostas funcionais, porém ainda "
            "oscilantes em termos de iniciativa, sustentação, reciprocidade e organização da interação musical."
        )

    return (
        "Os resultados indicam boa comunicabilidade musical, com presença de iniciativa, responsividade, engajamento e organização relacional "
        "no fazer musical."
    )


def interpretar_iaps(totais):
    texto = ""

    for area, valor in totals_to_order(totais).items():
        nivel = classificar(valor, 20)

        if nivel == "baixo":
            texto += f"{area}: desempenho baixo, sugerindo necessidade de intervenções estruturadas para ampliar recursos musicais, expressivos, perceptivos e relacionais.\n"
        elif nivel == "moderado":
            texto += f"{area}: desempenho moderado, indicando recursos presentes, porém ainda em processo de consolidação clínica.\n"
        else:
            texto += f"{area}: desempenho adequado, indicando boa disponibilidade funcional para esse domínio musical.\n"

    return texto


def interpretar_demuca(totais_demuca):
    texto = ""

    for dominio, valor in totais_demuca.items():
        maximo = DEMUCA_DOMINIOS[dominio]["maximo"]
        nivel = classificar(valor, maximo)

        if dominio == "Comportamentos Restritivos":
            if nivel == "baixo":
                texto += f"{dominio}: pontuação {valor}/{maximo}, classificada como baixa. Considerando a lógica invertida, sugere presença relevante de comportamentos restritivos interferindo na disponibilidade musical e relacional.\n"
            elif nivel == "moderado":
                texto += f"{dominio}: pontuação {valor}/{maximo}, classificada como moderada. Sugere presença parcial ou oscilante de comportamentos restritivos.\n"
            else:
                texto += f"{dominio}: pontuação {valor}/{maximo}, classificada como adequada. Sugere menor interferência de comportamentos restritivos no setting.\n"
        else:
            if nivel == "baixo":
                texto += f"{dominio}: pontuação {valor}/{maximo}, classificada como baixa. Indica habilidade pouco disponível ou pouco organizada no contexto sonoro-musical.\n"
            elif nivel == "moderado":
                texto += f"{dominio}: pontuação {valor}/{maximo}, classificada como moderada. Indica repertório em desenvolvimento.\n"
            else:
                texto += f"{dominio}: pontuação {valor}/{maximo}, classificada como adequada. Indica recursos funcionais importantes para o processo terapêutico.\n"

    return texto


def gerar_resumo_geral(nordoff_total, totais_iaps, totais_demuca):
    texto = (
        f"{nome if nome else 'O paciente'}, {idade} anos, com diagnóstico de {diagnostico if diagnostico else 'não informado'}, "
        f"foi avaliado em musicoterapia por meio dos instrumentos selecionados. A avaliação contemplou aspectos relacionados à comunicação "
        f"musical, responsividade, engajamento, interação, percepção rítmica, percepção sonora, exploração vocal, movimentação corporal, "
        f"expressividade e organização musical.\n\n"
    )

    if nordoff_total is not None:
        texto += f"Na Escala Nordoff-Robbins, o desempenho geral foi classificado como {classificar(nordoff_total, 50)}. {interpretar_nordoff(nordoff_dominios)}\n\n"

    if totais_iaps:
        texto += "Nos IAPS, observou-se o seguinte perfil:\n"
        texto += interpretar_iaps(totais_iaps) + "\n"

    if totais_demuca:
        texto += "Na DEMUCA, observou-se o seguinte perfil:\n"
        texto += interpretar_demuca(totais_demuca) + "\n"

    texto += (
        "A integração dos dados sugere a necessidade de um planejamento terapêutico individualizado, fundamentado nas respostas musicais "
        "observadas e na relação entre comunicação, escuta, corpo, voz, ritmo, interação e regulação. O plano deve priorizar as habilidades "
        "em prejuízo, utilizando os recursos preservados como vias de acesso terapêutico."
    )

    return texto


def gerar_gas_automatico(prejuizos):
    metas = []

    for idx, item in enumerate(prejuizos[:3], start=1):
        habilidade = item["habilidade"]
        objetivo = item["objetivo"]

        metas.append({
            "meta": f"META {idx:02d}: {habilidade.capitalize()}",
            "-2": f"Não demonstra {habilidade} de forma funcional durante as experiências musicoterapêuticas, mesmo com mediação intensa.",
            "-1": f"Demonstra {habilidade} de forma inconsistente, com necessidade de suporte máximo, alta previsibilidade e mediação contínua do musicoterapeuta.",
            "0": objetivo,
            "+1": f"Demonstra {habilidade} com maior consistência, necessitando de suporte moderado e apresentando respostas musicais mais organizadas e intencionais.",
            "+2": f"Demonstra {habilidade} de forma funcional, espontânea e generalizável no setting musicoterapêutico, com maior autonomia e qualidade relacional."
        })

    return metas


def gerar_objetivos(prejuizos):
    return [p["objetivo"] for p in prejuizos]


def gerar_prescricao_automatica(prejuizos):
    principais = ", ".join([p["area"] for p in prejuizos[:4]])

    texto = (
        "Recomenda-se acompanhamento musicoterapêutico regular, com frequência mínima de 1 a 2 sessões semanais, duração média de 40 a 50 minutos, "
        "podendo ser ajustada conforme tolerância, disponibilidade atencional, perfil sensorial e resposta clínica do paciente.\n\n"
        f"O plano terapêutico deverá priorizar os domínios identificados como mais prejudicados: {principais}. As intervenções deverão ser estruturadas "
        "a partir de experiências musicais graduadas, com objetivos funcionais claros, previsibilidade, repetição terapêutica, variação progressiva de "
        "complexidade e uso de repertório significativo para o paciente.\n\n"
        "A prescrição musicoterapêutica deverá incluir improvisação clínica, recriação musical, escuta ativa, composição guiada, exploração vocal, "
        "atividades rítmicas, propostas de movimento com música e intervenções de co-regulação sonoro-musical. O setting deverá ser organizado de forma "
        "a favorecer segurança, vínculo, engajamento, iniciativa, atenção compartilhada e participação ativa."
    )

    return texto


def gerar_conduta_automatizada(nordoff_total, totais_iaps, totais_demuca, prejuizos):
    texto = (
        "A partir da análise integrada dos dados avaliativos, recomenda-se a continuidade do acompanhamento musicoterapêutico com planejamento "
        "individualizado, considerando a relação entre comunicabilidade musical, expressão sonora, organização temporal, responsividade relacional, "
        "escuta, criatividade, recursos sensório-motores, exploração vocal, percepção rítmica, percepção sonora, comportamentos restritivos e "
        "possibilidades de autorregulação do paciente.\n\n"
    )

    texto += (
        "As áreas em prejuízo indicam necessidade de intervenções direcionadas, musicalmente estruturadas e clinicamente graduadas. O trabalho deverá "
        "partir de experiências sonoro-musicais acessíveis, favorecendo previsibilidade, vínculo terapêutico, co-regulação, iniciativa musical, "
        "organização da resposta, sustentação atencional, comunicação não verbal, exploração vocal e integração corpo-música.\n\n"
    )

    texto += "Os principais focos terapêuticos identificados foram:\n"
    for p in prejuizos[:5]:
        texto += f"- {p['area']} ({p['origem']}): {p['objetivo']}\n"

    texto += (
        "\nA conduta musicoterapêutica deverá integrar improvisação clínica, recriação musical, escuta ativa, composição guiada, exploração sonoro-corporal, "
        "exploração vocal, experiências rítmicas e atividades musicais significativas, sempre considerando o perfil sensorial, afetivo, cognitivo, motor, "
        "comunicativo e relacional do paciente. Recomenda-se reavaliação periódica dos objetivos terapêuticos, com ajustes conforme a evolução clínica, "
        "o engajamento nas sessões e a qualidade das respostas musicais observadas.\n\n"
        "Ao final deste parecer técnico, ressalta-se a importância de manter sessões de musicoterapia de forma regular, pois a continuidade do processo "
        "favorece a consolidação de habilidades, a ampliação da comunicação musical, a estabilidade regulatória e o desenvolvimento progressivo dos "
        "objetivos terapêuticos estabelecidos."
    )

    return texto


def gerar_estrategias(prejuizos):
    estrategias = [
        "Utilizar improvisação clínica estruturada para favorecer diálogo sonoro, turn-taking e responsividade musical.",
        "Empregar canções estruturadas para previsibilidade, organização temporal e sustentação da atenção.",
        "Realizar atividades de escuta ativa com contrastes sonoros, pausas e mediação terapêutica.",
        "Propor experiências de recriação musical com repertório significativo para o paciente.",
        "Desenvolver propostas de composição guiada para estimular autoria, simbolização e expressão emocional.",
        "Organizar o setting musicoterapêutico de forma previsível, acessível e ajustada às necessidades sensoriais e relacionais do paciente."
    ]

    for p in prejuizos:
        if "vocal" in p["habilidade"]:
            estrategias.append("Utilizar jogos vocais, vocalizações espelhadas, canções responsivas e imitação melódica para estimular expressão vocal.")
        if "rítmica" in p["habilidade"] or "temporal" in p["habilidade"]:
            estrategias.append("Aplicar padrões rítmicos simples, pulso marcado, alternância de andamento e atividades de apoio rítmico corporal/instrumental.")
        if "corporal" in p["habilidade"] or "motora" in p["habilidade"]:
            estrategias.append("Integrar movimento corporal, gestos, deslocamentos e instrumentos de fácil acesso para favorecer coordenação música-movimento.")
        if "regulação" in p["habilidade"]:
            estrategias.append("Utilizar músicas previsíveis, andamento estável, contorno melódico simples e pausas terapêuticas para co-regulação.")

    return list(dict.fromkeys(estrategias))


def referencias():
    refs = []

    if "Nordoff-Robbins" in escalas_escolhidas:
        refs.append("Nordoff, P., & Robbins, C. (2007). Creative Music Therapy. Gilsum: Barcelona Publishers.")

    if "IAPS" in escalas_escolhidas:
        refs.append("Bruscia, K. E. (1987). Improvisational Models of Music Therapy. Springfield: Charles C Thomas.")

    if "DEMUCA" in escalas_escolhidas:
        refs.append("DEMUCA. Escala de avaliação musicoterapêutica organizada por domínios funcionais: comportamentos restritivos, interação social/cognição, percepção rítmica, percepção sonora, exploração vocal e movimentação corporal com a música.")

    refs.extend([
        "Bruscia, K. (1998). Defining Music Therapy. Barcelona Publishers.",
        "Wigram, T., Pedersen, I. N., & Bonde, L. O. (2002). A Comprehensive Guide to Music Therapy. Jessica Kingsley Publishers.",
        "Kiresuk, T. J., & Sherman, R. E. (1968). Goal Attainment Scaling: A general method for evaluating comprehensive community mental health programs. Community Mental Health Journal."
    ])

    return refs


# ==========================================================
# WORD
# ==========================================================
def limpar_documento(doc):
    body = doc._body._element
    for child in list(body):
        if child.tag.endswith("sectPr"):
            continue
        body.remove(child)


def add_title(doc, text):
    p = doc.add_paragraph()
    p.alignment = 1
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(16)


def add_section(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(12)


def add_text(doc, text):
    p = doc.add_paragraph()
    p.alignment = 3
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
            p.alignment = 1
            run = p.add_run()
            run.add_picture(imagem, width=Inches(2.2))
            return


def criar_word_modelo(
    grafico_nordoff,
    grafico_iaps,
    grafico_demuca,
    grafico_radar,
    totais_iaps,
    totais_demuca,
    nordoff_total,
    resumo_geral,
    conduta_final,
    prescricao_auto,
    gas_auto,
    prejuizos
):
    try:
        doc = Document("modelo_relatorio.docx")
    except Exception:
        doc = Document()

    limpar_documento(doc)

    if metodo_intensivo in ["MIG", "TREINI"]:
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

    add_section(doc, "INSTRUMENTOS AVALIATIVOS")
    add_text(doc, explicar_escalas(escalas_escolhidas))

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
    add_section(doc, "Objetivos terapêuticos gerados automaticamente")

    for obj in gerar_objetivos(prejuizos):
        add_bullet(doc, obj)

    add_section(doc, "Estratégias terapêuticas sugeridas")
    for est in gerar_estrategias(prejuizos):
        add_bullet(doc, est)

    add_section(doc, "GAS / SMART gerado automaticamente")

    for meta in gas_auto:
        add_text(doc, meta["meta"])

        tabela_gas = doc.add_table(rows=1, cols=3)
        tabela_gas.style = "Table Grid"

        hdr = tabela_gas.rows[0].cells
        hdr[0].text = "Escore"
        hdr[1].text = "Descrição"
        hdr[2].text = "Objetivo / Meta prevista"

        dados = [
            ("-2", "Estado atual", meta["-2"]),
            ("-1", "Abaixo do esperado", meta["-1"]),
            ("0", "Desfecho esperado após intervenção", meta["0"]),
            ("+1", "Acima do esperado", meta["+1"]),
            ("+2", "Muito acima do esperado", meta["+2"]),
        ]

        for escore, descricao, objetivo in dados:
            row = tabela_gas.add_row().cells
            row[0].text = escore
            row[1].text = descricao
            row[2].text = objetivo

    add_section(doc, "Plano terapêutico e prescrição gerados automaticamente")
    add_text(doc, prescricao_auto)

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

    prejuizos = identificar_prejuizos(nordoff_total, totais_iaps, totais_demuca)
    gas_auto = gerar_gas_automatico(prejuizos)
    prescricao_auto = gerar_prescricao_automatica(prejuizos)

    resumo_geral = gerar_resumo_geral(nordoff_total, totais_iaps, totais_demuca)
    conduta_final = gerar_conduta_automatizada(nordoff_total, totais_iaps, totais_demuca, prejuizos)

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

    st.subheader("GAS / SMART gerado automaticamente")
    for meta in gas_auto:
        st.write(f"**{meta['meta']}**")
        st.write(meta)

    st.subheader("Plano terapêutico e prescrição automatizados")
    st.text_area("Plano terapêutico e prescrição", prescricao_auto, height=300)

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
        prescricao_auto=prescricao_auto,
        gas_auto=gas_auto,
        prejuizos=prejuizos
    )

    nome_arquivo = f"relatorio_{nome.replace(' ', '_') if nome else 'paciente'}.docx"

    st.download_button(
        "📄 Baixar relatório em Word",
        data=word,
        file_name=nome_arquivo,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )