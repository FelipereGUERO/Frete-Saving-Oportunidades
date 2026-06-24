# app.py
# =============================================================================
# Aplicativo Streamlit - Análise de Oportunidade de Frete Dedicado
# Branding Parker-Hannifin
#
# Instalação:
#   pip install streamlit pandas openpyxl plotly xlsxwriter
#
# Execução:
#   streamlit run app.py
# =============================================================================

import io
import re
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


# =============================================================================
# Configuração visual e branding
# =============================================================================

st.set_page_config(
    page_title="Parker-Hannifin | Oportunidade de Frete Dedicado",
    page_icon="🚚",
    layout="wide",
)

PARKER_RED = "#C8102E"
PARKER_DARK = "#1F2933"
PARKER_GRAY = "#F3F4F6"
PARKER_LIGHT_RED = "#FDECEF"

st.markdown(
    f"""
    <style>
        .main {{
            background-color: #FFFFFF;
        }}
        .block-container {{
            padding-top: 1.5rem;
            padding-bottom: 2rem;
        }}
        .parker-header {{
            background: linear-gradient(90deg, {PARKER_RED} 0%, #8A0018 100%);
            padding: 1.2rem 1.5rem;
            border-radius: 12px;
            color: white;
            margin-bottom: 1.2rem;
            box-shadow: 0 4px 12px rgba(0,0,0,0.12);
        }}
        .parker-header h1 {{
            margin: 0;
            font-size: 2.0rem;
            font-weight: 800;
        }}
        .parker-header p {{
            margin: 0.35rem 0 0 0;
            font-size: 1.0rem;
            opacity: 0.95;
        }}
        .metric-card {{
            background-color: {PARKER_GRAY};
            border-left: 5px solid {PARKER_RED};
            border-radius: 10px;
            padding: 1rem;
        }}
        div[data-testid="stMetricValue"] {{
            color: {PARKER_DARK};
            font-weight: 800;
        }}
        section[data-testid="stSidebar"] h1,
        section[data-testid="stSidebar"] h2,
        section[data-testid="stSidebar"] h3 {{
            color: {PARKER_RED};
        }}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="parker-header">
        <h1>Parker-Hannifin | Análise de Oportunidade de Frete Dedicado</h1>
        <p>Ferramenta de apoio à decisão logística entre frete fracionado e frete dedicado.</p>
    </div>
    """,
    unsafe_allow_html=True,
)


# =============================================================================
# Funções utilitárias
# =============================================================================

def formatar_moeda(valor):
    """Formata número como moeda brasileira."""
    try:
        if pd.isna(valor):
            return "R$ 0,00"
        valor = float(valor)
        return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"


def formatar_percentual(valor):
    """Formata número decimal como percentual brasileiro."""
    try:
        if pd.isna(valor):
            return "0,00%"
        return f"{float(valor) * 100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,00%"


def converter_numero(valor):
    """
    Converte números em formatos brasileiros e internacionais para float.

    Exemplos aceitos:
    - '1.234,56' -> 1234.56
    - '1234,56'  -> 1234.56
    - '1,234.56' -> 1234.56
    - 'R$ 5.000,00' -> 5000.0
    - pandas Series -> Series numérica
    """
    if isinstance(valor, pd.Series):
        return valor.apply(converter_numero)

    if valor is None:
        return np.nan

    try:
        if pd.isna(valor):
            return np.nan
    except Exception:
        pass

    if isinstance(valor, (int, float, np.integer, np.floating)):
        return float(valor)

    texto = str(valor).strip()
    if texto == "" or texto.lower() in {"nan", "none", "null", "n/a", "na", "-"}:
        return np.nan

    # Remove moeda, espaços e caracteres não numéricos relevantes.
    texto = texto.replace("R$", "").replace(" ", "")
    texto = re.sub(r"[^0-9,\.\-]", "", texto)

    if texto in {"", "-", ",", "."}:
        return np.nan

    try:
        # Se houver vírgula e ponto, o último separador tende a ser decimal.
        if "," in texto and "." in texto:
            if texto.rfind(",") > texto.rfind("."):
                # Formato brasileiro: 1.234,56
                texto = texto.replace(".", "").replace(",", ".")
            else:
                # Formato internacional: 1,234.56
                texto = texto.replace(",", "")
        elif "," in texto:
            # Apenas vírgula: assume vírgula decimal.
            texto = texto.replace(".", "").replace(",", ".")
        else:
            # Apenas ponto: pode ser decimal ou milhar.
            # Se houver múltiplos pontos, remove todos exceto o último.
            if texto.count(".") > 1:
                partes = texto.split(".")
                texto = "".join(partes[:-1]) + "." + partes[-1]

        return float(texto)
    except Exception:
        return np.nan


def carregar_arquivo(arquivo):
    """
    Carrega arquivo .xlsx, .xls ou .csv para DataFrame.
    Preserva todas as colunas inicialmente como texto para evitar perda de chaves/documentos.
    """
    if arquivo is None:
        return None

    nome = arquivo.name.lower()
    conteudo = arquivo.getvalue()

    try:
        if nome.endswith(".csv"):
            # Tenta diferentes encodings comuns no Brasil.
            ultimo_erro = None
            for encoding in ["utf-8-sig", "utf-8", "latin1", "cp1252"]:
                try:
                    buffer = io.BytesIO(conteudo)
                    df = pd.read_csv(
                        buffer,
                        dtype=str,
                        sep=None,
                        engine="python",
                        encoding=encoding,
                    )
                    return df
                except Exception as erro:
                    ultimo_erro = erro
            raise ultimo_erro

        if nome.endswith(".xlsx"):
            return pd.read_excel(io.BytesIO(conteudo), dtype=str, engine="openpyxl")

        if nome.endswith(".xls"):
            # Para .xls, o pandas pode requerer xlrd instalado no ambiente.
            return pd.read_excel(io.BytesIO(conteudo), dtype=str)

        st.error("Formato não suportado. Envie arquivo .xlsx, .xls ou .csv.")
        return None

    except Exception as erro:
        st.error(
            "Não foi possível carregar o arquivo. "
            "Verifique se o arquivo está íntegro e se as dependências estão instaladas. "
            f"Detalhe técnico: {erro}"
        )
        return None


def criar_download_excel(df_resultado, df_oportunidades=None, resumo=None):
    """Cria arquivo Excel em memória para download."""
    output = io.BytesIO()

    if df_oportunidades is None:
        df_oportunidades = pd.DataFrame()
    if resumo is None:
        resumo = {}

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pd.DataFrame([resumo]).to_excel(writer, index=False, sheet_name="Resumo")
        df_resultado.to_excel(writer, index=False, sheet_name="Analise_Completa")
        df_oportunidades.to_excel(writer, index=False, sheet_name="Oportunidades")

        workbook = writer.book
        header_format = workbook.add_format(
            {
                "bold": True,
                "text_wrap": True,
                "valign": "top",
                "fg_color": PARKER_RED,
                "font_color": "white",
                "border": 1,
            }
        )
        money_format = workbook.add_format({"num_format": "R$ #,##0.00"})
        percent_format = workbook.add_format({"num_format": "0.00%"})

        for sheet_name in ["Resumo", "Analise_Completa", "Oportunidades"]:
            worksheet = writer.sheets[sheet_name]
            df_sheet = pd.DataFrame([resumo]) if sheet_name == "Resumo" else (
                df_resultado if sheet_name == "Analise_Completa" else df_oportunidades
            )

            for col_num, col_name in enumerate(df_sheet.columns):
                worksheet.write(0, col_num, col_name, header_format)
                largura = min(max(len(str(col_name)) + 2, 14), 45)
                worksheet.set_column(col_num, col_num, largura)

            for idx, col_name in enumerate(df_sheet.columns):
                nome_col = str(col_name).lower()
                if any(palavra in nome_col for palavra in ["valor", "frete", "economia", "mercadoria"]):
                    worksheet.set_column(idx, idx, 16, money_format)
                if any(palavra in nome_col for palavra in ["percentual", "%"]):
                    worksheet.set_column(idx, idx, 14, percent_format)

    output.seek(0)
    return output.getvalue()


def aplicar_regras(df, parametros):
    """
    Aplica as regras de decisão em um DataFrame padronizado.

    Colunas esperadas, quando disponíveis:
    - documento
    - modalidade
    - valor_mercadoria
    - peso
    - frete_fracionado
    - frete_dedicado
    - origem
    - destino
    - data
    - transportadora
    """
    resultado = df.copy()

    colunas_padrao = [
        "documento",
        "modalidade",
        "valor_mercadoria",
        "peso",
        "frete_fracionado",
        "frete_dedicado",
        "origem",
        "destino",
        "data",
        "transportadora",
    ]
    for coluna in colunas_padrao:
        if coluna not in resultado.columns:
            resultado[coluna] = "" if coluna not in ["valor_mercadoria", "peso", "frete_fracionado", "frete_dedicado"] else np.nan

    # Preserva chaves/documentos como texto.
    resultado["documento"] = resultado["documento"].fillna("").astype(str)

    # Converte campos numéricos.
    for coluna in ["valor_mercadoria", "peso", "frete_fracionado", "frete_dedicado"]:
        resultado[coluna] = converter_numero(resultado[coluna])

    resultado["valor_mercadoria"] = resultado["valor_mercadoria"].fillna(0.0)
    resultado["peso"] = resultado["peso"].fillna(0.0)
    resultado["frete_fracionado"] = resultado["frete_fracionado"].fillna(0.0)

    # Valor dedicado informado somente quando positivo.
    resultado["frete_dedicado"] = resultado["frete_dedicado"].where(resultado["frete_dedicado"] > 0, np.nan)

    resultado["percentual_frete_sobre_mercadoria"] = np.where(
        resultado["valor_mercadoria"] > 0,
        resultado["frete_fracionado"] / resultado["valor_mercadoria"],
        0.0,
    )
    resultado["frete_fracionado_por_kg"] = np.where(
        resultado["peso"] > 0,
        resultado["frete_fracionado"] / resultado["peso"],
        0.0,
    )

    # Identifica se o registro é fracionado. Se modalidade não estiver disponível, assume fracionado.
    modalidade_texto = resultado["modalidade"].fillna("").astype(str).str.lower()
    modalidade_vazia = modalidade_texto.str.strip().eq("") | modalidade_texto.isin(["nan", "none", "não disponível", "nao disponivel"])
    resultado["eh_fracionado"] = modalidade_vazia | modalidade_texto.str.contains(
        "fracion|frac|ltl|consolid", regex=True, na=False
    )

    cond_a = resultado["valor_mercadoria"] >= parametros["limite_valor_mercadoria"]
    cond_b = resultado["peso"] >= parametros["limite_peso"]
    cond_c = resultado["percentual_frete_sobre_mercadoria"] >= parametros["limite_percentual_frete"]
    cond_d = (resultado["frete_fracionado_por_kg"] >= parametros["limite_frete_por_kg"]) & (
        resultado["peso"] >= parametros["limite_peso_para_frete_kg"]
    )
    cond_e = resultado["frete_fracionado"] >= parametros["limite_valor_frete"]

    resultado["regra_valor_mercadoria"] = cond_a
    resultado["regra_peso"] = cond_b
    resultado["regra_percentual_frete"] = cond_c
    resultado["regra_frete_por_kg"] = cond_d
    resultado["regra_valor_frete"] = cond_e

    resultado["bate_regra_cotacao"] = cond_a | cond_b | cond_c | cond_d | cond_e
    resultado["qtd_regras_acionadas"] = (
        cond_a.astype(int) + cond_b.astype(int) + cond_c.astype(int) + cond_d.astype(int) + cond_e.astype(int)
    )

    def regras_acionadas_linha(linha):
        regras = []
        if linha["regra_valor_mercadoria"]:
            regras.append("Valor mercadoria")
        if linha["regra_peso"]:
            regras.append("Peso")
        if linha["regra_percentual_frete"]:
            regras.append("% frete/mercadoria")
        if linha["regra_frete_por_kg"]:
            regras.append("R$/kg")
        if linha["regra_valor_frete"]:
            regras.append("Valor frete")
        return "; ".join(regras) if regras else "Sem gatilho"

    resultado["regras_acionadas"] = resultado.apply(regras_acionadas_linha, axis=1)

    existe_dedicado_menor = resultado["frete_dedicado"].notna() & (
        resultado["frete_dedicado"] < resultado["frete_fracionado"]
    )
    cotar_dedicado = resultado["eh_fracionado"] & resultado["bate_regra_cotacao"] & resultado["frete_dedicado"].isna()

    resultado["classificacao"] = np.select(
        [existe_dedicado_menor, cotar_dedicado],
        ["DEDICADO RECOMENDADO", "COTAR DEDICADO"],
        default="MANTER FRACIONADO",
    )

    resultado["economia_potencial"] = 0.0
    resultado.loc[existe_dedicado_menor, "economia_potencial"] = (
        resultado.loc[existe_dedicado_menor, "frete_fracionado"]
        - resultado.loc[existe_dedicado_menor, "frete_dedicado"]
    ).clip(lower=0)
    resultado.loc[cotar_dedicado, "economia_potencial"] = (
        resultado.loc[cotar_dedicado, "frete_fracionado"] * parametros["percentual_economia_estimada"]
    ).clip(lower=0)

    resultado["nivel_oportunidade"] = np.select(
        [
            resultado["classificacao"].eq("DEDICADO RECOMENDADO"),
            resultado["qtd_regras_acionadas"] >= 2,
            resultado["classificacao"].eq("COTAR DEDICADO"),
        ],
        ["ALTA", "ALTA", "MÉDIA"],
        default="BAIXA",
    )

    resultado["oportunidade_dedicado"] = resultado["classificacao"].isin(
        ["DEDICADO RECOMENDADO", "COTAR DEDICADO"]
    )

    return resultado


def sugerir_coluna(colunas, termos):
    """Sugere coluna a partir de uma lista de termos."""
    colunas_lower = {str(c).lower(): c for c in colunas}
    for termo in termos:
        termo_lower = termo.lower()
        for coluna_lower, coluna_original in colunas_lower.items():
            if termo_lower in coluna_lower:
                return coluna_original
    return "Não disponível"


def construir_df_padronizado(df_original, mapa_colunas):
    """Constrói DataFrame com nomes padronizados a partir do mapeamento do usuário."""
    df_padrao = pd.DataFrame(index=df_original.index)

    for destino, origem in mapa_colunas.items():
        if origem and origem != "Não disponível" and origem in df_original.columns:
            if destino == "documento":
                df_padrao[destino] = df_original[origem].fillna("").astype(str)
            else:
                df_padrao[destino] = df_original[origem]
        else:
            df_padrao[destino] = "" if destino not in ["valor_mercadoria", "peso", "frete_fracionado", "frete_dedicado"] else np.nan

    return df_padrao


def preparar_tabela_exibicao(df):
    """Cria cópia com colunas principais em ordem amigável para exibição/download."""
    colunas_preferidas = [
        "documento",
        "data",
        "modalidade",
        "transportadora",
        "origem",
        "destino",
        "valor_mercadoria",
        "peso",
        "frete_fracionado",
        "frete_dedicado",
        "percentual_frete_sobre_mercadoria",
        "frete_fracionado_por_kg",
        "classificacao",
        "nivel_oportunidade",
        "regras_acionadas",
        "economia_potencial",
    ]
    colunas = [c for c in colunas_preferidas if c in df.columns]
    demais = [c for c in df.columns if c not in colunas]
    return df[colunas + demais].copy()


# =============================================================================
# Sidebar de parâmetros configuráveis
# =============================================================================

st.sidebar.title("Parâmetros de Decisão")
st.sidebar.caption("Ajuste os gatilhos operacionais conforme a política logística vigente.")

limite_valor_mercadoria = st.sidebar.number_input(
    "Valor da mercadoria para cotar dedicado (R$)",
    min_value=0.0,
    value=100000.0,
    step=5000.0,
    format="%.2f",
)
limite_peso = st.sidebar.number_input(
    "Peso bruto para cotar dedicado (kg)",
    min_value=0.0,
    value=10000.0,
    step=500.0,
    format="%.2f",
)
limite_percentual_frete = st.sidebar.number_input(
    "Frete / mercadoria para cotar dedicado (%)",
    min_value=0.0,
    value=3.0,
    step=0.1,
    format="%.2f",
) / 100
limite_frete_por_kg = st.sidebar.number_input(
    "Frete fracionado por kg para cotar dedicado (R$/kg)",
    min_value=0.0,
    value=2.50,
    step=0.10,
    format="%.2f",
)
limite_peso_para_frete_kg = st.sidebar.number_input(
    "Peso mínimo para regra R$/kg (kg)",
    min_value=0.0,
    value=2000.0,
    step=100.0,
    format="%.2f",
)
limite_valor_frete = st.sidebar.number_input(
    "Valor do frete fracionado para cotar dedicado (R$)",
    min_value=0.0,
    value=5000.0,
    step=250.0,
    format="%.2f",
)
percentual_economia_estimada = st.sidebar.number_input(
    "Economia estimada quando não há cotação dedicada (%)",
    min_value=0.0,
    max_value=100.0,
    value=15.0,
    step=1.0,
    format="%.2f",
) / 100

parametros = {
    "limite_valor_mercadoria": limite_valor_mercadoria,
    "limite_peso": limite_peso,
    "limite_percentual_frete": limite_percentual_frete,
    "limite_frete_por_kg": limite_frete_por_kg,
    "limite_peso_para_frete_kg": limite_peso_para_frete_kg,
    "limite_valor_frete": limite_valor_frete,
    "percentual_economia_estimada": percentual_economia_estimada,
}

st.sidebar.markdown("---")
st.sidebar.info(
    "Recomendação: use os parâmetros como gatilhos de triagem. "
    "A decisão final deve considerar lead time, janela de coleta, cubagem, rota e disponibilidade de veículos."
)


# =============================================================================
# Abas principais
# =============================================================================

tab_manual, tab_lote = st.tabs(["1. Simulador Manual", "2. Análise em Lote"])


# =============================================================================
# Aba 1: Simulador Manual
# =============================================================================

with tab_manual:
    st.subheader("Simulador Manual de Decisão")
    st.write("Informe os dados do embarque para avaliar se vale manter frete fracionado ou abrir cotação de dedicado.")

    with st.form("form_simulador_manual"):
        col1, col2, col3 = st.columns(3)

        with col1:
            valor_mercadoria = st.number_input(
                "Valor da mercadoria (R$)", min_value=0.0, value=100000.0, step=1000.0, format="%.2f"
            )
            peso = st.number_input(
                "Peso bruto (kg)", min_value=0.0, value=2500.0, step=100.0, format="%.2f"
            )
            frete_fracionado = st.number_input(
                "Valor estimado do frete fracionado (R$)",
                min_value=0.0,
                value=5000.0,
                step=100.0,
                format="%.2f",
            )

        with col2:
            frete_dedicado = st.number_input(
                "Valor cotado/estimado do dedicado (opcional, R$)",
                min_value=0.0,
                value=0.0,
                step=100.0,
                format="%.2f",
                help="Deixe 0,00 quando ainda não houver cotação de dedicado.",
            )
            origem_uf = st.text_input("Origem - UF", value="SP", max_chars=2)
            origem_cidade = st.text_input("Origem - Cidade", value="São Paulo")

        with col3:
            destino_uf = st.text_input("Destino - UF", value="MG", max_chars=2)
            destino_cidade = st.text_input("Destino - Cidade", value="Belo Horizonte")
            transportadora = st.text_input("Transportadora (opcional)", value="")

        simular = st.form_submit_button("Calcular recomendação", type="primary")

    if simular:
        df_manual = pd.DataFrame(
            [
                {
                    "documento": "SIMULAÇÃO MANUAL",
                    "modalidade": "Fracionado",
                    "valor_mercadoria": valor_mercadoria,
                    "peso": peso,
                    "frete_fracionado": frete_fracionado,
                    "frete_dedicado": np.nan if frete_dedicado <= 0 else frete_dedicado,
                    "origem": f"{origem_cidade.strip()} / {origem_uf.strip().upper()}",
                    "destino": f"{destino_cidade.strip()} / {destino_uf.strip().upper()}",
                    "data": datetime.today().strftime("%Y-%m-%d"),
                    "transportadora": transportadora,
                }
            ]
        )

        resultado_manual = aplicar_regras(df_manual, parametros).iloc[0]

        st.markdown("### Resultado da Simulação")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Frete / Mercadoria", formatar_percentual(resultado_manual["percentual_frete_sobre_mercadoria"]))
        m2.metric("R$/kg Fracionado", formatar_moeda(resultado_manual["frete_fracionado_por_kg"]))
        m3.metric("Economia Potencial", formatar_moeda(resultado_manual["economia_potencial"]))
        m4.metric("Nível de Oportunidade", resultado_manual["nivel_oportunidade"])

        classificacao = resultado_manual["classificacao"]
        if classificacao == "DEDICADO RECOMENDADO":
            st.success(
                f"✅ **{classificacao}** — a cotação dedicada informada é menor que o frete fracionado."
            )
        elif classificacao == "COTAR DEDICADO":
            st.warning(
                f"⚠️ **{classificacao}** — o embarque acionou uma ou mais regras e não há valor dedicado informado."
            )
        else:
            st.info(f"ℹ️ **{classificacao}** — não foram identificados gatilhos suficientes para cotar dedicado.")

        st.markdown("#### Detalhes")
        detalhes = pd.DataFrame(
            {
                "Indicador": [
                    "Origem",
                    "Destino",
                    "Transportadora",
                    "Valor da mercadoria",
                    "Peso bruto",
                    "Frete fracionado",
                    "Frete dedicado informado",
                    "Regras acionadas",
                ],
                "Valor": [
                    resultado_manual["origem"],
                    resultado_manual["destino"],
                    resultado_manual["transportadora"] or "Não informado",
                    formatar_moeda(resultado_manual["valor_mercadoria"]),
                    f"{resultado_manual['peso']:,.2f} kg".replace(",", "X").replace(".", ",").replace("X", "."),
                    formatar_moeda(resultado_manual["frete_fracionado"]),
                    "Não informado" if pd.isna(resultado_manual["frete_dedicado"]) else formatar_moeda(resultado_manual["frete_dedicado"]),
                    resultado_manual["regras_acionadas"],
                ],
            }
        )
        st.dataframe(detalhes, use_container_width=True, hide_index=True)

    else:
        st.info("Preencha os campos e clique em **Calcular recomendação** para visualizar a análise.")


# =============================================================================
# Aba 2: Análise em Lote
# =============================================================================

with tab_lote:
    st.subheader("Análise em Lote de Histórico de Embarques")
    st.write(
        "Faça upload de um arquivo Excel ou CSV. Caso os nomes das colunas não estejam padronizados, "
        "use o mapeamento abaixo para indicar os campos corretos."
    )

    arquivo = st.file_uploader(
        "Upload de histórico de embarques (.xlsx, .xls ou .csv)",
        type=["xlsx", "xls", "csv"],
    )

    df_original = carregar_arquivo(arquivo) if arquivo is not None else None

    if df_original is not None and not df_original.empty:
        st.success(f"Arquivo carregado com sucesso: {len(df_original):,} linhas e {len(df_original.columns):,} colunas.".replace(",", "."))

        with st.expander("Pré-visualização do arquivo carregado", expanded=False):
            st.dataframe(df_original.head(50), use_container_width=True)

        colunas_disponiveis = ["Não disponível"] + list(df_original.columns)

        sugestoes = {
            "documento": sugerir_coluna(df_original.columns, ["nf", "nota", "documento", "doc", "cte", "pedido", "shipment"]),
            "modalidade": sugerir_coluna(df_original.columns, ["modalidade", "tipo frete", "tipo_frete", "frete tipo", "serviço", "servico"]),
            "valor_mercadoria": sugerir_coluna(df_original.columns, ["valor mercadoria", "vl mercadoria", "valor nf", "valor_nf", "mercadoria", "invoice"]),
            "peso": sugerir_coluna(df_original.columns, ["peso bruto", "peso", "kg", "weight"]),
            "frete_fracionado": sugerir_coluna(df_original.columns, ["valor frete", "vl frete", "frete", "freight"]),
            "frete_dedicado": sugerir_coluna(df_original.columns, ["dedicado", "cotacao dedicado", "cotação dedicado", "frete dedicado"]),
            "origem": sugerir_coluna(df_original.columns, ["origem", "cidade origem", "uf origem", "origin"]),
            "destino": sugerir_coluna(df_original.columns, ["destino", "cidade destino", "uf destino", "destination"]),
            "data": sugerir_coluna(df_original.columns, ["data", "emissao", "emissão", "dt", "date"]),
            "transportadora": sugerir_coluna(df_original.columns, ["transportadora", "carrier", "transp", "fornecedor"]),
        }

        st.markdown("### Mapeamento de Colunas")
        st.caption("Selecione **Não disponível** quando a informação não existir no arquivo.")

        mapa_colunas = {}
        c1, c2, c3 = st.columns(3)

        campos = [
            ("documento", "Chave NF / Documento"),
            ("modalidade", "Modalidade / Tipo de frete"),
            ("valor_mercadoria", "Valor da mercadoria"),
            ("peso", "Peso bruto"),
            ("frete_fracionado", "Valor do frete fracionado"),
            ("frete_dedicado", "Valor dedicado cotado/estimado"),
            ("origem", "Origem"),
            ("destino", "Destino"),
            ("data", "Data"),
            ("transportadora", "Transportadora"),
        ]

        for i, (campo, rotulo) in enumerate(campos):
            coluna_container = [c1, c2, c3][i % 3]
            with coluna_container:
                indice_sugerido = colunas_disponiveis.index(sugestoes[campo]) if sugestoes[campo] in colunas_disponiveis else 0
                mapa_colunas[campo] = st.selectbox(
                    rotulo,
                    options=colunas_disponiveis,
                    index=indice_sugerido,
                    key=f"mapa_{campo}",
                )

        col_obrig_1, col_obrig_2, col_obrig_3 = st.columns(3)
        with col_obrig_1:
            st.caption("Campo crítico: Valor da mercadoria")
            if mapa_colunas["valor_mercadoria"] == "Não disponível":
                st.warning("Sem esta coluna, a regra de % frete/mercadoria e valor da mercadoria não será efetiva.")
        with col_obrig_2:
            st.caption("Campo crítico: Peso")
            if mapa_colunas["peso"] == "Não disponível":
                st.warning("Sem esta coluna, as regras de peso e R$/kg não serão efetivas.")
        with col_obrig_3:
            st.caption("Campo crítico: Valor do frete")
            if mapa_colunas["frete_fracionado"] == "Não disponível":
                st.error("Sem valor de frete não é possível calcular oportunidades corretamente.")

        analisar = st.button("Aplicar regras no banco", type="primary")

        if analisar:
            df_padrao = construir_df_padronizado(df_original, mapa_colunas)
            df_resultado = aplicar_regras(df_padrao, parametros)
            df_exibicao = preparar_tabela_exibicao(df_resultado)

            df_oportunidades = df_exibicao[df_exibicao["oportunidade_dedicado"] == True].copy()
            df_oportunidades = df_oportunidades.sort_values(
                by=["economia_potencial", "frete_fracionado", "valor_mercadoria"],
                ascending=[False, False, False],
            )

            total_embarques = len(df_resultado)
            fretes_fracionados = int(df_resultado["eh_fracionado"].sum())
            qtd_oportunidades = int(df_resultado["oportunidade_dedicado"].sum())
            gasto_fracionado_oportunidade = float(
                df_resultado.loc[df_resultado["oportunidade_dedicado"], "frete_fracionado"].sum()
            )
            economia_potencial_total = float(df_resultado["economia_potencial"].sum())

            resumo = {
                "total_embarques": total_embarques,
                "fretes_fracionados_analisados": fretes_fracionados,
                "quantidade_com_oportunidade_dedicado": qtd_oportunidades,
                "gasto_total_fracionado_com_oportunidade": gasto_fracionado_oportunidade,
                "economia_potencial_estimada": economia_potencial_total,
                "percentual_economia_sobre_gasto_oportunidade": (
                    economia_potencial_total / gasto_fracionado_oportunidade if gasto_fracionado_oportunidade > 0 else 0
                ),
                "data_processamento": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }

            st.markdown("### KPIs da Análise")
            k1, k2, k3, k4, k5 = st.columns(5)
            k1.metric("Total de embarques", f"{total_embarques:,}".replace(",", "."))
            k2.metric("Fracionados analisados", f"{fretes_fracionados:,}".replace(",", "."))
            k3.metric("Oportunidades", f"{qtd_oportunidades:,}".replace(",", "."))
            k4.metric("Gasto fracionado c/ oportunidade", formatar_moeda(gasto_fracionado_oportunidade))
            k5.metric("Economia potencial", formatar_moeda(economia_potencial_total))

            st.markdown("### Visualizações")
            graf1, graf2 = st.columns(2)

            with graf1:
                df_classificacao = (
                    df_resultado.groupby("classificacao", dropna=False)
                    .size()
                    .reset_index(name="quantidade")
                    .sort_values("quantidade", ascending=False)
                )
                fig_pizza = px.pie(
                    df_classificacao,
                    names="classificacao",
                    values="quantidade",
                    title="Distribuição por recomendação",
                    color="classificacao",
                    color_discrete_map={
                        "DEDICADO RECOMENDADO": PARKER_RED,
                        "COTAR DEDICADO": "#F59E0B",
                        "MANTER FRACIONADO": "#6B7280",
                    },
                    hole=0.35,
                )
                fig_pizza.update_traces(textposition="inside", textinfo="percent+label")
                st.plotly_chart(fig_pizza, use_container_width=True)

            with graf2:
                if not df_oportunidades.empty:
                    df_top = df_oportunidades.head(15).copy()
                    df_top["documento_plot"] = df_top["documento"].replace("", np.nan).fillna("Sem documento")
                    fig_bar = px.bar(
                        df_top.sort_values("economia_potencial", ascending=True),
                        x="economia_potencial",
                        y="documento_plot",
                        orientation="h",
                        title="Top oportunidades por economia potencial",
                        labels={"economia_potencial": "Economia potencial (R$)", "documento_plot": "Documento"},
                        color="nivel_oportunidade",
                        color_discrete_map={"ALTA": PARKER_RED, "MÉDIA": "#F59E0B", "BAIXA": "#6B7280"},
                    )
                    st.plotly_chart(fig_bar, use_container_width=True)
                else:
                    st.info("Nenhuma oportunidade identificada para exibir no gráfico de barras.")

            st.markdown("### Tabela de Oportunidades Priorizadas")
            if df_oportunidades.empty:
                st.info("Nenhuma oportunidade de frete dedicado foi identificada com os parâmetros atuais.")
            else:
                st.dataframe(
                    df_oportunidades,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "valor_mercadoria": st.column_config.NumberColumn("Valor mercadoria", format="R$ %.2f"),
                        "frete_fracionado": st.column_config.NumberColumn("Frete fracionado", format="R$ %.2f"),
                        "frete_dedicado": st.column_config.NumberColumn("Frete dedicado", format="R$ %.2f"),
                        "economia_potencial": st.column_config.NumberColumn("Economia potencial", format="R$ %.2f"),
                        "percentual_frete_sobre_mercadoria": st.column_config.NumberColumn("Frete / mercadoria", format="%.2f%%"),
                        "frete_fracionado_por_kg": st.column_config.NumberColumn("R$/kg", format="R$ %.2f"),
                    },
                )

            st.markdown("### Download dos Resultados")
            excel_bytes = criar_download_excel(df_exibicao, df_oportunidades, resumo)
            csv_oportunidades = df_oportunidades.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")
            csv_completo = df_exibicao.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")

            d1, d2, d3 = st.columns(3)
            with d1:
                st.download_button(
                    "Baixar Excel completo",
                    data=excel_bytes,
                    file_name="parker_analise_frete_dedicado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            with d2:
                st.download_button(
                    "Baixar CSV de oportunidades",
                    data=csv_oportunidades,
                    file_name="parker_oportunidades_frete_dedicado.csv",
                    mime="text/csv",
                )
            with d3:
                st.download_button(
                    "Baixar CSV completo",
                    data=csv_completo,
                    file_name="parker_analise_completa_frete.csv",
                    mime="text/csv",
                )

            with st.expander("Ver análise completa", expanded=False):
                st.dataframe(df_exibicao, use_container_width=True, hide_index=True)

    elif arquivo is not None:
        st.warning("O arquivo foi carregado, mas não possui dados para análise.")
    else:
        st.info("Envie um arquivo para iniciar a análise em lote.")


st.markdown("---")
st.caption(
    "Parker-Hannifin | Ferramenta interna de apoio à decisão logística. "
    "Não substitui validação operacional, contratação formal ou análise de nível de serviço."
)
