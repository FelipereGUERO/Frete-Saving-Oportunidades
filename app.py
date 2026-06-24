"""
app.py

Aplicativo Streamlit para Simulador Manual e Análise em Massa.

Regras monetárias separadas:
- Simulador Manual: regra em R$ / BRL, sem conversão cambial.
- Análise em Massa: valores históricos da base em USD / US$, sem conversão cambial.

A regra operacional de peso >= 1.000 kg é válida em ambas as abas.
"""

from __future__ import annotations

import io
import re
import unicodedata
from typing import Dict, Iterable, Optional, Tuple

import pandas as pd
import streamlit as st


# =============================================================================
# Configuração geral
# =============================================================================

st.set_page_config(
    page_title="Matriz Fixa de Decisão",
    page_icon="📦",
    layout="wide",
)


# =============================================================================
# Constantes de regra — separadas por aba/moeda
# =============================================================================

# Regra operacional comum às duas abas
PESO_MINIMO_KG = 1_000.0

# Simulador Manual: sempre BRL / R$
MANUAL_VALOR_MERCADORIA_MIN_BRL = 50_000.0
MANUAL_FRETE_FRACIONADO_MIN_BRL = 500.0
MANUAL_AVALIACAO_COM_FRETE_MIN_BRL = 250.0

# Análise em Massa: sempre USD / US$ para valores históricos da base
MASSA_VALOR_MERCADORIA_MIN_USD = 50_000.0
MASSA_FRETE_FRACIONADO_MIN_USD = 500.0
MASSA_AVALIACAO_COM_FRETE_MIN_USD = 250.0


# =============================================================================
# Funções utilitárias
# =============================================================================

def format_brl(value: float) -> str:
    """Formata valor em reais, sem qualquer conversão cambial."""
    try:
        value = float(value)
    except Exception:
        value = 0.0
    formatted = f"R$ {value:,.2f}"
    return formatted.replace(",", "X").replace(".", ",").replace("X", ".")


def format_usd(value: float) -> str:
    """Formata valor em dólares históricos da base, sem qualquer conversão cambial."""
    try:
        value = float(value)
    except Exception:
        value = 0.0
    return f"US$ {value:,.2f}"


def format_kg(value: float) -> str:
    try:
        value = float(value)
    except Exception:
        value = 0.0
    formatted = f"{value:,.2f} kg"
    return formatted.replace(",", "X").replace(".", ",").replace("X", ".")


def normalize_text(text: object) -> str:
    """Normaliza nomes de colunas para detecção automática."""
    if text is None:
        return ""
    text = str(text).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return text.strip("_")


def parse_numeric_value(value: object) -> float:
    """Converte valores numéricos em formato BR/US para float.

    Esta função apenas interpreta o número escrito na base; não realiza conversão
    cambial. A moeda aplicável depende da aba:
    - Simulador Manual: BRL / R$.
    - Análise em Massa: USD / US$.
    """
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return 0.0

    text = re.sub(r"[^0-9,\.\-]", "", text)

    if "," in text and "." in text:
        # Formato brasileiro provável: 1.234,56
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        # Formato americano provável: 1,234.56
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")

    try:
        return float(text)
    except Exception:
        return 0.0


def numeric_series(series: pd.Series) -> pd.Series:
    return series.apply(parse_numeric_value).astype(float)


def detect_column(columns: Iterable[str], aliases: Iterable[str]) -> Optional[str]:
    """Detecta automaticamente uma coluna por aliases, sem controles manuais."""
    original_columns = list(columns)
    normalized_to_original: Dict[str, str] = {
        normalize_text(col): col for col in original_columns
    }
    normalized_aliases = [normalize_text(alias) for alias in aliases]

    # Primeiro tenta correspondência exata normalizada
    for alias in normalized_aliases:
        if alias in normalized_to_original:
            return normalized_to_original[alias]

    # Depois tenta alias contido no nome da coluna
    for norm_col, original in normalized_to_original.items():
        for alias in normalized_aliases:
            if alias and alias in norm_col:
                return original

    return None


def read_uploaded_file(uploaded_file) -> pd.DataFrame:
    file_name = uploaded_file.name.lower()
    if file_name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    if file_name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded_file)
    raise ValueError("Formato não suportado. Envie um arquivo CSV, XLSX ou XLS.")


def build_download_excel(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="analise_massa")
    return output.getvalue()


# =============================================================================
# Lógica de decisão
# =============================================================================

def avaliar_simulador_manual_brl(
    valor_mercadoria_brl: float,
    frete_fracionado_brl: float,
    avaliacao_com_frete_brl: float,
    peso_kg: float,
    frete_atual_brl: float,
) -> Dict[str, object]:
    """Matriz fixa do Simulador Manual.

    Todos os thresholds monetários são aplicados em R$ / BRL.
    Não há tratamento como USD e não há conversão cambial.
    """
    criterio_valor = valor_mercadoria_brl >= MANUAL_VALOR_MERCADORIA_MIN_BRL
    criterio_frete = frete_fracionado_brl >= MANUAL_FRETE_FRACIONADO_MIN_BRL
    criterio_avaliacao = avaliacao_com_frete_brl >= MANUAL_AVALIACAO_COM_FRETE_MIN_BRL
    criterio_peso = peso_kg >= PESO_MINIMO_KG

    custo_estimado_brl = frete_fracionado_brl + avaliacao_com_frete_brl
    economia_brl = frete_atual_brl - custo_estimado_brl

    aprovado = all([
        criterio_valor,
        criterio_frete,
        criterio_avaliacao,
        criterio_peso,
    ])

    return {
        "criterio_valor": criterio_valor,
        "criterio_frete": criterio_frete,
        "criterio_avaliacao": criterio_avaliacao,
        "criterio_peso": criterio_peso,
        "custo_estimado_brl": custo_estimado_brl,
        "economia_brl": economia_brl,
        "aprovado": aprovado,
    }


def avaliar_linha_massa_usd(row: pd.Series) -> Dict[str, object]:
    """Matriz fixa da Análise em Massa.

    Todos os thresholds monetários são aplicados em USD / US$ porque a base
    histórica permanece em USD / US$. Não há conversão cambial.
    """
    valor_mercadoria_usd = float(row.get("__valor_mercadoria_usd", 0.0))
    frete_fracionado_usd = float(row.get("__frete_fracionado_usd", 0.0))
    avaliacao_com_frete_usd = float(row.get("__avaliacao_com_frete_usd", 0.0))
    peso_kg = float(row.get("__peso_kg", 0.0))
    frete_atual_usd = float(row.get("__frete_atual_usd", 0.0))

    criterio_valor = valor_mercadoria_usd >= MASSA_VALOR_MERCADORIA_MIN_USD
    criterio_frete = frete_fracionado_usd >= MASSA_FRETE_FRACIONADO_MIN_USD
    criterio_avaliacao = avaliacao_com_frete_usd >= MASSA_AVALIACAO_COM_FRETE_MIN_USD
    criterio_peso = peso_kg >= PESO_MINIMO_KG

    custo_estimado_usd = frete_fracionado_usd + avaliacao_com_frete_usd
    economia_usd = frete_atual_usd - custo_estimado_usd

    aprovado = all([
        criterio_valor,
        criterio_frete,
        criterio_avaliacao,
        criterio_peso,
    ])

    return {
        "Valor mercadoria >= US$ 50.000": criterio_valor,
        "Frete fracionado estimado >= US$ 500": criterio_frete,
        "Avaliação com frete >= US$ 250": criterio_avaliacao,
        "Peso >= 1.000 kg": criterio_peso,
        "Custo estimado US$": custo_estimado_usd,
        "Economia US$": economia_usd,
        "Decisão matriz fixa": "Aprovado" if aprovado else "Não aprovado",
    }


# =============================================================================
# Layout
# =============================================================================

st.title("📦 Matriz Fixa de Decisão")
st.caption(
    "Simulador Manual com regra em R$ / BRL e Análise em Massa com valores "
    "históricos da base em USD / US$. Não há conversão cambial entre as abas."
)

with st.expander("Regras monetárias e operacionais", expanded=True):
    st.markdown(
        f"""
        **Simulador Manual: regra em R$**
        - Valor da mercadoria >= **{format_brl(MANUAL_VALOR_MERCADORIA_MIN_BRL)}**
        - Frete fracionado estimado >= **{format_brl(MANUAL_FRETE_FRACIONADO_MIN_BRL)}**
        - Avaliação com frete >= **{format_brl(MANUAL_AVALIACAO_COM_FRETE_MIN_BRL)}**
        - Economia calculada e exibida em **R$**
        - Sem tratar valores como USD e sem conversão cambial

        **Análise em Massa: valores da base em USD/US$**
        - Os valores históricos da planilha permanecem em **USD / US$**
        - Thresholds monetários aplicados à base em **USD / US$**
        - Economia calculada e exibida em **US$**
        - Sem conversão cambial

        **Regra operacional comum às duas abas**
        - Peso >= **{format_kg(PESO_MINIMO_KG)}**
        """
    )

aba_manual, aba_massa = st.tabs(["Simulador Manual", "Análise em Massa"])


# =============================================================================
# Aba: Simulador Manual
# =============================================================================

with aba_manual:
    st.subheader("Simulador Manual: regra em R$")
    st.write(
        "Preencha os campos abaixo em **R$ / BRL**. A matriz fixa do Simulador "
        "Manual não interpreta estes valores como USD e não realiza conversão cambial."
    )

    col1, col2 = st.columns(2)

    with col1:
        valor_mercadoria_brl = st.number_input(
            "Valor da mercadoria (R$)",
            min_value=0.0,
            value=50_000.0,
            step=1_000.0,
            format="%.2f",
            help="Simulador Manual: regra em R$. Threshold fixo: R$ 50.000.",
        )
        frete_fracionado_brl = st.number_input(
            "Frete fracionado estimado (R$)",
            min_value=0.0,
            value=500.0,
            step=50.0,
            format="%.2f",
            help="Simulador Manual: regra em R$. Threshold fixo: R$ 500.",
        )
        avaliacao_com_frete_brl = st.number_input(
            "Avaliação com frete (R$)",
            min_value=0.0,
            value=250.0,
            step=50.0,
            format="%.2f",
            help="Simulador Manual: regra em R$. Threshold fixo: R$ 250.",
        )

    with col2:
        peso_kg = st.number_input(
            "Peso (kg)",
            min_value=0.0,
            value=1_000.0,
            step=100.0,
            format="%.2f",
            help="Regra operacional válida em ambas as abas: peso >= 1.000 kg.",
        )
        frete_atual_brl = st.number_input(
            "Frete atual / referência (R$)",
            min_value=0.0,
            value=1_500.0,
            step=100.0,
            format="%.2f",
            help="Usado apenas para calcular a economia estimada em R$.",
        )

    resultado_manual = avaliar_simulador_manual_brl(
        valor_mercadoria_brl=valor_mercadoria_brl,
        frete_fracionado_brl=frete_fracionado_brl,
        avaliacao_com_frete_brl=avaliacao_com_frete_brl,
        peso_kg=peso_kg,
        frete_atual_brl=frete_atual_brl,
    )

    st.divider()

    metric_col1, metric_col2, metric_col3 = st.columns(3)
    metric_col1.metric(
        "Custo estimado (R$)",
        format_brl(resultado_manual["custo_estimado_brl"]),
    )
    metric_col2.metric(
        "Economia estimada (R$)",
        format_brl(resultado_manual["economia_brl"]),
    )
    metric_col3.metric(
        "Peso informado",
        format_kg(peso_kg),
    )

    criterios_manual = pd.DataFrame(
        [
            {
                "Critério": "Valor da mercadoria >= R$ 50.000",
                "Valor informado": format_brl(valor_mercadoria_brl),
                "Resultado": "Atendido" if resultado_manual["criterio_valor"] else "Não atendido",
            },
            {
                "Critério": "Frete fracionado estimado >= R$ 500",
                "Valor informado": format_brl(frete_fracionado_brl),
                "Resultado": "Atendido" if resultado_manual["criterio_frete"] else "Não atendido",
            },
            {
                "Critério": "Avaliação com frete >= R$ 250",
                "Valor informado": format_brl(avaliacao_com_frete_brl),
                "Resultado": "Atendido" if resultado_manual["criterio_avaliacao"] else "Não atendido",
            },
            {
                "Critério": "Peso >= 1.000 kg",
                "Valor informado": format_kg(peso_kg),
                "Resultado": "Atendido" if resultado_manual["criterio_peso"] else "Não atendido",
            },
        ]
    )

    st.markdown("#### Resultado da matriz fixa — Simulador Manual: regra em R$")
    st.dataframe(criterios_manual, use_container_width=True, hide_index=True)

    if resultado_manual["aprovado"]:
        st.success("Decisão matriz fixa: Aprovado no Simulador Manual em R$.")
    else:
        st.warning("Decisão matriz fixa: Não aprovado no Simulador Manual em R$.")


# =============================================================================
# Aba: Análise em Massa
# =============================================================================

with aba_massa:
    st.subheader("Análise em Massa: valores da base em USD/US$")
    st.write(
        "Envie uma planilha CSV, XLSX ou XLS. Os valores históricos da base são "
        "mantidos em **USD / US$** e os thresholds monetários aplicados à base "
        "também são em **USD / US$**. Não há conversão cambial."
    )

    st.info(
        "A detecção das colunas é automática, sem controles de mapeamento manual. "
        "Colunas esperadas: valor da mercadoria, frete fracionado estimado, "
        "avaliação com frete, peso e, opcionalmente, frete atual/referência."
    )

    uploaded_file = st.file_uploader(
        "Arquivo para análise em massa",
        type=["csv", "xlsx", "xls"],
    )

    if uploaded_file is not None:
        try:
            df_original = read_uploaded_file(uploaded_file)
            df = df_original.copy()
        except Exception as exc:
            st.error(f"Não foi possível ler o arquivo: {exc}")
            st.stop()

        if df.empty:
            st.warning("O arquivo enviado não possui linhas para análise.")
            st.stop()

        aliases = {
            "valor_mercadoria": [
                "valor_mercadoria",
                "valor da mercadoria",
                "valor mercadoria",
                "mercadoria",
                "invoice value",
                "goods value",
                "valor invoice",
                "vl mercadoria",
            ],
            "frete_fracionado": [
                "frete_fracionado_estimado",
                "frete fracionado estimado",
                "frete fracionado",
                "frete estimado",
                "estimated freight",
                "fractional freight",
            ],
            "avaliacao_com_frete": [
                "avaliacao_com_frete",
                "avaliação com frete",
                "avaliacao com frete",
                "avaliacao frete",
                "assessment freight",
                "evaluation freight",
            ],
            "peso": [
                "peso",
                "peso_kg",
                "peso kg",
                "gross weight",
                "weight",
                "kg",
            ],
            "frete_atual": [
                "frete_atual",
                "frete atual",
                "frete referencia",
                "frete referência",
                "freight current",
                "current freight",
                "frete pago",
            ],
        }

        col_valor = detect_column(df.columns, aliases["valor_mercadoria"])
        col_frete = detect_column(df.columns, aliases["frete_fracionado"])
        col_avaliacao = detect_column(df.columns, aliases["avaliacao_com_frete"])
        col_peso = detect_column(df.columns, aliases["peso"])
        col_frete_atual = detect_column(df.columns, aliases["frete_atual"])

        detected = pd.DataFrame(
            [
                {"Campo": "Valor da mercadoria — USD/US$", "Coluna detectada": col_valor or "Não detectada"},
                {"Campo": "Frete fracionado estimado — USD/US$", "Coluna detectada": col_frete or "Não detectada"},
                {"Campo": "Avaliação com frete — USD/US$", "Coluna detectada": col_avaliacao or "Não detectada"},
                {"Campo": "Peso — kg", "Coluna detectada": col_peso or "Não detectada"},
                {"Campo": "Frete atual/referência — USD/US$", "Coluna detectada": col_frete_atual or "Não detectada"},
            ]
        )

        st.markdown("#### Colunas detectadas automaticamente")
        st.dataframe(detected, use_container_width=True, hide_index=True)

        required_missing = []
        if col_valor is None:
            required_missing.append("valor da mercadoria")
        if col_frete is None:
            required_missing.append("frete fracionado estimado")
        if col_avaliacao is None:
            required_missing.append("avaliação com frete")
        if col_peso is None:
            required_missing.append("peso")

        if required_missing:
            st.error(
                "Não foi possível executar a análise porque a detecção automática "
                "não encontrou as seguintes colunas obrigatórias: "
                + ", ".join(required_missing)
                + ". Renomeie as colunas na planilha e envie novamente."
            )
            st.stop()

        df["__valor_mercadoria_usd"] = numeric_series(df[col_valor])
        df["__frete_fracionado_usd"] = numeric_series(df[col_frete])
        df["__avaliacao_com_frete_usd"] = numeric_series(df[col_avaliacao])
        df["__peso_kg"] = numeric_series(df[col_peso])
        df["__frete_atual_usd"] = numeric_series(df[col_frete_atual]) if col_frete_atual else 0.0

        resultados = df.apply(avaliar_linha_massa_usd, axis=1, result_type="expand")

        df_resultado = pd.concat(
            [
                df_original.reset_index(drop=True),
                resultados.reset_index(drop=True),
            ],
            axis=1,
        )

        total_linhas = len(df_resultado)
        total_aprovados = int((df_resultado["Decisão matriz fixa"] == "Aprovado").sum())
        total_nao_aprovados = total_linhas - total_aprovados
        economia_total_usd = float(df_resultado["Economia US$"].sum())

        st.divider()
        st.markdown("#### Resumo — Análise em Massa: valores da base em USD/US$")

        col_a, col_b, col_c, col_d = st.columns(4)
        col_a.metric("Linhas analisadas", f"{total_linhas:,}".replace(",", "."))
        col_b.metric("Aprovados", f"{total_aprovados:,}".replace(",", "."))
        col_c.metric("Não aprovados", f"{total_nao_aprovados:,}".replace(",", "."))
        col_d.metric("Economia total (US$)", format_usd(economia_total_usd))

        st.markdown(
            "#### Resultado detalhado da matriz fixa — Análise em Massa: valores da base em USD/US$"
        )
        st.dataframe(df_resultado, use_container_width=True, hide_index=True)

        excel_bytes = build_download_excel(df_resultado)
        st.download_button(
            label="Baixar resultado em Excel",
            data=excel_bytes,
            file_name="resultado_analise_massa_usd.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.caption(
            "Aguardando envio de arquivo. A Análise em Massa mantém os valores "
            "históricos da planilha em USD/US$ e aplica a regra de peso >= 1.000 kg."
        )
