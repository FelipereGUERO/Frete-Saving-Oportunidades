# app.py
# =============================================================================
# Aplicativo Streamlit - Análise de Oportunidade de Frete Dedicado
# Branding Parker-Hannifin
#
# Instalação:
#   pip install streamlit pandas numpy plotly openpyxl xlsxwriter
#
# Execução:
#   streamlit run app.py
#
# Logo Parker:
#   crie a pasta assets na raiz do projeto e salve a imagem como:
#   assets/parker_logo.png
# =============================================================================

import base64
import io
import os
import re
import unicodedata
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


# =============================================================================
# Configuração visual, branding e regras fixas
# =============================================================================

st.set_page_config(
    page_title='Parker-Hannifin | Oportunidade de Frete Dedicado',
    page_icon='🚚',
    layout='wide',
)

PARKER_RED = '#C8102E'
PARKER_DARK = '#111827'
PARKER_GRAY = '#F3F4F6'
PARKER_MID_GRAY = '#6B7280'
PARKER_LIGHT_RED = '#FDECEF'
PARKER_YELLOW = '#F59E0B'
PARKER_GREEN = '#16A34A'
PARKER_WHITE = '#FFFFFF'

# Regras fixas conforme matriz operacional Parker-Hannifin.
# Na análise em lote, os valores monetários da base permanecem em dólar/US$ conforme solicitado.
# No simulador manual, os mesmos cálculos são apresentados em moeda brasileira R$.
LIMITE_PESO_COTAR_OBRIGATORIO = 1000.0
LIMITE_FRETE_COTAR_OBRIGATORIO = 500.0
LIMITE_PESO_VALOR_COTAR = 1000.0
LIMITE_VALOR_MERCADORIA_COTAR = 50000.0
LIMITE_FRETE_VENDA_COTAR = 0.08
LIMITE_PESO_AVALIAR_MIN = 500.0
LIMITE_PESO_AVALIAR_MAX = 999.0
LIMITE_FRETE_AVALIAR = 250.0

TAXA_ECONOMIA_SIMULADA_ATIVA = True
TAXA_ECONOMIA_SIMULADA = 0.15

LIMITE_FRETE_KG_BAIXO_CARGA_PESADA = 0.35
LIMITE_FRETE_KG_MUITO_BAIXO = 0.20

st.markdown(
    f'''
    <style>
        .stApp {{ background-color: {PARKER_WHITE}; }}
        .block-container {{ padding-top: 1.1rem; padding-bottom: 2rem; }}
        .parker-header {{
            background: linear-gradient(120deg, {PARKER_DARK} 0%, #1F2937 42%, {PARKER_RED} 100%);
            padding: 1.25rem 1.45rem;
            border-radius: 16px;
            color: white;
            margin-bottom: 1.2rem;
            box-shadow: 0 8px 22px rgba(17,24,39,0.18);
            border: 1px solid rgba(255,255,255,0.12);
            min-height: 74px;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }}
        .parker-header h1 {{
            margin: 0;
            font-size: clamp(1.45rem, 2.4vw, 2.0rem);
            font-weight: 850;
            letter-spacing: -0.03em;
        }}
        .parker-header p {{ margin: 0.35rem 0 0 0; font-size: 1.0rem; opacity: 0.94; }}
        .logo-panel {{
            min-height: 74px;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 0.35rem 0.25rem;
        }}
        .parker-logo-img {{
            max-width: 180px;
            width: 100%;
            max-height: 65px;
            object-fit: contain;
            display: block;
        }}
        .logo-fallback {{
            height: 65px;
            max-width: 180px;
            width: 100%;
            border: 1px solid #E5E7EB;
            border-radius: 14px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: {PARKER_RED};
            font-weight: 900;
            background: {PARKER_WHITE};
            box-shadow: 0 2px 8px rgba(17,24,39,0.05);
            text-align: center;
            line-height: 1.1;
        }}
        .rule-card {{
            border-radius: 14px;
            padding: 1rem;
            border: 1px solid #E5E7EB;
            background-color: {PARKER_WHITE};
            min-height: 178px;
            box-shadow: 0 2px 10px rgba(17,24,39,0.05);
        }}
        .rule-red {{ border-top: 5px solid {PARKER_RED}; }}
        .rule-yellow {{ border-top: 5px solid {PARKER_YELLOW}; }}
        .rule-green {{ border-top: 5px solid {PARKER_GREEN}; }}
        .kpi-card {{
            background: {PARKER_WHITE};
            border: 1px solid #E5E7EB;
            border-radius: 14px;
            padding: 1rem;
            box-shadow: 0 2px 10px rgba(17,24,39,0.05);
            height: 100%;
        }}
        .kpi-label {{
            color: {PARKER_MID_GRAY};
            font-size: 0.78rem;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }}
        .kpi-value {{ color: {PARKER_DARK}; font-size: 1.42rem; font-weight: 850; margin-top: 0.25rem; }}
        .kpi-subtitle {{ color: {PARKER_MID_GRAY}; font-size: 0.82rem; margin-top: 0.25rem; }}
        .badge-red, .badge-yellow, .badge-green {{
            display: inline-block;
            padding: 0.35rem 0.6rem;
            border-radius: 999px;
            font-weight: 800;
        }}
        .badge-red {{ background: {PARKER_LIGHT_RED}; color: {PARKER_RED}; border: 1px solid #F8C7D0; }}
        .badge-yellow {{ background: #FFF7ED; color: #B45309; border: 1px solid #FED7AA; }}
        .badge-green {{ background: #ECFDF5; color: #047857; border: 1px solid #A7F3D0; }}
        div[data-testid='stMetricValue'] {{ color: {PARKER_DARK}; font-weight: 800; }}
        section[data-testid='stSidebar'] {{ background: {PARKER_GRAY}; }}
        section[data-testid='stSidebar'] h1,
        section[data-testid='stSidebar'] h2,
        section[data-testid='stSidebar'] h3 {{ color: {PARKER_RED}; }}
        .stButton > button,
        .stDownloadButton > button {{
            border-radius: 10px;
            border: 1px solid {PARKER_RED};
            color: {PARKER_RED};
            font-weight: 800;
        }}
        .stButton > button[kind='primary'],
        .stFormSubmitButton > button[kind='primary'] {{
            background: {PARKER_RED};
            border-color: {PARKER_RED};
            color: white;
            font-weight: 850;
        }}
        .stTabs [data-baseweb='tab-highlight'] {{ background-color: {PARKER_RED}; }}
        @media (max-width: 900px) {{
            .logo-panel {{ min-height: 55px; justify-content: flex-start; }}
            .parker-logo-img {{ max-width: 150px; max-height: 55px; }}
            .logo-fallback {{ height: 55px; max-width: 150px; }}
            .parker-header {{ min-height: auto; }}
        }}
    </style>
    ''',
    unsafe_allow_html=True,
)


# =============================================================================
# Funções utilitárias
# =============================================================================

def formatar_moeda(valor, prefixo='USD'):
    '''Formata número como moeda em padrão brasileiro com prefixo informado.'''
    try:
        if pd.isna(valor):
            return f'{prefixo} 0,00'
        valor = float(valor)
        return f'{prefixo} {valor:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    except Exception:
        return f'{prefixo} 0,00'


def formatar_percentual(valor):
    '''Formata número decimal como percentual brasileiro.'''
    try:
        if pd.isna(valor):
            return 'N/D'
        return f'{float(valor) * 100:,.2f}%'.replace(',', 'X').replace('.', ',').replace('X', '.')
    except Exception:
        return 'N/D'


def formatar_numero(valor, sufixo=''):
    '''Formata número em padrão brasileiro.'''
    try:
        if pd.isna(valor):
            return 'N/D'
        texto = f'{float(valor):,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
        return f'{texto}{sufixo}'
    except Exception:
        return 'N/D'


def converter_numero(valor):
    '''Converte números em formatos brasileiros e internacionais para float.'''
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
    if texto == '' or texto.lower() in {'nan', 'none', 'null', 'n/a', 'na', '-'}:
        return np.nan

    texto = texto.replace('R$', '').replace('US$', '').replace('USD', '').replace('$', '').replace(' ', '')
    texto = re.sub(r'[^0-9,\.\-]', '', texto)
    if texto in {'', '-', ',', '.'}:
        return np.nan

    try:
        if ',' in texto and '.' in texto:
            if texto.rfind(',') > texto.rfind('.'):
                texto = texto.replace('.', '').replace(',', '.')
            else:
                texto = texto.replace(',', '')
        elif ',' in texto:
            texto = texto.replace('.', '').replace(',', '.')
        elif texto.count('.') > 1:
            partes = texto.split('.')
            texto = ''.join(partes[:-1]) + '.' + partes[-1]
        return float(texto)
    except Exception:
        return np.nan


def carregar_arquivo(arquivo):
    '''Carrega arquivo .xlsx, .xls ou .csv para DataFrame preservando texto inicialmente.'''
    if arquivo is None:
        return None

    nome = arquivo.name.lower()
    conteudo = arquivo.getvalue()
    try:
        if nome.endswith('.csv'):
            ultimo_erro = None
            for encoding in ['utf-8-sig', 'utf-8', 'latin1', 'cp1252']:
                try:
                    return pd.read_csv(io.BytesIO(conteudo), dtype=str, sep=None, engine='python', encoding=encoding)
                except Exception as erro:
                    ultimo_erro = erro
            raise ultimo_erro
        if nome.endswith('.xlsx'):
            return pd.read_excel(io.BytesIO(conteudo), dtype=str, engine='openpyxl')
        if nome.endswith('.xls'):
            return pd.read_excel(io.BytesIO(conteudo), dtype=str)
        st.error('Formato não suportado. Envie arquivo .xlsx, .xls ou .csv.')
        return None
    except Exception as erro:
        st.error(
            'Não foi possível carregar o arquivo. Verifique se o arquivo está íntegro e se as dependências estão instaladas. '
            f'Detalhe técnico: {erro}'
        )
        return None


def normalizar_nome_coluna(texto):
    '''Normaliza nomes de colunas para detecção automática.'''
    texto = str(texto).lower().strip()
    texto = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('ascii')
    texto = re.sub(r'[^a-z0-9]+', ' ', texto)
    return re.sub(r'\s+', ' ', texto).strip()


def sugerir_coluna(colunas, termos):
    '''Sugere coluna a partir de uma lista de termos prováveis.'''
    colunas_norm = {normalizar_nome_coluna(c): c for c in colunas}
    termos_norm = [normalizar_nome_coluna(t) for t in termos]
    for termo in termos_norm:
        for coluna_norm, coluna_original in colunas_norm.items():
            if termo and termo in coluna_norm:
                return coluna_original
    return 'Não disponível'


def detectar_colunas_automaticamente(df_original):
    '''Detecta automaticamente as colunas esperadas, sem seletores manuais na interface.'''
    return {
        'documento': sugerir_coluna(df_original.columns, ['nf', 'nota', 'documento', 'doc', 'cte', 'pedido', 'shipment', 'invoice']),
        'modalidade': sugerir_coluna(df_original.columns, ['modalidade', 'tipo frete', 'tipo_frete', 'frete tipo', 'servico', 'serviço', 'mode', 'modal']),
        'valor_mercadoria': sugerir_coluna(df_original.columns, ['valor mercadoria', 'vl mercadoria', 'valor nf', 'valor_nf', 'mercadoria', 'venda', 'sales', 'invoice value', 'goods value']),
        'peso': sugerir_coluna(df_original.columns, ['peso bruto', 'peso', 'kg', 'weight', 'gross weight']),
        'frete_fracionado': sugerir_coluna(df_original.columns, ['valor frete', 'vl frete', 'frete fracionado', 'frete estimado', 'frete', 'freight', 'estimated freight']),
        'frete_dedicado': sugerir_coluna(df_original.columns, ['dedicado', 'cotacao dedicado', 'cotação dedicado', 'frete dedicado', 'dedicated freight', 'ftl quote']),
        'origem': sugerir_coluna(df_original.columns, ['origem', 'cidade origem', 'uf origem', 'origin', 'ship from']),
        'destino': sugerir_coluna(df_original.columns, ['destino', 'cidade destino', 'uf destino', 'destination', 'ship to']),
        'estado': sugerir_coluna(df_original.columns, ['estado', 'uf', 'state', 'province']),
        'data': sugerir_coluna(df_original.columns, ['data', 'emissao', 'emissão', 'dt', 'date', 'ship date']),
        'transportadora': sugerir_coluna(df_original.columns, ['transportadora', 'carrier', 'transp', 'fornecedor', 'vendor']),
        'divisao': sugerir_coluna(df_original.columns, ['divisão', 'divisao', 'division', 'business unit', 'bu']),
        'operacao_direcao': sugerir_coluna(df_original.columns, ['operação', 'operacao', 'direção', 'direcao', 'direction', 'inbound', 'outbound', 'tipo operação']),
        'cliente': sugerir_coluna(df_original.columns, ['cliente', 'customer', 'account', 'oem']),
        'observacoes': sugerir_coluna(df_original.columns, ['observação', 'observacao', 'obs', 'remarks', 'notes', 'comentario', 'comentário']),
    }


def extrair_estado(texto):
    '''Extrai UF/estado provável a partir de texto de origem/destino.'''
    if pd.isna(texto):
        return ''
    texto = str(texto).strip().upper()
    if texto == '':
        return ''
    match_barra = re.search(r'/\s*([A-Z]{2})\b', texto)
    if match_barra:
        return match_barra.group(1)
    match_final = re.search(r'\b([A-Z]{2})\b$', texto)
    if match_final:
        return match_final.group(1)
    return ''


def normalizar_texto_busca(texto):
    '''Normaliza texto para busca operacional simples.'''
    if pd.isna(texto):
        return ''
    texto = str(texto).lower().strip()
    texto = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('ascii')
    return re.sub(r'\s+', ' ', texto)


def identificar_rota_santos(df):
    '''Identifica registros com origem ou destino contendo Santos/SP ou Santos para exclusão em massa.'''
    origem = df['origem'].fillna('').astype(str).apply(normalizar_texto_busca) if 'origem' in df.columns else pd.Series('', index=df.index)
    destino = df['destino'].fillna('').astype(str).apply(normalizar_texto_busca) if 'destino' in df.columns else pd.Series('', index=df.index)
    padrao_santos = r'(^|[^a-z])santos([^a-z]|$)|santos\s*/\s*sp'
    return origem.str.contains(padrao_santos, regex=True, na=False) | destino.str.contains(padrao_santos, regex=True, na=False)


def construir_df_padronizado(df_original, mapa_colunas):
    '''Constrói DataFrame com nomes padronizados a partir do mapeamento automático.'''
    campos_texto = [
        'documento', 'modalidade', 'origem', 'destino', 'estado', 'data', 'transportadora',
        'divisao', 'operacao_direcao', 'cliente', 'observacoes'
    ]
    campos_numericos = ['valor_mercadoria', 'peso', 'frete_fracionado', 'frete_dedicado']
    df_padrao = pd.DataFrame(index=df_original.index)

    for destino, origem in mapa_colunas.items():
        if origem and origem != 'Não disponível' and origem in df_original.columns:
            df_padrao[destino] = df_original[origem].fillna('').astype(str) if destino in campos_texto else df_original[origem]
        else:
            df_padrao[destino] = '' if destino in campos_texto else np.nan

    for campo in campos_texto + campos_numericos:
        if campo not in df_padrao.columns:
            df_padrao[campo] = '' if campo in campos_texto else np.nan
    return df_padrao


def identificar_modal_estimado(resultado):
    '''Identifica fretes provavelmente dedicados por heurística operacional.'''
    palavras_dedicado = [
        'dedicado', 'especial', 'lotação', 'lotacao', 'exclusivo', 'exclusiva', 'carreta',
        'truck', 'toco', 'veículo dedicado', 'veiculo dedicado', 'ftl', 'carga fechada'
    ]
    palavras_fracionado = ['fracionado', 'fracionada', 'ltl', 'consolidado', 'consolidada', 'redespacho']
    texto_base = (
        resultado['modalidade'].fillna('').astype(str) + ' ' +
        resultado['transportadora'].fillna('').astype(str) + ' ' +
        resultado['observacoes'].fillna('').astype(str)
    ).str.lower()

    tem_dedicado = texto_base.apply(lambda x: any(p in x for p in palavras_dedicado))
    tem_fracionado = texto_base.apply(lambda x: any(p in x for p in palavras_fracionado))
    carga_pesada_padrao_fechado = (
        (resultado['peso'] >= LIMITE_PESO_COTAR_OBRIGATORIO) &
        resultado['frete_fracionado_por_kg'].notna() &
        (resultado['frete_fracionado_por_kg'] > 0) &
        (resultado['frete_fracionado_por_kg'] <= LIMITE_FRETE_KG_BAIXO_CARGA_PESADA)
    )
    frete_muito_baixo_carga_pesada = (
        (resultado['peso'] >= 2000) &
        resultado['frete_fracionado_por_kg'].notna() &
        (resultado['frete_fracionado_por_kg'] > 0) &
        (resultado['frete_fracionado_por_kg'] <= LIMITE_FRETE_KG_MUITO_BAIXO)
    )

    provavel_dedicado = tem_dedicado | carga_pesada_padrao_fechado | frete_muito_baixo_carga_pesada
    status = np.where(provavel_dedicado, 'Provável dedicado', 'Provável fracionado')

    explicacoes = []
    for idx in resultado.index:
        motivos = []
        if bool(tem_dedicado.loc[idx]):
            motivos.append('palavra-chave de dedicado/modal especial')
        if bool(carga_pesada_padrao_fechado.loc[idx]):
            motivos.append('peso alto e R$/kg baixo com padrão de carga fechada')
        if bool(frete_muito_baixo_carga_pesada.loc[idx]):
            motivos.append('R$/kg muito baixo para carga pesada')
        if not motivos and bool(tem_fracionado.loc[idx]):
            motivos.append('palavra-chave de fracionado/LTL')
        if not motivos:
            motivos.append('sem indício de dedicado; assumido como provável fracionado')
        explicacoes.append('; '.join(motivos))
    return status, explicacoes


def aplicar_regras(df):
    '''Aplica a matriz fixa Parker-Hannifin para cotação dedicada.'''
    resultado = df.copy()
    colunas_padrao = [
        'documento', 'modalidade', 'valor_mercadoria', 'peso', 'frete_fracionado', 'frete_dedicado',
        'origem', 'destino', 'estado', 'data', 'transportadora', 'divisao', 'operacao_direcao',
        'cliente', 'observacoes'
    ]
    for coluna in colunas_padrao:
        if coluna not in resultado.columns:
            resultado[coluna] = '' if coluna not in ['valor_mercadoria', 'peso', 'frete_fracionado', 'frete_dedicado'] else np.nan

    for coluna in ['documento', 'modalidade', 'origem', 'destino', 'estado', 'data', 'transportadora', 'divisao', 'operacao_direcao', 'cliente', 'observacoes']:
        resultado[coluna] = resultado[coluna].fillna('').astype(str)
    for coluna in ['valor_mercadoria', 'peso', 'frete_fracionado', 'frete_dedicado']:
        resultado[coluna] = converter_numero(resultado[coluna])
        resultado[coluna] = resultado[coluna].where(resultado[coluna] > 0, np.nan)

    resultado['percentual_frete_sobre_venda'] = np.where(
        resultado['frete_fracionado'].notna() & resultado['valor_mercadoria'].notna() & (resultado['valor_mercadoria'] > 0),
        resultado['frete_fracionado'] / resultado['valor_mercadoria'],
        np.nan,
    )
    resultado['frete_fracionado_por_kg'] = np.where(
        resultado['frete_fracionado'].notna() & resultado['peso'].notna() & (resultado['peso'] > 0),
        resultado['frete_fracionado'] / resultado['peso'],
        np.nan,
    )

    texto_excecao = (
        resultado['cliente'].fillna('').astype(str) + ' ' +
        resultado['observacoes'].fillna('').astype(str) + ' ' +
        resultado['modalidade'].fillna('').astype(str)
    ).str.lower()
    palavras_excecao = [
        'estratégico', 'estrategico', 'oem', 'crítica', 'critica', 'urgente', 'emergencial',
        'emergência', 'emergencia', 'priority', 'expedite', 'critical', 'urgent', 'vip', 'parada de linha'
    ]
    resultado['excecao_operacional'] = texto_excecao.apply(lambda x: any(p in x for p in palavras_excecao))

    cond_peso_vermelho = resultado['peso'].notna() & (resultado['peso'] >= LIMITE_PESO_COTAR_OBRIGATORIO)
    cond_frete_vermelho = resultado['frete_fracionado'].notna() & (resultado['frete_fracionado'] >= LIMITE_FRETE_COTAR_OBRIGATORIO)
    cond_peso_valor_vermelho = (
        resultado['peso'].notna() & resultado['valor_mercadoria'].notna() &
        (resultado['peso'] >= LIMITE_PESO_VALOR_COTAR) &
        (resultado['valor_mercadoria'] >= LIMITE_VALOR_MERCADORIA_COTAR)
    )
    cond_frete_venda_vermelho = resultado['percentual_frete_sobre_venda'].notna() & (resultado['percentual_frete_sobre_venda'] >= LIMITE_FRETE_VENDA_COTAR)
    cond_peso_amarelo = resultado['peso'].notna() & (resultado['peso'] >= LIMITE_PESO_AVALIAR_MIN) & (resultado['peso'] <= LIMITE_PESO_AVALIAR_MAX)
    cond_frete_amarelo = resultado['frete_fracionado'].notna() & (resultado['frete_fracionado'] >= LIMITE_FRETE_AVALIAR)
    cond_excecao_amarelo = resultado['excecao_operacional']

    resultado['regra_peso_1000'] = cond_peso_vermelho
    resultado['regra_frete_500'] = cond_frete_vermelho
    resultado['regra_peso_1000_valor_50000'] = cond_peso_valor_vermelho
    resultado['regra_frete_venda_8pct'] = cond_frete_venda_vermelho
    resultado['regra_peso_500_999'] = cond_peso_amarelo
    resultado['regra_frete_250'] = cond_frete_amarelo
    resultado['regra_excecao_operacional'] = cond_excecao_amarelo

    vermelho = cond_peso_vermelho | cond_frete_vermelho | cond_peso_valor_vermelho | cond_frete_venda_vermelho
    amarelo = (~vermelho) & (cond_peso_amarelo | cond_frete_amarelo | cond_excecao_amarelo)
    resultado['classificacao'] = np.select([vermelho, amarelo], ['Cotar dedicado obrigatório', 'Avaliar dedicado'], default='Manter fracionado')
    resultado['cor_matriz'] = np.select([vermelho, amarelo], ['Vermelho', 'Amarelo'], default='Verde')
    resultado['oportunidade_dedicado'] = resultado['classificacao'].isin(['Cotar dedicado obrigatório', 'Avaliar dedicado'])

    def regras_acionadas_linha(linha):
        regras = []
        if linha['regra_peso_1000']:
            regras.append('Peso >= 1000 kg')
        if linha['regra_frete_500']:
            regras.append('Frete estimado >= USD 500')
        if linha['regra_peso_1000_valor_50000']:
            regras.append('Peso >= 1000 kg e valor mercadoria >= USD 50000')
        if linha['regra_frete_venda_8pct']:
            regras.append('Frete/Venda >= 8%')
        if linha['regra_peso_500_999']:
            regras.append('Peso entre 500 e 999 kg')
        if linha['regra_frete_250']:
            regras.append('Frete estimado >= USD 250')
        if linha['regra_excecao_operacional']:
            regras.append('Cliente estratégico/OEM/carga crítica/entrega urgente')
        return '; '.join(regras) if regras else 'Sem gatilho operacional'

    resultado['regras_acionadas'] = resultado.apply(regras_acionadas_linha, axis=1)

    status_modal, explicacao_modal = identificar_modal_estimado(resultado)
    resultado['status_modal_estimado'] = status_modal
    resultado['explicacao_modal_estimado'] = explicacao_modal
    resultado['provavel_dedicado'] = resultado['status_modal_estimado'].eq('Provável dedicado')
    resultado['provavel_fracionado'] = resultado['status_modal_estimado'].eq('Provável fracionado')
    resultado['oportunidade_nao_realizada'] = resultado['oportunidade_dedicado'] & resultado['provavel_fracionado']
    resultado['oportunidade_ja_dedicado'] = resultado['oportunidade_dedicado'] & resultado['provavel_dedicado']

    resultado['economia_potencial'] = 0.0
    resultado['percentual_saving'] = np.nan
    resultado['origem_calculo_economia'] = 'Sem cálculo financeiro'
    tem_ambos = resultado['frete_fracionado'].notna() & resultado['frete_dedicado'].notna()
    economia_real = (resultado['frete_fracionado'] - resultado['frete_dedicado']).clip(lower=0)
    resultado.loc[tem_ambos, 'economia_potencial'] = economia_real.loc[tem_ambos]
    resultado.loc[tem_ambos & (resultado['frete_fracionado'] > 0), 'percentual_saving'] = resultado.loc[tem_ambos, 'economia_potencial'] / resultado.loc[tem_ambos, 'frete_fracionado']
    resultado.loc[tem_ambos, 'origem_calculo_economia'] = 'Frete fracionado e dedicado informados'

    if TAXA_ECONOMIA_SIMULADA_ATIVA:
        pode_simular = (~tem_ambos) & resultado['oportunidade_nao_realizada'] & resultado['frete_fracionado'].notna()
        resultado.loc[pode_simular, 'economia_potencial'] = resultado.loc[pode_simular, 'frete_fracionado'] * TAXA_ECONOMIA_SIMULADA
        resultado.loc[pode_simular, 'percentual_saving'] = TAXA_ECONOMIA_SIMULADA
        resultado.loc[pode_simular, 'origem_calculo_economia'] = f'Simulação fixa em código ({TAXA_ECONOMIA_SIMULADA:.0%})'

    resultado['data_convertida'] = pd.to_datetime(resultado['data'], errors='coerce', dayfirst=True)
    resultado['estado_final'] = resultado['estado'].where(resultado['estado'].str.strip() != '', resultado['destino'].apply(extrair_estado))
    resultado['rota'] = resultado['origem'].fillna('').astype(str).str.strip() + ' → ' + resultado['destino'].fillna('').astype(str).str.strip()
    resultado['rota'] = resultado['rota'].replace(' → ', 'Não informada')
    resultado['desconsiderado_regra_santos'] = identificar_rota_santos(resultado)
    return resultado


def preparar_tabela_exibicao(df):
    '''Cria cópia com colunas principais em ordem amigável para exibição/download.'''
    colunas_preferidas = [
        'documento', 'data', 'divisao', 'operacao_direcao', 'cliente', 'modalidade', 'status_modal_estimado',
        'transportadora', 'origem', 'destino', 'estado_final', 'valor_mercadoria', 'peso', 'frete_fracionado',
        'frete_dedicado', 'percentual_frete_sobre_venda', 'frete_fracionado_por_kg', 'classificacao',
        'cor_matriz', 'regras_acionadas', 'desconsiderado_regra_santos', 'oportunidade_nao_realizada',
        'oportunidade_ja_dedicado', 'economia_potencial', 'percentual_saving', 'origem_calculo_economia',
        'explicacao_modal_estimado', 'observacoes'
    ]
    colunas = [c for c in colunas_preferidas if c in df.columns]
    demais = [c for c in df.columns if c not in colunas]
    return df[colunas + demais].copy()


def criar_download_excel(df_resultado, df_oportunidades=None, resumo=None):
    '''Cria arquivo Excel em memória para download.'''
    output = io.BytesIO()
    df_oportunidades = pd.DataFrame() if df_oportunidades is None else df_oportunidades
    resumo = {} if resumo is None else resumo

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        pd.DataFrame([resumo]).to_excel(writer, index=False, sheet_name='Resumo')
        df_resultado.to_excel(writer, index=False, sheet_name='Analise_Completa')
        df_oportunidades.to_excel(writer, index=False, sheet_name='Oportunidades')
        workbook = writer.book
        header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': PARKER_RED, 'font_color': 'white', 'border': 1})
        money_format = workbook.add_format({'num_format': 'USD #,##0.00'})
        percent_format = workbook.add_format({'num_format': '0.00%'})

        for sheet_name in ['Resumo', 'Analise_Completa', 'Oportunidades']:
            worksheet = writer.sheets[sheet_name]
            df_sheet = pd.DataFrame([resumo]) if sheet_name == 'Resumo' else (df_resultado if sheet_name == 'Analise_Completa' else df_oportunidades)
            for col_num, col_name in enumerate(df_sheet.columns):
                worksheet.write(0, col_num, col_name, header_format)
                worksheet.set_column(col_num, col_num, min(max(len(str(col_name)) + 2, 14), 48))
            for idx, col_name in enumerate(df_sheet.columns):
                nome_col = str(col_name).lower()
                if any(palavra in nome_col for palavra in ['valor', 'frete', 'economia', 'mercadoria', 'gasto']):
                    worksheet.set_column(idx, idx, 16, money_format)
                if any(palavra in nome_col for palavra in ['percentual', 'saving', '%']):
                    worksheet.set_column(idx, idx, 14, percent_format)
    output.seek(0)
    return output.getvalue()


def kpi_card(label, value, subtitle='', accent=PARKER_RED):
    '''Renderiza card de KPI corporativo.'''
    st.markdown(
        f'''
        <div class='kpi-card' style='border-left: 5px solid {accent};'>
            <div class='kpi-label'>{label}</div>
            <div class='kpi-value'>{value}</div>
            <div class='kpi-subtitle'>{subtitle}</div>
        </div>
        ''',
        unsafe_allow_html=True,
    )


def badge_classificacao(classificacao):
    '''Retorna HTML para badge da recomendação operacional.'''
    if classificacao == 'Cotar dedicado obrigatório':
        classe = 'badge-red'
    elif classificacao == 'Avaliar dedicado':
        classe = 'badge-yellow'
    else:
        classe = 'badge-green'
    return f'<span class="{classe}">{classificacao}</span>'


def aplicar_filtro_multiselect(df, coluna, rotulo):
    '''Aplica filtro multiselect quando a coluna existir e tiver valores úteis.'''
    if coluna not in df.columns or df.empty:
        return df
    serie = df[coluna].fillna('').astype(str).str.strip()
    valores = sorted([v for v in serie.unique() if v != ''])
    if not valores:
        return df
    selecionados = st.multiselect(rotulo, valores, default=[])
    if selecionados:
        return df[serie.isin(selecionados)]
    return df


def safe_numeric(series):
    '''Converte série para numérico, removendo infinitos para uso seguro em gráficos.'''
    serie = pd.to_numeric(series, errors='coerce')
    return serie.replace([np.inf, -np.inf], np.nan)


def sanitizar_campos_grafico(df, numeric_cols=None, text_cols=None, required_numeric=None, positive_cols=None):
    '''Sanitiza campos usados antes de gerar gráficos Plotly.'''
    if df is None or df.empty:
        return pd.DataFrame()
    df_safe = df.copy()
    numeric_cols = numeric_cols or []
    text_cols = text_cols or []
    required_numeric = required_numeric or []
    positive_cols = positive_cols or []
    for col in numeric_cols:
        if col in df_safe.columns:
            df_safe[col] = safe_numeric(df_safe[col])
    for col in text_cols:
        if col in df_safe.columns:
            df_safe[col] = df_safe[col].fillna('').astype(str).replace('', 'Não informada')
    for col in required_numeric:
        if col in df_safe.columns:
            df_safe = df_safe[df_safe[col].notna()]
    for col in positive_cols:
        if col in df_safe.columns:
            df_safe = df_safe[df_safe[col].notna() & (df_safe[col] > 0)]
    return df_safe.copy()


def logo_html(caminho_logo):
    '''Retorna HTML do logo com tamanho limitado e responsivo.'''
    if caminho_logo and os.path.exists(caminho_logo):
        with open(caminho_logo, 'rb') as logo_file:
            encoded = base64.b64encode(logo_file.read()).decode('utf-8')
        return f"<div class='logo-panel'><img class='parker-logo-img' src='data:image/png;base64,{encoded}' alt='Parker Hannifin'></div>"
    return "<div class='logo-panel'><div class='logo-fallback'>Parker<br>Hannifin</div></div>"


# =============================================================================
# Cabeçalho corporativo e sidebar informativa
# =============================================================================

st.sidebar.title('Parker-Hannifin')
st.sidebar.caption('Branding e matriz fixa de decisão logística.')
st.sidebar.markdown('---')
st.sidebar.info('Os parâmetros da regra não são editáveis no aplicativo. A matriz está fixa no código conforme política informada.')

header_logo, header_text = st.columns([0.85, 5.15], vertical_alignment='center')
with header_logo:
    st.markdown(logo_html('assets/parker_logo.png'), unsafe_allow_html=True)
with header_text:
    st.markdown(
        '''
        <div class='parker-header'>
            <h1>Parker-Hannifin | Análise de Frete Dedicado</h1>
            <p>Matriz operacional fixa para decidir entre cotar dedicado, avaliar dedicado ou manter fracionado.</p>
        </div>
        ''',
        unsafe_allow_html=True,
    )

with st.expander('Matriz fixa de decisão operacional', expanded=True):
    r1, r2, r3 = st.columns(3)
    with r1:
        st.markdown(
            '''
            <div class='rule-card rule-red'>
                <h4>🔴 Cotar dedicado obrigatório</h4>
                <ul>
                    <li>Peso >= 1000 kg; ou</li>
                    <li>Frete fracionado estimado >= USD 500; ou</li>
                    <li>Peso >= 1000 kg e valor mercadoria >= USD 50000; ou</li>
                    <li>Frete/Venda >= 8%.</li>
                </ul>
            </div>
            ''',
            unsafe_allow_html=True,
        )
    with r2:
        st.markdown(
            '''
            <div class='rule-card rule-yellow'>
                <h4>🟡 Avaliar dedicado</h4>
                <ul>
                    <li>Peso entre 500 e 999 kg; ou</li>
                    <li>Frete fracionado estimado >= USD 250; ou</li>
                    <li>Cliente estratégico, OEM, carga crítica ou entrega urgente.</li>
                </ul>
            </div>
            ''',
            unsafe_allow_html=True,
        )
    with r3:
        st.markdown(
            '''
            <div class='rule-card rule-green'>
                <h4>🟢 Manter fracionado</h4>
                <ul>
                    <li>Baixo peso;</li>
                    <li>Baixo frete;</li>
                    <li>Frete/Venda < 8%;</li>
                    <li>Sem exceção operacional.</li>
                </ul>
            </div>
            ''',
            unsafe_allow_html=True,
        )

st.info(
    'O limite de **1000 kg** foi definido como gatilho operacional Parker para cotação dedicada. '
    'Esse parâmetro pode ser calibrado após validação de savings real e nível de serviço.'
)


# =============================================================================
# Abas principais
# =============================================================================

tab_manual, tab_lote = st.tabs(['1. Simulador Manual', '2. Análise em Lote'])


# =============================================================================
# Aba 1: Simulador Manual
# =============================================================================

with tab_manual:
    st.subheader('Simulador Manual de Decisão')
    st.write(
        'Informe os dados disponíveis do embarque. Frete fracionado e frete dedicado são opcionais; '
        'quando ambos forem preenchidos, o app calcula economia estimada e percentual de saving. '
        'Nesta simulação manual, os valores financeiros são tratados em moeda brasileira R$.'
    )

    with st.form('form_simulador_manual'):
        col1, col2, col3 = st.columns(3)
        with col1:
            valor_mercadoria = st.number_input('Valor da mercadoria (R$)', min_value=0.0, value=50000.0, step=1000.0, format='%.2f')
            peso = st.number_input('Peso bruto (kg)', min_value=0.0, value=1000.0, step=100.0, format='%.2f')
            frete_fracionado_txt = st.text_input('Frete fracionado informado (R$)', value='', placeholder='Ex.: 500,00')
        with col2:
            frete_dedicado_txt = st.text_input('Frete dedicado cotado (R$)', value='', placeholder='Ex.: 420,00')
            origem_uf = st.text_input('Origem - UF', value='SP', max_chars=2)
            origem_cidade = st.text_input('Origem - Cidade', value='São Paulo')
        with col3:
            destino_uf = st.text_input('Destino - UF', value='MG', max_chars=2)
            destino_cidade = st.text_input('Destino - Cidade', value='Belo Horizonte')
            transportadora = st.text_input('Transportadora (opcional)', value='')

        st.markdown('#### Exceções operacionais')
        e1, e2, e3, e4, e5 = st.columns(5)
        with e1:
            cliente_estrategico = st.checkbox('Cliente estratégico')
        with e2:
            cliente_oem = st.checkbox('OEM')
        with e3:
            carga_critica = st.checkbox('Carga crítica')
        with e4:
            entrega_urgente = st.checkbox('Entrega urgente')
        with e5:
            excecao_operacional = st.checkbox('Outra exceção')
        simular = st.form_submit_button('Calcular recomendação', type='primary')

    if simular:
        frete_fracionado = converter_numero(frete_fracionado_txt)
        frete_dedicado = converter_numero(frete_dedicado_txt)
        observacoes = []
        if cliente_estrategico:
            observacoes.append('cliente estratégico')
        if cliente_oem:
            observacoes.append('OEM')
        if carga_critica:
            observacoes.append('carga crítica')
        if entrega_urgente:
            observacoes.append('entrega urgente')
        if excecao_operacional:
            observacoes.append('exceção operacional')

        df_manual = pd.DataFrame([
            {
                'documento': 'SIMULAÇÃO MANUAL',
                'modalidade': 'Fracionado',
                'valor_mercadoria': valor_mercadoria,
                'peso': peso,
                'frete_fracionado': frete_fracionado,
                'frete_dedicado': frete_dedicado,
                'origem': f'{origem_cidade.strip()} / {origem_uf.strip().upper()}',
                'destino': f'{destino_cidade.strip()} / {destino_uf.strip().upper()}',
                'estado': destino_uf.strip().upper(),
                'data': datetime.today().strftime('%Y-%m-%d'),
                'transportadora': transportadora,
                'divisao': 'Simulação manual',
                'operacao_direcao': '',
                'cliente': 'OEM' if cliente_oem else '',
                'observacoes': '; '.join(observacoes),
            }
        ])

        resultado_manual = aplicar_regras(df_manual).iloc[0]
        tem_financeiro = pd.notna(resultado_manual['frete_fracionado']) and pd.notna(resultado_manual['frete_dedicado'])
        economia_manual = max(resultado_manual['frete_fracionado'] - resultado_manual['frete_dedicado'], 0) if tem_financeiro else np.nan
        saving_manual = economia_manual / resultado_manual['frete_fracionado'] if tem_financeiro and resultado_manual['frete_fracionado'] > 0 else np.nan
        regras_manual = str(resultado_manual['regras_acionadas']).replace('USD ', 'R$ ')

        st.markdown('### Resultado da Simulação')
        st.markdown(badge_classificacao(resultado_manual['classificacao']), unsafe_allow_html=True)
        st.write('')

        m1, m2, m3, m4 = st.columns(4)
        with m1:
            kpi_card('Frete/Venda', formatar_percentual(resultado_manual['percentual_frete_sobre_venda']), 'Calculado se frete e venda existirem')
        with m2:
            kpi_card('R$/kg', formatar_moeda(resultado_manual['frete_fracionado_por_kg'], 'R$'), 'Calculado se frete e peso existirem')
        with m3:
            kpi_card('Economia estimada (R$)', formatar_moeda(economia_manual, 'R$') if tem_financeiro else 'N/D', 'Exige frete fracionado e dedicado')
        with m4:
            kpi_card('Saving', formatar_percentual(saving_manual) if tem_financeiro else 'N/D', 'Exige frete fracionado e dedicado')

        if resultado_manual['classificacao'] == 'Cotar dedicado obrigatório':
            st.error('🔴 **Cotar dedicado obrigatório** — o embarque acionou pelo menos um gatilho vermelho da matriz Parker.')
        elif resultado_manual['classificacao'] == 'Avaliar dedicado':
            st.warning('🟡 **Avaliar dedicado** — existe gatilho amarelo ou exceção operacional que recomenda avaliação.')
        else:
            st.success('🟢 **Manter fracionado** — não foram identificados gatilhos para cotação dedicada, salvo decisão operacional.')
        if not tem_financeiro:
            st.info('Frete fracionado e/ou frete dedicado não preenchido. Resultado exibido apenas como recomendação operacional, sem cálculo financeiro.')

        detalhes = pd.DataFrame(
            {
                'Indicador': [
                    'Origem', 'Destino', 'Transportadora', 'Valor da mercadoria (R$)', 'Peso bruto',
                    'Frete fracionado informado (R$)', 'Frete dedicado cotado (R$)', 'Regras acionadas',
                    'Status estimado do modal', 'Explicação da heurística'
                ],
                'Valor': [
                    resultado_manual['origem'],
                    resultado_manual['destino'],
                    resultado_manual['transportadora'] or 'Não informado',
                    formatar_moeda(resultado_manual['valor_mercadoria'], 'R$'),
                    formatar_numero(resultado_manual['peso'], ' kg'),
                    'Não informado' if pd.isna(resultado_manual['frete_fracionado']) else formatar_moeda(resultado_manual['frete_fracionado'], 'R$'),
                    'Não informado' if pd.isna(resultado_manual['frete_dedicado']) else formatar_moeda(resultado_manual['frete_dedicado'], 'R$'),
                    regras_manual,
                    resultado_manual['status_modal_estimado'],
                    resultado_manual['explicacao_modal_estimado'],
                ],
            }
        )
        st.dataframe(detalhes, use_container_width=True, hide_index=True)
    else:
        st.info('Preencha os campos disponíveis e clique em **Calcular recomendação** para visualizar a análise.')


# =============================================================================
# Aba 2: Análise em Lote
# =============================================================================

with tab_lote:
    st.subheader('Análise em Lote de Histórico de Embarques')
    st.write(
        'Faça upload de um arquivo Excel ou CSV. O aplicativo detecta automaticamente as colunas prováveis '
        'em português ou inglês. Na análise em massa, os valores monetários da base são mantidos em dólar/US$, '
        'sem conversão. Rotas com origem ou destino contendo **Santos/SP** ou **Santos** são desconsideradas automaticamente.'
    )

    arquivo = st.file_uploader('Upload de histórico de embarques (.xlsx, .xls ou .csv)', type=['xlsx', 'xls', 'csv'])
    if arquivo is not None:
        file_id = f'{arquivo.name}_{arquivo.size}'
        if st.session_state.get('arquivo_lote_id') != file_id:
            st.session_state['arquivo_lote_id'] = file_id
            st.session_state.pop('df_resultado_lote', None)
            st.session_state.pop('df_exibicao_lote', None)

    df_original = carregar_arquivo(arquivo) if arquivo is not None else None

    if df_original is not None and not df_original.empty:
        st.success(f'Arquivo carregado com sucesso: {len(df_original):,} linhas e {len(df_original.columns):,} colunas.'.replace(',', '.'))
        with st.expander('Pré-visualização do arquivo carregado', expanded=False):
            st.dataframe(df_original.head(50), use_container_width=True)

        mapa_colunas = detectar_colunas_automaticamente(df_original)
        campos_criticos = {
            'peso': 'Peso bruto',
            'frete_fracionado': 'Frete fracionado / frete estimado',
            'valor_mercadoria': 'Valor da mercadoria / venda',
        }
        criticos_nao_detectados = [rotulo for campo, rotulo in campos_criticos.items() if mapa_colunas.get(campo) == 'Não disponível']
        criticos_detectados = [campo for campo in campos_criticos if mapa_colunas.get(campo) != 'Não disponível']

        st.markdown('### Detecção Automática de Colunas')
        if criticos_nao_detectados:
            st.warning(
                'A detecção automática não encontrou todos os campos críticos: '
                f'{", ".join(criticos_nao_detectados)}. '
                'Ajuste os nomes das colunas na planilha ou revise o arquivo para melhorar a análise. '
                'Não há mapeamento manual nesta versão.'
            )
        else:
            st.info('Colunas críticas detectadas automaticamente. Revise a pré-visualização acima caso deseje validar a estrutura do arquivo.')

        with st.expander('Ver colunas detectadas automaticamente', expanded=False):
            df_detectadas = pd.DataFrame(
                [{'Campo esperado': campo, 'Coluna detectada': coluna} for campo, coluna in mapa_colunas.items()]
            )
            st.dataframe(df_detectadas, use_container_width=True, hide_index=True)

        if not criticos_detectados:
            st.error(
                'Não foi possível iniciar a análise porque nenhum campo crítico foi detectado automaticamente. '
                'Revise o arquivo e use nomes de colunas como Peso, Valor Frete e Valor Mercadoria.'
            )
            analisar = False
        else:
            analisar = st.button('Aplicar matriz no banco', type='primary')

        if analisar:
            df_padrao = construir_df_padronizado(df_original, mapa_colunas)
            df_resultado = aplicar_regras(df_padrao)
            df_exibicao = preparar_tabela_exibicao(df_resultado)
            st.session_state['df_resultado_lote'] = df_resultado
            st.session_state['df_exibicao_lote'] = df_exibicao

        if 'df_resultado_lote' in st.session_state:
            df_resultado = st.session_state['df_resultado_lote'].copy()
            regra_santos = df_resultado['desconsiderado_regra_santos'] if 'desconsiderado_regra_santos' in df_resultado.columns else pd.Series(False, index=df_resultado.index)
            qtd_santos_desconsiderados = int(regra_santos.sum())
            df_base_analise = df_resultado[~regra_santos].copy()

            st.markdown('### Regra de Exclusão Automática')
            s1, s2 = st.columns([1, 3])
            with s1:
                kpi_card('Registros Santos desconsiderados', f'{qtd_santos_desconsiderados:,}'.replace(',', '.'), 'Origem ou destino Santos/SP ou Santos', PARKER_MID_GRAY)
            with s2:
                if qtd_santos_desconsiderados > 0:
                    st.warning(f'{qtd_santos_desconsiderados:,} registro(s) foram removidos automaticamente da análise em massa, KPIs e gráficos de oportunidade por regra Santos.'.replace(',', '.'))
                else:
                    st.info('Nenhum registro foi desconsiderado pela regra Santos.')

            st.markdown('### Filtros Operacionais')
            with st.expander('Filtrar análise', expanded=True):
                f1, f2, f3, f4 = st.columns(4)
                df_filtrado = df_base_analise.copy()
                with f1:
                    df_filtrado = aplicar_filtro_multiselect(df_filtrado, 'divisao', 'Divisão')
                    df_filtrado = aplicar_filtro_multiselect(df_filtrado, 'operacao_direcao', 'Operação / Direção')
                with f2:
                    df_filtrado = aplicar_filtro_multiselect(df_filtrado, 'transportadora', 'Transportadora')
                    df_filtrado = aplicar_filtro_multiselect(df_filtrado, 'estado_final', 'Estado')
                with f3:
                    df_filtrado = aplicar_filtro_multiselect(df_filtrado, 'origem', 'Origem')
                    df_filtrado = aplicar_filtro_multiselect(df_filtrado, 'destino', 'Destino')
                with f4:
                    df_filtrado = aplicar_filtro_multiselect(df_filtrado, 'classificacao', 'Classificação')
                    df_filtrado = aplicar_filtro_multiselect(df_filtrado, 'status_modal_estimado', 'Status estimado do modal usado')

                if 'data_convertida' in df_filtrado.columns and df_filtrado['data_convertida'].notna().any():
                    data_min = df_filtrado['data_convertida'].min().date()
                    data_max = df_filtrado['data_convertida'].max().date()
                    periodo = st.date_input('Período', value=(data_min, data_max), min_value=data_min, max_value=data_max)
                    if isinstance(periodo, tuple) and len(periodo) == 2:
                        inicio, fim = periodo
                        df_filtrado = df_filtrado[(df_filtrado['data_convertida'].dt.date >= inicio) & (df_filtrado['data_convertida'].dt.date <= fim)]

            df_exibicao_filtrado = preparar_tabela_exibicao(df_filtrado)
            oportunidades = df_filtrado[df_filtrado['oportunidade_dedicado']].copy() if 'oportunidade_dedicado' in df_filtrado.columns else pd.DataFrame()
            oportunidades_nao_realizadas = df_filtrado[df_filtrado['oportunidade_nao_realizada']].copy() if 'oportunidade_nao_realizada' in df_filtrado.columns else pd.DataFrame()
            oportunidades_ja_dedicado = df_filtrado[df_filtrado['oportunidade_ja_dedicado']].copy() if 'oportunidade_ja_dedicado' in df_filtrado.columns else pd.DataFrame()
            df_oportunidades_exibicao = preparar_tabela_exibicao(oportunidades).sort_values(
                by=[c for c in ['cor_matriz', 'economia_potencial', 'frete_fracionado', 'valor_mercadoria'] if c in oportunidades.columns],
                ascending=False,
            ) if not oportunidades.empty else pd.DataFrame()

            total_analisado = len(df_filtrado)
            provavel_fracionado = int(df_filtrado['provavel_fracionado'].sum()) if 'provavel_fracionado' in df_filtrado.columns else 0
            provavel_dedicado = int(df_filtrado['provavel_dedicado'].sum()) if 'provavel_dedicado' in df_filtrado.columns else 0
            qtd_oportunidades_nao_realizadas = int(df_filtrado['oportunidade_nao_realizada'].sum()) if 'oportunidade_nao_realizada' in df_filtrado.columns else 0
            qtd_oportunidades_ja_dedicado = int(df_filtrado['oportunidade_ja_dedicado'].sum()) if 'oportunidade_ja_dedicado' in df_filtrado.columns else 0
            gasto_fracionado_oportunidade = float(safe_numeric(oportunidades_nao_realizadas['frete_fracionado']).sum(skipna=True)) if not oportunidades_nao_realizadas.empty and 'frete_fracionado' in oportunidades_nao_realizadas.columns else 0.0
            economia_potencial_total = float(safe_numeric(oportunidades_nao_realizadas['economia_potencial']).sum(skipna=True)) if not oportunidades_nao_realizadas.empty and 'economia_potencial' in oportunidades_nao_realizadas.columns else 0.0

            resumo = {
                'total_analisado': total_analisado,
                'registros_desconsiderados_regra_santos': qtd_santos_desconsiderados,
                'provavel_fracionado': provavel_fracionado,
                'provavel_dedicado': provavel_dedicado,
                'oportunidades_em_provavel_fracionado': qtd_oportunidades_nao_realizadas,
                'oportunidades_ja_feitas_como_dedicado': qtd_oportunidades_ja_dedicado,
                'gasto_em_frete_fracionado_com_oportunidade': gasto_fracionado_oportunidade,
                'economia_estimada': economia_potencial_total,
                'taxa_simulada_ativa': TAXA_ECONOMIA_SIMULADA_ATIVA,
                'taxa_simulada_codigo': TAXA_ECONOMIA_SIMULADA,
                'data_processamento': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            }

            st.markdown('### KPIs da Análise')
            k1, k2, k3, k4 = st.columns(4)
            with k1:
                kpi_card('Total analisado', f'{total_analisado:,}'.replace(',', '.'), 'Após filtros e exclusão Santos')
            with k2:
                kpi_card('Provável fracionado', f'{provavel_fracionado:,}'.replace(',', '.'), 'Base elegível para oportunidade não realizada', PARKER_MID_GRAY)
            with k3:
                kpi_card('Provável dedicado', f'{provavel_dedicado:,}'.replace(',', '.'), 'Desconsiderado do KPI de não realizado', PARKER_DARK)
            with k4:
                kpi_card('Oportunidades em provável fracionado', f'{qtd_oportunidades_nao_realizadas:,}'.replace(',', '.'), 'Cotar ou avaliar dedicado', PARKER_RED)

            k5, k6, k7, k8 = st.columns(4)
            with k5:
                kpi_card('Oportunidades já como dedicado', f'{qtd_oportunidades_ja_dedicado:,}'.replace(',', '.'), 'Identificadas pela heurística', PARKER_GREEN)
            with k6:
                kpi_card('Gasto fracionado com oportunidade', formatar_moeda(gasto_fracionado_oportunidade), 'Somente provável fracionado')
            with k7:
                kpi_card('Economia estimada', formatar_moeda(economia_potencial_total), 'Real ou simulada conforme código')
            with k8:
                pct = economia_potencial_total / gasto_fracionado_oportunidade if gasto_fracionado_oportunidade > 0 else np.nan
                kpi_card('Saving sobre gasto elegível', formatar_percentual(pct), 'Base: oportunidades não realizadas')

            st.info(
                'A identificação de **Provável dedicado** é uma heurística. Ela usa palavras-chave em transportadora, observações ou modal; '
                'peso alto com padrão de carga fechada; e R$/kg muito baixo para cargas pesadas. Valide operacionalmente antes de concluir.'
            )

            st.markdown('### Visualizações')
            graf_tab1, graf_tab2, graf_tab3, graf_tab4 = st.tabs(['Resumo e oportunidades', 'Distribuições', 'Matriz peso x valor', 'Rotas e savings'])

            with graf_tab1:
                g1, g2 = st.columns(2)
                with g1:
                    if not df_filtrado.empty and 'classificacao' in df_filtrado.columns:
                        df_classificacao = df_filtrado.groupby('classificacao', dropna=False).size().reset_index(name='quantidade').sort_values('quantidade', ascending=False)
                        fig = px.pie(
                            df_classificacao,
                            names='classificacao',
                            values='quantidade',
                            title='Distribuição por recomendação operacional',
                            color='classificacao',
                            color_discrete_map={'Cotar dedicado obrigatório': PARKER_RED, 'Avaliar dedicado': PARKER_YELLOW, 'Manter fracionado': PARKER_GREEN},
                            hole=0.35,
                        )
                        fig.update_traces(textposition='inside', textinfo='percent+label')
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info('Sem dados de classificação para exibir.')
                with g2:
                    if not oportunidades_nao_realizadas.empty and 'divisao' in oportunidades_nao_realizadas.columns:
                        df_div = oportunidades_nao_realizadas.copy()
                        df_div['divisao'] = df_div['divisao'].fillna('').replace('', 'Não informada')
                        df_div = df_div.groupby('divisao').size().reset_index(name='oportunidades').sort_values('oportunidades', ascending=False).head(15)
                        fig = px.bar(df_div, x='divisao', y='oportunidades', title='Oportunidades por divisão', color_discrete_sequence=[PARKER_RED])
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info('Sem oportunidades por divisão para exibir.')

                g3, g4 = st.columns(2)
                with g3:
                    if not oportunidades_nao_realizadas.empty and 'transportadora' in oportunidades_nao_realizadas.columns:
                        df_transp = oportunidades_nao_realizadas.copy()
                        df_transp['transportadora'] = df_transp['transportadora'].fillna('').replace('', 'Não informada')
                        df_transp = df_transp.groupby('transportadora').size().reset_index(name='oportunidades').sort_values('oportunidades', ascending=False).head(15)
                        fig = px.bar(df_transp, x='oportunidades', y='transportadora', orientation='h', title='Oportunidades por transportadora', color_discrete_sequence=[PARKER_RED])
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info('Sem oportunidades por transportadora para exibir.')
                with g4:
                    if not oportunidades_nao_realizadas.empty and {'origem', 'destino'}.issubset(oportunidades_nao_realizadas.columns):
                        df_od = oportunidades_nao_realizadas.copy()
                        df_od['origem'] = df_od['origem'].fillna('').replace('', 'Não informada')
                        df_od['destino'] = df_od['destino'].fillna('').replace('', 'Não informada')
                        df_od = df_od.groupby(['origem', 'destino'], dropna=False).size().reset_index(name='oportunidades').sort_values('oportunidades', ascending=False).head(15)
                        df_od['origem_destino'] = df_od['origem'] + ' → ' + df_od['destino']
                        fig = px.bar(df_od, x='oportunidades', y='origem_destino', orientation='h', title='Oportunidades por origem/destino', color_discrete_sequence=[PARKER_RED])
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info('Sem oportunidades por origem/destino para exibir.')

            with graf_tab2:
                d1, d2, d3 = st.columns(3)
                with d1:
                    df_hist = sanitizar_campos_grafico(df_filtrado, numeric_cols=['peso'], text_cols=['classificacao'], required_numeric=['peso'])
                    if not df_hist.empty:
                        fig = px.histogram(df_hist, x='peso', nbins=30, title='Distribuição por peso', color='classificacao', color_discrete_map={'Cotar dedicado obrigatório': PARKER_RED, 'Avaliar dedicado': PARKER_YELLOW, 'Manter fracionado': PARKER_GREEN})
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info('Sem dados válidos de peso para exibir.')
                with d2:
                    df_hist = sanitizar_campos_grafico(df_filtrado, numeric_cols=['valor_mercadoria'], text_cols=['classificacao'], required_numeric=['valor_mercadoria'])
                    if not df_hist.empty:
                        fig = px.histogram(df_hist, x='valor_mercadoria', nbins=30, title='Distribuição por valor de mercadoria', color='classificacao', color_discrete_map={'Cotar dedicado obrigatório': PARKER_RED, 'Avaliar dedicado': PARKER_YELLOW, 'Manter fracionado': PARKER_GREEN})
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info('Sem dados válidos de valor de mercadoria para exibir.')
                with d3:
                    df_hist = sanitizar_campos_grafico(df_filtrado, numeric_cols=['frete_fracionado'], text_cols=['classificacao'], required_numeric=['frete_fracionado'])
                    if not df_hist.empty:
                        fig = px.histogram(df_hist, x='frete_fracionado', nbins=30, title='Distribuição por frete estimado', color='classificacao', color_discrete_map={'Cotar dedicado obrigatório': PARKER_RED, 'Avaliar dedicado': PARKER_YELLOW, 'Manter fracionado': PARKER_GREEN})
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info('Sem dados válidos de frete estimado para exibir.')

            with graf_tab3:
                df_scatter = sanitizar_campos_grafico(
                    df_filtrado,
                    numeric_cols=['peso', 'valor_mercadoria', 'frete_fracionado'],
                    text_cols=['classificacao', 'documento', 'transportadora', 'origem', 'destino', 'regras_acionadas', 'status_modal_estimado'],
                    required_numeric=['peso', 'valor_mercadoria'],
                    positive_cols=['peso', 'valor_mercadoria'],
                )
                if not df_scatter.empty:
                    df_scatter['_plot_size'] = safe_numeric(df_scatter.get('frete_fracionado', pd.Series(index=df_scatter.index))).fillna(12).clip(lower=6, upper=55)
                    hover_cols = [c for c in ['documento', 'transportadora', 'origem', 'destino', 'regras_acionadas', 'status_modal_estimado'] if c in df_scatter.columns]
                    fig = px.scatter(
                        df_scatter,
                        x='peso',
                        y='valor_mercadoria',
                        size='_plot_size',
                        color='classificacao' if 'classificacao' in df_scatter.columns else None,
                        hover_data=hover_cols,
                        title='Matriz peso versus valor da mercadoria',
                        color_discrete_map={'Cotar dedicado obrigatório': PARKER_RED, 'Avaliar dedicado': PARKER_YELLOW, 'Manter fracionado': PARKER_GREEN},
                    )
                    fig.add_vline(x=LIMITE_PESO_COTAR_OBRIGATORIO, line_dash='dash', line_color=PARKER_RED)
                    fig.add_hline(y=LIMITE_VALOR_MERCADORIA_COTAR, line_dash='dot', line_color=PARKER_YELLOW)
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info('Não há peso e valor de mercadoria suficientes para montar a matriz.')

            with graf_tab4:
                r1, r2 = st.columns(2)
                with r1:
                    if not oportunidades_nao_realizadas.empty and 'rota' in oportunidades_nao_realizadas.columns:
                        df_rotas = oportunidades_nao_realizadas.copy()
                        df_rotas['economia_potencial'] = safe_numeric(df_rotas['economia_potencial']).fillna(0)
                        df_rotas['frete_fracionado'] = safe_numeric(df_rotas['frete_fracionado']).fillna(0)
                        df_rotas['rota'] = df_rotas['rota'].fillna('').replace('', 'Não informada')
                        df_rotas = df_rotas.groupby('rota', dropna=False).agg(oportunidades=('rota', 'size'), economia=('economia_potencial', 'sum'), gasto=('frete_fracionado', 'sum')).reset_index().sort_values(['oportunidades', 'economia'], ascending=False).head(15)
                        fig = px.bar(df_rotas, x='oportunidades', y='rota', orientation='h', title='Top rotas com oportunidade', color='economia', color_continuous_scale=['#FDECEF', PARKER_RED])
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info('Sem rotas com oportunidade para exibir.')
                with r2:
                    df_savings = sanitizar_campos_grafico(
                        oportunidades_nao_realizadas,
                        numeric_cols=['economia_potencial'],
                        text_cols=['documento', 'origem_calculo_economia'],
                        required_numeric=['economia_potencial'],
                        positive_cols=['economia_potencial'],
                    )
                    if not df_savings.empty:
                        df_savings['documento_plot'] = df_savings['documento'].replace('', np.nan).fillna('Sem documento') if 'documento' in df_savings.columns else 'Sem documento'
                        df_savings = df_savings.sort_values('economia_potencial', ascending=False).head(15)
                        fig = px.bar(
                            df_savings.sort_values('economia_potencial'),
                            x='economia_potencial',
                            y='documento_plot',
                            orientation='h',
                            title='Ranking de savings',
                            color='origem_calculo_economia' if 'origem_calculo_economia' in df_savings.columns else None,
                            color_discrete_sequence=[PARKER_RED, PARKER_YELLOW, PARKER_GREEN],
                        )
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info('Ranking de savings disponível somente quando houver economia positiva.')

            st.markdown('### Tabela de Oportunidades Priorizadas')
            if df_oportunidades_exibicao.empty:
                st.info('Nenhuma oportunidade de frete dedicado foi identificada com a matriz fixa nos filtros atuais.')
            else:
                st.dataframe(
                    df_oportunidades_exibicao,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        'valor_mercadoria': st.column_config.NumberColumn('Valor mercadoria', format='USD %.2f'),
                        'frete_fracionado': st.column_config.NumberColumn('Frete estimado/fracionado', format='USD %.2f'),
                        'frete_dedicado': st.column_config.NumberColumn('Frete dedicado', format='USD %.2f'),
                        'economia_potencial': st.column_config.NumberColumn('Economia potencial', format='USD %.2f'),
                        'percentual_frete_sobre_venda': st.column_config.NumberColumn('Frete/Venda', format='%.2f'),
                        'percentual_saving': st.column_config.NumberColumn('Saving', format='%.2f'),
                        'frete_fracionado_por_kg': st.column_config.NumberColumn('USD/kg', format='USD %.2f'),
                    },
                )

            st.markdown('### Download dos Resultados')
            excel_bytes = criar_download_excel(df_exibicao_filtrado, df_oportunidades_exibicao, resumo)
            csv_oportunidades = df_oportunidades_exibicao.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')
            csv_completo = df_exibicao_filtrado.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')

            d1, d2, d3 = st.columns(3)
            with d1:
                st.download_button('Baixar Excel completo', data=excel_bytes, file_name='parker_analise_frete_dedicado.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            with d2:
                st.download_button('Baixar CSV de oportunidades', data=csv_oportunidades, file_name='parker_oportunidades_frete_dedicado.csv', mime='text/csv')
            with d3:
                st.download_button('Baixar CSV completo', data=csv_completo, file_name='parker_analise_completa_frete.csv', mime='text/csv')

            with st.expander('Ver análise completa filtrada', expanded=False):
                st.dataframe(df_exibicao_filtrado, use_container_width=True, hide_index=True)
            if qtd_santos_desconsiderados > 0:
                with st.expander('Ver registros desconsiderados pela regra Santos', expanded=False):
                    st.dataframe(preparar_tabela_exibicao(df_resultado[regra_santos].copy()), use_container_width=True, hide_index=True)

    elif arquivo is not None:
        st.warning('O arquivo foi carregado, mas não possui dados para análise.')
    else:
        st.info('Envie um arquivo para iniciar a análise em lote.')


st.markdown('---')
st.caption(
    'Parker-Hannifin | Ferramenta interna de apoio à decisão logística. '
    'A matriz de decisão está fixa no código e não substitui validação operacional, contratação formal ou análise de nível de serviço.'
)
