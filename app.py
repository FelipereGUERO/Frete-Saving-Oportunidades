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
# =============================================================================

import io
import os
import re
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

# Regras fixas conforme matriz operacional do usuário.
LIMITE_PESO_COTAR_OBRIGATORIO = 3000.0
LIMITE_FRETE_COTAR_OBRIGATORIO = 500.0
LIMITE_PESO_VALOR_COTAR = 1000.0
LIMITE_VALOR_MERCADORIA_COTAR = 50000.0
LIMITE_FRETE_VENDA_COTAR = 0.08
LIMITE_PESO_AVALIAR_MIN = 500.0
LIMITE_PESO_AVALIAR_MAX = 999.0
LIMITE_FRETE_AVALIAR = 250.0

# Taxa fixa em código para simulação gerencial em lote quando não houver valor dedicado.
# Altere no código somente se a política logística permitir simulação estimada.
TAXA_ECONOMIA_SIMULADA_ATIVA = True
TAXA_ECONOMIA_SIMULADA = 0.15

# Heurística fixa para identificar fretes provavelmente dedicados.
LIMITE_FRETE_KG_BAIXO_CARGA_PESADA = 0.35
LIMITE_FRETE_KG_MUITO_BAIXO = 0.20

st.markdown(
    f'''
    <style>
        .main {{
            background-color: {PARKER_WHITE};
        }}
        .block-container {{
            padding-top: 1.2rem;
            padding-bottom: 2rem;
        }}
        .parker-header {{
            background: linear-gradient(90deg, {PARKER_RED} 0%, #8A0018 100%);
            padding: 1.25rem 1.5rem;
            border-radius: 14px;
            color: white;
            margin-bottom: 1.2rem;
            box-shadow: 0 6px 18px rgba(17,24,39,0.16);
        }}
        .parker-header h1 {{
            margin: 0;
            font-size: 2.0rem;
            font-weight: 800;
            letter-spacing: -0.02em;
        }}
        .parker-header p {{
            margin: 0.35rem 0 0 0;
            font-size: 1.0rem;
            opacity: 0.95;
        }}
        .rule-card {{
            border-radius: 12px;
            padding: 1rem;
            border: 1px solid #E5E7EB;
            background-color: white;
            min-height: 170px;
            box-shadow: 0 2px 8px rgba(17,24,39,0.06);
        }}
        .rule-red {{ border-top: 5px solid {PARKER_RED}; }}
        .rule-yellow {{ border-top: 5px solid {PARKER_YELLOW}; }}
        .rule-green {{ border-top: 5px solid {PARKER_GREEN}; }}
        .kpi-card {{
            background: white;
            border: 1px solid #E5E7EB;
            border-radius: 12px;
            padding: 1rem;
            box-shadow: 0 2px 8px rgba(17,24,39,0.06);
        }}
        .kpi-label {{
            color: {PARKER_MID_GRAY};
            font-size: 0.82rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.04em;
        }}
        .kpi-value {{
            color: {PARKER_DARK};
            font-size: 1.45rem;
            font-weight: 850;
            margin-top: 0.25rem;
        }}
        .kpi-subtitle {{
            color: {PARKER_MID_GRAY};
            font-size: 0.82rem;
            margin-top: 0.25rem;
        }}
        .badge-red {{
            display: inline-block;
            background: {PARKER_LIGHT_RED};
            color: {PARKER_RED};
            border: 1px solid #F8C7D0;
            padding: 0.35rem 0.6rem;
            border-radius: 999px;
            font-weight: 800;
        }}
        .badge-yellow {{
            display: inline-block;
            background: #FFF7ED;
            color: #B45309;
            border: 1px solid #FED7AA;
            padding: 0.35rem 0.6rem;
            border-radius: 999px;
            font-weight: 800;
        }}
        .badge-green {{
            display: inline-block;
            background: #ECFDF5;
            color: #047857;
            border: 1px solid #A7F3D0;
            padding: 0.35rem 0.6rem;
            border-radius: 999px;
            font-weight: 800;
        }}
        div[data-testid='stMetricValue'] {{
            color: {PARKER_DARK};
            font-weight: 800;
        }}
        section[data-testid='stSidebar'] h1,
        section[data-testid='stSidebar'] h2,
        section[data-testid='stSidebar'] h3 {{
            color: {PARKER_RED};
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
    '''
    Converte números em formatos brasileiros e internacionais para float.

    Exemplos aceitos:
    - '1.234,56' -> 1234.56
    - '1234,56'  -> 1234.56
    - '1,234.56' -> 1234.56
    - 'USD 5,000.00' -> 5000.0
    - pandas Series -> Series numérica
    '''
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
        else:
            if texto.count('.') > 1:
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
                    buffer = io.BytesIO(conteudo)
                    return pd.read_csv(buffer, dtype=str, sep=None, engine='python', encoding=encoding)
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


def sugerir_coluna(colunas, termos):
    '''Sugere coluna a partir de uma lista de termos prováveis.'''
    colunas_lower = {str(c).lower(): c for c in colunas}
    for termo in termos:
        termo_lower = termo.lower()
        for coluna_lower, coluna_original in colunas_lower.items():
            if termo_lower in coluna_lower:
                return coluna_original
    return 'Não disponível'


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


def texto_contem(texto, palavras):
    '''Verifica se texto contém qualquer palavra-chave.'''
    if pd.isna(texto):
        return False
    texto_norm = str(texto).lower()
    return any(palavra in texto_norm for palavra in palavras)


def construir_df_padronizado(df_original, mapa_colunas):
    '''Constrói DataFrame com nomes padronizados a partir do mapeamento do usuário.'''
    campos_texto = [
        'documento', 'modalidade', 'origem', 'destino', 'estado', 'data', 'transportadora',
        'divisao', 'operacao_direcao', 'cliente', 'observacoes'
    ]
    campos_numericos = ['valor_mercadoria', 'peso', 'frete_fracionado', 'frete_dedicado']

    df_padrao = pd.DataFrame(index=df_original.index)

    for destino, origem in mapa_colunas.items():
        if origem and origem != 'Não disponível' and origem in df_original.columns:
            if destino in campos_texto:
                df_padrao[destino] = df_original[origem].fillna('').astype(str)
            else:
                df_padrao[destino] = df_original[origem]
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
    '''
    Aplica a matriz fixa Parker-Hannifin para cotação dedicada.

    Vermelho / Cotar dedicado obrigatório:
    - Peso >= 3000 kg; ou
    - Frete fracionado estimado >= USD 500; ou
    - Peso >= 1000 kg e valor da mercadoria >= USD 50000; ou
    - Frete/Venda >= 0,08.

    Amarelo / Avaliar dedicado:
    - Peso entre 500 e 999 kg e frete >= USD 250; ou
    - Cliente estratégico/OEM/carga crítica/entrega urgente/exceção operacional.

    Verde / Manter fracionado:
    - Sem gatilhos vermelho/amarelo, especialmente quando peso < 500 kg,
      frete < USD 250 e frete/venda < 8%, salvo exceção operacional.
    '''
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

    resultado['valor_mercadoria'] = resultado['valor_mercadoria'].where(resultado['valor_mercadoria'] > 0, np.nan)
    resultado['peso'] = resultado['peso'].where(resultado['peso'] > 0, np.nan)
    resultado['frete_fracionado'] = resultado['frete_fracionado'].where(resultado['frete_fracionado'] > 0, np.nan)
    resultado['frete_dedicado'] = resultado['frete_dedicado'].where(resultado['frete_dedicado'] > 0, np.nan)

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
    cond_frete_venda_vermelho = (
        resultado['percentual_frete_sobre_venda'].notna() &
        (resultado['percentual_frete_sobre_venda'] >= LIMITE_FRETE_VENDA_COTAR)
    )

    cond_peso_frete_amarelo = (
        resultado['peso'].notna() & resultado['frete_fracionado'].notna() &
        (resultado['peso'] >= LIMITE_PESO_AVALIAR_MIN) &
        (resultado['peso'] <= LIMITE_PESO_AVALIAR_MAX) &
        (resultado['frete_fracionado'] >= LIMITE_FRETE_AVALIAR)
    )
    cond_excecao_amarelo = resultado['excecao_operacional']

    resultado['regra_peso_3000'] = cond_peso_vermelho
    resultado['regra_frete_500'] = cond_frete_vermelho
    resultado['regra_peso_1000_valor_50000'] = cond_peso_valor_vermelho
    resultado['regra_frete_venda_8pct'] = cond_frete_venda_vermelho
    resultado['regra_peso_500_999_frete_250'] = cond_peso_frete_amarelo
    resultado['regra_excecao_operacional'] = cond_excecao_amarelo

    vermelho = cond_peso_vermelho | cond_frete_vermelho | cond_peso_valor_vermelho | cond_frete_venda_vermelho
    amarelo = (~vermelho) & (cond_peso_frete_amarelo | cond_excecao_amarelo)

    resultado['classificacao'] = np.select(
        [vermelho, amarelo],
        ['Cotar dedicado obrigatório', 'Avaliar dedicado'],
        default='Manter fracionado',
    )
    resultado['cor_matriz'] = np.select([vermelho, amarelo], ['Vermelho', 'Amarelo'], default='Verde')
    resultado['oportunidade_dedicado'] = resultado['classificacao'].isin(['Cotar dedicado obrigatório', 'Avaliar dedicado'])

    def regras_acionadas_linha(linha):
        regras = []
        if linha['regra_peso_3000']:
            regras.append('Peso >= 3000 kg')
        if linha['regra_frete_500']:
            regras.append('Frete estimado >= USD 500')
        if linha['regra_peso_1000_valor_50000']:
            regras.append('Peso >= 1000 kg e valor mercadoria >= USD 50000')
        if linha['regra_frete_venda_8pct']:
            regras.append('Frete/Venda >= 8%')
        if linha['regra_peso_500_999_frete_250']:
            regras.append('Peso 500-999 kg e frete >= USD 250')
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
    resultado.loc[tem_ambos & (resultado['frete_fracionado'] > 0), 'percentual_saving'] = (
        resultado.loc[tem_ambos, 'economia_potencial'] / resultado.loc[tem_ambos, 'frete_fracionado']
    )
    resultado.loc[tem_ambos, 'origem_calculo_economia'] = 'Frete fracionado e dedicado informados'

    if TAXA_ECONOMIA_SIMULADA_ATIVA:
        pode_simular = (
            (~tem_ambos) & resultado['oportunidade_nao_realizada'] & resultado['frete_fracionado'].notna()
        )
        resultado.loc[pode_simular, 'economia_potencial'] = (
            resultado.loc[pode_simular, 'frete_fracionado'] * TAXA_ECONOMIA_SIMULADA
        )
        resultado.loc[pode_simular, 'percentual_saving'] = TAXA_ECONOMIA_SIMULADA
        resultado.loc[pode_simular, 'origem_calculo_economia'] = f'Simulação fixa em código ({TAXA_ECONOMIA_SIMULADA:.0%})'

    resultado['data_convertida'] = pd.to_datetime(resultado['data'], errors='coerce', dayfirst=True)
    resultado['estado_final'] = resultado['estado'].where(resultado['estado'].str.strip() != '', resultado['destino'].apply(extrair_estado))
    resultado['rota'] = resultado['origem'].fillna('').astype(str).str.strip() + ' → ' + resultado['destino'].fillna('').astype(str).str.strip()
    resultado['rota'] = resultado['rota'].replace(' → ', 'Não informada')

    return resultado


def preparar_tabela_exibicao(df):
    '''Cria cópia com colunas principais em ordem amigável para exibição/download.'''
    colunas_preferidas = [
        'documento', 'data', 'divisao', 'operacao_direcao', 'cliente', 'modalidade', 'status_modal_estimado',
        'transportadora', 'origem', 'destino', 'estado_final', 'valor_mercadoria', 'peso', 'frete_fracionado',
        'frete_dedicado', 'percentual_frete_sobre_venda', 'frete_fracionado_por_kg', 'classificacao',
        'cor_matriz', 'regras_acionadas', 'oportunidade_nao_realizada', 'oportunidade_ja_dedicado',
        'economia_potencial', 'percentual_saving', 'origem_calculo_economia', 'explicacao_modal_estimado',
        'observacoes'
    ]
    colunas = [c for c in colunas_preferidas if c in df.columns]
    demais = [c for c in df.columns if c not in colunas]
    return df[colunas + demais].copy()


def criar_download_excel(df_resultado, df_oportunidades=None, resumo=None):
    '''Cria arquivo Excel em memória para download.'''
    output = io.BytesIO()

    if df_oportunidades is None:
        df_oportunidades = pd.DataFrame()
    if resumo is None:
        resumo = {}

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        pd.DataFrame([resumo]).to_excel(writer, index=False, sheet_name='Resumo')
        df_resultado.to_excel(writer, index=False, sheet_name='Analise_Completa')
        df_oportunidades.to_excel(writer, index=False, sheet_name='Oportunidades')

        workbook = writer.book
        header_format = workbook.add_format(
            {
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': PARKER_RED,
                'font_color': 'white',
                'border': 1,
            }
        )
        money_format = workbook.add_format({'num_format': 'USD #,##0.00'})
        percent_format = workbook.add_format({'num_format': '0.00%'})

        for sheet_name in ['Resumo', 'Analise_Completa', 'Oportunidades']:
            worksheet = writer.sheets[sheet_name]
            df_sheet = pd.DataFrame([resumo]) if sheet_name == 'Resumo' else (
                df_resultado if sheet_name == 'Analise_Completa' else df_oportunidades
            )

            for col_num, col_name in enumerate(df_sheet.columns):
                worksheet.write(0, col_num, col_name, header_format)
                largura = min(max(len(str(col_name)) + 2, 14), 48)
                worksheet.set_column(col_num, col_num, largura)

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
    return f"<span class='{classe}'>{classificacao}</span>"


def aplicar_filtro_multiselect(df, coluna, rotulo):
    '''Aplica filtro multiselect quando a coluna existir e tiver valores úteis.'''
    if coluna not in df.columns:
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


def safe_size_column(df, col, default=10):
    '''Cria série de tamanho sempre positiva e limitada para uso no parâmetro size do Plotly.'''
    if df is None or df.empty or col not in df.columns:
        return pd.Series(default, index=df.index if df is not None else None, dtype=float)
    tamanho = safe_numeric(df[col]).fillna(default)
    tamanho = tamanho.where(tamanho > 0, default)
    return tamanho.clip(lower=5, upper=60)


def safe_plot_df(df, required_cols):
    '''Retorna cópia segura para gráficos quando o DataFrame e as colunas obrigatórias existem.'''
    if df is None or df.empty:
        return pd.DataFrame()
    colunas_necessarias = [col for col in required_cols if col]
    if any(col not in df.columns for col in colunas_necessarias):
        return pd.DataFrame()
    return df.copy()


# =============================================================================
# Cabeçalho corporativo e sidebar informativa
# =============================================================================

st.sidebar.title('Parker-Hannifin')
st.sidebar.caption('Branding e matriz fixa de decisão logística.')
logo_upload = st.sidebar.file_uploader('Logo opcional (.png/.jpg)', type=['png', 'jpg', 'jpeg'])
st.sidebar.markdown('---')
st.sidebar.info(
    'Os parâmetros da regra não são editáveis no aplicativo. A matriz está fixa no código conforme política informada.'
)

logo_local = 'assets/parker_logo.png'
logo_source = None
if logo_upload is not None:
    logo_source = logo_upload
elif os.path.exists(logo_local):
    logo_source = logo_local

header_logo, header_text = st.columns([1, 5])
with header_logo:
    if logo_source is not None:
        st.image(logo_source, use_container_width=True)
    else:
        st.markdown(
            f'''
            <div style='height:84px;border:1px solid #E5E7EB;border-radius:12px;display:flex;align-items:center;justify-content:center;color:{PARKER_RED};font-weight:900;'>
                PARKER
            </div>
            ''',
            unsafe_allow_html=True,
        )
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
                    <li>Peso >= 3000 kg; ou</li>
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
                    <li>Peso entre 500 e 999 kg e frete >= USD 250; ou</li>
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
                    <li>Peso < 500 kg;</li>
                    <li>Frete < USD 250;</li>
                    <li>Frete/Venda < 8%;</li>
                    <li>Salvo exceção operacional.</li>
                </ul>
            </div>
            ''',
            unsafe_allow_html=True,
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
        'quando ambos forem preenchidos, o app calcula economia e percentual de saving.'
    )

    with st.form('form_simulador_manual'):
        col1, col2, col3 = st.columns(3)

        with col1:
            valor_mercadoria = st.number_input(
                'Valor da mercadoria (USD)', min_value=0.0, value=50000.0, step=1000.0, format='%.2f'
            )
            peso = st.number_input('Peso bruto (kg)', min_value=0.0, value=1000.0, step=100.0, format='%.2f')
            frete_fracionado_txt = st.text_input(
                'Frete fracionado estimado (opcional, USD)', value='', placeholder='Ex.: 500.00'
            )

        with col2:
            frete_dedicado_txt = st.text_input(
                'Frete dedicado cotado (opcional, USD)', value='', placeholder='Ex.: 420.00'
            )
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

        df_manual = pd.DataFrame(
            [
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
            ]
        )

        resultado_manual = aplicar_regras(df_manual).iloc[0]
        tem_financeiro = pd.notna(resultado_manual['frete_fracionado']) and pd.notna(resultado_manual['frete_dedicado'])
        economia_manual = max(resultado_manual['frete_fracionado'] - resultado_manual['frete_dedicado'], 0) if tem_financeiro else np.nan
        saving_manual = economia_manual / resultado_manual['frete_fracionado'] if tem_financeiro and resultado_manual['frete_fracionado'] > 0 else np.nan

        st.markdown('### Resultado da Simulação')
        st.markdown(badge_classificacao(resultado_manual['classificacao']), unsafe_allow_html=True)
        st.write('')

        m1, m2, m3, m4 = st.columns(4)
        with m1:
            kpi_card('Frete/Venda', formatar_percentual(resultado_manual['percentual_frete_sobre_venda']), 'Calculado se frete e venda existirem')
        with m2:
            kpi_card('R$/kg ou USD/kg', formatar_moeda(resultado_manual['frete_fracionado_por_kg']), 'Calculado se frete e peso existirem')
        with m3:
            kpi_card('Economia', formatar_moeda(economia_manual) if tem_financeiro else 'N/D', 'Exige frete fracionado e dedicado')
        with m4:
            kpi_card('Saving', formatar_percentual(saving_manual) if tem_financeiro else 'N/D', 'Exige frete fracionado e dedicado')

        if resultado_manual['classificacao'] == 'Cotar dedicado obrigatório':
            st.error('🔴 **Cotar dedicado obrigatório** — o embarque acionou pelo menos um gatilho vermelho da matriz.')
        elif resultado_manual['classificacao'] == 'Avaliar dedicado':
            st.warning('🟡 **Avaliar dedicado** — existe gatilho amarelo ou exceção operacional que recomenda avaliação.')
        else:
            st.success('🟢 **Manter fracionado** — não foram identificados gatilhos para cotação dedicada, salvo decisão operacional.')

        if not tem_financeiro:
            st.info('Frete fracionado e/ou frete dedicado não preenchido. Resultado exibido apenas como recomendação operacional, sem cálculo financeiro.')

        detalhes = pd.DataFrame(
            {
                'Indicador': [
                    'Origem', 'Destino', 'Transportadora', 'Valor da mercadoria', 'Peso bruto',
                    'Frete fracionado informado', 'Frete dedicado informado', 'Regras acionadas',
                    'Status estimado do modal', 'Explicação da heurística'
                ],
                'Valor': [
                    resultado_manual['origem'],
                    resultado_manual['destino'],
                    resultado_manual['transportadora'] or 'Não informado',
                    formatar_moeda(resultado_manual['valor_mercadoria']),
                    formatar_numero(resultado_manual['peso'], ' kg'),
                    'Não informado' if pd.isna(resultado_manual['frete_fracionado']) else formatar_moeda(resultado_manual['frete_fracionado']),
                    'Não informado' if pd.isna(resultado_manual['frete_dedicado']) else formatar_moeda(resultado_manual['frete_dedicado']),
                    resultado_manual['regras_acionadas'],
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
        'Faça upload de um arquivo Excel ou CSV. O aplicativo tenta detectar automaticamente as colunas prováveis '
        'em português ou inglês, mas permite ajuste manual quando necessário.'
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

        colunas_disponiveis = ['Não disponível'] + list(df_original.columns)

        sugestoes = {
            'documento': sugerir_coluna(df_original.columns, ['nf', 'nota', 'documento', 'doc', 'cte', 'pedido', 'shipment', 'invoice']),
            'modalidade': sugerir_coluna(df_original.columns, ['modalidade', 'tipo frete', 'tipo_frete', 'frete tipo', 'serviço', 'servico', 'mode', 'modal']),
            'valor_mercadoria': sugerir_coluna(df_original.columns, ['valor mercadoria', 'vl mercadoria', 'valor nf', 'valor_nf', 'mercadoria', 'venda', 'sales', 'invoice value', 'goods value']),
            'peso': sugerir_coluna(df_original.columns, ['peso bruto', 'peso', 'kg', 'weight', 'gross weight']),
            'frete_fracionado': sugerir_coluna(df_original.columns, ['valor frete', 'vl frete', 'frete', 'freight', 'estimated freight', 'frete estimado']),
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

        st.markdown('### Mapeamento de Colunas')
        st.caption('Selecione **Não disponível** quando a informação não existir no arquivo.')

        mapa_colunas = {}
        c1, c2, c3 = st.columns(3)

        campos = [
            ('documento', 'Chave NF / Documento'),
            ('modalidade', 'Modalidade / Tipo de frete'),
            ('valor_mercadoria', 'Valor da mercadoria / venda'),
            ('peso', 'Peso bruto'),
            ('frete_fracionado', 'Frete fracionado / frete estimado'),
            ('frete_dedicado', 'Frete dedicado cotado'),
            ('origem', 'Origem'),
            ('destino', 'Destino'),
            ('estado', 'Estado / UF'),
            ('data', 'Data'),
            ('transportadora', 'Transportadora'),
            ('divisao', 'Divisão'),
            ('operacao_direcao', 'Operação / Direção'),
            ('cliente', 'Cliente'),
            ('observacoes', 'Observações'),
        ]

        for i, (campo, rotulo) in enumerate(campos):
            coluna_container = [c1, c2, c3][i % 3]
            with coluna_container:
                indice_sugerido = colunas_disponiveis.index(sugestoes[campo]) if sugestoes[campo] in colunas_disponiveis else 0
                mapa_colunas[campo] = st.selectbox(rotulo, options=colunas_disponiveis, index=indice_sugerido, key=f'mapa_{campo}')

        col_obrig_1, col_obrig_2, col_obrig_3 = st.columns(3)
        with col_obrig_1:
            st.caption('Campo crítico: Peso')
            if mapa_colunas['peso'] == 'Não disponível':
                st.warning('Sem peso, as regras por faixa de kg não serão efetivas.')
        with col_obrig_2:
            st.caption('Campo crítico: Frete estimado')
            if mapa_colunas['frete_fracionado'] == 'Não disponível':
                st.warning('Sem frete, as regras USD 500, USD 250 e Frete/Venda não serão efetivas.')
        with col_obrig_3:
            st.caption('Campo crítico: Valor da mercadoria')
            if mapa_colunas['valor_mercadoria'] == 'Não disponível':
                st.warning('Sem valor da mercadoria, as regras Peso + Valor e Frete/Venda não serão efetivas.')

        analisar = st.button('Aplicar matriz no banco', type='primary')

        if analisar:
            df_padrao = construir_df_padronizado(df_original, mapa_colunas)
            df_resultado = aplicar_regras(df_padrao)
            df_exibicao = preparar_tabela_exibicao(df_resultado)
            st.session_state['df_resultado_lote'] = df_resultado
            st.session_state['df_exibicao_lote'] = df_exibicao

        if 'df_resultado_lote' in st.session_state:
            df_resultado = st.session_state['df_resultado_lote'].copy()
            df_exibicao = st.session_state['df_exibicao_lote'].copy()

            st.markdown('### Filtros Operacionais')
            with st.expander('Filtrar análise', expanded=True):
                f1, f2, f3, f4 = st.columns(4)
                with f1:
                    df_filtrado = aplicar_filtro_multiselect(df_resultado, 'divisao', 'Divisão')
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
                        df_filtrado = df_filtrado[
                            (df_filtrado['data_convertida'].dt.date >= inicio) &
                            (df_filtrado['data_convertida'].dt.date <= fim)
                        ]
            df_exibicao_filtrado = preparar_tabela_exibicao(df_filtrado)

            oportunidades = df_filtrado[df_filtrado['oportunidade_dedicado']].copy()
            oportunidades_nao_realizadas = df_filtrado[df_filtrado['oportunidade_nao_realizada']].copy()
            oportunidades_ja_dedicado = df_filtrado[df_filtrado['oportunidade_ja_dedicado']].copy()
            df_oportunidades_exibicao = preparar_tabela_exibicao(oportunidades).sort_values(
                by=['cor_matriz', 'economia_potencial', 'frete_fracionado', 'valor_mercadoria'],
                ascending=[False, False, False, False],
            )

            total_analisado = len(df_filtrado)
            provavel_fracionado = int(df_filtrado['provavel_fracionado'].sum())
            provavel_dedicado = int(df_filtrado['provavel_dedicado'].sum())
            qtd_oportunidades_nao_realizadas = int(df_filtrado['oportunidade_nao_realizada'].sum())
            qtd_oportunidades_ja_dedicado = int(df_filtrado['oportunidade_ja_dedicado'].sum())
            gasto_fracionado_oportunidade = float(oportunidades_nao_realizadas['frete_fracionado'].sum(skipna=True))
            economia_potencial_total = float(oportunidades_nao_realizadas['economia_potencial'].sum(skipna=True))

            resumo = {
                'total_analisado': total_analisado,
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
                kpi_card('Total analisado', f'{total_analisado:,}'.replace(',', '.'), 'Após filtros aplicados')
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
            graf_tab1, graf_tab2, graf_tab3, graf_tab4 = st.tabs([
                'Resumo e oportunidades', 'Distribuições', 'Matriz peso x valor', 'Rotas e savings'
            ])

            with graf_tab1:
                g1, g2 = st.columns(2)
                with g1:
                    df_pie_base = safe_plot_df(df_filtrado, ['classificacao'])
                    if not df_pie_base.empty:
                        df_classificacao = (
                            df_pie_base.groupby('classificacao', dropna=False)
                            .size()
                            .reset_index(name='quantidade')
                            .sort_values('quantidade', ascending=False)
                        )
                        df_classificacao['quantidade'] = safe_numeric(df_classificacao['quantidade']).fillna(0)
                        df_classificacao = df_classificacao[df_classificacao['quantidade'] > 0]
                        if not df_classificacao.empty:
                            fig = px.pie(
                                df_classificacao,
                                names='classificacao',
                                values='quantidade',
                                title='Distribuição por recomendação operacional',
                                color='classificacao',
                                color_discrete_map={
                                    'Cotar dedicado obrigatório': PARKER_RED,
                                    'Avaliar dedicado': PARKER_YELLOW,
                                    'Manter fracionado': PARKER_GREEN,
                                },
                                hole=0.35,
                            )
                            fig.update_traces(textposition='inside', textinfo='percent+label')
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.info('Sem dados válidos de classificação para exibir.')
                    else:
                        st.info('Sem dados de classificação para exibir.')

                with g2:
                    df_div_base = safe_plot_df(oportunidades_nao_realizadas, ['divisao'])
                    if not df_div_base.empty:
                        df_div = (
                            df_div_base.assign(divisao=df_div_base['divisao'].replace('', 'Não informada'))
                            .groupby('divisao')
                            .size()
                            .reset_index(name='oportunidades')
                            .sort_values('oportunidades', ascending=False)
                            .head(15)
                        )
                        df_div['oportunidades'] = safe_numeric(df_div['oportunidades']).fillna(0)
                        df_div = df_div[df_div['oportunidades'] > 0]
                        if not df_div.empty:
                            fig = px.bar(df_div, x='divisao', y='oportunidades', title='Oportunidades por divisão', color_discrete_sequence=[PARKER_RED])
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.info('Sem oportunidades por divisão para exibir.')
                    else:
                        st.info('Sem oportunidades por divisão para exibir.')

                g3, g4 = st.columns(2)
                with g3:
                    df_transp_base = safe_plot_df(oportunidades_nao_realizadas, ['transportadora'])
                    if not df_transp_base.empty:
                        df_transp = (
                            df_transp_base.assign(transportadora=df_transp_base['transportadora'].replace('', 'Não informada'))
                            .groupby('transportadora')
                            .size()
                            .reset_index(name='oportunidades')
                            .sort_values('oportunidades', ascending=False)
                            .head(15)
                        )
                        df_transp['oportunidades'] = safe_numeric(df_transp['oportunidades']).fillna(0)
                        df_transp = df_transp[df_transp['oportunidades'] > 0]
                        if not df_transp.empty:
                            fig = px.bar(df_transp, x='oportunidades', y='transportadora', orientation='h', title='Oportunidades por transportadora', color_discrete_sequence=[PARKER_RED])
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.info('Sem oportunidades por transportadora para exibir.')
                    else:
                        st.info('Sem oportunidades por transportadora para exibir.')
                with g4:
                    df_origem_destino_base = safe_plot_df(oportunidades_nao_realizadas, ['origem', 'destino'])
                    if not df_origem_destino_base.empty:
                        df_origem_destino = (
                            df_origem_destino_base.groupby(['origem', 'destino'], dropna=False)
                            .size()
                            .reset_index(name='oportunidades')
                            .sort_values('oportunidades', ascending=False)
                            .head(15)
                        )
                        df_origem_destino['oportunidades'] = safe_numeric(df_origem_destino['oportunidades']).fillna(0)
                        df_origem_destino = df_origem_destino[df_origem_destino['oportunidades'] > 0]
                        if not df_origem_destino.empty:
                            df_origem_destino['origem_destino'] = df_origem_destino['origem'].replace('', 'Não informada') + ' → ' + df_origem_destino['destino'].replace('', 'Não informada')
                            fig = px.bar(df_origem_destino, x='oportunidades', y='origem_destino', orientation='h', title='Oportunidades por origem/destino', color_discrete_sequence=[PARKER_RED])
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.info('Sem oportunidades por origem/destino para exibir.')
                    else:
                        st.info('Sem oportunidades por origem/destino para exibir.')

            with graf_tab2:
                d1, d2, d3 = st.columns(3)
                with d1:
                    df_hist_peso = safe_plot_df(df_filtrado, ['peso', 'classificacao'])
                    if not df_hist_peso.empty:
                        df_hist_peso['peso'] = safe_numeric(df_hist_peso['peso'])
                        df_hist_peso = df_hist_peso.dropna(subset=['peso'])
                        if not df_hist_peso.empty:
                            fig = px.histogram(df_hist_peso, x='peso', nbins=30, title='Distribuição por peso', color='classificacao', color_discrete_map={'Cotar dedicado obrigatório': PARKER_RED, 'Avaliar dedicado': PARKER_YELLOW, 'Manter fracionado': PARKER_GREEN})
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.info('Sem dados válidos de peso para exibir.')
                    else:
                        st.info('Sem dados de peso para exibir.')
                with d2:
                    df_hist_valor = safe_plot_df(df_filtrado, ['valor_mercadoria', 'classificacao'])
                    if not df_hist_valor.empty:
                        df_hist_valor['valor_mercadoria'] = safe_numeric(df_hist_valor['valor_mercadoria'])
                        df_hist_valor = df_hist_valor.dropna(subset=['valor_mercadoria'])
                        if not df_hist_valor.empty:
                            fig = px.histogram(df_hist_valor, x='valor_mercadoria', nbins=30, title='Distribuição por valor de mercadoria', color='classificacao', color_discrete_map={'Cotar dedicado obrigatório': PARKER_RED, 'Avaliar dedicado': PARKER_YELLOW, 'Manter fracionado': PARKER_GREEN})
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.info('Sem dados válidos de valor de mercadoria para exibir.')
                    else:
                        st.info('Sem dados de valor de mercadoria para exibir.')
                with d3:
                    df_hist_frete = safe_plot_df(df_filtrado, ['frete_fracionado', 'classificacao'])
                    if not df_hist_frete.empty:
                        df_hist_frete['frete_fracionado'] = safe_numeric(df_hist_frete['frete_fracionado'])
                        df_hist_frete = df_hist_frete.dropna(subset=['frete_fracionado'])
                        if not df_hist_frete.empty:
                            fig = px.histogram(df_hist_frete, x='frete_fracionado', nbins=30, title='Distribuição por frete estimado', color='classificacao', color_discrete_map={'Cotar dedicado obrigatório': PARKER_RED, 'Avaliar dedicado': PARKER_YELLOW, 'Manter fracionado': PARKER_GREEN})
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.info('Sem dados válidos de frete estimado para exibir.')
                    else:
                        st.info('Sem dados de frete estimado para exibir.')

            with graf_tab3:
                df_scatter = safe_plot_df(df_filtrado, ['peso', 'valor_mercadoria'])
                if not df_scatter.empty:
                    df_scatter['peso'] = pd.to_numeric(df_scatter['peso'], errors='coerce').replace([np.inf, -np.inf], np.nan)
                    df_scatter['valor_mercadoria'] = pd.to_numeric(df_scatter['valor_mercadoria'], errors='coerce').replace([np.inf, -np.inf], np.nan)
                    if 'frete_fracionado' in df_scatter.columns:
                        df_scatter['frete_fracionado'] = pd.to_numeric(df_scatter['frete_fracionado'], errors='coerce').replace([np.inf, -np.inf], np.nan)
                    else:
                        df_scatter['frete_fracionado'] = np.nan
                    df_scatter = df_scatter.dropna(subset=['peso', 'valor_mercadoria'])
                    df_scatter['_plot_size'] = safe_size_column(df_scatter, 'frete_fracionado', default=10)
                    df_scatter = df_scatter[df_scatter['_plot_size'].notna() & (df_scatter['_plot_size'] > 0)]
                    if not df_scatter.empty:
                        hover_data = [
                            coluna for coluna in ['documento', 'transportadora', 'origem', 'destino', 'regras_acionadas', 'status_modal_estimado']
                            if coluna in df_scatter.columns
                        ]
                        fig = px.scatter(
                            df_scatter,
                            x='peso',
                            y='valor_mercadoria',
                            size='_plot_size',
                            color='classificacao' if 'classificacao' in df_scatter.columns else None,
                            hover_data=hover_data,
                            title='Matriz peso versus valor da mercadoria',
                            color_discrete_map={
                                'Cotar dedicado obrigatório': PARKER_RED,
                                'Avaliar dedicado': PARKER_YELLOW,
                                'Manter fracionado': PARKER_GREEN,
                            },
                        )
                        fig.add_vline(x=LIMITE_PESO_COTAR_OBRIGATORIO, line_dash='dash', line_color=PARKER_RED)
                        fig.add_vline(x=LIMITE_PESO_VALOR_COTAR, line_dash='dot', line_color=PARKER_YELLOW)
                        fig.add_hline(y=LIMITE_VALOR_MERCADORIA_COTAR, line_dash='dot', line_color=PARKER_YELLOW)
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info('Não há peso e valor de mercadoria suficientes para montar a matriz após limpeza dos dados.')
                else:
                    st.info('Não há peso e valor de mercadoria suficientes para montar a matriz.')

            with graf_tab4:
                r1, r2 = st.columns(2)
                with r1:
                    df_rotas_base = safe_plot_df(oportunidades_nao_realizadas, ['rota', 'economia_potencial', 'frete_fracionado'])
                    if not df_rotas_base.empty:
                        df_rotas_base['economia_potencial'] = safe_numeric(df_rotas_base['economia_potencial']).fillna(0)
                        df_rotas_base['frete_fracionado'] = safe_numeric(df_rotas_base['frete_fracionado']).fillna(0)
                        df_rotas = (
                            df_rotas_base.groupby('rota', dropna=False)
                            .agg(oportunidades=('rota', 'size'), economia=('economia_potencial', 'sum'), gasto=('frete_fracionado', 'sum'))
                            .reset_index()
                            .sort_values(['oportunidades', 'economia'], ascending=False)
                            .head(15)
                        )
                        df_rotas['oportunidades'] = safe_numeric(df_rotas['oportunidades']).fillna(0)
                        df_rotas['economia'] = safe_numeric(df_rotas['economia']).fillna(0)
                        df_rotas = df_rotas[df_rotas['oportunidades'] > 0]
                        if not df_rotas.empty:
                            fig = px.bar(df_rotas, x='oportunidades', y='rota', orientation='h', title='Top rotas com oportunidade', color='economia', color_continuous_scale=['#FDECEF', PARKER_RED])
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.info('Sem rotas com oportunidade para exibir.')
                    else:
                        st.info('Sem rotas com oportunidade para exibir.')
                with r2:
                    df_savings = safe_plot_df(oportunidades_nao_realizadas, ['economia_potencial'])
                    if not df_savings.empty:
                        df_savings['economia_potencial'] = safe_numeric(df_savings['economia_potencial'])
                        df_savings = df_savings[df_savings['economia_potencial'] > 0].copy()
                        if not df_savings.empty:
                            if 'documento' in df_savings.columns:
                                df_savings['documento_plot'] = df_savings['documento'].replace('', np.nan).fillna('Sem documento')
                            else:
                                df_savings['documento_plot'] = 'Sem documento'
                            df_savings = df_savings.sort_values('economia_potencial', ascending=False).head(15)
                            color_col = 'origem_calculo_economia' if 'origem_calculo_economia' in df_savings.columns else None
                            fig = px.bar(
                                df_savings.sort_values('economia_potencial'),
                                x='economia_potencial',
                                y='documento_plot',
                                orientation='h',
                                title='Ranking de savings',
                                color=color_col,
                                color_discrete_sequence=[PARKER_RED, PARKER_YELLOW, PARKER_GREEN],
                            )
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.info('Ranking de savings disponível somente quando houver frete informado ou simulação ativa.')
                    else:
                        st.info('Ranking de savings disponível somente quando houver frete informado ou simulação ativa.')

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
                        'percentual_frete_sobre_venda': st.column_config.NumberColumn('Frete/Venda', format='%.2f%%'),
                        'percentual_saving': st.column_config.NumberColumn('Saving', format='%.2f%%'),
                        'frete_fracionado_por_kg': st.column_config.NumberColumn('USD/kg', format='USD %.2f'),
                    },
                )

            st.markdown('### Download dos Resultados')
            excel_bytes = criar_download_excel(df_exibicao_filtrado, df_oportunidades_exibicao, resumo)
            csv_oportunidades = df_oportunidades_exibicao.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')
            csv_completo = df_exibicao_filtrado.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')

            d1, d2, d3 = st.columns(3)
            with d1:
                st.download_button(
                    'Baixar Excel completo',
                    data=excel_bytes,
                    file_name='parker_analise_frete_dedicado.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                )
            with d2:
                st.download_button(
                    'Baixar CSV de oportunidades',
                    data=csv_oportunidades,
                    file_name='parker_oportunidades_frete_dedicado.csv',
                    mime='text/csv',
                )
            with d3:
                st.download_button(
                    'Baixar CSV completo',
                    data=csv_completo,
                    file_name='parker_analise_completa_frete.csv',
                    mime='text/csv',
                )

            with st.expander('Ver análise completa filtrada', expanded=False):
                st.dataframe(df_exibicao_filtrado, use_container_width=True, hide_index=True)

    elif arquivo is not None:
        st.warning('O arquivo foi carregado, mas não possui dados para análise.')
    else:
        st.info('Envie um arquivo para iniciar a análise em lote.')


st.markdown('---')
st.caption(
    'Parker-Hannifin | Ferramenta interna de apoio à decisão logística. '
    'A matriz de decisão está fixa no código e não substitui validação operacional, contratação formal ou análise de nível de serviço.'
)
