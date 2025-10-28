import streamlit as st
import pandas as pd
import dropbox
from datetime import datetime, date, timedelta
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import calendar
import time

st.set_page_config(
    page_title="Processamento Salarial v2.0",
    page_icon="💰",
    layout="wide"
)

# ==================== CONFIGURAÇÕES ====================

DROPBOX_APP_KEY = st.secrets["DROPBOX_APP_KEY"]
DROPBOX_APP_SECRET = st.secrets["DROPBOX_APP_SECRET"]
DROPBOX_REFRESH_TOKEN = st.secrets["DROPBOX_REFRESH_TOKEN"]
ADMIN_PASSWORD = st.secrets.get("ADMIN_PASSWORD", "adminpedro")

dbx = dropbox.Dropbox(
    app_key=DROPBOX_APP_KEY,
    app_secret=DROPBOX_APP_SECRET,
    oauth2_refresh_token=DROPBOX_REFRESH_TOKEN
)

EMPRESAS = {
    "Magnetic Sky Lda": {
        "path": "/Pedro Couto/Projectos/Alcalá_Arc_Amoreira/Gestão operacional/RH/Processamento Salários Magnetic/Gestão Colaboradores Magnetic.xlsx",
        "docs_path": "/Pedro Couto/Projectos/Alcalá_Arc_Amoreira/Gestão operacional/RH/Documentos_Baixas_Magnetic",
        "tem_horas_extras": False
    },
    "CCM Retail Lda": {
        "path": "/Pedro Couto/Projectos/Pingo Doce/Pingo Doce/2. Operação/1. Recursos Humanos/Processamento salarial/Gestão Colaboradores.xlsx",
        "docs_path": "/Pedro Couto/Projectos/Pingo Doce/Pingo Doce/2. Operação/1. Recursos Humanos/Documentos_Baixas",
        "tem_horas_extras": True
    }
}

FERIADOS_NACIONAIS_2025 = [
    date(2025, 1, 1), date(2025, 4, 18), date(2025, 4, 20), date(2025, 4, 25),
    date(2025, 5, 1), date(2025, 6, 10), date(2025, 6, 19), date(2025, 8, 15),
    date(2025, 10, 5), date(2025, 11, 1), date(2025, 12, 1), date(2025, 12, 8),
    date(2025, 12, 25)
]

MOTIVOS_RESCISAO = [
    "Denúncia pela entidade patronal - período experimental",
    "Denúncia pelo trabalhador - período experimental",
    "Caducidade contrato a termo",
    "Denúncia pelo trabalhador - aviso prévio parcial",
    "Denúncia pelo trabalhador - aviso prévio completo",
    "Denúncia pelo trabalhador - sem aviso prévio",
    "Denúncia pela entidade patronal - excesso faltas",
    "Outro (especificar em observações)"
]

# ==================== SESSION STATE ====================

if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'salario_minimo' not in st.session_state:
    st.session_state.salario_minimo = 870.0
if 'feriados_municipais' not in st.session_state:
    st.session_state.feriados_municipais = [date(2025, 1, 14)]

# ==================== FUNÇÕES DE AUTENTICAÇÃO ====================

def check_password():
    def password_entered():
        if st.session_state["password"] == ADMIN_PASSWORD:
            st.session_state.authenticated = True
            del st.session_state["password"]
        else:
            st.session_state.authenticated = False
    
    if not st.session_state.authenticated:
        st.title("🔒 Processamento Salarial - Login")
        st.markdown("---")
        st.text_input("Password de Administrador", type="password", key="password", on_change=password_entered)
        if "password" in st.session_state and not st.session_state.authenticated:
            st.error("❌ Password incorreta")
        return False
    return True

# ==================== FUNÇÕES DE CACHE ====================

def invalidar_cache_completo(empresa=None):
    """Limpa TODOS os caches após gravação"""
    keys_to_delete = []
    if empresa:
        keys_to_delete = [
            f"df_colaboradores_{empresa}",
            f"df_colaboradores_completo_{empresa}",
            f"historico_horas_{empresa}",
            f"banco_horas_{empresa}",
            f"rescisoes_{empresa}",
        ]
    else:
        keys_to_delete = [k for k in st.session_state.keys() if any(x in k for x in ['colaboradores', 'historico', 'banco', 'rescisoes', 'processamento'])]
    
    for key in keys_to_delete:
        if key in st.session_state and key != 'authenticated':
            del st.session_state[key]

# ==================== FUNÇÕES DE GESTÃO DE ABAS EXCEL ====================

def garantir_aba(empresa, nome_aba, colunas):
    """Garante que uma aba existe no Excel com as colunas especificadas"""
    try:
        file_path = EMPRESAS[empresa]["path"]
        _, response = dbx.files_download(file_path)
        wb = load_workbook(BytesIO(response.content))
        
        if nome_aba not in wb.sheetnames:
            ws = wb.create_sheet(nome_aba)
            ws.append(colunas)
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            dbx.files_upload(output.read(), file_path, mode=dropbox.files.WriteMode.overwrite)
            return True
        return True
    except Exception as e:
        st.error(f"❌ Erro ao criar aba {nome_aba}: {e}")
        return False

def garantir_todas_abas(empresa):
    """Garante que todas as abas necessárias existem"""
    abas_necessarias = {
        "Config_Colaboradores": ["Nome Completo", "Subsídio Alimentação Diário", "Número Pingo Doce", "Última Atualização"],
        "Historico_Horas_Semana": ["Nome Completo", "Horas Semana", "Data Início Vigência", "Data Fim Vigência", "Registado Em"],
        "Banco_Horas_Acumulado": ["Nome Completo", "Ano", "Mês", "Banco Horas Mês", "Saldo Acumulado", "Registado Em"],
        "Rescisoes": ["Nome Completo", "Data Rescisão", "Motivo", "Observações", "Dias Aviso Prévio", "Registado Em"],
        "Baixas_Transitorias": ["Nome Completo", "Data Início", "Data Fim", "Mês Origem", "Status", "Registado Em"]
    }
    
    for aba, colunas in abas_necessarias.items():
        garantir_aba(empresa, aba, colunas)

# ==================== FUNÇÕES DE CARREGAMENTO ====================

def carregar_colaboradores(empresa, force_reload=False):
    """Carrega dados base dos colaboradores"""
    cache_key = f"df_colaboradores_{empresa}"
    if force_reload or cache_key not in st.session_state:
        try:
            _, response = dbx.files_download(EMPRESAS[empresa]["path"])
            df = pd.read_excel(BytesIO(response.content), sheet_name="Colaboradores")
            st.session_state[cache_key] = df
            return df
        except Exception as e:
            st.error(f"❌ Erro ao carregar colaboradores: {e}")
            return pd.DataFrame()
    return st.session_state[cache_key]

def carregar_aba_excel(empresa, nome_aba, force_reload=False):
    """Carrega qualquer aba do Excel com cache"""
    cache_key = f"{nome_aba}_{empresa}"
    if force_reload or cache_key not in st.session_state:
        try:
            _, response = dbx.files_download(EMPRESAS[empresa]["path"])
            df = pd.read_excel(BytesIO(response.content), sheet_name=nome_aba)
            st.session_state[cache_key] = df
            return df
        except:
            return pd.DataFrame()
    return st.session_state[cache_key]

def carregar_colaboradores_ativos(empresa, data_referencia, force_reload=False):
    """Carrega apenas colaboradores ativos (não rescindidos) na data de referência"""
    df_base = carregar_colaboradores(empresa, force_reload)
    df_rescisoes = carregar_aba_excel(empresa, "Rescisoes", force_reload)
    
    if df_rescisoes.empty:
        return df_base
    
    # Filtrar rescindidos antes da data de referência
    rescindidos = df_rescisoes[pd.to_datetime(df_rescisoes['Data Rescisão']) <= pd.to_datetime(data_referencia)]['Nome Completo'].tolist()
    df_ativos = df_base[~df_base['Nome Completo'].isin(rescindidos)]
    
    return df_ativos

def carregar_horas_vigentes(empresa, colaborador, data_referencia):
    """Carrega horas/semana válidas para o colaborador na data de referência"""
    df_historico = carregar_aba_excel(empresa, "Historico_Horas_Semana")
    
    if df_historico.empty:
        # Sem histórico, usar dados base
        df_base = carregar_colaboradores(empresa)
        colab_data = df_base[df_base['Nome Completo'] == colaborador]
        if not colab_data.empty:
            return float(colab_data.iloc[0]['Nº Horas/Semana'])
        return 40.0
    
    # Filtrar histórico do colaborador
    df_colab = df_historico[df_historico['Nome Completo'] == colaborador].copy()
    df_colab['Data Início Vigência'] = pd.to_datetime(df_colab['Data Início Vigência'])
    df_colab['Data Fim Vigência'] = pd.to_datetime(df_colab['Data Fim Vigência'])
    
    # Encontrar registo válido
    data_ref = pd.to_datetime(data_referencia)
    df_valido = df_colab[
        (df_colab['Data Início Vigência'] <= data_ref) &
        ((df_colab['Data Fim Vigência'].isna()) | (df_colab['Data Fim Vigência'] >= data_ref))
    ]
    
    if not df_valido.empty:
        return float(df_valido.iloc[-1]['Horas Semana'])
    
    # Se não encontrou no histórico, buscar dados base
    df_base = carregar_colaboradores(empresa)
    colab_data = df_base[df_base['Nome Completo'] == colaborador]
    if not colab_data.empty:
        return float(colab_data.iloc[0]['Nº Horas/Semana'])
    
    return 40.0

def carregar_subsidio_vigente(empresa, colaborador):
    """Carrega subsídio alimentação do colaborador"""
    df_config = carregar_aba_excel(empresa, "Config_Colaboradores")
    
    if not df_config.empty:
        config = df_config[df_config['Nome Completo'] == colaborador]
        if not config.empty:
            return float(config.iloc[0]['Subsídio Alimentação Diário'])
    
    # Buscar dados base
    df_base = carregar_colaboradores(empresa)
    colab_data = df_base[df_base['Nome Completo'] == colaborador]
    if not colab_data.empty:
        return float(colab_data.iloc[0].get('Subsídio Alimentação Diário', 5.96))
    
    return 5.96

def carregar_numero_pingo_doce(empresa, colaborador):
    """Carrega número Pingo Doce do colaborador"""
    df_config = carregar_aba_excel(empresa, "Config_Colaboradores")
    
    if not df_config.empty:
        config = df_config[df_config['Nome Completo'] == colaborador]
        if not config.empty and 'Número Pingo Doce' in config.columns:
            num = config.iloc[0]['Número Pingo Doce']
            return str(num) if pd.notna(num) else ""
    return ""

def carregar_banco_horas_acumulado(empresa, colaborador, ate_mes, ate_ano):
    """Carrega saldo acumulado do banco de horas até determinado mês"""
    df_banco = carregar_aba_excel(empresa, "Banco_Horas_Acumulado")
    
    if df_banco.empty:
        return 0.0
    
    df_colab = df_banco[df_banco['Nome Completo'] == colaborador].copy()
    df_colab['Data'] = pd.to_datetime(df_colab['Ano'].astype(str) + '-' + df_colab['Mês'].astype(str) + '-01')
    data_limite = pd.to_datetime(f"{ate_ano}-{ate_mes:02d}-01")
    
    df_filtrado = df_colab[df_colab['Data'] < data_limite]
    
    if df_filtrado.empty:
        return 0.0
    
    return df_filtrado['Saldo Acumulado'].iloc[-1] if not df_filtrado.empty else 0.0

def carregar_baixas_transitorias(empresa, colaborador, mes, ano):
    """Carrega baixas que transitaram do mês anterior"""
    df_baixas = carregar_aba_excel(empresa, "Baixas_Transitorias")
    
    if df_baixas.empty:
        return []
    
    df_colab = df_baixas[
        (df_baixas['Nome Completo'] == colaborador) &
        (df_baixas['Status'] == 'Ativa')
    ].copy()
    
    baixas_no_mes = []
    for _, row in df_colab.iterrows():
        data_inicio = pd.to_datetime(row['Data Início']).date()
        data_fim = pd.to_datetime(row['Data Fim']).date()
        primeiro_dia_mes = date(ano, mes, 1)
        ultimo_dia_mes = date(ano, mes, calendar.monthrange(ano, mes)[1])
        
        # Se a baixa intercepta o mês atual
        if data_inicio <= ultimo_dia_mes and data_fim >= primeiro_dia_mes:
            inicio_no_mes = max(data_inicio, primeiro_dia_mes)
            fim_no_mes = min(data_fim, ultimo_dia_mes)
            baixas_no_mes.append((inicio_no_mes, fim_no_mes))
    
    return baixas_no_mes

# ==================== FUNÇÕES DE GRAVAÇÃO ====================

def gravar_em_aba(empresa, nome_aba, dados_dict):
    """Grava ou atualiza linha numa aba do Excel"""
    try:
        garantir_aba(empresa, nome_aba, list(dados_dict.keys()))
        
        file_path = EMPRESAS[empresa]["path"]
        _, response = dbx.files_download(file_path)
        wb = load_workbook(BytesIO(response.content))
        ws = wb[nome_aba]
        
        # Adicionar nova linha
        nova_linha = list(dados_dict.values())
        ws.append(nova_linha)
        
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        dbx.files_upload(output.read(), file_path, mode=dropbox.files.WriteMode.overwrite)
        
        return True
    except Exception as e:
        st.error(f"❌ Erro ao gravar em {nome_aba}: {e}")
        return False

def atualizar_subsidio_colaborador(empresa, colaborador, novo_valor):
    """Atualiza subsídio alimentação na aba Config_Colaboradores"""
    try:
        garantir_aba(empresa, "Config_Colaboradores", ["Nome Completo", "Subsídio Alimentação Diário", "Número Pingo Doce", "Última Atualização"])
        
        file_path = EMPRESAS[empresa]["path"]
        _, response = dbx.files_download(file_path)
        wb = load_workbook(BytesIO(response.content))
        ws = wb["Config_Colaboradores"]
        
        colaborador_row = None
        for row in range(2, ws.max_row + 1):
            if ws.cell(row, 1).value == colaborador:
                colaborador_row = row
                break
        
        if colaborador_row is None:
            colaborador_row = ws.max_row + 1
            ws.cell(colaborador_row, 1).value = colaborador
        
        ws.cell(colaborador_row, 2).value = float(novo_valor)
        ws.cell(colaborador_row, 4).value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        dbx.files_upload(output.read(), file_path, mode=dropbox.files.WriteMode.overwrite)
        
        return True
    except Exception as e:
        st.error(f"❌ Erro ao atualizar subsídio: {e}")
        return False

def atualizar_numero_pingo_doce(empresa, colaborador, numero):
    """Atualiza número Pingo Doce na aba Config_Colaboradores"""
    try:
        garantir_aba(empresa, "Config_Colaboradores", ["Nome Completo", "Subsídio Alimentação Diário", "Número Pingo Doce", "Última Atualização"])
        
        file_path = EMPRESAS[empresa]["path"]
        _, response = dbx.files_download(file_path)
        wb = load_workbook(BytesIO(response.content))
        ws = wb["Config_Colaboradores"]
        
        colaborador_row = None
        for row in range(2, ws.max_row + 1):
            if ws.cell(row, 1).value == colaborador:
                colaborador_row = row
                break
        
        if colaborador_row is None:
            colaborador_row = ws.max_row + 1
            ws.cell(colaborador_row, 1).value = colaborador
        
        ws.cell(colaborador_row, 3).value = str(numero)
        ws.cell(colaborador_row, 4).value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        dbx.files_upload(output.read(), file_path, mode=dropbox.files.WriteMode.overwrite)
        
        return True
    except Exception as e:
        st.error(f"❌ Erro ao atualizar número: {e}")
        return False

def registar_mudanca_horas(empresa, colaborador, horas_novas, data_inicio):
    """Registra mudança de horas/semana"""
    dados = {
        "Nome Completo": colaborador,
        "Horas Semana": horas_novas,
        "Data Início Vigência": data_inicio.strftime("%Y-%m-%d"),
        "Data Fim Vigência": None,
        "Registado Em": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    return gravar_em_aba(empresa, "Historico_Horas_Semana", dados)

def registar_rescisao(empresa, colaborador, data_rescisao, motivo, obs, dias_aviso):
    """Registra rescisão de colaborador"""
    dados = {
        "Nome Completo": colaborador,
        "Data Rescisão": data_rescisao.strftime("%Y-%m-%d"),
        "Motivo": motivo,
        "Observações": obs,
        "Dias Aviso Prévio": dias_aviso,
        "Registado Em": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    return gravar_em_aba(empresa, "Rescisoes", dados)

def registar_banco_horas(empresa, colaborador, ano, mes, banco_mes):
    """Registra banco de horas mensal e calcula acumulado"""
    saldo_anterior = carregar_banco_horas_acumulado(empresa, colaborador, mes, ano)
    saldo_novo = saldo_anterior + banco_mes
    
    dados = {
        "Nome Completo": colaborador,
        "Ano": ano,
        "Mês": mes,
        "Banco Horas Mês": banco_mes,
        "Saldo Acumulado": saldo_novo,
        "Registado Em": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    return gravar_em_aba(empresa, "Banco_Horas_Acumulado", dados)

def upload_documentos_baixa(empresa, colaborador, ano, mes, periodo_idx, files):
    """Faz upload de documentos de baixa para Dropbox"""
    try:
        docs_path = EMPRESAS[empresa]["docs_path"]
        caminho_base = f"{docs_path}/{ano}/{mes:02d}/{colaborador}/periodo_{periodo_idx}"
        
        links = []
        for idx, file in enumerate(files):
            nome_arquivo = f"baixa_{idx+1}_{file.name}"
            caminho_completo = f"{caminho_base}/{nome_arquivo}"
            
            dbx.files_upload(file.read(), caminho_completo, mode=dropbox.files.WriteMode.overwrite)
            links.append(caminho_completo)
        
        return links
    except Exception as e:
        st.error(f"❌ Erro ao fazer upload: {e}")
        return []

# ==================== FUNÇÕES DE CÁLCULO ====================

def calcular_dias_uteis(ano, mes, feriados_list):
    """Calcula dias úteis do mês"""
    num_dias = calendar.monthrange(ano, mes)[1]
    dias_uteis = 0
    for dia in range(1, num_dias + 1):
        data = date(ano, mes, dia)
        if data.weekday() < 5 and data not in feriados_list:
            dias_uteis += 1
    return dias_uteis

def calcular_dias_periodo(data_inicio, data_fim, apenas_uteis=False, feriados_list=None):
    """Calcula número de dias num período"""
    if data_inicio > data_fim:
        return 0
    
    dias = 0
    data_atual = data_inicio
    while data_atual <= data_fim:
        if apenas_uteis:
            if data_atual.weekday() < 5 and (feriados_list is None or data_atual not in feriados_list):
                dias += 1
        else:
            dias += 1
        data_atual += timedelta(days=1)
    
    return dias

def calcular_salario_base(horas_semana, salario_minimo):
    """Calcula salário base de acordo com horas/semana"""
    if horas_semana == 40:
        return salario_minimo
    elif horas_semana == 20:
        return salario_minimo / 2
    elif horas_semana == 16:
        return salario_minimo * 0.4
    return salario_minimo * (horas_semana / 40)

def calcular_vencimento_hora(salario_bruto, horas_semana):
    """Calcula vencimento por hora"""
    return (salario_bruto * 12) / (52 * horas_semana)

def calcular_dias_trabalhados_com_admissao(mes, ano, data_admissao, total_faltas, total_baixas):
    """Calcula dias trabalhados considerando data de admissão"""
    dias_no_mes = calendar.monthrange(ano, mes)[1]
    
    # Se admissão no mês
    if data_admissao.month == mes and data_admissao.year == ano:
        primeiro_dia_trabalho = data_admissao.day
        dias_possiveis = dias_no_mes - primeiro_dia_trabalho + 1
    else:
        dias_possiveis = dias_no_mes
    
    dias_trabalhados = dias_possiveis - total_faltas - total_baixas
    return max(dias_trabalhados, 0)

def processar_calculo_salario(dados_form):
    """Processa todos os cálculos salariais"""
    resultado = {}
    
    # Dados base
    salario_bruto = dados_form['salario_bruto']
    horas_semana = dados_form['horas_semana']
    sub_alimentacao_dia = dados_form['subsidio_alimentacao']
    vencimento_hora = calcular_vencimento_hora(salario_bruto, horas_semana)
    
    # Dias
    dias_uteis_mes = dados_form['dias_uteis_mes']
    dias_trabalhados = dados_form['dias_trabalhados']
    dias_uteis_trabalhados = dados_form['dias_uteis_trabalhados']
    
    # Horas extras
    horas_noturnas = dados_form.get('horas_noturnas', 0)
    horas_domingos = dados_form.get('horas_domingos', 0)
    horas_feriados = dados_form.get('horas_feriados', 0)
    horas_extra = dados_form.get('horas_extra', 0)
    
    # ===== REMUNERAÇÕES =====
    vencimento_ajustado = (salario_bruto / 30) * dias_trabalhados
    sub_alimentacao = sub_alimentacao_dia * dias_uteis_trabalhados
    trabalho_noturno = horas_noturnas * vencimento_hora * 0.25
    domingos = horas_domingos * vencimento_hora
    feriados = horas_feriados * vencimento_hora * 2
    
    # Subsídios Férias e Natal
    if dados_form['sub_ferias_tipo'] == 'Total':
        sub_ferias = salario_bruto
    else:
        sub_ferias = salario_bruto / 12
    
    if dados_form['sub_natal_tipo'] == 'Total':
        sub_natal = salario_bruto
    else:
        sub_natal = salario_bruto / 12
    
    banco_horas_valor = vencimento_hora * horas_extra
    outros_proveitos = dados_form.get('outros_proveitos', 0)
    
    total_remuneracoes = (vencimento_ajustado + sub_alimentacao + trabalho_noturno + 
                          domingos + feriados + sub_ferias + sub_natal + 
                          banco_horas_valor + outros_proveitos)
    
    # ===== DESCONTOS =====
    base_ss = total_remuneracoes - sub_alimentacao
    seg_social = base_ss * 0.11
    
    # IRS (simplificado - depois implementar tabela)
    irs = base_ss * 0.10  # Placeholder
    
    desconto_especie = sub_alimentacao if dados_form.get('desconto_especie', False) else 0
    
    total_descontos = seg_social + irs + desconto_especie
    
    # ===== LÍQUIDO =====
    liquido = total_remuneracoes - total_descontos
    
    resultado = {
        'vencimento_hora': vencimento_hora,
        'vencimento_ajustado': vencimento_ajustado,
        'sub_alimentacao': sub_alimentacao,
        'trabalho_noturno': trabalho_noturno,
        'domingos': domingos,
        'feriados': feriados,
        'sub_ferias': sub_ferias,
        'sub_natal': sub_natal,
        'banco_horas_valor': banco_horas_valor,
        'outros_proveitos': outros_proveitos,
        'total_remuneracoes': total_remuneracoes,
        'base_ss': base_ss,
        'seg_social': seg_social,
        'irs': irs,
        'desconto_especie': desconto_especie,
        'total_descontos': total_descontos,
        'liquido': liquido
    }
    
    return resultado

# ==================== INTERFACE ====================

if not check_password():
    st.stop()

st.title("💰 Processamento Salarial v2.0")
st.markdown("---")

menu = st.sidebar.radio(
    "Menu Principal",
    ["⚙️ Configurações", "💼 Processar Salários", "🚪 Rescisões", "📊 Relatórios"],
    index=0
)

# ==================== CONFIGURAÇÕES ====================

if menu == "⚙️ Configurações":
    st.header("⚙️ Configurações do Sistema")
    
    tab1, tab2, tab3 = st.tabs(["💶 Sistema", "👥 Colaboradores", "⏰ Horários"])
    
    # TAB 1: SISTEMA
    with tab1:
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("💶 Salário Mínimo Nacional")
            novo_salario = st.number_input(
                "Valor atual (€)",
                min_value=0.0,
                value=st.session_state.salario_minimo,
                step=10.0,
                format="%.2f"
            )
            if st.button("💾 Atualizar Salário Mínimo"):
                st.session_state.salario_minimo = novo_salario
                st.success(f"✅ Salário mínimo atualizado para {novo_salario}€")
        
        with col2:
            st.subheader("📅 Feriados Municipais")
            st.caption("Adicione até 3 feriados municipais")
            feriados_temp = []
            for i in range(3):
                valor_default = st.session_state.feriados_municipais[i] if i < len(st.session_state.feriados_municipais) else None
                feriado = st.date_input(
                    f"Feriado Municipal {i+1}",
                    value=valor_default,
                    key=f"feriado_mun_{i}"
                )
                if feriado:
                    feriados_temp.append(feriado)
            
            if st.button("💾 Atualizar Feriados"):
                st.session_state.feriados_municipais = feriados_temp
                st.success(f"✅ {len(feriados_temp)} feriados municipais configurados")
    
    # TAB 2: COLABORADORES
    with tab2:
        st.subheader("👥 Editar Dados de Colaboradores")
        
        empresa_sel = st.selectbox(
            "Selecione a Empresa",
            options=list(EMPRESAS.keys()),
            key="empresa_config_colab"
        )
        
        garantir_todas_abas(empresa_sel)
        df_colab = carregar_colaboradores(empresa_sel, force_reload=True)
        
        if not df_colab.empty:
            colaborador_sel = st.selectbox(
                "Selecione o Colaborador",
                options=df_colab['Nome Completo'].tolist(),
                key="colab_sel_config"
            )
            
            dados_base = df_colab[df_colab['Nome Completo'] == colaborador_sel].iloc[0]
            sub_atual = carregar_subsidio_vigente(empresa_sel, colaborador_sel)
            num_pingo_atual = carregar_numero_pingo_doce(empresa_sel, colaborador_sel)
            
            st.markdown("---")
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("📊 Subsídio Atual", f"{sub_atual:.2f}€")
            with col2:
                st.metric("🔢 Nº Pingo Doce", num_pingo_atual if num_pingo_atual else "N/A")
            
            with st.form("form_editar_colab"):
                st.markdown(f"### ✏️ Editar: {colaborador_sel}")
                
                col1, col2 = st.columns(2)
                with col1:
                    novo_sub = st.number_input(
                        "Subsídio Alimentação Diário (€)",
                        min_value=0.0,
                        value=float(sub_atual),
                        step=0.10,
                        format="%.2f"
                    )
                
                with col2:
                    novo_num_pingo = st.text_input(
                        "Número Pingo Doce",
                        value=num_pingo_atual
                    )
                
                submit = st.form_submit_button("💾 Guardar Alterações", use_container_width=True)
                
                if submit:
                    with st.spinner("🔄 A guardar alterações..."):
                        sucesso_sub = atualizar_subsidio_colaborador(empresa_sel, colaborador_sel, novo_sub)
                        sucesso_num = atualizar_numero_pingo_doce(empresa_sel, colaborador_sel, novo_num_pingo)
                        
                        if sucesso_sub and sucesso_num:
                            invalidar_cache_completo(empresa_sel)
                            st.success(f"✅ Dados atualizados com sucesso!")
                            st.balloons()
                            time.sleep(1)
                            st.rerun()
    
    # TAB 3: HORÁRIOS
    with tab3:
        st.subheader("⏰ Mudanças de Horário")
        st.info("📝 Registre aqui quando um colaborador muda de horário (ex: 20h → 40h)")
        
        empresa_sel_h = st.selectbox(
            "Selecione a Empresa",
            options=list(EMPRESAS.keys()),
            key="empresa_config_horas"
        )
        
        df_colab_h = carregar_colaboradores(empresa_sel_h)
        
        if not df_colab_h.empty:
            with st.form("form_mudanca_horas"):
                colaborador_h = st.selectbox(
                    "Colaborador",
                    options=df_colab_h['Nome Completo'].tolist()
                )
                
                dados_colab_h = df_colab_h[df_colab_h['Nome Completo'] == colaborador_h].iloc[0]
                horas_atuais = dados_colab_h['Nº Horas/Semana']
                
                st.info(f"⏰ Horário atual: **{horas_atuais}h/semana**")
                
                col1, col2 = st.columns(2)
                with col1:
                    novas_horas = st.selectbox(
                        "Novo Horário (h/semana)",
                        options=[16, 20, 40],
                        index=2
                    )
                
                with col2:
                    data_inicio_vigor = st.date_input(
                        "Data Início Vigência",
                        value=date.today()
                    )
                
                submit_horas = st.form_submit_button("💾 Registar Mudança", use_container_width=True)
                
                if submit_horas:
                    with st.spinner("🔄 A registar mudança..."):
                        sucesso = registar_mudanca_horas(empresa_sel_h, colaborador_h, novas_horas, data_inicio_vigor)
                        if sucesso:
                            invalidar_cache_completo(empresa_sel_h)
                            st.success(f"✅ Mudança registada: {horas_atuais}h → {novas_horas}h (vigência a partir de {data_inicio_vigor.strftime('%d/%m/%Y')})")
                            st.balloons()
                            time.sleep(2)
                            st.rerun()
        
        # Mostrar histórico
        st.markdown("---")
        st.subheader("📜 Histórico de Mudanças")
        df_historico = carregar_aba_excel(empresa_sel_h, "Historico_Horas_Semana", force_reload=True)
        if not df_historico.empty:
            st.dataframe(df_historico, use_container_width=True)
        else:
            st.info("📭 Nenhuma mudança registada ainda")

# ==================== PROCESSAR SALÁRIOS ====================

elif menu == "💼 Processar Salários":
    st.header("💼 Processamento Mensal de Salários")
    
    # Seleção inicial
    col1, col2, col3 = st.columns(3)
    
    with col1:
        empresa_proc = st.selectbox(
            "🏢 Empresa",
            options=list(EMPRESAS.keys()),
            key="empresa_processamento"
        )
    
    with col2:
        mes_proc = st.selectbox(
            "📅 Mês",
            options=list(range(1, 13)),
            format_func=lambda x: calendar.month_name[x],
            index=datetime.now().month - 1,
            key="mes_processamento"
        )
    
    with col3:
        ano_proc = st.selectbox(
            "📆 Ano",
            options=[2024, 2025, 2026],
            index=1,
            key="ano_processamento"
        )
    
    data_referencia = date(ano_proc, mes_proc, 1)
    df_ativos = carregar_colaboradores_ativos(empresa_proc, data_referencia)
    
    if df_ativos.empty:
        st.warning("⚠️ Nenhum colaborador ativo encontrado para esta data")
        st.stop()
    
    colaborador_proc = st.selectbox(
        "👤 Colaborador",
        options=df_ativos['Nome Completo'].tolist(),
        key="colaborador_processamento"
    )
    
    # Carregar dados do colaborador
    dados_colab = df_ativos[df_ativos['Nome Completo'] == colaborador_proc].iloc[0]
    horas_semana = carregar_horas_vigentes(empresa_proc, colaborador_proc, data_referencia)
    subsidio_alim = carregar_subsidio_vigente(empresa_proc, colaborador_proc)
    numero_pingo = carregar_numero_pingo_doce(empresa_proc, colaborador_proc)
    salario_bruto = calcular_salario_base(horas_semana, st.session_state.salario_minimo)
    vencimento_hora = calcular_vencimento_hora(salario_bruto, horas_semana)
    
    # Feriados do ano
    feriados_completos = FERIADOS_NACIONAIS_2025 + st.session_state.feriados_municipais
    dias_uteis_mes = calcular_dias_uteis(ano_proc, mes_proc, feriados_completos)
    
    st.markdown("---")
    
    # ===== DADOS BASE =====
    with st.expander("📋 **DADOS BASE DO COLABORADOR**", expanded=True):
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("💶 Salário Bruto", f"{salario_bruto:.2f}€")
        col2.metric("⏰ Horas/Semana", f"{horas_semana:.0f}h")
        col3.metric("💵 Vencimento/Hora", f"{vencimento_hora:.2f}€")
        col4.metric("🍽️ Sub. Alimentação", f"{subsidio_alim:.2f}€/dia")
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("📅 Dias Úteis Mês", dias_uteis_mes)
        col2.metric("🔢 NIF", dados_colab.get('NIF', 'N/A'))
        col3.metric("🔢 NISS", dados_colab.get('NISS', 'N/A'))
        if numero_pingo:
            col4.metric("🔢 Nº Pingo Doce", numero_pingo)
    
    st.markdown("---")
    
    # ===== OPÇÕES DE PROCESSAMENTO =====
    st.subheader("⚙️ Opções de Processamento")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        desconto_especie = st.checkbox("☑️ Desconto em Espécie", value=False)
    
    with col2:
        sub_ferias_tipo = st.selectbox("🏖️ Subsídio Férias", ["Duodécimos", "Total"])
    
    with col3:
        sub_natal_tipo = st.selectbox("🎄 Subsídio Natal", ["Duodécimos", "Total"])
    
    st.markdown("---")
    
    # ===== AUSÊNCIAS =====
    st.subheader("🏖️ Férias, Faltas e Baixas")
    
    tab_ferias, tab_faltas, tab_baixas = st.tabs(["🏖️ Férias", "🤒 Faltas", "🏥 Baixas"])
    
    # FÉRIAS
    with tab_ferias:
        st.caption("📅 Registre até 3 períodos de férias")
        periodos_ferias = []
        
        for i in range(3):
            col1, col2 = st.columns(2)
            with col1:
                inicio = st.date_input(
                    f"Início Período {i+1}",
                    value=None,
                    key=f"ferias_inicio_{i}"
                )
            with col2:
                fim = st.date_input(
                    f"Fim Período {i+1}",
                    value=None,
                    key=f"ferias_fim_{i}"
                )
            
            if inicio and fim:
                periodos_ferias.append((inicio, fim))
        
        # Calcular total
        total_dias_ferias = sum([calcular_dias_periodo(i, f) for i, f in periodos_ferias])
        total_dias_uteis_ferias = sum([calcular_dias_periodo(i, f, apenas_uteis=True, feriados_list=feriados_completos) for i, f in periodos_ferias])
        
        col1, col2 = st.columns(2)
        col1.metric("📊 Total Dias Férias", total_dias_ferias)
        col2.metric("📊 Total Dias Úteis Férias", total_dias_uteis_ferias)
    
    # FALTAS
    with tab_faltas:
        st.caption("📅 Registre até 3 períodos de faltas")
        periodos_faltas = []
        
        for i in range(3):
            col1, col2 = st.columns(2)
            with col1:
                inicio = st.date_input(
                    f"Início Período {i+1}",
                    value=None,
                    key=f"faltas_inicio_{i}"
                )
            with col2:
                fim = st.date_input(
                    f"Fim Período {i+1}",
                    value=None,
                    key=f"faltas_fim_{i}"
                )
            
            if inicio and fim:
                periodos_faltas.append((inicio, fim))
        
        total_dias_faltas = sum([calcular_dias_periodo(i, f) for i, f in periodos_faltas])
        total_dias_uteis_faltas = sum([calcular_dias_periodo(i, f, apenas_uteis=True, feriados_list=feriados_completos) for i, f in periodos_faltas])
        
        col1, col2 = st.columns(2)
        col1.metric("📊 Total Dias Faltas", total_dias_faltas)
        col2.metric("📊 Total Dias Úteis Faltas", total_dias_uteis_faltas)
    
    # BAIXAS
    with tab_baixas:
        st.caption("📅 Registre até 3 períodos de baixas + documentos")
        periodos_baixas = []
        docs_baixas = []
        
        # Verificar baixas transitórias
        baixas_transitorias = carregar_baixas_transitorias(empresa_proc, colaborador_proc, mes_proc, ano_proc)
        if baixas_transitorias:
            st.info(f"ℹ️ {len(baixas_transitorias)} baixa(s) transitória(s) do mês anterior detectada(s)")
            for idx, (inicio, fim) in enumerate(baixas_transitorias):
                st.text(f"Período transitório {idx+1}: {inicio.strftime('%d/%m/%Y')} → {fim.strftime('%d/%m/%Y')}")
                periodos_baixas.append((inicio, fim))
        
        for i in range(3):
            with st.container():
                st.markdown(f"**Período {i+1}**")
                col1, col2, col3 = st.columns([2, 2, 1])
                
                with col1:
                    inicio = st.date_input(
                        f"Início",
                        value=None,
                        key=f"baixas_inicio_{i}",
                        label_visibility="collapsed"
                    )
                
                with col2:
                    fim = st.date_input(
                        f"Fim",
                        value=None,
                        key=f"baixas_fim_{i}",
                        label_visibility="collapsed"
                    )
                
                with col3:
                    docs = st.file_uploader(
                        "Docs",
                        accept_multiple_files=True,
                        key=f"docs_baixas_{i}",
                        label_visibility="collapsed"
                    )
                
                if inicio and fim:
                    periodos_baixas.append((inicio, fim))
                    if docs:
                        docs_baixas.append((i, docs))
                
                st.markdown("---")
        
        total_dias_baixas = sum([calcular_dias_periodo(i, f) for i, f in periodos_baixas])
        total_dias_uteis_baixas = sum([calcular_dias_periodo(i, f, apenas_uteis=True, feriados_list=feriados_completos) for i, f in periodos_baixas])
        
        col1, col2 = st.columns(2)
        col1.metric("📊 Total Dias Baixas", total_dias_baixas)
        col2.metric("📊 Total Dias Úteis Baixas", total_dias_uteis_baixas)
    
    st.markdown("---")
    
    # ===== HORAS EXTRAS =====
    st.subheader("⏰ Horas Extras e Banco de Horas")
    
    modo_horas = st.radio(
        "Modo de Entrada",
        ["✍️ Manual", "📤 Importar Excel"],
        horizontal=True
    )
    
    if modo_horas == "✍️ Manual":
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            horas_noturnas = st.number_input("🌙 Horas Noturnas", min_value=0.0, value=0.0, step=0.5)
        
        with col2:
            horas_domingos = st.number_input("📅 Horas Domingos", min_value=0.0, value=0.0, step=0.5)
        
        with col3:
            horas_feriados = st.number_input("🎉 Horas Feriados", min_value=0.0, value=0.0, step=0.5)
        
        with col4:
            horas_extra = st.number_input("⚡ Horas Extra (pagas)", min_value=0.0, value=0.0, step=0.5)
    
    else:
        arquivo_horas = st.file_uploader("📤 Carregar ficheiro Excel com horas", type=['xlsx', 'xls'])
        if arquivo_horas:
            st.success("✅ Ficheiro carregado (funcionalidade de parsing em desenvolvimento)")
            horas_noturnas = 0.0
            horas_domingos = 0.0
            horas_feriados = 0.0
            horas_extra = 0.0
        else:
            horas_noturnas = 0.0
            horas_domingos = 0.0
            horas_feriados = 0.0
            horas_extra = 0.0
    
    # BANCO DE HORAS
    st.markdown("---")
    st.subheader("🏦 Banco de Horas")
    
    saldo_anterior = carregar_banco_horas_acumulado(empresa_proc, colaborador_proc, mes_proc, ano_proc)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("📊 Saldo Anterior", f"{saldo_anterior:.1f}h")
    
    with col2:
        banco_horas_mes = st.number_input(
            "Banco Horas Mês (+/-)",
            value=0.0,
            step=0.5,
            help="Positivo = acumular | Negativo = descontar"
        )
    
    with col3:
        saldo_novo = saldo_anterior + banco_horas_mes
        st.metric("📈 Novo Saldo", f"{saldo_novo:.1f}h")
    
    st.markdown("---")
    
    # ===== OUTROS PROVEITOS =====
    st.subheader("💰 Outros Proveitos com Descontos")
    outros_proveitos = st.number_input(
        "Valor Extra com Descontos (€)",
        min_value=0.0,
        value=0.0,
        step=10.0,
        help="Proveitos adicionais que sofrem descontos de SS e IRS"
    )
    
    st.markdown("---")
    
    # ===== CALCULAR DIAS TRABALHADOS =====
    data_admissao = pd.to_datetime(dados_colab.get('Data de Admissão', date.today())).date()
    dias_trabalhados = calcular_dias_trabalhados_com_admissao(
        mes_proc, ano_proc, data_admissao,
        total_dias_faltas, total_dias_baixas
    )
    dias_uteis_trabalhados = dias_uteis_mes - total_dias_uteis_faltas - total_dias_uteis_baixas - total_dias_uteis_ferias
    
    # ===== PREVIEW CÁLCULOS =====
    st.subheader("💵 Preview dos Cálculos")
    
    # Preparar dados para cálculo
    dados_calculo = {
        'salario_bruto': salario_bruto,
        'horas_semana': horas_semana,
        'subsidio_alimentacao': subsidio_alim,
        'dias_uteis_mes': dias_uteis_mes,
        'dias_trabalhados': dias_trabalhados,
        'dias_uteis_trabalhados': max(dias_uteis_trabalhados, 0),
        'horas_noturnas': horas_noturnas,
        'horas_domingos': horas_domingos,
        'horas_feriados': horas_feriados,
        'horas_extra': horas_extra,
        'sub_ferias_tipo': sub_ferias_tipo,
        'sub_natal_tipo': sub_natal_tipo,
        'desconto_especie': desconto_especie,
        'outros_proveitos': outros_proveitos
    }
    
    resultado = processar_calculo_salario(dados_calculo)
    
    # Mostrar resultados
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("### 💰 Remunerações")
        st.metric("Vencimento Ajustado", f"{resultado['vencimento_ajustado']:.2f}€")
        st.metric("Sub. Alimentação", f"{resultado['sub_alimentacao']:.2f}€")
        st.metric("Trabalho Noturno", f"{resultado['trabalho_noturno']:.2f}€")
        st.metric("Domingos", f"{resultado['domingos']:.2f}€")
        st.metric("Feriados", f"{resultado['feriados']:.2f}€")
        st.metric("Sub. Férias", f"{resultado['sub_ferias']:.2f}€")
        st.metric("Sub. Natal", f"{resultado['sub_natal']:.2f}€")
        st.metric("Horas Extra", f"{resultado['banco_horas_valor']:.2f}€")
        if outros_proveitos > 0:
            st.metric("Outros Proveitos", f"{resultado['outros_proveitos']:.2f}€")
        st.markdown("---")
        st.metric("**TOTAL REMUNERAÇÕES**", f"**{resultado['total_remuneracoes']:.2f}€**")
    
    with col2:
        st.markdown("### 📉 Descontos")
        st.metric("Base SS/IRS", f"{resultado['base_ss']:.2f}€")
        st.metric("Seg. Social (11%)", f"{resultado['seg_social']:.2f}€")
        st.metric("IRS", f"{resultado['irs']:.2f}€")
        if desconto_especie:
            st.metric("Desconto em Espécie", f"{resultado['desconto_especie']:.2f}€")
        st.markdown("---")
        st.metric("**TOTAL DESCONTOS**", f"**{resultado['total_descontos']:.2f}€**")
    
    with col3:
        st.markdown("### 💵 Resumo")
        st.metric("Dias Trabalhados", dias_trabalhados)
        st.metric("Dias Úteis Trabalhados", max(dias_uteis_trabalhados, 0))
        st.markdown("---")
        st.markdown("---")
        st.markdown("---")
        st.metric(
            "**💰 LÍQUIDO A RECEBER**",
            f"**{resultado['liquido']:.2f}€**",
            delta=None
        )
    
    st.markdown("---")
    
    # ===== AÇÕES FINAIS =====
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("💾 Guardar Processamento", use_container_width=True):
            st.info("🚧 Função de guardar em desenvolvimento...")
    
    with col2:
        if st.button("📄 Gerar PDF Recibo", use_container_width=True):
            st.info("🚧 Geração de PDF em desenvolvimento...")
    
    with col3:
        if st.button("📊 Export Excel", use_container_width=True):
            st.info("🚧 Export Excel em desenvolvimento...")

# ==================== RESCISÕES ====================

elif menu == "🚪 Rescisões":
    st.header("🚪 Gestão de Rescisões")
    
    empresa_resc = st.selectbox(
        "🏢 Empresa",
        options=list(EMPRESAS.keys()),
        key="empresa_rescisao"
    )
    
    garantir_todas_abas(empresa_resc)
    
    tab1, tab2 = st.tabs(["✍️ Nova Rescisão", "📜 Histórico"])
    
    # TAB 1: NOVA RESCISÃO
    with tab1:
        st.subheader("✍️ Registar Nova Rescisão")
        
        df_ativos_resc = carregar_colaboradores_ativos(empresa_resc, date.today())
        
        if df_ativos_resc.empty:
            st.warning("⚠️ Nenhum colaborador ativo encontrado")
        else:
            with st.form("form_rescisao"):
                colaborador_resc = st.selectbox(
                    "👤 Colaborador",
                    options=df_ativos_resc['Nome Completo'].tolist()
                )
                
                col1, col2 = st.columns(2)
                
                with col1:
                    data_rescisao = st.date_input(
                        "📅 Data da Rescisão",
                        value=date.today()
                    )
                
                with col2:
                    dias_aviso = st.number_input(
                        "Dias Aviso Prévio Cumpridos",
                        min_value=0,
                        value=0
                    )
                
                motivo_resc = st.selectbox(
                    "📋 Motivo da Rescisão",
                    options=MOTIVOS_RESCISAO
                )
                
                obs_resc = st.text_area(
                    "📝 Observações",
                    height=100,
                    placeholder="Informações adicionais sobre a rescisão..."
                )
                
                submit_resc = st.form_submit_button("💾 Registar Rescisão", use_container_width=True)
                
                if submit_resc:
                    with st.spinner("🔄 A registar rescisão..."):
                        sucesso = registar_rescisao(
                            empresa_resc,
                            colaborador_resc,
                            data_rescisao,
                            motivo_resc,
                            obs_resc,
                            dias_aviso
                        )
                        
                        if sucesso:
                            invalidar_cache_completo(empresa_resc)
                            st.success(f"✅ Rescisão de {colaborador_resc} registada com sucesso!")
                            st.info("ℹ️ Este colaborador não aparecerá mais nos processamentos após esta data")
                            time.sleep(2)
                            st.rerun()
    
    # TAB 2: HISTÓRICO
    with tab2:
        st.subheader("📜 Histórico de Rescisões")
        
        df_rescisoes = carregar_aba_excel(empresa_resc, "Rescisoes", force_reload=True)
        
        if not df_rescisoes.empty:
            # Formatar datas
            df_rescisoes['Data Rescisão'] = pd.to_datetime(df_rescisoes['Data Rescisão']).dt.strftime('%d/%m/%Y')
            
            st.dataframe(
                df_rescisoes,
                use_container_width=True,
                hide_index=True
            )
            
            st.metric("📊 Total de Rescisões", len(df_rescisoes))
        else:
            st.info("📭 Nenhuma rescisão registada ainda")

# ==================== RELATÓRIOS ====================

elif menu == "📊 Relatórios":
    st.header("📊 Relatórios e Análises")
    st.info("🚧 Módulo de relatórios em desenvolvimento...")
    
    st.markdown("""
    ### 📋 Funcionalidades Planeadas:
    
    - 📅 **Calendário Visual** com férias, faltas e baixas
    - 📜 **Histórico de Processamentos** completo
    - 📈 **Análises e Gráficos** (evolução salarial, custos, etc)
    - 📤 **Exportações** para Excel consolidado
    - 🏦 **Banco de Horas** - visualização de saldos
    """)

# ==================== SIDEBAR ====================

st.sidebar.markdown("---")
st.sidebar.info(f"👤 Sistema: v2.0\n💶 Salário Mínimo: {st.session_state.salario_minimo}€")

if st.sidebar.button("🚪 Logout", use_container_width=True):
    st.session_state.authenticated = False
    st.rerun()
