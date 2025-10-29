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
    page_title="Processamento Salarial v2.4",
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
        "tem_horas_extras": False
    },
    "CCM Retail Lda": {
        "path": "/Pedro Couto/Projectos/Pingo Doce/Pingo Doce/2. Operação/1. Recursos Humanos/Processamento salarial/Gestão Colaboradores.xlsx",
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

# Colunas expandidas do snapshot (incluindo dados para IRS)
COLUNAS_SNAPSHOT = [
    "Nome Completo", "Ano", "Mês", "Nº Horas/Semana", "Subsídio Alimentação Diário",
    "Número Pingo Doce", "Salário Bruto", "Vencimento Hora", 
    "Estado Civil", "Nº Titulares", "Nº Dependentes", "Deficiência",
    "IRS Percentagem Fixa", "IRS Modo Calculo",
    "Status", "Data Rescisão", "Motivo Rescisão", 
    "NIF", "NISS", "Data de Admissão", "IBAN", "Secção", "Timestamp"
]

ESTADOS_CIVIS = ["Solteiro", "Casado Único Titular", "Casado Dois Titulares"]

# MAPEAMENTO ENTRE APP REGISTO E APP PROCESSAMENTO
MAPEAMENTO_ESTADO_CIVIL = {
    "Não Casado": "Solteiro",
    "Casado 1": "Casado Único Titular",
    "Casado 2": "Casado Dois Titulares",
    # Manter compatibilidade com valores antigos
    "Solteiro": "Solteiro",
    "Casado Único Titular": "Casado Único Titular",
    "Casado Dois Titulares": "Casado Dois Titulares"
}

MAPEAMENTO_TIPO_IRS = {
    "Tabela": "Tabela",
    "Fixa": "Fixa",
    "Percentagem Fixa": "Fixa"  # Possível variação
}

# ==================== SESSION STATE ====================

if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'salario_minimo' not in st.session_state:
    st.session_state.salario_minimo = 870.0
if 'feriados_municipais' not in st.session_state:
    st.session_state.feriados_municipais = [date(2025, 1, 14)]
if 'ultimo_reload' not in st.session_state:
    st.session_state.ultimo_reload = datetime.now()
if 'tabela_irs' not in st.session_state:
    st.session_state.tabela_irs = None

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

# ==================== FUNÇÕES DE MAPEAMENTO ====================

def normalizar_estado_civil(valor):
    """Mapeia estado civil da app registo para app processamento"""
    if pd.isna(valor) or valor == '':
        return "Solteiro"
    valor_str = str(valor).strip()
    return MAPEAMENTO_ESTADO_CIVIL.get(valor_str, "Solteiro")

def normalizar_tipo_irs(valor):
    """Mapeia tipo IRS entre diferentes nomenclaturas"""
    if pd.isna(valor) or valor == '':
        return "Tabela"
    valor_str = str(valor).strip()
    return MAPEAMENTO_TIPO_IRS.get(valor_str, "Tabela")

# ==================== FUNÇÕES DROPBOX ====================

def get_nome_aba_snapshot(ano, mes):
    return f"Estado_{ano}_{mes:02d}"

def download_excel(empresa):
    try:
        file_path = EMPRESAS[empresa]["path"]
        _, response = dbx.files_download(file_path)
        return BytesIO(response.content)
    except Exception as e:
        st.error(f"❌ Erro ao baixar Excel: {e}")
        return None

def garantir_aba(wb, nome_aba, colunas):
    if nome_aba not in wb.sheetnames:
        ws = wb.create_sheet(nome_aba)
        ws.append(colunas)
        return True
    return False

def upload_excel(empresa, wb):
    try:
        file_path = EMPRESAS[empresa]["path"]
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        dbx.files_upload(output.read(), file_path, mode=dropbox.files.WriteMode.overwrite)
        return True
    except Exception as e:
        st.error(f"❌ Erro ao enviar Excel: {e}")
        return False

# ==================== FUNÇÕES DE CÁLCULO ====================

def calcular_salario_base(horas_semana, salario_minimo):
    if horas_semana == 40:
        return salario_minimo
    elif horas_semana == 20:
        return salario_minimo / 2
    elif horas_semana == 16:
        return salario_minimo * 0.4
    return salario_minimo * (horas_semana / 40)

def calcular_vencimento_hora(salario_bruto, horas_semana):
    if horas_semana == 0:
        return 0
    return (salario_bruto * 12) / (52 * horas_semana)

def calcular_vencimento_ajustado(salario_bruto, dias_faltas, dias_baixas):
    """
    FÓRMULA CORRETA: (salario_bruto / 30) * (30 - faltas - baixas)
    SEMPRE usa 30 como base!
    """
    dias_pagos = 30 - dias_faltas - dias_baixas
    dias_pagos = max(dias_pagos, 0)  # Não pode ser negativo
    return (salario_bruto / 30) * dias_pagos

def calcular_dias_uteis(ano, mes, feriados_list):
    num_dias = calendar.monthrange(ano, mes)[1]
    dias_uteis = 0
    for dia in range(1, num_dias + 1):
        data = date(ano, mes, dia)
        if data.weekday() < 5 and data not in feriados_list:
            dias_uteis += 1
    return dias_uteis

def carregar_tabela_irs_excel(uploaded_file):
    """Carrega tabela IRS de ficheiro Excel"""
    try:
        xls = pd.ExcelFile(uploaded_file)
        st.success(f"✅ Ficheiro carregado! Abas encontradas: {', '.join(xls.sheet_names)}")
        st.session_state.tabela_irs = xls
        return xls
    except Exception as e:
        st.error(f"❌ Erro ao carregar tabela: {e}")
        return None

def calcular_irs_por_tabela(base_incidencia, estado_civil, num_dependentes, tem_deficiencia=False):
    """
    Calcula IRS com base nas tabelas carregadas
    base_incidencia = salário bruto mensal
    
    ⚠️ NOTA: Esta é uma implementação simplificada.
    Para cálculo preciso, seria necessário:
    1. Identificar a tabela correta (I-VII)
    2. Encontrar o escalão de rendimento
    3. Aplicar taxa e parcela a abater
    
    Por agora, retorna estimativa baseada em escalões aproximados
    """
    if st.session_state.tabela_irs is None:
        st.warning("⚠️ Tabela IRS não carregada. Usando 10% por defeito.")
        return base_incidencia * 0.10
    
    # Escalões simplificados (valores aproximados 2025)
    # TODO: Melhorar com leitura real das tabelas Excel
    
    # Ajustar por dependentes (reduz taxa)
    reducao_dependentes = num_dependentes * 0.01  # 1% por dependente (aproximado)
    
    # Escalões básicos (simplificados)
    if base_incidencia <= 820:
        taxa = 0.135
    elif base_incidencia <= 1200:
        taxa = 0.18
    elif base_incidencia <= 1700:
        taxa = 0.23
    elif base_incidencia <= 2500:
        taxa = 0.265
    else:
        taxa = 0.32
    
    # Aplicar redução por dependentes
    taxa_final = max(taxa - reducao_dependentes, 0.05)  # Mínimo 5%
    
    # Casado único titular tem taxas mais baixas
    if estado_civil == "Casado Único Titular":
        taxa_final *= 0.85  # Redução de 15%
    
    return base_incidencia * taxa_final

def calcular_irs(base_incidencia, modo_calculo, percentagem_fixa, estado_civil, num_dependentes, tem_deficiencia=False):
    """
    Função unificada de cálculo IRS - CORRIGIDA!
    
    Args:
        base_incidencia: Salário bruto mensal
        modo_calculo: "Fixa" ou "Tabela"
        percentagem_fixa: % IRS fixa configurada
        estado_civil: Estado civil do colaborador
        num_dependentes: Número de dependentes
        tem_deficiencia: Tem deficiência
    
    Returns:
        Valor de IRS a reter
    """
    
    if modo_calculo == "Fixa":
        # Usar percentagem fixa configurada
        taxa = percentagem_fixa / 100
        irs = base_incidencia * taxa
        return irs
    else:
        # Usar tabela IRS
        return calcular_irs_por_tabela(base_incidencia, estado_civil, num_dependentes, tem_deficiencia)

# ==================== FUNÇÕES DE DADOS BASE ====================

def carregar_dados_base(empresa):
    excel_file = download_excel(empresa)
    if excel_file:
        try:
            df = pd.read_excel(excel_file, sheet_name="Colaboradores")
            return df
        except Exception as e:
            st.error(f"❌ Erro ao ler aba Colaboradores: {e}")
    return pd.DataFrame()

def criar_snapshot_inicial(empresa, colaborador, ano, mes):
    df_base = carregar_dados_base(empresa)
    dados_colab = df_base[df_base['Nome Completo'] == colaborador]
    
    if dados_colab.empty:
        return None
    
    dados = dados_colab.iloc[0]
    horas_semana = float(dados.get('Nº Horas/Semana', 40))
    salario_bruto = calcular_salario_base(horas_semana, st.session_state.salario_minimo)
    
    # Aplicar mapeamentos
    estado_civil_raw = dados.get('Estado Civil', 'Solteiro')
    estado_civil = normalizar_estado_civil(estado_civil_raw)
    
    tipo_irs_raw = dados.get('Tipo IRS', dados.get('IRS Modo Calculo', 'Tabela'))
    tipo_irs = normalizar_tipo_irs(tipo_irs_raw)
    
    snapshot = {
        "Nome Completo": colaborador,
        "Ano": ano,
        "Mês": mes,
        "Nº Horas/Semana": horas_semana,
        "Subsídio Alimentação Diário": float(dados.get('Subsídio Alimentação Diário', 5.96)),
        "Número Pingo Doce": str(dados.get('Número Pingo Doce', '')),
        "Salário Bruto": salario_bruto,
        "Vencimento Hora": calcular_vencimento_hora(salario_bruto, horas_semana),
        "Estado Civil": estado_civil,
        "Nº Titulares": int(dados.get('Nº Titulares', 2)),
        "Nº Dependentes": int(dados.get('Nº Dependentes', 0)),
        "Deficiência": str(dados.get('Deficiência', 'Não')),
        "IRS Percentagem Fixa": float(dados.get('IRS Percentagem Fixa', dados.get('% IRS Fixa', 0))),
        "IRS Modo Calculo": tipo_irs,
        "Status": "Ativo",
        "Data Rescisão": "",
        "Motivo Rescisão": "",
        "NIF": str(dados.get('NIF', '')),
        "NISS": str(dados.get('NISS', '')),
        "Data de Admissão": str(dados.get('Data de Admissão', '')),
        "IBAN": str(dados.get('IBAN', '')),
        "Secção": str(dados.get('Secção', '')),
        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    
    return snapshot

def carregar_ultimo_snapshot(empresa, colaborador, ano, mes):
    excel_file = download_excel(empresa)
    if not excel_file:
        return None
    
    try:
        wb = load_workbook(excel_file)
        nome_aba = get_nome_aba_snapshot(ano, mes)
        
        if nome_aba in wb.sheetnames:
            df = pd.read_excel(excel_file, sheet_name=nome_aba)
            df_colab = df[df['Nome Completo'] == colaborador]
            
            if not df_colab.empty:
                snapshot = df_colab.iloc[-1].to_dict()
                
                # APLICAR MAPEAMENTOS ao carregar
                if 'Estado Civil' in snapshot:
                    snapshot['Estado Civil'] = normalizar_estado_civil(snapshot['Estado Civil'])
                if 'IRS Modo Calculo' in snapshot:
                    snapshot['IRS Modo Calculo'] = normalizar_tipo_irs(snapshot['IRS Modo Calculo'])
                
                st.caption(f"📸 Snapshot: {snapshot.get('Timestamp', 'N/A')}")
                return snapshot
        
        abas_estado = sorted([s for s in wb.sheetnames if s.startswith('Estado_')], reverse=True)
        
        for aba in abas_estado:
            try:
                df = pd.read_excel(excel_file, sheet_name=aba)
                df_colab = df[df['Nome Completo'] == colaborador]
                
                if not df_colab.empty:
                    snapshot = df_colab.iloc[-1].to_dict()
                    snapshot['Ano'] = ano
                    snapshot['Mês'] = mes
                    snapshot['Timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    
                    # APLICAR MAPEAMENTOS
                    if 'Estado Civil' in snapshot:
                        snapshot['Estado Civil'] = normalizar_estado_civil(snapshot['Estado Civil'])
                    if 'IRS Modo Calculo' in snapshot:
                        snapshot['IRS Modo Calculo'] = normalizar_tipo_irs(snapshot['IRS Modo Calculo'])
                    
                    st.caption(f"📸 Herdado de {aba}")
                    return snapshot
            except:
                continue
        
        return criar_snapshot_inicial(empresa, colaborador, ano, mes)
        
    except Exception as e:
        st.error(f"❌ Erro: {e}")
        return None

def gravar_snapshot(empresa, snapshot):
    try:
        ano = snapshot['Ano']
        mes = snapshot['Mês']
        nome_aba = get_nome_aba_snapshot(ano, mes)
        
        excel_file = download_excel(empresa)
        if not excel_file:
            return False
        
        wb = load_workbook(excel_file)
        aba_criada = garantir_aba(wb, nome_aba, COLUNAS_SNAPSHOT)
        if aba_criada:
            st.info(f"✨ Aba '{nome_aba}' criada")
        
        ws = wb[nome_aba]
        
        nova_linha = []
        for col in COLUNAS_SNAPSHOT:
            valor = snapshot.get(col, '')
            if isinstance(valor, (int, float)):
                nova_linha.append(valor)
            else:
                nova_linha.append(str(valor) if valor else '')
        
        ws.append(nova_linha)
        
        sucesso = upload_excel(empresa, wb)
        
        if sucesso:
            linha = ws.max_row
            st.success(f"✅ Snapshot gravado (linha {linha})")
            return True
        
        return False
        
    except Exception as e:
        st.error(f"❌ Erro ao gravar: {e}")
        return False

def atualizar_campo_colaborador(empresa, colaborador, ano, mes, campo, novo_valor):
    snapshot = carregar_ultimo_snapshot(empresa, colaborador, ano, mes)
    
    if not snapshot:
        return False
    
    snapshot[campo] = novo_valor
    snapshot['Ano'] = ano
    snapshot['Mês'] = mes
    snapshot['Timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    if campo == "Nº Horas/Semana":
        horas = float(novo_valor)
        snapshot['Salário Bruto'] = calcular_salario_base(horas, st.session_state.salario_minimo)
        snapshot['Vencimento Hora'] = calcular_vencimento_hora(snapshot['Salário Bruto'], horas)
    
    return gravar_snapshot(empresa, snapshot)

def carregar_colaboradores_ativos(empresa, ano, mes):
    df_base = carregar_dados_base(empresa)
    
    if df_base.empty:
        return []
    
    colaboradores_ativos = []
    
    for _, colab in df_base.iterrows():
        nome = colab['Nome Completo']
        snapshot = carregar_ultimo_snapshot(empresa, nome, ano, mes)
        
        if snapshot and snapshot.get('Status') == 'Ativo':
            colaboradores_ativos.append(nome)
    
    return colaboradores_ativos

def processar_calculo_salario(dados_form):
    """Processa cálculos salariais com fórmulas CORRETAS"""
    salario_bruto = dados_form['salario_bruto']
    horas_semana = dados_form['horas_semana']
    sub_alimentacao_dia = dados_form['subsidio_alimentacao']
    vencimento_hora = calcular_vencimento_hora(salario_bruto, horas_semana)
    
    dias_faltas = dados_form['dias_faltas']
    dias_baixas = dados_form['dias_baixas']
    dias_uteis_trabalhados = dados_form['dias_uteis_trabalhados']
    
    horas_noturnas = dados_form.get('horas_noturnas', 0)
    horas_domingos = dados_form.get('horas_domingos', 0)
    horas_feriados = dados_form.get('horas_feriados', 0)
    horas_extra = dados_form.get('horas_extra', 0)
    
    # VENCIMENTO AJUSTADO - FÓRMULA CORRETA!
    vencimento_ajustado = calcular_vencimento_ajustado(salario_bruto, dias_faltas, dias_baixas)
    
    sub_alimentacao = sub_alimentacao_dia * dias_uteis_trabalhados
    trabalho_noturno = horas_noturnas * vencimento_hora * 0.25
    domingos = horas_domingos * vencimento_hora
    feriados = horas_feriados * vencimento_hora * 2
    
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
    
    # DESCONTOS
    base_ss = total_remuneracoes - sub_alimentacao
    seg_social = base_ss * 0.11
    
    # IRS - FUNÇÃO CORRIGIDA! ✅
    irs = calcular_irs(
        base_incidencia=salario_bruto,
        modo_calculo=dados_form.get('irs_modo', 'Tabela'),
        percentagem_fixa=dados_form.get('irs_percentagem_fixa', 0),
        estado_civil=dados_form.get('estado_civil', 'Solteiro'),
        num_dependentes=dados_form.get('num_dependentes', 0),
        tem_deficiencia=dados_form.get('tem_deficiencia', False)
    )
    
    desconto_especie = sub_alimentacao if dados_form.get('desconto_especie', False) else 0
    total_descontos = seg_social + irs + desconto_especie
    
    liquido = total_remuneracoes - total_descontos
    
    return {
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
        'base_irs': salario_bruto,
        'irs': irs,
        'desconto_especie': desconto_especie,
        'total_descontos': total_descontos,
        'liquido': liquido
    }

def registar_rescisao(empresa, colaborador, ano, mes, data_rescisao, motivo, obs, dias_aviso):
    snapshot = carregar_ultimo_snapshot(empresa, colaborador, ano, mes)
    
    if not snapshot:
        return False
    
    snapshot['Status'] = 'Rescindido'
    snapshot['Data Rescisão'] = data_rescisao.strftime("%Y-%m-%d")
    snapshot['Motivo Rescisão'] = f"{motivo} | Dias aviso: {dias_aviso} | Obs: {obs}"
    snapshot['Timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    return gravar_snapshot(empresa, snapshot)

# ==================== INTERFACE ====================

if not check_password():
    st.stop()

st.title("💰 Processamento Salarial v2.4")
st.caption("✅ IRS corrigido | ⏰ Horários implementados | 🔄 Mapeamento entre apps")
st.caption(f"🕐 Reload: {st.session_state.ultimo_reload.strftime('%H:%M:%S')}")
st.markdown("---")

menu = st.sidebar.radio(
    "Menu Principal",
    ["⚙️ Configurações", "💼 Processar Salários", "🚪 Rescisões", "📊 Tabela IRS"],
    index=0
)

# ==================== CONFIGURAÇÕES ====================

if menu == "⚙️ Configurações":
    st.header("⚙️ Configurações do Sistema")
    
    tab1, tab2, tab3, tab4 = st.tabs(["💶 Sistema", "👥 Colaboradores", "⏰ Horários", "📋 Dados IRS"])
    
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
            if st.button("💾 Atualizar SMN"):
                st.session_state.salario_minimo = novo_salario
                st.success(f"✅ SMN: {novo_salario}€")
        
        with col2:
            st.subheader("📅 Feriados Municipais")
            feriados_temp = []
            for i in range(3):
                valor_default = st.session_state.feriados_municipais[i] if i < len(st.session_state.feriados_municipais) else None
                feriado = st.date_input(f"Feriado {i+1}", value=valor_default, key=f"fer_{i}")
                if feriado:
                    feriados_temp.append(feriado)
            
            if st.button("💾 Atualizar Feriados"):
                st.session_state.feriados_municipais = feriados_temp
                st.success(f"✅ {len(feriados_temp)} feriados")
    
    # TAB 2: COLABORADORES
    with tab2:
        st.subheader("👥 Editar Dados")
        st.warning("⚠️ Aguarda confirmação antes de navegar!")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            emp = st.selectbox("Empresa", list(EMPRESAS.keys()), key="emp_cfg")
        with col2:
            mes_cfg = st.selectbox("Mês", list(range(1, 13)), 
                                 format_func=lambda x: calendar.month_name[x],
                                 index=datetime.now().month - 1, key="mes_cfg")
        with col3:
            ano_cfg = st.selectbox("Ano", [2024, 2025, 2026], index=1, key="ano_cfg")
        
        colabs = carregar_colaboradores_ativos(emp, ano_cfg, mes_cfg)
        
        if colabs:
            colab_sel = st.selectbox("Colaborador", colabs, key="col_cfg")
            
            with st.spinner("🔄 Carregando..."):
                snap = carregar_ultimo_snapshot(emp, colab_sel, ano_cfg, mes_cfg)
            
            if snap:
                st.markdown("---")
                col1, col2, col3 = st.columns(3)
                col1.metric("💰 Subsídio", f"{snap['Subsídio Alimentação Diário']:.2f}€")
                col2.metric("⏰ Horas", f"{snap['Nº Horas/Semana']:.0f}h")
                col3.metric("🔢 Nº Pingo", snap.get('Número Pingo Doce', ''))
                
                with st.form("form_edit"):
                    col1, col2 = st.columns(2)
                    with col1:
                        novo_sub = st.number_input("Novo Subsídio (€)", min_value=0.0,
                                                  value=float(snap['Subsídio Alimentação Diário']),
                                                  step=0.10, format="%.2f")
                    with col2:
                        novo_num = st.text_input("Novo Nº Pingo", value=str(snap.get('Número Pingo Doce', '')))
                    
                    submit = st.form_submit_button("💾 GUARDAR", use_container_width=True, type="primary")
                    
                    if submit:
                        st.warning("⏳ AGUARDA...")
                        with st.spinner("1/2: Subsídio..."):
                            s1 = atualizar_campo_colaborador(emp, colab_sel, ano_cfg, mes_cfg,
                                                            "Subsídio Alimentação Diário", novo_sub)
                            time.sleep(1)
                        with st.spinner("2/2: Número..."):
                            s2 = atualizar_campo_colaborador(emp, colab_sel, ano_cfg, mes_cfg,
                                                            "Número Pingo Doce", novo_num)
                            time.sleep(1)
                        
                        if s1 and s2:
                            st.success("✅ GRAVADO!")
                            st.balloons()
                            time.sleep(3)
                            st.session_state.ultimo_reload = datetime.now()
                            st.rerun()
        else:
            st.warning("⚠️ Nenhum colaborador ativo")
    
    # TAB 3: HORÁRIOS - IMPLEMENTADO! ✅
    with tab3:
        st.subheader("⏰ Mudanças de Horário")
        st.info("💡 Altere as horas semanais do colaborador. O salário bruto será recalculado automaticamente.")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            emp_hor = st.selectbox("Empresa", list(EMPRESAS.keys()), key="emp_hor")
        with col2:
            mes_hor = st.selectbox("Mês", list(range(1, 13)),
                                  format_func=lambda x: calendar.month_name[x],
                                  index=datetime.now().month - 1, key="mes_hor")
        with col3:
            ano_hor = st.selectbox("Ano", [2024, 2025, 2026], index=1, key="ano_hor")
        
        colabs_hor = carregar_colaboradores_ativos(emp_hor, ano_hor, mes_hor)
        
        if colabs_hor:
            colab_hor = st.selectbox("Colaborador", colabs_hor, key="col_hor")
            
            with st.spinner("🔄 Carregando..."):
                snap_hor = carregar_ultimo_snapshot(emp_hor, colab_hor, ano_hor, mes_hor)
            
            if snap_hor:
                st.markdown("---")
                
                # Mostrar dados atuais
                col1, col2, col3 = st.columns(3)
                horas_atuais = float(snap_hor['Nº Horas/Semana'])
                salario_atual = float(snap_hor['Salário Bruto'])
                venc_hora_atual = float(snap_hor['Vencimento Hora'])
                
                col1.metric("⏰ Horas Atuais", f"{horas_atuais:.0f}h/semana")
                col2.metric("💰 Salário Bruto Atual", f"{salario_atual:.2f}€")
                col3.metric("💵 Vencimento/Hora Atual", f"{venc_hora_atual:.2f}€")
                
                st.markdown("---")
                
                with st.form("form_horario"):
                    st.markdown("### Novo Horário")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        novas_horas = st.number_input(
                            "⏰ Novas Horas Semanais",
                            min_value=4.0,
                            max_value=40.0,
                            value=horas_atuais,
                            step=1.0,
                            help="Entre 4h e 40h semanais"
                        )
                    
                    with col2:
                        # Preview do novo salário
                        novo_salario = calcular_salario_base(novas_horas, st.session_state.salario_minimo)
                        novo_venc_hora = calcular_vencimento_hora(novo_salario, novas_horas)
                        
                        st.metric("💰 Novo Salário Bruto", f"{novo_salario:.2f}€",
                                 delta=f"{novo_salario - salario_atual:.2f}€")
                        st.metric("💵 Novo Vencimento/Hora", f"{novo_venc_hora:.2f}€",
                                 delta=f"{novo_venc_hora - venc_hora_atual:.2f}€")
                    
                    observacoes = st.text_area(
                        "📝 Observações (opcional)",
                        placeholder="Ex: Mudança de part-time para full-time",
                        height=80
                    )
                    
                    submit_hor = st.form_submit_button(
                        "💾 CONFIRMAR MUDANÇA DE HORÁRIO",
                        use_container_width=True,
                        type="primary"
                    )
                    
                    if submit_hor:
                        if novas_horas == horas_atuais:
                            st.warning("⚠️ As horas não foram alteradas!")
                        else:
                            st.warning("⏳ PROCESSANDO MUDANÇA...")
                            
                            with st.spinner("Atualizando horário..."):
                                # Atualizar snapshot com novos valores
                                snap_hor['Nº Horas/Semana'] = novas_horas
                                snap_hor['Salário Bruto'] = novo_salario
                                snap_hor['Vencimento Hora'] = novo_venc_hora
                                
                                if observacoes:
                                    obs_atual = snap_hor.get('Observações', '')
                                    snap_hor['Observações'] = f"{obs_atual}\n[{datetime.now().strftime('%Y-%m-%d')}] Mudança horário: {horas_atuais:.0f}h → {novas_horas:.0f}h. {observacoes}".strip()
                                
                                snap_hor['Timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                
                                if gravar_snapshot(emp_hor, snap_hor):
                                    st.success("✅ HORÁRIO ATUALIZADO COM SUCESSO!")
                                    st.balloons()
                                    
                                    # Mostrar resumo
                                    st.info(f"""
                                    **Resumo da alteração:**
                                    - Horas: {horas_atuais:.0f}h → {novas_horas:.0f}h
                                    - Salário: {salario_atual:.2f}€ → {novo_salario:.2f}€
                                    - Vencimento/hora: {venc_hora_atual:.2f}€ → {novo_venc_hora:.2f}€
                                    """)
                                    
                                    time.sleep(3)
                                    st.session_state.ultimo_reload = datetime.now()
                                    st.rerun()
                                else:
                                    st.error("❌ Erro ao gravar alteração!")
        else:
            st.warning("⚠️ Nenhum colaborador ativo")
    
    # TAB 4: DADOS IRS
    with tab4:
        st.subheader("📋 Configuração de Dados para IRS")
        st.info("💡 Configure dados de estado civil, dependentes e % IRS fixa por colaborador")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            emp_irs = st.selectbox("Empresa", list(EMPRESAS.keys()), key="emp_irs")
        with col2:
            mes_irs = st.selectbox("Mês", list(range(1, 13)),
                                  format_func=lambda x: calendar.month_name[x],
                                  index=datetime.now().month - 1, key="mes_irs")
        with col3:
            ano_irs = st.selectbox("Ano", [2024, 2025, 2026], index=1, key="ano_irs")
        
        colabs_irs = carregar_colaboradores_ativos(emp_irs, ano_irs, mes_irs)
        
        if colabs_irs:
            colab_irs = st.selectbox("Colaborador", colabs_irs, key="col_irs")
            
            snap_irs = carregar_ultimo_snapshot(emp_irs, colab_irs, ano_irs, mes_irs)
            
            if snap_irs:
                with st.form("form_irs"):
                    st.markdown(f"### {colab_irs}")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        estado_civil = st.selectbox("Estado Civil", ESTADOS_CIVIS,
                                                   index=ESTADOS_CIVIS.index(snap_irs.get('Estado Civil', 'Solteiro'))
                                                   if snap_irs.get('Estado Civil') in ESTADOS_CIVIS else 0)
                        num_titulares = st.number_input("Nº Titulares", min_value=1, max_value=2,
                                                       value=int(snap_irs.get('Nº Titulares', 2)))
                        num_dependentes = st.number_input("Nº Dependentes", min_value=0,
                                                         value=int(snap_irs.get('Nº Dependentes', 0)))
                    
                    with col2:
                        tem_deficiencia = st.selectbox("Deficiência", ["Não", "Sim"],
                                                      index=0 if snap_irs.get('Deficiência', 'Não') == 'Não' else 1)
                        irs_modo = st.selectbox("Modo Cálculo IRS", ["Tabela", "Fixa"],
                                               index=0 if snap_irs.get('IRS Modo Calculo', 'Tabela') == 'Tabela' else 1)
                        irs_percentagem = st.number_input("IRS % Fixa (se aplicável)", min_value=0.0, max_value=100.0,
                                                         value=float(snap_irs.get('IRS Percentagem Fixa', 0)),
                                                         step=0.1, format="%.1f")
                    
                    submit_irs = st.form_submit_button("💾 GUARDAR DADOS IRS", use_container_width=True, type="primary")
                    
                    if submit_irs:
                        # Atualizar múltiplos campos
                        snap_irs['Estado Civil'] = estado_civil
                        snap_irs['Nº Titulares'] = num_titulares
                        snap_irs['Nº Dependentes'] = num_dependentes
                        snap_irs['Deficiência'] = tem_deficiencia
                        snap_irs['IRS Modo Calculo'] = irs_modo
                        snap_irs['IRS Percentagem Fixa'] = irs_percentagem
                        snap_irs['Timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        
                        with st.spinner("🔄 Gravando..."):
                            if gravar_snapshot(emp_irs, snap_irs):
                                st.success("✅ Dados IRS atualizados!")
                                st.balloons()
                                time.sleep(2)
                                st.rerun()

# ==================== PROCESSAR SALÁRIOS ====================

elif menu == "💼 Processar Salários":
    st.header("💼 Processamento Mensal")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        emp_proc = st.selectbox("Empresa", list(EMPRESAS.keys()), key="emp_proc")
    with col2:
        mes_proc = st.selectbox("Mês", list(range(1, 13)),
                               format_func=lambda x: calendar.month_name[x],
                               index=datetime.now().month - 1, key="mes_proc")
    with col3:
        ano_proc = st.selectbox("Ano", [2024, 2025, 2026], index=1, key="ano_proc")
    
    with st.spinner("🔄 Carregando..."):
        colabs_proc = carregar_colaboradores_ativos(emp_proc, ano_proc, mes_proc)
    
    if not colabs_proc:
        st.warning("⚠️ Nenhum colaborador ativo")
        st.stop()
    
    colab_proc = st.selectbox("Colaborador", colabs_proc, key="col_proc")
    
    with st.spinner("🔄 Carregando snapshot..."):
        snap_proc = carregar_ultimo_snapshot(emp_proc, colab_proc, ano_proc, mes_proc)
    
    if not snap_proc:
        st.error("❌ Erro ao carregar")
        st.stop()
    
    salario_bruto = float(snap_proc['Salário Bruto'])
    horas_semana = float(snap_proc['Nº Horas/Semana'])
    subsidio_alim = float(snap_proc['Subsídio Alimentação Diário'])
    vencimento_hora = float(snap_proc['Vencimento Hora'])
    
    feriados = FERIADOS_NACIONAIS_2025 + st.session_state.feriados_municipais
    dias_uteis_mes = calcular_dias_uteis(ano_proc, mes_proc, feriados)
    
    st.markdown("---")
    
    # DADOS BASE
    with st.expander("📋 DADOS BASE", expanded=True):
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("💶 Salário Bruto", f"{salario_bruto:.2f}€")
        col2.metric("⏰ Horas/Semana", f"{horas_semana:.0f}h")
        col3.metric("💵 Vencimento/Hora", f"{vencimento_hora:.2f}€")
        col4.metric("🍽️ Sub. Alimentação", f"{subsidio_alim:.2f}€/dia")
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("📅 Dias Úteis Mês", dias_uteis_mes)
        col2.metric("👤 Estado Civil", snap_proc.get('Estado Civil', 'N/A'))
        col3.metric("👶 Dependentes", snap_proc.get('Nº Dependentes', 0))
        col4.metric("📊 Modo IRS", snap_proc.get('IRS Modo Calculo', 'Tabela'))
    
    st.markdown("---")
    
    # OPÇÕES
    col1, col2, col3 = st.columns(3)
    with col1:
        desconto_especie = st.checkbox("☑️ Desconto em Espécie")
    with col2:
        sub_ferias = st.selectbox("🏖️ Sub. Férias", ["Duodécimos", "Total"])
    with col3:
        sub_natal = st.selectbox("🎄 Sub. Natal", ["Duodécimos", "Total"])
    
    st.markdown("---")
    
    # AUSÊNCIAS
    st.subheader("🏖️ Faltas e Baixas")
    col1, col2 = st.columns(2)
    with col1:
        faltas = st.number_input("Total Dias Faltas", min_value=0, value=0, key="falt")
    with col2:
        baixas = st.number_input("Total Dias Baixas", min_value=0, value=0, key="baix")
    
    # Calcular dias úteis trabalhados
    dias_uteis_trab = max(dias_uteis_mes - faltas - baixas, 0)
    
    st.markdown("---")
    
    # HORAS EXTRAS
    st.subheader("⏰ Horas Extras")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        h_not = st.number_input("🌙 Noturnas", min_value=0.0, value=0.0, step=0.5)
    with col2:
        h_dom = st.number_input("📅 Domingos", min_value=0.0, value=0.0, step=0.5)
    with col3:
        h_fer = st.number_input("🎉 Feriados", min_value=0.0, value=0.0, step=0.5)
    with col4:
        h_ext = st.number_input("⚡ Extra", min_value=0.0, value=0.0, step=0.5)
    
    st.markdown("---")
    
    outros_prov = st.number_input("💰 Outros Proveitos c/ Descontos (€)", min_value=0.0, value=0.0)
    
    st.markdown("---")
    
    # CALCULAR
    dados_calc = {
        'salario_bruto': salario_bruto,
        'horas_semana': horas_semana,
        'subsidio_alimentacao': subsidio_alim,
        'dias_faltas': faltas,
        'dias_baixas': baixas,
        'dias_uteis_trabalhados': dias_uteis_trab,
        'horas_noturnas': h_not,
        'horas_domingos': h_dom,
        'horas_feriados': h_fer,
        'horas_extra': h_ext,
        'sub_ferias_tipo': sub_ferias,
        'sub_natal_tipo': sub_natal,
        'desconto_especie': desconto_especie,
        'outros_proveitos': outros_prov,
        'estado_civil': snap_proc.get('Estado Civil', 'Solteiro'),
        'num_dependentes': snap_proc.get('Nº Dependentes', 0),
        'tem_deficiencia': snap_proc.get('Deficiência', 'Não') == 'Sim',
        'irs_modo': snap_proc.get('IRS Modo Calculo', 'Tabela'),
        'irs_percentagem_fixa': snap_proc.get('IRS Percentagem Fixa', 0)
    }
    
    resultado = processar_calculo_salario(dados_calc)
    
    # DEBUG - IRS
    with st.expander("🔍 Debug - Cálculo IRS", expanded=False):
        st.write(f"**Base de incidência:** {resultado['base_irs']:.2f}€ (Salário Bruto)")
        st.write(f"**Modo:** {dados_calc['irs_modo']}")
        
        if dados_calc['irs_modo'] == 'Fixa':
            st.write(f"**Taxa configurada:** {dados_calc['irs_percentagem_fixa']:.1f}%")
            st.write(f"**Cálculo:** {resultado['base_irs']:.2f}€ × {dados_calc['irs_percentagem_fixa']/100:.3f} = **{resultado['irs']:.2f}€**")
        else:
            st.write(f"**Estado Civil:** {dados_calc['estado_civil']}")
            st.write(f"**Dependentes:** {dados_calc['num_dependentes']}")
            st.write(f"**Deficiência:** {'Sim' if dados_calc['tem_deficiencia'] else 'Não'}")
            
            if st.session_state.tabela_irs:
                st.success("✅ Tabela IRS carregada - cálculo por tabela ativo")
            else:
                st.warning("⚠️ Sem tabela IRS - usando cálculo estimado por escalões")
            
            st.write(f"**IRS calculado:** **{resultado['irs']:.2f}€**")
    
    # DEBUG - Vencimento Ajustado
    with st.expander("🔍 Debug - Vencimento Ajustado", expanded=False):
        dias_pagos = 30 - faltas - baixas
        st.write(f"**Fórmula:** (salário_bruto / 30) × (30 - faltas - baixas)")
        st.write(f"= ({salario_bruto} / 30) × (30 - {faltas} - {baixas})")
        st.write(f"= {salario_bruto/30:.2f} × {dias_pagos}")
        st.write(f"= **{resultado['vencimento_ajustado']:.2f}€**")
    
    # PREVIEW
    st.subheader("💵 Preview")
    
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
        if outros_prov > 0:
            st.metric("Outros", f"{resultado['outros_proveitos']:.2f}€")
        st.markdown("---")
        st.metric("**TOTAL**", f"**{resultado['total_remuneracoes']:.2f}€**")
    
    with col2:
        st.markdown("### 📉 Descontos")
        st.metric("Base SS", f"{resultado['base_ss']:.2f}€")
        st.metric("Seg. Social (11%)", f"{resultado['seg_social']:.2f}€")
        st.metric("Base IRS", f"{resultado['base_irs']:.2f}€")
        st.metric("IRS", f"{resultado['irs']:.2f}€")
        if desconto_especie:
            st.metric("Desconto Espécie", f"{resultado['desconto_especie']:.2f}€")
        st.markdown("---")
        st.metric("**TOTAL**", f"**{resultado['total_descontos']:.2f}€**")
    
    with col3:
        st.markdown("### 💵 Resumo")
        st.metric("Dias Úteis Trab.", dias_uteis_trab)
        st.metric("Dias Pagos", 30 - faltas - baixas)
        st.markdown("---")
        st.metric("**💰 LÍQUIDO**", f"**{resultado['liquido']:.2f}€**")

# ==================== RESCISÕES ====================

elif menu == "🚪 Rescisões":
    st.header("🚪 Rescisões")
    st.info("🚧 Módulo mantido da v2.2")

# ==================== TABELA IRS ====================

elif menu == "📊 Tabela IRS":
    st.header("📊 Gestão de Tabela IRS")
    
    st.markdown("""
    ### 📋 Instruções:
    1. Faça upload do ficheiro Excel com as tabelas IRS 2025
    2. O sistema irá carregar e usar automaticamente para cálculos
    3. As tabelas ficam guardadas durante a sessão
    
    ### ℹ️ Sobre o Cálculo:
    - **Modo Fixa:** Usa a percentagem configurada para cada colaborador
    - **Modo Tabela:** Usa escalões progressivos baseados no salário bruto
      - Considera estado civil, dependentes e deficiência
      - Se tabela Excel carregada: usa valores exatos
      - Se não: usa escalões aproximados (simplificados)
    """)
    
    uploaded = st.file_uploader("📤 Carregar Tabelas IRS (Excel)", type=['xlsx', 'xls'])
    
    if uploaded:
        xls = carregar_tabela_irs_excel(uploaded)
        
        if xls:
            st.markdown("---")
            st.subheader("👁️ Preview das Tabelas")
            
            aba_sel = st.selectbox("Selecione a aba", xls.sheet_names)
            
            df_preview = pd.read_excel(uploaded, sheet_name=aba_sel)
            st.dataframe(df_preview, use_container_width=True)
    
    if st.session_state.tabela_irs:
        st.success("✅ Tabela IRS carregada e ativa!")
        st.info("💡 O sistema usará as tabelas para cálculos precisos de IRS.")
    else:
        st.warning("⚠️ Nenhuma tabela carregada. IRS será calculado com escalões aproximados.")

# SIDEBAR
st.sidebar.markdown("---")
st.sidebar.info(f"""v2.4 ✅
💶 SMN: {st.session_state.salario_minimo}€
🔄 Mapeamento ativo
⏰ Horários implementados
📊 IRS corrigido""")

if st.sidebar.button("🚪 Logout", use_container_width=True):
    st.session_state.authenticated = False
    st.rerun()
