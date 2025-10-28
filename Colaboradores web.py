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
    page_title="Processamento Salarial",
    page_icon="💰",
    layout="wide"
)

# Configuração OAuth 2 Dropbox
DROPBOX_APP_KEY = st.secrets["DROPBOX_APP_KEY"]
DROPBOX_APP_SECRET = st.secrets["DROPBOX_APP_SECRET"]
DROPBOX_REFRESH_TOKEN = st.secrets["DROPBOX_REFRESH_TOKEN"]
ADMIN_PASSWORD = st.secrets.get("ADMIN_PASSWORD", "adminpedro")

dbx = dropbox.Dropbox(
    app_key=DROPBOX_APP_KEY,
    app_secret=DROPBOX_APP_SECRET,
    oauth2_refresh_token=DROPBOX_REFRESH_TOKEN
)

# Configuração das empresas
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

# Feriados nacionais 2025
FERIADOS_NACIONAIS_2025 = [
    date(2025, 1, 1), date(2025, 4, 18), date(2025, 4, 20), date(2025, 4, 25),
    date(2025, 5, 1), date(2025, 6, 10), date(2025, 6, 19), date(2025, 8, 15),
    date(2025, 10, 5), date(2025, 11, 1), date(2025, 12, 1), date(2025, 12, 8),
    date(2025, 12, 25)
]

# Inicializar session state
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'salario_minimo' not in st.session_state:
    st.session_state.salario_minimo = 870.0
if 'feriados_municipais' not in st.session_state:
    st.session_state.feriados_municipais = [date(2025, 1, 14)]
if 'dados_processamento' not in st.session_state:
    st.session_state.dados_processamento = {}

# ==================== FUNÇÕES ====================

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

def carregar_colaboradores(empresa, force_reload=False):
    cache_key = f"df_colaboradores_{empresa}"
    if force_reload or cache_key not in st.session_state:
        try:
            _, response = dbx.files_download(EMPRESAS[empresa]["path"])
            df = pd.read_excel(BytesIO(response.content), sheet_name="Colaboradores")
            st.session_state[cache_key] = df
            return df
        except Exception as e:
            st.error(f"Erro ao carregar colaboradores: {e}")
            return pd.DataFrame()
    return st.session_state[cache_key]

def garantir_aba_config(empresa):
    try:
        file_path = EMPRESAS[empresa]["path"]
        _, response = dbx.files_download(file_path)
        wb = load_workbook(BytesIO(response.content))
        
        if "Config_Colaboradores" not in wb.sheetnames:
            ws = wb.create_sheet("Config_Colaboradores")
            ws.append(["Nome Completo", "Subsídio Alimentação Diário", "Última Atualização"])
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            dbx.files_upload(output.read(), file_path, mode=dropbox.files.WriteMode.overwrite)
        return True
    except Exception as e:
        st.error(f"Erro ao criar aba de configurações: {e}")
        return False

def atualizar_subsidio_colaborador(empresa, nome_colaborador, novo_valor):
    try:
        file_path = EMPRESAS[empresa]["path"]
        garantir_aba_config(empresa)
        
        _, response = dbx.files_download(file_path)
        wb = load_workbook(BytesIO(response.content))
        ws = wb["Config_Colaboradores"]
        
        colaborador_row = None
        for row in range(2, ws.max_row + 1):
            if ws.cell(row, 1).value == nome_colaborador:
                colaborador_row = row
                break
        
        if colaborador_row is None:
            colaborador_row = ws.max_row + 1
            ws.cell(colaborador_row, 1).value = nome_colaborador
        
        ws.cell(colaborador_row, 2).value = float(novo_valor)
        ws.cell(colaborador_row, 3).value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        dbx.files_upload(output.read(), file_path, mode=dropbox.files.WriteMode.overwrite)
        
        st.success(f"✅ GRAVADO: {nome_colaborador} → {novo_valor}€ (Linha {colaborador_row})")
        return True
    except Exception as e:
        st.error(f"❌ ERRO: {str(e)}")
        return False

def carregar_colaboradores_completo(empresa, force_reload=False):
    cache_key = f"df_colaboradores_completo_{empresa}"
    if force_reload or cache_key not in st.session_state:
        try:
            df_base = carregar_colaboradores(empresa, force_reload=True)
            file_path = EMPRESAS[empresa]["path"]
            _, response = dbx.files_download(file_path)
            
            try:
                df_config = pd.read_excel(BytesIO(response.content), sheet_name="Config_Colaboradores")
                for idx, row in df_base.iterrows():
                    nome = row['Nome Completo']
                    config = df_config[df_config['Nome Completo'] == nome]
                    if not config.empty:
                        df_base.at[idx, 'Subsídio Alimentação Diário'] = config.iloc[0]['Subsídio Alimentação Diário']
            except:
                pass
            
            st.session_state[cache_key] = df_base
            return df_base
        except Exception as e:
            st.error(f"Erro ao carregar dados: {e}")
            return pd.DataFrame()
    return st.session_state[cache_key]

def carregar_horas_extras(empresa, mes, ano):
    try:
        _, response = dbx.files_download(EMPRESAS[empresa]["path"])
        df = pd.read_excel(BytesIO(response.content), sheet_name="Horas extra")
        df_filtrado = df[(df['Mês'] == mes) & (df['Ano'] == ano)]
        return df_filtrado
    except Exception as e:
        st.warning(f"Aviso: Não foi possível carregar horas extras. {e}")
        return pd.DataFrame()

def calcular_dias_uteis(ano, mes, feriados_list):
    num_dias = calendar.monthrange(ano, mes)[1]
    dias_uteis = 0
    for dia in range(1, num_dias + 1):
        data = date(ano, mes, dia)
        if data.weekday() < 5 and data not in feriados_list:
            dias_uteis += 1
    return dias_uteis

def calcular_salario_base(horas_semana, salario_minimo):
    if horas_semana == 40:
        return salario_minimo
    elif horas_semana == 20:
        return salario_minimo / 2
    elif horas_semana == 16:
        return salario_minimo * 0.4
    return 0

def guardar_processamento_dropbox(empresa, mes, ano, dados_processamento):
    try:
        file_path = EMPRESAS[empresa]["path"]
        _, response = dbx.files_download(file_path)
        wb = load_workbook(BytesIO(response.content))
        
        sheet_name = f"Processamento_{ano}_{mes:02d}"
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
        
        ws = wb.create_sheet(sheet_name)
        df = pd.DataFrame([dados_processamento])
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        dbx.files_upload(output.read(), file_path, mode=dropbox.files.WriteMode.overwrite)
        return True
    except Exception as e:
        st.error(f"Erro ao guardar processamento: {e}")
        return False

def carregar_processamento_dropbox(empresa, mes, ano):
    try:
        file_path = EMPRESAS[empresa]["path"]
        _, response = dbx.files_download(file_path)
        sheet_name = f"Processamento_{ano}_{mes:02d}"
        df = pd.read_excel(BytesIO(response.content), sheet_name=sheet_name)
        if not df.empty:
            return df.iloc[0].to_dict()
        return None
    except:
        return None

# ==================== VERIFICAR AUTENTICAÇÃO ====================
if not check_password():
    st.stop()

# ==================== INTERFACE PRINCIPAL ====================
st.title("💰 Processamento Salarial")
st.markdown("---")

menu = st.sidebar.radio("Menu", ["⚙️ Configurações", "💼 Processar Salários", "📊 Relatórios"])

# CONFIGURAÇÕES
if menu == "⚙️ Configurações":
    st.header("⚙️ Configurações do Sistema")
    tab_config1, tab_config2 = st.tabs(["💶 Sistema", "👥 Colaboradores"])
    
    with tab_config1:
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("💶 Salário Mínimo Nacional")
            novo_salario = st.number_input("Valor atual (€)", min_value=0.0, value=st.session_state.salario_minimo, step=10.0, format="%.2f")
            if st.button("Atualizar Salário Mínimo"):
                st.session_state.salario_minimo = novo_salario
                st.success(f"✅ Salário mínimo atualizado para {novo_salario}€")
        
        with col2:
            st.subheader("📅 Feriados Municipais")
            st.caption("Adicione até 3 feriados municipais")
            feriados_temp = []
            for i in range(3):
                feriado = st.date_input(f"Feriado Municipal {i+1}", value=st.session_state.feriados_municipais[i] if i < len(st.session_state.feriados_municipais) else None, key=f"feriado_{i}")
                if feriado:
                    feriados_temp.append(feriado)
            if st.button("Atualizar Feriados"):
                st.session_state.feriados_municipais = feriados_temp
                st.success(f"✅ {len(feriados_temp)} feriados municipais configurados")
    
    with tab_config2:
        st.subheader("👥 Editar Dados de Colaboradores")
        empresa_config = st.selectbox("Empresa", options=list(EMPRESAS.keys()), key="empresa_config")
        df_colab_config = carregar_colaboradores_completo(empresa_config, force_reload=True)
        
        if not df_colab_config.empty:
            colaborador_config = st.selectbox("Colaborador", options=df_colab_config['Nome Completo'].tolist(), key="colab_config")
            dados_atual = df_colab_config[df_colab_config['Nome Completo'] == colaborador_config].iloc[0]
            st.markdown("---")
            st.info(f"📊 Valor atual: {dados_atual.get('Subsídio Alimentação Diário', 'N/A')}€")
            
            with st.form("form_editar_colab"):
                st.markdown(f"### Editar: {colaborador_config}")
                novo_sub_alim = st.number_input("Subsídio de Alimentação Diário (€)", min_value=0.0, value=float(dados_atual.get('Subsídio Alimentação Diário', 0)), step=0.10, format="%.2f")
                
                if st.form_submit_button("💾 Guardar Alterações", use_container_width=True):
                    with st.spinner("🔄 A guardar na aba Config_Colaboradores..."):
                        sucesso = atualizar_subsidio_colaborador(empresa_config, colaborador_config, novo_sub_alim)
                        if sucesso:
                            for key in list(st.session_state.keys()):
                                if 'colaboradores' in key.lower() and key != 'authenticated':
                                    del st.session_state[key]
                            st.balloons()
                            time.sleep(2)
                            st.rerun()

# PROCESSAR SALÁRIOS
elif menu == "💼 Processar Salários":
    st.header("💼 Processamento Mensal de Salários")
    st.info("🚧 Módulo em construção - base implementada!")

# RELATÓRIOS
elif menu == "📊 Relatórios":
    st.header("📊 Relatórios e Histórico")
    st.info("🚧 Em desenvolvimento...")

st.sidebar.markdown("---")
if st.sidebar.button("🚪 Logout"):
    st.session_state.authenticated = False
    st.rerun()
