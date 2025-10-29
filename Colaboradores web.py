import streamlit as st
import pandas as pd
import dropbox
from datetime import datetime, date, timedelta
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from copy import copy
import calendar
import time

st.set_page_config(
    page_title="Processamento Salarial v2.8",
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
        "tem_horas_extras": False,
        "pasta_baixas": "/Pedro Couto/Projectos/Alcalá_Arc_Amoreira/Gestão operacional/RH/Baixas Médicas"
    },
    "CCM Retail Lda": {
        "path": "/Pedro Couto/Projectos/Pingo Doce/Pingo Doce/2. Operação/1. Recursos Humanos/Processamento salarial/Gestão Colaboradores.xlsx",
        "tem_horas_extras": True,
        "pasta_baixas": "/Pedro Couto/Projectos/Pingo Doce/Pingo Doce/2. Operação/1. Recursos Humanos/Baixas Médicas"
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
    "Despedimento por facto imputável ao trabalhador",
    "Despedimento colectivo",
    "Despedimento por extinção do posto de trabalho",
    "Despedimento por inadaptação",
    "Revogação por acordo",
    "Reforma por velhice",
    "Reforma por invalidez",
    "Falecimento",
    "Outro (especificar em observações)"
]

COLUNAS_SNAPSHOT = [
    "Nome Completo", "Ano", "Mês", "Nº Horas/Semana", "Subsídio Alimentação Diário",
    "Número Pingo Doce", "Salário Bruto", "Vencimento Hora", 
    "Estado Civil", "Nº Titulares", "Nº Dependentes", "Deficiência",
    "IRS Percentagem Fixa", "IRS Modo Calculo",
    "Status", "Data Rescisão", "Motivo Rescisão", 
    "NIF", "NISS", "Data de Admissão", "IBAN", "Secção", "Timestamp"
]

COLUNAS_FALTAS_BAIXAS = [
    "Nome Completo", "Ano", "Mês", "Tipo", "Data Início", "Data Fim", 
    "Dias Úteis", "Dias Totais", "Observações", "Ficheiro Anexo", "Timestamp"
]

ESTADOS_CIVIS = ["Solteiro", "Casado Único Titular", "Casado Dois Titulares"]
HORAS_PERMITIDAS = [16, 20, 40]

MAPEAMENTO_ESTADO_CIVIL = {
    "Não Casado": "Solteiro",
    "Casado 1": "Casado Único Titular",
    "Casado 2": "Casado Dois Titulares",
    "Solteiro": "Solteiro",
    "Casado Único Titular": "Casado Único Titular",
    "Casado Dois Titulares": "Casado Dois Titulares"
}

MAPEAMENTO_TIPO_IRS = {
    "Automático (por tabela)": "Tabela",
    "Percentagem fixa": "Fixa",
    "Tabela": "Tabela",
    "Fixa": "Fixa",
    "Percentagem Fixa": "Fixa"
}

MAPEAMENTO_DEFICIENCIA = {
    "Sim": "Sim",
    "Não": "Não",
    "sim": "Sim",
    "não": "Não",
    "S": "Sim",
    "N": "Não"
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
if 'dados_processamento' not in st.session_state:
    st.session_state.dados_processamento = {}

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
    if pd.isna(valor) or valor == '':
        return "Solteiro"
    valor_str = str(valor).strip()
    return MAPEAMENTO_ESTADO_CIVIL.get(valor_str, "Solteiro")

def normalizar_tipo_irs(valor):
    if pd.isna(valor) or valor == '':
        return "Tabela"
    valor_str = str(valor).strip()
    return MAPEAMENTO_TIPO_IRS.get(valor_str, "Tabela")

def normalizar_deficiencia(valor):
    if pd.isna(valor) or valor == '':
        return "Não"
    valor_str = str(valor).strip()
    return MAPEAMENTO_DEFICIENCIA.get(valor_str, "Não")

def normalizar_percentagem_irs(valor):
    if pd.isna(valor) or valor == '':
        return 0.0
    try:
        return float(valor)
    except:
        return 0.0

# ==================== FUNÇÕES DROPBOX ====================

def get_nome_aba_snapshot(ano, mes):
    return f"Estado_{ano}_{mes:02d}"

def get_nome_aba_faltas_baixas(ano, mes):
    return f"Faltas_Baixas_{ano}_{mes:02d}"

def criar_pasta_dropbox(path):
    """Cria pasta na Dropbox se não existir"""
    try:
        dbx.files_get_metadata(path)
        return True
    except:
        try:
            dbx.files_create_folder_v2(path)
            return True
        except Exception as e:
            st.error(f"Erro ao criar pasta: {e}")
            return False

def upload_ficheiro_baixa(empresa, ano, mes, colaborador, file):
    """Upload de ficheiro de baixa médica para Dropbox"""
    try:
        pasta_base = EMPRESAS[empresa]["pasta_baixas"]
        pasta_ano = f"{pasta_base}/{ano}"
        pasta_mes = f"{pasta_ano}/{mes:02d}_{calendar.month_name[mes]}"
        
        # Criar estrutura de pastas
        criar_pasta_dropbox(pasta_base)
        criar_pasta_dropbox(pasta_ano)
        criar_pasta_dropbox(pasta_mes)
        
        # Nome do ficheiro
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_limpo = colaborador.replace(" ", "_")
        extensao = file.name.split(".")[-1]
        nome_ficheiro = f"{nome_limpo}_{timestamp}.{extensao}"
        
        caminho_completo = f"{pasta_mes}/{nome_ficheiro}"
        
        # Upload
        file.seek(0)
        dbx.files_upload(file.read(), caminho_completo, mode=dropbox.files.WriteMode.overwrite)
        
        return caminho_completo
        
    except Exception as e:
        st.error(f"❌ Erro ao fazer upload: {e}")
        return None

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

def upload_excel_seguro(empresa, wb):
    """
    CRÍTICO: Upload com verificação de integridade
    Garante que aba Colaboradores não é corrompida
    """
    try:
        # Verificar se aba Colaboradores existe
        if "Colaboradores" not in wb.sheetnames:
            st.error("🚨 ERRO CRÍTICO: Aba 'Colaboradores' não encontrada no workbook!")
            st.error("Upload CANCELADO para proteger dados!")
            return False
        
        # Verificar se aba tem dados
        ws_colab = wb["Colaboradores"]
        if ws_colab.max_row < 2:  # Menos de 2 linhas = só header ou vazio
            st.error("🚨 ERRO CRÍTICO: Aba 'Colaboradores' está vazia!")
            st.error("Upload CANCELADO para proteger dados!")
            return False
        
        file_path = EMPRESAS[empresa]["path"]
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        dbx.files_upload(output.read(), file_path, mode=dropbox.files.WriteMode.overwrite)
        
        st.success(f"✅ Excel salvo com segurança ({ws_colab.max_row-1} colaboradores preservados)")
        return True
        
    except Exception as e:
        st.error(f"❌ Erro ao enviar Excel: {e}")
        st.error(f"Detalhes: {str(e)}")
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
    dias_pagos = 30 - dias_faltas - dias_baixas
    dias_pagos = max(dias_pagos, 0)
    return (salario_bruto / 30) * dias_pagos

def calcular_dias_entre_datas(data_inicio, data_fim, feriados_list):
    """Calcula dias úteis e totais entre duas datas"""
    if data_inicio > data_fim:
        return 0, 0
    
    dias_totais = (data_fim - data_inicio).days + 1
    dias_uteis = 0
    
    data_atual = data_inicio
    while data_atual <= data_fim:
        if data_atual.weekday() < 5 and data_atual not in feriados_list:
            dias_uteis += 1
        data_atual += timedelta(days=1)
    
    return dias_uteis, dias_totais

def calcular_dias_uteis(ano, mes, feriados_list):
    num_dias = calendar.monthrange(ano, mes)[1]
    dias_uteis = 0
    for dia in range(1, num_dias + 1):
        data = date(ano, mes, dia)
        if data.weekday() < 5 and data not in feriados_list:
            dias_uteis += 1
    return dias_uteis

def carregar_tabela_irs_excel(uploaded_file):
    try:
        xls = pd.ExcelFile(uploaded_file)
        st.success(f"✅ Ficheiro carregado! Abas encontradas: {', '.join(xls.sheet_names)}")
        st.session_state.tabela_irs = xls
        return xls
    except Exception as e:
        st.error(f"❌ Erro ao carregar tabela: {e}")
        return None

def calcular_irs_por_tabela(base_incidencia, estado_civil, num_dependentes, tem_deficiencia=False):
    if st.session_state.tabela_irs is None:
        pass  # Silencioso, apenas usa escalões
    
    reducao_dependentes = num_dependentes * 0.01
    
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
    
    taxa_final = max(taxa - reducao_dependentes, 0.05)
    
    if estado_civil == "Casado Único Titular":
        taxa_final *= 0.85
    
    return base_incidencia * taxa_final

def calcular_irs(base_incidencia, modo_calculo, percentagem_fixa, estado_civil, num_dependentes, tem_deficiencia=False):
    if modo_calculo == "Fixa":
        taxa = percentagem_fixa / 100
        irs = base_incidencia * taxa
        return irs
    else:
        return calcular_irs_por_tabela(base_incidencia, estado_civil, num_dependentes, tem_deficiencia)

# ==================== FUNÇÕES DE DADOS BASE ====================

def carregar_dados_base(empresa):
    """
    FONTE DA VERDADE: Lê sempre da aba 'Colaboradores'
    """
    excel_file = download_excel(empresa)
    if excel_file:
        try:
            df = pd.read_excel(excel_file, sheet_name="Colaboradores")
            
            if 'Status' not in df.columns:
                df['Status'] = 'Ativo'
            
            df.loc[df['Status'].isna() | (df['Status'] == ''), 'Status'] = 'Ativo'
            
            return df
        except Exception as e:
            st.error(f"❌ Erro ao ler aba Colaboradores: {e}")
    return pd.DataFrame()

def carregar_colaboradores_ativos(empresa, ano=None, mes=None):
    """
    Lê da aba 'Colaboradores' onde Status = 'Ativo'
    """
    df_base = carregar_dados_base(empresa)
    
    if df_base.empty:
        return []
    
    if 'Status' in df_base.columns:
        df_ativos = df_base[df_base['Status'] == 'Ativo']
    else:
        df_ativos = df_base
    
    colaboradores = df_ativos['Nome Completo'].tolist()
    
    return colaboradores

def atualizar_status_colaborador(empresa, colaborador, novo_status):
    """
    Atualiza Status APENAS na aba Colaboradores
    """
    try:
        excel_file = download_excel(empresa)
        if not excel_file:
            return False
        
        df = pd.read_excel(excel_file, sheet_name="Colaboradores")
        
        if 'Status' not in df.columns:
            df['Status'] = 'Ativo'
        
        mask = df['Nome Completo'] == colaborador
        if mask.any():
            df.loc[mask, 'Status'] = novo_status
        else:
            st.error(f"❌ Colaborador '{colaborador}' não encontrado")
            return False
        
        # Carregar workbook completo
        wb = load_workbook(excel_file, data_only=False)
        
        # Apagar e recriar APENAS aba Colaboradores
        if "Colaboradores" in wb.sheetnames:
            idx = wb.sheetnames.index("Colaboradores")
            del wb["Colaboradores"]
            ws = wb.create_sheet("Colaboradores", idx)
        else:
            ws = wb.create_sheet("Colaboradores", 0)
        
        # Escrever dados
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        
        if upload_excel_seguro(empresa, wb):
            st.success(f"✅ Status de '{colaborador}' → '{novo_status}'")
            return True
        
        return False
        
    except Exception as e:
        st.error(f"❌ Erro: {e}")
        return False

def garantir_coluna_status(empresa):
    """
    Adiciona coluna Status se não existir
    """
    try:
        excel_file = download_excel(empresa)
        if not excel_file:
            return False
        
        wb = load_workbook(excel_file, data_only=False)
        
        if "Colaboradores" not in wb.sheetnames:
            st.warning("⚠️ Aba 'Colaboradores' não encontrada")
            return False
        
        df = pd.read_excel(excel_file, sheet_name="Colaboradores")
        
        if 'Status' not in df.columns:
            st.info("🔧 Adicionando coluna 'Status'...")
            df['Status'] = 'Ativo'
            alterado = True
        else:
            status_vazios = df['Status'].isna() | (df['Status'] == '')
            if status_vazios.any():
                st.info(f"🔧 Preenchendo {status_vazios.sum()} registos...")
                df.loc[status_vazios, 'Status'] = 'Ativo'
                alterado = True
            else:
                alterado = False
        
        if alterado:
            if "Colaboradores" in wb.sheetnames:
                idx = wb.sheetnames.index("Colaboradores")
                del wb["Colaboradores"]
                ws = wb.create_sheet("Colaboradores", idx)
            else:
                ws = wb.create_sheet("Colaboradores", 0)
            
            for r in dataframe_to_rows(df, index=False, header=True):
                ws.append(r)
            
            if upload_excel_seguro(empresa, wb):
                st.success("✅ Coluna Status adicionada/atualizada!")
                return True
        
        return True
        
    except Exception as e:
        st.error(f"❌ Erro: {e}")
        return False

def criar_snapshot_inicial(empresa, colaborador, ano, mes):
    df_base = carregar_dados_base(empresa)
    dados_colab = df_base[df_base['Nome Completo'] == colaborador]
    
    if dados_colab.empty:
        return None
    
    dados = dados_colab.iloc[0]
    horas_semana = float(dados.get('Nº Horas/Semana', 40))
    salario_bruto = calcular_salario_base(horas_semana, st.session_state.salario_minimo)
    
    estado_civil_raw = dados.get('Estado Civil', 'Solteiro')
    estado_civil = normalizar_estado_civil(estado_civil_raw)
    
    tipo_irs_raw = dados.get('Tipo IRS', 'Tabela')
    tipo_irs = normalizar_tipo_irs(tipo_irs_raw)
    
    perc_irs_raw = dados.get('% IRS Fixa', 0)
    perc_irs = normalizar_percentagem_irs(perc_irs_raw)
    
    deficiencia_raw = dados.get('Pessoa com Deficiência', 'Não')
    deficiencia = normalizar_deficiencia(deficiencia_raw)
    
    status = dados.get('Status', 'Ativo')
    
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
        "Deficiência": deficiencia,
        "IRS Percentagem Fixa": perc_irs,
        "IRS Modo Calculo": tipo_irs,
        "Status": status,
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
    """Carrega último snapshot com dados ATUALIZADOS das configurações"""
    excel_file = download_excel(empresa)
    if not excel_file:
        return None
    
    try:
        wb = load_workbook(excel_file, data_only=False)
        nome_aba = get_nome_aba_snapshot(ano, mes)
        
        # Primeiro tenta aba do mês atual
        if nome_aba in wb.sheetnames:
            df = pd.read_excel(excel_file, sheet_name=nome_aba)
            df_colab = df[df['Nome Completo'] == colaborador]
            
            if not df_colab.empty:
                snapshot = df_colab.iloc[-1].to_dict()
                
                # ATUALIZAR com dados da aba Colaboradores
                df_base = carregar_dados_base(empresa)
                dados_colab = df_base[df_base['Nome Completo'] == colaborador]
                
                if not dados_colab.empty:
                    dados = dados_colab.iloc[0]
                    
                    # Atualizar campos que podem ter mudado
                    snapshot['Nº Horas/Semana'] = float(dados.get('Nº Horas/Semana', snapshot.get('Nº Horas/Semana', 40)))
                    snapshot['Subsídio Alimentação Diário'] = float(dados.get('Subsídio Alimentação Diário', snapshot.get('Subsídio Alimentação Diário', 5.96)))
                    snapshot['Número Pingo Doce'] = str(dados.get('Número Pingo Doce', snapshot.get('Número Pingo Doce', '')))
                    
                    # Recalcular salário com base nas horas atualizadas
                    horas = float(snapshot['Nº Horas/Semana'])
                    snapshot['Salário Bruto'] = calcular_salario_base(horas, st.session_state.salario_minimo)
                    snapshot['Vencimento Hora'] = calcular_vencimento_hora(snapshot['Salário Bruto'], horas)
                    
                    # Atualizar dados IRS
                    snapshot['Estado Civil'] = normalizar_estado_civil(dados.get('Estado Civil', snapshot.get('Estado Civil', 'Solteiro')))
                    snapshot['Nº Titulares'] = int(dados.get('Nº Titulares', snapshot.get('Nº Titulares', 2)))
                    snapshot['Nº Dependentes'] = int(dados.get('Nº Dependentes', snapshot.get('Nº Dependentes', 0)))
                    snapshot['Deficiência'] = normalizar_deficiencia(dados.get('Pessoa com Deficiência', snapshot.get('Deficiência', 'Não')))
                    snapshot['IRS Modo Calculo'] = normalizar_tipo_irs(dados.get('Tipo IRS', snapshot.get('IRS Modo Calculo', 'Tabela')))
                    snapshot['IRS Percentagem Fixa'] = normalizar_percentagem_irs(dados.get('% IRS Fixa', snapshot.get('IRS Percentagem Fixa', 0)))
                
                if 'Status' not in snapshot or pd.isna(snapshot['Status']) or snapshot['Status'] == '':
                    snapshot['Status'] = 'Ativo'
                
                st.caption(f"📸 Snapshot {ano}-{mes:02d}: {snapshot.get('Timestamp', 'N/A')} (dados atualizados)")
                return snapshot
        
        # Se não existe snapshot do mês, cria novo com dados da aba Colaboradores
        snapshot = criar_snapshot_inicial(empresa, colaborador, ano, mes)
        if snapshot:
            st.caption(f"📸 Criado da aba Colaboradores (dados atuais)")
        return snapshot
        
    except Exception as e:
        st.error(f"❌ Erro: {e}")
        return None

def gravar_snapshot(empresa, snapshot):
    """
    CRÍTICO: Grava snapshot SEM mexer na aba Colaboradores
    """
    try:
        if 'Status' not in snapshot or pd.isna(snapshot['Status']) or snapshot['Status'] == '':
            snapshot['Status'] = 'Ativo'
        
        ano = snapshot['Ano']
        mes = snapshot['Mês']
        nome_aba = get_nome_aba_snapshot(ano, mes)
        
        excel_file = download_excel(empresa)
        if not excel_file:
            return False
        
        # Carregar workbook COM SEGURANÇA
        wb = load_workbook(excel_file, data_only=False, keep_vba=True)
        
        # Verificar que aba Colaboradores existe
        if "Colaboradores" not in wb.sheetnames:
            st.error("🚨 ERRO: Aba Colaboradores não encontrada!")
            return False
        
        # Contar colaboradores ANTES
        ws_colab_antes = wb["Colaboradores"]
        num_colab_antes = ws_colab_antes.max_row - 1  # -1 para header
        
        # Criar/obter aba snapshot
        aba_criada = garantir_aba(wb, nome_aba, COLUNAS_SNAPSHOT)
        if aba_criada:
            st.info(f"✨ Aba '{nome_aba}' criada")
        
        ws = wb[nome_aba]
        
        # Adicionar linha ao snapshot
        nova_linha = []
        for col in COLUNAS_SNAPSHOT:
            valor = snapshot.get(col, '')
            if isinstance(valor, (int, float)):
                nova_linha.append(valor)
            else:
                nova_linha.append(str(valor) if valor else '')
        
        ws.append(nova_linha)
        
        # Verificar que aba Colaboradores AINDA existe
        if "Colaboradores" not in wb.sheetnames:
            st.error("🚨 ERRO: Aba Colaboradores desapareceu durante operação!")
            return False
        
        # Contar colaboradores DEPOIS
        ws_colab_depois = wb["Colaboradores"]
        num_colab_depois = ws_colab_depois.max_row - 1
        
        # Verificar integridade
        if num_colab_depois < num_colab_antes:
            st.error(f"🚨 ERRO: Perderam-se colaboradores! Antes: {num_colab_antes}, Depois: {num_colab_depois}")
            st.error("Upload CANCELADO!")
            return False
        
        # Upload com segurança
        sucesso = upload_excel_seguro(empresa, wb)
        
        if sucesso:
            linha = ws.max_row
            st.success(f"✅ Snapshot gravado (linha {linha})")
            return True
        
        return False
        
    except Exception as e:
        st.error(f"❌ Erro ao gravar: {e}")
        st.error(f"Detalhes: {str(e)}")
        return False

def gravar_falta_baixa(empresa, ano, mes, colaborador, tipo, data_inicio, data_fim, obs, ficheiro_path=None):
    """Grava registo de falta ou baixa"""
    try:
        excel_file = download_excel(empresa)
        if not excel_file:
            return False
        
        wb = load_workbook(excel_file, data_only=False, keep_vba=True)
        
        # Verificar integridade
        if "Colaboradores" not in wb.sheetnames:
            st.error("🚨 ERRO: Aba Colaboradores não encontrada!")
            return False
        
        nome_aba = get_nome_aba_faltas_baixas(ano, mes)
        garantir_aba(wb, nome_aba, COLUNAS_FALTAS_BAIXAS)
        
        ws = wb[nome_aba]
        
        # Calcular dias
        feriados = FERIADOS_NACIONAIS_2025 + st.session_state.feriados_municipais
        dias_uteis, dias_totais = calcular_dias_entre_datas(data_inicio, data_fim, feriados)
        
        # Nova linha
        nova_linha = [
            colaborador,
            ano,
            mes,
            tipo,
            data_inicio.strftime("%Y-%m-%d"),
            data_fim.strftime("%Y-%m-%d"),
            dias_uteis,
            dias_totais,
            obs,
            ficheiro_path if ficheiro_path else "",
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ]
        
        ws.append(nova_linha)
        
        # Upload com segurança
        if upload_excel_seguro(empresa, wb):
            st.success(f"✅ {tipo} registada: {dias_uteis} dias úteis / {dias_totais} dias totais")
            return True
        
        return False
        
    except Exception as e:
        st.error(f"❌ Erro ao gravar: {e}")
        return False

def carregar_faltas_baixas(empresa, ano, mes, colaborador=None):
    """Carrega faltas e baixas do mês"""
    try:
        excel_file = download_excel(empresa)
        if not excel_file:
            return pd.DataFrame()
        
        nome_aba = get_nome_aba_faltas_baixas(ano, mes)
        
        try:
            df = pd.read_excel(excel_file, sheet_name=nome_aba)
            
            if colaborador:
                df = df[df['Nome Completo'] == colaborador]
            
            return df
        except:
            return pd.DataFrame()
            
    except Exception as e:
        st.error(f"❌ Erro: {e}")
        return pd.DataFrame()

def atualizar_campo_colaborador(empresa, colaborador, ano, mes, campo, novo_valor):
    snapshot = carregar_ultimo_snapshot(empresa, colaborador, ano, mes)
    
    if not snapshot:
        return False
    
    status_original = snapshot.get('Status', 'Ativo')
    
    snapshot[campo] = novo_valor
    snapshot['Ano'] = ano
    snapshot['Mês'] = mes
    snapshot['Status'] = status_original
    snapshot['Timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    if campo == "Nº Horas/Semana":
        horas = float(novo_valor)
        snapshot['Salário Bruto'] = calcular_salario_base(horas, st.session_state.salario_minimo)
        snapshot['Vencimento Hora'] = calcular_vencimento_hora(snapshot['Salário Bruto'], horas)
    
    return gravar_snapshot(empresa, snapshot)

def processar_calculo_salario(dados_form):
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
    
    base_ss = total_remuneracoes - sub_alimentacao
    seg_social = base_ss * 0.11
    
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

def registar_rescisao(empresa, colaborador, ano, mes, data_rescisao, motivo, obs):
    """Regista rescisão e atualiza status"""
    try:
        # Atualizar snapshot
        snapshot = carregar_ultimo_snapshot(empresa, colaborador, ano, mes)
        
        if not snapshot:
            return False
        
        snapshot['Status'] = 'Rescindido'
        snapshot['Data Rescisão'] = data_rescisao.strftime("%Y-%m-%d")
        snapshot['Motivo Rescisão'] = f"{motivo} | Obs: {obs}"
        snapshot['Timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        if not gravar_snapshot(empresa, snapshot):
            return False
        
        # Atualizar status na aba Colaboradores
        if not atualizar_status_colaborador(empresa, colaborador, 'Rescindido'):
            return False
        
        return True
        
    except Exception as e:
        st.error(f"❌ Erro: {e}")
        return False

# ==================== INTERFACE ====================

if not check_password():
    st.stop()

st.title("💰 Processamento Salarial v2.8")
st.caption("✨ NOVO: Gestão de faltas/baixas com datas + Upload de documentos + Rescisões completo")
st.caption(f"🕐 Reload: {st.session_state.ultimo_reload.strftime('%H:%M:%S')}")

st.markdown("---")

menu = st.sidebar.radio(
    "Menu Principal",
    ["⚙️ Configurações", "💼 Processar Salários", "🔧 Gestão Status", "🚪 Rescisões", "📊 Tabela IRS"],
    index=0
)

# ==================== CONFIGURAÇÕES ====================

if menu == "⚙️ Configurações":
    st.header("⚙️ Configurações do Sistema")
    
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["💶 Sistema", "👥 Colaboradores", "⏰ Horários", "📋 Dados IRS", "🔧 Migrar Status"])
    
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
    
    with tab2:
        st.subheader("👥 Editar Dados")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            emp = st.selectbox("Empresa", list(EMPRESAS.keys()), key="emp_cfg")
        with col2:
            mes_cfg = st.selectbox("Mês", list(range(1, 13)), 
                                 format_func=lambda x: calendar.month_name[x],
                                 index=datetime.now().month - 1, key="mes_cfg")
        with col3:
            ano_cfg = st.selectbox("Ano", [2024, 2025, 2026], index=1, key="ano_cfg")
        
        colabs = carregar_colaboradores_ativos(emp)
        
        if colabs:
            st.success(f"✅ {len(colabs)} colaboradores ativos")
            
            colab_sel = st.selectbox("Colaborador", colabs, key="col_cfg")
            
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
                        # Atualizar na aba Colaboradores
                        df_base = carregar_dados_base(emp)
                        excel_file = download_excel(emp)
                        wb = load_workbook(excel_file, data_only=False)
                        
                        mask = df_base['Nome Completo'] == colab_sel
                        df_base.loc[mask, 'Subsídio Alimentação Diário'] = novo_sub
                        df_base.loc[mask, 'Número Pingo Doce'] = novo_num
                        
                        # Reescrever aba Colaboradores
                        if "Colaboradores" in wb.sheetnames:
                            idx = wb.sheetnames.index("Colaboradores")
                            del wb["Colaboradores"]
                            ws = wb.create_sheet("Colaboradores", idx)
                        else:
                            ws = wb.create_sheet("Colaboradores", 0)
                        
                        for r in dataframe_to_rows(df_base, index=False, header=True):
                            ws.append(r)
                        
                        if upload_excel_seguro(emp, wb):
                            st.success("✅ Dados atualizados na aba Colaboradores!")
                            st.info("💡 Mudanças serão refletidas no próximo processamento")
                            time.sleep(2)
                            st.rerun()
        else:
            st.warning("⚠️ Nenhum colaborador ativo")
    
    with tab3:
        st.subheader("⏰ Mudanças de Horário")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            emp_hor = st.selectbox("Empresa", list(EMPRESAS.keys()), key="emp_hor")
        with col2:
            mes_hor = st.selectbox("Mês", list(range(1, 13)),
                                  format_func=lambda x: calendar.month_name[x],
                                  index=datetime.now().month - 1, key="mes_hor")
        with col3:
            ano_hor = st.selectbox("Ano", [2024, 2025, 2026], index=1, key="ano_hor")
        
        colabs_hor = carregar_colaboradores_ativos(emp_hor)
        
        if colabs_hor:
            colab_hor = st.selectbox("Colaborador", colabs_hor, key="col_hor")
            
            snap_hor = carregar_ultimo_snapshot(emp_hor, colab_hor, ano_hor, mes_hor)
            
            if snap_hor:
                st.markdown("---")
                
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
                        novas_horas = st.selectbox(
                            "⏰ Novas Horas Semanais",
                            options=HORAS_PERMITIDAS,
                            index=HORAS_PERMITIDAS.index(int(horas_atuais)) if int(horas_atuais) in HORAS_PERMITIDAS else 2
                        )
                    
                    with col2:
                        novo_salario = calcular_salario_base(novas_horas, st.session_state.salario_minimo)
                        novo_venc_hora = calcular_vencimento_hora(novo_salario, novas_horas)
                        
                        st.metric("💰 Novo Salário Bruto", f"{novo_salario:.2f}€",
                                 delta=f"{novo_salario - salario_atual:.2f}€")
                        st.metric("💵 Novo Vencimento/Hora", f"{novo_venc_hora:.2f}€",
                                 delta=f"{novo_venc_hora - venc_hora_atual:.2f}€")
                    
                    submit_hor = st.form_submit_button("💾 CONFIRMAR", use_container_width=True, type="primary")
                    
                    if submit_hor:
                        if novas_horas == horas_atuais:
                            st.warning("⚠️ As horas não foram alteradas!")
                        else:
                            # Atualizar na aba Colaboradores
                            df_base = carregar_dados_base(emp_hor)
                            excel_file = download_excel(emp_hor)
                            wb = load_workbook(excel_file, data_only=False)
                            
                            mask = df_base['Nome Completo'] == colab_hor
                            df_base.loc[mask, 'Nº Horas/Semana'] = novas_horas
                            
                            # Reescrever aba Colaboradores
                            if "Colaboradores" in wb.sheetnames:
                                idx = wb.sheetnames.index("Colaboradores")
                                del wb["Colaboradores"]
                                ws = wb.create_sheet("Colaboradores", idx)
                            else:
                                ws = wb.create_sheet("Colaboradores", 0)
                            
                            for r in dataframe_to_rows(df_base, index=False, header=True):
                                ws.append(r)
                            
                            if upload_excel_seguro(emp_hor, wb):
                                st.success("✅ Horário atualizado na aba Colaboradores!")
                                st.info("💡 Novo salário será calculado no próximo processamento")
                                st.balloons()
                                time.sleep(2)
                                st.rerun()
        else:
            st.warning("⚠️ Nenhum colaborador ativo")
    
    with tab4:
        st.subheader("📋 Dados IRS")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            emp_irs = st.selectbox("Empresa", list(EMPRESAS.keys()), key="emp_irs")
        with col2:
            mes_irs = st.selectbox("Mês", list(range(1, 13)),
                                  format_func=lambda x: calendar.month_name[x],
                                  index=datetime.now().month - 1, key="mes_irs")
        with col3:
            ano_irs = st.selectbox("Ano", [2024, 2025, 2026], index=1, key="ano_irs")
        
        colabs_irs = carregar_colaboradores_ativos(emp_irs)
        
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
                        irs_percentagem = st.number_input("IRS % Fixa", min_value=0.0, max_value=100.0,
                                                         value=float(snap_irs.get('IRS Percentagem Fixa', 0)),
                                                         step=0.1, format="%.1f")
                    
                    submit_irs = st.form_submit_button("💾 GUARDAR", use_container_width=True, type="primary")
                    
                    if submit_irs:
                        # Atualizar na aba Colaboradores
                        df_base = carregar_dados_base(emp_irs)
                        excel_file = download_excel(emp_irs)
                        wb = load_workbook(excel_file, data_only=False)
                        
                        mask = df_base['Nome Completo'] == colab_irs
                        df_base.loc[mask, 'Estado Civil'] = estado_civil
                        df_base.loc[mask, 'Nº Titulares'] = num_titulares
                        df_base.loc[mask, 'Nº Dependentes'] = num_dependentes
                        df_base.loc[mask, 'Pessoa com Deficiência'] = tem_deficiencia
                        df_base.loc[mask, 'Tipo IRS'] = irs_modo
                        df_base.loc[mask, '% IRS Fixa'] = irs_percentagem
                        
                        # Reescrever aba Colaboradores
                        if "Colaboradores" in wb.sheetnames:
                            idx = wb.sheetnames.index("Colaboradores")
                            del wb["Colaboradores"]
                            ws = wb.create_sheet("Colaboradores", idx)
                        else:
                            ws = wb.create_sheet("Colaboradores", 0)
                        
                        for r in dataframe_to_rows(df_base, index=False, header=True):
                            ws.append(r)
                        
                        if upload_excel_seguro(emp_irs, wb):
                            st.success("✅ Dados IRS atualizados na aba Colaboradores!")
                            st.balloons()
                            time.sleep(2)
                            st.rerun()
        else:
            st.warning("⚠️ Nenhum colaborador ativo")
    
    with tab5:
        st.subheader("🔧 Migração: Adicionar Coluna Status")
        st.info("Execute isto UMA VEZ por empresa para adicionar coluna Status")
        
        emp_migrar = st.selectbox("Empresa", list(EMPRESAS.keys()), key="emp_migrar")
        
        if st.button("🔧 EXECUTAR MIGRAÇÃO", type="primary"):
            with st.spinner("Executando..."):
                if garantir_coluna_status(emp_migrar):
                    st.success("✅ Migração concluída!")
                    st.balloons()

# ==================== GESTÃO STATUS ====================

elif menu == "🔧 Gestão Status":
    st.header("🔧 Gestão de Status")
    
    col1, col2 = st.columns(2)
    with col1:
        emp_status = st.selectbox("Empresa", list(EMPRESAS.keys()), key="emp_status")
    with col2:
        mostrar = st.radio("Mostrar", ["Ativos", "Inativos", "Todos"], horizontal=True)
    
    df_base = carregar_dados_base(emp_status)
    
    if not df_base.empty:
        if mostrar == "Ativos":
            df_filtrado = df_base[df_base['Status'] == 'Ativo']
        elif mostrar == "Inativos":
            df_filtrado = df_base[df_base['Status'] == 'Inativo']
        else:
            df_filtrado = df_base
        
        st.markdown(f"**Total: {len(df_filtrado)} colaboradores**")
        st.markdown("---")
        
        if not df_filtrado.empty:
            for _, row in df_filtrado.iterrows():
                nome = row['Nome Completo']
                status_atual = row.get('Status', 'Ativo')
                
                col1, col2, col3, col4 = st.columns([3, 1, 1, 1])
                
                with col1:
                    st.write(f"**{nome}**")
                    st.caption(f"Secção: {row.get('Secção', 'N/A')}")
                
                with col2:
                    if status_atual == 'Ativo':
                        st.success("✅ Ativo")
                    else:
                        st.error("❌ Inativo")
                
                with col3:
                    if status_atual == 'Ativo':
                        if st.button("❌ Desativar", key=f"desativar_{nome}"):
                            if atualizar_status_colaborador(emp_status, nome, 'Inativo'):
                                st.rerun()
                
                with col4:
                    if status_atual == 'Inativo':
                        if st.button("✅ Ativar", key=f"ativar_{nome}"):
                            if atualizar_status_colaborador(emp_status, nome, 'Ativo'):
                                st.rerun()
                
                st.markdown("---")

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
    
    colabs_proc = carregar_colaboradores_ativos(emp_proc)
    
    if not colabs_proc:
        st.warning("⚠️ Nenhum colaborador ativo")
        st.info("💡 Execute a migração em Configurações → Migrar Status")
        st.stop()
    
    st.success(f"✅ {len(colabs_proc)} colaboradores ativos")
    
    colab_proc = st.selectbox("Colaborador", colabs_proc, key="col_proc")
    
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
    
    with st.expander("📋 DADOS BASE (atualizados das Configurações)", expanded=True):
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
    
    # FALTAS E BAIXAS COM DATAS
    st.subheader("🏖️ Faltas e Baixas Médicas")
    
    tab_faltas, tab_baixas, tab_historico = st.tabs(["➕ Nova Falta", "🏥 Nova Baixa", "📜 Histórico"])
    
    with tab_faltas:
        with st.form("form_nova_falta"):
            st.markdown("### Registar Falta")
            
            col1, col2 = st.columns(2)
            with col1:
                data_inicio_falta = st.date_input("📅 Data Início", value=date.today(), key="falta_inicio")
            with col2:
                data_fim_falta = st.date_input("📅 Data Fim", value=date.today(), key="falta_fim")
            
            obs_falta = st.text_area("📝 Observações", key="obs_falta")
            
            submit_falta = st.form_submit_button("💾 REGISTAR FALTA", use_container_width=True, type="primary")
            
            if submit_falta:
                if data_inicio_falta > data_fim_falta:
                    st.error("❌ Data de início deve ser anterior à data de fim!")
                else:
                    with st.spinner("A registar..."):
                        if gravar_falta_baixa(emp_proc, ano_proc, mes_proc, colab_proc, 
                                            "Falta", data_inicio_falta, data_fim_falta, obs_falta):
                            st.success("✅ Falta registada!")
                            time.sleep(1)
                            st.rerun()
    
    with tab_baixas:
        with st.form("form_nova_baixa"):
            st.markdown("### Registar Baixa Médica")
            
            col1, col2 = st.columns(2)
            with col1:
                data_inicio_baixa = st.date_input("📅 Data Início", value=date.today(), key="baixa_inicio")
            with col2:
                data_fim_baixa = st.date_input("📅 Data Fim", value=date.today(), key="baixa_fim")
            
            obs_baixa = st.text_area("📝 Observações", key="obs_baixa")
            
            ficheiro_baixa = st.file_uploader("📎 Anexar Documento (PDF, imagem, etc.)", 
                                             type=['pdf', 'jpg', 'jpeg', 'png', 'doc', 'docx'],
                                             key="file_baixa")
            
            submit_baixa = st.form_submit_button("💾 REGISTAR BAIXA", use_container_width=True, type="primary")
            
            if submit_baixa:
                if data_inicio_baixa > data_fim_baixa:
                    st.error("❌ Data de início deve ser anterior à data de fim!")
                else:
                    with st.spinner("A processar..."):
                        ficheiro_path = None
                        
                        if ficheiro_baixa:
                            st.info("📤 A fazer upload do documento...")
                            ficheiro_path = upload_ficheiro_baixa(emp_proc, ano_proc, mes_proc, 
                                                                 colab_proc, ficheiro_baixa)
                            if ficheiro_path:
                                st.success(f"✅ Documento guardado: {ficheiro_path}")
                        
                        if gravar_falta_baixa(emp_proc, ano_proc, mes_proc, colab_proc,
                                            "Baixa", data_inicio_baixa, data_fim_baixa, 
                                            obs_baixa, ficheiro_path):
                            st.success("✅ Baixa registada!")
                            if ficheiro_path:
                                st.info(f"📁 Documento: {ficheiro_path}")
                            time.sleep(2)
                            st.rerun()
    
    with tab_historico:
        st.markdown("### 📜 Histórico do Mês")
        
        df_historico = carregar_faltas_baixas(emp_proc, ano_proc, mes_proc, colab_proc)
        
        if not df_historico.empty:
            st.dataframe(
                df_historico[['Tipo', 'Data Início', 'Data Fim', 'Dias Úteis', 'Dias Totais', 'Observações']],
                use_container_width=True,
                hide_index=True
            )
            
            # Totais
            total_faltas_uteis = df_historico[df_historico['Tipo'] == 'Falta']['Dias Úteis'].sum()
            total_baixas_uteis = df_historico[df_historico['Tipo'] == 'Baixa']['Dias Úteis'].sum()
            
            col1, col2, col3 = st.columns(3)
            col1.metric("🏖️ Total Faltas (dias úteis)", int(total_faltas_uteis))
            col2.metric("🏥 Total Baixas (dias úteis)", int(total_baixas_uteis))
            col3.metric("📊 Total Geral", int(total_faltas_uteis + total_baixas_uteis))
        else:
            st.info("ℹ️ Sem registos para este mês")
    
    st.markdown("---")
    
    # CALCULAR DIAS ÚTEIS TRABALHADOS
    df_historico = carregar_faltas_baixas(emp_proc, ano_proc, mes_proc, colab_proc)
    
    if not df_historico.empty:
        total_faltas_uteis = int(df_historico[df_historico['Tipo'] == 'Falta']['Dias Úteis'].sum())
        total_baixas_uteis = int(df_historico[df_historico['Tipo'] == 'Baixa']['Dias Úteis'].sum())
    else:
        total_faltas_uteis = 0
        total_baixas_uteis = 0
    
    dias_uteis_trab = max(dias_uteis_mes - total_faltas_uteis - total_baixas_uteis, 0)
    
    st.info(f"📊 Dias úteis trabalhados: {dias_uteis_trab} (de {dias_uteis_mes} dias úteis no mês)")
    
    st.markdown("---")
    
    # RESTO DO PROCESSAMENTO
    col1, col2, col3 = st.columns(3)
    with col1:
        desconto_especie = st.checkbox("☑️ Desconto em Espécie")
    with col2:
        sub_ferias = st.selectbox("🏖️ Sub. Férias", ["Duodécimos", "Total"])
    with col3:
        sub_natal = st.selectbox("🎄 Sub. Natal", ["Duodécimos", "Total"])
    
    st.markdown("---")
    
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
    
    dados_calc = {
        'salario_bruto': salario_bruto,
        'horas_semana': horas_semana,
        'subsidio_alimentacao': subsidio_alim,
        'dias_faltas': total_faltas_uteis,
        'dias_baixas': total_baixas_uteis,
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
    
    st.subheader("💵 Preview do Processamento")
    
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
        st.metric("Dias Úteis Mês", dias_uteis_mes)
        st.metric("Faltas (dias úteis)", total_faltas_uteis)
        st.metric("Baixas (dias úteis)", total_baixas_uteis)
        st.metric("Dias Úteis Trab.", dias_uteis_trab)
        st.metric("Dias Pagos", 30 - total_faltas_uteis - total_baixas_uteis)
        st.markdown("---")
        st.metric("**💰 LÍQUIDO**", f"**{resultado['liquido']:.2f}€**")

# ==================== RESCISÕES ====================

elif menu == "🚪 Rescisões":
    st.header("🚪 Gestão de Rescisões")
    st.info("📋 Registar rescisão e preparar envio para contabilidade")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        emp_resc = st.selectbox("Empresa", list(EMPRESAS.keys()), key="emp_resc")
    with col2:
        mes_resc = st.selectbox("Mês", list(range(1, 13)),
                               format_func=lambda x: calendar.month_name[x],
                               index=datetime.now().month - 1, key="mes_resc")
    with col3:
        ano_resc = st.selectbox("Ano", [2024, 2025, 2026], index=1, key="ano_resc")
    
    colabs_ativos = carregar_colaboradores_ativos(emp_resc)
    
    if not colabs_ativos:
        st.warning("⚠️ Nenhum colaborador ativo")
        st.stop()
    
    st.markdown("---")
    
    with st.form("form_rescisao"):
        st.markdown("### 📝 Dados da Rescisão")
        
        colab_resc = st.selectbox("👤 Colaborador", colabs_ativos, key="col_resc")
        
        col1, col2 = st.columns(2)
        
        with col1:
            data_rescisao = st.date_input("📅 Data de Rescisão", value=date.today())
            motivo_rescisao = st.selectbox("📋 Motivo da Rescisão", MOTIVOS_RESCISAO)
        
        with col2:
            enviar_contabilidade = st.checkbox("📤 Preparar para envio à contabilidade", value=True)
            observacoes = st.text_area("📝 Observações / Detalhes adicionais", height=100)
        
        st.markdown("---")
        
        submit_rescisao = st.form_submit_button("🚪 REGISTAR RESCISÃO", use_container_width=True, type="primary")
        
        if submit_rescisao:
            st.warning("⚠️ Esta ação é IRREVERSÍVEL!")
            
            if st.button("✅ CONFIRMAR RESCISÃO", type="primary"):
                with st.spinner("A processar rescisão..."):
                    if registar_rescisao(emp_resc, colab_resc, ano_resc, mes_resc, 
                                       data_rescisao, motivo_rescisao, observacoes):
                        st.success(f"✅ Rescisão de '{colab_resc}' registada!")
                        st.info(f"📅 Data: {data_rescisao.strftime('%d/%m/%Y')}")
                        st.info(f"📋 Motivo: {motivo_rescisao}")
                        
                        if enviar_contabilidade:
                            st.success("📤 Rescisão marcada para envio à contabilidade")
                            st.info("💡 Exporte os dados do colaborador para enviar ao contabilista")
                        
                        st.balloons()
                        time.sleep(3)
                        st.rerun()
    
    st.markdown("---")
    st.subheader("📊 Rescisões Registadas")
    
    # Mostrar colaboradores rescindidos
    df_base = carregar_dados_base(emp_resc)
    df_rescindidos = df_base[df_base['Status'] == 'Rescindido']
    
    if not df_rescindidos.empty:
        for _, row in df_rescindidos.iterrows():
            nome = row['Nome Completo']
            
            # Buscar dados de rescisão no snapshot
            snap = carregar_ultimo_snapshot(emp_resc, nome, ano_resc, mes_resc)
            
            if snap and snap.get('Status') == 'Rescindido':
                col1, col2 = st.columns([3, 1])
                
                with col1:
                    st.markdown(f"### 👤 {nome}")
                    st.write(f"📅 **Data Rescisão:** {snap.get('Data Rescisão', 'N/A')}")
                    st.write(f"📋 **Motivo:** {snap.get('Motivo Rescisão', 'N/A')}")
                    st.write(f"🏢 **Secção:** {row.get('Secção', 'N/A')}")
                
                with col2:
                    st.error("🚪 Rescindido")
                    if st.button("📄 Exportar Dados", key=f"exp_{nome}"):
                        st.info("💡 Funcionalidade de exportação em desenvolvimento")
                
                st.markdown("---")
    else:
        st.info("ℹ️ Sem rescisões registadas")

# ==================== TABELA IRS ====================

elif menu == "📊 Tabela IRS":
    st.header("📊 Gestão de Tabela IRS")
    
    uploaded = st.file_uploader("📤 Carregar Tabelas IRS (Excel)", type=['xlsx', 'xls'])
    
    if uploaded:
        xls = carregar_tabela_irs_excel(uploaded)
        
        if xls:
            st.markdown("---")
            aba_sel = st.selectbox("Selecione a aba", xls.sheet_names)
            df_preview = pd.read_excel(uploaded, sheet_name=aba_sel)
            st.dataframe(df_preview, use_container_width=True)
    
    if st.session_state.tabela_irs:
        st.success("✅ Tabela IRS carregada!")
    else:
        st.warning("⚠️ IRS será calculado com escalões aproximados")

st.sidebar.markdown("---")
st.sidebar.info(f"""v2.8 ✨ MELHORIAS
💶 SMN: {st.session_state.salario_minimo}€
✅ Faltas/baixas com datas
📤 Upload de documentos
🚪 Rescisões completo
🔄 Configurações sincronizadas""")

if st.sidebar.button("🚪 Logout", use_container_width=True):
    st.session_state.authenticated = False
    st.rerun()
