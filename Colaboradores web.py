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
    page_title="Processamento Salarial v2.1",
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

# Colunas do snapshot completo
COLUNAS_SNAPSHOT = [
    "Nome Completo",
    "Ano",
    "Mês",
    "Nº Horas/Semana",
    "Subsídio Alimentação Diário",
    "Número Pingo Doce",
    "Salário Bruto",
    "Vencimento Hora",
    "Status",
    "Data Rescisão",
    "Motivo Rescisão",
    "NIF",
    "NISS",
    "Data de Admissão",
    "IBAN",
    "Secção",
    "Timestamp"
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

def invalidar_cache_completo():
    """Limpa TODOS os caches"""
    keys_to_delete = [k for k in st.session_state.keys() 
                     if k not in ['authenticated', 'salario_minimo', 'feriados_municipais', 'password']]
    for key in keys_to_delete:
        del st.session_state[key]

# ==================== FUNÇÕES DE GESTÃO DE ABAS ====================

def garantir_aba(empresa, nome_aba, colunas):
    """Garante que uma aba existe no Excel"""
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
            st.success(f"✅ Aba '{nome_aba}' criada")
            return True
        return True
    except Exception as e:
        st.error(f"❌ Erro ao criar aba {nome_aba}: {e}")
        return False

def get_nome_aba_snapshot(ano, mes):
    """Retorna nome da aba de snapshot do mês"""
    return f"Estado_{ano}_{mes:02d}"

# ==================== SISTEMA DE SNAPSHOTS ====================

def carregar_dados_base(empresa):
    """Carrega dados base dos colaboradores (aba Colaboradores)"""
    try:
        _, response = dbx.files_download(EMPRESAS[empresa]["path"])
        df = pd.read_excel(BytesIO(response.content), sheet_name="Colaboradores")
        return df
    except Exception as e:
        st.error(f"❌ Erro ao carregar dados base: {e}")
        return pd.DataFrame()

def carregar_ultimo_snapshot(empresa, colaborador, ano=None, mes=None):
    """Carrega o último snapshot do colaborador (mais recente ou de mês específico)"""
    try:
        file_path = EMPRESAS[empresa]["path"]
        _, response = dbx.files_download(file_path)
        wb = load_workbook(BytesIO(response.content))
        
        # Se mês/ano especificados, buscar naquela aba
        if ano and mes:
            nome_aba = get_nome_aba_snapshot(ano, mes)
            if nome_aba in wb.sheetnames:
                df = pd.read_excel(BytesIO(response.content), sheet_name=nome_aba)
                df_colab = df[df['Nome Completo'] == colaborador]
                if not df_colab.empty:
                    return df_colab.iloc[-1].to_dict()  # Última linha = mais recente
        
        # Buscar snapshot mais recente em qualquer mês
        abas_estado = [s for s in wb.sheetnames if s.startswith('Estado_')]
        if not abas_estado:
            return None
        
        # Ordenar abas por data (mais recente primeiro)
        abas_estado.sort(reverse=True)
        
        for aba in abas_estado:
            df = pd.read_excel(BytesIO(response.content), sheet_name=aba)
            df_colab = df[df['Nome Completo'] == colaborador]
            if not df_colab.empty:
                return df_colab.iloc[-1].to_dict()
        
        return None
    except Exception as e:
        st.warning(f"⚠️ Aviso ao carregar snapshot: {e}")
        return None

def criar_snapshot_inicial(empresa, colaborador, ano, mes):
    """Cria snapshot inicial a partir dos dados base"""
    df_base = carregar_dados_base(empresa)
    dados_colab = df_base[df_base['Nome Completo'] == colaborador]
    
    if dados_colab.empty:
        st.error(f"❌ Colaborador {colaborador} não encontrado nos dados base")
        return None
    
    dados = dados_colab.iloc[0]
    horas_semana = float(dados.get('Nº Horas/Semana', 40))
    salario_bruto = calcular_salario_base(horas_semana, st.session_state.salario_minimo)
    
    snapshot = {
        "Nome Completo": colaborador,
        "Ano": ano,
        "Mês": mes,
        "Nº Horas/Semana": horas_semana,
        "Subsídio Alimentação Diário": float(dados.get('Subsídio Alimentação Diário', 5.96)),
        "Número Pingo Doce": str(dados.get('Número Pingo Doce', '')),
        "Salário Bruto": salario_bruto,
        "Vencimento Hora": calcular_vencimento_hora(salario_bruto, horas_semana),
        "Status": "Ativo",
        "Data Rescisão": None,
        "Motivo Rescisão": None,
        "NIF": dados.get('NIF', ''),
        "NISS": dados.get('NISS', ''),
        "Data de Admissão": dados.get('Data de Admissão', ''),
        "IBAN": dados.get('IBAN', ''),
        "Secção": dados.get('Secção', ''),
        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    
    return snapshot

def gravar_snapshot(empresa, dados_snapshot):
    """Grava snapshot completo do colaborador no mês"""
    try:
        ano = dados_snapshot['Ano']
        mes = dados_snapshot['Mês']
        nome_aba = get_nome_aba_snapshot(ano, mes)
        
        # Garantir que a aba existe
        garantir_aba(empresa, nome_aba, COLUNAS_SNAPSHOT)
        
        file_path = EMPRESAS[empresa]["path"]
        _, response = dbx.files_download(file_path)
        wb = load_workbook(BytesIO(response.content))
        ws = wb[nome_aba]
        
        # Adicionar nova linha com snapshot completo
        nova_linha = [dados_snapshot.get(col, '') for col in COLUNAS_SNAPSHOT]
        ws.append(nova_linha)
        
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        dbx.files_upload(output.read(), file_path, mode=dropbox.files.WriteMode.overwrite)
        
        return True
    except Exception as e:
        st.error(f"❌ Erro ao gravar snapshot: {e}")
        import traceback
        st.error(traceback.format_exc())
        return False

def carregar_estado_colaborador(empresa, colaborador, ano, mes):
    """
    Carrega estado completo do colaborador no mês.
    Se não existir snapshot, cria um inicial dos dados base.
    """
    # Tentar carregar snapshot do mês específico
    snapshot = carregar_ultimo_snapshot(empresa, colaborador, ano, mes)
    
    if snapshot:
        return snapshot
    
    # Se não existe snapshot do mês, buscar o mais recente de qualquer mês
    snapshot_anterior = carregar_ultimo_snapshot(empresa, colaborador)
    
    if snapshot_anterior:
        # Atualizar mês/ano e timestamp
        snapshot_anterior['Ano'] = ano
        snapshot_anterior['Mês'] = mes
        snapshot_anterior['Timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # Gravar novo snapshot para o mês atual
        gravar_snapshot(empresa, snapshot_anterior)
        return snapshot_anterior
    
    # Se não existe nenhum snapshot, criar inicial
    snapshot_inicial = criar_snapshot_inicial(empresa, colaborador, ano, mes)
    if snapshot_inicial:
        gravar_snapshot(empresa, snapshot_inicial)
        return snapshot_inicial
    
    return None

def atualizar_campo_colaborador(empresa, colaborador, ano, mes, campo, novo_valor):
    """
    Atualiza um campo específico do colaborador.
    Carrega último estado, atualiza campo e grava novo snapshot.
    """
    # Carregar estado atual
    estado = carregar_estado_colaborador(empresa, colaborador, ano, mes)
    
    if not estado:
        st.error(f"❌ Não foi possível carregar estado do colaborador")
        return False
    
    # Atualizar campo
    estado[campo] = novo_valor
    estado['Ano'] = ano
    estado['Mês'] = mes
    estado['Timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Recalcular campos dependentes
    if campo == "Nº Horas/Semana":
        horas = float(novo_valor)
        estado['Salário Bruto'] = calcular_salario_base(horas, st.session_state.salario_minimo)
        estado['Vencimento Hora'] = calcular_vencimento_hora(estado['Salário Bruto'], horas)
    
    # Gravar novo snapshot
    return gravar_snapshot(empresa, estado)

def carregar_colaboradores_ativos(empresa, ano, mes):
    """Carrega todos os colaboradores ativos no mês"""
    df_base = carregar_dados_base(empresa)
    colaboradores_ativos = []
    
    for _, colab in df_base.iterrows():
        nome = colab['Nome Completo']
        estado = carregar_estado_colaborador(empresa, nome, ano, mes)
        
        if estado and estado.get('Status') == 'Ativo':
            colaboradores_ativos.append(estado)
    
    if colaboradores_ativos:
        return pd.DataFrame(colaboradores_ativos)
    return pd.DataFrame()

# ==================== FUNÇÕES DE CÁLCULO ====================

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
    salario_bruto = dados_form['salario_bruto']
    horas_semana = dados_form['horas_semana']
    sub_alimentacao_dia = dados_form['subsidio_alimentacao']
    vencimento_hora = calcular_vencimento_hora(salario_bruto, horas_semana)
    
    dias_trabalhados = dados_form['dias_trabalhados']
    dias_uteis_trabalhados = dados_form['dias_uteis_trabalhados']
    
    horas_noturnas = dados_form.get('horas_noturnas', 0)
    horas_domingos = dados_form.get('horas_domingos', 0)
    horas_feriados = dados_form.get('horas_feriados', 0)
    horas_extra = dados_form.get('horas_extra', 0)
    
    # REMUNERAÇÕES
    vencimento_ajustado = (salario_bruto / 30) * dias_trabalhados
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
    irs = base_ss * 0.10  # Placeholder
    desconto_especie = sub_alimentacao if dados_form.get('desconto_especie', False) else 0
    total_descontos = seg_social + irs + desconto_especie
    
    # LÍQUIDO
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
        'irs': irs,
        'desconto_especie': desconto_especie,
        'total_descontos': total_descontos,
        'liquido': liquido
    }

# ==================== FUNÇÕES AUXILIARES ====================

def registar_rescisao(empresa, colaborador, ano, mes, data_rescisao, motivo, obs, dias_aviso):
    """Registra rescisão atualizando snapshot do colaborador"""
    estado = carregar_estado_colaborador(empresa, colaborador, ano, mes)
    
    if not estado:
        return False
    
    estado['Status'] = 'Rescindido'
    estado['Data Rescisão'] = data_rescisao.strftime("%Y-%m-%d")
    estado['Motivo Rescisão'] = f"{motivo} | Dias aviso: {dias_aviso} | Obs: {obs}"
    estado['Timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    return gravar_snapshot(empresa, estado)

# ==================== INTERFACE ====================

if not check_password():
    st.stop()

st.title("💰 Processamento Salarial v2.1")
st.caption("🔄 Sistema com snapshots mensais - dados sempre atualizados")
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
        
        col1, col2, col3 = st.columns(3)
        with col1:
            empresa_sel = st.selectbox("Empresa", list(EMPRESAS.keys()), key="emp_config")
        with col2:
            mes_config = st.selectbox("Mês", list(range(1, 13)), 
                                     format_func=lambda x: calendar.month_name[x],
                                     index=datetime.now().month - 1, key="mes_config")
        with col3:
            ano_config = st.selectbox("Ano", [2024, 2025, 2026], index=1, key="ano_config")
        
        df_ativos = carregar_colaboradores_ativos(empresa_sel, ano_config, mes_config)
        
        if not df_ativos.empty:
            colaborador_sel = st.selectbox(
                "Colaborador",
                options=df_ativos['Nome Completo'].tolist(),
                key="colab_config"
            )
            
            estado_atual = carregar_estado_colaborador(empresa_sel, colaborador_sel, ano_config, mes_config)
            
            if estado_atual:
                st.markdown("---")
                col1, col2, col3 = st.columns(3)
                col1.metric("📊 Subsídio Atual", f"{estado_atual['Subsídio Alimentação Diário']:.2f}€")
                col2.metric("⏰ Horas/Semana", f"{estado_atual['Nº Horas/Semana']:.0f}h")
                col3.metric("🔢 Nº Pingo Doce", estado_atual.get('Número Pingo Doce', 'N/A'))
                
                with st.form("form_editar"):
                    st.markdown(f"### ✏️ Editar: {colaborador_sel} ({calendar.month_name[mes_config]}/{ano_config})")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        novo_sub = st.number_input(
                            "Subsídio Alimentação Diário (€)",
                            min_value=0.0,
                            value=float(estado_atual['Subsídio Alimentação Diário']),
                            step=0.10,
                            format="%.2f"
                        )
                    
                    with col2:
                        novo_num_pingo = st.text_input(
                            "Número Pingo Doce",
                            value=str(estado_atual.get('Número Pingo Doce', ''))
                        )
                    
                    submit = st.form_submit_button("💾 Guardar Alterações", use_container_width=True)
                    
                    if submit:
                        with st.spinner("🔄 A guardar snapshot atualizado..."):
                            sucesso1 = atualizar_campo_colaborador(empresa_sel, colaborador_sel, ano_config, mes_config, 
                                                                   "Subsídio Alimentação Diário", novo_sub)
                            sucesso2 = atualizar_campo_colaborador(empresa_sel, colaborador_sel, ano_config, mes_config,
                                                                   "Número Pingo Doce", novo_num_pingo)
                            
                            if sucesso1 and sucesso2:
                                invalidar_cache_completo()
                                st.success(f"✅ Snapshot atualizado! Dados gravados para {calendar.month_name[mes_config]}/{ano_config}")
                                st.info("ℹ️ Esta alteração estará disponível em todos os módulos imediatamente")
                                st.balloons()
                                time.sleep(2)
                                st.rerun()
        else:
            st.warning("⚠️ Nenhum colaborador ativo encontrado")
    
    # TAB 3: HORÁRIOS
    with tab3:
        st.subheader("⏰ Mudanças de Horário")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            empresa_h = st.selectbox("Empresa", list(EMPRESAS.keys()), key="emp_horas")
        with col2:
            mes_h = st.selectbox("Mês Vigência", list(range(1, 13)),
                                format_func=lambda x: calendar.month_name[x],
                                index=datetime.now().month - 1, key="mes_horas")
        with col3:
            ano_h = st.selectbox("Ano Vigência", [2024, 2025, 2026], index=1, key="ano_horas")
        
        df_ativos_h = carregar_colaboradores_ativos(empresa_h, ano_h, mes_h)
        
        if not df_ativos_h.empty:
            with st.form("form_horas"):
                colaborador_h = st.selectbox("Colaborador", df_ativos_h['Nome Completo'].tolist())
                
                estado_h = carregar_estado_colaborador(empresa_h, colaborador_h, ano_h, mes_h)
                horas_atuais = estado_h['Nº Horas/Semana'] if estado_h else 40
                
                st.info(f"⏰ Horário atual: **{horas_atuais:.0f}h/semana**")
                
                novas_horas = st.selectbox("Novo Horário (h/semana)", [16, 20, 40], index=2)
                
                submit_h = st.form_submit_button("💾 Registar Mudança", use_container_width=True)
                
                if submit_h:
                    with st.spinner("🔄 A atualizar snapshot..."):
                        sucesso = atualizar_campo_colaborador(empresa_h, colaborador_h, ano_h, mes_h,
                                                             "Nº Horas/Semana", float(novas_horas))
                        if sucesso:
                            invalidar_cache_completo()
                            st.success(f"✅ Horário atualizado: {horas_atuais:.0f}h → {novas_horas}h (vigência {calendar.month_name[mes_h]}/{ano_h})")
                            st.balloons()
                            time.sleep(2)
                            st.rerun()

# ==================== PROCESSAR SALÁRIOS ====================

elif menu == "💼 Processar Salários":
    st.header("💼 Processamento Mensal de Salários")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        empresa_proc = st.selectbox("🏢 Empresa", list(EMPRESAS.keys()), key="emp_proc")
    with col2:
        mes_proc = st.selectbox("📅 Mês", list(range(1, 13)),
                               format_func=lambda x: calendar.month_name[x],
                               index=datetime.now().month - 1, key="mes_proc")
    with col3:
        ano_proc = st.selectbox("📆 Ano", [2024, 2025, 2026], index=1, key="ano_proc")
    
    df_ativos_proc = carregar_colaboradores_ativos(empresa_proc, ano_proc, mes_proc)
    
    if df_ativos_proc.empty:
        st.warning("⚠️ Nenhum colaborador ativo encontrado")
        st.stop()
    
    colaborador_proc = st.selectbox("👤 Colaborador", df_ativos_proc['Nome Completo'].tolist(), key="colab_proc")
    
    # Carregar estado completo do colaborador no mês
    estado = carregar_estado_colaborador(empresa_proc, colaborador_proc, ano_proc, mes_proc)
    
    if not estado:
        st.error("❌ Erro ao carregar dados do colaborador")
        st.stop()
    
    salario_bruto = float(estado['Salário Bruto'])
    horas_semana = float(estado['Nº Horas/Semana'])
    subsidio_alim = float(estado['Subsídio Alimentação Diário'])
    vencimento_hora = float(estado['Vencimento Hora'])
    numero_pingo = estado.get('Número Pingo Doce', '')
    
    feriados_completos = FERIADOS_NACIONAIS_2025 + st.session_state.feriados_municipais
    dias_uteis_mes = calcular_dias_uteis(ano_proc, mes_proc, feriados_completos)
    
    st.markdown("---")
    
    # DADOS BASE
    with st.expander("📋 **DADOS BASE DO COLABORADOR**", expanded=True):
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("💶 Salário Bruto", f"{salario_bruto:.2f}€")
        col2.metric("⏰ Horas/Semana", f"{horas_semana:.0f}h")
        col3.metric("💵 Vencimento/Hora", f"{vencimento_hora:.2f}€")
        col4.metric("🍽️ Sub. Alimentação", f"{subsidio_alim:.2f}€/dia")
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("📅 Dias Úteis Mês", dias_uteis_mes)
        col2.metric("🔢 NIF", estado.get('NIF', 'N/A'))
        col3.metric("🔢 NISS", estado.get('NISS', 'N/A'))
        if numero_pingo:
            col4.metric("🔢 Nº Pingo Doce", numero_pingo)
        
        st.caption(f"📸 Snapshot carregado: {estado.get('Timestamp', 'N/A')}")
    
    st.markdown("---")
    
    # OPÇÕES
    st.subheader("⚙️ Opções de Processamento")
    col1, col2, col3 = st.columns(3)
    with col1:
        desconto_especie = st.checkbox("☑️ Desconto em Espécie", value=False)
    with col2:
        sub_ferias_tipo = st.selectbox("🏖️ Subsídio Férias", ["Duodécimos", "Total"])
    with col3:
        sub_natal_tipo = st.selectbox("🎄 Subsídio Natal", ["Duodécimos", "Total"])
    
    st.markdown("---")
    
    # AUSÊNCIAS (simplificado para o exemplo)
    st.subheader("🏖️ Faltas e Baixas")
    col1, col2 = st.columns(2)
    with col1:
        total_dias_faltas = st.number_input("Total Dias Faltas", min_value=0, value=0)
    with col2:
        total_dias_baixas = st.number_input("Total Dias Baixas", min_value=0, value=0)
    
    st.markdown("---")
    
    # HORAS EXTRAS
    st.subheader("⏰ Horas Extras")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        horas_noturnas = st.number_input("🌙 Noturnas", min_value=0.0, value=0.0, step=0.5)
    with col2:
        horas_domingos = st.number_input("📅 Domingos", min_value=0.0, value=0.0, step=0.5)
    with col3:
        horas_feriados = st.number_input("🎉 Feriados", min_value=0.0, value=0.0, step=0.5)
    with col4:
        horas_extra = st.number_input("⚡ Extra", min_value=0.0, value=0.0, step=0.5)
    
    st.markdown("---")
    
    # OUTROS
    outros_proveitos = st.number_input("💰 Outros Proveitos c/ Descontos (€)", min_value=0.0, value=0.0, step=10.0)
    
    st.markdown("---")
    
    # CALCULAR
    data_admissao = pd.to_datetime(estado.get('Data de Admissão', date.today())).date()
    dias_trabalhados = calcular_dias_trabalhados_com_admissao(mes_proc, ano_proc, data_admissao, 
                                                              total_dias_faltas, total_dias_baixas)
    dias_uteis_trabalhados = max(dias_uteis_mes - total_dias_faltas - total_dias_baixas, 0)
    
    dados_calculo = {
        'salario_bruto': salario_bruto,
        'horas_semana': horas_semana,
        'subsidio_alimentacao': subsidio_alim,
        'dias_uteis_mes': dias_uteis_mes,
        'dias_trabalhados': dias_trabalhados,
        'dias_uteis_trabalhados': dias_uteis_trabalhados,
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
    
    # PREVIEW
    st.subheader("💵 Preview dos Cálculos")
    
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
        st.metric("**TOTAL**", f"**{resultado['total_remuneracoes']:.2f}€**")
    
    with col2:
        st.markdown("### 📉 Descontos")
        st.metric("Base SS/IRS", f"{resultado['base_ss']:.2f}€")
        st.metric("Seg. Social (11%)", f"{resultado['seg_social']:.2f}€")
        st.metric("IRS", f"{resultado['irs']:.2f}€")
        if desconto_especie:
            st.metric("Desconto Espécie", f"{resultado['desconto_especie']:.2f}€")
        st.markdown("---")
        st.metric("**TOTAL**", f"**{resultado['total_descontos']:.2f}€**")
    
    with col3:
        st.markdown("### 💵 Resumo")
        st.metric("Dias Trabalhados", dias_trabalhados)
        st.metric("Dias Úteis Trab.", dias_uteis_trabalhados)
        st.markdown("---")
        st.metric("**💰 LÍQUIDO**", f"**{resultado['liquido']:.2f}€**")
    
    st.markdown("---")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("💾 Guardar", use_container_width=True):
            st.info("🚧 Em desenvolvimento...")
    with col2:
        if st.button("📄 Gerar PDF", use_container_width=True):
            st.info("🚧 Em desenvolvimento...")
    with col3:
        if st.button("📊 Export Excel", use_container_width=True):
            st.info("🚧 Em desenvolvimento...")

# ==================== RESCISÕES ====================

elif menu == "🚪 Rescisões":
    st.header("🚪 Gestão de Rescisões")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        empresa_resc = st.selectbox("Empresa", list(EMPRESAS.keys()), key="emp_resc")
    with col2:
        mes_resc = st.selectbox("Mês", list(range(1, 13)),
                               format_func=lambda x: calendar.month_name[x],
                               index=datetime.now().month - 1, key="mes_resc")
    with col3:
        ano_resc = st.selectbox("Ano", [2024, 2025, 2026], index=1, key="ano_resc")
    
    df_ativos_resc = carregar_colaboradores_ativos(empresa_resc, ano_resc, mes_resc)
    
    if df_ativos_resc.empty:
        st.warning("⚠️ Nenhum colaborador ativo")
    else:
        with st.form("form_resc"):
            colaborador_resc = st.selectbox("Colaborador", df_ativos_resc['Nome Completo'].tolist())
            
            col1, col2 = st.columns(2)
            with col1:
                data_rescisao = st.date_input("Data Rescisão", value=date.today())
            with col2:
                dias_aviso = st.number_input("Dias Aviso Prévio", min_value=0, value=0)
            
            motivo = st.selectbox("Motivo", MOTIVOS_RESCISAO)
            obs = st.text_area("Observações", height=100)
            
            submit = st.form_submit_button("💾 Registar Rescisão", use_container_width=True)
            
            if submit:
                with st.spinner("🔄 A registar..."):
                    sucesso = registar_rescisao(empresa_resc, colaborador_resc, ano_resc, mes_resc,
                                              data_rescisao, motivo, obs, dias_aviso)
                    if sucesso:
                        invalidar_cache_completo()
                        st.success(f"✅ Rescisão registada! {colaborador_resc} não aparecerá mais nos próximos meses")
                        time.sleep(2)
                        st.rerun()

# ==================== RELATÓRIOS ====================

elif menu == "📊 Relatórios":
    st.header("📊 Relatórios")
    st.info("🚧 Módulo em desenvolvimento...")
    
    st.markdown("### 🔍 Debug: Ver Snapshots")
    
    col1, col2 = st.columns(2)
    with col1:
        empresa_debug = st.selectbox("Empresa", list(EMPRESAS.keys()), key="emp_debug")
    with col2:
        try:
            _, response = dbx.files_download(EMPRESAS[empresa_debug]["path"])
            wb = load_workbook(BytesIO(response.content))
            abas_estado = [s for s in wb.sheetnames if s.startswith('Estado_')]
            if abas_estado:
                aba_sel = st.selectbox("Aba Estado", sorted(abas_estado, reverse=True))
                df_debug = pd.read_excel(BytesIO(response.content), sheet_name=aba_sel)
                st.dataframe(df_debug, use_container_width=True)
            else:
                st.info("📭 Nenhuma aba de estado encontrada ainda")
        except Exception as e:
            st.error(f"Erro: {e}")

# SIDEBAR
st.sidebar.markdown("---")
st.sidebar.info(f"👤 Sistema: v2.1 (Snapshots)\n💶 SMN: {st.session_state.salario_minimo}€")

if st.sidebar.button("🚪 Logout", use_container_width=True):
    st.session_state.authenticated = False
    st.rerun()
