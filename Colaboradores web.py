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
    page_title="Processamento Salarial v2.2",
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

COLUNAS_SNAPSHOT = [
    "Nome Completo", "Ano", "Mês", "Nº Horas/Semana", "Subsídio Alimentação Diário",
    "Número Pingo Doce", "Salário Bruto", "Vencimento Hora", "Status",
    "Data Rescisão", "Motivo Rescisão", "NIF", "NISS", "Data de Admissão",
    "IBAN", "Secção", "Timestamp"
]

# ==================== SESSION STATE ====================

if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'salario_minimo' not in st.session_state:
    st.session_state.salario_minimo = 870.0
if 'feriados_municipais' not in st.session_state:
    st.session_state.feriados_municipais = [date(2025, 1, 14)]
if 'ultimo_reload' not in st.session_state:
    st.session_state.ultimo_reload = datetime.now()

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

# ==================== FUNÇÕES DROPBOX (SEM CACHE) ====================

def get_nome_aba_snapshot(ano, mes):
    """Retorna nome da aba de snapshot do mês"""
    return f"Estado_{ano}_{mes:02d}"

def download_excel(empresa):
    """Download do Excel direto do Dropbox - SEM CACHE"""
    try:
        file_path = EMPRESAS[empresa]["path"]
        _, response = dbx.files_download(file_path)
        return BytesIO(response.content)
    except Exception as e:
        st.error(f"❌ Erro ao baixar Excel: {e}")
        return None

def garantir_aba(wb, nome_aba, colunas):
    """Garante que uma aba existe no workbook"""
    if nome_aba not in wb.sheetnames:
        ws = wb.create_sheet(nome_aba)
        ws.append(colunas)
        return True
    return False

def upload_excel(empresa, wb):
    """Upload do Excel para Dropbox"""
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

# ==================== FUNÇÕES DE DADOS BASE ====================

def carregar_dados_base(empresa):
    """Carrega aba Colaboradores - dados base originais"""
    excel_file = download_excel(empresa)
    if excel_file:
        try:
            df = pd.read_excel(excel_file, sheet_name="Colaboradores")
            return df
        except Exception as e:
            st.error(f"❌ Erro ao ler aba Colaboradores: {e}")
    return pd.DataFrame()

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
    if horas_semana == 0:
        return 0
    return (salario_bruto * 12) / (52 * horas_semana)

# ==================== SISTEMA DE SNAPSHOTS ====================

def criar_snapshot_inicial(empresa, colaborador, ano, mes):
    """Cria snapshot inicial a partir dos dados base"""
    df_base = carregar_dados_base(empresa)
    dados_colab = df_base[df_base['Nome Completo'] == colaborador]
    
    if dados_colab.empty:
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
    """
    Carrega o último snapshot do colaborador no mês.
    SEMPRE vai à Dropbox - SEM CACHE!
    """
    excel_file = download_excel(empresa)
    if not excel_file:
        return None
    
    try:
        wb = load_workbook(excel_file)
        nome_aba = get_nome_aba_snapshot(ano, mes)
        
        # Tentar carregar da aba do mês específico
        if nome_aba in wb.sheetnames:
            df = pd.read_excel(excel_file, sheet_name=nome_aba)
            df_colab = df[df['Nome Completo'] == colaborador]
            
            if not df_colab.empty:
                # ÚLTIMA linha = snapshot mais recente
                snapshot = df_colab.iloc[-1].to_dict()
                st.caption(f"📸 Snapshot carregado: {snapshot.get('Timestamp', 'N/A')} (Aba: {nome_aba})")
                return snapshot
        
        # Se não existe no mês atual, buscar em meses anteriores
        abas_estado = sorted([s for s in wb.sheetnames if s.startswith('Estado_')], reverse=True)
        
        for aba in abas_estado:
            try:
                df = pd.read_excel(excel_file, sheet_name=aba)
                df_colab = df[df['Nome Completo'] == colaborador]
                
                if not df_colab.empty:
                    snapshot = df_colab.iloc[-1].to_dict()
                    # Atualizar para mês atual
                    snapshot['Ano'] = ano
                    snapshot['Mês'] = mes
                    snapshot['Timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    st.caption(f"📸 Snapshot herdado de {aba}: {snapshot.get('Timestamp', 'N/A')}")
                    return snapshot
            except:
                continue
        
        # Se não encontrou nenhum snapshot, criar inicial
        st.warning(f"⚠️ Nenhum snapshot encontrado. Criando inicial...")
        return criar_snapshot_inicial(empresa, colaborador, ano, mes)
        
    except Exception as e:
        st.error(f"❌ Erro ao carregar snapshot: {e}")
        import traceback
        st.error(traceback.format_exc())
        return None

def gravar_snapshot(empresa, snapshot):
    """
    Grava snapshot DIRETO no Dropbox.
    Adiciona NOVA LINHA na aba Estado_YYYY_MM.
    """
    try:
        ano = snapshot['Ano']
        mes = snapshot['Mês']
        nome_aba = get_nome_aba_snapshot(ano, mes)
        
        # Baixar Excel atual
        excel_file = download_excel(empresa)
        if not excel_file:
            return False
        
        wb = load_workbook(excel_file)
        
        # Garantir que aba existe
        aba_criada = garantir_aba(wb, nome_aba, COLUNAS_SNAPSHOT)
        if aba_criada:
            st.info(f"✨ Aba '{nome_aba}' criada")
        
        ws = wb[nome_aba]
        
        # Preparar linha com snapshot
        nova_linha = []
        for col in COLUNAS_SNAPSHOT:
            valor = snapshot.get(col, '')
            # Converter tipos para garantir compatibilidade
            if isinstance(valor, (int, float)):
                nova_linha.append(valor)
            else:
                nova_linha.append(str(valor) if valor else '')
        
        # Adicionar linha
        ws.append(nova_linha)
        
        # Upload de volta
        sucesso = upload_excel(empresa, wb)
        
        if sucesso:
            linha = ws.max_row
            st.success(f"✅ Snapshot gravado na linha {linha} da aba '{nome_aba}'")
            return True
        
        return False
        
    except Exception as e:
        st.error(f"❌ Erro ao gravar snapshot: {e}")
        import traceback
        st.error(traceback.format_exc())
        return False

def atualizar_campo_colaborador(empresa, colaborador, ano, mes, campo, novo_valor):
    """
    Atualiza um campo do colaborador.
    1. Carrega último snapshot
    2. Atualiza campo
    3. Recalcula dependências
    4. Grava novo snapshot
    """
    # Carregar estado atual - DIRETO da Dropbox
    snapshot = carregar_ultimo_snapshot(empresa, colaborador, ano, mes)
    
    if not snapshot:
        st.error(f"❌ Não foi possível carregar snapshot de {colaborador}")
        return False
    
    # Atualizar campo
    snapshot[campo] = novo_valor
    snapshot['Ano'] = ano
    snapshot['Mês'] = mes
    snapshot['Timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Recalcular campos dependentes
    if campo == "Nº Horas/Semana":
        horas = float(novo_valor)
        snapshot['Salário Bruto'] = calcular_salario_base(horas, st.session_state.salario_minimo)
        snapshot['Vencimento Hora'] = calcular_vencimento_hora(snapshot['Salário Bruto'], horas)
    
    # Gravar - DIRETO na Dropbox
    return gravar_snapshot(empresa, snapshot)

def carregar_colaboradores_ativos(empresa, ano, mes):
    """Carrega lista de colaboradores ativos no mês"""
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

def calcular_dias_trabalhados_com_admissao(mes, ano, data_admissao, total_faltas, total_baixas):
    """Calcula dias trabalhados considerando data de admissão"""
    dias_no_mes = calendar.monthrange(ano, mes)[1]
    
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
    irs = base_ss * 0.10
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

def registar_rescisao(empresa, colaborador, ano, mes, data_rescisao, motivo, obs, dias_aviso):
    """Registra rescisão atualizando snapshot"""
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

st.title("💰 Processamento Salarial v2.2")
st.caption("🔄 Sistema SEM CACHE - dados sempre da Dropbox")
st.caption(f"🕐 Último reload: {st.session_state.ultimo_reload.strftime('%H:%M:%S')}")
st.markdown("---")

menu = st.sidebar.radio(
    "Menu Principal",
    ["⚙️ Configurações", "💼 Processar Salários", "🚪 Rescisões"],
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
                st.success(f"✅ {len(feriados_temp)} feriados configurados")
    
    # TAB 2: COLABORADORES
    with tab2:
        st.subheader("👥 Editar Dados de Colaboradores")
        st.warning("⚠️ ATENÇÃO: Ao clicar Guardar, aguarde a confirmação antes de fazer qualquer outra ação!")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            empresa_sel = st.selectbox("Empresa", list(EMPRESAS.keys()), key="emp_config")
        with col2:
            mes_config = st.selectbox("Mês", list(range(1, 13)), 
                                     format_func=lambda x: calendar.month_name[x],
                                     index=datetime.now().month - 1, key="mes_config")
        with col3:
            ano_config = st.selectbox("Ano", [2024, 2025, 2026], index=1, key="ano_config")
        
        st.info(f"📅 A trabalhar com: {calendar.month_name[mes_config]}/{ano_config}")
        
        colaboradores_ativos = carregar_colaboradores_ativos(empresa_sel, ano_config, mes_config)
        
        if not colaboradores_ativos:
            st.warning("⚠️ Nenhum colaborador ativo encontrado")
        else:
            colaborador_sel = st.selectbox(
                "Colaborador",
                options=colaboradores_ativos,
                key="colab_config"
            )
            
            # CARREGAR SNAPSHOT - SEM CACHE!
            with st.spinner("🔄 A carregar dados da Dropbox..."):
                snapshot = carregar_ultimo_snapshot(empresa_sel, colaborador_sel, ano_config, mes_config)
            
            if snapshot:
                st.markdown("---")
                st.markdown("### 📊 Dados Atuais")
                
                col1, col2, col3 = st.columns(3)
                col1.metric("💰 Subsídio Alimentação", f"{snapshot['Subsídio Alimentação Diário']:.2f}€")
                col2.metric("⏰ Horas/Semana", f"{snapshot['Nº Horas/Semana']:.0f}h")
                col3.metric("🔢 Nº Pingo Doce", snapshot.get('Número Pingo Doce', 'N/A'))
                
                st.markdown("---")
                
                with st.form("form_editar", clear_on_submit=False):
                    st.markdown(f"### ✏️ Editar: {colaborador_sel}")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        novo_sub = st.number_input(
                            "Novo Subsídio Alimentação Diário (€)",
                            min_value=0.0,
                            value=float(snapshot['Subsídio Alimentação Diário']),
                            step=0.10,
                            format="%.2f",
                            key="novo_sub_input"
                        )
                    
                    with col2:
                        novo_num_pingo = st.text_input(
                            "Novo Número Pingo Doce",
                            value=str(snapshot.get('Número Pingo Doce', '')),
                            key="novo_num_input"
                        )
                    
                    submit = st.form_submit_button("💾 GUARDAR ALTERAÇÕES", use_container_width=True, type="primary")
                    
                    if submit:
                        st.markdown("---")
                        st.warning("⏳ A GUARDAR... NÃO FECHE OU NAVEGUE!")
                        
                        with st.spinner("🔄 Passo 1/3: Atualizando subsídio..."):
                            sucesso1 = atualizar_campo_colaborador(
                                empresa_sel, colaborador_sel, ano_config, mes_config,
                                "Subsídio Alimentação Diário", novo_sub
                            )
                            time.sleep(1)
                        
                        with st.spinner("🔄 Passo 2/3: Atualizando número Pingo Doce..."):
                            sucesso2 = atualizar_campo_colaborador(
                                empresa_sel, colaborador_sel, ano_config, mes_config,
                                "Número Pingo Doce", novo_num_pingo
                            )
                            time.sleep(1)
                        
                        if sucesso1 and sucesso2:
                            st.success("✅ DADOS GRAVADOS COM SUCESSO!")
                            st.info("🔄 A recarregar página em 3 segundos...")
                            st.balloons()
                            time.sleep(3)
                            st.session_state.ultimo_reload = datetime.now()
                            st.rerun()
                        else:
                            st.error("❌ Erro ao gravar. Verifique as mensagens acima.")

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
        
        colaboradores_h = carregar_colaboradores_ativos(empresa_h, ano_h, mes_h)
        
        if not colaboradores_h:
            st.warning("⚠️ Nenhum colaborador ativo")
        else:
            with st.form("form_horas"):
                colaborador_h = st.selectbox("Colaborador", colaboradores_h)
                
                snapshot_h = carregar_ultimo_snapshot(empresa_h, colaborador_h, ano_h, mes_h)
                horas_atuais = snapshot_h['Nº Horas/Semana'] if snapshot_h else 40
                
                st.info(f"⏰ Horário atual: **{horas_atuais:.0f}h/semana**")
                
                novas_horas = st.selectbox("Novo Horário (h/semana)", [16, 20, 40], index=2)
                
                submit_h = st.form_submit_button("💾 REGISTAR MUDANÇA", use_container_width=True, type="primary")
                
                if submit_h:
                    with st.spinner("🔄 A gravar mudança de horário..."):
                        sucesso = atualizar_campo_colaborador(
                            empresa_h, colaborador_h, ano_h, mes_h,
                            "Nº Horas/Semana", float(novas_horas)
                        )
                        
                        if sucesso:
                            st.success(f"✅ Horário atualizado: {horas_atuais:.0f}h → {novas_horas}h")
                            st.balloons()
                            time.sleep(3)
                            st.session_state.ultimo_reload = datetime.now()
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
    
    with st.spinner("🔄 A carregar colaboradores ativos..."):
        colaboradores_proc = carregar_colaboradores_ativos(empresa_proc, ano_proc, mes_proc)
    
    if not colaboradores_proc:
        st.warning("⚠️ Nenhum colaborador ativo encontrado")
        st.stop()
    
    colaborador_proc = st.selectbox("👤 Colaborador", colaboradores_proc, key="colab_proc")
    
    # CARREGAR SNAPSHOT - SEM CACHE!
    with st.spinner("🔄 A carregar dados do colaborador da Dropbox..."):
        snapshot_proc = carregar_ultimo_snapshot(empresa_proc, colaborador_proc, ano_proc, mes_proc)
    
    if not snapshot_proc:
        st.error("❌ Erro ao carregar dados do colaborador")
        st.stop()
    
    salario_bruto = float(snapshot_proc['Salário Bruto'])
    horas_semana = float(snapshot_proc['Nº Horas/Semana'])
    subsidio_alim = float(snapshot_proc['Subsídio Alimentação Diário'])
    vencimento_hora = float(snapshot_proc['Vencimento Hora'])
    numero_pingo = snapshot_proc.get('Número Pingo Doce', '')
    
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
        col2.metric("🔢 NIF", snapshot_proc.get('NIF', 'N/A'))
        col3.metric("🔢 NISS", snapshot_proc.get('NISS', 'N/A'))
        if numero_pingo:
            col4.metric("🔢 Nº Pingo Doce", numero_pingo)
    
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
    
    # AUSÊNCIAS (simplificado)
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
    
    outros_proveitos = st.number_input("💰 Outros Proveitos c/ Descontos (€)", min_value=0.0, value=0.0)
    
    st.markdown("---")
    
    # CALCULAR
    data_admissao = pd.to_datetime(snapshot_proc.get('Data de Admissão', date.today())).date()
    dias_trabalhados = calcular_dias_trabalhados_com_admissao(
        mes_proc, ano_proc, data_admissao, total_dias_faltas, total_dias_baixas
    )
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
    
    colaboradores_resc = carregar_colaboradores_ativos(empresa_resc, ano_resc, mes_resc)
    
    if not colaboradores_resc:
        st.warning("⚠️ Nenhum colaborador ativo")
    else:
        with st.form("form_resc"):
            colaborador_resc = st.selectbox("Colaborador", colaboradores_resc)
            
            col1, col2 = st.columns(2)
            with col1:
                data_rescisao = st.date_input("Data Rescisão", value=date.today())
            with col2:
                dias_aviso = st.number_input("Dias Aviso Prévio", min_value=0, value=0)
            
            motivo = st.selectbox("Motivo", MOTIVOS_RESCISAO)
            obs = st.text_area("Observações", height=100)
            
            submit = st.form_submit_button("💾 REGISTAR RESCISÃO", use_container_width=True, type="primary")
            
            if submit:
                with st.spinner("🔄 A registar rescisão..."):
                    sucesso = registar_rescisao(
                        empresa_resc, colaborador_resc, ano_resc, mes_resc,
                        data_rescisao, motivo, obs, dias_aviso
                    )
                    
                    if sucesso:
                        st.success(f"✅ Rescisão de {colaborador_resc} registada!")
                        st.info("ℹ️ Este colaborador não aparecerá nos meses seguintes")
                        time.sleep(3)
                        st.session_state.ultimo_reload = datetime.now()
                        st.rerun()

# SIDEBAR
st.sidebar.markdown("---")
st.sidebar.info(f"👤 v2.2 (SEM CACHE)\n💶 SMN: {st.session_state.salario_minimo}€")
st.sidebar.caption("✅ Dados sempre atualizados da Dropbox")

if st.sidebar.button("🚪 Logout", use_container_width=True):
    st.session_state.authenticated = False
    st.rerun()
