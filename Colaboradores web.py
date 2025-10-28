import streamlit as st
import pandas as pd
import dropbox
from datetime import datetime, date, timedelta
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import calendar

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
    date(2025, 1, 1),   # Ano Novo
    date(2025, 4, 18),  # Sexta-feira Santa
    date(2025, 4, 20),  # Páscoa
    date(2025, 4, 25),  # 25 de Abril
    date(2025, 5, 1),   # Dia do Trabalhador
    date(2025, 6, 10),  # Dia de Portugal
    date(2025, 6, 19),  # Corpo de Deus
    date(2025, 8, 15),  # Assunção
    date(2025, 10, 5),  # Implantação República
    date(2025, 11, 1),  # Todos os Santos
    date(2025, 12, 1),  # Restauração
    date(2025, 12, 8),  # Imaculada Conceição
    date(2025, 12, 25), # Natal
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

# Função de autenticação
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
        st.text_input(
            "Password de Administrador",
            type="password",
            key="password",
            on_change=password_entered
        )
        if "password" in st.session_state and not st.session_state.authenticated:
            st.error("❌ Password incorreta")
        return False
    return True

# Função para carregar colaboradores
def carregar_colaboradores(empresa):
    try:
        _, response = dbx.files_download(EMPRESAS[empresa]["path"])
        df = pd.read_excel(BytesIO(response.content), sheet_name="Colaboradores")
        return df
    except Exception as e:
        st.error(f"Erro ao carregar colaboradores: {e}")
        return pd.DataFrame()

# Função para atualizar colaborador
def atualizar_colaborador_dropbox(empresa, nome_colaborador, dados_atualizados):
    try:
        file_path = EMPRESAS[empresa]["path"]
        _, response = dbx.files_download(file_path)
        wb = load_workbook(BytesIO(response.content))
        
        if "Colaboradores" in wb.sheetnames:
            ws = wb["Colaboradores"]
            
            # Encontrar linha do colaborador
            for row in range(2, ws.max_row + 1):
                if ws.cell(row, 1).value == nome_colaborador:
                    # Atualizar subsídio alimentação (coluna 19)
                    if 'Subsídio Alimentação Diário' in dados_atualizados:
                        ws.cell(row, 19).value = dados_atualizados['Subsídio Alimentação Diário']
                    break
            
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            dbx.files_upload(output.read(), file_path, mode=dropbox.files.WriteMode.overwrite)
            return True
    except Exception as e:
        st.error(f"Erro ao atualizar: {e}")
        return False
def carregar_horas_extras(empresa, mes, ano):
    try:
        _, response = dbx.files_download(EMPRESAS[empresa]["path"])
        df = pd.read_excel(BytesIO(response.content), sheet_name="Horas extra")
        # Filtrar pelo mês e ano
        df_filtrado = df[(df['Mês'] == mes) & (df['Ano'] == ano)]
        return df_filtrado
    except Exception as e:
        st.warning(f"Aviso: Não foi possível carregar horas extras. {e}")
        return pd.DataFrame()

# Função para calcular dias úteis
def calcular_dias_uteis(ano, mes, feriados_list):
    num_dias = calendar.monthrange(ano, mes)[1]
    dias_uteis = 0
    
    for dia in range(1, num_dias + 1):
        data = date(ano, mes, dia)
        # Se não for fim de semana (5=sábado, 6=domingo) e não for feriado
        if data.weekday() < 5 and data not in feriados_list:
            dias_uteis += 1
    
    return dias_uteis

# Função para calcular salário base
def calcular_salario_base(horas_semana, salario_minimo):
    if horas_semana == 40:
        return salario_minimo
    elif horas_semana == 20:
        return salario_minimo / 2
    elif horas_semana == 16:
        return salario_minimo * 0.4
    return 0

# Verificar autenticação
if not check_password():
    st.stop()

# Interface principal
st.title("💰 Processamento Salarial")
st.markdown("---")

# Sidebar - Navegação
menu = st.sidebar.radio(
    "Menu",
    ["⚙️ Configurações", "💼 Processar Salários", "📊 Relatórios"]
)

# PÁGINA DE CONFIGURAÇÕES
if menu == "⚙️ Configurações":
    st.header("⚙️ Configurações do Sistema")
    
    tab_config1, tab_config2 = st.tabs(["💶 Sistema", "👥 Colaboradores"])
    
    with tab_config1:
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
            if st.button("Atualizar Salário Mínimo"):
                st.session_state.salario_minimo = novo_salario
                st.success(f"✅ Salário mínimo atualizado para {novo_salario}€")
        
        with col2:
            st.subheader("📅 Feriados Municipais")
            st.caption("Adicione até 3 feriados municipais")
            
            feriados_temp = []
            for i in range(3):
                feriado = st.date_input(
                    f"Feriado Municipal {i+1}",
                    value=st.session_state.feriados_municipais[i] if i < len(st.session_state.feriados_municipais) else None,
                    key=f"feriado_{i}"
                )
                if feriado:
                    feriados_temp.append(feriado)
            
            if st.button("Atualizar Feriados"):
                st.session_state.feriados_municipais = feriados_temp
                st.success(f"✅ {len(feriados_temp)} feriados municipais configurados")
        
        st.markdown("---")
        st.subheader("📋 Feriados Nacionais 2025")
        st.dataframe(
            pd.DataFrame({
                "Data": [f.strftime("%d/%m/%Y") for f in FERIADOS_NACIONAIS_2025],
                "Descrição": ["Ano Novo", "Sexta-feira Santa", "Páscoa", "25 de Abril", 
                             "Dia do Trabalhador", "Dia de Portugal", "Corpo de Deus",
                             "Assunção", "Implantação República", "Todos os Santos",
                             "Restauração", "Imaculada Conceição", "Natal"]
            }),
            use_container_width=True
        )
    
    with tab_config2:
        st.subheader("👥 Editar Dados de Colaboradores")
        
        empresa_config = st.selectbox(
            "Empresa",
            options=list(EMPRESAS.keys()),
            key="empresa_config"
        )
        
        df_colab_config = carregar_colaboradores(empresa_config)
        
        if not df_colab_config.empty:
            colaborador_config = st.selectbox(
                "Colaborador",
                options=df_colab_config['Nome Completo'].tolist(),
                key="colab_config"
            )
            
            dados_atual = df_colab_config[df_colab_config['Nome Completo'] == colaborador_config].iloc[0]
            
            st.markdown("---")
            
            with st.form("form_editar_colab"):
                st.markdown(f"### Editar: {colaborador_config}")
                
                novo_sub_alim = st.number_input(
                    "Subsídio de Alimentação Diário (€)",
                    min_value=0.0,
                    value=float(dados_atual.get('Subsídio Alimentação Diário', 0)),
                    step=0.10,
                    format="%.2f"
                )
                
                if st.form_submit_button("💾 Guardar Alterações"):
                    if atualizar_colaborador_dropbox(
                        empresa_config,
                        colaborador_config,
                        {'Subsídio Alimentação Diário': novo_sub_alim}
                    ):
                        st.success("✅ Dados atualizados com sucesso!")
                    else:
                        st.error("❌ Erro ao atualizar dados")

# PÁGINA DE PROCESSAMENTO
elif menu == "💼 Processar Salários":
    st.header("💼 Processamento Mensal de Salários")
    
    # Seleção de empresa e período
    col1, col2, col3 = st.columns(3)
    
    with col1:
        empresa_selecionada = st.selectbox(
            "🏢 Empresa",
            options=list(EMPRESAS.keys())
        )
    
    with col2:
        mes_selecionado = st.selectbox(
            "📅 Mês",
            options=range(1, 13),
            format_func=lambda x: calendar.month_name[x]
        )
    
    with col3:
        ano_selecionado = st.number_input(
            "📆 Ano",
            min_value=2024,
            max_value=2030,
            value=2025
        )
    
    st.markdown("---")
    
    # Carregar colaboradores
    df_colaboradores = carregar_colaboradores(empresa_selecionada)
    
    if df_colaboradores.empty:
        st.warning("⚠️ Nenhum colaborador encontrado para esta empresa.")
        st.stop()
    
    # Carregar horas extras se aplicável
    df_horas = pd.DataFrame()
    if EMPRESAS[empresa_selecionada]["tem_horas_extras"]:
        df_horas = carregar_horas_extras(empresa_selecionada, mes_selecionado, ano_selecionado)
    
    # Calcular dias úteis do mês
    todos_feriados = FERIADOS_NACIONAIS_2025 + st.session_state.feriados_municipais
    dias_uteis_mes = calcular_dias_uteis(ano_selecionado, mes_selecionado, todos_feriados)
    
    st.info(f"📊 {len(df_colaboradores)} colaboradores | 📅 Dias úteis no mês: {dias_uteis_mes}")
    
    # Selecionar colaborador
    st.subheader("👤 Selecionar Colaborador")
    colaborador_selecionado = st.selectbox(
        "Nome",
        options=df_colaboradores['Nome Completo'].tolist(),
        key=f"colab_proc_{empresa_selecionada}_{mes_selecionado}_{ano_selecionado}"
    )
    
    # Chave única para dados deste colaborador/mês
    chave_dados = f"{empresa_selecionada}_{colaborador_selecionado}_{mes_selecionado}_{ano_selecionado}"
    
    # Inicializar dados se não existir
    if chave_dados not in st.session_state.dados_processamento:
        st.session_state.dados_processamento[chave_dados] = {
            'faltas_periodos': [],
            'ferias_periodos': [],
            'baixas_periodos': [],
            'sub_ferias': 'Duodécimos',
            'sub_natal': 'Duodécimos',
            'desconto_especie': False,
            'h_extra': 0.0,
            'h_noturnas': 0.0,
            'h_domingos': 0.0,
            'h_feriados': 0.0
        }
    
    dados_salvos = st.session_state.dados_processamento[chave_dados]
    
    # Obter dados do colaborador
    dados_colab = df_colaboradores[df_colaboradores['Nome Completo'] == colaborador_selecionado].iloc[0]
    
    st.markdown("---")
    st.subheader(f"💼 Processar: {colaborador_selecionado}")
    
    # Formulário de processamento
    with st.form("form_processamento"):
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### 📋 Dados Base")
            st.write(f"**Horas/Semana:** {dados_colab['Nº Horas/Semana']}h")
            st.write(f"**Subsídio Alimentação:** {dados_colab['Subsídio Alimentação Diário']}€/dia")
            st.write(f"**Estado Civil:** {dados_colab['Estado Civil']}")
            st.write(f"**Dependentes:** {dados_colab['Nº Dependentes']}")
            
            salario_bruto = calcular_salario_base(dados_colab['Nº Horas/Semana'], st.session_state.salario_minimo)
            st.write(f"**Salário Bruto:** {salario_bruto:.2f}€")
        
        with col2:
            st.markdown("#### 🏖️ Subsídios")
            sub_ferias = st.selectbox(
                "Subsídio de Férias",
                ["Duodécimos", "Total"],
                index=0 if st.session_state.dados_processamento[chave_dados]['sub_ferias'] == 'Duodécimos' else 1
            )
            sub_natal = st.selectbox(
                "Subsídio de Natal",
                ["Duodécimos", "Total"],
                index=0 if st.session_state.dados_processamento[chave_dados]['sub_natal'] == 'Duodécimos' else 1
            )
            desconto_especie = st.checkbox(
                "Desconto em espécie (cartão refeição)",
                value=st.session_state.dados_processamento[chave_dados]['desconto_especie']
            )
        
        st.markdown("#### 📅 Faltas, Férias e Baixas")
        
        tab1, tab2, tab3 = st.tabs(["🔴 Faltas", "🟢 Férias", "🟡 Baixas"])
        
        # FALTAS
        with tab1:
            faltas_periodos = []
            for i in range(3):
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    falta_inicio = st.date_input(
                        f"Falta {i+1} - Início",
                        value=None,
                        key=f"falta_inicio_{i}",
                        min_value=date(ano_selecionado, mes_selecionado, 1),
                        max_value=date(ano_selecionado, mes_selecionado, calendar.monthrange(ano_selecionado, mes_selecionado)[1])
                    )
                with col_f2:
                    falta_fim = st.date_input(
                        f"Falta {i+1} - Fim",
                        value=None,
                        key=f"falta_fim_{i}",
                        min_value=date(ano_selecionado, mes_selecionado, 1),
                        max_value=date(ano_selecionado, mes_selecionado, calendar.monthrange(ano_selecionado, mes_selecionado)[1])
                    )
                if falta_inicio and falta_fim:
                    faltas_periodos.append((falta_inicio, falta_fim))
        
        # FÉRIAS
        with tab2:
            ferias_periodos = []
            for i in range(3):
                col_v1, col_v2 = st.columns(2)
                with col_v1:
                    ferias_inicio = st.date_input(
                        f"Férias {i+1} - Início",
                        value=None,
                        key=f"ferias_inicio_{i}",
                        min_value=date(ano_selecionado, mes_selecionado, 1),
                        max_value=date(ano_selecionado, mes_selecionado, calendar.monthrange(ano_selecionado, mes_selecionado)[1])
                    )
                with col_v2:
                    ferias_fim = st.date_input(
                        f"Férias {i+1} - Fim",
                        value=None,
                        key=f"ferias_fim_{i}",
                        min_value=date(ano_selecionado, mes_selecionado, 1),
                        max_value=date(ano_selecionado, mes_selecionado, calendar.monthrange(ano_selecionado, mes_selecionado)[1])
                    )
                if ferias_inicio and ferias_fim:
                    ferias_periodos.append((ferias_inicio, ferias_fim))
        
        # BAIXAS
        with tab3:
            baixas_periodos = []
            for i in range(3):
                col_b1, col_b2 = st.columns(2)
                with col_b1:
                    baixa_inicio = st.date_input(
                        f"Baixa {i+1} - Início",
                        value=None,
                        key=f"baixa_inicio_{i}",
                        min_value=date(ano_selecionado, mes_selecionado, 1),
                        max_value=date(ano_selecionado, mes_selecionado, calendar.monthrange(ano_selecionado, mes_selecionado)[1])
                    )
                with col_b2:
                    baixa_fim = st.date_input(
                        f"Baixa {i+1} - Fim",
                        value=None,
                        key=f"baixa_fim_{i}",
                        min_value=date(ano_selecionado, mes_selecionado, 1),
                        max_value=date(ano_selecionado, mes_selecionado, calendar.monthrange(ano_selecionado, mes_selecionado)[1])
                    )
                if baixa_inicio and baixa_fim:
                    baixas_periodos.append((baixa_inicio, baixa_fim))
        
        st.markdown("---")
        
        # Horas extras
        st.markdown("#### ⏰ Horas Extras")
        
        if EMPRESAS[empresa_selecionada]["tem_horas_extras"] and not df_horas.empty:
            # Procurar dados do colaborador
            horas_colab = df_horas[df_horas['Nome Completo'] == colaborador_selecionado]
            if not horas_colab.empty:
                h_noturnas = horas_colab.iloc[0].get('Noturnas', 0)
                h_domingos = horas_colab.iloc[0].get('Domingos', 0)
                h_feriados = horas_colab.iloc[0].get('Feriados', 0)
                h_extra = horas_colab.iloc[0].get('Extra', 0)
                
                col_h1, col_h2, col_h3, col_h4 = st.columns(4)
                col_h1.metric("🌙 Noturnas", f"{h_noturnas}h")
                col_h2.metric("☀️ Domingos", f"{h_domingos}h")
                col_h3.metric("🎉 Feriados", f"{h_feriados}h")
                col_h4.metric("⏱️ Extra", f"{h_extra}h")
            else:
                st.info("ℹ️ Sem horas extras registadas para este colaborador/mês")
                h_noturnas = h_domingos = h_feriados = h_extra = 0
        else:
            # Magnetic Sky - apenas banco de horas manual
            h_extra = st.number_input(
                "Banco de Horas",
                min_value=0.0,
                step=0.5,
                format="%.2f",
                help="Número de horas extra a pagar"
            )
            h_noturnas = h_domingos = h_feriados = 0
        
        submitted = st.form_submit_button("💰 Calcular Recibo", use_container_width=True)
        
        if submitted:
            # Calcular dias de faltas, férias e baixas
            def contar_dias(periodos):
                total_dias = 0
                dias_uteis = 0
                for inicio, fim in periodos:
                    dias_periodo = (fim - inicio).days + 1
                    total_dias += dias_periodo
                    # Contar dias úteis (excluindo fins de semana e feriados)
                    for i in range(dias_periodo):
                        dia = inicio + timedelta(days=i)
                        if dia.weekday() < 5 and dia not in todos_feriados:
                            dias_uteis += 1
                return total_dias, dias_uteis
            
            # Guardar dados no session_state
            st.session_state.dados_processamento[chave_dados] = {
                'faltas': faltas_periodos,
                'ferias': ferias_periodos,
                'baixas': baixas_periodos,
                'sub_ferias': sub_ferias,
                'sub_natal': sub_natal,
                'desconto_especie': desconto_especie,
                'h_extra': h_extra,
                'h_noturnas': h_noturnas,
                'h_domingos': h_domingos,
                'h_feriados': h_feriados
            }
            
            dias_faltas, dias_faltas_uteis = contar_dias(faltas_periodos)
            dias_ferias, dias_ferias_uteis = contar_dias(ferias_periodos)
            dias_baixas, dias_baixas_uteis = contar_dias(baixas_periodos)
            
            # Calcular dias trabalhados
            num_dias_mes = calendar.monthrange(ano_selecionado, mes_selecionado)[1]
            dias_trabalhados = num_dias_mes - dias_faltas - dias_baixas
            dias_uteis_trabalhados = dias_uteis_mes - dias_faltas_uteis - dias_baixas_uteis - dias_ferias_uteis
            
            # Mostrar resumo
            st.markdown("---")
            st.subheader("📊 Resumo do Processamento")
            
            col_r1, col_r2, col_r3, col_r4 = st.columns(4)
            col_r1.metric("📅 Dias do Mês", num_dias_mes)
            col_r2.metric("🔴 Faltas", f"{dias_faltas} ({dias_faltas_uteis} úteis)")
            col_r3.metric("🟢 Férias", f"{dias_ferias} ({dias_ferias_uteis} úteis)")
            col_r4.metric("🟡 Baixas", f"{dias_baixas} ({dias_baixas_uteis} úteis)")
            
            col_r5, col_r6 = st.columns(2)
            col_r5.metric("💼 Dias Trabalhados", dias_trabalhados)
            col_r6.metric("📊 Dias Úteis Trabalhados", dias_uteis_trabalhados)
            
            st.info("🚧 Módulo 3 em construção: Cálculos de remunerações e descontos...")

# PÁGINA DE RELATÓRIOS
elif menu == "📊 Relatórios":
    st.header("📊 Relatórios e Histórico")
    st.info("🚧 Em desenvolvimento...")

st.sidebar.markdown("---")
if st.sidebar.button("🚪 Logout"):
    st.session_state.authenticated = False
    st.rerun()
