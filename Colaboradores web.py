import streamlit as st
import pandas as pd
import dropbox
from datetime import datetime
from io import BytesIO
from openpyxl import load_workbook, Workbook

# ===============================
# CONFIGURAÇÃO
# ===============================
st.set_page_config(page_title="Registo de Colaboradores", page_icon="📋", layout="centered")

DROPBOX_TOKEN = st.secrets["DROPBOX_TOKEN"]
DROPBOX_FILE_PATH = "/Gestão Colaboradores.xlsx"

dbx = dropbox.Dropbox(DROPBOX_TOKEN)

# ===============================
# FUNÇÕES DE VALIDAÇÃO
# ===============================
def validar_email(e):
    if "@" not in e:
        return False
    partes = e.split("@")
    return len(partes) == 2 and len(partes[0]) > 0 and len(partes[1]) > 0

def validar_nif(n): return len(str(n).replace(" ", "")) == 9 and str(n).isdigit()
def validar_niss(n): return len(str(n).replace(" ", "")) == 11 and str(n).isdigit()
def validar_tel(t): return len(str(t).replace(" ", "")) == 9 and str(t).isdigit()
def validar_iban(i):
    i = i.replace(" ", "")
    return i.startswith("PT50") and len(i) == 25 and i[4:].isdigit()

def validar_doc(cc):  # Documento de identificação (livre)
    return len(cc.strip()) > 0

# ===============================
# FUNÇÕES DE LEITURA/ESCRITA NO DROPBOX
# ===============================
def ficheiro_existe():
    try:
        dbx.files_get_metadata(DROPBOX_FILE_PATH)
        return True
    except dropbox.exceptions.ApiError:
        return False

def carregar_dados():
    try:
        _, response = dbx.files_download(DROPBOX_FILE_PATH)
        df = pd.read_excel(BytesIO(response.content), sheet_name="Colaboradores")
        return df
    except Exception:
        colunas = [
            "Nome Completo", "Secção", "Nº Horas/Semana", "E-mail", "Data de Nascimento",
            "NISS", "NIF", "Documento de Identificação", "Validade Documento",
            "Bairro Fiscal", "Estado Civil", "Nº Titulares", "Nº Dependentes",
            "Morada", "IBAN", "Data de Admissão", "Nacionalidade", "Telemóvel", "Data de Registo"
        ]
        return pd.DataFrame(columns=colunas)

def guardar_dados(df):
    try:
        # Verifica se ficheiro já existe
        if ficheiro_existe():
            _, response = dbx.files_download(DROPBOX_FILE_PATH)
            wb = load_workbook(BytesIO(response.content))
            if "Colaboradores" in wb.sheetnames:
                ws = wb["Colaboradores"]
                start_row = ws.max_row + 1
            else:
                ws = wb.create_sheet("Colaboradores")
                start_row = 1
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "Colaboradores"
            start_row = 1

        # Se for novo, escreve cabeçalhos
        if start_row == 1:
            for i, col_name in enumerate(df.columns, start=1):
                ws.cell(row=1, column=i, value=col_name)
            start_row = 2

        # Escreve novas linhas no final
        for r_idx, row in enumerate(df.itertuples(index=False), start=start_row):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        # Guardar e enviar para Dropbox
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        dbx.files_upload(output.read(), DROPBOX_FILE_PATH, mode=dropbox.files.WriteMode.overwrite)
        return True
    except Exception as e:
        st.error(f"Erro ao guardar no Dropbox: {e}")
        return False
# ===============================
# LISTA COMPLETA DOS BAIRROS FISCAIS (SERVIÇOS DE FINANÇAS)
# ===============================
BAIRROS_FISCAIS = [
    "01-AVEIRO – 19-ÁGUEDA", "01-AVEIRO – 27-ALBERGARIA-A-VELHA", "01-AVEIRO – 35-ANADIA",
    "01-AVEIRO – 43-AROUCA", "01-AVEIRO – 51-AVEIRO-1", "01-AVEIRO – 60-CASTELO DE PAIVA",
    "01-AVEIRO – 78-ESPINHO", "01-AVEIRO – 86-ESTARREJA", "01-AVEIRO – 94-ST. MARIA FEIRA-1",
    "01-AVEIRO – 108-ÍLHAVO", "01-AVEIRO – 116-MEALHADA", "01-AVEIRO – 124-MURTOSA",
    "01-AVEIRO – 132-OLIVEIRA AZEMÉIS", "01-AVEIRO – 140-OLIVEIRA DO BAIRRO",
    "01-AVEIRO – 159-OVAR", "01-AVEIRO – 167-S. JOÃO DA MADEIRA", "01-AVEIRO – 175-SEVER DO VOUGA",
    "01-AVEIRO – 183-VAGOS", "01-AVEIRO – 191-VALE DE CAMBRA",
    # … (restantes 400+ linhas idênticas do PDF oficial) …
    "18-VISEU – 2720-VISEU", "18-VISEU – 2739-VOUZELA",
    "19-ANGRA DO HEROÍSMO – 2747-ANGRA DO HEROÍSMO", "19-ANGRA DO HEROÍSMO – 2771-S. CRUZ DA GRACIOSA",
    "20-HORTA – 2917-HORTA", "21-PONTA DELGADA – 2992-PONTA DELGADA",
    "22-FUNCHAL – 2810-FUNCHAL-1", "22-FUNCHAL – 2895-SANTANA"
]
# ===============================
# INTERFACE STREAMLIT
# ===============================

st.title("📋 Registo de Colaboradores")
st.markdown("---")

if not ficheiro_existe():
    st.warning("⚠️ O ficheiro 'Gestão Colaboradores.xlsx' ainda não existe na Dropbox. Será criado automaticamente ao guardar o primeiro registo.")

with st.form("colab_form"):
    st.subheader("Dados Pessoais")
    col1, col2 = st.columns(2)

    with col1:
        nome = st.text_input("Nome Completo *", help="Nome completo do colaborador")
        email = st.text_input("E-mail *", help="Deve conter @ válido")
        data_nasc = st.date_input("Data de Nascimento *", min_value=datetime(1950, 1, 1).date(),
                                  max_value=datetime.now().date(), help="Formato dd/mm/aaaa")
        nif = st.text_input("NIF *", max_chars=9)
        niss = st.text_input("NISS *", max_chars=11)

    with col2:
        tel = st.text_input("Telemóvel *", max_chars=9)
        nac = st.text_input("Nacionalidade *")
        bairro = st.selectbox("Bairro Fiscal *", options=BAIRROS_FISCAIS)
        doc = st.text_input("Documento de Identificação *",
                            help="Formato CC: 12345678 0 ZW0 ou 'Passaporte' ou 'Cartão de Residência'")
        validade = st.date_input("Validade do Documento *", help="Formato dd/mm/aaaa")

    st.subheader("Situação Familiar")
    col3, col4 = st.columns(2)
    with col3:
        estado = st.selectbox("Estado Civil / Nº Titulares *", ["Casado 1", "Casado 2", "Não Casado"])
        titulares = st.number_input("Nº Titulares *", min_value=1, max_value=2, value=1)
    with col4:
        dependentes = st.number_input("Nº Dependentes *", min_value=0, value=0)

    st.subheader("Morada")
    morada = st.text_area("Morada Completa *",
                          help="Completa com rua, lote, porta, andar, código postal e cidade")

    st.subheader("Dados Profissionais")
    col5, col6 = st.columns(2)
    with col5:
        secao = st.selectbox(
            "Secção *",
            ["Charcutaria/Lacticínios", "Frente de Loja", "Frutas e Vegetais", "Gerência",
             "Não Perecíveis (reposição)", "Padaria e Take Away", "Peixaria", "Quiosque", "Talho"]
        )
        horas = st.selectbox("Nº Horas/Semana *", [16, 20, 40])
        admissao = st.date_input("Data de Admissão *", help="Formato dd/mm/aaaa")

    with col6:
        iban = st.text_input("IBAN *", max_chars=25,
                             placeholder="PT50 0000 0000 0000 0000 0000 0",
                             help="Formato PT50 + 21 dígitos (25 caracteres no total)")

    submitted = st.form_submit_button("✅ Submeter Registo", use_container_width=True)

# ===============================
# VALIDAÇÃO E GRAVAÇÃO
# ===============================
if submitted:
    erros = []
    if not nome: erros.append("Nome é obrigatório.")
    if not validar_email(email): erros.append("Email inválido (deve conter @).")
    if not validar_nif(nif): erros.append("NIF inválido (9 dígitos).")
    if not validar_niss(niss): erros.append("NISS inválido (11 dígitos).")
    if not validar_tel(tel): erros.append("Telemóvel inválido (9 dígitos).")
    if not validar_iban(iban): erros.append("IBAN inválido (PT50 + 21 dígitos).")
    if not morada: erros.append("Morada é obrigatória.")
    if not nac: erros.append("Nacionalidade é obrigatória.")
    if not doc: erros.append("Documento de Identificação é obrigatório.")

    if erros:
        st.error("Por favor corrija os seguintes erros:")
        for e in erros: st.error(f"• {e}")
    else:
        novo = {
            "Nome Completo": nome, "Secção": secao, "Nº Horas/Semana": horas,
            "E-mail": email, "Data de Nascimento": data_nasc.strftime("%d/%m/%Y"),
            "NISS": niss, "NIF": nif, "Documento de Identificação": doc,
            "Validade Documento": validade.strftime("%d/%m/%Y"), "Bairro Fiscal": bairro,
            "Estado Civil": estado, "Nº Titulares": titulares, "Nº Dependentes": dependentes,
            "Morada": morada, "IBAN": iban, "Data de Admissão": admissao.strftime("%d/%m/%Y"),
            "Nacionalidade": nac, "Telemóvel": tel,
            "Data de Registo": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        }

        with st.spinner("A guardar..."):
            df = carregar_dados()
            df = pd.concat([df, pd.DataFrame([novo])], ignore_index=True)
            if guardar_dados(df):
                st.success("✅ Registo guardado com sucesso!")
                st.balloons()
                st.info(f"Total de colaboradores registados: {len(df)}")
            else:
                st.error("❌ Erro ao guardar o registo.")
st.markdown("---")
st.caption("Formulário de Registo de Colaboradores | Dados guardados de forma segura no Dropbox")
