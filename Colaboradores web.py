import streamlit as st
import pandas as pd
import dropbox
from openpyxl import load_workbook
from datetime import datetime
from io import BytesIO

# ===============================
# CONFIGURAÇÃO GERAL
# ===============================
st.set_page_config(page_title="Registo de Colaboradores", page_icon="📋", layout="centered")

DROPBOX_TOKEN = st.secrets["DROPBOX_TOKEN"]
DROPBOX_FILE_PATH = "/Gestão Colaboradores.xlsx"

dbx = dropbox.Dropbox(DROPBOX_TOKEN)

# ===============================
# LISTA COMPLETA DE BAIRROS FISCAIS
# ===============================
BAIRROS_FISCAIS = [
    "01-AVEIRO - 19-AGUEDA","01-AVEIRO - 27-ALBERGARIA-A-VELHA","01-AVEIRO - 35-ANADIA","01-AVEIRO - 43-AROUCA",
    "01-AVEIRO - 51-AVEIRO-1","01-AVEIRO - 60-CASTELO DE PAIVA","01-AVEIRO - 78-ESPINHO","01-AVEIRO - 86-ESTARREJA",
    "01-AVEIRO - 94-ST. MARIA FEIRA-1","01-AVEIRO - 108-ILHAVO","01-AVEIRO - 116-MEALHADA","01-AVEIRO - 124-MURTOSA",
    "01-AVEIRO - 132-OLIVEIRA AZEMEIS","01-AVEIRO - 140-OLIVEIRA DO BAIRRO","01-AVEIRO - 159-OVAR","01-AVEIRO - 167-S. JOAO DA MADEIRA",
    "01-AVEIRO - 175-SEVER DO VOUGA","01-AVEIRO - 183-VAGOS","02-BEJA - 205-ALJUSTREL","02-BEJA - 213-ALMODOVAR",
    "02-BEJA - 221-ALVITO","02-BEJA - 230-BARRANCOS","02-BEJA - 248-BEJA","02-BEJA - 256-CASTRO VERDE","02-BEJA - 264-CUBA",
    "02-BEJA - 272-FERREIRA DO ALENTEJO","02-BEJA - 280-MERTOLA","02-BEJA - 299-MOURA","02-BEJA - 302-ODEMIRA","02-BEJA - 310-OURIQUE",
    "02-BEJA - 329-SERPA","02-BEJA - 337-VIDIGUEIRA","07-EVORA - 876-ALANDROAL","07-EVORA - 884-ARRAIOLOS","07-EVORA - 892-BORBA",
    "07-EVORA - 906-ESTREMOZ","07-EVORA - 914-EVORA","07-EVORA - 922-MONTEMOR-O-NOVO","07-EVORA - 930-MORA","07-EVORA - 949-MOURAO",
    "07-EVORA - 957-PORTEL","07-EVORA - 965-REDONDO","07-EVORA - 973-REGUENGOS DE MONSARAZ","07-EVORA - 981-VIANA DO ALENTEJO",
    "07-EVORA - 990-VILA VICOSA","12-PORTALEGRE - 1600-ALTER DO CHAO","12-PORTALEGRE - 1619-ARRONCHES","12-PORTALEGRE - 1627-AVIS",
    "12-PORTALEGRE - 1635-CAMPO MAIOR","12-PORTALEGRE - 1643-CASTELO DE VIDE","12-PORTALEGRE - 1651-CRATO","12-PORTALEGRE - 1660-ELVAS",
    "12-PORTALEGRE - 1678-FRONTEIRA","12-PORTALEGRE - 1686-GAVIAO","12-PORTALEGRE - 1694-MARVAO","12-PORTALEGRE - 1708-MONFORTE",
    "12-PORTALEGRE - 1716-NISA","12-PORTALEGRE - 1724-PONTE DE SOR","12-PORTALEGRE - 1732-PORTALEGRE","12-PORTALEGRE - 1740-SOUSEL"
]

# ===============================
# FUNÇÕES DE VALIDAÇÃO
# ===============================
def validar_email(e): return "@" in e and len(e.split("@")[0])>0 and len(e.split("@")[1])>0
def validar_nif(n): return len(str(n).replace(" ",""))==9 and str(n).isdigit()
def validar_niss(n): return len(str(n).replace(" ",""))==11 and str(n).isdigit()
def validar_tel(t): return len(str(t).replace(" ",""))==9 and str(t).isdigit()
def validar_iban(i): i=i.replace(" ",""); return i.startswith("PT50") and len(i)==25 and i[4:].isdigit()

# ===============================
# FUNÇÕES DROPBOX / EXCEL
# ===============================
def carregar_dados():
    try:
        _, response = dbx.files_download(DROPBOX_FILE_PATH)
        return pd.read_excel(BytesIO(response.content), sheet_name="Colaboradores")
    except Exception:
        colunas = ["Nome Completo","Secção","Nº Horas/Semana","E-mail","Data de Nascimento","NISS","NIF",
                   "Documento de Identificação","Validade Documento","Bairro Fiscal","Estado Civil","Nº Titulares",
                   "Nº Dependentes","Morada","IBAN","Data de Admissão","Nacionalidade","Telemóvel","Data de Registo"]
        return pd.DataFrame(columns=colunas)

def guardar_dados(df):
    try:
        # 1️⃣ Descarregar ficheiro existente
        _, response = dbx.files_download(DROPBOX_FILE_PATH)
        existing_file = BytesIO(response.content)
        wb = load_workbook(existing_file)

        # 2️⃣ Apagar folha antiga se existir
        if "Colaboradores" in wb.sheetnames:
            del wb["Colaboradores"]

        # 3️⃣ Criar nova aba com dados atualizados
        ws = wb.create_sheet("Colaboradores")
        for i, col_name in enumerate(df.columns, start=1):
            ws.cell(row=1, column=i, value=col_name)
        for r_idx, row in enumerate(df.itertuples(index=False), start=2):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        # 4️⃣ Garantir que todas as folhas permanecem visíveis
        for sheet in wb.worksheets:
            sheet.sheet_state = "visible"
        wb.active = 0

        # 5️⃣ Guardar no buffer e reenviar para Dropbox
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        dbx.files_upload(output.read(), DROPBOX_FILE_PATH, mode=dropbox.files.WriteMode.overwrite)
        return True

    except Exception as e:
        st.error(f"Erro ao guardar no Dropbox: {e}")
        return False

# ===============================
# INTERFACE STREAMLIT
# ===============================
st.title("📋 Registo de Colaboradores")
st.markdown("---")

with st.form("colab_form"):
    st.subheader("Dados Pessoais")
    col1,col2 = st.columns(2)
    with col1:
        nome = st.text_input("Nome Completo *")
        email = st.text_input("E-mail *", help="Deve conter @")
        data_nasc = st.date_input("Data de Nascimento *", min_value=datetime(1950,1,1).date(), max_value=datetime.now().date())
        nif = st.text_input("NIF *", max_chars=9)
        niss = st.text_input("NISS *", max_chars=11)
    with col2:
        tel = st.text_input("Telemóvel *", max_chars=9)
        nac = st.text_input("Nacionalidade *")
        bairro = st.selectbox("Bairro Fiscal *", options=BAIRROS_FISCAIS)
        doc = st.text_input("Documento de Identificação *", help="CC, Passaporte ou Cartão de Residência")
        validade = st.date_input("Validade do Documento *")

    st.subheader("Situação Familiar")
    col3,col4 = st.columns(2)
    with col3:
        estado = st.selectbox("Estado Civil / Nº Titulares *", ["Casado 1","Casado 2","Não Casado"])
        titulares = st.number_input("Nº Titulares *", min_value=1, max_value=2, value=1)
    with col4:
        dependentes = st.number_input("Nº Dependentes *", min_value=0, value=0)

    st.subheader("Morada")
    morada = st.text_area("Morada Completa *", help="Completa com rua, lote, porta, andar, código postal e cidade")

    st.subheader("Dados Profissionais")
    col5,col6 = st.columns(2)
    with col5:
        secao = st.selectbox("Secção *", ["Charcutaria/Lacticínios","Frente de Loja","Frutas e Vegetais","Gerência",
                                          "Não Perecíveis (reposição)","Padaria e Take Away","Peixaria","Quiosque","Talho"])
        horas = st.selectbox("Nº Horas/Semana *", [16,20,40])
        admissao = st.date_input("Data de Admissão *")
    with col6:
        iban = st.text_input("IBAN *", max_chars=25, placeholder="PT50 0000 0000 0000 0000 0000 0")

    submitted = st.form_submit_button("✅ Submeter Registo", use_container_width=True)

if submitted:
    erros = []
    if not nome: erros.append("Nome é obrigatório")
    if not validar_email(email): erros.append("Email inválido")
    if not validar_nif(nif): erros.append("NIF inválido")
    if not validar_niss(niss): erros.append("NISS inválido")
    if not validar_tel(tel): erros.append("Telemóvel inválido")
    if not validar_iban(iban): erros.append("IBAN inválido")
    if not morada: erros.append("Morada obrigatória")
    if not nac: erros.append("Nacionalidade obrigatória")

    if erros:
        st.error("Por favor corrija os seguintes erros:")
        for e in erros: st.error(f"• {e}")
    else:
        novo = {
            "Nome Completo": nome, "Secção": secao, "Nº Horas/Semana": horas, "E-mail": email,
            "Data de Nascimento": data_nasc.strftime("%d/%m/%Y"), "NISS": niss, "NIF": nif,
            "Documento de Identificação": doc, "Validade Documento": validade.strftime("%d/%m/%Y"),
            "Bairro Fiscal": bairro, "Estado Civil": estado, "Nº Titulares": titulares,
            "Nº Dependentes": dependentes, "Morada": morada, "IBAN": iban,
            "Data de Admissão": admissao.strftime("%d/%m/%Y"), "Nacionalidade": nac,
            "Telemóvel": tel, "Data de Registo": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
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
