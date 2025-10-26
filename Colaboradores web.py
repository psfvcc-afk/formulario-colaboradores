import streamlit as st
import pandas as pd
import dropbox
from datetime import datetime
import re
from io import BytesIO

# Configuração da página
st.set_page_config(
    page_title="Registo de Colaboradores",
    page_icon="📋",
    layout="centered"
)

# Token do Dropbox - IMPORTANTE: Usar secrets no Streamlit Cloud
DROPBOX_TOKEN = st.secrets["DROPBOX_TOKEN"]
DROPBOX_FILE_PATH = "/colaboradores.xlsx"

# Inicializar cliente Dropbox
dbx = dropbox.Dropbox(DROPBOX_TOKEN)

# Lista de Bairros Fiscais (Serviços de Finanças)
BAIRROS_FISCAIS = [
    "01-AVEIRO - 19-AGUEDA", "01-AVEIRO - 27-ALBERGARIA-A-VELHA", "01-AVEIRO - 35-ANADIA",
    "01-AVEIRO - 43-AROUCA", "01-AVEIRO - 51-AVEIRO-1", "01-AVEIRO - 60-CASTELO DE PAIVA",
    "01-AVEIRO - 78-ESPINHO", "01-AVEIRO - 86-ESTARREJA", "01-AVEIRO - 94-ST. MARIA FEIRA-1",
    "01-AVEIRO - 108-ILHAVO", "01-AVEIRO - 116-MEALHADA", "01-AVEIRO - 124-MURTOSA",
    "01-AVEIRO - 132-OLIVEIRA AZEMEIS", "01-AVEIRO - 140-OLIVEIRA DO BAIRRO", "01-AVEIRO - 159-OVAR",
    "01-AVEIRO - 167-S. JOAO DA MADEIRA", "01-AVEIRO - 175-SEVER DO VOUGA", "01-AVEIRO - 183-VAGOS",
    "01-AVEIRO - 191-VALE DE CAMBRA", "02-BEJA - 205-ALJUSTREL", "02-BEJA - 213-ALMODOVAR",
    "02-BEJA - 221-ALVITO", "02-BEJA - 230-BARRANCOS", "02-BEJA - 248-BEJA", "02-BEJA - 256-CASTRO VERDE",
    "02-BEJA - 264-CUBA", "02-BEJA - 272-FERREIRA DO ALENTEJO", "02-BEJA - 280-MERTOLA",
    "02-BEJA - 299-MOURA", "02-BEJA - 302-ODEMIRA", "02-BEJA - 310-OURIQUE", "02-BEJA - 329-SERPA",
    "02-BEJA - 337-VIDIGUEIRA", "03-BRAGA - 345-AMARES", "03-BRAGA - 353-BARCELOS", "03-BRAGA - 361-BRAGA-1",
    "03-BRAGA - 400-FAFE", "03-BRAGA - 418-GUIMARAES-1", "03-BRAGA - 442-VIEIRA DO MINHO",
    "07-EVORA - 906-ESTREMOZ", "07-EVORA - 914-EVORA", "07-EVORA - 922-MONTEMOR-O-NOVO",
    "07-EVORA - 965-REDONDO", "07-EVORA - 973-REGUENGOS DE MONSARAZ", "07-EVORA - 990-VILA VICOSA",
    "12-PORTALEGRE - 1635-CAMPO MAIOR", "12-PORTALEGRE - 1660-ELVAS", "12-PORTALEGRE - 1732-PORTALEGRE",
    "15-SETUBAL - 2232-SETUBAL-1", "15-SETUBAL - 2240-SESIMBRA", "15-SETUBAL - 2259-SINES",
    "18-VISEU - 2720-VISEU"
]

# ---------------------- FUNÇÕES DE VALIDAÇÃO ----------------------

def validar_email(email):
    """Valida que o email contém '@' com texto antes e depois"""
    if "@" not in email:
        return False
    partes = email.split("@")
    return len(partes) == 2 and len(partes[0]) > 0 and len(partes[1]) > 0

def validar_nif(nif):
    nif_clean = str(nif).replace(" ", "")
    return len(nif_clean) == 9 and nif_clean.isdigit()

def validar_niss(niss):
    niss_clean = str(niss).replace(" ", "")
    return len(niss_clean) == 11 and niss_clean.isdigit()

def validar_telemovel(tel):
    tel_clean = str(tel).replace(" ", "")
    return len(tel_clean) == 9 and tel_clean.isdigit()

def validar_iban(iban):
    """Valida que começa com PT50 e tem 25 caracteres (PT50 + 21 dígitos)"""
    iban_clean = iban.replace(" ", "")
    if not iban_clean.startswith("PT50"):
        return False
    if len(iban_clean) != 25:
        return False
    return iban_clean[4:].isdigit()

def validar_cc(cc):
    """Aceita qualquer texto — apenas verifica que não está vazio"""
    return len(cc.strip()) > 0

# ---------------------- FUNÇÕES DE LER/GRAVAR ----------------------

def carregar_dados_dropbox():
    """Carrega o Excel de colaboradores, ou cria um novo se não existir"""
    try:
        _, response = dbx.files_download(DROPBOX_FILE_PATH)
        data = response.content
        df = pd.read_excel(BytesIO(data))
        return df
    except:
        colunas = [
            "Nome Completo", "Secção", "Nº Horas/Semana", "E-mail", "Data de Nascimento",
            "NISS", "NIF", "Documento de Identificação", "Validade Documento",
            "Bairro Fiscal", "Estado Civil", "Nº Titulares", "Nº Dependentes",
            "Morada", "IBAN", "Data de Admissão", "Nacionalidade",
            "Telemóvel", "Data de Registo"
        ]
        return pd.DataFrame(columns=colunas)

def guardar_dados_dropbox(df):
    """Guarda o DataFrame no Dropbox (substitui ficheiro anterior)"""
    try:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Colaboradores')
        output.seek(0)
        dbx.files_upload(
            output.read(),
            DROPBOX_FILE_PATH,
            mode=dropbox.files.WriteMode.overwrite
        )
        return True
    except Exception as e:
        st.error(f"Erro ao guardar no Dropbox: {e}")
        return False
# ---------------------- INTERFACE STREAMLIT ----------------------

st.title("📋 Registo de Colaboradores")
st.markdown("---")

# Formulário principal
with st.form("formulario_colaborador"):
    st.subheader("Dados Pessoais")

    col1, col2 = st.columns(2)

    with col1:
        nome = st.text_input("Nome Completo *", help="Nome completo do colaborador")
        email = st.text_input("E-mail *", help="Email corporativo ou pessoal (deve conter @)")
        data_nascimento = st.date_input(
            "Data de Nascimento *",
            min_value=datetime(1950, 1, 1).date(),
            max_value=datetime.now().date(),
            help="Formato: dd/mm/aaaa"
        )
        nif = st.text_input("NIF *", max_chars=9, help="9 dígitos")
        niss = st.text_input("NISS *", max_chars=11, help="11 dígitos")

    with col2:
        telemovel = st.text_input("Telemóvel *", max_chars=9, help="9 dígitos")
        nacionalidade = st.text_input("Nacionalidade *", help="Ex: Portuguesa")
        bairro_fiscal = st.selectbox(
            "Bairro Fiscal *",
            options=BAIRROS_FISCAIS,
            help="Serviço de finanças da área de residência"
        )
        doc_identificacao = st.text_input(
            "Documento de Identificação *",
            help="Formato CC: 12345678 0 ZW0 ou 'Passaporte' ou 'Cartão de Residência'"
        )
        validade_doc = st.date_input("Validade do Documento *", help="Formato: dd/mm/aaaa")

    st.subheader("Situação Familiar")
    col3, col4 = st.columns(2)

    with col3:
        estado_civil = st.selectbox(
            "Estado Civil / Nº Titulares *",
            ["Casado 1", "Casado 2", "Não Casado"],
            help="Casado 1: único titular casado | Casado 2: ambos titulares | Não Casado"
        )
        num_titulares = st.number_input(
            "Nº Titulares *", min_value=1, max_value=2, value=1,
            help="Número de titulares do agregado familiar"
        )

    with col4:
        num_dependentes = st.number_input(
            "Nº Dependentes *", min_value=0, value=0,
            help="Número de dependentes a cargo"
        )

    st.subheader("Morada")
    morada = st.text_area(
        "Morada Completa *",
        help="Completa com rua, lote, porta, andar, código postal e cidade"
    )

    st.subheader("Dados Profissionais")

    col5, col6 = st.columns(2)

    with col5:
        secao = st.selectbox(
            "Secção *",
            options=[
                "Charcutaria/Lacticínios", "Frente de Loja", "Frutas e Vegetais",
                "Gerência", "Não Perecíveis (reposição)", "Padaria e Take Away",
                "Peixaria", "Quiosque", "Talho"
            ],
            help="Departamento ou secção do colaborador"
        )
        horas_semana = st.selectbox(
            "Nº Horas/Semana *",
            [16, 20, 40],
            help="Horas de trabalho semanais (16h, 20h ou 40h)"
        )
        data_admissao = st.date_input("Data de Admissão *", help="Formato: dd/mm/aaaa")

    with col6:
        iban = st.text_input(
            "IBAN *",
            max_chars=25,
            placeholder="PT50 0000 0000 0000 0000 0000 0",
            help="Formato: PT50 seguido de 21 dígitos (25 caracteres no total)"
        )

    st.markdown("---")
    st.caption("* Campos obrigatórios")

    submitted = st.form_submit_button("✅ Submeter Registo", use_container_width=True)

    # ---------------------- VALIDAÇÕES ----------------------
    if submitted:
        erros = []

        if not nome or len(nome) < 3:
            erros.append("Nome completo é obrigatório")
        if not email or not validar_email(email):
            erros.append("Email inválido (deve conter @)")
        if not nif or not validar_nif(nif):
            erros.append("NIF deve ter 9 dígitos")
        if not niss or not validar_niss(niss):
            erros.append("NISS deve ter 11 dígitos")
        if not telemovel or not validar_telemovel(telemovel):
            erros.append("Telemóvel deve ter 9 dígitos")
        if not doc_identificacao or not validar_cc(doc_identificacao):
            erros.append("Documento de identificação em formato inválido")
        if not iban or not validar_iban(iban):
            erros.append("IBAN deve estar no formato PT50 seguido de 21 dígitos")
        if not morada or len(morada) < 10:
            erros.append("Morada completa é obrigatória")
        if not nacionalidade:
            erros.append("Nacionalidade é obrigatória")

        # ---------------------- RESULTADOS ----------------------
        if erros:
            st.error("Por favor corrija os seguintes erros:")
            for erro in erros:
                st.error(f"• {erro}")
        else:
            novo_registo = {
                "Nome Completo": nome,
                "Secção": secao,
                "Nº Horas/Semana": horas_semana,
                "E-mail": email,
                "Data de Nascimento": data_nascimento.strftime("%d/%m/%Y"),
                "NISS": niss,
                "NIF": nif,
                "Documento de Identificação": doc_identificacao,
                "Validade Documento": validade_doc.strftime("%d/%m/%Y"),
                "Bairro Fiscal": bairro_fiscal,
                "Estado Civil": estado_civil,
                "Nº Titulares": num_titulares,
                "Nº Dependentes": num_dependentes,
                "Morada": morada,
                "IBAN": iban,
                "Data de Admissão": data_admissao.strftime("%d/%m/%Y"),
                "Nacionalidade": nacionalidade,
                "Telemóvel": telemovel,
                "Data de Registo": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            }

            with st.spinner("A guardar..."):
                df = carregar_dados_dropbox()
                df = pd.concat([df, pd.DataFrame([novo_registo])], ignore_index=True)

                if guardar_dados_dropbox(df):
                    st.success("✅ Registo guardado com sucesso!")
                    st.balloons()
                    st.info(f"Total de colaboradores registados: {len(df)}")
                else:
                    st.error("❌ Erro ao guardar o registo. Tente novamente.")

# Rodapé
st.markdown("---")
st.caption("Formulário de Registo de Colaboradores | Dados guardados de forma segura no Dropbox")
