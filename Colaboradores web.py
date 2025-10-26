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
    "01-AVEIRO - 19-AGUEDA",
    "01-AVEIRO - 27-ALBERGARIA-A-VELHA",
    "01-AVEIRO - 35-ANADIA",
    "01-AVEIRO - 43-AROUCA",
    "01-AVEIRO - 51-AVEIRO-1",
    "01-AVEIRO - 60-CASTELO DE PAIVA",
    "01-AVEIRO - 78-ESPINHO",
    "01-AVEIRO - 86-ESTARREJA",
    "01-AVEIRO - 94-ST. MARIA FEIRA-1",
    "01-AVEIRO - 108-ILHAVO",
    "01-AVEIRO - 116-MEALHADA",
    "01-AVEIRO - 124-MURTOSA",
    "01-AVEIRO - 132-OLIVEIRA AZEMEIS",
    "01-AVEIRO - 140-OLIVEIRA DO BAIRRO",
    "01-AVEIRO - 159-OVAR",
    "01-AVEIRO - 167-S. JOAO DA MADEIRA",
    "01-AVEIRO - 175-SEVER DO VOUGA",
    "01-AVEIRO - 183-VAGOS",
    "01-AVEIRO - 191-VALE DE CAMBRA",
    "01-AVEIRO - 3417-AVEIRO-2",
    "01-AVEIRO - 3441-ST. MARIA FEIRA-2",
    "01-AVEIRO - 3735-ST. MARIA FEIRA 3",
    "01-AVEIRO - 4170-ST. MARIA FEIRA 4",
    "02-BEJA - 205-ALJUSTREL",
    "02-BEJA - 213-ALMODOVAR",
    "02-BEJA - 221-ALVITO",
    "02-BEJA - 230-BARRANCOS",
    "02-BEJA - 248-BEJA",
    "02-BEJA - 256-CASTRO VERDE",
    "02-BEJA - 264-CUBA",
    "02-BEJA - 272-FERREIRA DO ALENTEJO",
    "02-BEJA - 280-MERTOLA",
    "02-BEJA - 299-MOURA",
    "02-BEJA - 302-ODEMIRA",
    "02-BEJA - 310-OURIQUE",
    "02-BEJA - 329-SERPA",
    "02-BEJA - 337-VIDIGUEIRA",
    "03-BRAGA - 345-AMARES",
    "03-BRAGA - 353-BARCELOS",
    "03-BRAGA - 361-BRAGA-1",
    "03-BRAGA - 370-CABECEIRAS DE BASTO",
    "03-BRAGA - 388-CELORICO DE BASTO",
    "03-BRAGA - 396-ESPOSENDE",
    "03-BRAGA - 400-FAFE",
    "03-BRAGA - 418-GUIMARAES-1",
    "03-BRAGA - 426-POVOA DE LANHOSO",
    "03-BRAGA - 434-TERRAS DE BOURO",
    "03-BRAGA - 442-VIEIRA DO MINHO",
    "03-BRAGA - 450-VILA N.FAMALICAO-1",
    "03-BRAGA - 469-VILA VERDE",
    "03-BRAGA - 3425-BRAGA-2",
    "03-BRAGA - 3476-GUIMARAES-2",
    "03-BRAGA - 3590-VILA N.FAMALICAO 2",
    "03-BRAGA - 4200-VIZELA",
    "04-BRAGANCA - 477-ALFANDEGA DA FE",
    "04-BRAGANCA - 485-BRAGANCA",
    "04-BRAGANCA - 493-CARRAZEDA DE ANSIAES",
    "04-BRAGANCA - 507-FR DE ESPADA A CINTA",
    "04-BRAGANCA - 515-MACEDO DE CAVALEIROS",
    "04-BRAGANCA - 523-MIRANDA DO DOURO",
    "04-BRAGANCA - 531-MIRANDELA",
    "04-BRAGANCA - 540-MOGADOURO",
    "04-BRAGANCA - 558-TORRE DE MONCORVO",
    "04-BRAGANCA - 566-VILA FLOR",
    "04-BRAGANCA - 574-VIMIOSO",
    "04-BRAGANCA - 582-VINHAIS",
    "05-C BRANCO - 590-BELMONTE",
    "05-C BRANCO - 604-CASTELO BRANCO-1",
    "05-C BRANCO - 612-COVILHA",
    "05-C BRANCO - 620-FUNDAO",
    "05-C BRANCO - 639-IDANHA-A-NOVA",
    "05-C BRANCO - 647-OLEIROS",
    "05-C BRANCO - 655-PENAMACOR",
    "05-C BRANCO - 663-PROENCA-A-NOVA",
    "05-C BRANCO - 671-SERTA",
    "05-C BRANCO - 680-VILA DE REI",
    "05-C BRANCO - 698-VILA VELHA DE RODAO",
    "05-C BRANCO - 3794-CASTELO BRANCO-2",
    "06-COIMBRA - 701-ARGANIL",
    "06-COIMBRA - 710-CANTANHEDE",
    "06-COIMBRA - 728-COIMBRA-1",
    "06-COIMBRA - 736-CONDEIXA-A-NOVA",
    "06-COIMBRA - 744-FIGUEIRA DA FOZ-1",
    "06-COIMBRA - 752-GOIS",
    "06-COIMBRA - 760-LOUSA",
    "06-COIMBRA - 779-MIRA",
    "06-COIMBRA - 787-MIRANDA DO CORVO",
    "06-COIMBRA - 795-MONTEMOR-O-VELHO",
    "06-COIMBRA - 809-OLIVEIRA DO HOSPITAL",
    "06-COIMBRA - 817-PAMPILHOSA DA SERRA",
    "06-COIMBRA - 825-PENACOVA",
    "06-COIMBRA - 833-PENELA",
    "06-COIMBRA - 841-VILA NOVA DE POIARES",
    "06-COIMBRA - 850-SOURE",
    "06-COIMBRA - 868-TABUA",
    "06-COIMBRA - 3050-COIMBRA-2",
    "06-COIMBRA - 3824-FIGUEIRA DA FOZ 2",
    "07-EVORA - 876-ALANDROAL",
    "07-EVORA - 884-ARRAIOLOS",
    "07-EVORA - 892-BORBA",
    "07-EVORA - 906-ESTREMOZ",
    "07-EVORA - 914-EVORA",
    "07-EVORA - 922-MONTEMOR-O-NOVO",
    "07-EVORA - 930-MORA",
    "07-EVORA - 949-MOURAO",
    "07-EVORA - 957-PORTEL",
    "07-EVORA - 965-REDONDO",
    "07-EVORA - 973-REGUENGOS DE MONSARAZ",
    "07-EVORA - 981-VIANA DO ALENTEJO",
    "07-EVORA - 990-VILA VICOSA",
    "07-EVORA - 3042-VENDAS NOVAS",
    "08-FARO - 1007-ALBUFEIRA",
    "08-FARO - 1015-ALCOUTIM",
    "08-FARO - 1023-ALJEZUR",
    "08-FARO - 1031-S.BRAS DE ALPORTEL",
    "08-FARO - 1040-CASTRO MARIM",
    "08-FARO - 1058-FARO",
    "08-FARO - 1066-LAGOA (ALGARVE)",
    "08-FARO - 1074-LAGOS",
    "08-FARO - 1082-LOULE-1",
    "08-FARO - 1090-MONCHIQUE",
    "08-FARO - 1104-OLHAO",
    "08-FARO - 1112-PORTIMAO",
    "08-FARO - 1120-SILVES",
    "08-FARO - 1139-TAVIRA",
    "08-FARO - 1147-VILA DO BISPO",
    "08-FARO - 1155-VILA REAL S.ANTONIO",
    "08-FARO - 3859-LOULE-2",
    "09-GUARDA - 1163-AGUIAR DA BEIRA",
    "09-GUARDA - 1171-ALMEIDA",
    "09-GUARDA - 1180-CELORICO DA BEIRA",
    "09-GUARDA - 1198-FIG. CASTELO RODRIGO",
    "09-GUARDA - 1201-FORNOS DE ALGODRES",
    "09-GUARDA - 1210-GOUVEIA",
    "09-GUARDA - 1228-GUARDA",
    "09-GUARDA - 1236-MANTEIGAS",
    "09-GUARDA - 1244-MEDA",
    "09-GUARDA - 1252-PINHEL",
    "09-GUARDA - 1260-SABUGAL",
    "09-GUARDA - 1279-SEIA",
    "09-GUARDA - 1287-TRANCOSO",
    "09-GUARDA - 1295-VILA NOVA DE FOZ COA",
    "10-LEIRIA - 1309-ALCOBACA",
    "10-LEIRIA - 1317-ALVAIAZERE",
    "10-LEIRIA - 1325-ANSIAO",
    "10-LEIRIA - 1333-BATALHA",
    "10-LEIRIA - 1341-BOMBARRAL",
    "10-LEIRIA - 1350-CALDAS DA RAINHA",
    "10-LEIRIA - 1368-CASTANHEIRA DE PERA",
    "10-LEIRIA - 1376-FIGUEIRO DOS VINHOS",
    "10-LEIRIA - 1384-LEIRIA-1",
    "10-LEIRIA - 1392-MARINHA GRANDE",
    "10-LEIRIA - 1406-NAZARE",
    "10-LEIRIA - 1414-OBIDOS",
    "10-LEIRIA - 1422-PEDROGAO GRANDE",
    "10-LEIRIA - 1430-PENICHE",
    "10-LEIRIA - 1449-POMBAL",
    "10-LEIRIA - 1457-PORTO DE MOS",
    "10-LEIRIA - 3603-LEIRIA-2",
    "11-LISBOA - 1465-ALENQUER",
    "11-LISBOA - 1473-ARRUDA DOS VINHOS",
    "11-LISBOA - 1481-AZAMBUJA",
    "11-LISBOA - 1490-CADAVAL",
    "11-LISBOA - 1503-CASCAIS-1",
    "11-LISBOA - 1520-LOURES-1",
    "11-LISBOA - 1538-LOURINHA",
    "11-LISBOA - 1546-MAFRA",
    "11-LISBOA - 1554-OEIRAS-1",
    "11-LISBOA - 1562-SINTRA-1",
    "11-LISBOA - 1570-SOBRAL MONTE AGRACO",
    "11-LISBOA - 1589-TORRES VEDRAS",
    "11-LISBOA - 1597-VILA FRANCA XIRA-1",
    "11-LISBOA - 3069-LISBOA-1 BAIRRO",
    "11-LISBOA - 3085-LISBOA-3 BAIRRO",
    "11-LISBOA - 3107-LISBOA-8 BAIRRO",
    "11-LISBOA - 3131-AMADORA-1",
    "11-LISBOA - 3140-AMADORA-2",
    "11-LISBOA - 3158-LOURES-3. MOSCAVIDE",
    "11-LISBOA - 3166-SINTRA-4. QUELUZ",
    "11-LISBOA - 3239-LISBOA-7 BAIRRO",
    "11-LISBOA - 3247-LISBOA-2 BAIRRO",
    "11-LISBOA - 3255-LISBOA-10 BAIRRO",
    "11-LISBOA - 3263-LISBOA-5 BAIRRO",
    "11-LISBOA - 3301-LISBOA-4 BAIRRO",
    "11-LISBOA - 3328-LISBOA-9 BAIRRO",
    "11-LISBOA - 3336-LISBOA-6 BAIRRO",
    "11-LISBOA - 3344-LISBOA-11 BAIRRO",
    "11-LISBOA - 3433-CASCAIS-2",
    "11-LISBOA - 3492-LOURES-4",
    "11-LISBOA - 3522-OEIRAS-3.ALGES",
    "11-LISBOA - 3549-SINTRA-2. ALGUEIRAO",
    "11-LISBOA - 3557-SINTRA-3. CACEM",
    "11-LISBOA - 3573-VILA FRANCA XIRA-2",
    "11-LISBOA - 3611-AMADORA-3",
    "11-LISBOA - 3654-OEIRAS-2",
    "11-LISBOA - 4227-ODIVELAS",
    "12-PORTALEGRE - 1600-ALTER DO CHAO",
    "12-PORTALEGRE - 1619-ARRONCHES",
    "12-PORTALEGRE - 1627-AVIS",
    "12-PORTALEGRE - 1635-CAMPO MAIOR",
    "12-PORTALEGRE - 1643-CASTELO DE VIDE",
    "12-PORTALEGRE - 1651-CRATO",
    "12-PORTALEGRE - 1660-ELVAS",
    "12-PORTALEGRE - 1678-FRONTEIRA",
    "12-PORTALEGRE - 1686-GAVIAO",
    "12-PORTALEGRE - 1694-MARVAO",
    "12-PORTALEGRE - 1708-MONFORTE",
    "12-PORTALEGRE - 1716-NISA",
    "12-PORTALEGRE - 1724-PONTE DE SOR",
    "12-PORTALEGRE - 1732-PORTALEGRE",
    "12-PORTALEGRE - 1740-SOUSEL",
    "13-PORTO - 1759-AMARANTE",
    "13-PORTO - 1767-BAIAO",
    "13-PORTO - 1775-FELGUEIRAS",
    "13-PORTO - 1783-GONDOMAR-1",
    "13-PORTO - 1791-LOUSADA",
    "13-PORTO - 1805-MAIA",
    "13-PORTO - 1813-MARCO DE CANAVESES",
    "13-PORTO - 1821-MATOSINHOS-1",
    "13-PORTO - 1830-PACOS DE FERREIRA",
    "13-PORTO - 1848-PAREDES",
    "13-PORTO - 1856-PENAFIEL",
    "13-PORTO - 1872-POVOA DE VARZIM",
    "13-PORTO - 1880-SANTO TIRSO",
    "13-PORTO - 1899-VALONGO-1",
    "13-PORTO - 1902-VILA DO CONDE",
    "13-PORTO - 1910-VILA NOVA DE GAIA-1",
    "13-PORTO - 3174-PORTO-1 BAIRRO",
    "13-PORTO - 3182-PORTO-2 BAIRRO",
    "13-PORTO - 3190-PORTO-5 BAIRRO",
    "13-PORTO - 3204-VILA NOVA DE GAIA-2",
    "13-PORTO - 3360-PORTO-3 BAIRRO",
    "13-PORTO - 3387-PORTO-4 BAIRRO",
    "13-PORTO - 3468-GONDOMAR-2",
    "13-PORTO - 3514-MATOSINHOS-2",
    "13-PORTO - 3565-VALONGO-2. ERMESINDE",
    "13-PORTO - 3964-VILA NOVA DE GAIA 3",
    "13-PORTO - 4219-TROFA",
    "14-SANTAREM - 1929-ABRANTES-1",
    "14-SANTAREM - 1937-ALCANENA",
    "14-SANTAREM - 1945-ALMEIRIM",
    "14-SANTAREM - 1953-ALPIARCA",
    "14-SANTAREM - 1961-VILA N. DA BARQUINHA",
    "14-SANTAREM - 1970-BENAVENTE",
    "14-SANTAREM - 1988-CARTAXO",
    "14-SANTAREM - 1996-CHAMUSCA",
    "14-SANTAREM - 2003-CONSTANCIA",
    "14-SANTAREM - 2011-CORUCHE",
    "14-SANTAREM - 2020-ENTRONCAMENTO",
    "14-SANTAREM - 2038-FERREIRA DO ZEZERE",
    "14-SANTAREM - 2046-GOLEGA",
    "14-SANTAREM - 2054-MACAO",
    "14-SANTAREM - 2062-RIO MAIOR",
    "14-SANTAREM - 2070-SALVATERRA DE MAGOS",
    "14-SANTAREM - 2089-SANTAREM",
    "14-SANTAREM - 2097-SARDOAL",
    "14-SANTAREM - 2100-TOMAR",
    "14-SANTAREM - 2119-TORRES NOVAS",
    "14-SANTAREM - 2127-OUREM",
    "15-SETUBAL - 2135-ALCACER DO SAL",
    "15-SETUBAL - 2143-ALCOCHETE",
    "15-SETUBAL - 2151-ALMADA-1",
    "15-SETUBAL - 2160-BARREIRO",
    "15-SETUBAL - 2178-GRANDOLA",
    "15-SETUBAL - 2186-MOITA",
    "15-SETUBAL - 2194-MONTIJO",
    "15-SETUBAL - 2208-PALMELA",
    "15-SETUBAL - 2216-SANTIAGO DO CACEM",
    "15-SETUBAL - 2224-SEIXAL 1",
    "15-SETUBAL - 2232-SETUBAL-1",
    "15-SETUBAL - 2240-SESIMBRA",
    "15-SETUBAL - 2259-SINES",
    "15-SETUBAL - 3212-ALMADA-2. C.PIEDADE",
    "15-SETUBAL - 3409-ALMADA-3. C.CAPARICA",
    "15-SETUBAL - 3530-SETUBAL 2",
    "15-SETUBAL - 3697-SEIXAL-2",
    "16-VIANA DO CASTELO - 2267-ARCOS DE VALDEVEZ",
    "16-VIANA DO CASTELO - 2275-CAMINHA",
    "16-VIANA DO CASTELO - 2283-MELGACO",
    "16-VIANA DO CASTELO - 2291-MONCAO",
    "16-VIANA DO CASTELO - 2305-PAREDES DE COURA",
    "16-VIANA DO CASTELO - 2313-PONTE DA BARCA",
    "16-VIANA DO CASTELO - 2321-PONTE DE LIMA",
    "16-VIANA DO CASTELO - 2330-VALENCA",
    "16-VIANA DO CASTELO - 2348-VIANA DO CASTELO",
    "16-VIANA DO CASTELO - 2356-VILA NOVA CERVEIRA",
    "17-VILA REAL - 2364-ALIJO",
    "17-VILA REAL - 2372-BOTICAS",
    "17-VILA REAL - 2380-CHAVES",
    "17-VILA REAL - 2399-MESAO FRIO",
    "17-VILA REAL - 2402-MONDIM DE BASTO",
    "17-VILA REAL - 2410-MONTALEGRE",
    "17-VILA REAL - 2429-MURCA",
    "17-VILA REAL - 2437-PESO DA REGUA",
    "17-VILA REAL - 2445-RIBEIRA DE PENA",
    "17-VILA REAL - 2453-SABROSA",
    "17-VILA REAL - 2461-SANTA MARTA PENAGUIAO",
    "17-VILA REAL - 2470-VALPACOS",
    "17-VILA REAL - 2488-VILA POUCA DE AGUIAR",
    "17-VILA REAL - 2496-VILA REAL",
    "18-VISEU - 2500-ARMAMAR",
    "18-VISEU - 2518-CARREGAL DO SAL",
    "18-VISEU - 2526-CASTRO DAIRE",
    "18-VISEU - 2534-CINFAES",
    "18-VISEU - 2542-LAMEGO",
    "18-VISEU - 2550-MANGUALDE",
    "18-VISEU - 2569-MOIMENTA DA BEIRA",
    "18-VISEU - 2577-MORTAGUA",
    "18-VISEU - 2585-NELAS",
    "18-VISEU - 2593-OLIVEIRA DE FRADES",
    "18-VISEU - 2607-PENALVA DO CASTELO",
    "18-VISEU - 2615-PENEDONO",
    "18-VISEU - 2623-RESENDE",
    "18-VISEU - 2631-S.JOAO DA PESQUEIRA",
    "18-VISEU - 2640-S.PEDRO DO SUL",
    "18-VISEU - 2658-SANTA COMBA DAO",
    "18-VISEU - 2666-SATAO",
    "18-VISEU - 2674-SERNANCELHE",
    "18-VISEU - 2682-TABUACO",
    "18-VISEU - 2690-TAROUCA",
    "18-VISEU - 2704-TONDELA",
    "18-VISEU - 2712-VILA NOVA DE PAIVA",
    "18-VISEU - 2720-VISEU",
    "18-VISEU - 2739-VOUZELA",
    "19-ANGRA DO HEROISMO - 2747-ANGRA DO HEROISMO",
    "19-ANGRA DO HEROISMO - 2755-CALHETA ( S.JORGE )",
    "19-ANGRA DO HEROISMO - 2763-PRAIA DA VITORIA",
    "19-ANGRA DO HEROISMO - 2771-S.CRUZ DA GRACIOSA",
    "19-ANGRA DO HEROISMO - 2780-VELAS",
    "20-HORTA - 2909-CORVO",
    "20-HORTA - 2917-HORTA",
    "20-HORTA - 2925-LAJES DAS FLORES",
    "20-HORTA - 2933-LAGES DO PICO",
    "20-HORTA - 2941-MADALENA",
    "20-HORTA - 2950-S.ROQUE DO PICO",
    "20-HORTA - 2968-S.CRUZ DAS FLORES",
    "21-PONTA DELGADA - 2976-LAGOA (S. MIGUEL)",
    "21-PONTA DELGADA - 2984-NORDESTE",
    "21-PONTA DELGADA - 2992-PONTA DELGADA",
    "21-PONTA DELGADA - 3000-POVOACAO",
    "21-PONTA DELGADA - 3018-RIBEIRA GRANDE",
    "21-PONTA DELGADA - 3026-VILA FRANCA DO CAMPO",
    "21-PONTA DELGADA - 3034-VILA DO PORTO",
    "22-FUNCHAL - 2798-CALHETA - MADEIRA",
    "22-FUNCHAL - 2801-CAMARA DE LOBOS",
    "22-FUNCHAL - 2810-FUNCHAL-1",
    "22-FUNCHAL - 2828-MACHICO",
    "22-FUNCHAL - 2836-PONTA DO SOL",
    "22-FUNCHAL - 2844-PORTO MONIZ",
    "22-FUNCHAL - 2852-PORTO SANTO",
    "22-FUNCHAL - 2860-RIBEIRA BRAVA",
    "22-FUNCHAL - 2879-S.VICENTE (MADEIRA)",
    "22-FUNCHAL - 2887-SANTA CRUZ (MADEIRA)",
    "22-FUNCHAL - 2895-SANTANA",
    "22-FUNCHAL - 3450-FUNCHAL-2"
]

# Funções de validação
def validar_email(email):
    # Email deve conter @ e ter pelo menos 3 caracteres
    if "@" not in email:
        return False
    partes = email.split("@")
    # Deve ter algo antes e depois do @
    return len(partes) == 2 and len(partes[0]) > 0 and len(partes[1]) > 0

def validar_nif(nif):
    return len(str(nif).replace(" ", "")) == 9 and str(nif).isdigit()

def validar_niss(niss):
    return len(str(niss).replace(" ", "")) == 11 and str(niss).isdigit()

def validar_telemovel(tel):
    tel_clean = str(tel).replace(" ", "")
    return len(tel_clean) == 9 and tel_clean.isdigit()

def validar_iban(iban):
    iban_clean = iban.replace(" ", "")
    # Deve começar com PT50 e ter 25 caracteres no total (PT50 + 21 dígitos)
    if not iban_clean.startswith("PT50"):
        return False
    if len(iban_clean) != 25:
        return False
    # Verificar se os últimos 21 caracteres são dígitos
    return iban_clean[4:].isdigit()

def validar_cc(cc):
    # Sem validação de formato - aceita qualquer texto
    return len(cc.strip()) > 0

# Função para carregar dados existentes do Dropbox
def carregar_dados_dropbox():
    try:
        _, response = dbx.files_download(DROPBOX_FILE_PATH)
        data = response.content
        df = pd.read_excel(BytesIO(data))
        return df
    except:
        # Se o ficheiro não existe, criar DataFrame vazio com as colunas
        colunas = [
            "Nome Completo", "Secção", "Nº Horas/Semana", "E-mail", 
            "Data de Nascimento", "NISS", "NIF", "Documento de Identificação",
            "Validade Documento", "Bairro Fiscal", "Estado Civil", 
            "Nº Titulares", "Nº Dependentes", "Morada",
            "IBAN", "Data de Admissão", "Nacionalidade", "Telemóvel",
            "Data de Registo"
        ]
        return pd.DataFrame(columns=colunas)

# Função para guardar dados no Dropbox
def guardar_dados_dropbox(df):
    try:
        # Converter DataFrame para Excel em memória
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Colaboradores')
        output.seek(0)
        
        # Upload para Dropbox (sobrescreve se existir)
        dbx.files_upload(
            output.read(),
            DROPBOX_FILE_PATH,
            mode=dropbox.files.WriteMode.overwrite
        )
        return True
    except Exception as e:
        st.error(f"Erro ao guardar no Dropbox: {e}")
        return False

# Interface da aplicação
st.title("📋 Registo de Colaboradores")
st.markdown("---")

# Formulário
with st.form("formulario_colaborador"):
    st.subheader("Dados Pessoais")
    
    col1, col2 = st.columns(2)
    
    with col1:
        nome = st.text_input("Nome Completo *", help="Nome completo do colaborador")
        email = st.text_input("E-mail *", help="Email corporativo ou pessoal (deve conter @)")
        data_nascimento = st.date_input("Data de Nascimento *", 
                                        min_value=datetime(1950, 1, 1).date(),
                                        max_value=datetime.now().date(),
                                        help="Formato: dd/mm/aaaa")
        nif = st.text_input("NIF *", max_chars=9, help="9 dígitos")
        niss = st.text_input("NISS *", max_chars=11, help="11 dígitos")
    
    with col2:
        telemovel = st.text_input("Telemóvel *", max_chars=9, help="9 dígitos")
        nacionalidade = st.text_input("Nacionalidade *")
        bairro_fiscal = st.selectbox("Bairro Fiscal *", 
                                     options=BAIRROS_FISCAIS,
                                     help="Serviço de finanças da área de residência")
        doc_identificacao = st.text_input("Documento de Identificação *", 
                                          help="Formato CC: 12345678 0 ZW0 ou 'Passaporte' ou 'Cartão de Residência'")
        validade_doc = st.date_input("Validade do Documento *", help="Formato: dd/mm/aaaa")
    
    st.subheader("Situação Familiar")
    
    col3, col4 = st.columns(2)
    
    with col3:
        estado_civil = st.selectbox("Estado Civil / Nº Titulares *", 
                                    ["Casado 1", "Casado 2", "Não Casado"],
                                    help="Casado 1: único titular casado | Casado 2: ambos titulares | Não Casado")
        num_titulares = st.number_input("Nº Titulares *", min_value=1, max_value=2, value=1,
                                        help="Número de titulares do agregado familiar")
    
    with col4:
        num_dependentes = st.number_input("Nº Dependentes *", min_value=0, value=0,
                                         help="Número de dependentes a cargo")
    
    st.subheader("Morada")
    
    morada = st.text_area("Morada Completa *", 
                         help="Completa com rua, lote, porta, andar, código postal e cidade")
    
    st.subheader("Dados Profissionais")
    
    col5, col6 = st.columns(2)
    
    with col5:
        secao = st.selectbox("Secção *", 
                            options=[
                                "Charcutaria/Lacticínios",
                                "Frente de Loja",
                                "Frutas e Vegetais",
                                "Gerência",
                                "Não Perecíveis (reposição)",
                                "Padaria e Take Away",
                                "Peixaria",
                                "Quiosque",
                                "Talho"
                            ],
                            help="Departamento ou secção do colaborador")
        horas_semana = st.selectbox("Nº Horas/Semana *", [16, 20, 40],
                                    help="Horas de trabalho semanais")
        data_admissao = st.date_input("Data de Admissão *", help="Formato: dd/mm/aaaa")
    
    with col6:
        iban = st.text_input("IBAN *",
                            max_chars=25,
                            placeholder="PT50 0000 0000 0000 0000 0000 0",
                            help="Formato: PT50 seguido de 21 dígitos (25 caracteres no total)")
    
    st.markdown("---")
    st.caption("* Campos obrigatórios")
    
    submitted = st.form_submit_button("✅ Submeter Registo", use_container_width=True)
    
    if submitted:
        # Validações
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
            erros.append("Documento de identificação é obrigatório")
        if not iban or not validar_iban(iban):
            erros.append("IBAN deve estar no formato PT50 seguido de 21 dígitos")
        if not morada or len(morada) < 10:
            erros.append("Morada completa é obrigatória")
        if not nacionalidade:
            erros.append("Nacionalidade é obrigatória")
        
        if erros:
            st.error("Por favor corrija os seguintes erros:")
            for erro in erros:
                st.error(f"• {erro}")
        else:
            # Preparar dados
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
            
            # Carregar dados existentes
            with st.spinner("A guardar..."):
                df = carregar_dados_dropbox()
                
                # Adicionar novo registo
                df = pd.concat([df, pd.DataFrame([novo_registo])], ignore_index=True)
                
                # Guardar no Dropbox
                if guardar_dados_dropbox(df):
                    st.success("✅ Registo guardado com sucesso!")
                    st.balloons()
                    st.info(f"Total de colaboradores registados: {len(df)}")
                else:
                    st.error("❌ Erro ao guardar o registo. Tente novamente.")

# Rodapé
st.markdown("---")
st.caption("Formulário de Registo de Colaboradores | Dados guardados de forma segura no Dropbox")
