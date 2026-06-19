import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# --- Senha ---
SENHA_CORRETA = "RodobrasIA"
senha_input = st.text_input("Digite a senha para acessar o app:", type="password")

# Verifica se a senha está correta
if senha_input != SENHA_CORRETA:
    st.warning("🔒 Senha incorreta ou não informada!")
    st.stop()  # Para a execução do app até a senha correta ser digitada

# --- Configuração da página ---
st.set_page_config(layout="wide", page_title="Conciliação de Fretes - Rodobras")
st.title("Conciliação de Fretes - Rodobras")

# --- Função para converter DataFrame em Excel ---
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Conciliação')
    processed_data = output.getvalue()
    return processed_data

# --- Upload do arquivo TMS ---
uploaded_file = st.file_uploader("Escolha o arquivo exportado do TMS", type=["xlsx", "csv"])

if uploaded_file:
    # --- Ler TMS ---
    if uploaded_file.name.endswith(".xlsx"):
        df = pd.read_excel(uploaded_file, header=None)
    else:
        df = pd.read_csv(uploaded_file, header=None)

    # --- Limpeza inicial ---
    df = df.iloc[1:].reset_index(drop=True)
    df = df[~df.iloc[:, 8].astype(str).str.upper().str.contains("CORTESIA")]

    # --- Excluir colunas desnecessárias ---
    colunas_para_excluir = [0,1,4,5,7,8,9,10,11,12,14,15,16,17,18,20,21,22,23,26,27,28,29,30,31,32,33,34,36,39,47,48,49,51,53,54,56,57,58,59,60,61,62,63,64,65,66,67,68,69,70,72,74,75,76,77,78,79,80,81,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97,98,99,100,101,102,103,104,105,106,107,108,109,112,113,114,115,116,117,118,119,120,121,122,123,124,125,126,127,128,129,130,131,132,133,134,135,136,137,138]
    
    colunas_para_excluir = [i for i in colunas_para_excluir if i < len(df.columns)]
    df.drop(df.columns[colunas_para_excluir], axis=1, inplace=True)

    # --- Mover coluna 110 para início ---
    if 110 in df.columns:
        col_110 = df.pop(110)
        df.insert(0, 110, col_110)

    # --- Renomear colunas ---
    nova_ordem_colunas = ["SHIPMENT","CTRC/SUBCON","NUMERO CT-e","DATA DE AUTORIZACAO","LOGIN","CNPJ DESTINATARIO","CIDADE DESTINO","UF","TIPO DE MERCADORIA","NFISCAL","QTDE VOL","PESO REAL","M3","PESO CUBADO","PESO CALC","FRETE PESO","ADEVALOREM","DEDICADO/PARADA","DESCARGA","PEDAGIO","IMPOSTOS","VAL FRETE/COMIS","TABELA DE CALCULO","VAL.MERCADORIA","OBSERVACAO 2" ]
    
    df.columns = nova_ordem_colunas[:len(df.columns)]
    df = df.iloc[1:].reset_index(drop=True)
    df = df.dropna(how="all").reset_index(drop=True)

    # --- Converter colunas numéricas ---
    num_cols = ["PESO CALC", "VAL.MERCADORIA", "FRETE PESO", "ADEVALOREM", "DESCARGA", "PEDAGIO"]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        else:
            df[col] = 0.0

    # --- Ler Tabela MDLZ ---
    caminho_mdlz = "https://raw.githubusercontent.com/AlexsandroQueiroz/CMDLZ/main/Tabela_MDLZ.xlsx"
    df_mdlz = pd.read_excel(caminho_mdlz, sheet_name="Tabela")
    df_mdlz.columns = df_mdlz.columns.str.strip()
    df_mdlz[["TIPO DE OFERTA","TIPO DE CARGA","CIDADE","UF"]] = df_mdlz[["TIPO DE OFERTA","TIPO DE CARGA","CIDADE","UF"]].apply(lambda x: x.astype(str).str.strip().str.upper())
    df["CIDADE DESTINO"] = df["CIDADE DESTINO"].astype(str).str.strip().str.upper()
    df["UF"] = df["UF"].astype(str).str.strip().str.upper()

    # --- Mapear tipo de mercadoria ---
    def map_tipo_mercadoria(tipo):
        tipo = str(tipo).upper()
        if "SUPERFRIO/ARCOR CARRETA REFRIGERADA" in tipo:
            return "CARRETA","REEFER"
        elif "SUPERFRIO/ARCOR CARRETA SECA" in tipo:
            return "CARRETA","DRY"
        elif tipo == "REFRIGERADA":
            return "FRACIONADO","REEFER"
        elif tipo == "CARGA SECA":
            return "FRACIONADO","DRY"
        else:
            return "",""
    df[["TIPO_OFERTA","TIPO_CARGA"]] = df["TIPO DE MERCADORIA"].apply(lambda x: pd.Series(map_tipo_mercadoria(x)))

    # --- Ler planilha online ---
    url_online = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRdZRJ5YctVpKyawelp6CT1ZEkqIAbkqjyRh9DBElof0X0hadYs9ujvKgdqUanPFg/pub?output=csv"
    df_online = pd.read_csv(url_online, header=0)
    df_online.columns = df_online.columns.str.strip()
    def extrair_tipo_carga(valor):    valor = str(valor).upper()
    if "REEFER" in valor:
        return "REEFER"

    if "DRY" in valor:
        return "DRY"

    return ""

    df_online["TIPO_CARGA_ONLINE"] = (
    df_online.iloc[:, 9]
    .apply(extrair_tipo_carga)
)
    df_online["TIPO_OFERTA_ONLINE"] = df_online.iloc[:, 24].str.upper()
    df_online["Tabela_correta"] = df_online["TIPO_OFERTA_ONLINE"].str.capitalize() + " + " + df_online["TIPO_CARGA_ONLINE"].str.capitalize()
    df.insert(2, "Tabela usada", df["TIPO_OFERTA"].str.capitalize() + " + " + df["TIPO_CARGA"].str.capitalize())
    df["SHIPMENT"] = df["SHIPMENT"].astype(str).str.strip()
    df_online["Nº Buy Shipment"] = df_online["Nº Buy Shipment"].astype(str).str.strip()
    df_online_aux = df_online[["Nº Buy Shipment", "Peso", "Frete Sem Impostos", "Tabela_correta"]].copy()
    df_online_aux = df_online_aux.rename(columns={"Nº Buy Shipment": "SHIPMENT","Peso": "PESO_OFER","Frete Sem Impostos": "FRETE_OFER"})
    df = df.merge(df_online_aux, on="SHIPMENT", how="left")

    # --- Extrair tabela correta da oferta Mondelez ---
    df["TIPO_OFERTA_CORRETO"] = (
        df["Tabela_correta"]
        .fillna("")
        .str.split(" + ", regex=False)
        .str[0]
        .str.upper()
    )

    df["TIPO_CARGA_CORRETO"] = (
        df["Tabela_correta"]
        .fillna("")
        .str.split(" + ", regex=False)
        .str[1]
        .str.upper()
    )

    df["DIV_TABELA"] = np.where(
        df["Tabela usada"] == df["Tabela_correta"],
        "OK",
        "ERRO"
    )


    # --- Função calcular frete ---
    colunas_peso = [("V_0A1000",0,1000),("V_1000A3500",1001,3500),("V_3500A5000",3501,5000),("V_5000A7000",5001,7000),("V_7000A9000",7001,9000),("V_9000A11000",9001,11000),("V_11000A13000",11001,13000),("V_13000A15000",13001,15000),("V_15000A100000",15001,100000)]
    def calcular_frete(row):
        oferta, carga, peso, cidade, uf = row["TIPO_OFERTA_CORRETO"], row["TIPO_CARGA_CORRETO"], row["PESO CALC"], row["CIDADE DESTINO"], row["UF"]
        filtro = df_mdlz[(df_mdlz["TIPO DE OFERTA"]==oferta)&(df_mdlz["TIPO DE CARGA"]==carga)&(df_mdlz["CIDADE"]==cidade)&(df_mdlz["UF"]==uf)]
        if filtro.empty: return np.nan
        if oferta=="CARRETA" or peso<=1000: return round(filtro.iloc[0]["V_0A1000"],2)
        for col,min_p,max_p in colunas_peso:
            if min_p<=peso<=max_p: return round((peso/1000)*filtro.iloc[0][col],2)
        return np.nan

    # --- CNPJs especiais ---
    cnpj_especiais = {
        "47960950168901": {"DEDICADO": 0.0, "PE": 149.22},
        "79379491002399": {"DEDICADO": 925.45, "PE": 0.0},
        "82647165000629": {"DEDICADO": 0.0, "PE": 149.22},
        "76430438004673": {"DEDICADO": 925.45, "PE": 0.0},
        "16417108000114": {"DEDICADO": 0.0, "PE": 149.22},
        "76062488000739": {"DEDICADO": 0.0, "PE": 149.22},
        "83646604001370": {"DEDICADO": 0.0, "PE": 149.22},
        "13495487000253": {"DEDICADO": 925.45, "PE": 149.22},
        "45453214002286": {"DEDICADO": 0.0, "PE": 149.22},
        "61940292000218": {"DEDICADO": 925.45, "PE": 0.0},
        "83261420001554": {"DEDICADO": 0.0, "PE": 149.22},
        "82647165003806": {"DEDICADO": 0.0, "PE": 149.22},
        "82647165004020": {"DEDICADO": 0.0, "PE": 149.22},
        "56228356019070": {"DEDICADO": 0.0, "PE": 149.22},
        "11517841004931": {"DEDICADO": 925.45, "PE": 0.0},
        "45543915098998": {"DEDICADO": 0.0, "PE": 149.22},
        "83646984006906": {"DEDICADO": 0.0, "PE": 149.22},
        "61585865065440": {"DEDICADO": 925.45, "PE": 0.0},
        "56228356005958": {"DEDICADO": 0.0, "PE": 149.22},
        "56228356015910": {"DEDICADO": 0.0, "PE": 149.22},
        "02831172001880": {"DEDICADO": 0.0, "PE": 149.22},
        "82647165002672": {"DEDICADO": 0.0, "PE": 149.22},
    }

    # --- Calcular frete correto ---
    df["FRETE_CORRETO"] = df.apply(calcular_frete, axis=1)

    # --- Inicializar DED/PAR_CORR ---
    def calcular_dedicado(row):
        if row["TIPO_OFERTA_CORRETO"] != "FRACIONADO":
            return 0.0

        cnpj = str(row["CNPJ DESTINATARIO"]).strip().replace("\xa0", "")

        return cnpj_especiais.get(
        cnpj,
        {}
    ).get("DEDICADO", 0.0)

    df["DED/PAR_CORR"] = df.apply(calcular_dedicado, axis=1)


    # --- Identificar Multistop ---
    df["SHIP_PREFIX"] = df["SHIPMENT"].str[:-2]
    df["SHIP_SUFFIX"] = df["SHIPMENT"].str[-2:]
    df["MULTISTOP"] = False
    df.loc[df["SHIP_SUFFIX"].isin(["13","14"]), "MULTISTOP"] = True
    prefixos_13 = df.loc[df["SHIP_SUFFIX"]=="13", "SHIP_PREFIX"].unique()
    df.loc[(df["SHIP_SUFFIX"]=="12") & (df["SHIP_PREFIX"].isin(prefixos_13)), "MULTISTOP"] = True
    df["MT"] = np.where(df["MULTISTOP"], "S", "N")

    # --- Rateio Multistop ---
    df_multistop = df[df["MULTISTOP"]].copy()
    if not df_multistop.empty:
        for prefix in df_multistop["SHIP_PREFIX"].unique():
            grupo_idx = df[df["SHIP_PREFIX"] == prefix].index
            grupo = df.loc[grupo_idx].copy().sort_values("SHIPMENT")

            grupo["PESO_BASE"] = np.maximum(
                grupo["PESO REAL"],
                grupo["M3"] * (312*(grupo["TIPO_CARGA"]=="REEFER") + 300*(grupo["TIPO_CARGA"]=="DRY"))
            )
            total_peso = grupo["PESO_BASE"].sum()
            frete_faixa = calcular_frete(grupo.iloc[0])
            grupo["FRETE_RATEIO"] = (grupo["PESO_BASE"]/total_peso) * frete_faixa

            VALOR_PARADA = 1041.13
            grupo["PARADA_TOTAL"] = [0] + [VALOR_PARADA]*(len(grupo)-1)
            grupo["DED/PAR_CORR"] = (grupo["PESO_BASE"]/total_peso) * grupo["PARADA_TOTAL"].sum()

            df.loc[grupo_idx, ["FRETE_CORRETO","DED/PAR_CORR"]] = grupo[["FRETE_RATEIO","DED/PAR_CORR"]].values

    df.drop(columns=["SHIP_PREFIX","SHIP_SUFFIX"], inplace=True)

    # --- PALLET_CORR e TEM_PE? ---
    df["PALLET_CORR"] = df.apply(lambda x: round(x["FRETE_CORRETO"]*cnpj_especiais.get(str(x["CNPJ DESTINATARIO"]).strip().replace("\xa0",""), {}).get("PE",0.0)/1000,2), axis=1)
    df["TEM_PE?"] = np.where(df["PALLET_CORR"]>0,"SIM","NAO")

    # --- Divergências ---
    df["DIV_DED/PAR"] = np.where(abs(df["DEDICADO/PARADA"]-df["DED/PAR_CORR"])<=0.05,"OK","ERRO")
    df["DESCARGA_CORR"] = (df["PESO CALC"]*0.0589828749351323).round(2)
    df["ADEVALOREM_CORR"] = (df["VAL.MERCADORIA"]*0.0003).round(2)
    df["PEDAGIO_CORR"] = df.apply(lambda row: round(row["PESO CALC"]*df_mdlz.loc[(df_mdlz["TIPO DE OFERTA"]=="FRACIONADO") &
                                                                                 (df_mdlz["TIPO DE CARGA"]==row["TIPO_CARGA_CORRETO"]) &
                                                                                 (df_mdlz["CIDADE"]==row["CIDADE DESTINO"]) &
                                                                                 (df_mdlz["UF"]==row["UF"])].iloc[0]["PEDAGIO"],2)
                                  if row["TIPO_OFERTA_CORRETO"]=="FRACIONADO" and not df_mdlz.loc[(df_mdlz["TIPO DE OFERTA"]=="FRACIONADO") &
                                                                                         (df_mdlz["TIPO DE CARGA"]==row["TIPO_CARGA_CORRETO"]) &
                                                                                         (df_mdlz["CIDADE"]==row["CIDADE DESTINO"]) &
                                                                                         (df_mdlz["UF"]==row["UF"])].empty else 0.0, axis=1)
    df["DIV_FRETE"] = np.where(abs(df["FRETE PESO"]-df["FRETE_CORRETO"])<=0.05,"OK","ERRO")
    df["DIV_DESCARGA"] = np.where(abs(df["DESCARGA"]-df["DESCARGA_CORR"])<=0.05,"OK","ERRO")
    df["DIV_ADEVALOREM"] = np.where(abs(df["ADEVALOREM"]-df["ADEVALOREM_CORR"])<=0.05,"OK","ERRO")
    df["DIV_PEDAGIO"] = np.where(abs(df["PEDAGIO"]-df["PEDAGIO_CORR"])<=0.05,"OK","ERRO")   
    
    # --- Formatando data ---
    df["DATA DE AUTORIZACAO"] = pd.to_datetime(df["DATA DE AUTORIZACAO"], errors="coerce").dt.strftime("%d/%m/%Y")

    # --- Colunas finais ---
    df["PESO_CT-e"] = df["PESO REAL"]
    colunas_validas = [
        "SHIPMENT","MT","CTRC/SUBCON","Tabela usada","Tabela_correta","DIV_TABELA","DATA DE AUTORIZACAO",
        "CIDADE DESTINO","UF","PESO_OFER","PESO_CT-e","FRETE_OFER","FRETE PESO","FRETE_CORRETO",
        "DIV_FRETE","DESCARGA","DESCARGA_CORR","DIV_DESCARGA","ADEVALOREM","ADEVALOREM_CORR",
        "DIV_ADEVALOREM","PEDAGIO","PEDAGIO_CORR","DIV_PEDAGIO","DEDICADO/PARADA","DED/PAR_CORR",
        "DIV_DED/PAR","TEM_PE?","PALLET_CORR"
    ]   
    df_conciliacao = df[colunas_validas].copy()
    df_conciliacao.rename(columns={"PALLET_CORR":"PALLET_CORR($)","DED/PAR_CORR":"DED/PAR_CORR($)"}, inplace=True)

    # --- Filtros para limpar dados indesejados ---

    # 1️⃣ Manter apenas Shipments que começam com '8'
    if "SHIPMENT" in df_conciliacao.columns:
        df_conciliacao = df_conciliacao[df_conciliacao["SHIPMENT"].astype(str).str.startswith("8")]

    # 2️⃣ Manter apenas Tabelas Usadas válidas
    if "Tabela usada" in df_conciliacao.columns:
        tabelas_validas = [
            "Fracionado + Reefer",
            "Carreta + Reefer",
            "Fracionado + Dry",
            "Carreta + Dry"
        ]
        df_conciliacao = df_conciliacao[df_conciliacao["Tabela usada"].isin(tabelas_validas)]

    # Resetar o índice após filtrar
    df_conciliacao.reset_index(drop=True, inplace=True)


    # --- Garantir 2 casas decimais para colunas de valores monetários ---
    for col in ["DED/PAR_CORR($)", "PALLET_CORR($)"]:
        if col in df_conciliacao.columns:
            df_conciliacao[col] = df_conciliacao[col].round(2)
            df_conciliacao[col] = df_conciliacao[col].apply(lambda x: f"{x:.2f}" if pd.notnull(x) else "0.00")

        # --- Estilização e formatação ---
    def colorir_divergencias(val):
        if val == "ERRO":
            return "background-color: #ff4d4d; color: white;"
        elif val == "OK":
            return "background-color: #4CAF50; color: white;"
        return ""

    # Colunas numéricas a formatar
    colunas_formatar = [
        "PESO_CT-e","FRETE PESO","FRETE_CORRETO","FRETE_OFER",
        "DESCARGA","DESCARGA_CORR","ADEVALOREM","ADEVALOREM_CORR",
        "PEDAGIO","PEDAGIO_CORR","DEDICADO/PARADA","DED/PAR_CORR","PALLET_CORR($)"
    ]

    # Função para limpar e converter valores
    def limpar_valor(val):
        val = str(val).strip()
        if val=="" or val.lower()=="nan":
            return np.nan
        if "," in val:
            val = val.replace(".","").replace(",",".")
        else:
            val = val.replace(",","")
        try:
            return float(val)
        except:
            return np.nan

    # Aplicar limpeza e formatação
    for col in colunas_formatar:
        if col in df_conciliacao.columns:
            df_conciliacao[col] = df_conciliacao[col].apply(limpar_valor)
            if col == "FRETE_OFER":
                df_conciliacao[col] = df_conciliacao[col] / 0.9635  # ajuste específico
            df_conciliacao[col] = df_conciliacao[col].apply(lambda x: f"{x:.2f}" if pd.notnull(x) else "0.00")


    # --- Exibir DataFrame com estilo ---
    st.dataframe(
        df_conciliacao.style.map(
            colorir_divergencias,
            subset=["DIV_TABELA","DIV_FRETE","DIV_DESCARGA","DIV_ADEVALOREM","DIV_PEDAGIO","DIV_DED/PAR"]
        ),
        use_container_width=True
    )

    # --- Botão de download ---
    excel_data = to_excel(df_conciliacao)
    st.download_button(
        label="📥 Baixar base em Excel",
        data=excel_data,
        file_name="Conciliação_Fretes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("👆 Faça o upload de um arquivo TMS para iniciar a conciliação.")           