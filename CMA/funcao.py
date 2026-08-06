from pathlib import Path
import pandas as pd
import sys
import gspread
from google.oauth2.service_account import Credentials

# ==========================================
# COLUNAS ESPERADAS
# ==========================================

colunas_hm = [
    "þÿNr atendimento",
    "Nm pac",
    "Nasc",
    "Inicio real",
    "Fim real",
    "Resp cred cir"
]

colunas_processados = [
    "Data/Hora",
    "AMB",
    "Procedimento",
    "Nome medico",
    "Anestesista",
    "Nome convenio",
    "Anestesista cma ?",
    "Atendimento"
]

# ==========================================
# LEITURA
# ==========================================

def ler_planilhas(arquivo_hm, arquivo_processado):

    try:

        df_hm = pd.read_excel(arquivo_hm)
        df_processado = pd.read_excel(arquivo_processado)

    except Exception as erro:
        raise Exception(f"Erro ao abrir as planilhas.\n{erro}")

    return df_hm, df_processado


# ==========================================
# VALIDAÇÃO
# ==========================================

def validar_colunas(df, colunas_esperadas, nome):

    faltando = [
        coluna
        for coluna in colunas_esperadas
        if coluna not in df.columns
    ]

    if faltando:
        raise ValueError(
            f"A planilha '{nome}' não possui as colunas:\n\n{faltando}"
        )


# ==========================================
# FILTRO
# ==========================================

def filtrar(df, coluna, valor):

    return df[
        df[coluna]
        .astype(str)
        .str.strip()
        .str.upper()
        == valor.upper()
    ]


# ==========================================
# PADRONIZAÇÃO
# ==========================================

def padronizar_chave(df, coluna):

    df[coluna] = (
        df[coluna]
        .astype(str)
        .str.strip()
    )

    return df


# ==========================================
# MERGE
# ==========================================

def unir_planilhas(df_hm, df_processado):

    return pd.merge(
        df_hm,
        df_processado,
        left_on="þÿNr atendimento",
        right_on="Atendimento",
        how="inner"
    )


# ==========================================
# PLANILHA FINAL
# ==========================================

def construir_planilha_final(df):

    mapa = {
            "Data/Hora": "Data/Hora",
            "Inicio real": "Inicio real",
            "Fim real": "Fim real",
            "þÿNr atendimento": "Atendimento",
            "Nm pac": "Paciente",
            "Nasc": "Nascimento",
            "Nome convenio": "Nome convenio",
            "AMB": "AMB",
            "Procedimento": "Procedimento",
            "Nome medico": "Nome medico",
            "Anestesista": "Anestesista",
            "Resp cred cir": "Cirurgião",
            "Anestesista cma ?": "Staff"
    }

    return (
        df[list(mapa.keys())]
        .rename(columns=mapa)
        .copy()
    )

# ==========================================
# LISTA STAFF
# ==========================================

def recurso(nome):
    if getattr(sys, "frozen", False):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).parent

    return base / nome

def carregar_lista_anestesistas():

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly"
    ]

    caminho_json = recurso("credenciais.json")

    credenciais = Credentials.from_service_account_file(
        caminho_json,
        scopes=scopes
    )

    cliente = gspread.authorize(credenciais)

    planilha = cliente.open_by_key(
        "1I4SdxCwAmfWd8us0Hd3UZ8fRPX6MjI98dV30RAcWpAI"
    )

    aba = planilha.worksheet("ANESTESISTAS")

    dados = aba.get_all_records()

    df = pd.DataFrame(dados)

    coluna_nome = "NOME COMPLETO"

    if coluna_nome not in df.columns:
        raise Exception(
            f"Coluna '{coluna_nome}' não encontrada."
        )

    df[coluna_nome] = (
        df[coluna_nome]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    return df, coluna_nome

def padronizar_anestesistas(df):

    df["Anestesista"] = (
        df["Anestesista"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    return df


def filtrar_anestesistas(df, lista, coluna_nome):

    return df[
        df["Anestesista"].isin(lista[coluna_nome])
    ]
# ==========================================
# SALVAR
# ==========================================

def salvar_planilha(df, arquivo_hm):

    arquivo_hm = Path(arquivo_hm)

    nome_saida = (
        arquivo_hm.parent /
        f"{arquivo_hm.stem}_FINAL.xlsx"
    )

    df.to_excel(nome_saida, index=False)

    return nome_saida


# ==========================================
# FUNÇÃO PRINCIPAL
# ==========================================

def gerar_relatorio(arquivo_hm, arquivo_processado):

    # Lê as planilhas escolhidas pelo usuário
    df_hm, df_processado = ler_planilhas(
        arquivo_hm,
        arquivo_processado
    )

    validar_colunas(df_hm, colunas_hm, "HM")
    validar_colunas(df_processado, colunas_processados, "Processado")

    df_processado = filtrar(
        df_processado,
        "Anestesista cma ?",
        "Staff"
    )

    df_hm = padronizar_chave(
        df_hm,
        "þÿNr atendimento"
    )

    df_processado = padronizar_chave(
        df_processado,
        "Atendimento"
    )

    df_merge = unir_planilhas(
        df_hm,
        df_processado
    )

    df_final = construir_planilha_final(
        df_merge
    )
    
    df_final = df_final.drop_duplicates(
        
        subset=["Atendimento"],
        keep="first"
    )

    lista, coluna_nome = carregar_lista_anestesistas()

    df_final = padronizar_anestesistas(df_final)

    df_final = filtrar_anestesistas(
        df_final,
        lista,
        coluna_nome
    )
    
    arquivo_saida = salvar_planilha(
        df_final,
        Path(arquivo_hm)
    )

    return f"Relatório criado!\n{arquivo_saida.name}"