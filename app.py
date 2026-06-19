# -*- coding: utf-8 -*-
"""
Painel de Vendas - Controle de prospecção e vendas da equipe
Banco de dados: Google Sheets (no Google Drive)
"""

from datetime import datetime
from zoneinfo import ZoneInfo

import io

import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

# ----------------------------------------------------------------------------
# Configurações gerais
# ----------------------------------------------------------------------------
TZ = ZoneInfo("America/Sao_Paulo")

ABA_REGISTROS  = "Registros"
ABA_VENDEDORES = "Vendedores"
ABA_CLIENTES   = "Clientes"
ABA_CARGAS     = "Cargas"
ABA_DIARIA     = "Análise Diária"
ABA_MENSAL     = "Análise Mensal"

CABECALHO_REGISTROS = [
    "Data", "Hora", "Vendedor", "Cliente", "Tipo cliente", "Com quem falou",
    "Resultado", "Situação", "Kg", "Valor (R$)", "R$/kg", "Carga",
    "Contato do Cliente", "Cliente Recorrente?", "Motivo da Perda",
    "Data Resposta Orçamento",
]

CABECALHO_CARGAS   = ["Carga", "Data Entrega", "Vendedor", "Meta (Kg)"]
CABECALHO_CLIENTES = ["Cliente", "Cadastrado em", "Cadastrado por"]

SITUACAO_ABERTO   = "Em aberto"
SITUACAO_APROVADO = "Aprovado"
SITUACAO_PERDIDO  = "Perdido"

RESULTADOS = ["Só contato", "Orçamento enviado", "Venda fechada"]

MOTIVOS_PERDA = [
    "Preço",
    "Concorrente entregou antes",
    "Cliente achou qualidade baixa",
    "Sem resposta / desistiu",
    "Outro",
]

VENDEDORES_INICIAIS = [
    ["ANA PAULA", "1010", "SIM"],
    ["CAIO",      "2020", "SIM"],
    ["VANDERLEI", "3030", "SIM"],
    ["JESUS",     "4040", "SIM"],
    ["JONATAN",   "5050", "SIM"],
    ["MARCIO",    "6060", "SIM"],
    ["RENATA",    "7070", "SIM"],
]

st.set_page_config(page_title="Painel de Vendas", page_icon="📊", layout="wide")


# ----------------------------------------------------------------------------
# Utilidades
# ----------------------------------------------------------------------------
def br(valor, casas=2):
    if valor is None or pd.isna(valor):
        return "-"
    s = f"{valor:,.{casas}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def agora():
    return datetime.now(TZ)


# ----------------------------------------------------------------------------
# Conexão com o Google Sheets
# ----------------------------------------------------------------------------
@st.cache_resource
def abrir_planilha():
    escopos = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=escopos
    )
    gc = gspread.authorize(creds)
    sh = gc.open_by_url(st.secrets["app"]["sheet_url"])
    garantir_estrutura(sh)
    return sh


def _adicionar_colunas_faltantes(ws, cabecalho_esperado):
    """Adiciona ao sheet as colunas presentes em cabecalho_esperado mas ausentes na planilha."""
    atual = ws.row_values(1)
    for col in cabecalho_esperado:
        if col not in atual:
            ws.update_cell(1, len(atual) + 1, col)
            atual.append(col)


def garantir_estrutura(sh):
    titulos = [ws.title for ws in sh.worksheets()]

    if ABA_REGISTROS not in titulos:
        ws = sh.add_worksheet(title=ABA_REGISTROS, rows=2000, cols=20)
        ws.update(values=[CABECALHO_REGISTROS], range_name="A1")
        ws.format("A1:P1", {"textFormat": {"bold": True}})
    else:
        ws = sh.worksheet(ABA_REGISTROS)
        # Sobrescreve o cabeçalho para corrigir colunas extras ou corrompidas
        ws.update('A1', [CABECALHO_REGISTROS])

    if ABA_CLIENTES not in titulos:
        ws = sh.add_worksheet(title=ABA_CLIENTES, rows=2000, cols=5)
        ws.update(values=[CABECALHO_CLIENTES], range_name="A1")
        ws.format("A1:C1", {"textFormat": {"bold": True}})

    if ABA_VENDEDORES not in titulos:
        ws = sh.add_worksheet(title=ABA_VENDEDORES, rows=50, cols=5)
        ws.update(
            values=[["Nome", "PIN", "Ativo"]] + VENDEDORES_INICIAIS,
            range_name="A1",
        )
        ws.format("A1:C1", {"textFormat": {"bold": True}})

    if ABA_CARGAS not in titulos:
        ws = sh.add_worksheet(title=ABA_CARGAS, rows=200, cols=5)
        ws.update(values=[CABECALHO_CARGAS], range_name="A1")
        ws.format("A1:D1", {"textFormat": {"bold": True}})

    if ABA_DIARIA not in titulos:
        sh.add_worksheet(title=ABA_DIARIA, rows=100, cols=15)

    if ABA_MENSAL not in titulos:
        sh.add_worksheet(title=ABA_MENSAL, rows=500, cols=15)


# ----------------------------------------------------------------------------
# Cargas
# ----------------------------------------------------------------------------
@st.cache_data(ttl=60)
def carregar_cargas():
    sh = abrir_planilha()
    titulos = [ws.title for ws in sh.worksheets()]
    if ABA_CARGAS not in titulos:
        ws = sh.add_worksheet(title=ABA_CARGAS, rows=200, cols=5)
        ws.update(values=[CABECALHO_CARGAS], range_name="A1")
        ws.format("A1:D1", {"textFormat": {"bold": True}})
        return pd.DataFrame(columns=CABECALHO_CARGAS + ["_linha"])
    dados = sh.worksheet(ABA_CARGAS).get_all_records(
        expected_headers=CABECALHO_CARGAS
    )
    df = pd.DataFrame(dados) if dados else pd.DataFrame(columns=CABECALHO_CARGAS)
    df["_linha"] = df.index + 2
    df["Meta (Kg)"] = pd.to_numeric(df["Meta (Kg)"], errors="coerce").fillna(0.0)
    return df


def salvar_carga(nome, data_entrega, vendedor, meta_kg):
    sh = abrir_planilha()
    sh.worksheet(ABA_CARGAS).append_row(
        [nome.strip(), data_entrega, vendedor, round(float(meta_kg), 2)],
        value_input_option="RAW",
    )
    carregar_cargas.clear()


def deletar_carga(linha_planilha):
    sh = abrir_planilha()
    sh.worksheet(ABA_CARGAS).delete_rows(int(linha_planilha))
    carregar_cargas.clear()


def progresso_carga(carga_nome, df_registros, vendedor=None):
    """Kg de vendas fechadas vinculadas a esta carga (e opcionalmente a um vendedor)."""
    mask = (
        (df_registros["Carga"] == carga_nome) &
        (df_registros["Resultado"] == "Venda fechada")
    )
    if vendedor:
        mask &= (df_registros["Vendedor"] == vendedor)
    return float(df_registros.loc[mask, "Kg"].sum())


# ----------------------------------------------------------------------------
# Vendedores / Clientes / Registros
# ----------------------------------------------------------------------------
@st.cache_data(ttl=300)
def carregar_vendedores():
    sh = abrir_planilha()
    dados = sh.worksheet(ABA_VENDEDORES).get_all_records()
    ativos = {}
    for linha in dados:
        nome  = str(linha.get("Nome", "")).strip()
        ativo = str(linha.get("Ativo", "")).strip().upper()
        if nome and ativo in ("SIM", "S", "1", "X"):
            ativos[nome] = str(linha.get("PIN", "")).strip()
    return ativos


@st.cache_data(ttl=60)
def carregar_clientes():
    sh = abrir_planilha()
    dados = sh.worksheet(ABA_CLIENTES).get_all_records(
        expected_headers=CABECALHO_CLIENTES
    )
    nomes, vistos = [], set()
    for linha in dados:
        nome  = str(linha.get("Cliente", "")).strip()
        chave = nome.upper()
        if nome and chave not in vistos:
            vistos.add(chave)
            nomes.append(nome)
    return sorted(nomes, key=str.upper)


def cadastrar_cliente(nome, vendedor):
    sh = abrir_planilha()
    sh.worksheet(ABA_CLIENTES).append_row(
        [nome.strip(), agora().strftime("%d/%m/%Y"), vendedor],
        value_input_option="RAW",
    )
    carregar_clientes.clear()


@st.cache_data(ttl=30)
def carregar_registros():
    sh     = abrir_planilha()
    ws     = sh.worksheet(ABA_REGISTROS)
    # get_all_values evita GSpreadException por cabeçalhos duplicados/vazios
    valores = ws.get_all_values()
    if len(valores) <= 1:
        colunas_extras = ["_data", "_mes", "_linha",
                          "Dias em Aberto", "Tempo até Resposta (dias)"]
        return pd.DataFrame(columns=CABECALHO_REGISTROS + colunas_extras)
    # Usa os cabeçalhos reais da planilha, ignorando duplicatas silenciosamente
    headers = valores[0]
    seen: dict = {}
    headers_uniq = []
    for h in headers:
        h = h.strip()
        if h in seen:
            seen[h] += 1
            headers_uniq.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            headers_uniq.append(h)
    df = pd.DataFrame(valores[1:], columns=headers_uniq)

    colunas_extras = ["_data", "_mes", "_linha", "Dias em Aberto", "Tempo até Resposta (dias)"]
    if df.empty:
        return pd.DataFrame(columns=CABECALHO_REGISTROS + colunas_extras)

    df["_linha"] = df.index + 2

    # Garante colunas novas no DataFrame para registros antigos
    for col in ("Carga", "Contato do Cliente", "Cliente Recorrente?",
                "Motivo da Perda", "Data Resposta Orçamento"):
        if col not in df.columns:
            df[col] = ""

    for col in ("Kg", "Valor (R$)", "R$/kg"):
        if col not in df.columns:
            df[col] = 0.0
        if df[col].dtype == object:
            df[col] = (
                df[col].astype(str)
                .str.replace(".", "", regex=False)
                .str.replace(",", ".", regex=False)
            )
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    df["_data"] = pd.to_datetime(df["Data"], format="%d/%m/%Y", errors="coerce")
    df["_mes"]  = df["_data"].dt.strftime("%m/%Y")

    # Campo calculado: dias em aberto (somente orçamentos em aberto)
    hoje_dt = pd.Timestamp(agora().date())
    df["Dias em Aberto"] = pd.NA
    mask_aberto = (
        (df["Resultado"] == "Orçamento enviado") &
        (df["Situação"] == SITUACAO_ABERTO) &
        df["_data"].notna()
    )
    df.loc[mask_aberto, "Dias em Aberto"] = (hoje_dt - df.loc[mask_aberto, "_data"]).dt.days

    # Campo calculado: tempo até resposta (linhas com Data Resposta preenchida)
    df["Tempo até Resposta (dias)"] = pd.NA
    mask_resp = df["Data Resposta Orçamento"].astype(str).str.strip() != ""
    if mask_resp.any():
        data_resp = pd.to_datetime(
            df.loc[mask_resp, "Data Resposta Orçamento"],
            format="%d/%m/%Y", errors="coerce"
        )
        df.loc[mask_resp, "Tempo até Resposta (dias)"] = (
            data_resp - df.loc[mask_resp, "_data"]
        ).dt.days

    return df


def salvar_registro(vendedor, cliente, cliente_novo, contato,
                    resultado, kg, valor, carga="",
                    contato_cliente="", recorrente=False, motivo_perda=""):
    sh = abrir_planilha()
    ws = sh.worksheet(ABA_REGISTROS)
    # Repara o cabeçalho se estiver corrompido (colunas extras/vazias)
    if ws.row_values(1) != CABECALHO_REGISTROS:
        ws.update('A1', [CABECALHO_REGISTROS])
    ts = agora()
    preco_kg = round(valor / kg, 2) if kg and valor else 0
    if cliente_novo and cliente.strip().upper() not in {
            c.upper() for c in carregar_clientes()}:
        cadastrar_cliente(cliente, vendedor)
    # table_range='A1' ancora a detecção de tabela na coluna A,
    # evitando que o Sheets escreva em colunas erradas por desalinhamento
    ws.append_row(
        [
            ts.strftime("%d/%m/%Y"),                                   # Data
            ts.strftime("%H:%M"),                                      # Hora
            vendedor,                                                  # Vendedor
            cliente.strip(),                                           # Cliente
            "Novo" if cliente_novo else "Carteira",                    # Tipo cliente
            contato.strip(),                                           # Com quem falou
            resultado,                                                 # Resultado
            SITUACAO_ABERTO if resultado == "Orçamento enviado" else "-",  # Situação
            round(float(kg or 0), 2),                                  # Kg
            round(float(valor or 0), 2),                               # Valor (R$)
            preco_kg,                                                  # R$/kg
            carga or "",                                               # Carga
            contato_cliente or "",                                     # Contato do Cliente
            "SIM" if recorrente else "",                               # Cliente Recorrente?
            motivo_perda or "",                                        # Motivo da Perda
            "",                                                        # Data Resposta Orçamento
        ],
        value_input_option="RAW",
        table_range="A1",
    )
    carregar_registros.clear()


def _atualizar_celula(linha_planilha, nome_col, valor):
    sh  = abrir_planilha()
    col = CABECALHO_REGISTROS.index(nome_col) + 1
    sh.worksheet(ABA_REGISTROS).update_cell(int(linha_planilha), col, valor)


def mudar_situacao(linha_planilha, nova):
    _atualizar_celula(linha_planilha, "Situação", nova)
    carregar_registros.clear()


def aprovar_orcamento(reg):
    mudar_situacao(reg["_linha"], SITUACAO_APROVADO)
    _atualizar_celula(reg["_linha"], "Data Resposta Orçamento",
                      agora().strftime("%d/%m/%Y"))
    salvar_registro(
        reg["Vendedor"], reg["Cliente"], False,
        str(reg["Com quem falou"]), "Venda fechada",
        float(reg["Kg"]), float(reg["Valor (R$)"]),
        str(reg.get("Carga", "")),
    )


def perder_orcamento(reg, motivo=""):
    mudar_situacao(reg["_linha"], SITUACAO_PERDIDO)
    _atualizar_celula(reg["_linha"], "Data Resposta Orçamento",
                      agora().strftime("%d/%m/%Y"))
    if motivo:
        _atualizar_celula(reg["_linha"], "Motivo da Perda", motivo)
    carregar_registros.clear()


def deletar_registro(linha_planilha):
    sh = abrir_planilha()
    sh.worksheet(ABA_REGISTROS).delete_rows(int(linha_planilha))
    carregar_registros.clear()


# ----------------------------------------------------------------------------
# Resumos
# ----------------------------------------------------------------------------
COLUNAS_RESUMO = [
    "Vendedor", "Lançamentos", "Clientes novos", "Orçamentos",
    "Vendas", "Kg vendido", "R$ vendido", "R$/kg médio",
    "Conversão (%)", "Taxa Perda (%)", "Recorrentes (%)", "Score",
]


def resumir(df):
    if df.empty:
        return pd.DataFrame(columns=COLUNAS_RESUMO)
    linhas = []
    for vend, g in df.groupby("Vendedor"):
        vendas   = g[g["Resultado"] == "Venda fechada"]
        orcs     = g[g["Resultado"] == "Orçamento enviado"]
        perdidos = g[g["Situação"] == SITUACAO_PERDIDO] if "Situação" in g.columns else pd.DataFrame()
        novos    = g[g["Tipo cliente"] == "Novo"] if "Tipo cliente" in g.columns else pd.DataFrame()
        recorr   = (g[g["Cliente Recorrente?"].astype(str).str.upper() == "SIM"]
                    if "Cliente Recorrente?" in g.columns else pd.DataFrame())
        kg  = vendas["Kg"].sum()
        rs  = vendas["Valor (R$)"].sum()
        # Orçamentos aprovados já viraram venda — excluir do denominador para não duplicar
        orcs_pendentes = orcs[orcs["Situação"] != SITUACAO_APROVADO] if "Situação" in orcs.columns else orcs
        dec_conv  = len(vendas) + len(orcs_pendentes)
        dec_perda = len(vendas) + len(perdidos)
        conv  = 100 * len(vendas) / dec_conv  if dec_conv  > 0 else 0.0
        perda = 100 * len(perdidos) / dec_perda if dec_perda > 0 else 0.0
        recorr_pct = 100 * len(recorr) / len(g) if len(g) > 0 else 0.0
        score = round(conv * 0.5 + max(100 - perda, 0) * 0.3 + recorr_pct * 0.2, 1)
        linhas.append([
            vend, len(g), len(novos), len(orcs), len(vendas),
            round(kg, 2), round(rs, 2),
            round(rs / kg, 2) if kg else 0.0,
            round(conv, 1), round(perda, 1), round(recorr_pct, 1), score,
        ])
    out = pd.DataFrame(linhas, columns=COLUNAS_RESUMO)
    return out.sort_values(["Kg vendido", "R$ vendido", "Lançamentos"],
                           ascending=False).reset_index(drop=True)


def _kpis(df):
    """KPIs consolidados para um período."""
    vendas   = df[df["Resultado"] == "Venda fechada"]
    orcs     = df[df["Resultado"] == "Orçamento enviado"]
    perdidos = df[df["Situação"] == SITUACAO_PERDIDO] if "Situação" in df.columns else pd.DataFrame()
    recorr   = (df[df["Cliente Recorrente?"].astype(str).str.upper() == "SIM"]
                if "Cliente Recorrente?" in df.columns else pd.DataFrame())
    kg  = vendas["Kg"].sum()
    rs  = vendas["Valor (R$)"].sum()
    orcs_pendentes = orcs[orcs["Situação"] != SITUACAO_APROVADO] if "Situação" in orcs.columns else orcs
    dec_conv  = len(vendas) + len(orcs_pendentes)
    dec_perda = len(vendas) + len(perdidos)
    return {
        "n_vendas":   len(vendas),
        "n_orcs":     len(orcs),
        "n_perdidos": len(perdidos),
        "kg":         kg,
        "rs":         rs,
        "ticket":     rs / len(vendas) if len(vendas) > 0 else 0.0,
        "conv":       100 * len(vendas) / dec_conv  if dec_conv  > 0 else 0.0,
        "perda_pct":  100 * len(perdidos) / dec_perda if dec_perda > 0 else 0.0,
        "recorr_pct": 100 * len(recorr) / len(df)  if len(df)   > 0 else 0.0,
    }


def atualizar_abas_analise():
    sh   = abrir_planilha()
    df   = carregar_registros()
    hoje = agora().strftime("%d/%m/%Y")

    df_hoje  = df[df["Data"] == hoje]
    res_dia  = resumir(df_hoje)
    sem_lanc = sorted(set(carregar_vendedores()) - set(res_dia["Vendedor"]))
    ws = sh.worksheet(ABA_DIARIA)
    ws.clear()
    valores  = [[f"ANÁLISE DO DIA {hoje}"], [], COLUNAS_RESUMO]
    valores += res_dia.astype(object).values.tolist()
    valores += [[], ["Sem lançamentos hoje:", ", ".join(sem_lanc) or "ninguém"]]
    ws.update(values=valores, range_name="A1")

    ws = sh.worksheet(ABA_MENSAL)
    ws.clear()
    valores = [["ANÁLISE MENSAL (todos os meses)"], [], ["Mês"] + COLUNAS_RESUMO]
    for mes in sorted(df["_mes"].dropna().unique(), reverse=True):
        res_mes = resumir(df[df["_mes"] == mes])
        for linha in res_mes.astype(object).values.tolist():
            valores.append([mes] + linha)
    ws.update(values=valores, range_name="A1")


# ----------------------------------------------------------------------------
# Inteligência de cargas
# ----------------------------------------------------------------------------
def _intel_carga(data_entrega_str, meta_kg, realizado_kg, df_registros, carga_nome, vendedor=None):
    """Retorna (dias_restantes, forecast_dias, is_overdue, pace_kg_dia)."""
    try:
        data_ent = datetime.strptime(data_entrega_str.strip(), "%d/%m/%Y").date()
    except (ValueError, AttributeError):
        return None, None, False, 0.0
    hoje = agora().date()
    dias_rest = (data_ent - hoje).days
    is_overdue = dias_rest < 0 and realizado_kg < meta_kg
    faltam = max(meta_kg - realizado_kg, 0.0)
    mask = (
        (df_registros["Carga"] == carga_nome) &
        (df_registros["Resultado"] == "Venda fechada") &
        df_registros["_data"].notna()
    )
    if vendedor:
        mask &= (df_registros["Vendedor"] == vendedor)
    vendas_c = df_registros[mask]
    if not vendas_c.empty:
        data_ini = vendas_c["_data"].min().date()
        dias_ativos = max((hoje - data_ini).days, 1)
        pace = realizado_kg / dias_ativos
    else:
        pace = 0.0
    if pace > 0 and faltam > 0:
        forecast = int(faltam / pace)
    elif faltam == 0:
        forecast = 0
    else:
        forecast = None
    return dias_rest, forecast, is_overdue, pace


# ----------------------------------------------------------------------------
# Componentes de tela
# ----------------------------------------------------------------------------
def tabela_ranking(res, destaque=None):
    if res.empty:
        st.info("Nenhum lançamento ainda.")
        return
    res = res.copy()
    res.insert(0, "Posição", [f"{i}º" for i in range(1, len(res) + 1)])
    res["Kg vendido"]      = res["Kg vendido"].map(br)
    res["R$ vendido"]      = res["R$ vendido"].map(lambda v: "R$ " + br(v))
    res["R$/kg médio"]     = res["R$/kg médio"].map(lambda v: "R$ " + br(v))
    res["Conversão (%)"]   = res["Conversão (%)"].map(lambda v: f"{v:.1f}%")
    res["Taxa Perda (%)"]  = res["Taxa Perda (%)"].map(lambda v: f"{v:.1f}%")
    res["Recorrentes (%)"] = res["Recorrentes (%)"].map(lambda v: f"{v:.1f}%")
    res["Score"]           = res["Score"].map(lambda v: f"{v:.1f}")
    st.dataframe(res, hide_index=True, use_container_width=True)
    if destaque is not None and destaque in res["Vendedor"].values:
        pos = res.loc[res["Vendedor"] == destaque, "Posição"].iloc[0]
        st.caption(f"Sua posição: {pos} de {len(res)}")


def widget_cargas(df_cargas, df_registros, filtro_vendedor=None):
    if df_cargas.empty:
        st.info("Nenhuma carga cadastrada ainda.")
        return
    for _, carga in df_cargas.iterrows():
        if filtro_vendedor and carga["Vendedor"] != filtro_vendedor:
            continue
        realizado = progresso_carga(carga["Carga"], df_registros, vendedor=carga["Vendedor"])
        meta      = float(carga["Meta (Kg)"]) or 1.0
        pct       = min(realizado / meta, 1.0)
        cor       = "🟢" if pct >= 1.0 else ("🟡" if pct >= 0.6 else "🔴")
        dias_rest, forecast, is_overdue, pace = _intel_carga(
            carga["Data Entrega"], meta, realizado, df_registros,
            carga["Carga"], vendedor=carga["Vendedor"]
        )
        titulo = f"**{carga['Carga']}** — {carga['Data Entrega']} — Meta: {br(meta)} kg"
        if is_overdue:
            st.error(f"🚨 {titulo} *(prazo vencido!)*")
        else:
            st.markdown(titulo)
        st.progress(pct, text=f"{cor} {br(realizado)} / {br(meta)} kg  ({pct*100:.0f}%)")
        if dias_rest is not None and pct < 1.0:
            if is_overdue:
                st.caption(f"⛔ Prazo vencido há {abs(dias_rest)} dia(s). Faltam {br(meta - realizado)} kg.")
            elif forecast is None:
                st.caption(f"⏳ {dias_rest} dia(s) até entrega — sem histórico para previsão.")
            elif forecast <= dias_rest:
                st.caption(f"✅ No ritmo atual ({br(pace)} kg/dia), fecha em ~{forecast} dia(s). Prazo: {dias_rest} dia(s).")
            else:
                st.caption(f"⚠️ Ritmo atual ({br(pace)} kg/dia) prevê fechamento em {forecast} dia(s), mas prazo é {dias_rest} dia(s).")
        st.divider()


# ----------------------------------------------------------------------------
# Tela Login
# ----------------------------------------------------------------------------
def tela_login():
    st.title("📊 Painel de Vendas")
    vendedores = carregar_vendedores()
    nomes = list(vendedores.keys()) + ["GESTOR"]

    col, _ = st.columns([1, 1])
    with col:
        nome  = st.selectbox("Quem é você?", nomes)
        senha = st.text_input(
            "Senha" if nome == "GESTOR" else "PIN",
            type="password", max_chars=20,
        )
        if st.button("Entrar", type="primary", use_container_width=True):
            ok = (
                senha == st.secrets["app"]["senha_gestor"]
                if nome == "GESTOR"
                else senha == vendedores.get(nome, "")
            )
            if ok and senha:
                st.session_state["usuario"] = nome
                st.rerun()
            else:
                st.error("PIN/senha incorreto. Tente novamente.")


# ----------------------------------------------------------------------------
# Tela Vendedor
# ----------------------------------------------------------------------------
def tela_vendedor(nome):
    st.title(f"Olá, {nome.title()}!")
    if "flash" in st.session_state:
        st.success(st.session_state.pop("flash"))

    aba_nova, aba_orc, aba_hoje, aba_cargas, aba_rank = st.tabs(
        ["➕ Nova venda", "📄 Orçamentos em aberto",
         "📋 Meus lançamentos de hoje", "🚛 Cargas", "🏆 Ranking"]
    )

    df        = carregar_registros()
    df_cargas = carregar_cargas()
    hoje      = agora().strftime("%d/%m/%Y")
    mes       = agora().strftime("%m/%Y")
    fv        = st.session_state.get("form_ver", 0)

    # --- Nova venda ---
    with aba_nova:
        opcoes_carga = list(df_cargas["Carga"].unique()) if not df_cargas.empty else []
        if opcoes_carga:
            carga_val = st.selectbox("🚛 Carga *", opcoes_carga, key=f"carga_{fv}")
        else:
            st.warning("Nenhuma carga cadastrada. Peça ao gestor para cadastrar antes de lançar.")
            carga_val = ""

        cliente = st.text_input("Cliente *", key=f"cli_{fv}")
        _clientes_set = set(c.upper() for c in carregar_clientes())
        _sugerido = "Carteira" if cliente.strip().upper() in _clientes_set and cliente.strip() else "Novo"
        tipo = st.radio(
            "Tipo de cliente *",
            ["Carteira", "Novo"],
            index=["Carteira", "Novo"].index(_sugerido),
            horizontal=True, key=f"tipo_{fv}",
        )
        cliente_novo = tipo == "Novo"
        with st.expander("Com quem falou / detalhes"):
            contato = st.text_input("Com quem falou", key=f"cont_{fv}")
        resultado        = st.radio("Resultado do contato *", RESULTADOS,
                                    horizontal=True, key=f"res_{fv}")

        kg = valor = 0.0
        if resultado != "Só contato":
            c1, c2, c3 = st.columns(3)
            kg       = c1.number_input("Kg *", min_value=0.0, step=10.0,
                                       format="%.2f", key=f"kg_{fv}")
            preco_kg = c2.number_input("R$/kg *", min_value=0.0,
                                       step=0.5, format="%.2f", key=f"preco_{fv}")
            valor    = round(kg * preco_kg, 2)
            c3.metric("Valor total (automático)", "R$ " + br(valor))

        if st.button("✅ Registrar", type="primary"):
            if not carga_val:
                st.error("Selecione a carga antes de registrar.")
            elif not cliente.strip():
                st.error("Informe o nome do cliente.")
            elif resultado != "Só contato" and (kg <= 0 or preco_kg <= 0):
                st.error("Para orçamento ou venda, informe Kg e R$/kg.")
            else:
                salvar_registro(
                    nome, cliente, cliente_novo, contato,
                    resultado, kg, valor, carga_val,
                )
                msg = "Lançamento registrado!"
                if cliente_novo:
                    msg += f" Cliente novo '{cliente.strip()}' cadastrado."
                st.session_state["flash"] = msg + " Data e hora gravadas automaticamente."
                st.session_state["form_ver"] = fv + 1
                st.balloons()
                st.rerun()

    # --- Orçamentos em aberto ---
    with aba_orc:
        abertos = df[
            (df["Vendedor"].astype(str).str.strip() == nome) &
            (df["Resultado"].astype(str).str.strip() == "Orçamento enviado") &
            (df["Situação"].astype(str).str.strip() == SITUACAO_ABERTO)
        ]
        if abertos.empty:
            n_meus = len(df[df["Vendedor"].astype(str).str.strip() == nome])
            st.info("Nenhum orçamento em aberto.")
            if n_meus > 0:
                st.caption(f"Você tem {n_meus} lançamento(s) no total. Orçamentos aparecem apenas quando o resultado for **'Orçamento enviado'** — confira a aba **Meus lançamentos de hoje**.")
        else:
            st.caption("Quando o cliente responder, atualize aqui — sem redigitar nada.")
            for _, reg in abertos.iterrows():
                dias = reg.get("Dias em Aberto", pd.NA)
                dias_str = ""
                if pd.notna(dias) and dias != "":
                    dias_num = int(dias)
                    dias_str = f" ⚠️ {dias_num} dias aberto" if dias_num > 5 else f" ({dias_num}d)"
                c1, c2, c3 = st.columns([3, 1, 1])
                c1.write(
                    f"**{reg['Cliente']}** — {reg['Data']} — "
                    f"{br(reg['Kg'])} kg — R$ {br(reg['Valor (R$)'])}{dias_str}"
                )
                if c2.button("✅ Aprovou", key=f"ap{reg['_linha']}"):
                    aprovar_orcamento(reg)
                    st.session_state["flash"] = (
                        f"Orçamento de {reg['Cliente']} convertido em venda fechada!")
                    st.rerun()

                # Perdido: abre dropdown de motivo antes de confirmar
                chave_perd = f"perd_motivo_{reg['_linha']}"
                if st.session_state.get(chave_perd):
                    cm1, cm2, cm3 = st.columns([3, 1, 1])
                    motivo_sel = cm1.selectbox(
                        "Motivo da perda (opcional)",
                        ["— Não informado —"] + MOTIVOS_PERDA,
                        key=f"mot_{reg['_linha']}",
                    )
                    if cm2.button("Confirmar", key=f"conf_perd_{reg['_linha']}", type="primary"):
                        motivo_final = "" if motivo_sel == "— Não informado —" else motivo_sel
                        perder_orcamento(reg, motivo=motivo_final)
                        st.session_state.pop(chave_perd, None)
                        st.session_state["flash"] = (
                            f"Orçamento de {reg['Cliente']} marcado como perdido.")
                        st.rerun()
                    if cm3.button("Cancelar", key=f"canc_perd_{reg['_linha']}"):
                        st.session_state.pop(chave_perd, None)
                else:
                    if c3.button("❌ Perdido", key=f"pd{reg['_linha']}"):
                        st.session_state[chave_perd] = True

    # --- Meus lançamentos de hoje ---
    with aba_hoje:
        meus = df[(df["Vendedor"] == nome) & (df["Data"] == hoje)]
        st.metric("Lançamentos hoje", len(meus))

        if meus.empty:
            st.info("Nenhum lançamento hoje.")
        else:
            for _, reg in meus.iterrows():
                c1, c2 = st.columns([5, 1])
                carga_info = f" | 🚛 {reg['Carga']}" if str(reg.get("Carga", "")).strip() else ""
                c1.write(
                    f"**{reg['Hora']}** — {reg['Cliente']} — {reg['Resultado']} — "
                    f"{br(reg['Kg'])} kg — R$ {br(reg['Valor (R$)'])}{carga_info}"
                )
                chave      = f"del_{reg['_linha']}"
                chave_conf = f"conf_{reg['_linha']}"
                if st.session_state.get(chave_conf):
                    cc1, cc2, cc3 = st.columns([3, 1, 1])
                    cc1.warning("Confirma exclusão deste lançamento?")
                    if cc2.button("Sim, apagar", key=f"sim_{reg['_linha']}", type="primary"):
                        deletar_registro(reg["_linha"])
                        st.session_state.pop(chave_conf, None)
                        st.session_state["flash"] = "Lançamento apagado."
                        st.rerun()
                    if cc3.button("Cancelar", key=f"nao_{reg['_linha']}"):
                        st.session_state.pop(chave_conf, None)
                        st.rerun()
                else:
                    if c2.button("🗑️ Apagar", key=chave):
                        st.session_state[chave_conf] = True
                        st.rerun()

    # --- Cargas do vendedor ---
    with aba_cargas:
        minhas_cargas = df_cargas[df_cargas["Vendedor"] == nome]

        # Seção 1: cargas atribuídas com meta
        if minhas_cargas.empty:
            st.info("Você não tem cargas atribuídas no momento.")
        else:
            st.caption("Progresso das suas cargas:")
            widget_cargas(minhas_cargas, df)

        # Seção 2: cargas em que vendeu mas não está atribuído
        nomes_minhas = set(minhas_cargas["Carga"].tolist())
        vendas_meu = df[
            (df["Vendedor"] == nome) &
            (df["Resultado"] == "Venda fechada") &
            (df["Carga"].astype(str).str.strip() != "")
        ]
        outras_nomes = sorted(set(vendas_meu["Carga"].unique()) - nomes_minhas)
        if outras_nomes:
            st.divider()
            st.caption("Vendas em outras cargas (sem meta individual):")
            for carga_nome in outras_nomes:
                kg_meu = float(vendas_meu[vendas_meu["Carga"] == carga_nome]["Kg"].sum())
                info = df_cargas[df_cargas["Carga"] == carga_nome]
                data_ent = info["Data Entrega"].iloc[0] if not info.empty else "-"
                st.markdown(f"**{carga_nome}** — {data_ent}")
                st.caption(f"✅ {br(kg_meu)} kg vendido por você")
                st.divider()

    # --- Ranking ---
    with aba_rank:
        st.subheader(f"Hoje ({hoje})")
        tabela_ranking(resumir(df[df["Data"] == hoje]), destaque=nome)
        st.subheader(f"Mês ({mes})")
        tabela_ranking(resumir(df[df["_mes"] == mes]), destaque=nome)


# ----------------------------------------------------------------------------
# Tela Gestor
# ----------------------------------------------------------------------------
def tela_gestor():
    st.title("Painel do Gestor")
    df        = carregar_registros()
    df_cargas = carregar_cargas()
    hoje      = agora().strftime("%d/%m/%Y")
    mes       = agora().strftime("%m/%Y")
    df_hoje   = df[df["Data"] == hoje]
    df_mes    = df[df["_mes"] == mes]

    # Orçamentos vencidos (todos, não só hoje)
    orc_todos_abertos = df[
        (df["Resultado"] == "Orçamento enviado") &
        (df["Situação"] == SITUACAO_ABERTO)
    ]
    dias_num  = pd.to_numeric(orc_todos_abertos.get("Dias em Aberto", pd.Series(dtype=str)),
                               errors="coerce")
    n_vencidos = int((dias_num > 5).sum())

    aba_dia, aba_mes, aba_orc, aba_cargas, aba_funil, aba_perdas, aba_dados = st.tabs(
        ["📅 Hoje", "📈 Mês", "💰 Orçamentos", "🚛 Cargas",
         "🔻 Funil", "📊 Análise de Perdas", "🗂 Dados completos"]
    )

    # --- Hoje ---
    with aba_dia:
        kpis = _kpis(df_hoje)
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("Lançamentos",     len(df_hoje))
        c2.metric("Clientes novos",  int((df_hoje["Tipo cliente"] == "Novo").sum())
                  if "Tipo cliente" in df_hoje.columns else 0)
        c3.metric("Vendas fechadas", kpis["n_vendas"])
        c4.metric("Kg vendido",      br(kpis["kg"]))
        c5.metric("R$ vendido",      "R$ " + br(kpis["rs"]))
        c6.metric("Ticket médio",    "R$ " + br(kpis["ticket"]))

        c1b, c2b, c3b, c4b = st.columns(4)
        c1b.metric("Taxa conversão",        f"{kpis['conv']:.0f}%")
        c2b.metric("Taxa de perda",         f"{kpis['perda_pct']:.0f}%")
        c3b.metric("Recorrência",           f"{kpis['recorr_pct']:.0f}%")
        c4b.metric("Orçamentos vencidos +5d", n_vencidos)

        n_vend = len(carregar_vendedores()) or 1
        contatos_vend = len(df_hoje) / n_vend
        tempo_medio_df = pd.to_numeric(
            df.loc[(df["Resultado"] == "Orçamento enviado") & (df["Situação"] == SITUACAO_APROVADO),
                   "Tempo até Resposta (dias)"], errors="coerce"
        )
        tempo_medio = tempo_medio_df.mean()
        novos_pct = 100 * int((df_hoje["Tipo cliente"] == "Novo").sum()) / len(df_hoje) if len(df_hoje) else 0

        c1c, c2c, c3c = st.columns(3)
        c1c.metric("Contatos/vendedor hoje", f"{contatos_vend:.1f}",
                   delta="meta >2,5", delta_color="off")
        c2c.metric("Tempo médio aprovação",
                   f"{tempo_medio:.0f} dias" if pd.notna(tempo_medio) else "—")
        c3c.metric("Mix hoje (novos)", f"{novos_pct:.0f}%")

        if n_vencidos > 0:
            st.warning(f"⚠️ {n_vencidos} orçamento(s) sem resposta há mais de 5 dias — veja a aba Orçamentos.")

        sem = sorted(set(carregar_vendedores()) - set(df_hoje["Vendedor"]))
        if sem:
            st.warning("⚠️ Sem lançamentos hoje: " + ", ".join(sem))
        else:
            st.success("Todos os vendedores lançaram hoje.")
        tabela_ranking(resumir(df_hoje))

    # --- Mês ---
    with aba_mes:
        kpis = _kpis(df_mes)
        vendas_mes = df_mes[df_mes["Resultado"] == "Venda fechada"]
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("Lançamentos",     len(df_mes))
        c2.metric("Clientes novos",  int((df_mes["Tipo cliente"] == "Novo").sum())
                  if "Tipo cliente" in df_mes.columns else 0)
        c3.metric("Vendas fechadas", kpis["n_vendas"])
        c4.metric("Kg vendido",      br(kpis["kg"]))
        c5.metric("R$ vendido",      "R$ " + br(kpis["rs"]))
        c6.metric("Ticket médio",    "R$ " + br(kpis["ticket"]))

        c1b, c2b, c3b = st.columns(3)
        c1b.metric("Taxa conversão", f"{kpis['conv']:.0f}%")
        c2b.metric("Taxa de perda",  f"{kpis['perda_pct']:.0f}%")
        c3b.metric("Recorrência",    f"{kpis['recorr_pct']:.0f}%")

        tabela_ranking(resumir(df_mes))
        if not vendas_mes.empty:
            st.subheader("Kg vendido por dia")
            por_dia = vendas_mes.groupby("Data")["Kg"].sum()
            por_dia.index = pd.to_datetime(por_dia.index, format="%d/%m/%Y")
            st.bar_chart(por_dia.sort_index())

    # --- Orçamentos ---
    with aba_orc:
        orc     = df[df["Resultado"] == "Orçamento enviado"]
        abertos = orc[orc["Situação"] == SITUACAO_ABERTO]
        aprov   = orc[orc["Situação"] == SITUACAO_APROVADO]
        perd    = orc[orc["Situação"] == SITUACAO_PERDIDO]
        decididos = len(aprov) + len(perd)
        taxa = 100 * len(aprov) / decididos if decididos else 0
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Em aberto",          len(abertos))
        c2.metric("R$ em aberto",       "R$ " + br(abertos["Valor (R$)"].sum()))
        c3.metric("Aprovados / Perdidos", f"{len(aprov)} / {len(perd)}")
        c4.metric("Taxa de aprovação",  f"{taxa:.0f}%")

        if n_vencidos > 0:
            vencidos_df = orc_todos_abertos[
                pd.to_numeric(orc_todos_abertos.get("Dias em Aberto",
                              pd.Series(dtype=str)), errors="coerce") > 5
            ]
            st.warning(f"⚠️ {n_vencidos} orçamento(s) sem resposta há mais de 5 dias:")
            cols_v = [c for c in ["Data", "Vendedor", "Cliente", "Kg",
                                   "Valor (R$)", "Dias em Aberto"] if c in vencidos_df.columns]
            st.dataframe(vencidos_df[cols_v], hide_index=True, use_container_width=True)

        if not orc.empty:
            st.subheader("Por vendedor")
            linhas = []
            for vend, g in orc.groupby("Vendedor"):
                ab = g[g["Situação"] == SITUACAO_ABERTO]
                ap = g[g["Situação"] == SITUACAO_APROVADO]
                pe = g[g["Situação"] == SITUACAO_PERDIDO]
                dec = len(ap) + len(pe)
                linhas.append([
                    vend, len(ab), "R$ " + br(ab["Valor (R$)"].sum()),
                    len(ap), len(pe),
                    f"{100 * len(ap) / dec:.0f}%" if dec else "-",
                ])
            st.dataframe(
                pd.DataFrame(linhas, columns=[
                    "Vendedor", "Em aberto", "R$ em aberto",
                    "Aprovados", "Perdidos", "Taxa de aprovação",
                ]),
                hide_index=True, use_container_width=True,
            )
        if not abertos.empty:
            st.subheader("Orçamentos aguardando resposta")
            cols_a = [c for c in ["Data", "Vendedor", "Cliente", "Kg",
                                   "Valor (R$)", "Dias em Aberto"] if c in abertos.columns]
            st.dataframe(abertos[cols_a], hide_index=True, use_container_width=True)

        if not perd.empty:
            st.subheader(f"❌ Orçamentos perdidos ({len(perd)})")
            cols_p = [c for c in ["Data", "Vendedor", "Cliente", "Contato do Cliente",
                                   "Kg", "Valor (R$)", "Motivo da Perda",
                                   "Data Resposta Orçamento"] if c in perd.columns]
            perd_show = perd[cols_p].copy()
            perd_show["Valor (R$)"] = perd_show["Valor (R$)"].map(lambda v: "R$ " + br(v))
            perd_show["Kg"] = perd_show["Kg"].map(br)
            st.dataframe(perd_show, hide_index=True, use_container_width=True)

    # --- Cargas ---
    with aba_cargas:
        st.subheader("Cadastrar nova carga")
        vendedores_lista = list(carregar_vendedores().keys())
        with st.form("form_nova_carga", clear_on_submit=True):
            c1, c2, c3, c4 = st.columns(4)
            nc_nome    = c1.text_input("Nome da carga *", placeholder="Ex: BAURU")
            nc_data    = c2.text_input("Data de entrega *", placeholder="Ex: 20/06/2026")
            nc_vend    = c3.selectbox("Vendedor responsável *", vendedores_lista)
            nc_meta    = c4.number_input("Meta (Kg) *", min_value=0.0, step=100.0, format="%.0f")
            salvar_btn = st.form_submit_button("➕ Adicionar carga", type="primary")
            if salvar_btn:
                if not nc_nome.strip() or not nc_data.strip() or nc_meta <= 0:
                    st.error("Preencha todos os campos obrigatórios.")
                else:
                    salvar_carga(nc_nome, nc_data, nc_vend, nc_meta)
                    st.success(f"Carga '{nc_nome}' adicionada!")
                    st.rerun()

        st.divider()
        st.subheader(f"Cargas cadastradas ({len(df_cargas)})")
        if df_cargas.empty:
            st.info("Nenhuma carga ainda.")
        else:
            for nome_carga, grupo in df_cargas.groupby("Carga", sort=False):
                data_entrega   = grupo["Data Entrega"].iloc[0]
                vendedores_str = ", ".join(grupo["Vendedor"].tolist())
                meta_total     = float(grupo["Meta (Kg)"].sum())
                realiz_total   = progresso_carga(nome_carga, df)
                pct_total = min(realiz_total / meta_total, 1.0) if meta_total else 0.0
                cor_total = "🟢" if pct_total >= 1.0 else ("🟡" if pct_total >= 0.6 else "🔴")

                st.markdown(
                    f"### 🚛 {nome_carga} &nbsp;|&nbsp; {data_entrega} &nbsp;|&nbsp; "
                    f"Vendedores: {vendedores_str} &nbsp;|&nbsp; Total meta: **{br(meta_total)} kg**"
                )
                st.progress(
                    pct_total,
                    text=f"{cor_total} Total realizado: {br(realiz_total)} / {br(meta_total)} kg  ({pct_total*100:.0f}%)"
                )

                for _, carga in grupo.iterrows():
                    realizado = progresso_carga(carga["Carga"], df, vendedor=carga["Vendedor"])
                    meta      = float(carga["Meta (Kg)"]) or 1.0
                    pct       = min(realizado / meta, 1.0)
                    cor       = "🟢" if pct >= 1.0 else ("🟡" if pct >= 0.6 else "🔴")
                    c1, c2    = st.columns([6, 1])
                    c1.markdown(f"↳ **{carga['Vendedor']}** — Meta individual: {br(meta)} kg")
                    c1.progress(pct, text=f"{cor} {br(realizado)} / {br(meta)} kg  ({pct*100:.0f}%)")
                    chave_del  = f"gd_{carga['_linha']}"
                    chave_conf = f"gdc_{carga['_linha']}"
                    if st.session_state.get(chave_conf):
                        c2.warning("Apagar?")
                        if c2.button("Sim", key=f"gsim_{carga['_linha']}"):
                            deletar_carga(carga["_linha"])
                            st.session_state.pop(chave_conf, None)
                            st.rerun()
                        if c2.button("Não", key=f"gnao_{carga['_linha']}"):
                            st.session_state.pop(chave_conf, None)
                            st.rerun()
                    else:
                        if c2.button("🗑️", key=chave_del, help="Apagar"):
                            st.session_state[chave_conf] = True
                            st.rerun()
                # Redistribuição sugerida (quando há múltiplos vendedores na carga)
                if len(grupo) > 1:
                    pcts_vend = []
                    for _, c in grupo.iterrows():
                        r = progresso_carga(c["Carga"], df, vendedor=c["Vendedor"])
                        m = float(c["Meta (Kg)"]) or 1.0
                        pcts_vend.append((c["Vendedor"], r / m * 100))
                    melhor = max(pcts_vend, key=lambda x: x[1])
                    pior   = min(pcts_vend, key=lambda x: x[1])
                    if melhor[1] - pior[1] > 30:
                        st.info(
                            f"💡 {melhor[0]} está em {melhor[1]:.0f}% e "
                            f"{pior[0]} em {pior[1]:.0f}% — considerar redistribuição?"
                        )
                st.divider()

    # --- Funil ---
    with aba_funil:
        st.caption("Contatos → Orçamentos → Vendas (mês atual)")
        res = resumir(df_mes)
        if res.empty:
            st.info("Sem dados no mês.")
        else:
            fun = res[["Vendedor", "Lançamentos", "Orçamentos", "Vendas",
                       "Conversão (%)", "Taxa Perda (%)"]]
            st.dataframe(fun, hide_index=True, use_container_width=True)
            st.bar_chart(res.set_index("Vendedor")[["Lançamentos", "Orçamentos", "Vendas"]])

    # --- Análise de Perdas ---
    with aba_perdas:
        st.subheader("Análise de Perdas")
        perdidos_df = df[df["Situação"] == SITUACAO_PERDIDO].copy()

        if perdidos_df.empty:
            st.info("Nenhuma perda registrada ainda.")
        else:
            periodo = st.radio("Período:", ["Mês atual", "Todos os meses"],
                               horizontal=True, key="perd_periodo")
            if periodo == "Mês atual":
                perdidos_df = perdidos_df[perdidos_df["_mes"] == mes]

            if perdidos_df.empty:
                st.info("Nenhuma perda no período selecionado.")
            else:
                total_perd = len(perdidos_df)
                vendas_periodo = df[df["Resultado"] == "Venda fechada"]
                if periodo == "Mês atual":
                    vendas_periodo = vendas_periodo[vendas_periodo["_mes"] == mes]
                total_vend = len(vendas_periodo)
                dec_total  = total_vend + total_perd
                taxa_perd  = 100 * total_perd / dec_total if dec_total > 0 else 0.0

                c1, c2, c3 = st.columns(3)
                c1.metric("Perdas no período", total_perd)
                c2.metric("Vendas fechadas",   total_vend)
                c3.metric("Taxa de perda",     f"{taxa_perd:.0f}%")
                if taxa_perd > 25:
                    st.warning("⚠️ Taxa de perda acima de 25% — investigar causas.")

                st.divider()

                # Por motivo
                col_m = "Motivo da Perda"
                motivos = (perdidos_df[col_m].astype(str).str.strip()
                           .replace("", "Não informado"))
                contagem = motivos.value_counts().reset_index()
                contagem.columns = ["Motivo", "Qtd"]
                contagem["% do Total"] = (
                    contagem["Qtd"] / total_perd * 100
                ).round(1).astype(str) + "%"

                st.subheader("Por motivo")
                st.dataframe(contagem, hide_index=True, use_container_width=True)
                st.bar_chart(contagem.set_index("Motivo")["Qtd"])

                # Por vendedor
                st.subheader("Por vendedor")
                linhas_v = []
                for vend, g in perdidos_df.groupby("Vendedor"):
                    vend_vendas = vendas_periodo[vendas_periodo["Vendedor"] == vend]
                    dec_v = len(vend_vendas) + len(g)
                    taxa_v = 100 * len(g) / dec_v if dec_v > 0 else 0.0
                    top_mot = (g[col_m].astype(str).str.strip()
                               .replace("", "Não informado")
                               .value_counts())
                    motivo_top = top_mot.index[0] if not top_mot.empty else "—"
                    linhas_v.append([vend, len(g), f"{taxa_v:.0f}%", motivo_top])
                st.dataframe(
                    pd.DataFrame(linhas_v, columns=[
                        "Vendedor", "Perdas", "Taxa Perda", "Motivo principal"
                    ]),
                    hide_index=True, use_container_width=True,
                )

                # Recomendação automática
                top_motivo = motivos.value_counts().index[0] if not motivos.empty else ""
                recomendacoes = {
                    "Preço":
                        "**Preço** é o motivo principal. Revise a estratégia de precificação ou o mix de produto.",
                    "Concorrente entregou antes":
                        "**Timing** é o gargalo. Verifique se há como acelerar produção ou prazos de entrega.",
                    "Cliente achou qualidade baixa":
                        "**Qualidade percebida** precisa de atenção. Revise especificações ou comunicação técnica.",
                    "Sem resposta / desistiu":
                        "Muitos clientes desapareceram. Melhore o **follow-up** nos orçamentos em aberto.",
                    "Outro":
                        "Motivos variados — incentive os vendedores a detalhar melhor o motivo da perda.",
                    "Não informado":
                        "Maioria das perdas sem motivo registrado — oriente os vendedores a informar ao marcar como perdido.",
                }
                rec = recomendacoes.get(top_motivo, "")
                if rec:
                    st.divider()
                    st.info(f"💡 Recomendação: {rec}")

    # --- Dados completos ---
    with aba_dados:
        c1, c2 = st.columns(2)
        f_vend = c1.multiselect(
            "Vendedor",
            sorted(df["Vendedor"].unique()) if not df.empty else [],
        )
        f_res  = c2.multiselect("Resultado", RESULTADOS)
        dados  = df.copy()
        if f_vend:
            dados = dados[dados["Vendedor"].isin(f_vend)]
        if f_res:
            dados = dados[dados["Resultado"].isin(f_res)]
        _ocultar = {"Contato do Cliente", "Cliente Recorrente?"}
        colunas_exibir = [c for c in CABECALHO_REGISTROS if c in dados.columns and c not in _ocultar]
        st.dataframe(dados[colunas_exibir], hide_index=True, use_container_width=True)
        c_dl1, c_dl2 = st.columns(2)
        c_dl1.download_button(
            "⬇️ Baixar CSV",
            dados[colunas_exibir].to_csv(index=False, sep=";",
                                          decimal=",").encode("utf-8-sig"),
            file_name=f"vendas_{agora().strftime('%Y%m%d')}.csv",
        )
        _buf = io.BytesIO()
        with pd.ExcelWriter(_buf, engine="openpyxl") as _w:
            dados[colunas_exibir].to_excel(_w, index=False, sheet_name="Registros")
        c_dl2.download_button(
            "⬇️ Baixar Excel",
            _buf.getvalue(),
            file_name=f"vendas_{agora().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        if st.button("🔄 Atualizar abas de análise na planilha"):
            atualizar_abas_analise()
            st.success("Abas 'Análise Diária' e 'Análise Mensal' atualizadas.")


# ----------------------------------------------------------------------------
# Fluxo principal
# ----------------------------------------------------------------------------
def main():
    usuario = st.session_state.get("usuario")
    if not usuario:
        tela_login()
        return

    with st.sidebar:
        st.write(f"Conectado como **{usuario.title()}**")
        if st.button("Sair"):
            st.session_state.pop("usuario", None)
            st.rerun()
        if st.button("Atualizar dados"):
            carregar_registros.clear()
            carregar_vendedores.clear()
            carregar_clientes.clear()
            carregar_cargas.clear()
            st.rerun()

    if usuario == "GESTOR":
        tela_gestor()
    else:
        tela_vendedor(usuario)


if __name__ == "__main__":
    main()
