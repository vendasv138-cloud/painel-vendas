# -*- coding: utf-8 -*-
"""
Painel de Vendas - Controle de prospecção e vendas da equipe
Banco de dados: Google Sheets (no Google Drive)
"""

from datetime import datetime
from zoneinfo import ZoneInfo

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
]

CABECALHO_CARGAS   = ["Carga", "Data Entrega", "Vendedor", "Meta (Kg)"]
CABECALHO_CLIENTES = ["Cliente", "Cadastrado em", "Cadastrado por"]

SITUACAO_ABERTO   = "Em aberto"
SITUACAO_APROVADO = "Aprovado"
SITUACAO_PERDIDO  = "Perdido"

RESULTADOS = ["Só contato", "Orçamento enviado", "Venda fechada"]

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


def garantir_estrutura(sh):
    titulos = [ws.title for ws in sh.worksheets()]

    if ABA_REGISTROS not in titulos:
        ws = sh.add_worksheet(title=ABA_REGISTROS, rows=2000, cols=13)
        ws.update(values=[CABECALHO_REGISTROS], range_name="A1")
        ws.format("A1:L1", {"textFormat": {"bold": True}})
    else:
        # Garante coluna "Carga" em sheets já existentes
        ws = sh.worksheet(ABA_REGISTROS)
        cabecalho_atual = ws.row_values(1)
        if "Carga" not in cabecalho_atual:
            col = len(cabecalho_atual) + 1
            ws.update_cell(1, col, "Carga")

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
        sh.add_worksheet(title=ABA_DIARIA, rows=100, cols=10)

    if ABA_MENSAL not in titulos:
        sh.add_worksheet(title=ABA_MENSAL, rows=500, cols=12)


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


def progresso_carga(carga_nome, df_registros):
    """Kg de vendas fechadas vinculadas a esta carga."""
    mask = (
        (df_registros["Carga"] == carga_nome) &
        (df_registros["Resultado"] == "Venda fechada")
    )
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
    sh = abrir_planilha()
    dados = sh.worksheet(ABA_REGISTROS).get_all_records()
    df = pd.DataFrame(dados)
    if df.empty:
        return pd.DataFrame(columns=CABECALHO_REGISTROS + ["_data", "_mes", "_linha"])
    df["_linha"] = df.index + 2

    # Garante coluna Carga mesmo em sheets antigas
    if "Carga" not in df.columns:
        df["Carga"] = ""

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
    return df


def salvar_registro(vendedor, cliente, cliente_novo, contato,
                    resultado, kg, valor, carga=""):
    sh = abrir_planilha()
    ws = sh.worksheet(ABA_REGISTROS)
    ts = agora()
    preco_kg = round(valor / kg, 2) if kg and valor else 0
    if cliente_novo and cliente.strip().upper() not in {
            c.upper() for c in carregar_clientes()}:
        cadastrar_cliente(cliente, vendedor)
    ws.append_row(
        [
            ts.strftime("%d/%m/%Y"),
            ts.strftime("%H:%M"),
            vendedor,
            cliente.strip(),
            "Novo" if cliente_novo else "Carteira",
            contato.strip(),
            resultado,
            SITUACAO_ABERTO if resultado == "Orçamento enviado" else "-",
            round(float(kg or 0), 2),
            round(float(valor or 0), 2),
            preco_kg,
            carga or "",
        ],
        value_input_option="RAW",
    )
    carregar_registros.clear()
    try:
        atualizar_abas_analise()
    except Exception:
        pass


def mudar_situacao(linha_planilha, nova):
    sh = abrir_planilha()
    col = CABECALHO_REGISTROS.index("Situação") + 1
    sh.worksheet(ABA_REGISTROS).update_cell(int(linha_planilha), col, nova)
    carregar_registros.clear()


def aprovar_orcamento(reg):
    mudar_situacao(reg["_linha"], SITUACAO_APROVADO)
    salvar_registro(reg["Vendedor"], reg["Cliente"], False,
                    str(reg["Com quem falou"]), "Venda fechada",
                    float(reg["Kg"]), float(reg["Valor (R$)"]),
                    str(reg.get("Carga", "")))


def perder_orcamento(reg):
    mudar_situacao(reg["_linha"], SITUACAO_PERDIDO)
    try:
        atualizar_abas_analise()
    except Exception:
        pass


def deletar_registro(linha_planilha):
    sh = abrir_planilha()
    sh.worksheet(ABA_REGISTROS).delete_rows(int(linha_planilha))
    carregar_registros.clear()
    try:
        atualizar_abas_analise()
    except Exception:
        pass


# ----------------------------------------------------------------------------
# Resumos
# ----------------------------------------------------------------------------
COLUNAS_RESUMO = ["Vendedor", "Lançamentos", "Clientes novos", "Orçamentos",
                  "Vendas", "Kg vendido", "R$ vendido", "R$/kg médio",
                  "Conversão (%)"]


def resumir(df):
    if df.empty:
        return pd.DataFrame(columns=COLUNAS_RESUMO)
    linhas = []
    for vend, g in df.groupby("Vendedor"):
        vendas = g[g["Resultado"] == "Venda fechada"]
        orcs   = g[g["Resultado"] == "Orçamento enviado"]
        novos  = g[g["Tipo cliente"] == "Novo"]
        kg     = vendas["Kg"].sum()
        rs     = vendas["Valor (R$)"].sum()
        conv   = 100 * len(vendas) / len(g) if len(g) else 0
        linhas.append([
            vend, len(g), len(novos), len(orcs), len(vendas),
            round(kg, 2), round(rs, 2),
            round(rs / kg, 2) if kg else 0, round(conv, 1),
        ])
    out = pd.DataFrame(linhas, columns=COLUNAS_RESUMO)
    return out.sort_values(["Kg vendido", "R$ vendido", "Lançamentos"],
                           ascending=False).reset_index(drop=True)


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
    ws.format("A3:I3", {"textFormat": {"bold": True}})

    ws = sh.worksheet(ABA_MENSAL)
    ws.clear()
    valores = [["ANÁLISE MENSAL (todos os meses)"], [], ["Mês"] + COLUNAS_RESUMO]
    for mes in sorted(df["_mes"].dropna().unique(), reverse=True):
        res_mes = resumir(df[df["_mes"] == mes])
        for linha in res_mes.astype(object).values.tolist():
            valores.append([mes] + linha)
    ws.update(values=valores, range_name="A1")
    ws.format("A3:J3", {"textFormat": {"bold": True}})


# ----------------------------------------------------------------------------
# Componentes de tela
# ----------------------------------------------------------------------------
def tabela_ranking(res, destaque=None):
    if res.empty:
        st.info("Nenhum lançamento ainda.")
        return
    res = res.copy()
    res.insert(0, "Posição", [f"{i}º" for i in range(1, len(res) + 1)])
    res["Kg vendido"]  = res["Kg vendido"].map(lambda v: br(v))
    res["R$ vendido"]  = res["R$ vendido"].map(lambda v: "R$ " + br(v))
    res["R$/kg médio"] = res["R$/kg médio"].map(lambda v: "R$ " + br(v))
    st.dataframe(res, hide_index=True, use_container_width=True)
    if destaque is not None and destaque in res["Vendedor"].values:
        pos = res.loc[res["Vendedor"] == destaque, "Posição"].iloc[0]
        st.caption(f"Sua posição: {pos} de {len(res)}")


def widget_cargas(df_cargas, df_registros, filtro_vendedor=None):
    """Exibe cada carga com barra de progresso."""
    if df_cargas.empty:
        st.info("Nenhuma carga cadastrada ainda.")
        return
    for _, carga in df_cargas.iterrows():
        if filtro_vendedor and carga["Vendedor"] != filtro_vendedor:
            continue
        realizado = progresso_carga(carga["Carga"], df_registros)
        meta      = float(carga["Meta (Kg)"]) or 1.0
        pct       = min(realizado / meta, 1.0)
        cor       = "🟢" if pct >= 1.0 else ("🟡" if pct >= 0.6 else "🔴")
        st.markdown(
            f"**{carga['Carga']}** — {carga['Data Entrega']} — "
            f"{carga['Vendedor']} — Meta: {br(meta)} kg"
        )
        st.progress(pct, text=f"{cor} {br(realizado)} / {br(meta)} kg  ({pct*100:.0f}%)")
        st.divider()


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


def tela_vendedor(nome):
    st.title(f"Olá, {nome.title()}!")
    if "flash" in st.session_state:
        st.success(st.session_state.pop("flash"))

    aba_nova, aba_orc, aba_hoje, aba_cargas, aba_rank = st.tabs(
        ["➕ Nova venda", "📄 Orçamentos em aberto",
         "📋 Meus lançamentos de hoje", "🚛 Cargas", "🏆 Ranking"]
    )

    df         = carregar_registros()
    df_cargas  = carregar_cargas()
    hoje       = agora().strftime("%d/%m/%Y")
    mes        = agora().strftime("%m/%Y")

    # Chave de versão para limpar o formulário após salvar
    fv = st.session_state.get("form_ver", 0)

    # --- Nova venda ---
    with aba_nova:
        # Selectbox de carga (acima do cliente)
        opcoes_carga = ["— Nenhuma carga —"] + list(df_cargas["Carga"])
        carga_sel = st.selectbox(
            "🚛 Vincular a uma carga (opcional)",
            opcoes_carga,
            key=f"carga_{fv}",
        )
        carga_val = "" if carga_sel == "— Nenhuma carga —" else carga_sel

        c1, c2 = st.columns(2)
        cliente     = c1.text_input("Cliente *", key=f"cli_{fv}")
        tipo        = c1.radio("Tipo de cliente *", ["Carteira", "Novo"],
                               horizontal=True, key=f"tipo_{fv}")
        cliente_novo = tipo == "Novo"
        contato     = c2.text_input("Com quem falou", key=f"cont_{fv}")
        resultado   = st.radio("Resultado do contato *", RESULTADOS,
                               horizontal=True, key=f"res_{fv}")

        kg = valor = 0.0
        if resultado != "Só contato":
            c1, c2, c3 = st.columns(3)
            kg    = c1.number_input("Kg *", min_value=0.0, step=10.0,
                                    format="%.2f", key=f"kg_{fv}")
            valor = c2.number_input("Valor total (R$) *", min_value=0.0,
                                    step=100.0, format="%.2f", key=f"val_{fv}")
            preco = valor / kg if kg else 0
            c3.metric("R$/kg (automático)", "R$ " + br(preco))

        if st.button("✅ Registrar", type="primary"):
            if not cliente.strip():
                st.error("Informe o nome do cliente.")
            elif resultado != "Só contato" and (kg <= 0 or valor <= 0):
                st.error("Para orçamento ou venda, informe Kg e Valor.")
            else:
                salvar_registro(nome, cliente, cliente_novo, contato,
                                resultado, kg, valor, carga_val)
                msg = "Lançamento registrado!"
                if cliente_novo:
                    msg += f" Cliente novo '{cliente.strip()}' cadastrado."
                st.session_state["flash"] = msg + " Data e hora gravadas automaticamente."
                st.session_state["form_ver"] = fv + 1
                st.balloons()
                st.rerun()

    # --- Orçamentos em aberto ---
    with aba_orc:
        abertos = df[(df["Vendedor"] == nome)
                     & (df["Resultado"] == "Orçamento enviado")
                     & (df["Situação"] == SITUACAO_ABERTO)]
        if abertos.empty:
            st.info("Nenhum orçamento em aberto.")
        else:
            st.caption("Quando o cliente responder, atualize aqui — sem redigitar nada.")
            for _, reg in abertos.iterrows():
                c1, c2, c3 = st.columns([3, 1, 1])
                c1.write(
                    f"**{reg['Cliente']}** — {reg['Data']} — "
                    f"{br(reg['Kg'])} kg — R$ {br(reg['Valor (R$)'])}"
                )
                if c2.button("✅ Aprovou", key=f"ap{reg['_linha']}"):
                    aprovar_orcamento(reg)
                    st.session_state["flash"] = (
                        f"Orçamento de {reg['Cliente']} convertido em venda fechada!")
                    st.rerun()
                if c3.button("❌ Perdido", key=f"pd{reg['_linha']}"):
                    perder_orcamento(reg)
                    st.session_state["flash"] = (
                        f"Orçamento de {reg['Cliente']} marcado como perdido.")
                    st.rerun()

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
        if minhas_cargas.empty:
            st.info("Você não tem cargas atribuídas no momento.")
        else:
            st.caption("Acompanhe o progresso das suas cargas.")
            widget_cargas(minhas_cargas, df)

    # --- Ranking ---
    with aba_rank:
        st.subheader(f"Hoje ({hoje})")
        tabela_ranking(resumir(df[df["Data"] == hoje]), destaque=nome)
        st.subheader(f"Mês ({mes})")
        tabela_ranking(resumir(df[df["_mes"] == mes]), destaque=nome)


def tela_gestor():
    st.title("Painel do Gestor")
    df        = carregar_registros()
    df_cargas = carregar_cargas()
    hoje      = agora().strftime("%d/%m/%Y")
    mes       = agora().strftime("%m/%Y")
    df_hoje   = df[df["Data"] == hoje]
    df_mes    = df[df["_mes"] == mes]

    aba_dia, aba_mes, aba_orc, aba_cargas, aba_funil, aba_dados = st.tabs(
        ["📅 Hoje", "📈 Mês", "💰 Orçamentos", "🚛 Cargas", "🔻 Funil", "🗂 Dados completos"]
    )

    with aba_dia:
        vendas = df_hoje[df_hoje["Resultado"] == "Venda fechada"]
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Lançamentos", len(df_hoje))
        c2.metric("Clientes novos", int((df_hoje["Tipo cliente"] == "Novo").sum()))
        c3.metric("Vendas fechadas", len(vendas))
        c4.metric("Kg vendido", br(vendas["Kg"].sum()))
        c5.metric("R$ vendido", "R$ " + br(vendas["Valor (R$)"].sum()))
        sem = sorted(set(carregar_vendedores()) - set(df_hoje["Vendedor"]))
        if sem:
            st.warning("⚠️ Sem lançamentos hoje: " + ", ".join(sem))
        else:
            st.success("Todos os vendedores lançaram hoje.")
        tabela_ranking(resumir(df_hoje))

    with aba_mes:
        vendas = df_mes[df_mes["Resultado"] == "Venda fechada"]
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Lançamentos", len(df_mes))
        c2.metric("Clientes novos", int((df_mes["Tipo cliente"] == "Novo").sum()))
        c3.metric("Vendas fechadas", len(vendas))
        c4.metric("Kg vendido", br(vendas["Kg"].sum()))
        c5.metric("R$ vendido", "R$ " + br(vendas["Valor (R$)"].sum()))
        tabela_ranking(resumir(df_mes))
        if not vendas.empty:
            st.subheader("Kg vendido por dia")
            por_dia = vendas.groupby("Data")["Kg"].sum()
            por_dia.index = pd.to_datetime(por_dia.index, format="%d/%m/%Y")
            st.bar_chart(por_dia.sort_index())

    with aba_orc:
        orc     = df[df["Resultado"] == "Orçamento enviado"]
        abertos = orc[orc["Situação"] == SITUACAO_ABERTO]
        aprov   = orc[orc["Situação"] == SITUACAO_APROVADO]
        perd    = orc[orc["Situação"] == SITUACAO_PERDIDO]
        decididos = len(aprov) + len(perd)
        taxa = 100 * len(aprov) / decididos if decididos else 0
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Em aberto", len(abertos))
        c2.metric("R$ em aberto", "R$ " + br(abertos["Valor (R$)"].sum()))
        c3.metric("Aprovados / Perdidos", f"{len(aprov)} / {len(perd)}")
        c4.metric("Taxa de aprovação", f"{taxa:.0f}%")
        if not orc.empty:
            st.subheader("Por vendedor")
            linhas = []
            for vend, g in orc.groupby("Vendedor"):
                ab = g[g["Situação"] == SITUACAO_ABERTO]
                ap = g[g["Situação"] == SITUACAO_APROVADO]
                pe = g[g["Situação"] == SITUACAO_PERDIDO]
                dec = len(ap) + len(pe)
                linhas.append([vend, len(ab), "R$ " + br(ab["Valor (R$)"].sum()),
                               len(ap), len(pe),
                               f"{100 * len(ap) / dec:.0f}%" if dec else "-"])
            st.dataframe(pd.DataFrame(linhas, columns=[
                "Vendedor", "Em aberto", "R$ em aberto", "Aprovados",
                "Perdidos", "Taxa de aprovação"]),
                hide_index=True, use_container_width=True)
        if not abertos.empty:
            st.subheader("Orçamentos aguardando resposta")
            st.dataframe(
                abertos[["Data", "Vendedor", "Cliente", "Kg", "Valor (R$)"]],
                hide_index=True, use_container_width=True)

    # --- Aba Cargas ---
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
            for _, carga in df_cargas.iterrows():
                realizado = progresso_carga(carga["Carga"], df)
                meta      = float(carga["Meta (Kg)"]) or 1.0
                pct       = min(realizado / meta, 1.0)
                cor       = "🟢" if pct >= 1.0 else ("🟡" if pct >= 0.6 else "🔴")
                c1, c2 = st.columns([6, 1])
                c1.markdown(
                    f"**{carga['Carga']}** — {carga['Data Entrega']} — "
                    f"{carga['Vendedor']} — Meta: {br(meta)} kg"
                )
                c1.progress(pct,
                    text=f"{cor} {br(realizado)} / {br(meta)} kg  ({pct*100:.0f}%)")
                chave_del  = f"gd_{carga['_linha']}"
                chave_conf = f"gdc_{carga['_linha']}"
                if st.session_state.get(chave_conf):
                    cc1, cc2, cc3 = c2.columns([1, 1, 1]) if False else (c2, c2, c2)
                    c2.warning("Apagar?")
                    if c2.button("Sim", key=f"gsim_{carga['_linha']}"):
                        deletar_carga(carga["_linha"])
                        st.session_state.pop(chave_conf, None)
                        st.rerun()
                    if c2.button("Não", key=f"gnao_{carga['_linha']}"):
                        st.session_state.pop(chave_conf, None)
                        st.rerun()
                else:
                    if c2.button("🗑️", key=chave_del, help="Apagar carga"):
                        st.session_state[chave_conf] = True
                        st.rerun()
                st.divider()

    with aba_funil:
        st.caption("Contatos → Orçamentos → Vendas (mês atual)")
        res = resumir(df_mes)
        if res.empty:
            st.info("Sem dados no mês.")
        else:
            fun = res[["Vendedor", "Lançamentos", "Orçamentos", "Vendas",
                       "Conversão (%)"]]
            st.dataframe(fun, hide_index=True, use_container_width=True)
            st.bar_chart(res.set_index("Vendedor")[
                ["Lançamentos", "Orçamentos", "Vendas"]])

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
        colunas_exibir = [c for c in CABECALHO_REGISTROS if c in dados.columns]
        st.dataframe(dados[colunas_exibir], hide_index=True, use_container_width=True)
        st.download_button(
            "⬇️ Baixar CSV",
            dados[colunas_exibir].to_csv(index=False, sep=";",
                                         decimal=",").encode("utf-8-sig"),
            file_name=f"vendas_{agora().strftime('%Y%m%d')}.csv",
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
