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

ABA_REGISTROS = "Registros"
ABA_VENDEDORES = "Vendedores"
ABA_CLIENTES = "Clientes"
ABA_DIARIA = "Análise Diária"
ABA_MENSAL = "Análise Mensal"

CABECALHO_REGISTROS = [
    "Data", "Hora", "Vendedor", "Cliente", "Tipo cliente", "Com quem falou",
    "Resultado", "Situação", "Kg", "Valor (R$)", "R$/kg",
]

SITUACAO_ABERTO = "Em aberto"
SITUACAO_APROVADO = "Aprovado"
SITUACAO_PERDIDO = "Perdido"

CABECALHO_CLIENTES = ["Cliente", "Cadastrado em", "Cadastrado por"]

RESULTADOS = ["Só contato", "Orçamento enviado", "Venda fechada"]

# Vendedores criados automaticamente na primeira execução.
# Depois, gerencie direto na aba "Vendedores" da planilha (nome, PIN, ativo).
VENDEDORES_INICIAIS = [
    ["ANA PAULA", "1010", "SIM"],
    ["CAIO", "2020", "SIM"],
    ["VANDERLEI", "3030", "SIM"],
    ["JESUS", "4040", "SIM"],
    ["JONATAN", "5050", "SIM"],
    ["MARCIO", "6060", "SIM"],
    ["RENATA", "7070", "SIM"],
]

st.set_page_config(page_title="Painel de Vendas", page_icon="📊", layout="wide")


# ----------------------------------------------------------------------------
# Utilidades
# ----------------------------------------------------------------------------
def br(valor, casas=2):
    """Formata número no padrão brasileiro: 1.234,56"""
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
    """Cria as abas necessárias na primeira execução."""
    titulos = [ws.title for ws in sh.worksheets()]

    if ABA_REGISTROS not in titulos:
        ws = sh.add_worksheet(title=ABA_REGISTROS, rows=2000, cols=12)
        ws.update(values=[CABECALHO_REGISTROS], range_name="A1")
        ws.format("A1:K1", {"textFormat": {"bold": True}})

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

    if ABA_DIARIA not in titulos:
        sh.add_worksheet(title=ABA_DIARIA, rows=100, cols=10)

    if ABA_MENSAL not in titulos:
        sh.add_worksheet(title=ABA_MENSAL, rows=500, cols=12)


@st.cache_data(ttl=300)
def carregar_vendedores():
    sh = abrir_planilha()
    dados = sh.worksheet(ABA_VENDEDORES).get_all_records()
    ativos = {}
    for linha in dados:
        nome = str(linha.get("Nome", "")).strip()
        ativo = str(linha.get("Ativo", "")).strip().upper()
        if nome and ativo in ("SIM", "S", "1", "X"):
            ativos[nome] = str(linha.get("PIN", "")).strip()
    return ativos


@st.cache_data(ttl=60)
def carregar_clientes():
    """Lista de clientes cadastrados (aba Clientes da planilha)."""
    sh = abrir_planilha()
    dados = sh.worksheet(ABA_CLIENTES).get_all_records(
        expected_headers=CABECALHO_CLIENTES
    )
    nomes = []
    vistos = set()
    for linha in dados:
        nome = str(linha.get("Cliente", "")).strip()
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
    dados = sh.worksheet(ABA_REGISTROS).get_all_records(
        expected_headers=CABECALHO_REGISTROS
    )
    df = pd.DataFrame(dados)
    if df.empty:
        return pd.DataFrame(columns=CABECALHO_REGISTROS + ["_data", "_mes", "_linha"])
    df["_linha"] = df.index + 2

    for col in ("Kg", "Valor (R$)", "R$/kg"):
        if df[col].dtype == object:
            df[col] = (
                df[col].astype(str)
                .str.replace(".", "", regex=False)
                .str.replace(",", ".", regex=False)
            )
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    df["_data"] = pd.to_datetime(df["Data"], format="%d/%m/%Y", errors="coerce")
    df["_mes"] = df["_data"].dt.strftime("%m/%Y")
    return df


def salvar_registro(vendedor, cliente, cliente_novo, contato, resultado, kg, valor):
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
        ],
        value_input_option="RAW",
    )
    carregar_registros.clear()
    try:
        atualizar_abas_analise()
    except Exception:
        pass  # análise na planilha é secundária; não trava o lançamento


def mudar_situacao(linha_planilha, nova):
    sh = abrir_planilha()
    col = CABECALHO_REGISTROS.index("Situação") + 1
    sh.worksheet(ABA_REGISTROS).update_cell(int(linha_planilha), col, nova)
    carregar_registros.clear()


def aprovar_orcamento(reg):
    """Marca o orçamento como aprovado e registra a venda automaticamente."""
    mudar_situacao(reg["_linha"], SITUACAO_APROVADO)
    salvar_registro(reg["Vendedor"], reg["Cliente"], False,
                    str(reg["Com quem falou"]), "Venda fechada",
                    float(reg["Kg"]), float(reg["Valor (R$)"]))


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
# Resumos (usados no app e gravados na planilha)
# ----------------------------------------------------------------------------
COLUNAS_RESUMO = ["Vendedor", "Lançamentos", "Clientes novos", "Orçamentos",
                  "Vendas", "Kg vendido", "R$ vendido", "R$/kg médio",
                  "Conversão (%)"]


def resumir(df):
    """Resumo por vendedor: lançamentos, orçamentos, vendas, kg, R$, conversão."""
    if df.empty:
        return pd.DataFrame(columns=COLUNAS_RESUMO)
    linhas = []
    for vend, g in df.groupby("Vendedor"):
        vendas = g[g["Resultado"] == "Venda fechada"]
        orcs = g[g["Resultado"] == "Orçamento enviado"]
        novos = g[g["Tipo cliente"] == "Novo"]
        kg = vendas["Kg"].sum()
        rs = vendas["Valor (R$)"].sum()
        conv = 100 * len(vendas) / len(g) if len(g) else 0
        linhas.append([
            vend, len(g), len(novos), len(orcs), len(vendas),
            round(kg, 2), round(rs, 2),
            round(rs / kg, 2) if kg else 0, round(conv, 1),
        ])
    out = pd.DataFrame(linhas, columns=COLUNAS_RESUMO)
    return out.sort_values(["Kg vendido", "R$ vendido", "Lançamentos"],
                           ascending=False).reset_index(drop=True)


def atualizar_abas_analise():
    """Reescreve as abas de análise diária e mensal na planilha."""
    sh = abrir_planilha()
    df = carregar_registros()
    hoje = agora().strftime("%d/%m/%Y")

    # --- Diária ---
    df_hoje = df[df["Data"] == hoje]
    res_dia = resumir(df_hoje)
    sem_lancamento = sorted(set(carregar_vendedores()) - set(res_dia["Vendedor"]))
    ws = sh.worksheet(ABA_DIARIA)
    ws.clear()
    valores = [[f"ANÁLISE DO DIA {hoje}"], [], COLUNAS_RESUMO]
    valores += res_dia.astype(object).values.tolist()
    valores += [[], ["Sem lançamentos hoje:", ", ".join(sem_lancamento) or "ninguém"]]
    ws.update(values=valores, range_name="A1")
    ws.format("A3:I3", {"textFormat": {"bold": True}})

    # --- Mensal ---
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
    res["Kg vendido"] = res["Kg vendido"].map(lambda v: br(v))
    res["R$ vendido"] = res["R$ vendido"].map(lambda v: "R$ " + br(v))
    res["R$/kg médio"] = res["R$/kg médio"].map(lambda v: "R$ " + br(v))
    st.dataframe(res, hide_index=True, use_container_width=True)
    if destaque is not None and destaque in res["Vendedor"].values:
        pos = res.loc[res["Vendedor"] == destaque, "Posição"].iloc[0]
        st.caption(f"Sua posição: {pos} de {len(res)}")


def tela_login():
    st.title("📊 Painel de Vendas")
    vendedores = carregar_vendedores()
    nomes = list(vendedores.keys()) + ["GESTOR"]

    col, _ = st.columns([1, 1])
    with col:
        nome = st.selectbox("Quem é você?", nomes)
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
    aba_novo, aba_orc, aba_hoje, aba_rank = st.tabs(
        ["➕ Novo lançamento", "📄 Orçamentos em aberto",
         "📋 Meus lançamentos de hoje", "🏆 Ranking"]
    )

    # --- Novo lançamento ---
    with aba_novo:
        c1, c2 = st.columns(2)
        cliente = c1.text_input("Cliente *")
        tipo = c1.radio("Tipo de cliente *", ["Carteira", "Novo"], horizontal=True)
        cliente_novo = tipo == "Novo"
        contato = c2.text_input("Com quem falou")
        resultado = st.radio("Resultado do contato *", RESULTADOS, horizontal=True)

        kg = valor = 0.0
        if resultado != "Só contato":
            c1, c2, c3 = st.columns(3)
            kg = c1.number_input("Kg *", min_value=0.0, step=10.0, format="%.2f")
            valor = c2.number_input("Valor total (R$) *", min_value=0.0,
                                    step=100.0, format="%.2f")
            preco = valor / kg if kg else 0
            c3.metric("R$/kg (automático)", "R$ " + br(preco))

        if st.button("✅ Registrar", type="primary"):
            if not cliente.strip():
                st.error("Informe o nome do cliente.")
            elif resultado != "Só contato" and (kg <= 0 or valor <= 0):
                st.error("Para orçamento ou venda, informe Kg e Valor.")
            else:
                salvar_registro(nome, cliente, cliente_novo, contato,
                                resultado, kg, valor)
                msg = "Lançamento registrado!"
                if cliente_novo:
                    msg += f" Cliente novo '{cliente.strip()}' cadastrado."
                st.success(msg + " Data e hora gravadas automaticamente.")
                st.balloons()

    df = carregar_registros()
    hoje = agora().strftime("%d/%m/%Y")
    mes = agora().strftime("%m/%Y")

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
                c1.write(
                    f"**{reg['Hora']}** — {reg['Cliente']} — {reg['Resultado']} — "
                    f"{br(reg['Kg'])} kg — R$ {br(reg['Valor (R$)'])}"
                )
                chave = f"del_{reg['_linha']}"
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

    # --- Ranking ---
    with aba_rank:
        st.subheader(f"Hoje ({hoje})")
        tabela_ranking(resumir(df[df["Data"] == hoje]), destaque=nome)
        st.subheader(f"Mês ({mes})")
        tabela_ranking(resumir(df[df["_mes"] == mes]), destaque=nome)


def tela_gestor():
    st.title("Painel do Gestor")
    df = carregar_registros()
    hoje = agora().strftime("%d/%m/%Y")
    mes = agora().strftime("%m/%Y")
    df_hoje = df[df["Data"] == hoje]
    df_mes = df[df["_mes"] == mes]

    aba_dia, aba_mes, aba_orc, aba_funil, aba_dados = st.tabs(
        ["📅 Hoje", "📈 Mês", "💰 Orçamentos", "🔻 Funil", "🗂 Dados completos"]
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
        orc = df[df["Resultado"] == "Orçamento enviado"]
        abertos = orc[orc["Situação"] == SITUACAO_ABERTO]
        aprov = orc[orc["Situação"] == SITUACAO_APROVADO]
        perd = orc[orc["Situação"] == SITUACAO_PERDIDO]
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
        f_res = c2.multiselect("Resultado", RESULTADOS)
        dados = df.copy()
        if f_vend:
            dados = dados[dados["Vendedor"].isin(f_vend)]
        if f_res:
            dados = dados[dados["Resultado"].isin(f_res)]
        st.dataframe(dados[CABECALHO_REGISTROS], hide_index=True,
                     use_container_width=True)
        st.download_button(
            "⬇️ Baixar CSV",
            dados[CABECALHO_REGISTROS].to_csv(index=False, sep=";",
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
            st.rerun()

    if usuario == "GESTOR":
        tela_gestor()
    else:
        tela_vendedor(usuario)


if __name__ == "__main__":
    main()
