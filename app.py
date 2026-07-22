"""
Orçamentos perdidos por falta de vínculo — app web (Streamlit)

Cada pessoa do time abre o site, escolhe o cliente que quer auditar, e
clica em "Gerar relatório". O app consulta o Metabase, filtra os leads que
aquele cliente não recebeu (produto por produto) e devolve um Excel pra
baixar.

O login do Metabase NÃO é digitado na tela — fica guardado nos "Secrets"
do Streamlit Cloud (Settings > Secrets do app), como:

    metabase_username = "seu_usuario"
    metabase_password = "sua_senha"

Só quem administra o app vê isso. Ninguém que usa o link precisa saber
usuário/senha nenhum.

Além do login, quem hospeda o app precisa preencher a seção
"CONFIGURAÇÃO FIXA" lá embaixo, uma única vez.
"""
import copy
import io

import pandas as pd
import requests
import streamlit as st

# ============================================================
# CONFIGURAÇÃO FIXA — preencha isto uma vez só, antes de publicar o app.
# Isto NÃO é login de usuário, é a estrutura das consultas do Metabase
# (a mesma pra todo mundo que usar o app).
# ============================================================

METABASE_URL = "https://metabase.ferramentademarketing.com.br"

# Question 287 - "Produto por Cliente" (filtro: chave_unica)
PRODUTOS_CARD_ID = 287
PRODUTOS_PARAM_TEMPLATE = {
    "type": "category",
    "target": ["variable", ["template-tag", "chave_unica"]],
}
COLUNA_PRODUTO = "Produto"

# Question 39 - "Growth - Relatório de Orçamentos Únicos" (filtro: produto)
LEADS_CARD_ID = 39
LEADS_PARAM_TEMPLATE = {
    "type": "category",
    "target": ["variable", ["template-tag", "produto"]],
}
# Essa question exige outros parâmetros além de "produto" (aparecem na URL
# do Metabase: data_inicio, data_final, announcements, mensagem, satellite).
# data_inicio/data_final agora vêm da tela (período escolhido pelo usuário);
# os outros três não usamos pra filtrar, então mandamos vazios.
LEADS_EXTRA_PARAMS = [
    {"type": "category", "target": ["variable", ["template-tag", "announcements"]], "value": ""},
    {"type": "category", "target": ["variable", ["template-tag", "mensagem"]], "value": ""},
    {"type": "category", "target": ["variable", ["template-tag", "satellite"]], "value": ""},
]
COLUNA_EMPRESAS_QUE_RECEBERAM = "Empresas Recebedoras"
SEPARADOR_EMPRESAS = ","

# ============================================================
# Daqui pra baixo é lógica do app — não precisa editar.
# ============================================================


def login(username: str, password: str) -> str:
    resp = requests.post(
        f"{METABASE_URL.rstrip('/')}/api/session",
        json={"username": username, "password": password},
        timeout=30,
    )
    if resp.status_code != 200:
        raise RuntimeError(
            "Não consegui logar no Metabase com esse usuário/senha. "
            f"Detalhe técnico: {resp.status_code} - {resp.text}"
        )
    token = resp.json().get("id")
    if not token:
        raise RuntimeError("Login aceito, mas o Metabase não devolveu um token de sessão.")
    return token


def run_card(session_token: str, card_id: int, parameters: list) -> pd.DataFrame:
    resp = requests.post(
        f"{METABASE_URL.rstrip('/')}/api/card/{card_id}/query",
        headers={"X-Metabase-Session": session_token},
        json={"parameters": parameters},
        timeout=60,
    )
    if resp.status_code not in (200, 202):
        raise RuntimeError(
            f"Erro ao consultar a question {card_id} no Metabase: "
            f"{resp.status_code} - {resp.text}"
        )
    data = resp.json()
    rows = data.get("data", {}).get("rows", [])
    cols = [c["display_name"] for c in data.get("data", {}).get("cols", [])]
    if not rows:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame(rows, columns=cols)


def get_client_products(session_token: str, chave_cliente: str) -> list:
    param = copy.deepcopy(PRODUTOS_PARAM_TEMPLATE)
    param["value"] = chave_cliente
    df = run_card(session_token, PRODUTOS_CARD_ID, [param])
    if COLUNA_PRODUTO not in df.columns:
        raise RuntimeError(
            f"A coluna '{COLUNA_PRODUTO}' não existe na question de produtos. "
            f"Colunas encontradas: {list(df.columns)}"
        )
    return sorted(set(df[COLUNA_PRODUTO].dropna().astype(str).str.strip()))


def get_leads_not_received(
    session_token: str, produto: str, chave_cliente: str, data_inicio: str, data_final: str
) -> pd.DataFrame:
    param_produto = copy.deepcopy(LEADS_PARAM_TEMPLATE)
    param_produto["value"] = produto
    param_data_inicio = {
        "type": "date/single",
        "target": ["variable", ["template-tag", "data_inicio"]],
        "value": data_inicio,
    }
    param_data_final = {
        "type": "date/single",
        "target": ["variable", ["template-tag", "data_final"]],
        "value": data_final,
    }
    params = [param_produto, param_data_inicio, param_data_final] + copy.deepcopy(LEADS_EXTRA_PARAMS)
    df = run_card(session_token, LEADS_CARD_ID, params)
    if df.empty:
        return df
    if COLUNA_EMPRESAS_QUE_RECEBERAM not in df.columns:
        raise RuntimeError(
            f"A coluna '{COLUNA_EMPRESAS_QUE_RECEBERAM}' não existe na question de leads. "
            f"Colunas encontradas: {list(df.columns)}"
        )

    def recebeu(valor) -> bool:
        if pd.isna(valor):
            return False
        chaves = [v.strip() for v in str(valor).split(SEPARADOR_EMPRESAS)]
        return chave_cliente in chaves

    faltantes = df[~df[COLUNA_EMPRESAS_QUE_RECEBERAM].apply(recebeu)].copy()
    if not faltantes.empty:
        faltantes.insert(0, "Produto Consultado", produto)
    return faltantes


def gerar_excel(resultado: pd.DataFrame) -> bytes:
    resumo = (
        resultado.groupby("Produto Consultado")
        .size()
        .reset_index(name="Qtde Não Recebidos")
        .sort_values("Qtde Não Recebidos", ascending=False)
    )
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        resumo.to_excel(writer, sheet_name="Resumo por Produto", index=False)
        resultado.to_excel(writer, sheet_name="Leads Não Recebidos", index=False)
    return buffer.getvalue()


# ============================================================
# Interface
# ============================================================

st.set_page_config(page_title="Orçamentos perdidos por vínculo", page_icon="📋")
st.title("📋 Orçamentos perdidos por falta de vínculo")
st.caption(
    "Descubra quantos orçamentos um cliente deixou de receber, produto por produto, "
    "comparando com o que já foi entregue no Metabase."
)

if "metabase_username" not in st.secrets or "metabase_password" not in st.secrets:
    st.error(
        "Login do Metabase não configurado. Quem administra o app precisa cadastrar "
        "'metabase_username' e 'metabase_password' em Settings > Secrets."
    )
    st.stop()

st.subheader("Cliente a auditar")
chave_cliente = st.text_input("Chave única (ID) do cliente no Metabase")

st.subheader("Período a analisar")
col_data1, col_data2 = st.columns(2)
with col_data1:
    data_inicio = st.date_input("Data início", value=None, format="DD/MM/YYYY")
with col_data2:
    data_final = st.date_input("Data fim", value=None, format="DD/MM/YYYY")

periodo_ok = data_inicio and data_final

if st.button("Gerar relatório", type="primary", disabled=not (chave_cliente and periodo_ok)):
    try:
        with st.spinner("Fazendo login no Metabase..."):
            token = login(st.secrets["metabase_username"], st.secrets["metabase_password"])

        with st.spinner("Buscando produtos cadastrados do cliente..."):
            produtos = get_client_products(token, chave_cliente)

        if not produtos:
            st.warning("Não encontrei nenhum produto cadastrado para essa chave de cliente.")
        else:
            st.write(f"**{len(produtos)} produto(s) encontrado(s):** {', '.join(produtos)}")

            progress = st.progress(0.0)
            status_area = st.empty()
            resultados = []
            for i, produto in enumerate(produtos):
                status_area.text(f"Consultando leads do produto: {produto}")
                faltantes = get_leads_not_received(
                    token,
                    produto,
                    chave_cliente,
                    data_inicio.isoformat(),
                    data_final.isoformat(),
                )
                if not faltantes.empty:
                    resultados.append(faltantes)
                progress.progress((i + 1) / len(produtos))

            status_area.empty()
            progress.empty()

            if not resultados:
                st.success("Nenhum orçamento perdido encontrado para os produtos deste cliente. 🎉")
            else:
                resultado_final = pd.concat(resultados, ignore_index=True)
                st.success(f"{len(resultado_final)} orçamento(s) não recebido(s) encontrado(s).")
                st.dataframe(resultado_final, use_container_width=True)

                excel_bytes = gerar_excel(resultado_final)
                st.download_button(
                    "⬇️ Baixar Excel",
                    data=excel_bytes,
                    file_name=f"orcamentos_perdidos_{chave_cliente}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
    except RuntimeError as e:
        st.error(str(e))
