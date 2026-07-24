"""
Orçamentos perdidos por falta de vínculo — app web (Streamlit)

Cada pessoa do time abre o site, escolhe o cliente que quer auditar, o
período e a região de atuação dele, e clica em "Gerar relatório". O app
consulta o Metabase, filtra os leads que aquele cliente não recebeu
(produto por produto, e anúncio por anúncio) e devolve um Excel pra baixar.

O login do Metabase NÃO é digitado na tela — fica guardado nos "Secrets"
do Streamlit Cloud (Settings > Secrets do app), como:

    metabase_username = "seu_usuario"
    metabase_password = "sua_senha"

Só quem administra o app vê isso. Ninguém que usa o link precisa saber
usuário/senha nenhum.

Além do login, quem hospeda o app precisa preencher a seção
"CONFIGURAÇÃO FIXA" lá embaixo, uma única vez.

Regras de negócio aplicadas (definidas em conversa com o time):
1. Só entram na análise produtos do cliente com pelo menos 1 anúncio ativo
   (Status do Anuncio = 'Ativo' na question 287) — produto sem anúncio
   ativo é ignorado, porque não faz sentido cobrar por algo que nem estava
   sendo anunciado.
2. Para cada produto, a busca de leads é feita duas vezes na question 39:
   uma filtrando pelo campo "produto", outra filtrando pelo campo
   "announcements" (anúncio) com o mesmo nome — porque a classificação de
   Produto do lead pode não bater com o nome exato do anúncio que gerou
   aquele lead. Os dois resultados são unidos e deduplicados pelo
   Orçamento ID.
3. Filtro de região: o analista escolhe em quais estados/regiões o cliente
   atua (Nacional, uma ou mais das 5 regiões, e/ou estados específicos).
   Leads de fora dessa cobertura não contam como "perdidos", porque o
   cliente nunca atenderia ali de qualquer forma. A comparação usa a
   coluna "region" (sigla), convertendo a seleção da tela (nomes por
   extenso) pra sigla internamente.
4. Palavras bloqueadas: campo de texto onde o analista pode digitar
   palavras (separadas por vírgula) que, se aparecerem no texto da coluna
   Anúncio, tiram aquele lead da lista de perdidos — para casos em que o
   cliente está numa categoria de produto mas não trabalha com uma
   variação específica dela (ex: vende "sacola" mas não sacola reciclável
   nem de plástico).
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
COLUNA_STATUS_ANUNCIO = "Status do Anuncio"
STATUS_ATIVO = "ativo"  # comparação é feita em minúsculo, sem sensibilidade a maiúscula

# Question 39 - "Growth - Relatório de Orçamentos Únicos"
LEADS_CARD_ID = 39
# Nomes dos template-tags dessa question, vistos na URL do Metabase:
# data_inicio, data_final, produto, announcements, mensagem, satellite
TAG_PRODUTO = "produto"
TAG_ANUNCIO = "announcements"
TAG_MENSAGEM = "mensagem"
TAG_SATELITE = "satellite"
TAG_DATA_INICIO = "data_inicio"
TAG_DATA_FINAL = "data_final"

COLUNA_ORCAMENTO_ID = "Orçamento ID"
COLUNA_ANUNCIO = "Anúncio"
COLUNA_EMPRESAS_QUE_RECEBERAM = "Empresas Recebedoras"
SEPARADOR_EMPRESAS = ","
COLUNA_UF = "region"  # sigla do estado (ex: "BA", "SP")

# ============================================================
# Tabela fixa de estados e regiões do Brasil (não vem do Metabase).
# ============================================================

ESTADOS = [
    ("Acre", "AC", "Norte"),
    ("Amapá", "AP", "Norte"),
    ("Amazonas", "AM", "Norte"),
    ("Pará", "PA", "Norte"),
    ("Rondônia", "RO", "Norte"),
    ("Roraima", "RR", "Norte"),
    ("Tocantins", "TO", "Norte"),
    ("Alagoas", "AL", "Nordeste"),
    ("Bahia", "BA", "Nordeste"),
    ("Ceará", "CE", "Nordeste"),
    ("Maranhão", "MA", "Nordeste"),
    ("Paraíba", "PB", "Nordeste"),
    ("Pernambuco", "PE", "Nordeste"),
    ("Piauí", "PI", "Nordeste"),
    ("Rio Grande do Norte", "RN", "Nordeste"),
    ("Sergipe", "SE", "Nordeste"),
    ("Distrito Federal", "DF", "Centro-Oeste"),
    ("Goiás", "GO", "Centro-Oeste"),
    ("Mato Grosso", "MT", "Centro-Oeste"),
    ("Mato Grosso do Sul", "MS", "Centro-Oeste"),
    ("Espírito Santo", "ES", "Sudeste"),
    ("Minas Gerais", "MG", "Sudeste"),
    ("Rio de Janeiro", "RJ", "Sudeste"),
    ("São Paulo", "SP", "Sudeste"),
    ("Paraná", "PR", "Sul"),
    ("Rio Grande do Sul", "RS", "Sul"),
    ("Santa Catarina", "SC", "Sul"),
]
REGIOES_ORDEM = ["Norte", "Nordeste", "Centro-Oeste", "Sudeste", "Sul"]
NACIONAL = "Nacional"

OPCOES_REGIAO = (
    [NACIONAL]
    + REGIOES_ORDEM
    + [nome for nome, _sigla, _regiao in sorted(ESTADOS, key=lambda e: e[0])]
)


def resolver_ufs(selecionados: list) -> set | None:
    """Converte o que foi marcado na tela (Nacional/Região/Estado, por
    extenso) num conjunto de siglas de UF pra filtrar. Devolve None quando
    não deve filtrar nada (Nacional selecionado, ou nada selecionado)."""
    if not selecionados or NACIONAL in selecionados:
        return None
    ufs = set()
    for escolha in selecionados:
        if escolha in REGIOES_ORDEM:
            ufs.update(sigla for _nome, sigla, regiao in ESTADOS if regiao == escolha)
        else:
            match = next((sigla for nome, sigla, _regiao in ESTADOS if nome == escolha), None)
            if match:
                ufs.add(match)
    return ufs


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
    """Produtos cadastrados do cliente que têm pelo menos 1 anúncio ativo.
    Produto sem nenhum anúncio ativo é descartado (regra 1)."""
    param = copy.deepcopy(PRODUTOS_PARAM_TEMPLATE)
    param["value"] = chave_cliente
    df = run_card(session_token, PRODUTOS_CARD_ID, [param])

    faltando = [c for c in (COLUNA_PRODUTO, COLUNA_STATUS_ANUNCIO) if c not in df.columns]
    if faltando:
        raise RuntimeError(
            f"Coluna(s) {faltando} não encontrada(s) na question de produtos. "
            f"Colunas encontradas: {list(df.columns)}"
        )

    df["_status_norm"] = df[COLUNA_STATUS_ANUNCIO].astype(str).str.strip().str.lower()
    produtos_com_ativo = df.loc[df["_status_norm"] == STATUS_ATIVO, COLUNA_PRODUTO]
    return sorted(set(produtos_com_ativo.dropna().astype(str).str.strip()))


def _parametro(tag: str, valor: str) -> dict:
    return {"type": "category", "target": ["variable", ["template-tag", tag]], "value": valor}


def _buscar_leads_por_campo(
    session_token: str, tag_filtro: str, valor: str, data_inicio: str, data_final: str
) -> pd.DataFrame:
    """Roda a question 39 filtrando por UM campo (produto OU anúncio),
    deixando o outro e os demais filtros de texto vazios."""
    params = [
        _parametro(TAG_PRODUTO, valor if tag_filtro == TAG_PRODUTO else ""),
        _parametro(TAG_ANUNCIO, valor if tag_filtro == TAG_ANUNCIO else ""),
        _parametro(TAG_MENSAGEM, ""),
        _parametro(TAG_SATELITE, ""),
        {
            "type": "date/single",
            "target": ["variable", ["template-tag", TAG_DATA_INICIO]],
            "value": data_inicio,
        },
        {
            "type": "date/single",
            "target": ["variable", ["template-tag", TAG_DATA_FINAL]],
            "value": data_final,
        },
    ]
    return run_card(session_token, LEADS_CARD_ID, params)


def get_leads_not_received(
    session_token: str,
    produto: str,
    chave_cliente: str,
    data_inicio: str,
    data_final: str,
    ufs_permitidas: set | None,
    palavras_bloqueadas: list,
) -> pd.DataFrame:
    # Regra 2: duas passadas — por produto e por anúncio — juntando e
    # removendo duplicata pelo Orçamento ID.
    df_por_produto = _buscar_leads_por_campo(session_token, TAG_PRODUTO, produto, data_inicio, data_final)
    df_por_anuncio = _buscar_leads_por_campo(session_token, TAG_ANUNCIO, produto, data_inicio, data_final)
    df = pd.concat([df_por_produto, df_por_anuncio], ignore_index=True)
    if df.empty:
        return df

    if COLUNA_ORCAMENTO_ID in df.columns:
        df = df.drop_duplicates(subset=COLUNA_ORCAMENTO_ID)

    faltando = [c for c in (COLUNA_EMPRESAS_QUE_RECEBERAM, COLUNA_ANUNCIO, COLUNA_UF) if c not in df.columns]
    if faltando:
        raise RuntimeError(
            f"Coluna(s) {faltando} não encontrada(s) na question de leads. "
            f"Colunas encontradas: {list(df.columns)}"
        )

    def recebeu(valor) -> bool:
        if pd.isna(valor):
            return False
        chaves = [v.strip() for v in str(valor).split(SEPARADOR_EMPRESAS)]
        return chave_cliente in chaves

    faltantes = df[~df[COLUNA_EMPRESAS_QUE_RECEBERAM].apply(recebeu)].copy()

    # Regra 3: só considera dentro da cobertura de região do cliente.
    if ufs_permitidas is not None and not faltantes.empty:
        faltantes = faltantes[faltantes[COLUNA_UF].isin(ufs_permitidas)]

    # Regra 4: remove leads cujo Anúncio contenha alguma palavra bloqueada.
    if palavras_bloqueadas and not faltantes.empty:
        def tem_palavra_bloqueada(anuncio) -> bool:
            texto = "" if pd.isna(anuncio) else str(anuncio).lower()
            return any(p in texto for p in palavras_bloqueadas)

        faltantes = faltantes[~faltantes[COLUNA_ANUNCIO].apply(tem_palavra_bloqueada)]

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

st.subheader("Onde o cliente atua")
regioes_selecionadas = st.multiselect(
    "Nacional, região(ões) e/ou estado(s) específico(s)",
    options=OPCOES_REGIAO,
    default=[NACIONAL],
    help="Leads fora dessa cobertura não entram na lista de perdidos.",
)

st.subheader("Palavras bloqueadas (opcional)")
palavras_texto = st.text_input(
    "Palavras que, se aparecerem no Anúncio, tiram o lead da lista (separadas por vírgula)",
    placeholder="ex: reciclável, plástico",
)

periodo_ok = data_inicio and data_final

if st.button("Gerar relatório", type="primary", disabled=not (chave_cliente and periodo_ok)):
    try:
        ufs_permitidas = resolver_ufs(regioes_selecionadas)
        palavras_bloqueadas = [p.strip().lower() for p in palavras_texto.split(",") if p.strip()]

        with st.spinner("Fazendo login no Metabase..."):
            token = login(st.secrets["metabase_username"], st.secrets["metabase_password"])

        with st.spinner("Buscando produtos ativos cadastrados do cliente..."):
            produtos = get_client_products(token, chave_cliente)

        if not produtos:
            st.warning(
                "Não encontrei nenhum produto com anúncio ativo para essa chave de cliente."
            )
        else:
            st.write(f"**{len(produtos)} produto(s) com anúncio ativo:** {', '.join(produtos)}")

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
                    ufs_permitidas,
                    palavras_bloqueadas,
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
