import streamlit as st
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from datetime import datetime, date
from io import BytesIO

# ──────────────────────────────────────────────
# CONFIGURAÇÕES
# ──────────────────────────────────────────────

FILIAIS_POR_SEGMENTO = {
    "MERCADO": [
        "M IVASKO GETÚLIO", "M MARI IRA", "M MARI SMS", "M IVASKO DEZENOVE",
        "M IVASKO NOCA", "M IVASKO VICENTE", "M MARI LDS", "M NEW",
        "M MARI PG02", "M MARI PG"
    ],
    "POSTOS": [
        "P S CAROLINE", "P C CALED IMBITUVA", "P C SALDANHA", "P C CALED XV",
        "P C CAROLINE FILIAL", "P C ALADIM", "P C PITANGA", "P C KENNEDY",
        "P C PALMEIRA ST FELICIDADE", "P C PGPOSECOL", "P C MASTER",
        "P C RVERNALHA", "P C MOTIVAÇÃO", "P C PALMEIRA", "P C POSECOL",
        "P S P CENTRAL", "P S P MANSA", "P C PRUDE", "P C TRAJANO", "P C VICENTE"
    ],
    "GÁS": [
        "T G LEVE GAS", "T G LEVE LAG", "T O GAS MAIS TB", "T O P GAS"
    ],
    "OUTROS": [
        "GI - SERVIÇOS", "Y Y GYMNAMIC", "U U KM LIVRE", "U U CEASA", "U U CD"
    ]
}

PECAS = [
    "CAMISETA", "CALÇA", "JAQUETA", "JAQUETA TÉRMICA", "BOTA", "AVENTAL",
    "OUTROS", "LUVAS VAQUETA", "LUVAS PU", "LUVAS RANHURADA", "LUVAS", "CAPA DE CHUVA"
]

STATUS = [
    "SOLICITADO", "APROVADO", "REPROVADO",
    "DISPONÍVEL PARA RETIRADA", "ENTREGUE", "ENVIADO"
]

# ──────────────────────────────────────────────
# CONEXÃO COM GOOGLE SHEETS
# ──────────────────────────────────────────────

@st.cache_resource
def conectar_sheets():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes
    )
    client = gspread.authorize(creds)
    planilha = client.open("Uniformes")
    aba_sol = planilha.worksheet("solicitacoes")
    aba_ite = planilha.worksheet("itens")
    return aba_sol, aba_ite


def get_proximo_id(aba):
    """Retorna o próximo ID disponível (máximo atual + 1)."""
    valores = aba.col_values(1)[1:]  # ignora cabeçalho
    ids = [int(v) for v in valores if v.strip().isdigit()]
    return max(ids, default=0) + 1


# ──────────────────────────────────────────────
# FUNÇÕES DE DADOS
# ──────────────────────────────────────────────

def salvar_solicitacao(nome, matricula, cargo, segmento, filial, itens):
    aba_sol, aba_ite = conectar_sheets()

    sol_id = get_proximo_id(aba_sol)

    aba_sol.append_row([
        sol_id,
        datetime.now().strftime("%d/%m/%Y %H:%M"),
        nome.strip(),
        matricula.strip(),
        cargo.strip(),
        segmento,
        filial,
        "SOLICITADO",
        "",
        ""
    ])

    item_id = get_proximo_id(aba_ite)
    for item in itens:
        aba_ite.append_row([
            item_id,
            sol_id,
            item["Peça"],
            int(item["Quantidade"]),
            item["Tamanho"].strip()
        ])
        item_id += 1

    return sol_id


def carregar_solicitacoes():
    aba_sol, _ = conectar_sheets()
    dados = aba_sol.get_all_records()
    if not dados:
        return pd.DataFrame(columns=[
            "id", "data_solicitacao", "nome", "matricula", "cargo",
            "segmento", "filial", "status", "data_retirada", "observacao_rh"
        ])
    df = pd.DataFrame(dados)
    df["id"] = pd.to_numeric(df["id"], errors="coerce")
    return df.sort_values("id", ascending=False).reset_index(drop=True)


def carregar_itens(solicitacao_id):
    _, aba_ite = conectar_sheets()
    dados = aba_ite.get_all_records()
    if not dados:
        return pd.DataFrame(columns=["Peça", "Quantidade", "Tamanho"])
    df = pd.DataFrame(dados)
    df["solicitacao_id"] = pd.to_numeric(df["solicitacao_id"], errors="coerce")
    df_filtrado = df[df["solicitacao_id"] == solicitacao_id][["peca", "quantidade", "tamanho"]]
    df_filtrado.columns = ["Peça", "Quantidade", "Tamanho"]
    return df_filtrado.reset_index(drop=True)


def atualizar_status(solicitacao_id, status, data_retirada, observacao):
    aba_sol, _ = conectar_sheets()
    ids = aba_sol.col_values(1)  # coluna A = id
    try:
        linha = ids.index(str(solicitacao_id)) + 1  # +1 porque Sheets começa em 1
    except ValueError:
        st.error("Solicitação não encontrada.")
        return

    # Colunas: id=1, data=2, nome=3, matricula=4, cargo=5, segmento=6,
    #          filial=7, status=8, data_retirada=9, observacao_rh=10
    aba_sol.update_cell(linha, 8, status)
    aba_sol.update_cell(linha, 9, data_retirada)
    aba_sol.update_cell(linha, 10, observacao.strip())


def gerar_excel_solicitacoes(df_solicitacoes):
    linhas = []

    for _, row in df_solicitacoes.iterrows():
        itens = carregar_itens(row["id"])

        if itens.empty:
            linhas.append({
                "Protocolo": row["id"],
                "Data Solicitação": row["data_solicitacao"],
                "Nome": row["nome"],
                "Matrícula": row["matricula"],
                "Função": row["cargo"],
                "Segmento": row["segmento"],
                "Filial": row["filial"],
                "Peça": "",
                "Quantidade": "",
                "Tamanho": "",
                "Status": row["status"],
                "Data Retirada": row["data_retirada"],
                "Observação RH": row["observacao_rh"]
            })
        else:
            for _, item in itens.iterrows():
                linhas.append({
                    "Protocolo": row["id"],
                    "Data Solicitação": row["data_solicitacao"],
                    "Nome": row["nome"],
                    "Matrícula": row["matricula"],
                    "Função": row["cargo"],
                    "Segmento": row["segmento"],
                    "Filial": row["filial"],
                    "Peça": item["Peça"],
                    "Quantidade": item["Quantidade"],
                    "Tamanho": item["Tamanho"],
                    "Status": row["status"],
                    "Data Retirada": row["data_retirada"],
                    "Observação RH": row["observacao_rh"]
                })

    df_excel = pd.DataFrame(linhas)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_excel.to_excel(writer, index=False, sheet_name="Solicitações")
    output.seek(0)
    return output


# ──────────────────────────────────────────────
# INTERFACE
# ──────────────────────────────────────────────

st.set_page_config(page_title="Solicitação de Uniformes", layout="wide")
st.title("Solicitação de Uniformes")

menu = st.sidebar.radio("Menu", ["Nova Solicitação", "Acompanhar Solicitação", "Área do RH"])

# ── NOVA SOLICITAÇÃO ──────────────────────────
if menu == "Nova Solicitação":
    st.subheader("Nova Solicitação")

    if "itens_temp" not in st.session_state:
        st.session_state.itens_temp = []

    col1, col2 = st.columns(2)

    with col1:
        nome = st.text_input("Nome do colaborador")
        matricula = st.text_input("Matrícula")
        cargo = st.text_input("Função")

    with col2:
        segmento = st.selectbox("Segmento", list(FILIAIS_POR_SEGMENTO.keys()))
        filial = st.selectbox("Filial", FILIAIS_POR_SEGMENTO[segmento])

    st.markdown("### Adicionar peças")

    c1, c2, c3 = st.columns(3)

    with c1:
        peca = st.selectbox("Peça", PECAS)
    with c2:
        quantidade = st.number_input("Quantidade", min_value=1, max_value=50, value=1)
    with c3:
        tamanho = st.text_input("Tamanho", placeholder="Ex.: P, M, G, GG, 40, 41")

    descricao_outros = ""
    if peca == "OUTROS":
        descricao_outros = st.text_input(
            "Descreva a peça desejada",
            placeholder="Ex.: Camisa social, mangote, protetor..."
        )

    if st.button("Adicionar peça na solicitação"):
        if peca == "OUTROS" and not descricao_outros.strip():
            st.error("Descreva a peça desejada no campo OUTROS.")
        elif not tamanho.strip():
            st.error("Informe o tamanho da peça.")
        else:
            peca_final = descricao_outros.strip().upper() if peca == "OUTROS" else peca
            st.session_state.itens_temp.append({
                "Peça": peca_final,
                "Quantidade": quantidade,
                "Tamanho": tamanho
            })
            st.success("Peça adicionada.")

    st.markdown("### Peças adicionadas")

    if st.session_state.itens_temp:
        df_itens = pd.DataFrame(st.session_state.itens_temp)
        st.dataframe(df_itens, use_container_width=True, hide_index=True)

        item_remover = st.number_input(
            "Número do item para remover",
            min_value=1,
            max_value=len(st.session_state.itens_temp),
            value=1
        )

        if st.button("Remover item selecionado"):
            st.session_state.itens_temp.pop(item_remover - 1)
            st.rerun()
    else:
        st.info("Nenhuma peça adicionada ainda.")

    st.markdown("---")

    if st.button("Finalizar Solicitação"):
        if not nome.strip() or not cargo.strip():
            st.error("Preencha nome e cargo.")
        elif not st.session_state.itens_temp:
            st.error("Adicione pelo menos uma peça antes de finalizar.")
        else:
            with st.spinner("Salvando solicitação..."):
                protocolo = salvar_solicitacao(
                    nome, matricula, cargo, segmento, filial,
                    st.session_state.itens_temp
                )
            st.session_state.itens_temp = []
            st.success(f"Solicitação finalizada com sucesso. Protocolo: #{protocolo}")

# ── ACOMPANHAR SOLICITAÇÃO ────────────────────
elif menu == "Acompanhar Solicitação":
    st.subheader("Acompanhar Solicitação")

    col1, col2 = st.columns(2)

    with col1:
        matricula_busca = st.text_input("Buscar por matrícula")
    with col2:
        protocolo_busca = st.text_input("Buscar por protocolo")

    if matricula_busca or protocolo_busca:
        with st.spinner("Carregando..."):
            df = carregar_solicitacoes()

        if protocolo_busca.strip():
            try:
                df = df[df["id"] == int(protocolo_busca)]
            except ValueError:
                df = df.iloc[0:0]
        else:
            df = df[df["matricula"].astype(str) == matricula_busca.strip()]

        if df.empty:
            st.warning("Nenhuma solicitação encontrada.")
        else:
            for _, row in df.iterrows():
                st.markdown("---")
                st.markdown(f"### Protocolo #{row['id']}")
                st.metric("Status RH", row["status"])
                st.write(f"**Data:** {row['data_solicitacao']}")
                st.write(f"**Colaborador:** {row['nome']}")
                st.write(f"**Matrícula:** {row['matricula']}")
                st.write(f"**Cargo:** {row['cargo']}")
                st.write(f"**Segmento:** {row['segmento']}")
                st.write(f"**Filial:** {row['filial']}")

                if row["data_retirada"]:
                    st.write(f"**Data de retirada:** {row['data_retirada']}")
                if row["observacao_rh"]:
                    st.info(f"Observação RH: {row['observacao_rh']}")

                st.markdown("**Itens solicitados:**")
                st.dataframe(carregar_itens(row["id"]), use_container_width=True, hide_index=True)

# ── ÁREA DO RH ────────────────────────────────
elif menu == "Área do RH":
    st.subheader("Área do RH")

    senha = st.text_input("Senha RH", type="password")

    if senha == "rh123":
        with st.spinner("Carregando solicitações..."):
            df = carregar_solicitacoes()

        col1, col2, col3 = st.columns(3)

        with col1:
            filtro_segmento = st.selectbox(
                "Filtrar por segmento",
                ["TODOS"] + list(FILIAIS_POR_SEGMENTO.keys())
            )
        with col2:
            filtro_status = st.selectbox("Filtrar por status", ["TODOS"] + STATUS)
        with col3:
            if filtro_segmento != "TODOS":
                lista_filiais = FILIAIS_POR_SEGMENTO[filtro_segmento]
            else:
                lista_filiais = sorted({
                    filial
                    for filiais in FILIAIS_POR_SEGMENTO.values()
                    for filial in filiais
                })
            filtro_filial = st.selectbox("Filtrar por filial", ["TODAS"] + lista_filiais)

        st.markdown("### Filtro por data da solicitação")
        col_data1, col_data2 = st.columns(2)

        with col_data1:
            filtro_data_inicio = st.date_input("Data inicial", value=date.today())
        with col_data2:
            filtro_data_fim = st.date_input("Data final", value=date.today())

        df_filtrado = df.copy()

        if filtro_segmento != "TODOS":
            df_filtrado = df_filtrado[df_filtrado["segmento"] == filtro_segmento]
        if filtro_status != "TODOS":
            df_filtrado = df_filtrado[df_filtrado["status"] == filtro_status]
        if filtro_filial != "TODAS":
            df_filtrado = df_filtrado[df_filtrado["filial"] == filtro_filial]

        if not df_filtrado.empty:
            df_filtrado["_data_filtro"] = pd.to_datetime(
                df_filtrado["data_solicitacao"],
                format="%d/%m/%Y %H:%M",
                errors="coerce"
            ).dt.date

            df_filtrado = df_filtrado[
                (df_filtrado["_data_filtro"] >= filtro_data_inicio) &
                (df_filtrado["_data_filtro"] <= filtro_data_fim)
            ]
            df_filtrado = df_filtrado.drop(columns=["_data_filtro"])

        st.markdown("### Solicitações")
        st.dataframe(df_filtrado, use_container_width=True, hide_index=True)

        if not df_filtrado.empty:
            arquivo_excel = gerar_excel_solicitacoes(df_filtrado)
            st.download_button(
                label="Baixar solicitações filtradas em Excel",
                data=arquivo_excel,
                file_name="solicitacoes_uniformes_filtradas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.markdown("### Atualizar Status")

        if df_filtrado.empty:
            st.warning("Nenhuma solicitação encontrada.")
        else:
            solicitacao_id = st.selectbox(
                "Selecione o protocolo",
                df_filtrado["id"].tolist()
            )

            dados = df_filtrado[df_filtrado["id"] == solicitacao_id].iloc[0]

            st.write(f"**Colaborador:** {dados['nome']}")
            st.write(f"**Segmento:** {dados['segmento']}")
            st.write(f"**Filial:** {dados['filial']}")
            st.write(f"**Status atual:** {dados['status']}")

            st.markdown("#### Itens")
            st.dataframe(carregar_itens(solicitacao_id), use_container_width=True, hide_index=True)

            novo_status = st.selectbox("Novo status", STATUS)
            data_retirada = st.date_input("Data de retirada", value=date.today())
            observacao = st.text_area("Observação RH")

            if st.button("Salvar atualização"):
                with st.spinner("Salvando..."):
                    atualizar_status(
                        solicitacao_id,
                        novo_status,
                        data_retirada.strftime("%d/%m/%Y"),
                        observacao
                    )
                st.success("Status atualizado com sucesso.")
                st.rerun()

    elif senha:
        st.error("Senha incorreta.")
