import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime, date
from io import BytesIO

DB = "uniformes.db"

FILIAIS_POR_SEGMENTO = {
    "MERCADO": [
        "M IVASKO GETÚLIO",
        "M MARI IRA",
        "M MARI SMS",
        "M IVASKO DEZENOVE",
        "M IVASKO NOCA",
        "M IVASKO VICENTE",
        "M MARI LDS",
        "M NEW",
        "M MARI PG02",
        "M MARI PG"
    ],
    "POSTOS": [
        "P S CAROLINE",
        "P C CALED IMBITUVA",
        "P C SALDANHA",
        "P C CALED XV",
        "P C CAROLINE FILIAL",
        "P C ALADIM",
        "P C PITANGA",
        "P C KENNEDY",
        "P C PALMEIRA ST FELICIDADE",
        "P C PGPOSECOL",
        "P C MASTER",
        "P C RVERNALHA",
        "P C MOTIVAÇÃO",
        "P C PALMEIRA",
        "P C POSECOL",
        "P S P CENTRAL",
        "P S P MANSA",
        "P C PRUDE",
        "P C TRAJANO",
        "P C VICENTE"
    ],
    "GÁS": [
        "T G LEVE GAS",
        "T G LEVE LAG",
        "T O GAS MAIS TB",
        "T O P GAS"
    ],
    "OUTROS": [
        "GI - SERVIÇOS",
        "Y Y GYMNAMIC",
        "U U KM LIVRE",
        "U U CEASA",
        "U U CD"
    ]
}

PECAS = [
    "CAMISETA",
    "CALÇA",
    "JAQUETA",
    "JAQUETA TÉRMICA",
    "BOTA",
    "AVENTAL",
    "OUTROS",
    "LUVAS VAQUETA",
    "LUVAS PU",
    "LUVAS RANHURADA",
    "LUVAS",
    "CAPA DE CHUVA"
]

STATUS = [
    "SOLICITADO",
    "APROVADO",
    "REPROVADO",
    "DISPONÍVEL PARA RETIRADA",
    "ENTREGUE",
    "ENVIADO"
]


def conectar():
    return sqlite3.connect(DB)


def criar_banco():
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS solicitacoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_solicitacao TEXT,
            nome TEXT,
            matricula TEXT,
            cargo TEXT,
            segmento TEXT,
            filial TEXT,
            status TEXT,
            data_retirada TEXT,
            observacao_rh TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS itens (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            solicitacao_id INTEGER,
            peca TEXT,
            quantidade INTEGER,
            tamanho TEXT,
            FOREIGN KEY (solicitacao_id) REFERENCES solicitacoes(id)
        )
    """)

    conn.commit()
    conn.close()


def salvar_solicitacao(nome, matricula, cargo, segmento, filial, itens):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO solicitacoes 
        (data_solicitacao, nome, matricula, cargo, segmento, filial, status, data_retirada, observacao_rh)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        datetime.now().strftime("%d/%m/%Y %H:%M"),
        nome.strip(),
        matricula.strip(),
        cargo.strip(),
        segmento,
        filial,
        "SOLICITADO",
        "",
        ""
    ))

    solicitacao_id = cursor.lastrowid

    for item in itens:
        cursor.execute("""
            INSERT INTO itens 
            (solicitacao_id, peca, quantidade, tamanho)
            VALUES (?, ?, ?, ?)
        """, (
            solicitacao_id,
            item["Peça"],
            int(item["Quantidade"]),
            item["Tamanho"].strip()
        ))

    conn.commit()
    conn.close()
    return solicitacao_id


def carregar_solicitacoes():
    conn = conectar()
    df = pd.read_sql_query("SELECT * FROM solicitacoes ORDER BY id DESC", conn)
    conn.close()
    return df


def carregar_itens(solicitacao_id):
    conn = conectar()
    df = pd.read_sql_query("""
        SELECT peca AS Peça, quantidade AS Quantidade, tamanho AS Tamanho
        FROM itens
        WHERE solicitacao_id = ?
    """, conn, params=(solicitacao_id,))
    conn.close()
    return df



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


def atualizar_status(solicitacao_id, status, data_retirada, observacao):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE solicitacoes
        SET status = ?, data_retirada = ?, observacao_rh = ?
        WHERE id = ?
    """, (
        status,
        data_retirada,
        observacao.strip(),
        solicitacao_id
    ))

    conn.commit()
    conn.close()


criar_banco()

st.set_page_config(page_title="Solicitação de Uniformes", layout="wide")

st.title("Solicitação de Uniformes")

menu = st.sidebar.radio(
    "Menu",
    ["Nova Solicitação", "Acompanhar Solicitação", "Área do RH"]
)

if menu == "Nova Solicitação":
    st.subheader("Nova Solicitação")

    if "itens_temp" not in st.session_state:
        st.session_state.itens_temp = []

    col1, col2 = st.columns(2)

    with col1:
        nome = st.text_input("Nome do colaborador")
        matricula = st.text_input("Matrícula")
        cargo = st.text_input("Funcao")

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
            placeholder="Ex.: Camisa social, mangote, protetor, outro item..."
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
            protocolo = salvar_solicitacao(
                nome,
                matricula,
                cargo,
                segmento,
                filial,
                st.session_state.itens_temp
            )
            st.session_state.itens_temp = []
            st.success(f"Solicitação finalizada com sucesso. Protocolo: #{protocolo}")

elif menu == "Acompanhar Solicitação":
    st.subheader("Acompanhar Solicitação")

    col1, col2 = st.columns(2)

    with col1:
        matricula_busca = st.text_input("Buscar por matrícula")

    with col2:
        protocolo_busca = st.text_input("Buscar por protocolo")

    if matricula_busca or protocolo_busca:
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

elif menu == "Área do RH":
    st.subheader("Área do RH")

    senha = st.text_input("Senha RH", type="password")

    if senha == "rh123":
        df = carregar_solicitacoes()

        col1, col2, col3 = st.columns(3)

        with col1:
            filtro_segmento = st.selectbox(
                "Filtrar por segmento",
                ["TODOS"] + list(FILIAIS_POR_SEGMENTO.keys())
            )

        with col2:
            filtro_status = st.selectbox(
                "Filtrar por status",
                ["TODOS"] + STATUS
            )

        with col3:
            if filtro_segmento != "TODOS":
                lista_filiais = FILIAIS_POR_SEGMENTO[filtro_segmento]
            else:
                lista_filiais = sorted({
                    filial
                    for filiais in FILIAIS_POR_SEGMENTO.values()
                    for filial in filiais
                })

            filtro_filial = st.selectbox(
                "Filtrar por filial",
                ["TODAS"] + lista_filiais
            )

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
