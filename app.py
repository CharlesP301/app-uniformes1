import streamlit as st
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from datetime import datetime, date, timedelta
from io import BytesIO

# ──────────────────────────────────────────────
# FILIAIS DE IRATI
# ──────────────────────────────────────────────

FILIAIS_IRATI = {
    "T G LEVE GAS", "T G LEVE LAG",
    "M NEW", "M IVASKO NOCA", "M IVASKO VICENTE",
    "M IVASKO DEZENOVE", "M IVASKO GETÚLIO", "M MARI IRA",
    "GI - SERVIÇOS", "U U KM LIVRE", "U U CD", "Y Y GYMNAMIC",
    "P C POSECOL", "P C TRAJANO", "P C MASTER", "P C VICENTE"
}


def calcular_data_entrega(data_solicitacao: datetime) -> date:
    """Retorna a data de entrega = data_solicitacao + 3 dias úteis (seg–sex)."""
    dias_uteis = 0
    data = data_solicitacao.date()
    while dias_uteis < 3:
        data += timedelta(days=1)
        if data.weekday() < 5:  # 0=seg ... 4=sex
            dias_uteis += 1
    return data

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
        "P S P CENTRAL", "P S P MANSA", "P C PRUDE", "P C TRAJANO", "P C VICENTE", "P C BOEING"
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

CABECALHO = [
    "protocolo", "data_solicitacao", "nome", "matricula", "cargo",
    "segmento", "filial", "peca", "quantidade", "tamanho",
    "status", "data_retirada", "observacao_rh"
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
    aba = client.open("Uniformes").worksheet("solicitacoes")
    return aba


def get_proximo_protocolo(aba):
    valores = aba.col_values(1)[1:]  # ignora cabeçalho
    ids = [int(v) for v in valores if str(v).strip().isdigit()]
    return max(ids, default=0) + 1


# ──────────────────────────────────────────────
# FUNÇÕES DE DADOS
# ──────────────────────────────────────────────

def salvar_solicitacao(nome, matricula, cargo, segmento, filial, itens):
    aba = conectar_sheets()
    protocolo = get_proximo_protocolo(aba)
    data = datetime.now().strftime("%d/%m/%Y %H:%M")

    linhas = []
    for item in itens:
        linhas.append([
            protocolo,
            data,
            nome.strip(),
            matricula.strip(),
            cargo.strip(),
            segmento,
            filial,
            item["Peça"],
            int(item["Quantidade"]),
            item["Tamanho"].strip(),
            "SOLICITADO",
            "",
            ""
        ])

    aba.append_rows(linhas)
    return protocolo


def carregar_solicitacoes():
    aba = conectar_sheets()
    dados = aba.get_all_records()
    if not dados:
        return pd.DataFrame(columns=CABECALHO)
    df = pd.DataFrame(dados)
    df["protocolo"] = pd.to_numeric(df["protocolo"], errors="coerce")
    return df.sort_values("protocolo", ascending=False).reset_index(drop=True)


def atualizar_status(protocolo, status, data_retirada, observacao):
    aba = conectar_sheets()
    protocolos = aba.col_values(1)[1:]  # ignora cabeçalho

    linhas_para_atualizar = [
        i + 2  # +2 = pula cabeçalho e ajusta índice base-1
        for i, v in enumerate(protocolos)
        if str(v).strip() == str(protocolo)
    ]

    for linha in linhas_para_atualizar:
        aba.update_cell(linha, 11, status)
        aba.update_cell(linha, 12, data_retirada)
        aba.update_cell(linha, 13, observacao.strip())


def gerar_excel(df):
    output = BytesIO()
    df_export = df.copy()
    df_export.columns = [
        "Protocolo", "Data Solicitação", "Nome", "Matrícula", "Função",
        "Segmento", "Filial", "Peça", "Quantidade", "Tamanho",
        "Status", "Data Retirada", "Observação RH"
    ]
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Solicitações")
    output.seek(0)
    return output


# ──────────────────────────────────────────────
# INTERFACE
# ──────────────────────────────────────────────

st.set_page_config(page_title="Solicitação de Uniformes", layout="wide")
st.title("Solicitação de Uniformes")

menu = st.sidebar.radio("Menu", ["Nova Solicitação", "Área do RH"])

# ── NOVA SOLICITAÇÃO ──────────────────────────
if menu == "Nova Solicitação":
    st.subheader("Nova Solicitação")

    if "itens_temp" not in st.session_state:
        st.session_state.itens_temp = []

    col1, col2 = st.columns(2)

    with col1:
        # ── ALTERAÇÃO: label, placeholder e help atualizado ──
        nome = st.text_input(
            "Nome completo do colaborador",
            placeholder="Ex.: Lucas Oliveira Silva",
            help="Informe nome e sobrenome para facilitar a identificação."
        )
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
        # ── ALTERAÇÃO: validação exige nome completo (mínimo 2 palavras) ──
        palavras_nome = [p for p in nome.strip().split() if p]
        if len(palavras_nome) < 2:
            st.error("Informe o nome completo do colaborador (nome e sobrenome).")
        elif not cargo.strip():
            st.error("Preencha o cargo.")
        elif not st.session_state.itens_temp:
            st.error("Adicione pelo menos uma peça antes de finalizar.")
        else:
            with st.spinner("Salvando solicitação..."):
                protocolo = salvar_solicitacao(
                    nome, matricula, cargo, segmento, filial,
                    st.session_state.itens_temp
                )
            st.session_state.itens_temp = []

            # ── AVISO DE CONFIRMAÇÃO ──────────────────────────
            agora = datetime.now()
            eh_irati = filial in FILIAIS_IRATI
            data_entrega = calcular_data_entrega(agora)

            if eh_irati:
                bloco_entrega = (
                    '<div style="background:#eff6ff; border-radius:8px; padding:0.85rem 1rem; font-size:0.88rem; color:#1e40af;">'
                    '🗓️ <strong>Previsão de entrega:</strong> ' + data_entrega.strftime("%d/%m/%Y") + ' (3 dias úteis)<br><br>'
                    '⏰ <strong>Horários exclusivos de retirada em Irati:</strong><br>'
                    '&nbsp;&nbsp;&nbsp;• <strong>08:30 às 09:00</strong><br>'
                    '&nbsp;&nbsp;&nbsp;• <strong>16:30 às 17:00</strong><br><br>'
                    'ℹ️ O atendimento ocorre <strong>exclusivamente</strong> nesses horários, mediante disponibilidade em estoque.'
                    '</div>'
                )
            else:
                bloco_entrega = (
                    '<div style="background:#fefce8; border-radius:8px; padding:0.85rem 1rem; font-size:0.88rem; color:#854d0e;">'
                    '🚚 <strong>Entrega via logística</strong><br><br>'
                    'Se disponível no estoque, seu pedido será encaminhado à sua unidade em <strong>7 a 15 dias úteis</strong> após a data da solicitação.'
                    '</div>'
                )

            html_aviso = (
                '<div style="border:1px solid #d1fae5; border-radius:12px; padding:1.5rem; margin-top:1.5rem;">'
                '<div style="display:flex; align-items:center; gap:12px; margin-bottom:1rem;">'
                '<span style="font-size:2rem;">✅</span>'
                '<div>'
                '<div style="font-size:1.1rem; font-weight:600; color:#065f46;">Solicitação enviada com sucesso!</div>'
                '<div style="font-size:0.9rem; color:#6b7280;">Protocolo: <strong>#' + str(protocolo) + '</strong> — ' + agora.strftime("%d/%m/%Y às %H:%M") + '</div>'
                '</div>'
                '</div>'
                '<hr style="border:none; border-top:1px solid #e5e7eb; margin:0.75rem 0;">'
                '<div style="font-size:0.9rem; color:#374151; margin-bottom:0.75rem;">'
                '<strong>Colaborador:</strong> ' + nome.strip() + '<br>'
                '<strong>Filial:</strong> ' + filial + ' — ' + segmento +
                '</div>'
                '<hr style="border:none; border-top:1px solid #e5e7eb; margin:0.75rem 0;">'
                + bloco_entrega +
                '</div>'
            )

            st.markdown(html_aviso, unsafe_allow_html=True)

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
            arquivo_excel = gerar_excel(df_filtrado)
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
            protocolos_unicos = sorted(df_filtrado["protocolo"].unique().tolist(), reverse=True)
            protocolo_sel = st.selectbox("Selecione o protocolo", protocolos_unicos)

            dados = df_filtrado[df_filtrado["protocolo"] == protocolo_sel].iloc[0]

            st.write(f"**Colaborador:** {dados['nome']}")
            st.write(f"**Segmento:** {dados['segmento']}")
            st.write(f"**Filial:** {dados['filial']}")
            st.write(f"**Status atual:** {dados['status']}")

            st.markdown("#### Itens")
            itens = df_filtrado[df_filtrado["protocolo"] == protocolo_sel][["peca", "quantidade", "tamanho"]].copy()
            itens.columns = ["Peça", "Quantidade", "Tamanho"]
            st.dataframe(itens, use_container_width=True, hide_index=True)

            novo_status = st.selectbox("Novo status", STATUS)
            data_retirada = st.date_input("Data de retirada", value=date.today())
            observacao = st.text_area("Observação RH")

            if st.button("Salvar atualização"):
                with st.spinner("Salvando..."):
                    atualizar_status(
                        protocolo_sel,
                        novo_status,
                        data_retirada.strftime("%d/%m/%Y"),
                        observacao
                    )
                st.success("Status atualizado com sucesso.")
                st.rerun()

    elif senha:
        st.error("Senha incorreta.")
