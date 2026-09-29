import streamlit as st
from openpyxl import load_workbook
from io import BytesIO

# ============================================================
# CONFIGURAÇÃO
# ============================================================

st.set_page_config(
    page_title="Vindex | Preenchimento de Quantidades",
    page_icon="📦",
    layout="wide"
)

# ============================================================
# TÍTULO
# ============================================================

st.title("📦 Vindex — Preenchimento de Quantidades")

st.write(
    "Envie a planilha Excel e clique em **Processar preenchimento**. "
    "O sistema fará automaticamente o preenchimento da segunda aba."
)

# ============================================================
# UPLOAD
# ============================================================

uploaded_file = st.file_uploader(
    "📂 Envie sua planilha Excel",
    type=["xlsx"]
)

# ============================================================
# PROCESSAMENTO
# ============================================================

if uploaded_file is not None:

    st.success(
        f"Planilha carregada: **{uploaded_file.name}**"
    )

    if st.button(
        "🚀 Processar preenchimento",
        type="primary",
        use_container_width=True
    ):

        try:

            # ------------------------------------------------
            # ABRIR EXCEL
            # ------------------------------------------------

            file_bytes = uploaded_file.getvalue()

            wb = load_workbook(
                BytesIO(file_bytes)
            )

            # ------------------------------------------------
            # VERIFICAR SEGUNDA ABA
            # ------------------------------------------------

            if len(wb.sheetnames) < 2:

                st.error(
                    "❌ A planilha precisa ter pelo menos 2 abas."
                )

                st.stop()

            # ------------------------------------------------
            # USAR AUTOMATICAMENTE A SEGUNDA ABA
            # ------------------------------------------------

            ws = wb[wb.sheetnames[1]]

            # ------------------------------------------------
            # CONFIGURAÇÃO
            # ------------------------------------------------

            # J = quantidade prevista
            COLUNA_QUANTIDADE = 10

            # M = primeira caixa
            COLUNA_INICIO = 13

            # ------------------------------------------------
            # VARIÁVEIS
            # ------------------------------------------------

            # Esta variável é MUITO importante.
            #
            # Ela controla a sequência das caixas para
            # TODOS os produtos.
            #
            # Não reinicia em M a cada linha.

            coluna_atual = COLUNA_INICIO

            total_produtos = 0
            total_unidades = 0

            # ------------------------------------------------
            # PERCORRER AS LINHAS
            # ------------------------------------------------

            for row in range(
                1,
                ws.max_row + 1
            ):

                quantidade = ws.cell(
                    row=row,
                    column=COLUNA_QUANTIDADE
                ).value

                # Ignorar células vazias
                if quantidade is None:
                    continue

                # Tentar transformar em número
                try:

                    quantidade = int(
                        float(quantidade)
                    )

                except (
                    ValueError,
                    TypeError
                ):

                    continue

                # Ignorar zero ou negativos
                if quantidade <= 0:
                    continue

                # ------------------------------------------------
                # LIMPAR O PREENCHIMENTO ANTERIOR DA LINHA
                # ------------------------------------------------

                # Isso evita que uma planilha que já tenha
                # algum preenchimento fique com dados sobrando.

                for coluna in range(
                    COLUNA_INICIO,
                    ws.max_column + 1
                ):

                    ws.cell(
                        row=row,
                        column=coluna
                    ).value = None

                # ------------------------------------------------
                # PREENCHER AS UNIDADES
                # ------------------------------------------------

                for i in range(quantidade):

                    coluna_destino = (
                        coluna_atual + i
                    )

                    ws.cell(
                        row=row,
                        column=coluna_destino
                    ).value = 1

                # ------------------------------------------------
                # AVANÇAR PARA O PRÓXIMO PRODUTO
                # ------------------------------------------------

                coluna_atual += quantidade

                total_produtos += 1
                total_unidades += quantidade

            # ====================================================
            # GERAR NOVO EXCEL
            # ====================================================

            output = BytesIO()

            wb.save(output)

            output.seek(0)

            # ====================================================
            # RESULTADO
            # ====================================================

            st.success(
                "✅ Preenchimento concluído com sucesso!"
            )

            col1, col2 = st.columns(2)

            with col1:

                st.metric(
                    "Produtos processados",
                    total_produtos
                )

            with col2:

                st.metric(
                    "Total de unidades",
                    total_unidades
                )

            # ====================================================
            # NOME DO ARQUIVO
            # ====================================================

            nome_original = uploaded_file.name.rsplit(
                ".",
                1
            )[0]

            nome_saida = (
                f"{nome_original}_preenchida.xlsx"
            )

            # ====================================================
            # DOWNLOAD
            # ====================================================

            st.download_button(
                label="⬇️ Baixar planilha preenchida",

                data=output.getvalue(),

                file_name=nome_saida,

                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),

                use_container_width=True
            )

        except Exception as e:

            st.error(
                f"❌ Não foi possível processar a planilha:\n\n{e}"
            )

# ============================================================
# TELA INICIAL
# ============================================================

else:

    st.info(
        "👆 Envie a planilha acima para começar."
    )
