import io
import os
from datetime import datetime

import pandas as pd
import streamlit as st
from fpdf import FPDF
from unidecode import unidecode

# ========= Leitura robusta de planilhas =========
def read_excel_auto(file, filename_hint: str = ""):
    """
    Tenta abrir com openpyxl (xlsx), cai para xlrd (xls),
    e trata o caso .xls salvo com extensão .xlsx (OLE2).
    Aceita 'file' como bytes (UploadedFile) ou file-like.
    """
    # Detecta extensão a partir do nome (se houver)
    ext = ""
    if filename_hint:
        ext = os.path.splitext(filename_hint)[1].lower()

    # Normaliza input para BytesIO
    if hasattr(file, "read"):
        raw = file.read()
    else:
        raw = file
    bio = io.BytesIO(raw)

    # Fluxos de tentativa
    try:
        if ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
            try:
                return pd.read_excel(io.BytesIO(raw), engine="openpyxl")
            except Exception as e:
                err = str(e).lower()
                if "zip file" in err or "ole2" in err:
                    # .xls disfarçado de .xlsx: tenta xlrd
                    try:
                        return pd.read_excel(io.BytesIO(raw), engine="xlrd")
                    except Exception as e2:
                        raise RuntimeError(
                            "Esse arquivo parece ser .xls salvo com extensão .xlsx. "
                            "Abra no Excel e 'Salvar como' .xlsx."
                        ) from e2
                else:
                    raise
        elif ext == ".xls":
            try:
                return pd.read_excel(io.BytesIO(raw), engine="xlrd")
            except Exception as e:
                raise RuntimeError(
                    "Arquivo .xls (formato antigo) não pôde ser carregado via xlrd."
                ) from e
        elif ext == ".ods":
            try:
                return pd.read_excel(io.BytesIO(raw), engine="odf")
            except Exception as e:
                raise RuntimeError("Falha ao abrir ODS via odf.") from e
        else:
            # Sem extensão (ou desconhecida): tenta openpyxl -> xlrd -> default
            try:
                return pd.read_excel(io.BytesIO(raw), engine="openpyxl")
            except:
                try:
                    return pd.read_excel(io.BytesIO(raw), engine="xlrd")
                except:
                    return pd.read_excel(io.BytesIO(raw))
    except Exception as e:
        raise RuntimeError(f"Não foi possível abrir '{filename_hint or 'arquivo'}': {e}")

# ========= Utilidades de UI/Export =========
def exportar_pdf(texto_volumetria: str, texto_resumo: str) -> bytes:
    pdf = FPDF()
    pdf.set_font("Arial", size=12)

    if texto_volumetria.strip():
        pdf.add_page()
        pdf.multi_cell(0, 10, txt="Relatório de Volumetria:\n".encode('latin-1','replace').decode('latin-1'))
        for line in texto_volumetria.splitlines():
            pdf.multi_cell(0, 10, txt=line.encode('latin-1','replace').decode('latin-1'))

    if texto_resumo.strip():
        pdf.add_page()
        pdf.multi_cell(0, 10, txt="Resumo Total de Serviços:\n".encode('latin-1','replace').decode('latin-1'))
        for line in texto_resumo.splitlines():
            pdf.multi_cell(0, 10, txt=line.encode('latin-1','replace').decode('latin-1'))

    mem = io.BytesIO()
    pdf.output(mem)
    return mem.getvalue()

st.set_page_config(page_title="Leitor/Analisador de Volume - Web", layout="wide")

# ===== Sidebar: Uploads e Filtros =====
st.sidebar.title("Leitor Supremo Web")
st.sidebar.caption("Envie as planilhas e aplique filtros")

volumetria_file = st.sidebar.file_uploader("Volumetria (.xlsx/.xls/.ods)", type=["xlsx","xls","ods"], key="volumetria")
usuarios_file = st.sidebar.file_uploader("USUARIOS_SICOOB (.xlsx/.xls/.ods)", type=["xlsx","xls","ods"], key="usuarios")

st.sidebar.markdown("---")
st.sidebar.subheader("Compilador (opcional)")
comp1_file = st.sidebar.file_uploader("Arquivo 1", type=["xlsx","xls","ods"], key="comp1")
comp2_file = st.sidebar.file_uploader("Arquivo 2", type=["xlsx","xls","ods"], key="comp2")
usar_usuarios_compilador = st.sidebar.checkbox("Filtrar apenas USUARIOS_SICOOB autorizados", value=False)
usuarios_compilador_file = st.sidebar.file_uploader("Planilha USUARIOS_SICOOB p/ filtro", type=["xlsx","xls","ods"], key="usuarios_comp")

st.sidebar.markdown("---")
processar_btn = st.sidebar.button("Processar Volumetria")
compilar_btn = st.sidebar.button("Compilar Arquivos")

# ===== Área principal (abas) =====
tab1, tab2, tab3 = st.tabs(["📊 Volumetria", "🧮 Resumo do Time", "🧩 Compilador"])

# Estado
if "df_usuarios" not in st.session_state:
    st.session_state.df_usuarios = None
if "mapa_liderancas" not in st.session_state:
    st.session_state.mapa_liderancas = {}
if "lista_autorizados" not in st.session_state:
    st.session_state.lista_autorizados = []
if "has_lideranca" not in st.session_state:
    st.session_state.has_lideranca = False

# ===== Carregar USUARIOS_SICOOB =====
if usuarios_file is not None:
    try:
        df_usuarios = read_excel_auto(usuarios_file, usuarios_file.name)
        if "USUARIO" not in df_usuarios.columns:
            st.error("A planilha de usuários precisa ter a coluna 'USUARIO'.")
        else:
            df_usuarios["USUARIO"] = df_usuarios["USUARIO"].astype(str).str.strip().str.upper()
            st.session_state.df_usuarios = df_usuarios
            st.session_state.lista_autorizados = df_usuarios["USUARIO"].dropna().unique().tolist()

            has_lideranca = "LIDERANCA" in df_usuarios.columns
            st.session_state.has_lideranca = has_lideranca
            mapa = {}
            if has_lideranca:
                df_usuarios["LIDERANCA"] = df_usuarios["LIDERANCA"].astype(str).str.strip()
                for _, row in df_usuarios.iterrows():
                    u = row["USUARIO"]
                    l = row["LIDERANCA"]
                    if pd.isna(u) or pd.isna(l):
                        continue
                    mapa.setdefault(l, []).append(u)
            st.session_state.mapa_liderancas = mapa
            st.success("Planilha de usuários carregada com sucesso.")
    except Exception as e:
        st.error(f"Erro ao carregar USUARIOS_SICOOB: {e}")

# ===== Filtros =====
with tab1:
    st.subheader("Configurações / Filtros")
    col_a, col_b = st.columns(2)

    if st.session_state.df_usuarios is None:
        st.info("Envie a planilha **USUARIOS_SICOOB** na barra lateral para habilitar os filtros.")
        filtro_val = None
    else:
        if st.session_state.has_lideranca:
            # Filtro por liderança
            opcoes = ["TODOS"] + sorted(st.session_state.mapa_liderancas.keys())
            filtro_val = col_a.selectbox("Liderança", options=opcoes, index=0)
        else:
            # Filtro por colaborador
            opcoes = ["TODOS"] + sorted(st.session_state.lista_autorizados)
            filtro_val = col_a.selectbox("Colaborador", options=opcoes, index=0)

    # ===== Processar Volumetria =====
    st.markdown("---")
    st.subheader("Resultado da Volumetria")
    placeholder_texto = st.empty()

    if processar_btn:
        if volumetria_file is None:
            st.error("Envie o arquivo de **Volumetria** na barra lateral.")
        elif st.session_state.df_usuarios is None:
            st.error("Envie a planilha **USUARIOS_SICOOB** para cruzar os dados.")
        else:
            try:
                df_principal = read_excel_auto(volumetria_file, volumetria_file.name)

                if "JUSTIFICATIVA" in df_principal.columns:
                    df_principal["JUSTIFICATIVA"] = df_principal["JUSTIFICATIVA"].apply(
                        lambda x: x.replace(";", ":") if isinstance(x, str) else x
                    )
                if "USUARIO" not in df_principal.columns:
                    st.error("A planilha principal não contém a coluna 'USUARIO'.")
                else:
                    if "IDREGISTROCONTROLADO" in df_principal.columns:
                        df_principal["IDREGISTROCONTROLADO"] = (
                            df_principal["IDREGISTROCONTROLADO"].astype(str).str.split("#").str[0]
                        )
                    df_principal["USUARIO"] = df_principal["USUARIO"].astype(str).str.strip().str.upper()

                    autorizados = st.session_state.lista_autorizados
                    mapa = st.session_state.mapa_liderancas
                    has_l = st.session_state.has_lideranca

                    df_filtrado = df_principal[df_principal["USUARIO"].isin(autorizados)].copy()

                    if filtro_val in (None, "TODOS"):
                        colaboradores = sorted(autorizados)
                    elif has_l:
                        colaboradores = mapa.get(filtro_val, [])
                        df_filtrado = df_filtrado[df_filtrado["USUARIO"].isin(colaboradores)]
                    else:
                        colaboradores = [filtro_val]
                        df_filtrado = df_filtrado[df_filtrado["USUARIO"] == filtro_val]

                    # resumo por colaborador e tipo
                    total_por_tipo = {}
                    linhas_texto = []
                    total_geral = 0

                    for colab in sorted(colaboradores):
                        df_c = df_filtrado[df_filtrado["USUARIO"] == colab]
                        linhas_texto.append(f"Colaborador: {colab}")
                        if df_c.empty:
                            linhas_texto.append("  Total de Serviços: 0")
                            linhas_texto.append("    - Nenhum fluxo encontrado.\n")
                        else:
                            cont = df_c["IDREGISTROCONTROLADO"].value_counts().reset_index()
                            cont.columns = ["Tipo de Serviço", "Quantidade"]
                            total_colab = cont["Quantidade"].sum()
                            total_geral += total_colab
                            linhas_texto.append(f"  Total de Serviços: {total_colab}")
                            for _, r in cont.iterrows():
                                linhas_texto.append(f"    - {r['Tipo de Serviço']}: {r['Quantidade']}")
                                total_por_tipo[r['Tipo de Serviço']] = total_por_tipo.get(r['Tipo de Serviço'], 0) + int(r['Quantidade'])
                            linhas_texto.append("")

                    placeholder_texto.code("\n".join(linhas_texto), language="text")

                    # Tab Resumo
                    with tab2:
                        st.subheader("Resumo por Tipo de Serviço (Time)")
                        if total_por_tipo:
                            df_resumo = pd.DataFrame(
                                [{"Tipo de Serviço": k, "Quantidade": v, "DIAS": 1, "MÉDIA": float(v)/1.0}
                                 for k, v in sorted(total_por_tipo.items())]
                            )
                            st.dataframe(df_resumo, use_container_width=True)
                            st.info(f"TOTAL GERAL DE SERVIÇOS: {total_geral}")
                        else:
                            st.info("Sem serviços para o filtro atual.")

                        # Exportar PDF
                        texto_vol = "\n".join(linhas_texto)
                        texto_res = ""
                        if total_por_tipo:
                            texto_res = "TOTAL POR SERVIÇO DO TIME:\n" + "\n".join(
                                [f"- {k}: {v}, DIAS: 1, MÉDIA: {float(v):.2f}" for k, v in sorted(total_por_tipo.items())]
                            ) + f"\n\nTOTAL GERAL DE SERVIÇOS: {total_geral}\n"
                        pdf_bytes = exportar_pdf(texto_vol, texto_res)
                        st.download_button(
                            label="⬇️ Baixar PDF do Resultado",
                            data=pdf_bytes,
                            file_name=f"relatorio_volumetria_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                            mime="application/pdf"
                        )

            except Exception as e:
                st.error(f"Erro ao processar volumetria: {e}")

# ===== Compilador =====
with tab3:
    st.subheader("Compilador de Planilhas")
    if compilar_btn:
        if comp1_file is None or comp2_file is None:
            st.error("Envie os dois arquivos para compilar.")
        else:
            try:
                df1 = read_excel_auto(comp1_file, comp1_file.name)
                df2 = read_excel_auto(comp2_file, comp2_file.name)
                df_combinado = pd.concat([df1, df2], ignore_index=True)

                if usar_usuarios_compilador:
                    if usuarios_compilador_file is None:
                        st.error("Selecione a planilha USUARIOS_SICOOB para filtrar.")
                    else:
                        df_uc = read_excel_auto(usuarios_compilador_file, usuarios_compilador_file.name)
                        if "USUARIO" not in df_uc.columns:
                            st.error("A planilha de usuários para filtro precisa ter a coluna 'USUARIO'.")
                        else:
                            lista_aut = (
                                df_uc["USUARIO"].astype(str).str.strip().str.upper().dropna().unique().tolist()
                            )
                            if "USUARIO" in df_combinado.columns:
                                df_combinado["USUARIO"] = df_combinado["USUARIO"].astype(str).str.strip().str.upper()
                                df_combinado = df_combinado[df_combinado["USUARIO"].isin(lista_aut)]
                            else:
                                st.warning("Coluna 'USUARIO' não encontrada em um dos arquivos a compilar.")

                if df_combinado.empty:
                    st.warning("Nenhum dado para salvar após os filtros.")
                else:
                    # Ajuste opcional de formatação de datas no Excel de saída
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl", datetime_format="dd/mm/yyyy hh:mm:ss") as writer:
                        df_combinado.to_excel(writer, index=False, sheet_name="Dados")
                        ws = writer.sheets["Dados"]
                        for col_name in ["DATAHORAINICIOATIVIDADE", "DATAHORAFIMATIVIDADE"]:
                            if col_name in df_combinado.columns:
                                col_idx = df_combinado.columns.get_loc(col_name) + 1
                                for row in range(2, len(df_combinado) + 2):
                                    cell = ws.cell(row=row, column=col_idx)
                                    cell.number_format = "DD/MM/YYYY HH:MM:SS"

                    st.success("Compilação concluída!")
                    st.download_button(
                        label="⬇️ Baixar Excel Compilado",
                        data=output.getvalue(),
                        file_name=f"compilado_personalizado_{datetime.now().strftime('%d-%m-%Y_%H-%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except Exception as e:
                st.error(f"Erro durante a compilação: {e}")
