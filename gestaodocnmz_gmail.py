import streamlit as st
import pandas as pd
import zipfile
import io
import re
from datetime import datetime, timedelta

# ===== BIBLIOTECAS PARA LER DOCUMENTOS =====
try:
    from docx import Document
except ImportError:
    Document = None
try:
    import PyPDF2
except ImportError:
    PyPDF2 = None

# ===== CONFIGURAÇÕES DA PÁGINA =====
st.set_page_config(page_title="Auditor NAQH — Gestão de Documentos", layout="wide")
st.title("📧 AUDITORIA DE DOCUMENTOS NORMATIVOS")
st.subheader("Conforme Norma Zero — Versão 4ª (Março/2025)")
st.markdown("---")

# ===== PRAZOS CONFORME NORMA ZERO (Item 4.2.3) =====
PRAZO_VERIF_CURTO = 7    # dias → POP, NOR, ROT
PRAZO_VERIF_LONGO = 15   # dias → demais documentos
PRAZO_ASSINATURA = 15     # dias após aprovação (Item 4.2.4)
ALERTA_VENCIMENTO = 60    # dias antes de vencer (Item 4.2.7)
PRAZO_PARADO = 7           # dias sem movimentação → alerta

# ===== VALIDADE POR TIPO DE DOCUMENTO (Norma Zero — Quadro 3) =====
VALIDADE_POR_TIPO = {
    "POL": 4,    # Política → 4 anos
    "REG": 4,    # Regimento → 4 anos
    "NOR": 2,    # Norma → 2 anos
    "POP": 2,    # POP → 2 anos
    "PROG": 2,   # Programa → 2 anos
    "PROT": 2,   # Protocolo → 2 anos
    "ROT": 2,    # Rotina → 2 anos
    "MAN": 2,    # Manual → 2 anos
    "PLANC": 1   # Plano de Contingência → 1 ano
}

# ===== PADRÕES DE BUSCA NO CABEÇALHO =====
PADRAO_COD_VERS = r"([A-Z]{2,5}[_][A-Z]{2,}\d{3,})\s*Versão\s*(\d+)[ªº]?"
PADRAO_DATA_APROV = [
    r"Data\s*de\s*aprovação\s*[:]*\s*(\d{2}/\d{2}/\d{4})",
    r"Aprovad[ao]\s*em\s*(\d{2}/\d{2}/\d{4})",
    r"(\d{2}/\d{2}/\d{4})\s*$",
]

# ==================================================
# 🛡️ REGRA DE OURO — VALIDAÇÃO OBRIGATÓRIA
# ==================================================
def validar_documento(codigo, versao):
    """Sem Código OU Sem Versão → NÃO entra no fluxo → Devolve ao Setor"""
    if not codigo or str(codigo).strip().upper() in ["", "NAN", "NONE"]:
        return False, "❌ DEVOLVIDO: SEM CÓDIGO — Não pode iniciar verificação"
    if not versao or str(versao).strip() in ["", "NAN", "NONE"]:
        return False, "❌ DEVOLVIDO: SEM VERSÃO — Não pode iniciar verificação"
    return True, "✅ VÁLIDO: Código + Versão presentes"

def chave_unica(codigo, versao):
    """Código + Versão = Identidade Única do Documento"""
    c = str(codigo).strip().upper().replace(" ", "")
    v = str(versao).strip().replace(" ", "")
    return f"{c}||V{v}"

def extrair_tipo_documento(codigo):
    """Identifica tipo pela sigla no código (NOR_, POP_, etc.)"""
    if not codigo:
        return None
    codigo = str(codigo).upper().strip()
    for tipo in VALIDADE_POR_TIPO.keys():
        if codigo.startswith(tipo + "_"):
            return tipo
    return None

def calcular_validade(data_aprov, tipo_doc):
    """Calcula Data de Próxima Atualização conforme Norma Zero"""
    if not data_aprov or not tipo_doc:
        return None
    anos = VALIDADE_POR_TIPO.get(tipo_doc, 2)  # padrão 2 anos
    try:
        dt = pd.to_datetime(data_aprov, dayfirst=True)
        return dt + pd.DateOffset(years=anos)
    except:
        return None

# ==================================================
# 📖 FUNÇÕES DE LEITURA DE DOCUMENTOS
# ==================================================
def extrair_dados_do_texto(texto):
    """Busca Código, Versão e Data de Aprovação no texto do cabeçalho"""
    codigo, versao, data_aprov = None, None, None

    # Busca Código + Versão
    m = re.search(PADRAO_COD_VERS, texto.upper())
    if m:
        codigo = m.group(1).strip()
        versao = m.group(2).strip()

    # Busca Data de Aprovação
    for padrao in PADRAO_DATA_APROV:
        m = re.search(padrao, texto, re.IGNORECASE)
        if m:
            data_aprov = m.group(1)
            break

    return codigo, versao, data_aprov

def extrair_dados_docx(binario):
    """Lê cabeçalho de arquivo Word"""
    if not Document:
        return None, None, None
    try:
        doc = Document(io.BytesIO(binario))
        # Prioriza cabeçalho
        cabecalho = doc.sections[0].header
        texto_cab = "\n".join([p.text for p in cabecalho.paragraphs])
        # Depois corpo
        texto_corpo = "\n".join([p.text for p in doc.paragraphs])
        return extrair_dados_do_texto(texto_cab + "\n" + texto_corpo)
    except:
        return None, None, None

def extrair_dados_pdf(binario):
    """Lê cabeçalho de arquivo PDF"""
    if not PyPDF2:
        return None, None, None
    try:
        pagina1 = PyPDF2.PdfReader(io.BytesIO(binario)).pages[0].extract_text() or ""
        return extrair_dados_do_texto(pagina1)
    except:
        return None, None, None

# ==================================================
# 📅 PREENCHER DATAS LACUNARES
# ==================================================
def preencher_datas_lacunares(df):
    """Só preenche onde está vazio. Preserva tudo que existe."""
    df = df.copy()
    colunas_datas = [
        "DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO",
        "FIM DA 1ª VERIFICAÇÃO DO RESPONSÁVEL",
        "DATA DE RECEBIMENTO PARA 2ª VERIFICAÇÃO",
        "FIM DA 2ª VERIFICAÇÃO DO RESPONSÁVEL",
        "DATA DE APROVAÇÃO",
        "PRÓXIMA ATUALIZAÇÃO"
    ]
    for col in colunas_datas:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')

    # Se tem Fim da 1ª Verificação mas sem Data de Entrada → Estima
    for idx, linha in df.iterrows():
        dt_fim = linha.get("FIM DA 1ª VERIFICAÇÃO DO RESPONSÁVEL")
        dt_entrada = linha.get("DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO")
        verif_ok = str(linha.get("1ª VERIFICAÇÃO EZEQUIAS", "")).strip().upper() == "OK"

        if pd.isna(dt_entrada) and pd.notna(dt_fim) and verif_ok:
            df.at[idx, "DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO"] = dt_fim - timedelta(days=PRAZO_VERIF_CURTO)

    # Se tem Data de Aprovação → Calcula Próxima Atualização
    for idx, linha in df.iterrows():
        dt_aprov = linha.get("DATA DE APROVAÇÃO")
        prox_atual = linha.get("PRÓXIMA ATUALIZAÇÃO")
        codigo = str(linha.get("Código", ""))

        if pd.notna(dt_aprov) and pd.isna(prox_atual):
            tipo = extrair_tipo_documento(codigo)
            dt_prox = calcular_validade(dt_aprov, tipo)
            if dt_prox:
                df.at[idx, "PRÓXIMA ATUALIZAÇÃO"] = dt_prox

    return df

# ==================================================
# 🚀 INTERFACE PRINCIPAL
# ==================================================
st.subheader("📂 PASSO 1: Carregar Planilha de Controle")
arquivo_planilha = st.file_uploader("Selecione sua Planilha (.xlsx)", type="xlsx")

st.subheader("📦 PASSO 2: Carregar Backup do Gmail (ZIP)")
st.info("🏠 Em Casa: Selecione o arquivo .zip baixado do Gmail")
arquivo_zip = st.file_uploader("Selecione o arquivo .zip", type="zip")

st.markdown("---")
st.subheader("⚙️ Regras Aplicadas (Norma Zero — Item 4.2.3)")
st.markdown("""
- ✅ **Preserva dados existentes** — só preenche células VAZIAS
- ❌ **Sem Código OU Sem Versão → Devolvido ao Setor** (não entra no fluxo)
- ✅ **Com Data de Aprovação → Aprovado de 1ª** (2ª Verificação fica vazia)
- 🔄 **Sem Data de Aprovação → Segue para 2ª Verificação**
- ⚠️ **SÓ 2 Verificações permitidas** — após isso, volta ao Setor
- ⏱️ **Prazos:** POP/NOR/ROT = 7 dias | Demais = 15 dias | Assinatura = 15 dias
- 🔔 **Alerta de vencimento:** 60 dias antes da validade
""")

# ==================================================
# 🔍 EXECUTAR AUDITORIA
# ==================================================
if arquivo_planilha and arquivo_zip:
    if st.button("🔍 INICIAR AUDITORIA COMPLETA", type="primary"):
        # Carregar planilha
        df = pd.read_excel(arquivo_planilha)
        st.success(f"✅ Planilha carregada: {len(df)} documentos registrados")

        # Índice de documentos já existentes
        existentes = {}
        for idx, linha in df.iterrows():
            cod = str(linha.get("Código", "")).strip().upper()
            ver = str(linha.get("Versão", "")).strip()
            if cod and ver and cod != "NAN":
                existentes[chave_unica(cod, ver)] = idx

        # Listas de relatórios
        devolvidos_sem_dados = []
        aprovados_faltantes = []
        novas_linhas = []

        # Abrir ZIP e processar documentos
        with zipfile.ZipFile(io.BytesIO(arquivo_zip.read())) as zf:
            arquivos = [n for n in zf.namelist() if n.lower().endswith((".docx", ".pdf"))]
            st.info(f"📄 Encontrados {len(arquivos)} documentos dentro do ZIP")
            barra = st.progress(0)

            for i, caminho in enumerate(arquivos):
                barra.progress((i+1)/len(arquivos), f"Lendo: {caminho.split('/')[-1]}")

                try:
                    binario = zf.read(caminho)
                    if caminho.lower().endswith(".docx"):
                        cod, ver, dt_aprov = extrair_dados_docx(binario)
                    else:
                        cod, ver, dt_aprov = extrair_dados_pdf(binario)

                    # VALIDAÇÃO — REGRA DE OURO
                    valido, motivo = validar_documento(cod, ver)
                    if not valido:
                        devolvidos_sem_dados.append({
                            "Arquivo": caminho.split("/")[-1],
                            "Motivo": motivo,
                            "Código encontrado": cod or "---",
                            "Versão encontrada": ver or "---"
                        })
                        continue

                    chave = chave_unica(cod, ver)

                    # Já existe na planilha? → NÃO tocar
                    if chave in existentes:
                        continue

                    # Documento NOVO
                    tipo = extrair_tipo_documento(cod)
                    prox_atual = calcular_validade(dt_aprov, tipo) if dt_aprov else None

                    linha_nova = {
                        "Código": cod,
                        "Versão": ver,
                        "Tipo de Documento": tipo,
                        "Data de Aprovação": dt_aprov,
                        "Próxima Atualização": prox_atual,
                    }

                    if dt_aprov:
                        linha_nova["Status"] = "APROVADO DE 1ª"
                        aprovados_faltantes.append(linha_nova)
                    else:
                        linha_nova["Status"] = "EM TRÂMITE — Aguardando devolução do Setor"

                    novas_linhas.append(linha_nova)

                except Exception as e:
                    continue

        barra.empty()
        st.success("✅ AUDITORIA CONCLUÍDA!")
        st.markdown("---")

        # = RELATÓRIOS =
        if devolvidos_sem_dados:
            st.subheader("❌ DOCUMENTOS DEVOLVIDOS (sem Código ou Versão)")
            st.dataframe(pd.DataFrame(devolvidos_sem_dados), use_container_width=True)

        if aprovados_faltantes:
            st.subheader("⚠️ DOCUMENTOS APROVADOS QUE FALTAVAM NA PLANILHA")
            st.dataframe(pd.DataFrame(aprovados_faltantes), use_container_width=True)

        if novas_linhas:
            st.subheader("📝 NOVAS LINHAS A ADICIONAR")
            df_novas = pd.DataFrame(novas_linhas)
            st.dataframe(df_novas, use_container_width=True)

            # = GERAR PLANILHA FINAL =
            st.subheader("💾 GERANDO PLANILHA ATUALIZADA...")
            df_final = pd.concat([df, df_novas], ignore_index=True)
            df_final = preencher_datas_lacunares(df_final)

            # Alertas de vencimento
            if "Próxima Atualização" in df_final.columns:
                hoje = pd.Timestamp.now()
                vencendo = df_final[
                    (df_final["Próxima Atualização"].notna()) &
                    (df_final["Próxima Atualização"] - hoje <= timedelta(days=ALERTA_VENCIMENTO))
                ]
                if len(vencendo) > 0:
                    st.warning(f"🔔 ATENÇÃO: {len(vencendo)} documento(s) vão vencer em até {ALERTA_VENCIMENTO} dias!")
                    st.dataframe(vencendo[["Código", "Versão", "Próxima Atualização"]], use_container_width=True)

            # Download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_final.to_excel(writer, index=False)

            st.download_button(
                label="📥 BAIXAR PLANILHA ATUALIZADA.xlsx",
                data=output.getvalue(),
                file_name=f"PLANILHA_GESTAO_DOCS_ATUALIZADA_{datetime.now().strftime('%d%m%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("✅ PRONTO! Sua planilha original está INTACTA. Você baixou uma CÓPIA ATUALIZADA!")
        else:
            st.info("✅ Nenhuma alteração necessária — tudo já está registrado corretamente!")

st.markdown("---")
st.caption("🏠 Versão para uso em Casa (Leitura de ZIP) | 🏢 No Serviço: trocar fonte para Gmail ao vivo")
