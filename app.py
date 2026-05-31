import streamlit as st
import tempfile
import os
from pathlib import Path
from conversor import convert

# ── Configuração da página ────────────────────────────────────────────────────
st.set_page_config(
    page_title="Conversor DOU — MGI/SSC",
    page_icon="📄",
    layout="centered"
)

# ── Estilo visual ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
  /* Fonte Segoe UI em todo o app */
  html, body, [class*="css"] { font-family: 'Segoe UI', sans-serif; }

  /* Cabeçalho institucional */
  .header-box {
    background: #1351b4;
    color: white;
    padding: 18px 24px;
    border-radius: 8px;
    margin-bottom: 24px;
  }
  .header-box h1 { font-size: 20px; font-weight: 700; margin: 0 0 4px 0; }
  .header-box p  { font-size: 12px; margin: 0; opacity: .85; }

  /* Badge */
  .badge {
    display: inline-block;
    background: rgba(255,255,255,.2);
    border-radius: 4px;
    padding: 2px 8px;
    font-size: 10px;
    font-weight: 700;
    letter-spacing: .8px;
    text-transform: uppercase;
    margin-bottom: 8px;
  }

  /* Specs */
  .spec-grid {
    display: grid;
    grid-template-columns: 1fr 1fr 1fr;
    gap: 10px;
    margin: 12px 0;
  }
  .spec-item {
    background: #f5f7ff;
    border-left: 3px solid #1351b4;
    border-radius: 6px;
    padding: 10px 12px;
  }
  .spec-label { font-size: 10px; text-transform: uppercase; color: #888; margin-bottom: 2px; }
  .spec-value { font-size: 13px; font-weight: 700; color: #1351b4; }

  /* Oculta menu e rodapé do Streamlit */
  #MainMenu { visibility: hidden; }
  footer     { visibility: hidden; }
  header     { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Cabeçalho ─────────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-box">
  <div class="badge">MGI · SSC · DGP</div>
  <h1>📄 Conversor DOU</h1>
  <p>Converte PDFs exportados pelo SEI em DOCX formatado no padrão da Imprensa Nacional</p>
</div>
""", unsafe_allow_html=True)

# ── Especificações DOU ────────────────────────────────────────────────────────
st.markdown("#### Padrão aplicado automaticamente")
st.markdown("""
<div class="spec-grid">
  <div class="spec-item">
    <div class="spec-label">Fonte</div>
    <div class="spec-value">Calibri 9pt</div>
  </div>
  <div class="spec-item">
    <div class="spec-label">Entrelinhamento</div>
    <div class="spec-value">Espaço simples</div>
  </div>
  <div class="spec-item">
    <div class="spec-label">Tabela</div>
    <div class="spec-value">25 cm · Bordas simples</div>
  </div>
  <div class="spec-item">
    <div class="spec-label">Margens</div>
    <div class="spec-value">2 cm (todos os lados)</div>
  </div>
  <div class="spec-item">
    <div class="spec-label">Marcadores</div>
    <div class="spec-value">Hífen (–)</div>
  </div>
  <div class="spec-item">
    <div class="spec-label">Referência</div>
    <div class="spec-value">Portaria IN/CC/PR nº 1/2024</div>
  </div>
</div>
""", unsafe_allow_html=True)

st.divider()

# ── Upload ────────────────────────────────────────────────────────────────────
st.markdown("#### 1. Selecione o PDF do SEI")
arquivo = st.file_uploader(
    "Arraste o arquivo ou clique para selecionar",
    type=["pdf"],
    label_visibility="collapsed"
)

if arquivo:
    col1, col2 = st.columns([3, 1])
    with col1:
        st.success(f"✅ **{arquivo.name}** · {arquivo.size // 1024} KB")
    with col2:
        st.caption(f"{arquivo.size // 1024} KB")

st.divider()

# ── Formato ────────────────────────────────────────────────────────────────────
st.markdown("#### 2. Formato de publicação")

col1, col2 = st.columns(2)
with col1:
    btn_inteira = st.button(
        "📄 Página inteira\n\n25cm · Tabelas grandes · Portarias com Anexo CCE/FCE",
        use_container_width=True,
        disabled=not arquivo,
        key="btn_inteira"
    )
with col2:
    btn_meia = st.button(
        "📋 Meia página\n\n12cm · Atos simples · Sem tabela ou tabela pequena",
        use_container_width=True,
        disabled=not arquivo,
        key="btn_meia"
    )

st.divider()

# ── Converter ─────────────────────────────────────────────────────────────────
def processar(formato):
    with st.spinner("Processando…"):
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                arquivo.seek(0)
                tmp_pdf.write(arquivo.read())
                pdf_path = tmp_pdf.name

            docx_path = pdf_path.replace(".pdf", "_DOU.docx")
            progress = st.progress(0, text="Lendo o PDF…")
            convert(pdf_path, docx_path, formato=formato)
            progress.progress(100, text="Concluído!")

            with open(docx_path, "rb") as f:
                docx_bytes = f.read()

            os.unlink(pdf_path)
            os.unlink(docx_path)

            nome_saida = Path(arquivo.name).stem + f"_DOU_{formato}.docx"
            fmt_label  = "Página inteira (25cm)" if formato == "inteira" else "Meia página (12cm)"

            st.success("✅ DOCX gerado com sucesso!")
            st.caption(f"{fmt_label} · Calibri 9pt · Espaço simples · Portaria IN/CC/PR nº 1/2024")

            st.download_button(
                label="⬇️ Baixar DOCX",
                data=docx_bytes,
                file_name=nome_saida,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"⚠️ Erro ao processar o PDF: {e}")
            try: os.unlink(pdf_path)
            except: pass

if btn_inteira:
    processar("inteira")
elif btn_meia:
    processar("meia")
elif not arquivo:
    st.info("⬆️ Selecione um PDF acima para habilitar a conversão.")

st.divider()
st.caption("MGI · Secretaria de Serviços Compartilhados · DGP — uso interno")
