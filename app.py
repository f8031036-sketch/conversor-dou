import streamlit as st
import tempfile, os
from pathlib import Path
from conversor import convert
from conversor_docx import convert_docx

st.set_page_config(
    page_title="Conversor DOU — MGI/SSC",
    page_icon="📄",
    layout="centered"
)

st.markdown("""
<style>
  html, body, [class*="css"] { font-family: 'Segoe UI', sans-serif; }
  .header-box {
    background: #1351b4; color: white;
    padding: 18px 24px; border-radius: 8px; margin-bottom: 24px;
  }
  .header-box h1 { font-size: 20px; font-weight: 700; margin: 0 0 4px 0; }
  .header-box p  { font-size: 12px; margin: 0; opacity: .85; }
  .badge {
    display: inline-block; background: rgba(255,255,255,.2);
    border-radius: 4px; padding: 2px 8px; font-size: 10px;
    font-weight: 700; letter-spacing: .8px; text-transform: uppercase; margin-bottom: 8px;
  }
  .spec-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 10px; margin: 12px 0; }
  .spec-item {
    background: #f5f7ff; border-left: 3px solid #1351b4;
    border-radius: 6px; padding: 10px 12px;
  }
  .spec-label { font-size: 10px; text-transform: uppercase; color: #888; margin-bottom: 2px; }
  .spec-value { font-size: 13px; font-weight: 700; color: #1351b4; }
  .tip-box {
    background: #fff8e1; border-left: 3px solid #f9a825;
    border-radius: 6px; padding: 12px 14px; font-size: 12px;
    color: #555; margin: 12px 0;
  }
  #MainMenu { visibility: hidden; }
  footer     { visibility: hidden; }
  header     { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="header-box">
  <div class="badge">MGI · SSC · DGP</div>
  <h1>📄 Conversor DOU</h1>
  <p>Converte documentos do SEI em DOCX formatado no padrão da Imprensa Nacional</p>
</div>
""", unsafe_allow_html=True)

st.markdown("#### Padrão aplicado automaticamente")
st.markdown("""
<div class="spec-grid">
  <div class="spec-item"><div class="spec-label">Fonte</div><div class="spec-value">Calibri 9pt</div></div>
  <div class="spec-item"><div class="spec-label">Entrelinhamento</div><div class="spec-value">Espaço simples</div></div>
  <div class="spec-item"><div class="spec-label">Margens</div><div class="spec-value">1 cm laterais</div></div>
  <div class="spec-item"><div class="spec-label">Página</div><div class="spec-value">27 × 35 cm (DOU)</div></div>
  <div class="spec-item"><div class="spec-label">Tabela</div><div class="spec-value">Autofit · Bordas simples</div></div>
  <div class="spec-item"><div class="spec-label">Referência</div><div class="spec-value">Portaria IN/CC/PR nº 1/2024</div></div>
</div>
""", unsafe_allow_html=True)

st.divider()

# ── Abas: PDF ou DOCX ─────────────────────────────────────────────────────────
aba_pdf, aba_docx = st.tabs(["📄 Upload de PDF", "📝 Upload de DOCX (Word)"])

with aba_pdf:
    st.markdown("##### Faça o upload do PDF exportado pelo SEI")
    st.markdown("""
    <div class="tip-box">
    ⚠️ O PDF pode ter imperfeições na extração de células longas. Se o resultado não ficar perfeito, use a aba <strong>DOCX (Word)</strong>.
    </div>
    """, unsafe_allow_html=True)
    arquivo_pdf = st.file_uploader("Selecione o PDF", type=["pdf"],
                                    label_visibility="collapsed", key="up_pdf")
    if arquivo_pdf:
        st.success(f"✅ **{arquivo_pdf.name}** · {arquivo_pdf.size // 1024} KB")

with aba_docx:
    st.markdown("##### Faça o upload do DOCX gerado pelo Word")
    st.markdown("""
    <div class="tip-box">
    ✅ <strong>Resultado mais preciso!</strong> Passos:<br>
    1. Abra o PDF do SEI no <strong>Word</strong> (Arquivo → Abrir → selecione o PDF)<br>
    2. O Word converte automaticamente — clique em <strong>OK</strong><br>
    3. Salve como <strong>.docx</strong> (Arquivo → Salvar como → Word)<br>
    4. Faça o upload do .docx aqui
    </div>
    """, unsafe_allow_html=True)
    arquivo_docx = st.file_uploader("Selecione o DOCX", type=["docx"],
                                     label_visibility="collapsed", key="up_docx")
    if arquivo_docx:
        st.success(f"✅ **{arquivo_docx.name}** · {arquivo_docx.size // 1024} KB")

st.divider()

# ── Formato de publicação ─────────────────────────────────────────────────────
st.markdown("#### Formato de publicação")

tem_arquivo = arquivo_pdf is not None or arquivo_docx is not None

col1, col2 = st.columns(2)
with col1:
    btn_inteira = st.button(
        "📄 Página inteira\n\n25cm · Portarias com Anexo CCE/FCE",
        use_container_width=True,
        disabled=not tem_arquivo,
        key="btn_inteira"
    )
with col2:
    btn_meia = st.button(
        "📋 Meia página\n\n12cm · Atos simples · Sem tabela grande",
        use_container_width=True,
        disabled=not tem_arquivo,
        key="btn_meia"
    )

# ── Conversão ─────────────────────────────────────────────────────────────────
def processar(formato):
    # Determina qual arquivo usar
    usar_docx = arquivo_docx is not None
    arquivo   = arquivo_docx if usar_docx else arquivo_pdf
    sufixo    = ".docx" if usar_docx else ".pdf"

    with st.spinner("Processando…"):
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=sufixo) as tmp:
                arquivo.seek(0)
                tmp.write(arquivo.read())
                src_path = tmp.name

            out_path = src_path.replace(sufixo, "_DOU.docx")
            prog     = st.progress(0, text="Lendo o arquivo…")

            if usar_docx:
                convert_docx(src_path, out_path, formato=formato)
            else:
                convert(src_path, out_path, formato=formato)

            prog.progress(100, text="Concluído!")

            with open(out_path, "rb") as f:
                docx_bytes = f.read()

            os.unlink(src_path)
            os.unlink(out_path)

            nome_saida = Path(arquivo.name).stem + f"_DOU_{formato}.docx"
            fmt_label  = "Página inteira (25cm)" if formato == "inteira" else "Meia página (12cm)"
            metodo     = "DOCX via Word" if usar_docx else "PDF direto"

            st.success("✅ DOCX gerado com sucesso!")
            st.caption(f"{fmt_label} · {metodo} · Calibri 9pt · Portaria IN/CC/PR nº 1/2024")

            st.download_button(
                label="⬇️ Baixar DOCX",
                data=docx_bytes,
                file_name=nome_saida,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"⚠️ Erro ao processar: {e}")
            try: os.unlink(src_path)
            except: pass

if btn_inteira:
    processar("inteira")
elif btn_meia:
    processar("meia")
elif not tem_arquivo:
    st.info("⬆️ Selecione um arquivo acima para habilitar a conversão.")

st.divider()
st.caption("MGI · Secretaria de Serviços Compartilhados · DGP — uso interno")
