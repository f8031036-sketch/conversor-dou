#!/usr/bin/env python3
"""
sei2dou_final.py — Converte PDF do SEI para DOCX no padrão DOU
Imprensa Nacional: Calibri 9pt, espaço simples, tabela única de 25cm
"""

import sys, re
from pathlib import Path
from collections import defaultdict

import pdfplumber
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT_NAME = "Calibri"
FONT_SIZE = Pt(9)

# Fronteiras reais das colunas (medidas do PDF)
COL_BOUNDS    = [0, 247, 354, 405, 461, 9999]
COL_WIDTHS_CM = [6.5, 5.5, 2.5, 3.5, 7.0]   # soma = 25cm

FOOTER_RE = re.compile(
    r'(^Minuta de Portaria\s|^Este conteúdo não substitui|^Referência: Processo)',
    re.IGNORECASE
)
HEADER_ROW = ['UNIDADE/ÓRGÃO', 'CARGO/FUNÇÃO', 'NÍVEL', 'Nº DOC SEI', 'Nº PROCESSO SEI']

# ── Formatação DOCX ───────────────────────────────────────────────────────────

def set_run(run, bold=False):
    run.font.name = FONT_NAME
    run.font.size = FONT_SIZE
    run.font.bold = bool(bold)
    rpr = run._r.get_or_add_rPr()
    old = rpr.find(qn('w:rFonts'))
    if old is not None: rpr.remove(old)
    rf = OxmlElement('w:rFonts')
    for a in ('w:ascii','w:hAnsi','w:cs'): rf.set(qn(a), FONT_NAME)
    rpr.insert(0, rf)


def set_para(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, fi=None, li=None):
    pf = p.paragraph_format
    pf.alignment    = align
    pf.space_before = Pt(0)
    pf.space_after  = Pt(0)
    pf.line_spacing = Pt(10)
    if fi is not None: pf.first_line_indent = Cm(fi)
    if li is not None: pf.left_indent       = Cm(li)


def add_p(doc, text, bold=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, fi=None, li=None):
    p = doc.add_paragraph()
    set_para(p, align=align, fi=fi, li=li)
    if text:
        r = p.add_run(text)
        set_run(r, bold=bool(bold))
    return p


def set_cell_w(cell, cm):
    tc  = cell._tc
    tcp = tc.get_or_add_tcPr()
    w   = OxmlElement('w:tcW')
    w.set(qn('w:w'), str(int(Cm(cm).pt * 20)))
    w.set(qn('w:type'), 'dxa')
    old = tcp.find(qn('w:tcW'))
    if old is not None: tcp.remove(old)
    tcp.append(w)


def add_table(doc, rows_data):
    tbl = doc.add_table(rows=0, cols=5)
    tbl.style = 'Table Grid'
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    # Largura total
    tp = tbl._tbl.find(qn('w:tblPr'))
    if tp is None: tp = OxmlElement('w:tblPr'); tbl._tbl.insert(0, tp)
    tw = OxmlElement('w:tblW')
    tw.set(qn('w:w'), str(int(Cm(25).pt * 20)))
    tw.set(qn('w:type'), 'dxa')
    old = tp.find(qn('w:tblW'))
    if old is not None: tp.remove(old)
    tp.append(tw)

    for row_cells in rows_data:
        row = tbl.add_row()
        is_hdr = (list(row_cells) == HEADER_ROW)
        for ci, ct in enumerate(row_cells[:5]):
            cell = row.cells[ci]
            set_cell_w(cell, COL_WIDTHS_CM[ci])
            p = cell.paragraphs[0]
            set_para(p, align=WD_ALIGN_PARAGRAPH.LEFT)
            r = p.add_run(str(ct or ''))
            set_run(r, bold=bool(is_hdr))
    return tbl


# ── Extração de PDF ───────────────────────────────────────────────────────────

def words_to_lines(words):
    bkt = defaultdict(list)
    for w in words:
        bkt[round(w['top'] / 4) * 4].append(w)
    lines = []
    for key in sorted(bkt):
        ws = sorted(bkt[key], key=lambda x: x['x0'])
        lines.append({
            'top': key, 'words': ws,
            'text': ' '.join(w['text'] for w in ws),
            'x0_min': min(w['x0'] for w in ws),
        })
    return lines


def col_of(x):
    for i in range(len(COL_BOUNDS)-1):
        if COL_BOUNDS[i] <= x < COL_BOUNDS[i+1]: return i
    return 4


def extract_cells(line):
    cols = ['','','','','']
    for w in line['words']:
        ci = col_of(w['x0'])
        cols[ci] = (cols[ci]+' '+w['text']).strip()
    return cols


def has_data_cols(line):
    """True se há conteúdo característico de tabela em colunas 2-4."""
    t = line['text']
    if re.search(r'\b[CF]CE\s+\d+\.\d+', t): return True
    if re.search(r'\d{5,}/\d{4}-\d{2}', t): return True
    for w in line['words']:
        if w['x0'] >= 405 and re.match(r'^\d{7,9}$', w['text']): return True
    return False


def extract_pages(pdf_path):
    """
    Extrai linhas e classifica em dois passes:
    Passe 1: classifica pré-tabela (org, title, ementa, text, article, etc.)
    Quando encontra o cabeçalho CARGO/FUNÇÃO, entra em modo tabela.
    Em modo tabela: TUDO é linha de tabela.
    """
    all_lines = []
    in_table = False

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(
                x_tolerance=3, y_tolerance=3,
                keep_blank_chars=False, use_text_flow=False)
            for line in words_to_lines(words):
                t = line['text'].strip()
                if not t or FOOTER_RE.match(t):
                    line['kind'] = 'skip'; all_lines.append(line); continue

                # Detecta início da tabela
                if not in_table and 'CARGO/FUNÇÃO' in t and 'NÍVEL' in t:
                    in_table = True

                if in_table:
                    line['kind'] = 'table_row'
                else:
                    line['kind'] = classify_pre_table(line)

                all_lines.append(line)
    return all_lines


def classify_pre_table(line):
    t, x = line['text'].strip(), line['x0_min']
    if re.match(r'^MINISTÉRIO|^Secretaria de Serviços', t) and x > 100: return 'org'
    if re.match(r'^PORTARIA\s+\S+\s+Nº', t, re.I): return 'title'
    if re.match(r'^ANEXO\b', t, re.I) and x > 200: return 'section_label'
    if re.match(r'^(CAPÍTULO|SEÇÃO|SUBSEÇÃO)\s', t, re.I): return 'chapter'
    if re.match(r'^Art\.\s+\d+', t): return 'article'
    if re.match(r'^(Parágrafo único|§\s*\d+)', t): return 'paragrafo'
    if re.match(r'^[IVXivx]+\s*[-–]\s+', t): return 'inciso'
    if re.match(r'^[a-z]\)\s+', t): return 'alinea'
    if re.match(r'^[A-ZÁÉÍÓÚÃÂÊÎÔÛÇÀ\s]+$', t) and 4 < len(t) < 60 and x > 200:
        return 'signature'
    if x > 260: return 'ementa'
    return 'text'


# ── Agrupamento ───────────────────────────────────────────────────────────────

def merge_pre_table(lines):
    """Une linhas de texto/ementa/artigo em blocos lógicos."""
    result, i = [], 0
    while i < len(lines):
        line = lines[i]; k = line['kind']
        if k == 'skip': i += 1; continue

        if k == 'ementa':
            m = line['text']
            j = i+1
            while j < len(lines) and lines[j]['kind'] == 'ementa':
                m += ' ' + lines[j]['text'].strip(); j += 1
            result.append({'kind':'ementa','text':m.strip()}); i=j; continue

        if k == 'text':
            m = line['text']
            j = i+1
            while j < len(lines):
                nk,nt = lines[j]['kind'], lines[j]['text'].strip()
                if nk=='skip': j+=1; continue
                if nk=='text':
                    ends = m.rstrip()[-1:] in '.;'
                    if not ends or (nt and nt[0].islower()):
                        m += ' '+nt; j+=1; continue
                break
            result.append({'kind':'text','text':m.strip()}); i=j; continue

        if k in ('article','paragrafo'):
            m = line['text']
            j = i+1
            while j < len(lines):
                nk,nt = lines[j]['kind'], lines[j]['text'].strip()
                if nk=='skip': j+=1; continue
                if nk=='text': m += ' '+nt; j+=1; continue
                break
            result.append({'kind':k,'text':m.strip()}); i=j; continue

        result.append({'kind':k,'text':line['text'],'line':line}); i+=1
    return result


def build_table_from_rows(table_lines):
    """
    Constrói lista de [c0,c1,c2,c3,c4] a partir das linhas brutas da tabela.
    Regra:
    - Linha com dados (c2/c3/c4 preenchidos): linha de dado.
      Se c1 vazio, busca pending_c1 acumulado antes.
    - Linha sem dados, só c1: pode ser (a) cargo antes dos dados ou (b) continuação
      do cargo da linha anterior. Guardamos em pending_c1.
    - Linha sem dados, só c0: nome de unidade / cabeçalho de bloco → pending_c0.
    - Cabeçalho CARGO/FUNÇÃO: linha fixa.
    """
    rows = []
    pending_c0 = ''
    pending_c1 = ''  # cargo acumulado antes de ver os dados

    for line in table_lines:
        t = line['text'].strip()
        cells = extract_cells(line)
        c0, c1, c2, c3, c4 = cells
        has_data = bool(c2 or c3 or c4)

        # Cabeçalho de colunas
        if 'CARGO/FUNÇÃO' in t and 'NÍVEL' in t:
            rows.append(list(HEADER_ROW))
            pending_c0 = pending_c1 = ''; continue

        if not has_data:
            # Linha-cargo sem dados: pode ser "Secretário-Executivo" (antes)
            # ou "Adjunto" (depois, continuação)
            if c1 and not c0:
                # Se a última linha já tem cargo preenchido → complemento (ex: "Adjunto")
                if rows and rows[-1][1] and rows[-1] != list(HEADER_ROW):
                    rows[-1][1] = (rows[-1][1]+' '+c1).strip()
                else:
                    # Cargo que precede os dados
                    pending_c1 = (pending_c1+' '+c1).strip()
            elif c0 and not c1:
                # Nome de unidade
                pending_c0 = (pending_c0+' '+c0).strip()
            elif c0 and c1:
                pending_c0 = (pending_c0+' '+c0).strip()
                pending_c1 = (pending_c1+' '+c1).strip()
            continue

        # Linha com dados
        final_c0 = (pending_c0+' '+c0).strip() if c0 else pending_c0
        final_c1 = (pending_c1+' '+c1).strip() if c1 else pending_c1
        pending_c0 = pending_c1 = ''
        rows.append([final_c0, final_c1, c2, c3, c4])

    return rows


def process(pdf_path):
    print(f'Lendo {Path(pdf_path).name}...')
    all_lines = extract_pages(pdf_path)

    pre_lines  = [l for l in all_lines if l['kind'] != 'table_row']
    table_lines = [l for l in all_lines if l['kind'] == 'table_row']

    print(f'  {len(pre_lines)} linhas de texto, {len(table_lines)} linhas de tabela')

    pre_items  = merge_pre_table(pre_lines)
    table_rows = build_table_from_rows(table_lines)
    print(f'  {len(pre_items)} blocos de texto, {len(table_rows)} linhas de tabela')

    return pre_items, table_rows


# ── Construtor DOCX ───────────────────────────────────────────────────────────

def build_docx(pre_items, table_rows, out):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width    = Cm(21);  sec.page_height   = Cm(29.7)
    sec.left_margin   = Cm(2);   sec.right_margin  = Cm(2)
    sec.top_margin    = Cm(2);   sec.bottom_margin = Cm(2)
    for p in sec.header.paragraphs: p.clear()
    for p in sec.footer.paragraphs: p.clear()

    sty = doc.styles['Normal']
    sty.font.name = FONT_NAME; sty.font.size = FONT_SIZE
    pf = sty.paragraph_format
    pf.space_before = Pt(0); pf.space_after = Pt(0); pf.line_spacing = Pt(10)

    C, J = WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.JUSTIFY
    prev = None

    for item in pre_items:
        k, t = item['kind'], item.get('text','').strip()
        if   k == 'org':          add_p(doc, t, bold=True, align=C)
        elif k == 'title':
            if prev not in (None,'org'): add_p(doc,'')
            add_p(doc, t, bold=True, align=C)
        elif k == 'ementa':       add_p(doc, t, align=J, li=8.5)
        elif k == 'text':         add_p(doc, t, align=J)
        elif k == 'signature':    add_p(doc,''); add_p(doc, t, bold=True, align=C)
        elif k == 'section_label': add_p(doc,''); add_p(doc, t, bold=True, align=C)
        elif k == 'chapter':      add_p(doc, t, align=C)
        elif k == 'article':      add_p(doc, t, align=J, fi=1.25)
        elif k == 'paragrafo':    add_p(doc, t, align=J, fi=1.25)
        elif k == 'inciso':       add_p(doc, t, align=J, li=1.25)
        elif k == 'alinea':       add_p(doc, t, align=J, li=2.5)
        prev = k

    # Tabela única com todos os dados
    if table_rows:
        add_p(doc,'')
        add_table(doc, table_rows)
        add_p(doc,'')

    doc.save(out)
    print(f'Salvo: {out}')


def convert(pdf_path, docx_path=None):
    pdf_path  = Path(pdf_path)
    docx_path = Path(docx_path or (pdf_path.stem+'_DOU.docx'))
    pre, tbl  = process(pdf_path)
    build_docx(pre, tbl, str(docx_path))
    return str(docx_path)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Uso: python3 sei2dou_final.py arquivo.pdf [saida.docx]'); sys.exit(1)
    print('Concluído:', convert(sys.argv[1], sys.argv[2] if len(sys.argv)>2 else None))
