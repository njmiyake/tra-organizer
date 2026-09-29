import re
import io
from pathlib import Path
from collections import OrderedDict

import streamlit as st

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="TRA Report Organizer",
    page_icon="📋",
    layout="centered",
)

st.markdown("""
<style>
    .main { max-width: 720px; margin: auto; }
    .stButton > button {
        background-color: #1F4E79;
        color: white;
        font-weight: bold;
        border-radius: 6px;
        padding: 0.5rem 2rem;
        border: none;
    }
    .stButton > button:hover { background-color: #2E6DA4; }
    .stDownloadButton > button {
        background-color: #2E7D32;
        color: white;
        font-weight: bold;
        border-radius: 6px;
        padding: 0.5rem 2rem;
        border: none;
        width: 100%;
    }
    .stDownloadButton > button:hover { background-color: #388E3C; }
</style>
""", unsafe_allow_html=True)


# ── Organizer logic ───────────────────────────────────────────────────────────

def parse_comment(text):
    if not text or not text.strip():
        return [{"notebook": "", "procedure": "", "result": "", "recommendation": ""}]
    segments = text.split(" | ")
    studies = []
    pending_rec = ""
    for seg in segments:
        seg = seg.strip()
        if not seg:
            continue
        if (seg.startswith("Recommendation:") and
                "Notebook page:" not in seg and "Procedure:" not in seg):
            pending_rec = seg[len("Recommendation:"):].strip()
            continue
        sub_parts = re.split(r"(?=Notebook page:)", seg)
        for sp in sub_parts:
            sp = sp.strip()
            if not sp:
                continue
            nb_m = re.search(r"Notebook page:\s*(.+?)(?:\s{2,}|(?=Procedure:)|$)", sp)
            pr_m = re.search(r"Procedure:\s*(.+?)(?:\s{2,}(?:Result)|$)", sp, re.DOTALL)
            re_m = re.search(r"Result(?:/conclusion)?:\s*(.+?)(?:\s{2,}(?:Recommendation:|Notebook page:)|$)", sp, re.DOTALL)
            rc_m = re.search(r"Recommendation:\s*(.+?)$", sp, re.DOTALL)
            if nb_m or pr_m or re_m or rc_m:
                studies.append({
                    "notebook":       nb_m.group(1).strip() if nb_m else "",
                    "procedure":      pr_m.group(1).strip() if pr_m else "",
                    "result":         re_m.group(1).strip() if re_m else "",
                    "recommendation": rc_m.group(1).strip() if rc_m else "",
                })
            else:
                if re.match(r"^(If |Clean |Consult )", sp):
                    studies.append({"notebook": "", "procedure": "", "result": "", "recommendation": sp})
                else:
                    studies.append({"notebook": "", "procedure": "", "result": sp, "recommendation": ""})
    if not studies:
        studies = [{"notebook": "", "procedure": "", "result": "", "recommendation": text.strip()}]
    if pending_rec:
        studies[-1]["recommendation"] = (
            studies[-1]["recommendation"] + "\n\n" + pending_rec
            if studies[-1]["recommendation"] else pending_rec
        )
    return studies


def read_tra(file_obj):
    doc = Document(file_obj)
    title, section = "", ""
    for para in doc.paragraphs:
        t = para.text.strip()
        if t:
            if not title:
                title = t
            else:
                section = t
                break
    if not doc.tables:
        raise ValueError("No table found in the document.")
    rows = []
    for i, row in enumerate(doc.tables[0].rows):
        if i == 0:
            continue
        cells = [c.text.strip() for c in row.cells]
        step    = cells[0] if len(cells) > 0 else ""
        risk    = cells[1] if len(cells) > 1 else ""
        comment = cells[2] if len(cells) > 2 else ""
        if step or risk:
            rows.append((step, risk, parse_comment(comment)))
    return title, section, rows


def build_groups(rows):
    ordinals = ["1st", "2nd", "3rd", "4th", "5th"]
    groups, key_history = OrderedDict(), {}
    prev_step = current_key = None
    for step, risk, studies in rows:
        if step == prev_step:
            groups[current_key].append((risk, studies))
        else:
            history = key_history.setdefault(step, [])
            occ = len(history)
            if occ == 0:
                key = step
            else:
                if occ == 1:
                    old_key = history[0]
                    new_key = f"{step} – {ordinals[0]}"
                    groups = OrderedDict((new_key if k == old_key else k, v) for k, v in groups.items())
                    history[0] = new_key
                label = ordinals[occ] if occ < len(ordinals) else f"{occ+1}th"
                key = f"{step} – {label}"
            history.append(key)
            groups[key] = [(risk, studies)]
            current_key = key
            prev_step = step
    return groups


HEADER_BG = "1F4E79"; HEADER_FG = "FFFFFF"
ALT_BG    = "EBF3FB"; BORDER_COL = "9DC3E6"; BLUE = "1F4E79"

def _shading(cell, fill):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(shd)

def _border(cell, color=BORDER_COL):
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "single"); b.set(qn("w:sz"), "4")
        b.set(qn("w:space"), "0");   b.set(qn("w:color"), color)
        tcBorders.append(b)
    cell._tc.get_or_add_tcPr().append(tcBorders)

def _text(cell, text, bold=False, size_pt=9, color=None, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ""
    p = cell.paragraphs[0]; p.alignment = align
    p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text); r.bold = bold; r.font.size = Pt(size_pt)
    if color: r.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

def _paras(cell, items, size_pt=9):
    cell.text = ""
    first = True
    for item in items:
        p = cell.paragraphs[0] if first else cell.add_paragraph(); first = False
        p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(1)
        p.add_run(item).font.size = Pt(size_pt)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

def build_doc(title, section, groups):
    doc = Document()
    s = doc.sections[0]
    s.orientation = 1; s.page_width = Cm(29.7); s.page_height = Cm(21.0)
    s.left_margin = s.right_margin = s.top_margin = s.bottom_margin = Cm(1.5)

    tp = doc.add_paragraph(); tp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = tp.add_run(title); r.bold = True; r.font.size = Pt(16)
    r.font.color.rgb = RGBColor.from_string(BLUE)

    if section:
        doc.add_paragraph()
        sh = doc.add_heading(section, level=1)
        sh.runs[0].font.size = Pt(14)
        sh.runs[0].font.color.rgb = RGBColor.from_string(BLUE)

    COL_W   = [Cm(5.2), Cm(3.0), Cm(5.8), Cm(6.2), Cm(6.5)]
    HEADERS = ["Risk", "Notebook Page", "Procedure", "Result / Conclusion", "Recommendation"]

    for step_num, (step_name, risks) in enumerate(groups.items(), 1):
        doc.add_paragraph()
        hd = doc.add_heading(f"Step {step_num}: {step_name}", level=2)
        hd.runs[0].font.size = Pt(11)
        hd.runs[0].font.color.rgb = RGBColor.from_string(BLUE)

        tbl = doc.add_table(rows=1, cols=5)
        tbl.style = "Table Grid"; tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
        for i, (cell, hdr) in enumerate(zip(tbl.rows[0].cells, HEADERS)):
            cell.width = COL_W[i]; _shading(cell, HEADER_BG); _border(cell, "FFFFFF")
            _text(cell, hdr, bold=True, size_pt=9, color=HEADER_FG, align=WD_ALIGN_PARAGRAPH.CENTER)

        for idx, (risk, studies) in enumerate(risks):
            bg = ALT_BG if idx % 2 == 1 else "FFFFFF"

            def collect(key, _s=studies):
                vals = [s[key] for s in _s if s[key]]
                out = []
                for i, v in enumerate(vals):
                    if i > 0: out.append("")
                    out.append(v)
                return out or [""]

            dr = tbl.add_row()
            for i, cell in enumerate(dr.cells):
                cell.width = COL_W[i]; _shading(cell, bg); _border(cell)
            _text(dr.cells[0], risk, size_pt=9)
            _paras(dr.cells[1], collect("notebook"))
            _paras(dr.cells[2], collect("procedure"))
            _paras(dr.cells[3], collect("result"))
            _paras(dr.cells[4], collect("recommendation"))

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ── Streamlit UI ──────────────────────────────────────────────────────────────

st.title("📋 TRA Report Organizer")
st.markdown(
    "Upload a TRA report in the standard format "
    "(**Process Step | Risk | Comments | ...**) "
    "and download a reorganised version with section headings and split columns."
)
st.divider()

uploaded = st.file_uploader(
    "Upload your TRA report (.docx)",
    type=["docx"],
    help="The file is processed in your browser session and never stored.",
)

if uploaded:
    st.success(f"✔ Loaded: **{uploaded.name}**")

    with st.spinner("Organising document…"):
        try:
            title, section, rows = read_tra(uploaded)
            groups = build_groups(rows)
            out_buf = build_doc(title, section, groups)

            col1, col2, col3 = st.columns(3)
            col1.metric("Rows", len(rows))
            col2.metric("Process steps", len(groups))
            col3.metric("Sections created", len(groups))

            out_name = Path(uploaded.name).stem + "_organised.docx"
            st.divider()
            st.download_button(
                label="⬇ Download organised report",
                data=out_buf,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

        except Exception as e:
            st.error(f"Something went wrong: {e}")
            st.info("Make sure the document contains a table with columns: "
                    "Process Step | Risk | Comments")

st.divider()
st.caption(
    "Files are processed in memory and never saved or shared. "
    "Each session is independent."
)
