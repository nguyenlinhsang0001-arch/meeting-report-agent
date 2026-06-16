"""
report_builder.py — Sinh 2 ban bao cao tu transcript va xuat ra file Word.
  - build_full_docx    : bien ban DAY DU theo format ARTIUS
  - build_summary_docx : ban TOM TAT 1 trang cho CEO
  - generate_reports   : goi Claude, tra ve JSON {full, summary}

Tach rieng khoi app.py (khong import streamlit) de co the test doc lap.
"""

import io
import json

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

BLUE = RGBColor(0x2E, 0x75, 0xB6)
RED = RGBColor(0xC0, 0x00, 0x00)


# ============================================================
# GOI CLAUDE -> JSON 2 BAN
# ============================================================
def generate_reports(transcript: str, title: str, objective: str,
                     anthropic_key: str, model: str = "claude-sonnet-4-6") -> dict:
    from anthropic import Anthropic  # import lazy de module van test duoc khi chua cai anthropic

    schema = """{
  "full": {
    "topics": [
      {"heading": "Ten chu de ngan", "points": ["y chi tiet 1", "y chi tiet 2"]}
    ],
    "action_plan": [
      {"task": "viec can lam", "owner": "ten - vai tro / don vi", "deadline": "ngay cu the hoac 'Chua xac dinh'"}
    ]
  },
  "summary": {
    "summary": "2-3 cau tom tat dieu hanh",
    "risks_approvals": ["rui ro / viec can CEO duyet"],
    "progress_budget": ["cap nhat tien do - ngan sach"],
    "decisions": ["quyet dinh da thong qua"],
    "action_items": [{"owner": "ten - vai tro", "task": "viec", "deadline": "han chot"}]
  }
}"""

    prompt = f"""Ban la thu ky chuyen nghiep cua cong ty thiet ke va thi cong noi that ARTIUS.
Tu transcript ben duoi (da ghi ro ten + vai tro nguoi noi), hay tao DONG THOI 2 tai lieu
va tra ve DUY NHAT mot JSON hop le (khong markdown, khong giai thich) theo schema:

{schema}

Yeu cau:
- "full" la bien ban DAY DU: gom noi dung thao luan thanh cac CHU DE (topics) theo
  linh vuc / hang muc cong viec, moi topic co heading ngan va points la cac y chi tiet.
  "action_plan" liet ke viec can lam kem nguoi phu trach va han chot cu the neu co.
- "summary" la ban TOM TAT 1 trang cho CEO, uu tien rui ro/viec can duyet va tien do/ngan sach.
- Viet tieng Viet co dau day du, giu nguyen thuat ngu ky thuat (MEP, FCU, CDT...).
- TRUNG THUC voi transcript, KHONG bia them thong tin khong co trong transcript.

Tieu de cuoc hop: {title}
Muc tieu: {objective}

--- TRANSCRIPT ---
{transcript}
"""
    client = Anthropic(api_key=anthropic_key)
    msg = client.messages.create(
        model=model, max_tokens=4000,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = msg.content[0].text.strip().replace("```json", "").replace("```", "").strip()
    return json.loads(raw)


# ============================================================
# HELPERS DOCX
# ============================================================
def _add_bottom_border(paragraph):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    pPr = paragraph._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    for k, v in (("val", "single"), ("sz", "12"), ("space", "4"), ("color", "2E75B6")):
        bottom.set(qn("w:" + k), v)
    pbdr.append(bottom)
    pPr.append(pbdr)


def _setup(doc):
    s = doc.styles["Normal"]
    s.font.name = "Arial"
    s.font.size = Pt(11)


def _letterhead(doc, company, logo_path=None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if logo_path:
        try:
            p.add_run().add_picture(logo_path, width=Inches(1.8))
        except Exception:
            logo_path = None
    if not logo_path:
        r = p.add_run(company.upper())
        r.bold = True
        r.font.size = Pt(16)
        r.font.color.rgb = BLUE
    _add_bottom_border(p)


def _section(doc, text):
    """Tieu de muc lon, in dam (vd: '1. DIA DIEM...')"""
    h = doc.add_paragraph()
    h.paragraph_format.space_before = Pt(10)
    r = h.add_run(text)
    r.bold = True
    r.font.size = Pt(12)


def _labelline(doc, label, value):
    """Dong dang 'Nhan: gia tri' voi nhan in dam."""
    p = doc.add_paragraph()
    p.add_run(label + ": ").bold = True
    p.add_run(value)


# ============================================================
# BAN DAY DU (format ARTIUS)
# ============================================================
def build_full_docx(meta: dict, full: dict) -> bytes:
    doc = Document()
    _setup(doc)
    _letterhead(doc, meta.get("company", ""), meta.get("logo_path"))

    # Tieu de
    t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = t.add_run("BIEN BAN CUOC HOP"); tr.bold = True; tr.font.size = Pt(14)
    st = doc.add_paragraph(); st.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = st.add_run(meta.get("title", "").upper()); sr.bold = True; sr.font.size = Pt(12)

    # 1. Dia diem + thanh phan
    _section(doc, "1. DIA DIEM CUOC HOP VA THANH PHAN THAM DU")
    if meta.get("date"):
        _labelline(doc, "Ngay hop", meta["date"])
    if meta.get("location"):
        _labelline(doc, "Dia diem", meta["location"])
    doc.add_paragraph().add_run("Thanh phan tham du:").bold = True
    for line in [l.strip() for l in meta.get("attendees", "").splitlines() if l.strip()]:
        p = doc.add_paragraph(style="List Bullet")
        if ":" in line:  # 'BP Du an: Ms. Nhung, Mr. Sang' -> in dam phan truoc dau ':'
            lbl, val = line.split(":", 1)
            p.add_run(lbl + ":").bold = True
            p.add_run(val)
        else:
            p.add_run(line)

    # 2. Muc tieu
    _section(doc, "2. MUC TIEU CUOC HOP")
    objs = [l.strip() for l in meta.get("objective", "").splitlines() if l.strip()]
    for o in objs or ["(Khong ghi)"]:
        doc.add_paragraph(o, style="List Bullet")

    # 3. Noi dung
    _section(doc, "3. NOI DUNG CUOC HOP")
    sub = doc.add_paragraph(); sub.add_run("3.1. Cac noi dung chinh:").bold = True
    for topic in full.get("topics", []):
        p = doc.add_paragraph(style="List Bullet")
        p.add_run((topic.get("heading", "") + ": ")).bold = True
        p.add_run(" ".join(topic.get("points", [])))
    if not full.get("topics"):
        doc.add_paragraph("(Khong co)", style="List Bullet")

    sub2 = doc.add_paragraph(); sub2.add_run("3.2. Ke hoach trien khai").bold = True
    for a in full.get("action_plan", []):
        p = doc.add_paragraph(style="List Bullet")
        p.add_run(a.get("task", ""))
        if a.get("owner"):
            p.add_run(f" (Phu trach: {a['owner']})")
        if a.get("deadline"):
            p.add_run("  Han: ")
            p.add_run(a["deadline"]).font.color.rgb = RED  # ngay to do giong ban goc
    if not full.get("action_plan"):
        doc.add_paragraph("(Khong co)", style="List Bullet")

    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()


# ============================================================
# BAN TOM TAT (1 trang cho CEO)
# ============================================================
def build_summary_docx(meta: dict, s: dict) -> bytes:
    doc = Document()
    _setup(doc)
    _letterhead(doc, meta.get("company", ""), meta.get("logo_path"))

    t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = t.add_run("BAO CAO TOM TAT - CEO"); tr.bold = True; tr.font.size = Pt(14)
    m = doc.add_paragraph(); m.alignment = WD_ALIGN_PARAGRAPH.CENTER
    m.add_run(f"{meta.get('title','')}  -  {meta.get('date','')}").italic = True

    def heading(txt):
        h = doc.add_heading(txt, level=2)
        for r in h.runs:
            r.font.color.rgb = BLUE

    def bullets(items):
        if not items:
            doc.add_paragraph("(Khong co)")
        for it in items:
            doc.add_paragraph(str(it), style="List Bullet")

    heading("Tom tat dieu hanh"); doc.add_paragraph(s.get("summary", ""))
    heading("Rui ro & Viec can CEO duyet"); bullets(s.get("risks_approvals", []))
    heading("Tien do & Ngan sach du an"); bullets(s.get("progress_budget", []))
    heading("Quyet dinh da thong qua"); bullets(s.get("decisions", []))

    heading("Action items")
    actions = s.get("action_items", [])
    if actions:
        table = doc.add_table(rows=1, cols=3); table.style = "Light Grid Accent 1"
        hdr = table.rows[0].cells
        hdr[0].text, hdr[1].text, hdr[2].text = "Nguoi phu trach", "Cong viec", "Han chot"
        for a in actions:
            c = table.add_row().cells
            c[0].text = str(a.get("owner", "")); c[1].text = str(a.get("task", "")); c[2].text = str(a.get("deadline", ""))
    else:
        doc.add_paragraph("(Khong co)")

    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()
