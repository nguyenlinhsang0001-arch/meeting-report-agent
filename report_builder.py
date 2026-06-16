"""
report_builder.py — Sinh 2 ban bao cao tu transcript va xuat Word theo format ARTIUS.
Khong import streamlit -> test doc lap duoc.
"""

import io
import json

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

BLUE = RGBColor(0x2E, 0x75, 0xB6)
RED = RGBColor(0xC0, 0x00, 0x00)
GRAY = RGBColor(0x55, 0x55, 0x55)


# ============================================================
# GOI CLAUDE -> JSON 2 BAN
# ============================================================
def generate_reports(transcript, title, objective, anthropic_key,
                     model="claude-sonnet-4-6"):
    from anthropic import Anthropic  # import lazy

    schema = """{
  "full": {
    "content_heading": "Ten nhom noi dung (vd: Cac noi dung MEP)",
    "topics": [{"heading": "Ten chu de", "points": ["y chi tiet 1", "y chi tiet 2"]}],
    "action_plan": [{"task": "viec can lam", "owner": "ten/don vi phu trach", "deadline": "ngay cu the hoac 'Chua xac dinh'"}]
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
Tu transcript ben duoi (da ghi ro ten + vai tro nguoi noi), tao DONG THOI 2 tai lieu va
tra ve DUY NHAT mot JSON hop le (khong markdown) theo schema:

{schema}

Yeu cau:
- "full" la bien ban DAY DU: dat "content_heading" cho phu hop loai hop; gom noi dung thao
  luan thanh cac CHU DE (topics), moi topic co heading ngan va points la cac y chi tiet.
  "action_plan" liet ke viec can lam kem nguoi phu trach va han chot cu the neu co.
- "summary" la ban TOM TAT 1 trang cho CEO, uu tien rui ro/viec can duyet va tien do/ngan sach.
- Viet tieng Viet co dau day du, giu nguyen thuat ngu ky thuat (MEP, FCU, CDT, PCCC...).
- TRUNG THUC voi transcript, KHONG bia them.

Tieu de cuoc hop: {title}
Muc tieu: {objective}

--- TRANSCRIPT ---
{transcript}
"""
    client = Anthropic(api_key=anthropic_key)
    msg = client.messages.create(model=model, max_tokens=4000,
                                 messages=[{"role": "user", "content": prompt}])
    raw = msg.content[0].text.strip().replace("```json", "").replace("```", "").strip()
    return json.loads(raw)


# ============================================================
# HELPERS DOCX
# ============================================================
def _bottom_border(paragraph, color="000000", sz="8"):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    pPr = paragraph._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    for k, v in (("val", "single"), ("sz", sz), ("space", "4"), ("color", color)):
        bottom.set(qn("w:" + k), v)
    pbdr.append(bottom)
    pPr.append(pbdr)


def _setup(doc):
    s = doc.styles["Normal"]
    s.font.name = "Arial"
    s.font.size = Pt(11)


def _letterhead(doc, company, tagline="", logo_path=None):
    """Dat o HEADER cua section -> lap lai dau moi trang, can trai (giong template)."""
    header = doc.sections[0].header
    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    used_image = False
    if logo_path:
        try:
            p.add_run().add_picture(logo_path, width=Inches(2.0))
            used_image = True
        except Exception:
            used_image = False
    if not used_image:
        # Tai hien wordmark bang chu (gian cach chu cho giong logo)
        r = p.add_run(" ".join(company.upper()))
        r.font.size = Pt(20)
        r.font.color.rgb = RGBColor(0, 0, 0)
        if tagline:
            sub = header.add_paragraph()
            sr = sub.add_run("   ".join(tagline.upper().split()))
            sr.font.size = Pt(7.5)
            sr.font.color.rgb = GRAY
            _bottom_border(sub)
            return
    _bottom_border(p)


def _section(doc, text):
    h = doc.add_paragraph()
    h.paragraph_format.space_before = Pt(12)
    h.paragraph_format.space_after = Pt(4)
    r = h.add_run(text)
    r.bold = True
    r.font.size = Pt(12)


def _sub(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.add_run(text).bold = True


def _label(doc, label, value):
    p = doc.add_paragraph()
    p.add_run(label + ": ").bold = True
    p.add_run(value)


def _dash(doc):
    """Tao 1 gach dau dong kieu dau '-' co thut hang treo (giong template)."""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Inches(0.5)
    pf.first_line_indent = Inches(-0.22)
    pf.space_after = Pt(4)
    p.add_run("-\t")
    return p


# ============================================================
# BAN DAY DU (format ARTIUS)
# ============================================================
def build_full_docx(meta, full):
    doc = Document()
    _setup(doc)
    _letterhead(doc, meta.get("company", ""), meta.get("tagline", ""), meta.get("logo_path"))

    t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t.paragraph_format.space_before = Pt(18)
    tr = t.add_run("BIÊN BẢN CUỘC HỌP"); tr.bold = True; tr.font.size = Pt(18)
    s = doc.add_paragraph(); s.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = s.add_run(meta.get("title", "").upper()); sr.bold = True; sr.font.size = Pt(12)

    # 1.
    _section(doc, "1. ĐỊA ĐIỂM CUỘC HỌP VÀ THÀNH PHẦN THAM DỰ")
    if meta.get("date"):
        _label(doc, "Ngày họp", meta["date"])
    if meta.get("location"):
        _label(doc, "Địa điểm", meta["location"])
    doc.add_paragraph().add_run("Thành phần tham dự:").bold = True
    for line in [l.strip() for l in meta.get("attendees", "").splitlines() if l.strip()]:
        p = _dash(doc)
        if ":" in line:
            lbl, val = line.split(":", 1)
            p.add_run(lbl + ":").bold = True
            p.add_run(val)
        else:
            p.add_run(line)

    # 2.
    _section(doc, "2. MỤC TIÊU CUỘC HỌP")
    for o in [l.strip() for l in meta.get("objective", "").splitlines() if l.strip()] or ["(Không ghi)"]:
        _dash(doc).add_run(o)

    # 3.
    _section(doc, "3. NỘI DUNG CUỘC HỌP")
    _sub(doc, "3.1. " + (full.get("content_heading") or "Các nội dung chính") + ":")
    for topic in full.get("topics", []):
        p = _dash(doc)
        p.add_run(topic.get("heading", "") + ": ").bold = True
        p.add_run(" ".join(topic.get("points", [])))
    if not full.get("topics"):
        _dash(doc).add_run("(Không có)")

    _sub(doc, "3.2. Kế hoạch triển khai")
    for a in full.get("action_plan", []):
        p = _dash(doc)
        p.add_run(a.get("task", ""))
        if a.get("owner"):
            p.add_run(f" (Phụ trách: {a['owner']})")
        if a.get("deadline"):
            p.add_run("  Hạn: ")
            p.add_run(a["deadline"]).font.color.rgb = RED
    if not full.get("action_plan"):
        _dash(doc).add_run("(Không có)")

    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()


# ============================================================
# BAN TOM TAT (1 trang cho CEO)
# ============================================================
def build_summary_docx(meta, s):
    doc = Document()
    _setup(doc)
    _letterhead(doc, meta.get("company", ""), meta.get("tagline", ""), meta.get("logo_path"))

    t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t.paragraph_format.space_before = Pt(18)
    tr = t.add_run("BÁO CÁO TÓM TẮT - CEO"); tr.bold = True; tr.font.size = Pt(16)
    m = doc.add_paragraph(); m.alignment = WD_ALIGN_PARAGRAPH.CENTER
    m.add_run(f"{meta.get('title','')}  -  {meta.get('date','')}").italic = True

    def heading(txt):
        h = doc.add_heading(txt, level=2)
        for r in h.runs:
            r.font.color.rgb = BLUE

    def bullets(items):
        if not items:
            _dash(doc).add_run("(Không có)")
        for it in items:
            _dash(doc).add_run(str(it))

    heading("Tóm tắt điều hành"); doc.add_paragraph(s.get("summary", ""))
    heading("Rủi ro & Việc cần CEO duyệt"); bullets(s.get("risks_approvals", []))
    heading("Tiến độ & Ngân sách dự án"); bullets(s.get("progress_budget", []))
    heading("Quyết định đã thông qua"); bullets(s.get("decisions", []))

    heading("Action items")
    actions = s.get("action_items", [])
    if actions:
        table = doc.add_table(rows=1, cols=3); table.style = "Light Grid Accent 1"
        hdr = table.rows[0].cells
        hdr[0].text, hdr[1].text, hdr[2].text = "Người phụ trách", "Công việc", "Hạn chót"
        for a in actions:
            c = table.add_row().cells
            c[0].text = str(a.get("owner", "")); c[1].text = str(a.get("task", "")); c[2].text = str(a.get("deadline", ""))
    else:
        doc.add_paragraph("(Không có)")

    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()
