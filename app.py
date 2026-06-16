"""
app.py — Web app tao bien ban cuoc hop (giao dien Streamlit).
Xuat 2 ban tu 1 lan ghi am:
  - Bien ban DAY DU (format ARTIUS)
  - Bao cao TOM TAT 1 trang cho CEO

Cai dat: pip install streamlit assemblyai anthropic python-docx
Chay:    streamlit run app.py
"""

import os
import tempfile
import json

import streamlit as st
import assemblyai as aai

import report_builder as rb

# ============================================================
# CAU HINH — doc key tu st.secrets (.streamlit/secrets.toml khi local;
# muc "Secrets" tren Streamlit Cloud khi deploy)
# ============================================================
ASSEMBLYAI_API_KEY = st.secrets["ASSEMBLYAI_API_KEY"]
ANTHROPIC_API_KEY = st.secrets["ANTHROPIC_API_KEY"]
CLAUDE_MODEL = st.secrets.get("CLAUDE_MODEL", "claude-sonnet-4-6")

# Tu vung chuyen nganh -> giup STT nghe dung thuat ngu & viet tat.
# THEM ten rieng / du an / nha thau hay gap vao day de tang do chinh xac.
WORD_BOOST = [
    "MEP", "FCU", "CĐT", "PCCC", "ARTIUS", "VIAS", "BOD",
    "Spa", "steam", "dự toán", "thầu phụ", "cục nóng",
]

aai.settings.api_key = ASSEMBLYAI_API_KEY

# Logo mac dinh di kem app (de canh app.py). Upload trong app se ghi de logo nay.
DEFAULT_LOGO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.jpg")


# ============================================================
# STT + DIARIZATION (co custom vocabulary)
# ============================================================
def transcribe(audio_bytes, suffix, speakers_expected):
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(audio_bytes)
        path = tmp.name
    config = aai.TranscriptionConfig(
        language_code="vi",
        speaker_labels=True,
        speakers_expected=speakers_expected,
        word_boost=WORD_BOOST,
        boost_param="high",
    )
    t = aai.Transcriber().transcribe(path, config)
    if t.status == aai.TranscriptStatus.error:
        raise RuntimeError(t.error)
    return t.utterances


# ============================================================
# GIAO DIEN
# ============================================================
st.set_page_config(page_title="Biên bản cuộc họp", page_icon="📝")
st.title("📝 Tạo biên bản cuộc họp")

with st.sidebar:
    st.header("Thông tin cuộc họp")
    company = st.text_input("Tên công ty", "ARTIUS")
    tagline = st.text_input("Slogan dưới logo", "BEYOND DESIGN AND BUILD")
    logo_file = st.file_uploader("Logo công ty (tùy chọn, nền trắng/trong suốt)",
                                 type=["png", "jpg", "jpeg"])
    title = st.text_input("Tiêu đề cuộc họp", "Họp giao ban tuần")
    date = st.text_input("Ngày họp", "")
    location = st.text_input("Địa điểm", "")
    objective = st.text_area("Mục tiêu cuộc họp", height=70,
                             help="Mỗi dòng một ý")
    attendees = st.text_area(
        "Thành phần tham dự", height=130,
        help="Mỗi dòng một nhóm, ví dụ:\nBP Dự án: Ms. Nhung, Mr. Sang",
    )
    n_speakers = st.number_input("Số người dự kiến", 1, 15, 4)

uploaded = st.file_uploader(
    "Kéo-thả file ghi âm vào đây",
    type=["mp3", "wav", "m4a", "mp4", "aac"],
)

# --- Buoc 1: transcribe ---
if uploaded and st.button("1️⃣ Tách lời & nhận diện người nói", type="primary"):
    with st.spinner("Đang xử lý (có thể mất vài phút)..."):
        suffix = "." + uploaded.name.split(".")[-1]
        utts = transcribe(uploaded.getvalue(), suffix, int(n_speakers))
        st.session_state["utts"] = [{"speaker": u.speaker, "text": u.text} for u in utts]
        st.success("Xong! Cuộn xuống để gán tên người nói.")

# --- Buoc 2: gan ten ---
if "utts" in st.session_state:
    utts = st.session_state["utts"]
    speakers = sorted({u["speaker"] for u in utts})

    st.subheader("2️⃣ Gán tên & vai trò")
    st.caption('Đọc câu mẫu rồi điền tên. Ví dụ: "Mr. Sang - BP Dự án"')
    mapping = {}
    for spk in speakers:
        sample = next(u["text"] for u in utts if u["speaker"] == spk)
        st.markdown(f"**Speaker {spk}** — _mẫu:_ \"{sample[:90]}...\"")
        mapping[spk] = st.text_input(f"Tên cho Speaker {spk}", spk, key=f"name_{spk}")

    # --- Buoc 3: tao 2 ban ---
    st.subheader("3️⃣ Tạo biên bản")
    if st.button("📄 Tạo 2 bản (đầy đủ + tóm tắt)", type="primary"):
        labeled = "\n".join(f"[{mapping[u['speaker']]}] {u['text']}" for u in utts)

        # Luu logo ra file tam (neu co)
        logo_path = DEFAULT_LOGO if os.path.exists(DEFAULT_LOGO) else None
        if logo_file:  # nguoi dung upload -> ghi de logo mac dinh
            suffix = "." + logo_file.name.split(".")[-1]
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as lf:
                lf.write(logo_file.getvalue())
                logo_path = lf.name

        meta = {
            "company": company, "tagline": tagline, "title": title, "date": date,
            "location": location, "objective": objective,
            "attendees": attendees, "logo_path": logo_path,
        }

        with st.spinner("Claude đang viết 2 bản báo cáo..."):
            try:
                data = rb.generate_reports(labeled, title, objective,
                                           ANTHROPIC_API_KEY, CLAUDE_MODEL)
            except json.JSONDecodeError:
                st.error("Claude trả về không đúng JSON. Thử lại hoặc đổi sang model opus.")
                st.stop()

        full_bytes = rb.build_full_docx(meta, data.get("full", {}))
        summary_bytes = rb.build_summary_docx(meta, data.get("summary", {}))

        st.success("Đã tạo xong 2 bản!")
        st.markdown("**Tóm tắt nhanh:** " + data.get("summary", {}).get("summary", ""))

        DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        c1, c2 = st.columns(2)
        c1.download_button("⬇️ Biên bản ĐẦY ĐỦ (.docx)", full_bytes,
                           file_name="bien_ban_day_du.docx", mime=DOCX)
        c2.download_button("⬇️ Báo cáo TÓM TẮT (.docx)", summary_bytes,
                           file_name="bao_cao_tom_tat.docx", mime=DOCX)
        st.download_button("⬇️ Transcript thô (.txt)", labeled.encode("utf-8"),
                           file_name="transcript.txt", mime="text/plain")
