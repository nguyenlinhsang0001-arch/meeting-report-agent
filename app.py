"""
app.py — ARTIUS Meeting Report Agent
Upload audio -> AssemblyAI (STT tieng Viet + speaker diarization)
            -> gan ten/vai tro cho tung speaker
            -> Claude sinh 1 ban bien ban (report_builder.generate_report)
            -> xuat 1 file Word (.docx)

Khop voi report_builder.py ban 1-file:
  - generate_report(transcript, title, key, model) -> 1 dict
  - build_report_docx(meta, report) -> bytes .docx
"""

import os
import tempfile

import streamlit as st
import assemblyai as aai

import report_builder as rb

# ============================================================
# CAU HINH CO DINH
# ============================================================
st.set_page_config(page_title="ARTIUS - Bien ban cuoc hop", page_icon="📝", layout="centered")

CLAUDE_MODELS = {
    "Sonnet (nhanh, mac dinh)": "claude-sonnet-4-6",
    "Opus (manh hon, cham hon)": "claude-opus-4-8",
}

# Tu khoa chuyen nganh -> giup AssemblyAI nhan dung thuat ngu
WORD_BOOST = [
    "ARTIUS", "VIAS", "MEP", "FCU", "AHU", "PAU", "CDT", "CĐT", "PCCC",
    "BOD", "BQL", "spa", "steam", "GYM", "thau phu", "bao gia", "du toan",
    "thi cong", "ban giao", "nghiem thu", "ngan sach", "tien do",
]

# ============================================================
# DOC API KEY TU SECRETS (dung TEN secret, khong dung gia tri)
# ============================================================
def _get_secret(name):
    try:
        return st.secrets[name].strip()
    except Exception:
        return ""

ASSEMBLYAI_API_KEY = _get_secret("ASSEMBLYAI_API_KEY")
ANTHROPIC_API_KEY = _get_secret("ANTHROPIC_API_KEY")

missing = [n for n, v in [("ASSEMBLYAI_API_KEY", ASSEMBLYAI_API_KEY),
                          ("ANTHROPIC_API_KEY", ANTHROPIC_API_KEY)] if not v]
if missing:
    st.error(
        "Thieu API key trong Secrets: " + ", ".join(missing) +
        ".\nVao Manage app -> Settings -> Secrets va them dong dang:\n"
        'ANTHROPIC_API_KEY = "sk-ant-..."\nASSEMBLYAI_API_KEY = "..."'
    )
    st.stop()

aai.settings.api_key = ASSEMBLYAI_API_KEY

# ============================================================
# GIAO DIEN
# ============================================================
st.title("📝 Biên bản cuộc họp ARTIUS")
st.caption("Upload file ghi âm → tự phiên âm, nhận diện người nói → sinh biên bản Word.")

with st.sidebar:
    st.subheader("Cấu hình")
    model_label = st.selectbox("Model Claude", list(CLAUDE_MODELS.keys()), index=0)
    CLAUDE_MODEL = CLAUDE_MODELS[model_label]

# --- Thong tin cuoc hop (chi nhap toi thieu; muc tieu & thanh phan agent tu rut ra) ---
title = st.text_input("Tiêu đề cuộc họp", placeholder="VD: Họp thầu phụ MEP VIAS Vân Phong GĐ3")
col1, col2 = st.columns(2)
with col1:
    meeting_date = st.text_input("Ngày họp", placeholder="VD: 15/06/2026")
with col2:
    meeting_location = st.text_input("Địa điểm", placeholder="VD: Phòng họp lớn, ARTIUS")

audio_file = st.file_uploader("File ghi âm", type=["mp3", "wav", "m4a", "mp4", "aac", "ogg", "flac"])

# ============================================================
# BUOC 1: PHIEN AM (chay AssemblyAI, luu vao session_state)
# ============================================================
def transcribe(file):
    """Tra ve list utterances (moi phan tu co .speaker, .text) hoac raise loi ro rang."""
    suffix = os.path.splitext(file.name)[1] or ".mp3"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(file.getvalue())
        audio_path = tmp.name

    config = aai.TranscriptionConfig(
        language_code="vi",        # tieng Viet
        speaker_labels=True,       # nhan dien nguoi noi (diarization)
        word_boost=WORD_BOOST,
        boost_param="high",
    )
    transcript = aai.Transcriber().transcribe(audio_path, config=config)

    # Bat loi ro rang thay vi de TranscriptError bi redact
    if transcript.status == aai.TranscriptStatus.error:
        raise RuntimeError(f"AssemblyAI báo lỗi: {transcript.error}")
    if not transcript.utterances:
        raise RuntimeError("Không tách được người nói. Kiểm tra file âm thanh có tiếng nói rõ không.")
    return transcript.utterances


if st.button("① Phiên âm", type="primary", disabled=audio_file is None):
    try:
        with st.spinner("Đang phiên âm và nhận diện người nói..."):
            utts = transcribe(audio_file)
        # Luu utterances duoi dang list dict (de session_state giu duoc)
        st.session_state["utterances"] = [{"speaker": u.speaker, "text": u.text} for u in utts]
        # Reset mapping cu
        st.session_state.pop("speaker_names", None)
        st.success(f"Xong! Phát hiện {len({u.speaker for u in utts})} người nói.")
    except Exception as e:
        st.error(str(e))

# ============================================================
# BUOC 2: GAN TEN / VAI TRO CHO TUNG SPEAKER
# ============================================================
utterances = st.session_state.get("utterances")

if utterances:
    speakers = sorted({u["speaker"] for u in utterances})
    speaker_count = len(speakers)  # (3) so nguoi tu dem trong file ghi am

    st.divider()
    st.subheader(f"② Gán tên người nói ({speaker_count} người)")
    st.caption("Nghe lại nếu cần, rồi điền tên + vai trò cho từng giọng. Để trống sẽ giữ nhãn mặc định.")

    names = {}
    for spk in speakers:
        # Hien 1 cau mau cua speaker do de de nhan ra giong
        sample = next((u["text"] for u in utterances if u["speaker"] == spk), "")
        c1, c2 = st.columns(2)
        with c1:
            nm = st.text_input(f"Người nói {spk} — Tên", key=f"name_{spk}",
                               placeholder="VD: Mr. Sang")
        with c2:
            role = st.text_input(f"Người nói {spk} — Vai trò/Bộ phận", key=f"role_{spk}",
                                 placeholder="VD: BP Dự án")
        st.caption(f"🔊 Mẫu: “{sample[:90]}…”" if sample else "")
        names[spk] = (nm.strip(), role.strip())

    # ----- Dung transcript co gan ten de gui cho Claude -----
    def build_labeled_transcript():
        lines = []
        for u in utterances:
            nm, role = names.get(u["speaker"], ("", ""))
            who = nm or f"Người nói {u['speaker']}"
            if role:
                who = f"{who} ({role})"
            lines.append(f"{who}: {u['text']}")
        return "\n".join(lines)

    st.divider()

    # ============================================================
    # BUOC 3: SINH BIEN BAN -> XUAT 1 FILE WORD
    # ============================================================
    if st.button("③ Tạo biên bản", type="primary", disabled=not title):
        labeled = build_labeled_transcript()
        meta = {
            "title": title,
            "date": meeting_date,
            "location": meeting_location,
            "speaker_count": speaker_count,   # (3) uu tien so nay khi dem nguoi du
        }
        try:
            with st.spinner("Claude đang soạn biên bản..."):
                report = rb.generate_report(labeled, title, ANTHROPIC_API_KEY, CLAUDE_MODEL)
                docx_bytes = rb.build_report_docx(meta, report)
            st.success("Đã tạo xong biên bản!")
            st.download_button(
                "📄 Tải biên bản (.docx)",
                docx_bytes,
                file_name="bien_ban.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        except Exception as e:
            st.error(str(e))
    elif not title:
        st.info("Nhập **Tiêu đề cuộc họp** ở trên trước khi tạo biên bản.")
