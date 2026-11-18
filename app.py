import io
import re
import streamlit as st
from docx import Document

st.set_page_config(page_title="DOCX Hidden Text & ZWSP Sanitizer", page_icon="🛡️", layout="centered")

st.title("🛡️ DOCX Hidden Text & ZWSP Sanitizer")
st.markdown(
    "Word(`.docx`) 문서의 **각주(Footnotes), 표, 서식을 그대로 유지**하면서 "
    "투명 문자(Zero-Width) 및 숨겨진 텍스트(숨김 속성, 흰색 글씨, 초소형 폰트, 주석)를 탐지하고 제거합니다."
)

# 제어 문자 및 제로위스/방향성 제어 문자 패턴
ZW_PAT = re.compile(r"[\u200B-\u200F\u202A-\u202E\u2060-\u206F\uFEFF\u180E]")

with st.sidebar:
    st.header("정제 옵션 설정")
    opt_zwsp = st.checkbox("제로위스(Zero-Width) 및 제어 문자 제거", value=True)
    opt_hidden = st.checkbox("워드 '숨김 글꼴(Hidden)' 텍스트 제거", value=True)
    opt_white = st.checkbox("흰색 은닉 텍스트(#FFFFFF) 제거", value=True)
    opt_tiny = st.checkbox("초소형 폰트(< 2pt) 텍스트 제거", value=True)
    opt_comments = st.checkbox("숨겨진 주석(Comments) 영역 초기화", value=True)


def sanitize_docx(file_bytes: bytes):
    doc = Document(io.BytesIO(file_bytes))
    stats = {
        "zero_width_chars": 0,
        "hidden_runs": 0,
        "white_text_runs": 0,
        "tiny_font_runs": 0,
        "comments_cleaned": 0,
    }
    detected_snippets = []

    def clean_run(run):
        if not run.text:
            return

        # 1. 워드 숨김 속성
        if opt_hidden and run.font.hidden:
            stats["hidden_runs"] += 1
            detected_snippets.append(f"[숨김 속성] {run.text[:80]}")
            run.text = ""
            return

        # 2. 흰색(#FFFFFF) 은닉 텍스트
        if opt_white and run.font.color and run.font.color.rgb:
            rgb_str = str(run.font.color.rgb).upper()
            if rgb_str in ["FFFFFF", "RGBCOLOR(255, 255, 255)"]:
                stats["white_text_runs"] += 1
                detected_snippets.append(f"[흰색 텍스트] {run.text[:80]}")
                run.text = ""
                return

        # 3. 2pt 미만 초소형 폰트
        if opt_tiny and run.font.size and run.font.size.pt < 2.0:
            stats["tiny_font_runs"] += 1
            detected_snippets.append(f"[초소형 폰트({run.font.size.pt}pt)] {run.text[:80]}")
            run.text = ""
            return

        # 4. 제로위스 및 투명 제어 문자
        if opt_zwsp:
            matches = len(ZW_PAT.findall(run.text))
            if matches > 0:
                stats["zero_width_chars"] += matches
                run.text = ZW_PAT.sub("", run.text)

    # 본문 단락 정제
    for p in doc.paragraphs:
        for r in p.runs:
            clean_run(r)

    # 표(Table) 내부 셀 정제
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        clean_run(r)

    # 머리글 / 바닥글 정제
    for section in doc.sections:
        for header_footer in [section.header, section.footer]:
            for p in header_footer.paragraphs:
                for r in p.runs:
                    clean_run(r)

    # 각주(Footnotes) XML 파트 정제 (OSCOLA 각주 번호 및 서식 유지)
    if opt_zwsp:
        for rel in doc.part.rels.values():
            if "footnotes" in rel.target_ref:
                fn_part = rel.target_part
                fn_xml = fn_part.blob.decode("utf-8")
                cleaned_xml, count = ZW_PAT.subn("", fn_xml)
                if count > 0:
                    stats["zero_width_chars"] += count
                    fn_part._blob = cleaned_xml.encode("utf-8")

    # 주석(Comments) 영역 초기화
    if opt_comments:
        for rel in list(doc.part.rels.values()):
            if "comments" in rel.target_ref:
                stats["comments_cleaned"] += 1
                empty_comments_xml = (
                    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    b'<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
                )
                rel.target_part._blob = empty_comments_xml

    out_buffer = io.BytesIO()
    doc.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer, stats, detected_snippets


uploaded_file = st.file_uploader("정제할 Word 파일(.docx)을 업로드하세요", type=["docx"])

if uploaded_file is not None:
    if st.button("문서 정제 실행", type="primary"):
        with st.spinner("문서 스캔 및 정제 중..."):
            cleaned_buffer, stats, snippets = sanitize_docx(uploaded_file.read())

        st.success("정제가 완료되었습니다. 각주와 원본 서식은 그대로 유지되었습니다.")

        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("제로위스 문자", f"{stats['zero_width_chars']}개")
        col2.metric("숨김 속성", f"{stats['hidden_runs']}건")
        col3.metric("흰색 텍스트", f"{stats['white_text_runs']}건")
        col4.metric("초소형 폰트", f"{stats['tiny_font_runs']}건")
        col5.metric("주석 초기화", f"{stats['comments_cleaned']}건")

        if snippets:
            with st.expander("제거된 은닉 구절 미리보기"):
                for item in snippets:
                    st.code(item)

        base_name = uploaded_file.name.rsplit(".", 1)[0]
        st.download_button(
            label="📥 정제된 Word 파일(.docx) 다운로드",
            data=cleaned_buffer,
            file_name=f"{base_name}_cleaned.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
