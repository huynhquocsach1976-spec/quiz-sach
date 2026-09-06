import streamlit as st
import os
import json
import time
from docx import Document
from io import BytesIO
import pypdf
from google import genai

# =============================================================================
# 1. CẤU HÌNH TRANG WEB & GEMINI CLIENT
# =============================================================================
st.set_page_config(
    page_title="Hệ Thống Đề Thi Trắc Nghiệm AI",
    page_icon="📝",
    layout="wide"
)

TEACHER_PASSWORD = "123456"
QUIZ_FILE_PATH = "quiz_data.json"

api_key = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")

@st.cache_resource
def get_gemini_client(key):
    if not key:
        return None
    return genai.Client(api_key=key)

client = get_gemini_client(api_key)

# =============================================================================
# 2. HÀM ĐỌC / GHI ĐỀ THI
# =============================================================================
def save_quiz_to_file(quiz_data, topic):
    payload = {
        "topic": topic,
        "quiz_data": quiz_data
    }
    with open(QUIZ_FILE_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

def load_quiz_from_file():
    if os.path.exists(QUIZ_FILE_PATH):
        try:
            with open(QUIZ_FILE_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None
    return None

# =============================================================================
# 3. HÀM ĐỌC FILE VĂN BẢN
# =============================================================================
def extract_text_from_file(uploaded_file):
    text = ""
    file_type = uploaded_file.name.split('.')[-1].lower()
    
    try:
        if file_type == 'txt':
            text = uploaded_file.read().decode("utf-8")
        elif file_type == 'pdf':
            pdf_reader = pypdf.PdfReader(uploaded_file)
            for page in pdf_reader.pages:
                text += page.extract_text() + "\n"
        elif file_type in ['docx', 'doc']:
            doc = Document(uploaded_file)
            for para in doc.paragraphs:
                text += para.text + "\n"
    except Exception as e:
        st.error(f"Lỗi khi đọc file: {e}")
    return text

# =============================================================================
# 4. HÀM TẠO ĐỀ THI BẰNG GEMINI AI
# =============================================================================
def generate_quiz_json(document_text, num_questions, difficulty, topic):
    if not client:
        return None, "⚠️ Chưa cấu hình GEMINI_API_KEY trong Secrets."

    prompt = f"""
    Bạn là một chuyên gia ra đề thi giáo dục. Hãy đọc tài liệu dưới đây và tạo ra chính xác {num_questions} câu hỏi trắc nghiệm.
    
    Chủ đề/Môn học: {topic}
    Mức độ khó: {difficulty}
    
    Nội dung tài liệu tham khảo:
    \"\"\"
    {document_text[:10000]}
    \"\"\"
    
    YÊU CẦU ĐẦU RA:
    Trả về kết quả dưới dạng một chuỗi JSON thuần túy (JSON Array), KHÔNG chứa ký tự markdown ```json ở đầu hoặc cuối.
    Cấu trúc mỗi phần tử trong mảng JSON như sau:
    [
      {{
        "id": 1,
        "question": "Nội dung câu hỏi?",
        "options": {{
          "A": "Đáp án A",
          "B": "Đáp án B",
          "C": "Đáp án C",
          "D": "Đáp án D"
        }},
        "correct_answer": "A",
        "explanation": "Giải thích chi tiết tại sao A đúng dựa vào tài liệu."
      }}
    ]
    """

    max_retries = 3
    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
                model='gemini-3.6-flash',
                contents=prompt,
            )
            clean_json = response.text.strip().replace("```json", "").replace("```", "")
            quiz_data = json.loads(clean_json)
            return quiz_data, None
        except Exception as e:
            error_msg = str(e)
            if "503" in error_msg or "UNAVAILABLE" in error_msg:
                if attempt < max_retries - 1:
                    time.sleep(2 * (attempt + 1))
                    continue
                return None, "⏳ Máy chủ Gemini đang quá tải tạm thời. Vui lòng bấm thử lại!"
            return None, f"Lỗi tạo câu hỏi từ Gemini: {e}"

# =============================================================================
# 5. HÀM XUẤT FILE WORD (.DOCX)
# =============================================================================
def create_word_file(quiz_data, title="ĐỀ THI TRẮC NGHIỆM"):
    doc = Document()
    doc.add_heading(title, level=0)
    
    doc.add_heading("PHẦN 1: CÂU HỎI TRẮC NGHIỆM", level=1)
    for q in quiz_data:
        p = doc.add_paragraph()
        p.add_run(f"Câu {q['id']}: {q['question']}").bold = True
        for key, value in q['options'].items():
            doc.add_paragraph(f"  {key}. {value}")
        doc.add_paragraph("")
        
    doc.add_page_break()
    doc.add_heading("PHẦN 2: ĐÁP ÁN & GIẢI THÍCH CHI TIẾT", level=1)
    for q in quiz_data:
        p = doc.add_paragraph()
        p.add_run(f"Câu {q['id']}: Đáp án đúng {q['correct_answer']}").bold = True
        doc.add_paragraph(f"Giải thích: {q['explanation']}")
        doc.add_paragraph("")
        
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# =============================================================================
# 6. PHÂN LUỒNG TRUY CẬP TỪ URL (LINK RIÊNG)
# =============================================================================
query_params = st.query_params
mode_param = query_params.get("mode", None)

if mode_param == "student":
    role = "👨‍🎓 Sinh viên (Làm bài)"
else:
    st.sidebar.title("🎮 Chế Độ Sử Dụng")
    role = st.sidebar.radio("Bạn là:", ["👨‍🎓 Sinh viên (Làm bài)", "👨‍🏫 Giáo viên (Soạn đề)"])

# -----------------------------------------------------------------------------
# CHẾ ĐỘ 1: GIÁO VIÊN SOẠN ĐỀ
# -----------------------------------------------------------------------------
if role == "👨‍🏫 Giáo viên (Soạn đề)":
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔒 Xác thực Quyền Giáo viên")
    input_password = st.sidebar.text_input("Nhập mật khẩu Giáo viên:", type="password")

    if input_password != TEACHER_PASSWORD:
        st.title("🔒 Chế Độ Giáo Viên Bị Khóa")
        st.warning("⚠️ Vui lòng nhập đúng mật khẩu Giáo viên ở thanh menu bên trái.")
        st.stop()

    st.title("📝 Hệ Thống Soạn Đề Thi Trắc Nghiệm Tự Động")
    st.success("🔓 Xác thực thành công! Quyền Giáo viên đã được kích hoạt.")
    
    st.sidebar.markdown("---")
    st.sidebar.header("⚙️ Cấu Hình Đề Thi")
    uploaded_file = st.sidebar.file_uploader("Tải tài liệu bài giảng (PDF, DOCX, TXT):", type=["pdf", "docx", "txt"])
    topic_input = st.sidebar.text_input("Tên môn học / Bài học:", value="UML - Chapter 5")
    num_q = st.sidebar.slider("Số lượng câu hỏi:", min_value=3, max_value=20, value=5)
    difficulty_level = st.sidebar.selectbox("Độ khó:", ["Nhận biết (Dễ)", "Thông hiểu (Trung bình)", "Vận dụng (Khó)"])

    if uploaded_file is not None:
        doc_text = extract_text_from_file(uploaded_file)
        if len(doc_text.strip()) < 50:
            st.warning("⚠️ Tài liệu quá ngắn hoặc không đọc được nội dung.")
        else:
            st.success(f"Trích xuất thành công {len(doc_text)} ký tự từ tài liệu: **{uploaded_file.name}**")
            
            if st.button("🚀 Bắt Đầu Tạo Đề Thi Trắc Nghiệm", type="primary"):
                with st.spinner("Gemini AI đang nghiên cứu tài liệu và soạn câu hỏi..."):
                    quiz_data, err = generate_quiz_json(doc_text, num_q, difficulty_level, topic_input)
                    if err:
                        st.error(err)
                    else:
                        save_quiz_to_file(quiz_data, topic_input)
                        st.success("✨ Đã tạo & KÍCH HOẠT ĐỀ THI thành công!")

    saved_quiz = load_quiz_from_file()
    if saved_quiz:
        st.markdown("---")
        st.header(f"📋 Đề Thi Đang Đang Được Kích Hoạt: {saved_quiz['topic']}")
        
        # HIỂN THỊ LINK ĐỂ GIÁO VIÊN COPY GỬI SINH VIÊN
        base_url = "[https://quiz-sach.streamlit.app](https://quiz-sach.streamlit.app)"  # Thay bằng URL thực tế của bạn nếu khác
        student_link = f"{base_url}/?mode=student"
        
        st.info(f"🔗 **Đường link riêng gửi cho Sinh viên làm bài:**\n`{student_link}`")
        
        for q in saved_quiz['quiz_data']:
            st.markdown(f"**Câu {q['id']}: {q['question']}**")
            for key, val in q['options'].items():
                st.write(f"- **{key}.** {val}")
            
            with st.expander(f"🔑 Xem đáp án Câu {q['id']}"):
                st.write(f"**Đáp án đúng:** `{q['correct_answer']}`")
                st.write(f"**Giải thích:** {q['explanation']}")
            st.write("")
            
        st.markdown("---")
        docx_file = create_word_file(saved_quiz['quiz_data'], f"ĐỀ THI: {saved_quiz['topic'].upper()}")
        st.download_button(
            label="📥 Tải Đề Thi File Word (.docx)",
            data=docx_file,
            file_name=f"De_Thi_{saved_quiz['topic']}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

# -----------------------------------------------------------------------------
# CHẾ ĐỘ 2: SINH VIÊN LÀM BÀI TRỰC TIẾP
# -----------------------------------------------------------------------------
else:
    st.title("✍️ BÀI THI TRẮC NGHIỆM TRỰC TUYẾN")
    
    saved_quiz = load_quiz_from_file()
    
    if not saved_quiz:
        st.info("📌 Hiện tại chưa có đề thi nào được kích hoạt. Vui lòng nhờ Giáo viên tạo đề trước!")
    else:
        quiz_data = saved_quiz['quiz_data']
        topic = saved_quiz['topic']
        st.subheader(f"Môn học / Chủ đề: {topic}")
        st.write(f"Tổng số câu hỏi: **{len(quiz_data)} câu**")
        st.markdown("---")

        with st.form("quiz_form"):
            user_answers = {}
            for q in quiz_data:
                st.markdown(f"##### Câu {q['id']}: {q['question']}")
                options_fmt = [f"{k}. {v}" for k, v in q['options'].items()]
                
                selected = st.radio(
                    f"Chọn đáp án cho câu {q['id']}:",
                    options_fmt,
                    index=None,
                    key=f"q_{q['id']}"
                )
                if selected:
                    user_answers[q['id']] = selected.split('.')[0]
                st.write("")

            submitted = st.form_submit_button("🏁 NỘP BÀI THI", type="primary")

        if submitted:
            score = 0
            total = len(quiz_data)
            
            st.markdown("---")
            st.header("📊 KẾT QUẢ BÀI THI")
            
            for q in quiz_data:
                q_id = q['id']
                user_ans = user_answers.get(q_id, "Chưa chọn")
                correct_ans = q['correct_answer']
                
                if user_ans == correct_ans:
                    score += 1
                    st.success(f"✅ **Câu {q_id}: ĐÚNG** | Bạn chọn: **{user_ans}**")
                else:
                    st.error(f"❌ **Câu {q_id}: SAI** | Bạn chọn: **{user_ans}** | Đáp án đúng: **{correct_ans}**")
                
                st.write(f"💡 *Giải thích:* {q['explanation']}")
                st.write("")

            final_score = round((score / total) * 10, 2)
            st.balloons()
            st.metric("Tổng Điểm Của Bạn", f"{final_score} / 10 Điểm", delta=f"Đúng {score}/{total} câu")