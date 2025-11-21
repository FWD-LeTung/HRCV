import os
import json
#import pdfplumber
from openai import OpenAI
from datetime import datetime

# --- CẤU HÌNH ---
# 1. Điền API Key DeepSeek của bạn
DEEPSEEK_API_KEY = "" 

# 2. Đường dẫn thư mục
INPUT_FOLDER = "D:/CV" 
OUTPUT_FILE = "test_tool_2.xlsx"

# Khởi tạo client
client = OpenAI(
    api_key=DEEPSEEK_API_KEY,
    base_url="https://api.deepseek.com"
)

# Định nghĩa danh sách cột chuẩn (0 -> 15) left to right
COLUMNS = [
    "Dấu thời gian",                                                # 0: Empty
    "Họ và tên của bạn?",                                           # 1: AI Extract
    "Ngày tháng năm sinh của bạn?",                                 # 2: AI Extract
    "Giới tính của bạn?",                                           # 3: AI Extract
    "Số điện thoại của bạn?",                                       # 4: AI Extract
    "Email của bạn",                                                # 5: AI Extract
    "Bạn đang là sinh viên năm thứ mấy?",                           # 6: AI Extract (Infer)
    "Bạn học chuyên ngành gì? Và theo học tại trường nào?",         # 7: AI Extract (Combined)
    "Điểm GPA (Tính đến thời điểm hiện tại) của bạn?",              # 8: AI Extract
    "Tiếng Anh hiện tại của bạn? (Điểm TOEIC/IELTS...)",            # 9: AI Extract
    "Bạn biết đến thông tin chương trình qua đâu?",                 # 10: Empty
    "Nếu có thể trở thành thực tập sinh...",                        # 11: Empty
    "Bạn hãy upload CV ứng tuyển của bạn nhé...",                   # 12: Empty
    "Phone Interview",                                              # 13: Empty
    "Hr Review",                                                    # 14: Empty
    "Hiring Manager Review"                                         # 15: Empty
]

def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except Exception as e:
        print(f"Lỗi đọc file {pdf_path}: {e}")
    return text

def analyze_cv_with_deepseek(cv_text):
    # Lấy ngày hiện tại để AI tính toán năm học
    current_date = datetime.now().strftime("%m/%Y")
    
    prompt = f"""
    Thời điểm hiện tại là: {current_date}.
    Nhiệm vụ: Trích xuất thông tin từ CV dưới đây để điền vào Excel tuyển dụng.
    
    Yêu cầu quan trọng:
    1. Trả về JSON object.
    2. Nếu thông tin nào KHÔNG tìm thấy trong CV, hãy điền giá trị chính xác là chuỗi: "Can't find" (Không được để null hay trống).
    3. Logic trích xuất:
       - 'student_year': Dựa vào niên khóa trong CV và thời điểm hiện tại ({current_date}) để tính xem đang là sinh viên năm mấy (Ví dụ: "Năm 3", "Năm 4", "Đã tốt nghiệp").
       - 'major_university': Gộp "Chuyên ngành" và "Tên trường" thành 1 câu (Ví dụ: "CNTT - ĐH Bách Khoa").
       - 'gender' nếu không có hãy dựa vào tên để nhận biết là 'Nam' hay 'Nữ'. Nếu lọc được male hoặc female thì đổi sang tiếng việt tương ứng
       - 'name' sau khi có thông tin nên chỉnh về cùng một định dạng: chỉ viết hoa chữ cái đầu (ví dụ: LE VAN TUNG -> Lê Văn Tùng)
    
    Các trường JSON cần trả về:
    {{
        "full_name": "...",
        "dob": "...",
        "gender": "...",
        "phone": "...",
        "email": "...",
        "student_year": "...",
        "major_university": "...",
        "gpa": "...",
        "english_skill": "..."
    }}

    Nội dung CV:
    {cv_text}
    """

    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "Bạn là trợ lý HR chính xác tuyệt đối. Output JSON only."},
                {"role": "user", "content": prompt},
            ],
            temperature=0.1,
            max_tokens=1024,
            response_format={ "type": "json_object" }
        )
        return json.loads(response.choices[0].message.content)
    except Exception as e:
        print(f"Lỗi API: {e}")
        return None

def main():
    if not os.path.exists(INPUT_FOLDER):
        os.makedirs(INPUT_FOLDER)
        print(f"Đã tạo folder '{INPUT_FOLDER}'. Vui lòng copy PDF vào đây.")
        return

    files = [f for f in os.listdir(INPUT_FOLDER) if f.lower().endswith('.pdf')]
    if not files:
        print("Không có file PDF nào.")
        return

    results = []
    print(f"Bắt đầu xử lý {len(files)} hồ sơ...")

    for filename in files:
        print(f" >> Đang đọc: {filename}")
        text = extract_text_from_pdf(os.path.join(INPUT_FOLDER, filename))
        
        if not text.strip():
            print("    -> File trống hoặc dạng ảnh (scan). Bỏ qua.")
            continue
            
        data = analyze_cv_with_deepseek(text)
        
        if data:
            # Ánh xạ dữ liệu từ JSON vào đúng thứ tự cột Excel (0 -> 15)
            row = {
                COLUMNS[0]: "",                      # Dấu thời gian
                COLUMNS[1]: data.get("full_name", "Can't find"),
                COLUMNS[2]: data.get("dob", "Can't find"),
                COLUMNS[3]: data.get("gender", "Can't find"),
                COLUMNS[4]: data.get("phone", "Can't find"),
                COLUMNS[5]: data.get("email", "Can't find"),
                COLUMNS[6]: data.get("student_year", "Can't find"),
                COLUMNS[7]: data.get("major_university", "Can't find"),
                COLUMNS[8]: data.get("gpa", "Can't find"),
                COLUMNS[9]: data.get("english_skill", "Can't find"),
                COLUMNS[10]: "",                     # Nguồn tin
                COLUMNS[11]: "",                     # Thời gian làm việc
                COLUMNS[12]: "",                     # Upload CV (User request empty)
                COLUMNS[13]: "",                     # Phone Interview
                COLUMNS[14]: "",                     # Hr Review
                COLUMNS[15]: ""                      # HM Review
            }
            results.append(row)
        else:
            print("    -> Lỗi xử lý AI.")

    # Xuất Excel
    if results:
        df = pd.DataFrame(results)
        df = df[COLUMNS] 
        df.to_excel(OUTPUT_FILE, index=False)
        print(f"\nHoàn tất! File kết quả: {OUTPUT_FILE}")
    else:
        print("\nKhông có dữ liệu.")

if __name__ == "__main__":
    main()