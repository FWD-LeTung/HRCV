import os
import json
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime

# Thư viện giao diện & logic
import customtkinter as ctk
import openpyxl
import pdfplumber
from openai import OpenAI

ctk.set_appearance_mode("System")  
ctk.set_default_color_theme("blue")  

class CVParserApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        
        self.title("CV-Scanner")
        self.geometry("600x550")
        self.resizable(False, False)

        self.folder_path = tk.StringVar()
        self.api_key = tk.StringVar()
        self.is_running = False

        
        
        # Title
        self.label_title = ctk.CTkLabel(self, text="Automatic CV Scanner", font=ctk.CTkFont(size=20, weight="bold"))
        self.label_title.pack(pady=20)

        # Insert API key
        self.frame_api = ctk.CTkFrame(self)
        self.frame_api.pack(pady=10, padx=20, fill="x")
        
        self.lbl_api = ctk.CTkLabel(self.frame_api, text="DeepSeek API Key:")
        self.lbl_api.pack(side="left", padx=10)
        
        self.entry_api = ctk.CTkEntry(self.frame_api, textvariable=self.api_key, placeholder_text="sk-...", width=350, show="*")
        self.entry_api.pack(side="right", padx=10, pady=10)

        # Import Folder
        self.frame_folder = ctk.CTkFrame(self)
        self.frame_folder.pack(pady=10, padx=20, fill="x")
        
        self.btn_browse = ctk.CTkButton(self.frame_folder, text="Import CV Folder", command=self.browse_folder)
        self.btn_browse.pack(side="left", padx=10, pady=10)
        
        self.lbl_folder = ctk.CTkLabel(self.frame_folder, textvariable=self.folder_path, text_color="gray")
        self.lbl_folder.pack(side="left", padx=10)

        # Status
        self.progress_bar = ctk.CTkProgressBar(self, width=500)
        self.progress_bar.pack(pady=20)
        self.progress_bar.set(0)
        
        self.lbl_status = ctk.CTkLabel(self, text="Ready", text_color="green")
        self.lbl_status.pack(pady=5)

        # Log 
        self.textbox_log = ctk.CTkTextbox(self, width=560, height=150)
        self.textbox_log.pack(pady=10)
        self.textbox_log.configure(state="disabled") # Chỉ đọc

        # start button
        self.btn_start = ctk.CTkButton(self, text="Start Processing", command=self.start_thread, height=40, font=ctk.CTkFont(size=15, weight="bold"))
        self.btn_start.pack(pady=10)

    def log(self, message):
        """Hàm ghi log ra màn hình"""
        self.textbox_log.configure(state="normal")
        self.textbox_log.insert("end", message + "\n")
        self.textbox_log.see("end")
        self.textbox_log.configure(state="disabled")

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)
            self.log(f"Selected folder: {folder_selected}")

    def start_thread(self):
        if self.is_running:
            return
        
        key = self.api_key.get().strip()
        folder = self.folder_path.get().strip()

        if not key:
            messagebox.showerror("Error", "Insert API Key please!")
            return
        if not folder:
            messagebox.showerror("Error", "Import folder with pdf file!")
            return
            
        self.is_running = True
        self.btn_start.configure(state="disabled", text="Đang chạy...")
        self.progress_bar.set(0)
        
        threading.Thread(target=self.process_cvs, args=(key, folder), daemon=True).start()

    def process_cvs(self, api_key, input_folder):
        try:
            client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
            
            files = [f for f in os.listdir(input_folder) if f.lower().endswith('.pdf')]
            total_files = len(files)
            
            if total_files == 0:
                self.log("Error: No pdf files found")
                self.reset_ui()
                return

            self.log(f"There are {total_files} file PDF. Start scanning...")
            
            output_file = os.path.join(input_folder, f"Ket_qua_Loc_CV_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Ung Vien"
            
            headers = [
                "Dấu thời gian", "Họ và tên", "Ngày sinh", "Giới tính", "SĐT", "Email",
                "Sinh viên năm", "Chuyên ngành - Trường", "GPA", "Tiếng Anh",
                "Nguồn tin", "Thời gian làm việc", "File CV", "Phone Interview", "Hr Review", "HM Review"
            ]
            ws.append(headers)

           
            for index, filename in enumerate(files):
                self.lbl_status.configure(text=f"Đang xử lý ({index + 1}/{total_files}): {filename}")
                self.log(f">> Đang đọc: {filename}")
                
                file_path = os.path.join(input_folder, filename)
                
                # A. Đọc Text
                cv_text = self.extract_text(file_path)
                if not cv_text:
                    self.log(f"   -> Bỏ qua (Không đọc được text/Scan).")
                    self.update_progress(index + 1, total_files)
                    continue
                
                data_json = self.call_ai(client, cv_text)
                
                row_data = [""] * 16
                
                if data_json:
                    row_data[1] = data_json.get("full_name", "Can't find")
                    row_data[2] = data_json.get("dob", "Can't find")
                    row_data[3] = data_json.get("gender", "Can't find")
                    row_data[4] = data_json.get("phone", "Can't find")
                    row_data[5] = data_json.get("email", "Can't find")
                    row_data[6] = data_json.get("student_year", "Can't find")
                    row_data[7] = data_json.get("major_university", "Can't find")
                    row_data[8] = data_json.get("gpa", "Can't find")
                    row_data[9] = data_json.get("english_skill", "Can't find")
                    row_data[12] = "" 
                    
                    ws.append(row_data)
                    self.log("   -> Xong.")
                else:
                    self.log("   -> Lỗi phân tích AI.")

                self.update_progress(index + 1, total_files)

            wb.save(output_file)
            self.log(f"\nDone! Excel file path:\n{output_file}")
            messagebox.showinfo("Done", f"Processed {total_files} cv!\nFile save at: {output_file}")
            
            os.startfile(output_file)

        except Exception as e:
            self.log(f"LỖI NGHIÊM TRỌNG: {str(e)}")
            messagebox.showerror("Lỗi Crash", str(e))
        finally:
            self.reset_ui()

    def update_progress(self, current, total):
        val = current / total
        self.progress_bar.set(val)

    def reset_ui(self):
        self.is_running = False
        self.btn_start.configure(state="normal", text="START PROCESSING")
        self.lbl_status.configure(text="Hoàn tất", text_color="green")

    def extract_text(self, pdf_path):
        text = ""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t: text += t + "\n"
        except:
            return None
        return text.strip()

    def call_ai(self, client, text):
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
        {text[:6000]} 
        """
        # Lưu ý: text[:4000] tiết kiệm token
        try:
            response = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": "JSON Extractor."},
                    {"role": "user", "content": prompt},
                ],
                temperature=0.1,
                response_format={ "type": "json_object" }
            )
            return json.loads(response.choices[0].message.content)
        except Exception as e:
            self.log(f"Lỗi API: {e}")
            return None

if __name__ == "__main__":
    app = CVParserApp()
    app.mainloop()