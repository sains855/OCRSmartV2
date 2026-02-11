import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from google import genai 
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image
from dotenv import load_dotenv
import threading
import re

# --- LOAD KONFIGURASI ---
load_dotenv()
api_key = os.getenv("GEMINI_API_KEY")
client = genai.Client(api_key=api_key) if api_key else None

class GeminiOCRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AI Form Digitizer - Pro Layout v4")
        self.root.geometry("500x400")
        self.root.configure(bg="#ffffff")
        
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure("Action.TButton", background="#2b579a", foreground="white", font=("Segoe UI", 10, "bold"))

        self.file_path = ""
        self.container = tk.Frame(root, bg="#ffffff", padx=40, pady=40)
        self.container.pack(expand=True, fill="both")

        tk.Label(self.container, text="Form OCR Precision v4", font=("Segoe UI", 16, "bold"), bg="#ffffff").pack(pady=(0, 20))
        
        ttk.Button(self.container, text="1. Pilih Foto Formulir", command=self.select_file).pack(fill="x", pady=5)
        self.btn_process = ttk.Button(self.container, text="2. Ekstrak & Tampilkan Tabel", style="Action.TButton", command=self.start_processing)
        self.btn_process.pack(fill="x", pady=20)

        self.status_var = tk.StringVar(value="Ready")
        tk.Label(self.container, textvariable=self.status_var, font=("Segoe UI", 9), bg="#ffffff", fg="#666666").pack()

    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Images", "*.jpg *.jpeg *.png")])
        if self.file_path:
            self.status_var.set(f"File: {os.path.basename(self.file_path)}")

    def start_processing(self):
        if not self.file_path: return
        self.status_var.set("Membangun tabel dan grid...")
        threading.Thread(target=self.process_ocr, daemon=True).start()

    def format_run(self, paragraph, text, is_bold=False, size=10):
        """Helper untuk memastikan font setiap baris konsisten di dalam tabel"""
        run = paragraph.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(size)
        run.bold = is_bold
        return run

    def process_ocr(self):
        try:
            img = Image.open(self.file_path)
            
            prompt = (
                "Lakukan OCR pada formulir ini. Berikan output dengan format:\n"
                "[HEADER]\n"
                "(Teks header instansi)\n"
                "[BODY]\n"
                "(Format 'Label: Isi'. Untuk kotak/checkbox, gunakan simbol [ ] atau [X] jika terisi. "
                "PENTING: JANGAN sertakan titik-titik pengisi ....)\n"
                "Abaikan teks penjelasan AI."
            )
            
            # Menggunakan model terbaru gemini-2.0-flash
            response = client.models.generate_content(model="gemini-2.5-flash", contents=[prompt, img])
            raw_text = response.text.strip()

            doc = Document()
            # Set default font dokumen
            doc.styles['Normal'].font.name = 'Arial'
            doc.styles['Normal'].font.size = Pt(10)

            lines = raw_text.split('\n')
            mode = "HEADER"
            table = None

            for line in lines:
                line = line.strip()
                if not line: continue
                
                if "[HEADER]" in line:
                    mode = "HEADER"
                    continue
                if "[BODY]" in line:
                    mode = "BODY"
                    # --- FIX: MENAMPILKAN GARIS TABEL ---
                    table = doc.add_table(rows=0, cols=2)
                    table.style = 'Table Grid' # Membuat garis tabel terlihat 
                    table.autofit = False
                    table.columns[0].width = Inches(2.2)
                    table.columns[1].width = Inches(3.8)
                    continue

                if mode == "HEADER":
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    is_big = any(x in line.upper() for x in ["OMBUDSMAN", "FORMULIR"])
                    self.format_run(p, line, is_bold=is_big, size=11 if is_big else 9)

                elif mode == "BODY" and table:
                    if ":" in line:
                        parts = line.split(":", 1)
                        label = parts[0].strip()
                        value = re.sub(r'\.{2,}', '', parts[1]).strip()

                        if label.isupper() and not value:
                            row = table.add_row().cells
                            cell = row[0].merge(row[1])
                            self.format_run(cell.paragraphs[0], label, is_bold=True, size=10)
                        else:
                            row_cells = table.add_row().cells
                            self.format_run(row_cells[0].paragraphs[0], label, is_bold=True)
                            
                            # Mengonversi checkbox teks ke simbol kotak agar rapi
                            clean_val = value.replace("[ ]", "□").replace("[X]", "▣")
                            self.format_run(row_cells[1].paragraphs[0], ": " + (clean_val if clean_val else "-"))
                    else:
                        row = table.add_row().cells
                        self.format_run(row[0].merge(row[1]).paragraphs[0], line)

            self.root.after(0, self.save_document, doc)

        except Exception as e:
            self.root.after(0, lambda err=e: messagebox.showerror("Error", str(err)))
        finally:
            self.root.after(0, self.reset_ui)

    def save_document(self, doc):
        save_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word", "*.docx")])
        if save_path:
            doc.save(save_path)
            messagebox.showinfo("Sukses", "Tabel dan data berhasil dibuat dengan garis yang terlihat!")
        self.status_var.set("Ready")

    def reset_ui(self):
        self.btn_process.state(['!disabled'])

if __name__ == "__main__":
    root = tk.Tk()
    app = GeminiOCRApp(root)
    root.mainloop()