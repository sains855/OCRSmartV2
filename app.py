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
        self.root.title("AI Form Digitizer - Auto Header")
        self.root.geometry("500x400")
        self.root.configure(bg="#ffffff")
        
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure("Action.TButton", background="#2b579a", foreground="white", font=("Segoe UI", 10, "bold"))

        self.file_path = ""
        self.container = tk.Frame(root, bg="#ffffff", padx=40, pady=40)
        self.container.pack(expand=True, fill="both")

        tk.Label(self.container, text="Form OCR Precision", font=("Segoe UI", 16, "bold"), bg="#ffffff").pack(pady=(0, 20))
        
        ttk.Button(self.container, text="1. Pilih Foto Formulir", command=self.select_file).pack(fill="x", pady=5)
        self.btn_process = ttk.Button(self.container, text="2. Ekstrak & Rapikan", style="Action.TButton", command=self.start_processing)
        self.btn_process.pack(fill="x", pady=20)

        self.status_var = tk.StringVar(value="Ready")
        tk.Label(self.container, textvariable=self.status_var, font=("Segoe UI", 9), bg="#ffffff", fg="#666666").pack()

    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Images", "*.jpg *.jpeg *.png")])
        if self.file_path:
            self.status_var.set(f"File: {os.path.basename(self.file_path)}")

    def start_processing(self):
        if not self.file_path: return
        self.status_var.set("AI sedang menganalisis tata letak...")
        threading.Thread(target=self.process_ocr, daemon=True).start()

    def process_ocr(self):
        try:
            img = Image.open(self.file_path)
            
            # Prompt ditingkatkan untuk memisahkan Header dan Body secara eksplisit
            prompt = (
                "Lakukan OCR pada formulir ini. Berikan output dengan format berikut:\n"
                "[HEADER]\n"
                "(Tulis semua teks header instansi/logo teks, alamat, telp di sini secara berurutan)\n"
                "[BODY]\n"
                "(Tulis semua isian formulir dengan format 'Label: Isi'. Hapus titik-titik pengisi ....)\n"
                "PENTING: Jangan berikan komentar pembuka atau penutup."
            )
            
            response = client.models.generate_content(model="gemini-2.5-flash", contents=[prompt, img])
            raw_text = response.text.strip()

            doc = Document()
            style = doc.styles['Normal']
            style.font.name = 'Arial'
            style.font.size = Pt(10)

            # --- PARSING LOGIC ---
            lines = raw_text.split('\n')
            mode = "HEADER" # Default mode awal
            
            # Buat tabel untuk bagian BODY nanti
            table = None

            for line in lines:
                line = line.strip()
                if not line: continue
                
                # Cek perubahan mode
                if "[HEADER]" in line:
                    mode = "HEADER"
                    continue
                if "[BODY]" in line:
                    mode = "BODY"
                    # Inisialisasi tabel saat masuk ke mode body
                    table = doc.add_table(rows=0, cols=2)
                    table.autofit = False
                    table.columns[0].width = Inches(2.2)
                    table.columns[1].width = Inches(3.8)
                    continue

                if mode == "HEADER":
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run(line)
                    # Tebalkan jika baris pertama (biasanya nama instansi)
                    if "OMBUDSMAN" in line.upper() or "REPUBLIK" in line.upper() or "FORMULIR" in line.upper():
                        run.bold = True
                        run.font.size = Pt(11)

                elif mode == "BODY" and table:
                    if ":" in line:
                        parts = line.split(":", 1)
                        label = parts[0].strip()
                        value = re.sub(r'\.{2,}', '', parts[1]).strip()

                        # Deteksi judul section (seperti IDENTITAS PELAPOR)
                        if label.isupper() and not value:
                            row = table.add_row().cells
                            cell = row[0].merge(row[1])
                            p = cell.paragraphs[0]
                            p.add_run(label).bold = True
                            p.add_run().underline = True
                        else:
                            row_cells = table.add_row().cells
                            row_cells[0].text = label
                            row_cells[0].paragraphs[0].runs[0].bold = True
                            row_cells[1].text = ": " + (value if value else "-")
                    else:
                        # Jika teks biasa tanpa titik dua di dalam body
                        row = table.add_row().cells
                        row[0].merge(row[1]).text = line

            self.root.after(0, self.save_document, doc)

        except Exception as e:
            self.root.after(0, lambda err=e: messagebox.showerror("Error", str(err)))
        finally:
            self.root.after(0, self.reset_ui)

    def save_document(self, doc):
        save_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word", "*.docx")])
        if save_path:
            doc.save(save_path)
            messagebox.showinfo("Sukses", "Dokumen otomatis berhasil dibuat!")
        self.status_var.set("Ready")

    def reset_ui(self):
        self.btn_process.state(['!disabled'])

if __name__ == "__main__":
    root = tk.Tk()
    app = GeminiOCRApp(root)
    root.mainloop()