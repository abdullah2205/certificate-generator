from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
import pandas as pd
import os
from datetime import datetime
from generator import generate_from_word
from pypdf import PdfWriter
from docx2pdf import convert
import pythoncom
import locale

# Set locale to Indonesian
try:
    locale.setlocale(locale.LC_TIME, 'id_ID.UTF-8')
except locale.Error:
    try:
        # Windows fallback
        locale.setlocale(locale.LC_TIME, 'Indonesian_Indonesia.1252')
    except locale.Error:
        print("Warning: Could not set Indonesian locale. Dates may be in English.")

app = FastAPI()

TEMPLATE_DIR = "file"

@app.get("/")
async def read_index():
    return FileResponse("index.html")

@app.get("/templates")
async def list_templates():
    files = [f for f in os.listdir(TEMPLATE_DIR) if f.endswith('.docx')]
    return files

# Temporary directories
UPLOAD_DIR = "uploads"
OUTPUT_DIR = "output"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

@app.post("/generate")
async def generate_documents(
    template_name: str = Form(...),
    tanggal_ujian: str = Form(...),
    tanggal_pengesahan: str = Form(...),
    data_file: UploadFile = File(...)
):
    # Initialize COM for docx2pdf if on Windows
    pythoncom.CoInitialize()
    
    # Format tanggal
    tgl_ujian_obj = datetime.strptime(tanggal_ujian, "%Y-%m-%d")
    tanggal_ujian_str = tgl_ujian_obj.strftime("%d %B %Y")
    
    tgl_pengesahan_obj = datetime.strptime(tanggal_pengesahan, "%Y-%m-%d")
    tanggal_pengesahan_str = tgl_pengesahan_obj.strftime("%d %B %Y")
    
    template_path = os.path.join(TEMPLATE_DIR, template_name)
    data_path = os.path.join(UPLOAD_DIR, data_file.filename)
    
    with open(data_path, "wb") as buffer:
        buffer.write(await data_file.read())
    
    # Process data
    if data_file.filename.endswith('.csv'):
        df = pd.read_csv(data_path, encoding='utf-8-sig')
    else:
        df = pd.read_excel(data_path)
    
    # Bersihkan data: ganti NaN dengan string kosong, bersihkan nama kolom
    df = df.fillna('')
    df.columns = [str(c).strip().lower() for c in df.columns]
    
    merger = PdfWriter()
    temp_files = []
    
    for index, row in df.iterrows():
        # Konversi baris ke dict dan bersihkan spasi
        data = {str(k): str(v).strip() for k, v in row.to_dict().items()}
        
        # Add tanggal to data
        data['tanggal_ujian'] = tanggal_ujian_str
        data['tanggal_pengesahan'] = tanggal_pengesahan_str
        
        # Scores calculation
        score_columns = [
            'nilai_teori', 'nilai_salam', 'nilai_dasar', 
            'nilai_kombinasi', 'nilai_jurus', 'nilai_jatuhan', 
            'nilai_bantingan', 'nilai_serang_bela', 'nilai_senjata', 
            'nilai_fisik'
        ]
        
        available_scores = []
        for col in score_columns:
            val = data.get(col, "")
            if val != "":
                try:
                    available_scores.append(float(str(val).replace(',', '.')))
                except: pass
                    
        if available_scores:
            avg = sum(available_scores) / len(available_scores)
            # Hasil rata-rata menggunakan koma (1 digit)
            data['rata_rata_nilai'] = f"{avg:.1f}".replace('.', ',')
        else:
            data['rata_rata_nilai'] = '0,0'
        
        # Dates and names
        if not data.get('tanggal_pembuatan'):
            data['tanggal_pembuatan'] = tanggal_pengesahan_str

        # Penggabungan tempat dan tanggal lahir
        tl = data.get('tempat_lahir', '')
        tgl = data.get('tanggal_lahir', '')
        if tl and tgl:
            data['tempat_tanggal_lahir'] = f"{tl}, {tgl}"
        else:
            data['tempat_tanggal_lahir'] = tl or tgl

        name = data.get('nama', f"doc_{index}")
        
        # Nama file sementara
        docx_temp = os.path.join(OUTPUT_DIR, f"temp_{index}.docx")
        pdf_temp = os.path.join(OUTPUT_DIR, f"temp_{index}.pdf")
        
        # Generate Word
        generate_from_word(template_path, docx_temp, data)
        
        # Konversi Word ke PDF dan gabungkan
        try:
            convert(docx_temp, pdf_temp)
            merger.append(pdf_temp)
            temp_files.extend([docx_temp, pdf_temp])
        except Exception as e:
            print(f"Gagal mengonversi {name}: {e}")
            
    # Simpan hasil gabungan PDF
    # Format: Nama Template (Capitalized) - Tanggal Pengesahan.pdf
    base_name = os.path.splitext(template_name)[0]
    formatted_name = base_name.replace('-', ' ').title()
    final_filename = f"{formatted_name} - {tanggal_pengesahan_str}.pdf"
    final_pdf_path = os.path.join(OUTPUT_DIR, final_filename)
    merger.write(final_pdf_path)
    merger.close()
            
    # Use standard response with filename
    from fastapi.responses import FileResponse
    return FileResponse(final_pdf_path, media_type='application/pdf', filename=final_filename)
