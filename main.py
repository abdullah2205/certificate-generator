from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
import pandas as pd
import os
import zipfile
from datetime import datetime
from generator import generate_certificate, generate_from_word

app = FastAPI()

@app.get("/")
async def read_index():
    return FileResponse("index.html")

# Temporary directories
UPLOAD_DIR = "uploads"
OUTPUT_DIR = "output"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

@app.post("/generate")
async def generate_documents(
    template: UploadFile = File(...),
    data_file: UploadFile = File(...),
    # Optional fields for Image-based certificates
    font_path: str = Form(None),
    font_size: int = Form(30),
    name_x: int = Form(None),
    name_y: int = Form(None)
):
    # Save files
    template_path = os.path.join(UPLOAD_DIR, template.filename)
    data_path = os.path.join(UPLOAD_DIR, data_file.filename)
    
    with open(template_path, "wb") as buffer:
        buffer.write(await template.read())
    with open(data_path, "wb") as buffer:
        buffer.write(await data_file.read())
    
    # Process data
    if data_file.filename.endswith('.csv'):
        # Using engine='python' can be more robust for some CSV variations
        df = pd.read_csv(data_path)
    else:
        df = pd.read_excel(data_path)
    
    # FIX 1: Hapus semua tulisan 'nan' dengan menggantinya jadi string kosong
    df = df.fillna('')
    
    is_word = template.filename.endswith('.docx')
    generated_files = []
    
    for index, row in df.iterrows():
        # Convert row to dictionary and strip whitespace from keys and values
        data = {str(k).strip(): (str(v).strip() if pd.notna(v) else "") for k, v in row.to_dict().items()}
        
        # FIX 2: Sinkronkan nama kolom nilai dengan header nilai.csv Anda
        score_columns = [
            'nilai_teori', 'nilai_salam', 'nilai_dasar', 
            'nilai_kombinasi', 'nilai_jurus', 'nilai_jatuhan', 
            'nilai_bantingan', 'nilai_serang_bela', 'nilai_senjata', 
            'nilai_fisik'
        ]
        
        # Hitung rata-rata
        available_scores = []
        for col in score_columns:
            val = data.get(col, "")
            if val != "":
                try:
                    # Ganti koma ke titik dulu jika inputnya pakai koma (misal: 85,5)
                    clean_val = str(val).replace(',', '.')
                    available_scores.append(float(clean_val))
                except (ValueError, TypeError):
                    pass
                    
        if available_scores:
            avg = sum(available_scores) / len(available_scores)
            # Hasil rata-rata tetap pakai format Indonesia (koma)
            data['rata_rata_nilai'] = f"{avg:.1f}".replace('.', ',')
        else:
            data['rata_rata_nilai'] = '0,0'
        
        # Tanggal otomatis
        if not data.get('tanggal_pembuatan'):
            data['tanggal_pembuatan'] = datetime.now().strftime("%d %B %Y")

        # FIX 3: Gabungkan tempat dan tanggal lahir otomatis
        tl = data.get('tempat_lahir', '')
        tgl = data.get('tanggal_lahir', '')
        if tl and tgl:
            data['tempat_tanggal_lahir'] = f"{tl}, {tgl}"
        else:
            data['tempat_tanggal_lahir'] = tl or tgl

        # Nama file output
        name = data.get('nama', f"document_{index}")
        
        if is_word:
            output_filename = f"Surat_{name}.docx"
            output_path = os.path.join(OUTPUT_DIR, output_filename)
            generate_from_word(template_path, output_path, data)
        else:
            output_filename = f"cert_{name}.pdf"
            output_path = os.path.join(OUTPUT_DIR, output_filename)
            coords = {"nama": (name_x or 100, name_y or 100)}
            # Map 'nama' for image generation if needed
            cert_data = {"nama": name}
            generate_certificate(template_path, output_path, cert_data, font_path or "arial.ttf", font_size, coords)
            
        generated_files.append(output_path)
    
    # Zip files
    zip_filename = "hasil_generate.zip"
    zip_path = os.path.join(OUTPUT_DIR, zip_filename)
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for file in generated_files:
            zipf.write(file, os.path.basename(file))
            
    return FileResponse(zip_path, media_type='application/zip', filename=zip_filename)
