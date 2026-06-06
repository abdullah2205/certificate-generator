from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
import pandas as pd
import os
from datetime import datetime
from generator import generate_from_word
from pypdf import PdfWriter
from docx2pdf import convert
import pythoncom
import re

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
    data_file: UploadFile = File(...)
):
    # Initialize COM for docx2pdf if on Windows
    pythoncom.CoInitialize()
    
    # Save files
    template_path = os.path.join(UPLOAD_DIR, template.filename)
    data_path = os.path.join(UPLOAD_DIR, data_file.filename)
    
    with open(template_path, "wb") as buffer:
        buffer.write(await template.read())
    with open(data_path, "wb") as buffer:
        buffer.write(await data_file.read())
    
    # Process data
    if data_file.filename.endswith('.csv'):
        df = pd.read_csv(data_path, encoding='utf-8-sig')
    else:
        df = pd.read_excel(data_path)
    
    df = df.fillna('')
    df.columns = [str(c).strip() for c in df.columns]
    
    merger = PdfWriter()
    
    for index, row in df.iterrows():
        data = {str(k).strip(): str(v).strip() for k, v in row.to_dict().items()}
        
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
            data['rata_rata_nilai'] = f"{avg:.1f}".replace('.', ',')
        else:
            data['rata_rata_nilai'] = '0,0'
        
        # Dates and names
        if not data.get('tanggal_pembuatan'):
            data['tanggal_pembuatan'] = datetime.now().strftime("%d %B %Y")

        tl = data.get('tempat_lahir', '')
        tgl = data.get('tanggal_lahir', '')
        if tl and tgl:
            data['tempat_tanggal_lahir'] = f"{tl}, {tgl}"
        else:
            data['tempat_tanggal_lahir'] = tl or tgl

        name = data.get('nama', f"doc_{index}")
        
        docx_temp = os.path.join(OUTPUT_DIR, f"temp_{index}.docx")
        pdf_temp = os.path.join(OUTPUT_DIR, f"temp_{index}.pdf")
        
        generate_from_word(template_path, docx_temp, data)
        try:
            convert(docx_temp, pdf_temp)
            merger.append(pdf_temp)
        except Exception as e:
            print(f"Error converting {name}: {e}")
            
    final_pdf_path = os.path.join(OUTPUT_DIR, "SKHU_LENGKAP.pdf")
    merger.write(final_pdf_path)
    merger.close()
            
    return FileResponse(final_pdf_path, media_type='application/pdf', filename='SKHU_LENGKAP.pdf')
