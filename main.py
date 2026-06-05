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
    # Support CSV or Excel
    if data_file.filename.endswith('.csv'):
        df = pd.read_csv(data_path)
    else:
        df = pd.read_excel(data_path)
    
    is_word = template.filename.endswith('.docx')
    generated_files = []
    
    for index, row in df.iterrows():
        data = row.to_dict()
        
        # Calculate Average for Graduation Letter if columns exist
        score_columns = [
            'nilai_teori', 'nilai', 'nilai_dasar', 
            'nilai_kombinasi', 'nilai_jurus', 'nilai_jatuhan', 
            'nilai_bantingan', 'nilai_serang_bela', 'nilai_senjata', 
            'nilai_fisik'
        ]
        
        # Check if we have these columns to calculate average
        available_scores = [data[col] for col in score_columns if col in data]
        if available_scores:
            data['rata_rata_nilai'] = sum(available_scores) / len(available_scores)
        
        # Add creation date if not present
        if 'tanggal_pembuatan' not in data or pd.isna(data['tanggal_pembuatan']):
            data['tanggal_pembuatan'] = datetime.now().strftime("%d %B %Y")

        name = data.get('nama', data.get('name', f"doc_{index}"))
        
        if is_word:
            output_filename = f"Surat_{name}.docx"
            output_path = os.path.join(OUTPUT_DIR, output_filename)
            generate_from_word(template_path, output_path, data)
        else:
            # Fallback to image-based certificate
            output_filename = f"cert_{name}.pdf"
            output_path = os.path.join(OUTPUT_DIR, output_filename)
            # For simplicity, we use name_x/y if provided, else defaults
            coords = {"nama": (name_x or 100, name_y or 100)}
            generate_certificate(
                template_path, 
                output_path, 
                {"nama": name}, 
                font_path or "arial.ttf", 
                font_size, 
                coords
            )
            
        generated_files.append(output_path)
    
    # Zip files
    zip_filename = "hasil_generate.zip"
    zip_path = os.path.join(OUTPUT_DIR, zip_filename)
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for file in generated_files:
            zipf.write(file, os.path.basename(file))
            
    return FileResponse(zip_path, media_type='application/zip', filename=zip_filename)
