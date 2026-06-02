from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
import pandas as pd
import os
import zipfile
from generator import generate_certificate

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
async def generate_certificates(
    template: UploadFile = File(...),
    data_file: UploadFile = File(...),
    font_path: str = Form(...),
    font_size: int = Form(30),
    # Simple form fields for coordinates (could be expanded)
    name_x: int = Form(100),
    name_y: int = Form(100)
):
    # Save files
    template_path = os.path.join(UPLOAD_DIR, template.filename)
    data_path = os.path.join(UPLOAD_DIR, data_file.filename)
    
    with open(template_path, "wb") as buffer:
        buffer.write(await template.read())
    with open(data_path, "wb") as buffer:
        buffer.write(await data_file.read())
    
    # Process data
    df = pd.read_csv(data_path) # Assuming CSV for now
    
    generated_files = []
    for index, row in df.iterrows():
        name = row['name']
        output_filename = f"cert_{name}.pdf"
        output_path = os.path.join(OUTPUT_DIR, output_filename)
        
        generate_certificate(
            template_path, 
            output_path, 
            {"name": name}, 
            font_path, 
            font_size, 
            {"name": (name_x, name_y)}
        )
        generated_files.append(output_path)
    
    # Zip files
    zip_path = "certificates.zip"
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for file in generated_files:
            zipf.write(file, os.path.basename(file))
            
    return FileResponse(zip_path, media_type='application/zip', filename='certificates.zip')
