from fastapi import FastAPI, UploadFile, File, Form, Request
from fastapi.responses import FileResponse
from fastapi.templating import Jinja2Templates
import pandas as pd
import os
import subprocess
from datetime import datetime
from generator import generate_from_word
from pypdf import PdfWriter
from docx2pdf import convert
import pythoncom
import locale
import io

# Konfigurasi Microsoft Edge untuk PDF
EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

def convert_html_to_pdf_edge(html_path, pdf_path):
    """Menggunakan Microsoft Edge untuk convert HTML ke PDF."""
    if not os.path.exists(EDGE_PATH):
        raise Exception(f"Microsoft Edge tidak ditemukan di {EDGE_PATH}")
    
    # Gunakan absolute path untuk file HTML
    abs_html_path = os.path.abspath(html_path)
    # Gunakan format file:/// untuk memastikan Edge bisa membuka file lokal
    file_url = f"file:///{abs_html_path.replace(os.sep, '/')}"
    
    cmd = [
        EDGE_PATH,
        "--headless",
        "--disable-gpu",
        "--no-sandbox",
        "--run-all-compositor-stages-before-draw",
        "--no-pdf-header-footer",      # Menghilangkan teks tanggal/link di pinggir PDF
        "--virtual-time-budget=2000",  # Memberi waktu 2 detik agar gambar/font selesai render
        f"--print-to-pdf={os.path.abspath(pdf_path)}",
        file_url
    ]
    print(f"Menjalankan perintah: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode != 0:
        print(f"Edge Error Output: {result.stderr}")
        raise Exception(f"Edge gagal: {result.stderr}")
    
    if not os.path.exists(pdf_path):
        print(f"Edge selesai tapi file PDF tidak ditemukan di: {pdf_path}")
        raise Exception("Edge gagal membuat file PDF.")

def format_tanggal_indonesia(tgl_val):
    """Mengubah format tanggal ke DD BULAN YYYY (Indonesian)."""
    if not tgl_val or str(tgl_val).strip() == "":
        return ""
    try:
        # Jika input adalah datetime object
        if isinstance(tgl_val, datetime):
            return tgl_val.strftime("%d %B %Y").upper()
        
        # Jika string, coba parse (menangani format timestamp excel juga)
        tgl_str = str(tgl_val).split(' ')[0] # Ambil bagian tanggal saja jika ada jam
        formats = ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"]
        for fmt in formats:
            try:
                dt = datetime.strptime(tgl_str, fmt)
                return dt.strftime("%d %B %Y").upper()
            except: continue
        return str(tgl_val).upper()
    except:
        return str(tgl_val).upper()

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
templates = Jinja2Templates(directory=["file", "."])

TEMPLATE_DIR = "file"

@app.get("/")
@app.get("/index.html")
async def read_index():
    return FileResponse("index.html")

@app.get("/templates")
async def list_templates():
    # File Word untuk SKHU
    word_files = [f for f in os.listdir(TEMPLATE_DIR) if f.endswith('.docx')]
    # File HTML untuk Sertifikat
    html_files = [f for f in os.listdir(TEMPLATE_DIR) if f.endswith('.html')]
    return {"word": word_files, "html": html_files}

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
    
    for index, row in df.iterrows():
        # Konversi baris ke dict dan bersihkan spasi
        data = {str(k): str(v).strip() for k, v in row.to_dict().items()}
        
        # Format Tanggal Lahir jika ada
        if 'tanggal_lahir' in data:
            data['tanggal_lahir'] = format_tanggal_indonesia(data['tanggal_lahir'])
        
        # Add tanggal to data
        data['tanggal_ujian'] = tanggal_ujian_str
        data['tanggal_pengesahan'] = tanggal_pengesahan_str
        
        # Scores calculation
        # Mendukung variasi nama kolom (dasar vs gerak_dasar, kombinasi vs gerak_kombinasi)
        score_mapping = {
            'nilai_teori': ['nilai_teori', 'teori'],
            'nilai_salam': ['nilai_salam', 'salam'],
            'nilai_gerak_dasar': ['nilai_gerak_dasar', 'nilai_dasar', 'dasar'],
            'nilai_gerak_kombinasi': ['nilai_gerak_kombinasi', 'nilai_kombinasi', 'kombinasi'],
            'nilai_jurus': ['nilai_jurus', 'jurus'],
            'nilai_jatuhan': ['nilai_jatuhan', 'jatuhan'],
            'nilai_bantingan': ['nilai_bantingan', 'bantingan'],
            'nilai_serang_bela': ['nilai_serang_bela', 'serang_bela', 'sab'],
            'nilai_senjata': ['nilai_senjata', 'senjata'],
            'nilai_fisik': ['nilai_fisik', 'fisik']
        }
        
        available_scores = []
        for target_col, aliases in score_mapping.items():
            val = ""
            for alias in aliases:
                if data.get(alias):
                    val = data[alias]
                    # Pastikan data juga tersedia di target_col agar template Word bisa baca
                    data[target_col] = val
                    break
            
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
        except Exception as e:
            print(f"Gagal mengonversi {name}: {e}")
            
    # Simpan hasil gabungan PDF
    # Format: Nama Template (Capitalized) - Tanggal Ujian.pdf
    base_name = os.path.splitext(template_name)[0]
    formatted_name = base_name.replace('-', ' ').title()
    final_filename = f"{formatted_name} - {tanggal_ujian_str}.pdf"
    final_pdf_path = os.path.join(OUTPUT_DIR, final_filename)
    merger.write(final_pdf_path)
    merger.close()
            
    response = FileResponse(final_pdf_path, media_type='application/pdf', filename=final_filename)
    response.headers["X-Filename"] = final_filename
    return response

@app.post("/generate-sertifikat")
async def generate_sertifikat(
    request: Request,
    template_name: str = Form(...),
    jenis_ujian: str = Form(...),
    tempat_ujian: str = Form(...),
    tanggal_ujian: str = Form(...),
    tanggal_pengesahan: str = Form(...) ,
    data_file: UploadFile = File(...)
):
    try:
        data_path = os.path.join(UPLOAD_DIR, data_file.filename)
        with open(data_path, "wb") as buffer:
            buffer.write(await data_file.read())
        
        # Format tanggal (Indonesian)
        try:
            tgl_ujian_obj = datetime.strptime(tanggal_ujian, "%Y-%m-%d")
            tanggal_ujian_str = tgl_ujian_obj.strftime("%d %B %Y")
            
            tgl_pengesahan_obj = datetime.strptime(tanggal_pengesahan, "%Y-%m-%d")
            tanggal_pengesahan_str = tgl_pengesahan_obj.strftime("%d %B %Y")
        except:
            tanggal_ujian_str = tanggal_ujian
            tanggal_pengesahan_str = tanggal_pengesahan

        # Process data
        if data_file.filename.endswith('.csv'):
            df = pd.read_csv(data_path, encoding='utf-8-sig')
        else:
            df = pd.read_excel(data_path)
        
        # Bersihkan data: ganti NaN dengan string kosong, bersihkan nama kolom
        df = df.fillna('')
        df.columns = [str(c).strip().lower() for c in df.columns]
        print(f"Kolom yang dideteksi di Excel: {list(df.columns)}")
        
        merger = PdfWriter()
        temp_pdfs = []
        
        # Render dan konversi tiap baris
        for index, row in df.iterrows():
            try:
                # Ambil semua kolom dari baris excel sebagai dictionary
                row_data = {str(k): str(v).strip() for k, v in row.to_dict().items()}
                
                # Pastikan field utama ada walaupun nama kolom di excel sedikit berbeda
                # (Misal: 'Nama ' atau 'NAMA' atau 'nama')
                nama_val = row_data.get('nama', row_data.get('nama ', ''))
                tempat_lahir_val = row_data.get('tempat_lahir', row_data.get('tempat lahir', ''))
                tanggal_lahir_val = row_data.get('tanggal_lahir', row_data.get('tanggal lahir', ''))

                # Format Tanggal Lahir jika ada
                if tanggal_lahir_val:
                    tanggal_lahir_val = format_tanggal_indonesia(tanggal_lahir_val)
                
                # Siapkan data context (gabungkan data form + data excel)
                context = {
                    **row_data,  # Pindahkan ke atas agar tidak menimpa variabel di bawahnya
                    "request": request,
                    "jenis_ujian": jenis_ujian,
                    "tempat_ujian": tempat_ujian,
                    "tanggal_ujian": tanggal_ujian_str,
                    "tanggal_pengesahan": tanggal_pengesahan_str,
                    "nama": nama_val,
                    "tempat_lahir": tempat_lahir_val,
                    "tanggal_lahir": tanggal_lahir_val
                }
                
                # Render HTML menggunakan Jinja2
                template_obj = templates.get_template(template_name)
                html_content = template_obj.render(context)
                
                # Simpan HTML sementara di root agar path gambar (images/...) valid
                temp_html_path = f"temp_render_{index}.html"
                with open(temp_html_path, "w", encoding="utf-8") as f:
                    f.write(html_content)
                
                # Konversi ke PDF menggunakan Edge
                pdf_path = os.path.join(OUTPUT_DIR, f"sertifikat_{index}.pdf")
                convert_html_to_pdf_edge(temp_html_path, pdf_path)
                
                if os.path.exists(pdf_path):
                    merger.append(pdf_path)
                    temp_pdfs.append(pdf_path)
                
                # Hapus HTML sementara
                if os.path.exists(temp_html_path):
                    os.remove(temp_html_path)
            except Exception as line_error:
                print(f"Error pada baris {index}: {str(line_error)}")
                continue
                
        if not temp_pdfs:
            raise Exception("Tidak ada PDF yang berhasil dibuat. Periksa log server.")

        final_filename = "Sertifikat_Gabungan.pdf"
        final_path = os.path.join(OUTPUT_DIR, final_filename)
        merger.write(final_path)
        merger.close()
        
        # Hapus PDF individual
        for p in temp_pdfs:
            if os.path.exists(p):
                os.remove(p)
        
        response = FileResponse(final_path, media_type='application/pdf', filename=final_filename)
        response.headers["X-Filename"] = final_filename
        return response
    except Exception as e:
        print(f"CRITICAL ERROR: {str(e)}")
        from fastapi import HTTPException
        raise HTTPException(status_code=500, detail=str(e))
