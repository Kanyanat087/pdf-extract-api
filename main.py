from fastapi import FastAPI, UploadFile
from fastapi.staticfiles import StaticFiles
import fitz  # PyMuPDF
import uuid
import os

app = FastAPI()

# 1. สร้างโฟลเดอร์สำหรับเก็บรูปภาพที่สกัดมาได้
IMAGE_DIR = "static/extracted_images"
os.makedirs(IMAGE_DIR, exist_ok=True)

# 2. เปิดให้เข้าถึงรูปภาพผ่าน URL ได้ (เช่น http://localhost:8000/static/extracted_images/xxx.png)
app.mount("/static", StaticFiles(directory="static"), name="static")

@app.post("/extract")
async def extract_pdf(file: UploadFile):
    pdf_bytes = await file.read()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    results = []
    for page_num, page in enumerate(doc):
        text = page.get_text()
        image_urls = []
        
        for img in page.get_images(full=True):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image["ext"]
            
            # ตั้งชื่อไฟล์รูปภาพด้วย UUID เพื่อไม่ให้ชื่อซ้ำกัน
            filename = f"page_{page_num + 1}_{uuid.uuid4().hex[:8]}.{image_ext}"
            filepath = os.path.join(IMAGE_DIR, filename)
            
            # บันทึกไฟล์รูปลงดิสก์
            with open(filepath, "wb") as f:
                f.write(image_bytes)
            
            # สร้าง URL ของรูปภาพเพื่อส่งกลับไปให้ n8n เก็บลง metadata
            # (ถ้าเปลี่ยนไป deploy บน Render/Server ให้เปลี่ยน domain ตรงนี้)
            img_url = f"http://127.0.0.1:8000/static/extracted_images/{filename}"
            image_urls.append(img_url)

        results.append({
            "page": page_num + 1,
            "text": text,
            "images": image_urls  # ส่งกลับเฉพาะลิสต์ URL รูปภาพ สบายๆ คลีนๆ
        })

    doc.close()
    return {"status": "ok", "pages": results}

@app.get("/health")
def health_check():
    return {"status": "running"}