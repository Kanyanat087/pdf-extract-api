from fastapi import FastAPI, UploadFile, HTTPException, Form
from fastapi.staticfiles import StaticFiles
import fitz  # PyMuPDF
import docx  # python-docx
from pptx import Presentation  # python-pptx
import pandas as pd  # pandas
import gc
import uuid
import io
import os
import hashlib

# ---------------------------------------------------------------
# ตั้งค่าโฟลเดอร์เก็บรูป
# ---------------------------------------------------------------
IMAGE_DIR = "static/extracted_images"
os.makedirs(IMAGE_DIR, exist_ok=True)

# หมายเหตุ: API คืนค่าเป็น "path" อย่างเดียว ไม่มี domain
#   เช่น /static/extracted_images/page_1_abc.jpeg
# เพราะ IP ของเครื่องอาจเปลี่ยนได้ ถ้าเก็บ IP ลงฐานข้อมูลแล้ว IP เปลี่ยน
# จะต้อง ingest ข้อมูลใหม่ทั้งหมด
# ให้ n8n เป็นคนเติม domain เองตอนจะใช้งานจริง

# ข้ามรูปที่เล็กกว่านี้ (โลโก้ ไอคอน เส้นคั่น) หน่วยเป็น byte
MIN_IMAGE_SIZE = 15000

app = FastAPI(
    title="Multi-Format Document Extraction API",
    description="API สำหรับสกัดข้อความและรูปออกจากไฟล์ PDF, DOCX, PPTX, XLSX, TXT เพื่อนำไปใช้ทำ Vector DB / RAG",
    version="3.1.0"
)

# เปิดให้เข้าถึงรูปผ่าน URL ได้
app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/")
def root():
    return {"message": "Document Extraction Service is running perfectly!"}


@app.get("/health")
def health_check():
    return {"status": "running"}


def save_page_images(doc, page, page_num):
    """
    ดึงรูปทั้งหมดในหน้านั้น เซฟลงดิสก์จริง แล้วคืน list ของ path
    """
    image_list = []

    for img in page.get_images(full=True):
        xref = img[0]

        try:
            base_image = doc.extract_image(xref)
        except Exception:
            continue

        image_bytes = base_image["image"]
        image_ext = base_image["ext"]

        # ข้ามรูปเล็ก ๆ ที่ไม่ใช่เนื้อหาจริง
        if len(image_bytes) < MIN_IMAGE_SIZE:
            continue

        image_filename = f"page_{page_num}_{uuid.uuid4().hex[:8]}.{image_ext}"
        image_path = os.path.join(IMAGE_DIR, image_filename)

        # เขียนไฟล์ลงดิสก์จริง
        with open(image_path, "wb") as f:
            f.write(image_bytes)

        # เก็บเฉพาะ path ไม่มี domain
        image_list.append(f"/static/extracted_images/{image_filename}")

    return image_list


@app.post("/extract")
async def extract_file(
    file: UploadFile,
    system: str = Form("unknown"),      # เช่น ACCA, e-Payment
    doc_type: str = Form("manual"),     # manual หรือ faq
):
    try:
        content = await file.read()
        file_hash = hashlib.sha256(content).hexdigest()
        results = []
        filename = file.filename.lower()

        # 1. PDF (ดึงข้อความ + สกัดรูป)
        if filename.endswith(".pdf"):
            doc = fitz.open(stream=content, filetype="pdf")

            for page_index, page in enumerate(doc):
                page_num = page_index + 1
                raw_text = page.get_text()

                # ใส่หัวเรื่องนำหน้าเนื้อหา เพื่อให้ embedding มีบริบท
                # (embedding มองเห็นแค่ text ไม่เห็น metadata)
                text = f"[{system} - {doc_type} - หน้า {page_num}]\n{raw_text}"

                image_list = save_page_images(doc, page, page_num)

                results.append({
                    "page": page_num,
                    "text": text,
                    "images": image_list,
                })

            doc.close()

        # 2. ไฟล์ข้อความ (.txt, .md, .json, .log, .csv)
        elif filename.endswith((".txt", ".md", ".json", ".log", ".csv")):
            raw_text = content.decode("utf-8", errors="ignore")
            results.append({
                "page": 1,
                "text": f"[{system} - {doc_type}]\n{raw_text}",
                "images": [],
            })

        # 3. Word (.docx)
        elif filename.endswith(".docx"):
            doc = docx.Document(io.BytesIO(content))
            full_text = [p.text for p in doc.paragraphs if p.text.strip()]
            results.append({
                "page": 1,
                "text": f"[{system} - {doc_type}]\n" + "\n".join(full_text),
                "images": [],
            })

        # 4. PowerPoint (.pptx)
        elif filename.endswith(".pptx"):
            prs = Presentation(io.BytesIO(content))
            for idx, slide in enumerate(prs.slides):
                slide_text = [
                    shape.text for shape in slide.shapes
                    if hasattr(shape, "text") and shape.text
                ]
                results.append({
                    "page": idx + 1,
                    "text": f"[{system} - {doc_type} - สไลด์ {idx + 1}]\n" + "\n".join(slide_text),
                    "images": [],
                })

        # 5. Excel (.xlsx, .xls)
        elif filename.endswith((".xlsx", ".xls")):
            excel_file = pd.ExcelFile(io.BytesIO(content))
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                text = (
                    f"[{system} - {doc_type} - Sheet: {sheet_name}]\n"
                    + df.to_csv(index=False)
                )
                results.append({
                    "page": sheet_name,
                    "text": text,
                    "images": [],
                })

        else:
            raise HTTPException(
                status_code=400,
                detail=f"ไม่รองรับไฟล์ประเภทนี้: {file.filename}"
            )

        del content
        gc.collect()

        return {
            "status": "ok",
            "filename": file.filename,
            "file_hash": file_hash,     # ใช้เช็กไฟล์ซ้ำได้
            "system": system,
            "doc_type": doc_type,
            "total_pages": len(results),
            "pages": results,
        }

    except HTTPException as he:
        raise he
    except Exception as e:
        gc.collect()
        raise HTTPException(
            status_code=500,
            detail=f"เกิดข้อผิดพลาดในการประมวลผลไฟล์: {str(e)}"
        )