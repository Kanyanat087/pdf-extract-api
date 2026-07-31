from fastapi import FastAPI, UploadFile, HTTPException
import fitz  # PyMuPDF
import docx  # python-docx
from pptx import Presentation  # python-pptx
import pandas as pd  # pandas
import gc
import uuid
import io

app = FastAPI(
    title="Multi-Format Document Extraction API",
    description="API สำหรับสกัดข้อความออกจากไฟล์ PDF, DOCX, PPTX, XLSX, TXT เพื่อนำไปใช้ทำ Vector DB / RAG",
    version="2.0.0"
)

@app.get("/")
def root():
    return {"message": "Document Extraction Service is running perfectly!"}

@app.get("/health")
def health_check():
    return {"status": "running"}

@app.post("/extract")
async def extract_file(file: UploadFile):
    try:
        content = await file.read()
        results = []
        filename = file.filename.lower()

        # 1. กรณีเป็นไฟล์ PDF (ดึงข้อความ + สกัดรูป)
        if filename.endswith(".pdf"):
            doc = fitz.open(stream=content, filetype="pdf")
            for page_num, page in enumerate(doc):
                text = page.get_text()

                image_list = []
                for img_index, img in enumerate(page.get_images(full=True)):
                    xref = img[0]
                    base_image = doc.extract_image(xref)
                    image_ext = base_image["ext"]

                    image_filename = f"page_{page_num+1}_{uuid.uuid4().hex[:8]}.{image_ext}"
                    image_url = f"http://127.0.0.1:8000/static/extracted_images/{image_filename}"
                    image_list.append(image_url)

                results.append({
                    "page": page_num + 1,
                    "text": text,
                    "images": image_list  # ส่งกลับเป็น List/Array
                })
            doc.close()

        # 2. กรณีเป็นไฟล์ข้อความ (.txt, .csv, .json, .md)
        elif filename.endswith((".txt", ".md", ".json", ".log", ".csv")):
            text = content.decode("utf-8", errors="ignore")
            results.append({
                "page": 1,
                "text": text,
                "images": []
            })

        # 3. กรณีเป็นไฟล์ Word (.docx)
        elif filename.endswith(".docx"):
            doc = docx.Document(io.BytesIO(content))
            full_text = [p.text for p in doc.paragraphs if p.text.strip()]
            results.append({
                "page": 1,
                "text": "\n".join(full_text),
                "images": []
            })

        # 4. กรณีเป็นไฟล์ PowerPoint (.pptx)
        elif filename.endswith(".pptx"):
            prs = Presentation(io.BytesIO(content))
            for idx, slide in enumerate(prs.slides):
                slide_text = [shape.text for shape in slide.shapes if hasattr(shape, "text") and shape.text]
                results.append({
                    "page": idx + 1,
                    "text": "\n".join(slide_text),
                    "images": []
                })

        # 5. กรณีเป็นไฟล์ Excel (.xlsx, .xls)
        elif filename.endswith((".xlsx", ".xls")):
            excel_file = pd.ExcelFile(io.BytesIO(content))
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                text = f"--- Sheet: {sheet_name} ---\n" + df.to_csv(index=False)
                results.append({
                    "page": sheet_name,
                    "text": text,
                    "images": []
                })

        else:
            raise HTTPException(status_code=400, detail=f"ไม่รองรับไฟล์ประเภทนี้: {file.filename}")

        del content
        gc.collect()

        return {
            "status": "ok",
            "filename": file.filename,
            "total_pages": len(results),
            "pages": results
        }

    except HTTPException as he:
        raise he
    except Exception as e:
        gc.collect()
        raise HTTPException(status_code=500, detail=f"เกิดข้อผิดพลาดในการประมวลผลไฟล์: {str(e)}")