from fastapi import FastAPI, UploadFile, HTTPException
import fitz  # PyMuPDF
import gc

app = FastAPI(
    title="PDF Text Extraction API",
    description="API สำหรับสกัดข้อความออกจากไฟล์ PDF เพื่อนำไปใช้ทำ Vector DB / RAG",
    version="1.0.0"
)

@app.get("/")
def root():
    return {"message": "PDF Extraction Service is running perfectly!"}

@app.get("/health")
def health_check():
    return {"status": "running"}

@app.post("/extract")
async def extract_pdf(file: UploadFile):
    # ตรวจสอบว่าเป็นไฟล์ PDF หรือไม่
    if not file.filename.endswith('.pdf'):
        raise HTTPException(status_code=400, detail="ไฟล์ที่อัปโหลดต้องเป็นนามสกุล .pdf เท่านั้น")

    try:
        # อ่านไฟล์เป็น Bytes
        pdf_bytes = await file.read()
        
        # เปิดไฟล์ PDF ผ่าน PyMuPDF ในรูปแบบ Stream
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")

        results = []
        for page_num, page in enumerate(doc):
            # ดึงข้อความจากแต่ละหน้า
            text = page.get_text()
            
            results.append({
                "page": page_num + 1,
                "text": text,
                "images": []  # ส่ง list ว่างเพื่อป้องกันปัญหา Memory/Base64 หลุดเข้า Vector DB
            })

        # ปิดไฟล์เพื่อคืน Memory
        doc.close()
        
        # เคลียร์ตัวแปรและเรียก Garbage Collector เพื่อประหยัด RAM บน Render Free Tier
        del pdf_bytes
        gc.collect()

        return {
            "status": "ok",
            "filename": file.filename,
            "total_pages": len(results),
            "pages": results
        }

    except Exception as e:
        # ล้าง RAM หากเกิด Error
        gc.collect()
        raise HTTPException(status_code=500, detail=f"เกิดข้อผิดพลาดในการประมวลผล PDF: {str(e)}")