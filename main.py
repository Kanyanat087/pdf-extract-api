from fastapi import FastAPI, UploadFile
import fitz  # PyMuPDF
import base64

app = FastAPI()

@app.post("/extract")
async def extract_pdf(file: UploadFile):
    pdf_bytes = await file.read()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    results = []
    for page_num, page in enumerate(doc):
        text = page.get_text()
        images = []
        for img in page.get_images(full=True):
            xref = img[0]
            base_image = doc.extract_image(xref)
            img_b64 = base64.b64encode(base_image["image"]).decode()
            images.append({
                "data": img_b64,
                "ext": base_image["ext"]
            })
        results.append({
            "page": page_num,
            "text": text,
            "images": images
        })

    doc.close()
    return {"status": "ok", "pages": results}

@app.get("/health")
def health_check():
    return {"status": "running"}