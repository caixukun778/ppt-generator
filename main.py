import os
import uuid
import json
import shutil
import time
import httpx
from fastapi import FastAPI, HTTPException, Request, UploadFile, File as FastAPIFile
from fastapi.responses import FileResponse
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

app = FastAPI()
OUTPUT_DIR = "outputs"
os.makedirs(OUTPUT_DIR, exist_ok=True)

COLOR_MAP = {
    "深蓝色": RGBColor(0, 51, 102),
    "蓝色": RGBColor(0, 102, 204),
    "科技蓝": RGBColor(0, 153, 255),
}

def get_color(name: str) -> RGBColor:
    for key, value in COLOR_MAP.items():
        if key in name:
            return value
    return RGBColor(0, 51, 102)

def build_pptx(title: str, style_desc: str, slides_json_str: str):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    primary_color = get_color(style_desc)

    slides_data = []
    try:
        slides_data = json.loads(slides_json_str)
        if not isinstance(slides_data, list):
            slides_data = [{"title": slides_json_str}]
    except:
        slides_data = [{"title": "内容", "blocks": [{"heading": "", "bullets": [slides_json_str]}]}]

    for slide_data in slides_data:
        slide_type = slide_data.get("type", "content")
        slide_layout = prs.slide_layouts[6] if slide_type == "cover" else prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)

        if slide_type == "cover":
            bg = slide.background
            fill = bg.fill
            fill.solid()
            fill.fore_color.rgb = primary_color
            txBox = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(11), Inches(1.5))
            p = txBox.text_frame.paragraphs[0]
            p.text = slide_data.get("title", title)
            p.font.size = Pt(44)
            p.font.color.rgb = RGBColor(255, 255, 255)
            p.font.bold = True
            p.alignment = PP_ALIGN.CENTER
            sub = slide_data.get("subtitle", "")
            if sub:
                txBox2 = slide.shapes.add_textbox(Inches(1), Inches(5), Inches(11), Inches(1))
                p2 = txBox2.text_frame.paragraphs[0]
                p2.text = sub
                p2.font.size = Pt(24)
                p2.font.color.rgb = RGBColor(220, 220, 220)
                p2.alignment = PP_ALIGN.CENTER
        else:
            slide.shapes.title.text = slide_data.get("title", "内容")
            body = slide.shapes.placeholders[1].text_frame
            body.clear()
            blocks = slide_data.get("blocks", [])
            items = slide_data.get("items", [])
            if blocks:
                for block in blocks:
                    h = block.get("heading", "")
                    if h:
                        p = body.add_paragraph()
                        p.text = h
                        p.font.size = Pt(20)
                        p.font.bold = True
                        p.font.color.rgb = primary_color
                        p.space_after = Pt(8)
                    for bullet in block.get("bullets", []):
                        p = body.add_paragraph()
                        p.text = f"• {bullet}"
                        p.font.size = Pt(16)
                        p.space_after = Pt(4)
            elif items:
                for item in items:
                    p = body.add_paragraph()
                    p.text = f"• {item}"
                    p.font.size = Pt(16)
                    p.space_after = Pt(4)

    filename = f"{uuid.uuid4()}.pptx"
    filepath = os.path.join(OUTPUT_DIR, filename)
    prs.save(filepath)
    return filepath, filename

@app.post("/generate_pptx")
async def generate_pptx(request: Request):
    try:
        data = await request.json()
        title = data.get("title", "PPT")
        style_desc = data.get("style", "")
        slides_str = data.get("slides", "[]")
        
        filepath, filename = build_pptx(title, style_desc, slides_str)
        download_url = f"/download/{filename}"
        return {
            "success": True,
            "download_url": download_url,
            "file_name": filename
        }
    except Exception as e:
        return {
            "success": False,
            "error": str(e)
        }

@app.get("/download/{filename}")
async def download_file(filename: str):
    filepath = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(filepath):
        raise HTTPException(status_code=404, detail="File not found or expired")
    return FileResponse(filepath, filename=filename)
    
# ========== 素材库 ==========
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

IMAGES_DB = "images_db.json"
if not os.path.exists(IMAGES_DB):
    with open(IMAGES_DB, "w", encoding="utf-8") as f:
        json.dump([], f)


import httpx  # 需要新增这个导入，用于下载链接

@app.post("/upload_image")
async def upload_image(request: Request, tag: str = ""):
    """上传图片，支持直接传文件或传图片链接"""
    try:
        form = await request.form()
        file = form.get("file")
        
        if file and hasattr(file, "filename"):
            ext = file.filename.split(".")[-1] if "." in file.filename else "png"
            save_name = f"{uuid.uuid4()}.{ext}"
            save_path = os.path.join(UPLOAD_DIR, save_name)
            with open(save_path, "wb") as f:
                f.write(await file.read())
        else:
            data = await request.json()
            image_url = data.get("file") or data.get("image_url")
            if not image_url:
                return {"success": False, "error": "未提供图片文件或链接"}
            async with httpx.AsyncClient() as client:
                response = await client.get(image_url)
                if response.status_code != 200:
                    return {"success": False, "error": f"下载图片失败，状态码: {response.status_code}"}
                content_type = response.headers.get("content-type", "")
                ext = "png"
                if "jpeg" in content_type or "jpg" in content_type:
                    ext = "jpg"
                elif "webp" in content_type:
                    ext = "webp"
                save_name = f"{uuid.uuid4()}.{ext}"
                save_path = os.path.join(UPLOAD_DIR, save_name)
                with open(save_path, "wb") as f:
                    f.write(response.content)
        
        with open(IMAGES_DB, "r", encoding="utf-8") as f:
            images = json.load(f)
        
        images.append({
            "id": save_name,
            "tag": tag,
            "url": f"/uploads/{save_name}",
            "uploaded_at": time.strftime("%Y-%m-%d %H:%M:%S")
        })
        
        with open(IMAGES_DB, "w", encoding="utf-8") as f:
            json.dump(images, f, ensure_ascii=False, indent=2)
        
        return {"success": True, "image_url": f"/uploads/{save_name}", "id": save_name}
    except Exception as e:
        return {"success": False, "error": str(e)}
        
@app.get("/list_images")
async def list_images(tag: str = ""):
    """查看图片列表，可按标签筛选"""
    try:
        with open(IMAGES_DB, "r", encoding="utf-8") as f:
            images = json.load(f)
        
        if tag:
            images = [img for img in images if img.get("tag") == tag]
        
        for img in images:
            img["full_url"] = f"https://ppt-generator-production-a9fd.up.railway.app{img['url']}"
        
        return {"success": True, "images": images, "count": len(images)}
    except Exception as e:
        return {"success": False, "error": str(e)}
@app.get("/list_images")
async def list_images(tag: str = ""):
    """查看图片列表，可按标签筛选"""
    try:
        with open(IMAGES_DB, "r", encoding="utf-8") as f:
            images = json.load(f)
        
        if tag:
            images = [img for img in images if img.get("tag") == tag]
        
        for img in images:
            img["full_url"] = f"https://ppt-generator-production-a9fd.up.railway.app{img['url']}"
        
        return {"success": True, "images": images, "count": len(images)}
    except Exception as e:
        return {"success": False, "error": str(e)}
@app.get("/")
def root():
    return {"status": "running"}
