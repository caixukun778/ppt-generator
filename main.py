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

# ========== 颜色映射 ==========
COLOR_MAP = {
    "深蓝色": RGBColor(0, 51, 102),
    "蓝色": RGBColor(0, 102, 204),
    "科技蓝": RGBColor(0, 153, 255),
    "灰色": RGBColor(128, 128, 128),
    "深灰色": RGBColor(64, 64, 64),
    "白色": RGBColor(255, 255, 255),
    "黑色": RGBColor(0, 0, 0),
}

def get_color(name: str) -> RGBColor:
    for key, value in COLOR_MAP.items():
        if key in name:
            return value
    return RGBColor(0, 51, 102)


# ========== PPT 生成核心 ==========
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
        return {"success": True, "download_url": download_url}
    except Exception as e:
        return {"success": False, "error": str(e)}


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

@app.post("/upload_image")
async def upload_image(request: Request):
    """上传图片：支持直接传文件 + 传图片链接，两者可同时进行"""
    saved = []
    tag = ""

    try:
        content_type = request.headers.get("content-type", "")

        # 1. 处理文件上传（multipart/form-data）
        if "multipart" in content_type:
            form = await request.form()
            tag = form.get("tag", "")
            file = form.get("file")
            if file and hasattr(file, "filename"):
                ext = file.filename.split(".")[-1] if "." in file.filename else "png"
                save_name = f"{uuid.uuid4()}.{ext}"
                save_path = os.path.join(UPLOAD_DIR, save_name)
                with open(save_path, "wb") as f:
                    f.write(await file.read())
                saved.append({"name": save_name, "url": f"/uploads/{save_name}"})

        # 2. 处理 JSON 链接上传（可以单独，也可以和文件流同时存在）
        # 如果content-type是application/json，或者form里包含image_url
        if "json" in content_type:
            data = await request.json()
            tag = data.get("tag", tag)
            image_url = data.get("image_url", "")
            if image_url:
                async with httpx.AsyncClient() as client:
                    resp = await client.get(image_url)
                    if resp.status_code == 200:
                        ct = resp.headers.get("content-type", "")
                        ext = "png"
                        if "jpeg" in ct or "jpg" in ct:
                            ext = "jpg"
                        elif "webp" in ct:
                            ext = "webp"
                        save_name = f"{uuid.uuid4()}.{ext}"
                        save_path = os.path.join(UPLOAD_DIR, save_name)
                        with open(save_path, "wb") as f:
                            f.write(resp.content)
                        saved.append({"name": save_name, "url": f"/uploads/{save_name}"})
        else:
            # 如果multipart里也传了image_url（非标准但可能发生）
            form = await request.form()
            tag = tag or form.get("tag", "")
            image_url = form.get("image_url", "")
            if image_url:
                async with httpx.AsyncClient() as client:
                    resp = await client.get(str(image_url))
                    if resp.status_code == 200:
                        ct = resp.headers.get("content-type", "")
                        ext = "png"
                        if "jpeg" in ct or "jpg" in ct:
                            ext = "jpg"
                        elif "webp" in ct:
                            ext = "webp"
                        save_name = f"{uuid.uuid4()}.{ext}"
                        save_path = os.path.join(UPLOAD_DIR, save_name)
                        with open(save_path, "wb") as f:
                            f.write(resp.content)
                        saved.append({"name": save_name, "url": f"/uploads/{save_name}"})

        if not saved:
            return {"success": False, "error": "未提供图片文件或链接"}

        # 3. 保存记录
        with open(IMAGES_DB, "r", encoding="utf-8") as f:
            images = json.load(f)

        for s in saved:
            images.append({
                "id": s["name"],
                "tag": tag,
                "url": s["url"],
                "uploaded_at": time.strftime("%Y-%m-%d %H:%M:%S")
            })

        with open(IMAGES_DB, "w", encoding="utf-8") as f:
            json.dump(images, f, ensure_ascii=False, indent=2)

        return {
            "success": True,
            "saved": len(saved),
            "images": [{"id": s["name"], "url": s["url"]} for s in saved]
        }

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


@app.get("/uploads/{filename}")
async def get_uploaded_file(filename: str):
    """访问上传的图片"""
    filepath = os.path.join(UPLOAD_DIR, filename)
    if not os.path.exists(filepath):
        raise HTTPException(status_code=404, detail="Image not found")
    return FileResponse(filepath)


@app.get("/cleanup")
def cleanup():
    """重置图片数据库"""
    try:
        with open(IMAGES_DB, "w", encoding="utf-8") as f:
            json.dump([], f)
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


@app.get("/")
def root():
    return {"status": "running"}
