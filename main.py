import os
import uuid
import json
import time
import threading
from fastapi import FastAPI, HTTPException, Request
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
    "灰色": RGBColor(128, 128, 128),
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

def delete_file_later(filepath: str, delay: int = 300):
    """延迟删除文件（默认5分钟）"""
    def _del():
        time.sleep(delay)
        if os.path.exists(filepath):
            os.remove(filepath)
    threading.Thread(target=_del, daemon=True).start()

@app.post("/generate_pptx")
async def generate_pptx(request: Request):
    try:
        data = await request.json()
        title = data.get("title", "PPT")
        style_desc = data.get("style", "")
        slides_str = data.get("slides", "[]")
        
        filepath, filename = build_pptx(title, style_desc, slides_str)
        # 5分钟后自动删除文件
        delete_file_later(filepath, delay=300)
        
        download_url = f"/download/{filename}"
        return {
            "success": True,
            "download_url": download_url
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

@app.get("/")
def root():
    return {"status": "running"}
