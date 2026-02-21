from fastapi import FastAPI, HTTPException
from fastapi.responses import Response, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import traceback, sys, os

# Add backend directory to path so ppt_builder can find anim_engine
sys.path.insert(0, os.path.dirname(__file__))

app = FastAPI(title="AI PPT Generator")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

TEMPLATES_INFO = {
    "futuristic":  {"name":"Futuristic",  "desc":"Dark cyber-tech, orange accents, glowing circles, grid lines"},
    "minimalist":  {"name":"Minimalist",  "desc":"White space, fine hairlines, restrained type-forward layout"},
    "artistic":    {"name":"Artistic",    "desc":"Overlapping geometry, editorial composition, expressive shapes"},
    "misty":       {"name":"Misty Peaks", "desc":"Layered mountain silhouettes, atmospheric depth, nature palette"},
    "monochrome":  {"name":"Monochrome",  "desc":"Single hue + stark contrast, heavy block type, bold structure"},
    "hardware":    {"name":"Hardware",    "desc":"Circuit blueprint grid, dark industrial, monospace font"},
    "magazine":    {"name":"Magazine",    "desc":"Bold editorial, full-bleed asymmetry, oversized impact type"},
    "neon":        {"name":"Neon Nights", "desc":"Outlined neon shapes, synthwave glow, pulse & flip effects"},
}

THEMES = {
    1:  {"name":"Midnight Orange","bg":"#0a0a0f","card_bg":"#13131f","header_bg":"#0f0f1a","title_color":"#ffffff","body_color":"#c8cfe0","muted_color":"#7a8299","accent":"#ff7820","accent2":"#ff4500","font_head":"Trebuchet MS","font_body":"Calibri"},
    2:  {"name":"Cyber Cyan","bg":"#050d1a","card_bg":"#0a1628","header_bg":"#071220","title_color":"#e0f4ff","body_color":"#a8d8ea","muted_color":"#5a8fa8","accent":"#00d4ff","accent2":"#0099cc","font_head":"Calibri","font_body":"Calibri Light"},
    3:  {"name":"Neon Green","bg":"#060f06","card_bg":"#0d1f0d","header_bg":"#0a180a","title_color":"#e8ffe8","body_color":"#a8d4a8","muted_color":"#5a8a5a","accent":"#39ff14","accent2":"#00cc00","font_head":"Trebuchet MS","font_body":"Calibri"},
    4:  {"name":"Electric Purple","bg":"#0a0515","card_bg":"#130a20","header_bg":"#0f0818","title_color":"#f0e8ff","body_color":"#c8a8f0","muted_color":"#8060b0","accent":"#bf5fff","accent2":"#8800ff","font_head":"Calibri","font_body":"Calibri Light"},
    5:  {"name":"Clean White","bg":"#fafafa","card_bg":"#ffffff","header_bg":"#f0f0f0","title_color":"#111111","body_color":"#333333","muted_color":"#888888","accent":"#ff7820","accent2":"#e05510","font_head":"Trebuchet MS","font_body":"Calibri"},
    6:  {"name":"Warm Paper","bg":"#fdf8f0","card_bg":"#ffffff","header_bg":"#f5ede0","title_color":"#2c1810","body_color":"#5a3a28","muted_color":"#a07860","accent":"#c8500a","accent2":"#a03800","font_head":"Calibri","font_body":"Calibri Light"},
    7:  {"name":"Executive Navy","bg":"#1e2761","card_bg":"#253070","header_bg":"#1a2258","title_color":"#ffffff","body_color":"#cadcfc","muted_color":"#8099cc","accent":"#4da6ff","accent2":"#cadcfc","font_head":"Calibri","font_body":"Calibri Light"},
    8:  {"name":"Black Gold","bg":"#0a0800","card_bg":"#151000","header_bg":"#100c00","title_color":"#ffd700","body_color":"#c8a800","muted_color":"#886e00","accent":"#ffd700","accent2":"#c8a000","font_head":"Calibri","font_body":"Calibri Light"},
    9:  {"name":"Sepia Classic","bg":"#2c1a00","card_bg":"#3d2a00","header_bg":"#352200","title_color":"#f5deb3","body_color":"#d4af70","muted_color":"#a07840","accent":"#cd853f","accent2":"#8b6914","font_head":"Calibri","font_body":"Calibri Light"},
    10: {"name":"Synthwave","bg":"#0d0221","card_bg":"#1a0435","header_bg":"#14032a","title_color":"#ff7eee","body_color":"#df73ff","muted_color":"#9940cc","accent":"#08f7fe","accent2":"#09fbd3","font_head":"Trebuchet MS","font_body":"Calibri Light"},
    11: {"name":"Terminal Green","bg":"#000a00","card_bg":"#001400","header_bg":"#000f00","title_color":"#00ff00","body_color":"#00cc00","muted_color":"#008800","accent":"#00ff44","accent2":"#00cc33","font_head":"Courier New","font_body":"Courier New"},
    12: {"name":"Misty Mountain","bg":"#c0c8d8","card_bg":"#d0d8e8","header_bg":"#b8c0d0","title_color":"#1a2030","body_color":"#3a4050","muted_color":"#7080a0","accent":"#3060a0","accent2":"#204880","font_head":"Calibri","font_body":"Calibri Light"},
    13: {"name":"Obsidian","bg":"#080808","card_bg":"#111111","header_bg":"#0d0d0d","title_color":"#ffffff","body_color":"#dddddd","muted_color":"#888888","accent":"#ff8c00","accent2":"#dd5500","font_head":"Trebuchet MS","font_body":"Calibri"},
    14: {"name":"Neon Tokyo","bg":"#05001a","card_bg":"#0a0030","header_bg":"#070025","title_color":"#ff00aa","body_color":"#cc0088","muted_color":"#880055","accent":"#00eeff","accent2":"#00ccdd","font_head":"Trebuchet MS","font_body":"Calibri Light"},
    15: {"name":"Rose Gold","bg":"#1a0a0f","card_bg":"#2a1018","header_bg":"#220d14","title_color":"#ffddcc","body_color":"#e8b090","muted_color":"#c07050","accent":"#e8927c","accent2":"#c87060","font_head":"Calibri","font_body":"Calibri Light"},
}

class PPTRequest(BaseModel):
    topic: str
    slides: int = 10
    theme_id: int = 1
    template: str = "futuristic"

class PreviewRequest(BaseModel):
    topic: str
    slides: int = 10
    theme_id: int = 1
    template: str = "futuristic"

@app.get("/templates")
def list_templates():
    return [{"key": k, **v} for k, v in TEMPLATES_INFO.items()]

@app.get("/themes")
def list_themes():
    return [{"id": k, "name": v["name"], "bg": v["bg"],
             "accent": v["accent"], "accent2": v["accent2"]} for k, v in THEMES.items()]

@app.post("/preview")
async def preview_ppt(req: PreviewRequest):
    try:
        from agent.planner import generate_outline
        outline = generate_outline(req.topic, req.slides)
        theme = THEMES.get(req.theme_id, THEMES[1])
        return JSONResponse({
            "outline": outline,
            "theme": theme,
            "template": req.template,
        })
    except Exception as e:
        print(traceback.format_exc())
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/generate-ppt")
async def generate_ppt(req: PPTRequest):
    try:
        from agent.planner import generate_outline
        from agent.ppt_builder import build_ppt
        print(f"[1/3] Generating outline '{req.topic}' ({req.slides} slides, template={req.template})...")
        outline = generate_outline(req.topic, req.slides)
        theme = THEMES.get(req.theme_id, THEMES[1])
        print(f"[2/3] Theme: {theme['name']}")
        print(f"[3/3] Building {req.template} template...")
        ppt_bytes = build_ppt(outline, theme, req.template)
        print("Done!")
        fname = f"{req.topic.replace(' ','_')}_{req.template}.pptx"
        return Response(
            content=ppt_bytes,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": f"attachment; filename={fname}"}
        )
    except Exception as e:
        print(traceback.format_exc())
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/health")
def health(): return {"status": "running"}