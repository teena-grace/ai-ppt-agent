from fastapi import FastAPI, HTTPException
from fastapi.responses import Response, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import traceback, base64, io

app = FastAPI(title="AI PPT Generator")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

class PPTRequest(BaseModel):
    topic: str
    slides: int = 10
    theme_id: int = 1  # from themes list
    style: str = "futuristic"

class PreviewRequest(BaseModel):
    topic: str
    slides: int = 10
    theme_id: int = 1
    style: str = "futuristic"

@app.get("/themes")
def list_themes():
    from themes import get_themes_list
    return get_themes_list()

@app.post("/preview")
async def preview_ppt(req: PreviewRequest):
    """Returns slide data as JSON for frontend preview rendering."""
    try:
        from agent.planner import generate_outline
        from themes import get_theme_by_id
        outline = generate_outline(req.topic, req.slides)
        theme = get_theme_by_id(req.theme_id)
        # Return slide data + theme for frontend to render preview cards
        return JSONResponse({
            "outline": outline,
            "theme": {
                "id": theme["id"],
                "name": theme["name"],
                "bg": theme["bg"],
                "accent": theme["accent"],
                "accent2": theme["accent2"],
                "title_color": theme["title_color"],
                "body_color": theme["body_color"],
                "card_bg": theme["card_bg"],
                "header_bg": theme["header_bg"],
                "muted_color": theme["muted_color"],
            }
        })
    except Exception as e:
        print(traceback.format_exc())
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/generate-ppt")
async def generate_ppt(req: PPTRequest):
    try:
        from agent.planner import generate_outline
        from agent.ppt_builder import build_ppt
        from themes import get_theme_by_id

        print(f"[1/3] Outline for '{req.topic}' ({req.slides} slides)...")
        outline = generate_outline(req.topic, req.slides)
        print(f"[2/3] Theme {req.theme_id}...")
        theme = get_theme_by_id(req.theme_id)
        print(f"[3/3] Building PPT...")
        ppt_bytes = build_ppt(outline, theme, req.style)
        print("Done!")
        filename = f"{req.topic.replace(' ','_')}_presentation.pptx"
        return Response(
            content=ppt_bytes,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        print(traceback.format_exc())
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/health")
def health(): return {"status": "running"}