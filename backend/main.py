from fastapi import FastAPI, HTTPException
from fastapi.responses import Response
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import traceback

app = FastAPI(title="AI PPT Generator")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

class PPTRequest(BaseModel):
    topic: str
    slides: int = 10   # default 10
    style: str = "futuristic"

@app.post("/generate-ppt")
async def generate_ppt(req: PPTRequest):
    try:
        from agent.planner import generate_outline
        from agent.designer import generate_theme
        from agent.ppt_builder import build_ppt

        print(f"[1/3] Generating outline for '{req.topic}' ({req.slides} slides)...")
        outline = generate_outline(req.topic, req.slides)
        print(f"[2/3] Building theme '{req.style}'...")
        theme = generate_theme(req.style)
        print(f"[3/3] Building PPT...")
        ppt_bytes = build_ppt(outline, theme, req.style)
        print("Done!")

        filename = f"{req.topic.replace(' ', '_')}_presentation.pptx"
        return Response(
            content=ppt_bytes,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        print(traceback.format_exc())
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/health")
def health():
    return {"status": "running"}