import os
import json
from groq import Groq
from dotenv import load_dotenv

load_dotenv()
client = Groq(api_key=os.getenv("GROQ_API_KEY"))

def generate_outline(topic: str, slides: int = 10) -> list:
    prompt = f"""
You are an expert presentation designer and subject matter expert.
Create a detailed, informative {slides}-slide presentation about: "{topic}".

Return ONLY a valid JSON array with exactly {slides} objects.
No markdown code blocks, no explanation — just the raw JSON array starting with [ and ending with ].

Each object must have ALL of these keys:
- "title": compelling slide title (max 8 words)
- "subtitle": one punchy line summarizing the slide (max 15 words)
- "points": list of exactly 4 bullet points. Each bullet MUST be a complete sentence of 20-30 words that fully explains the concept with specific details, facts, or examples.
- "detail": a paragraph of 50-70 words expanding on the slide topic with real facts, statistics, or concrete examples. This should be educational and informative.
- "notes": speaker notes of 60-80 words with talking points, extra context, and suggestions for the presenter.
- "layout": choose the BEST layout from this list based on content type:
    "title_hero"   - for introduction or conclusion slides
    "two_column"   - for explanatory content with deep detail
    "icon_grid"    - for listing 4 key concepts or features
    "stat_callout" - for numbered steps or ranked items
    "timeline"     - for sequential steps, history, or process
    "full_detail"  - for complex topics needing full explanation

Rules:
- NEVER use the same layout twice in a row
- Use "title_hero" for slide 1 and the last slide
- Vary layouts across all {slides} slides
- Make content expert-level, educational, and deeply informative
- Include real-world examples, statistics, or named technologies where relevant
- Each bullet point must be substantive — no vague filler phrases

Return ONLY the JSON array. No markdown, no backticks, no explanation.
"""

    response = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.7,
        max_tokens=6000,
    )
    raw = response.choices[0].message.content.strip()

    # Strip markdown fences
    if "```" in raw:
        for part in raw.split("```"):
            part = part.strip()
            if part.startswith("json"):
                part = part[4:].strip()
            if part.startswith("["):
                raw = part
                break

    # Find JSON boundaries
    start = raw.find("[")
    end   = raw.rfind("]") + 1
    if start != -1 and end > start:
        raw = raw[start:end]

    data = json.loads(raw)

    # Ensure we have exactly the right number of slides
    # and all required keys exist
    required_keys = ["title", "subtitle", "points", "detail", "notes", "layout"]
    valid_layouts  = {"title_hero","two_column","icon_grid","stat_callout","timeline","full_detail"}
    cleaned = []
    for i, item in enumerate(data[:slides]):
        for key in required_keys:
            if key not in item:
                if key == "points":
                    item[key] = [f"Key concept about {item.get('title','this topic')}"] * 4
                elif key == "layout":
                    fallbacks = ["two_column","icon_grid","stat_callout","timeline","full_detail"]
                    item[key] = fallbacks[i % len(fallbacks)]
                else:
                    item[key] = ""
        if item["layout"] not in valid_layouts:
            item["layout"] = "two_column"
        # Ensure 4 points
        while len(item["points"]) < 4:
            item["points"].append(f"Additional insight about {item.get('title','this topic')}.")
        item["points"] = item["points"][:4]
        cleaned.append(item)

    return cleaned