THEMES = {
    "futuristic": {
        "bg":          "#0a0a0f",
        "header_bg":   "#0f0f1a",
        "card_bg":     "#13131f",
        "title_color": "#ffffff",
        "body_color":  "#c8cfe0",
        "muted_color": "#7a8299",
        "accent":      "#ff7820",
        "font_head":   "Trebuchet MS",
        "font_body":   "Calibri",
    },
    "neon": {
        "bg":          "#060610",
        "header_bg":   "#0c0c20",
        "card_bg":     "#10102a",
        "title_color": "#ffffff",
        "body_color":  "#d0d8f0",
        "muted_color": "#7080aa",
        "accent":      "#c084fc",
        "font_head":   "Calibri",
        "font_body":   "Calibri Light",
    },
    "minimal": {
        "bg":          "#fafafa",
        "header_bg":   "#f0f0f0",
        "card_bg":     "#ffffff",
        "title_color": "#111111",
        "body_color":  "#333333",
        "muted_color": "#888888",
        "accent":      "#ff7820",
        "font_head":   "Calibri",
        "font_body":   "Calibri",
    },
}

def generate_theme(style: str = "futuristic") -> dict:
    return THEMES.get(style, THEMES["futuristic"])