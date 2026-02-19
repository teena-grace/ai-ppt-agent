"""
themes.py — 100+ unique PPT design themes.
Each theme defines: bg, card_bg, header_bg, title_color, body_color,
muted_color, accent, accent2, font_head, font_body, style_tag.
style_tag groups: dark, light, neon, pastel, corporate, retro, minimal, bold
"""

import random

THEMES = [
    # ── DARK FUTURISTIC ─────────────────────────────────────────────────────
    {"id":1,"name":"Midnight Orange","bg":"#0a0a0f","card_bg":"#13131f","header_bg":"#0f0f1a","title_color":"#ffffff","body_color":"#c8cfe0","muted_color":"#7a8299","accent":"#ff7820","accent2":"#ff4500","font_head":"Trebuchet MS","font_body":"Calibri","tag":"dark"},
    {"id":2,"name":"Cyber Cyan","bg":"#050d1a","card_bg":"#0a1628","header_bg":"#071220","title_color":"#e0f4ff","body_color":"#a8d8ea","muted_color":"#5a8fa8","accent":"#00d4ff","accent2":"#0099cc","font_head":"Calibri","font_body":"Calibri Light","tag":"dark"},
    {"id":3,"name":"Neon Green","bg":"#060f06","card_bg":"#0d1f0d","header_bg":"#0a180a","title_color":"#e8ffe8","body_color":"#a8d4a8","muted_color":"#5a8a5a","accent":"#39ff14","accent2":"#00cc00","font_head":"Trebuchet MS","font_body":"Calibri","tag":"neon"},
    {"id":4,"name":"Electric Purple","bg":"#0a0515","card_bg":"#130a20","header_bg":"#0f0818","title_color":"#f0e8ff","body_color":"#c8a8f0","muted_color":"#8060b0","accent":"#bf5fff","accent2":"#8800ff","font_head":"Calibri","font_body":"Calibri Light","tag":"neon"},
    {"id":5,"name":"Magenta Noir","bg":"#0f050f","card_bg":"#1a0a1a","header_bg":"#140814","title_color":"#ffe8ff","body_color":"#d4a8d4","muted_color":"#885088","accent":"#ff00aa","accent2":"#cc0077","font_head":"Trebuchet MS","font_body":"Calibri","tag":"neon"},
    {"id":6,"name":"Solar Flare","bg":"#0f0800","card_bg":"#1f1200","header_bg":"#180e00","title_color":"#fff8e8","body_color":"#f0d090","muted_color":"#a07830","accent":"#ffa500","accent2":"#ff6600","font_head":"Calibri","font_body":"Calibri","tag":"dark"},
    {"id":7,"name":"Arctic Blue","bg":"#040d1f","card_bg":"#091828","header_bg":"#061420","title_color":"#e8f4ff","body_color":"#a8c8f0","muted_color":"#5080b0","accent":"#4da8ff","accent2":"#0066cc","font_head":"Trebuchet MS","font_body":"Calibri Light","tag":"dark"},
    {"id":8,"name":"Blood Red","bg":"#0f0000","card_bg":"#1f0505","header_bg":"#180303","title_color":"#ffe8e8","body_color":"#f0a8a8","muted_color":"#a05050","accent":"#ff2222","accent2":"#cc0000","font_head":"Calibri","font_body":"Calibri","tag":"dark"},
    {"id":9,"name":"Teal Depths","bg":"#020f0f","card_bg":"#071e1e","header_bg":"#051818","title_color":"#e8ffff","body_color":"#a8e0e0","muted_color":"#508080","accent":"#00c8c8","accent2":"#009999","font_head":"Trebuchet MS","font_body":"Calibri","tag":"dark"},
    {"id":10,"name":"Gold Rush","bg":"#0a0800","card_bg":"#1a1400","header_bg":"#150f00","title_color":"#fffbe8","body_color":"#f0e090","muted_color":"#a09030","accent":"#ffd700","accent2":"#c8a800","font_head":"Calibri","font_body":"Calibri Light","tag":"dark"},

    # ── CORPORATE / PROFESSIONAL ─────────────────────────────────────────────
    {"id":11,"name":"Executive Navy","bg":"#1e2761","card_bg":"#253070","header_bg":"#1a2258","title_color":"#ffffff","body_color":"#cadcfc","muted_color":"#8099cc","accent":"#4da6ff","accent2":"#cadcfc","font_head":"Calibri","font_body":"Calibri Light","tag":"corporate"},
    {"id":12,"name":"Slate Pro","bg":"#2c3e50","card_bg":"#34495e","header_bg":"#2c3e50","title_color":"#ecf0f1","body_color":"#bdc3c7","muted_color":"#7f8c8d","accent":"#3498db","accent2":"#2980b9","font_head":"Calibri","font_body":"Calibri","tag":"corporate"},
    {"id":13,"name":"Forest Executive","bg":"#1a2f1a","card_bg":"#233023","header_bg":"#1e2a1e","title_color":"#e8f5e8","body_color":"#a8c8a8","muted_color":"#608060","accent":"#4caf50","accent2":"#388e3c","font_head":"Trebuchet MS","font_body":"Calibri","tag":"corporate"},
    {"id":14,"name":"Burgundy Suite","bg":"#2d0a0a","card_bg":"#3d1010","header_bg":"#350c0c","title_color":"#ffe8e8","body_color":"#d4a8a8","muted_color":"#885050","accent":"#c0392b","accent2":"#e74c3c","font_head":"Calibri","font_body":"Calibri Light","tag":"corporate"},
    {"id":15,"name":"Charcoal Minimal","bg":"#1a1a1a","card_bg":"#252525","header_bg":"#1f1f1f","title_color":"#ffffff","body_color":"#cccccc","muted_color":"#888888","accent":"#ff7820","accent2":"#e06010","font_head":"Trebuchet MS","font_body":"Calibri","tag":"corporate"},
    {"id":16,"name":"Ocean Depths","bg":"#003366","card_bg":"#004080","header_bg":"#002d5a","title_color":"#e8f4ff","body_color":"#99ccff","muted_color":"#5599cc","accent":"#00aaff","accent2":"#0088cc","font_head":"Calibri","font_body":"Calibri Light","tag":"corporate"},
    {"id":17,"name":"Emerald Trust","bg":"#013220","card_bg":"#024d30","header_bg":"#012a1a","title_color":"#e8fff0","body_color":"#90ee90","muted_color":"#408040","accent":"#00c853","accent2":"#009624","font_head":"Trebuchet MS","font_body":"Calibri","tag":"corporate"},
    {"id":18,"name":"Stone Cold","bg":"#36454f","card_bg":"#40505c","header_bg":"#2e3c45","title_color":"#f5f5f5","body_color":"#b0bec5","muted_color":"#78909c","accent":"#ff8f00","accent2":"#e65100","font_head":"Calibri","font_body":"Calibri","tag":"corporate"},
    {"id":19,"name":"Cobalt Authority","bg":"#002171","card_bg":"#0d47a1","header_bg":"#001a5c","title_color":"#e3f2fd","body_color":"#90caf9","muted_color":"#5090cc","accent":"#82b1ff","accent2":"#448aff","font_head":"Trebuchet MS","font_body":"Calibri Light","tag":"corporate"},
    {"id":20,"name":"Graphite Edge","bg":"#212121","card_bg":"#303030","header_bg":"#292929","title_color":"#fafafa","body_color":"#e0e0e0","muted_color":"#9e9e9e","accent":"#ff5722","accent2":"#e64a19","font_head":"Calibri","font_body":"Calibri","tag":"corporate"},

    # ── LIGHT / MINIMAL ──────────────────────────────────────────────────────
    {"id":21,"name":"Clean White","bg":"#fafafa","card_bg":"#ffffff","header_bg":"#f0f0f0","title_color":"#111111","body_color":"#333333","muted_color":"#888888","accent":"#ff7820","accent2":"#e05510","font_head":"Trebuchet MS","font_body":"Calibri","tag":"light"},
    {"id":22,"name":"Warm Paper","bg":"#fdf8f0","card_bg":"#ffffff","header_bg":"#f5ede0","title_color":"#2c1810","body_color":"#5a3a28","muted_color":"#a07860","accent":"#c8500a","accent2":"#a03800","font_head":"Calibri","font_body":"Calibri Light","tag":"light"},
    {"id":23,"name":"Mint Fresh","bg":"#f0fff4","card_bg":"#ffffff","header_bg":"#e0f7ea","title_color":"#1a3a2a","body_color":"#2d6a4f","muted_color":"#74a98c","accent":"#00b37e","accent2":"#00916e","font_head":"Trebuchet MS","font_body":"Calibri","tag":"light"},
    {"id":24,"name":"Sky Light","bg":"#f0f8ff","card_bg":"#ffffff","header_bg":"#e0f0ff","title_color":"#0a2a4a","body_color":"#1a4a7a","muted_color":"#5080aa","accent":"#0077cc","accent2":"#005599","font_head":"Calibri","font_body":"Calibri Light","tag":"light"},
    {"id":25,"name":"Blush Rose","bg":"#fff5f5","card_bg":"#ffffff","header_bg":"#ffe8e8","title_color":"#4a0a0a","body_color":"#7a2a2a","muted_color":"#aa6060","accent":"#e91e63","accent2":"#c2185b","font_head":"Trebuchet MS","font_body":"Calibri","tag":"light"},
    {"id":26,"name":"Lavender Mist","bg":"#f8f4ff","card_bg":"#ffffff","header_bg":"#f0e8ff","title_color":"#2a0a4a","body_color":"#5a2a8a","muted_color":"#9060c0","accent":"#7c4dff","accent2":"#651fff","font_head":"Calibri","font_body":"Calibri Light","tag":"light"},
    {"id":27,"name":"Sand Dune","bg":"#faf5e8","card_bg":"#ffffff","header_bg":"#f5edd5","title_color":"#3a2a10","body_color":"#6a4a20","muted_color":"#a07848","accent":"#c8860a","accent2":"#a06800","font_head":"Trebuchet MS","font_body":"Calibri","tag":"light"},
    {"id":28,"name":"Ice White","bg":"#f4f9ff","card_bg":"#ffffff","header_bg":"#eaf3ff","title_color":"#0a1a3a","body_color":"#2a4a7a","muted_color":"#6080aa","accent":"#1565c0","accent2":"#003c8f","font_head":"Calibri","font_body":"Calibri","tag":"light"},
    {"id":29,"name":"Parchment","bg":"#fdf5e6","card_bg":"#fff8f0","header_bg":"#f5e8d0","title_color":"#2a1a00","body_color":"#5a3a10","muted_color":"#906040","accent":"#8b4513","accent2":"#6b3010","font_head":"Trebuchet MS","font_body":"Calibri Light","tag":"light"},
    {"id":30,"name":"Arctic White","bg":"#f0f4f8","card_bg":"#ffffff","header_bg":"#e4ecf4","title_color":"#102030","body_color":"#304050","muted_color":"#708090","accent":"#007acc","accent2":"#005a9e","font_head":"Calibri","font_body":"Calibri","tag":"light"},

    # ── PASTEL / SOFT ────────────────────────────────────────────────────────
    {"id":31,"name":"Cotton Candy","bg":"#fff0f8","card_bg":"#ffe8f5","header_bg":"#ffd8ee","title_color":"#4a0a3a","body_color":"#7a2a6a","muted_color":"#b070a0","accent":"#ff69b4","accent2":"#ff1493","font_head":"Trebuchet MS","font_body":"Calibri Light","tag":"pastel"},
    {"id":32,"name":"Baby Blue","bg":"#e8f4ff","card_bg":"#d8ecff","header_bg":"#c8e4ff","title_color":"#0a2a5a","body_color":"#2a4a8a","muted_color":"#6080c0","accent":"#4a90d9","accent2":"#2a70b9","font_head":"Calibri","font_body":"Calibri Light","tag":"pastel"},
    {"id":33,"name":"Peach Dream","bg":"#fff5e8","card_bg":"#ffede0","header_bg":"#ffe0c8","title_color":"#4a1a00","body_color":"#7a3a10","muted_color":"#b07840","accent":"#ff7043","accent2":"#e64a19","font_head":"Trebuchet MS","font_body":"Calibri","tag":"pastel"},
    {"id":34,"name":"Sage Garden","bg":"#f0f8f0","card_bg":"#e4f4e4","header_bg":"#d8ecd8","title_color":"#0a2a0a","body_color":"#2a5a2a","muted_color":"#608060","accent":"#66bb6a","accent2":"#43a047","font_head":"Calibri","font_body":"Calibri Light","tag":"pastel"},
    {"id":35,"name":"Lilac Fields","bg":"#f5f0ff","card_bg":"#ede8ff","header_bg":"#e0d8ff","title_color":"#1a0a4a","body_color":"#3a2a7a","muted_color":"#7060b0","accent":"#9575cd","accent2":"#7e57c2","font_head":"Trebuchet MS","font_body":"Calibri","tag":"pastel"},
    {"id":36,"name":"Lemon Zest","bg":"#fffde8","card_bg":"#fff9d0","header_bg":"#fff5b8","title_color":"#2a2000","body_color":"#5a4a00","muted_color":"#908030","accent":"#f9c700","accent2":"#d4a000","font_head":"Calibri","font_body":"Calibri Light","tag":"pastel"},
    {"id":37,"name":"Rose Quartz","bg":"#fef0f4","card_bg":"#fde4ec","header_bg":"#fbd4e2","title_color":"#3a0a1a","body_color":"#6a2a3a","muted_color":"#a06070","accent":"#ec407a","accent2":"#d81b60","font_head":"Trebuchet MS","font_body":"Calibri","tag":"pastel"},
    {"id":38,"name":"Sky Mint","bg":"#e8fff8","card_bg":"#d8ffee","header_bg":"#c8ffe4","title_color":"#003a2a","body_color":"#0a5a3a","muted_color":"#409070","accent":"#26a69a","accent2":"#00897b","font_head":"Calibri","font_body":"Calibri Light","tag":"pastel"},
    {"id":39,"name":"Apricot","bg":"#fff8f0","card_bg":"#fff0e0","header_bg":"#ffe8cc","title_color":"#3a1a00","body_color":"#6a3010","muted_color":"#a06030","accent":"#ffa040","accent2":"#e07020","font_head":"Trebuchet MS","font_body":"Calibri","tag":"pastel"},
    {"id":40,"name":"Powder Blue","bg":"#eef4ff","card_bg":"#e0ecff","header_bg":"#d0e4ff","title_color":"#0a1a40","body_color":"#2a3a70","muted_color":"#6070a0","accent":"#5c85d6","accent2":"#3a63b4","font_head":"Calibri","font_body":"Calibri","tag":"pastel"},

    # ── RETRO / VINTAGE ──────────────────────────────────────────────────────
    {"id":41,"name":"Retro Arcade","bg":"#1a0a2e","card_bg":"#2a1040","header_bg":"#220d38","title_color":"#ffee00","body_color":"#ff9900","muted_color":"#cc6600","accent":"#ff0066","accent2":"#ff00cc","font_head":"Calibri","font_body":"Courier New","tag":"retro"},
    {"id":42,"name":"VHS Purple","bg":"#0d0020","card_bg":"#1a0035","header_bg":"#150028","title_color":"#ff80ff","body_color":"#cc99ff","muted_color":"#8844cc","accent":"#ff00ff","accent2":"#cc00cc","font_head":"Courier New","font_body":"Courier New","tag":"retro"},
    {"id":43,"name":"Sepia Classic","bg":"#2c1a00","card_bg":"#3d2a00","header_bg":"#352200","title_color":"#f5deb3","body_color":"#d4af70","muted_color":"#a07840","accent":"#cd853f","accent2":"#8b6914","font_head":"Calibri","font_body":"Calibri Light","tag":"retro"},
    {"id":44,"name":"Film Noir","bg":"#111111","card_bg":"#1e1e1e","header_bg":"#191919","title_color":"#f5f5f5","body_color":"#cccccc","muted_color":"#888888","accent":"#c8a84b","accent2":"#a08030","font_head":"Trebuchet MS","font_body":"Calibri","tag":"retro"},
    {"id":45,"name":"Pop Art","bg":"#fff700","card_bg":"#ffdd00","header_bg":"#ffc800","title_color":"#000000","body_color":"#1a1a1a","muted_color":"#444444","accent":"#ff0000","accent2":"#cc0000","font_head":"Trebuchet MS","font_body":"Calibri","tag":"retro"},
    {"id":46,"name":"Retro Sunset","bg":"#1a0033","card_bg":"#2a0050","header_bg":"#220040","title_color":"#ff9966","body_color":"#ffcc99","muted_color":"#cc7744","accent":"#ff6600","accent2":"#ff3300","font_head":"Calibri","font_body":"Calibri Light","tag":"retro"},
    {"id":47,"name":"Typewriter","bg":"#f4f0e8","card_bg":"#ede8d8","header_bg":"#e5dfc8","title_color":"#1a1000","body_color":"#3a2800","muted_color":"#7a6040","accent":"#5a3a10","accent2":"#3a2000","font_head":"Courier New","font_body":"Courier New","tag":"retro"},
    {"id":48,"name":"Disco Fever","bg":"#0a001a","card_bg":"#150030","header_bg":"#100025","title_color":"#ffdd00","body_color":"#ff9900","muted_color":"#cc6600","accent":"#ff00ff","accent2":"#cc00cc","font_head":"Trebuchet MS","font_body":"Calibri","tag":"retro"},
    {"id":49,"name":"80s Wave","bg":"#0d0030","card_bg":"#1a0050","header_bg":"#150040","title_color":"#00ffcc","body_color":"#00ccaa","muted_color":"#008870","accent":"#ff0099","accent2":"#cc0077","font_head":"Calibri","font_body":"Courier New","tag":"retro"},
    {"id":50,"name":"Comic Book","bg":"#fff9c4","card_bg":"#fff176","header_bg":"#ffee58","title_color":"#0a0a0a","body_color":"#1a1a1a","muted_color":"#444444","accent":"#e53935","accent2":"#b71c1c","font_head":"Trebuchet MS","font_body":"Calibri","tag":"retro"},

    # ── BOLD / HIGH CONTRAST ─────────────────────────────────────────────────
    {"id":51,"name":"Fire Engine","bg":"#cc0000","card_bg":"#dd1111","header_bg":"#bb0000","title_color":"#ffffff","body_color":"#ffe0e0","muted_color":"#ffaaaa","accent":"#ffff00","accent2":"#ffdd00","font_head":"Trebuchet MS","font_body":"Calibri","tag":"bold"},
    {"id":52,"name":"Deep Space","bg":"#000033","card_bg":"#000055","header_bg":"#000044","title_color":"#ffffff","body_color":"#ccddff","muted_color":"#8899cc","accent":"#00ccff","accent2":"#0099cc","font_head":"Calibri","font_body":"Calibri Light","tag":"bold"},
    {"id":53,"name":"Jungle Green","bg":"#003300","card_bg":"#005500","header_bg":"#004400","title_color":"#ffffff","body_color":"#ccffcc","muted_color":"#88cc88","accent":"#00ff44","accent2":"#00cc33","font_head":"Trebuchet MS","font_body":"Calibri","tag":"bold"},
    {"id":54,"name":"Royal Purple","bg":"#1a0066","card_bg":"#2a0088","header_bg":"#220077","title_color":"#ffffff","body_color":"#ddc8ff","muted_color":"#9970dd","accent":"#cc88ff","accent2":"#aa44ff","font_head":"Calibri","font_body":"Calibri Light","tag":"bold"},
    {"id":55,"name":"Tangerine Pop","bg":"#cc4400","card_bg":"#dd5500","header_bg":"#bb3300","title_color":"#ffffff","body_color":"#ffe8d8","muted_color":"#ffbbaa","accent":"#ffff00","accent2":"#ffee00","font_head":"Trebuchet MS","font_body":"Calibri","tag":"bold"},
    {"id":56,"name":"Electric Blue","bg":"#0000aa","card_bg":"#0000cc","header_bg":"#000099","title_color":"#ffffff","body_color":"#ccddff","muted_color":"#8899ee","accent":"#ff6600","accent2":"#ff4400","font_head":"Calibri","font_body":"Calibri","tag":"bold"},
    {"id":57,"name":"Hot Pink","bg":"#880044","card_bg":"#aa0055","header_bg":"#770033","title_color":"#ffffff","body_color":"#ffd8e8","muted_color":"#ffaacc","accent":"#ffff00","accent2":"#ffee00","font_head":"Trebuchet MS","font_body":"Calibri Light","tag":"bold"},
    {"id":58,"name":"Toxic Lime","bg":"#1a3300","card_bg":"#254400","header_bg":"#1f3800","title_color":"#ccff00","body_color":"#99cc00","muted_color":"#667700","accent":"#00ff00","accent2":"#00cc00","font_head":"Calibri","font_body":"Courier New","tag":"bold"},
    {"id":59,"name":"Crimson Night","bg":"#1a0000","card_bg":"#2d0000","header_bg":"#220000","title_color":"#ff4444","body_color":"#ff8888","muted_color":"#cc4444","accent":"#ff0000","accent2":"#dd0000","font_head":"Trebuchet MS","font_body":"Calibri","tag":"bold"},
    {"id":60,"name":"Ultraviolet","bg":"#0d0033","card_bg":"#1a0055","header_bg":"#140044","title_color":"#cc99ff","body_color":"#9966cc","muted_color":"#664499","accent":"#ff00cc","accent2":"#cc0099","font_head":"Calibri","font_body":"Calibri Light","tag":"bold"},

    # ── GRADIENT / RICH ──────────────────────────────────────────────────────
    {"id":61,"name":"Sunset Gradient","bg":"#1a0530","card_bg":"#2a0840","header_bg":"#200635","title_color":"#ffddaa","body_color":"#ffaa66","muted_color":"#cc7733","accent":"#ff6b35","accent2":"#f7931e","font_head":"Trebuchet MS","font_body":"Calibri","tag":"dark"},
    {"id":62,"name":"Aurora","bg":"#001a1a","card_bg":"#002a2a","header_bg":"#002020","title_color":"#aaffee","body_color":"#66ccbb","muted_color":"#339988","accent":"#00ffcc","accent2":"#00ddaa","font_head":"Calibri","font_body":"Calibri Light","tag":"dark"},
    {"id":63,"name":"Lava Flow","bg":"#1a0500","card_bg":"#2d0800","header_bg":"#220600","title_color":"#ff8c00","body_color":"#ff6600","muted_color":"#cc4400","accent":"#ff2200","accent2":"#dd0000","font_head":"Trebuchet MS","font_body":"Calibri","tag":"dark"},
    {"id":64,"name":"Galaxy","bg":"#050010","card_bg":"#0a0020","header_bg":"#080018","title_color":"#e8e0ff","body_color":"#c0b0ff","muted_color":"#8070cc","accent":"#a855f7","accent2":"#9333ea","font_head":"Calibri","font_body":"Calibri Light","tag":"dark"},
    {"id":65,"name":"Bioluminescent","bg":"#001020","card_bg":"#001830","header_bg":"#001428","title_color":"#80ffee","body_color":"#40ccbb","muted_color":"#208878","accent":"#00e5cc","accent2":"#00bfaa","font_head":"Trebuchet MS","font_body":"Calibri","tag":"dark"},
    {"id":66,"name":"Nebula","bg":"#0f0020","card_bg":"#180035","header_bg":"#130028","title_color":"#ffccff","body_color":"#dd99ff","muted_color":"#aa55cc","accent":"#ff77ff","accent2":"#ff44dd","font_head":"Calibri","font_body":"Calibri Light","tag":"dark"},
    {"id":67,"name":"Deep Ocean","bg":"#001830","card_bg":"#002040","header_bg":"#001c38","title_color":"#a0d8ff","body_color":"#6090cc","muted_color":"#3060a0","accent":"#00aaff","accent2":"#0088dd","font_head":"Trebuchet MS","font_body":"Calibri","tag":"dark"},
    {"id":68,"name":"Ember","bg":"#1a0a00","card_bg":"#2a1200","header_bg":"#220e00","title_color":"#ffcc88","body_color":"#ff9944","muted_color":"#cc6620","accent":"#ff5500","accent2":"#dd3300","font_head":"Calibri","font_body":"Calibri","tag":"dark"},
    {"id":69,"name":"Glacial","bg":"#0a1520","card_bg":"#102030","header_bg":"#0d1a28","title_color":"#d0eeff","body_color":"#90c0e0","muted_color":"#5090b0","accent":"#60c8ff","accent2":"#40aaee","font_head":"Trebuchet MS","font_body":"Calibri Light","tag":"dark"},
    {"id":70,"name":"Prism","bg":"#0a0010","card_bg":"#140020","header_bg":"#0f0018","title_color":"#ffffff","body_color":"#e0d0ff","muted_color":"#a090cc","accent":"#ff3388","accent2":"#cc0066","font_head":"Calibri","font_body":"Calibri Light","tag":"dark"},

    # ── NATURE / ORGANIC ─────────────────────────────────────────────────────
    {"id":71,"name":"Rainforest","bg":"#0d2010","card_bg":"#153020","header_bg":"#102818","title_color":"#ccffcc","body_color":"#88cc88","muted_color":"#448844","accent":"#44ff88","accent2":"#22dd66","font_head":"Trebuchet MS","font_body":"Calibri","tag":"dark"},
    {"id":72,"name":"Desert Storm","bg":"#2a1a00","card_bg":"#3d2800","header_bg":"#332000","title_color":"#fff5cc","body_color":"#ddcc88","muted_color":"#aa9944","accent":"#d4a017","accent2":"#b08000","font_head":"Calibri","font_body":"Calibri Light","tag":"dark"},
    {"id":73,"name":"Cherry Blossom","bg":"#fff0f5","card_bg":"#ffe0ee","header_bg":"#ffd0e5","title_color":"#4a0a2a","body_color":"#8a2a5a","muted_color":"#c07090","accent":"#ff6b9d","accent2":"#ff4480","font_head":"Trebuchet MS","font_body":"Calibri","tag":"light"},
    {"id":74,"name":"Mountain Fog","bg":"#c0c8d8","card_bg":"#d0d8e8","header_bg":"#b8c0d0","title_color":"#1a2030","body_color":"#3a4050","muted_color":"#7080a0","accent":"#3060a0","accent2":"#204880","font_head":"Calibri","font_body":"Calibri Light","tag":"light"},
    {"id":75,"name":"Autumn Harvest","bg":"#1a1000","card_bg":"#2d1c00","header_bg":"#221500","title_color":"#ffcc66","body_color":"#ff9933","muted_color":"#cc6600","accent":"#ff6600","accent2":"#dd4400","font_head":"Trebuchet MS","font_body":"Calibri","tag":"dark"},
    {"id":76,"name":"Ocean Breeze","bg":"#e8f5ff","card_bg":"#d5edff","header_bg":"#c0e3ff","title_color":"#002244","body_color":"#003366","muted_color":"#406080","accent":"#0066cc","accent2":"#004499","font_head":"Calibri","font_body":"Calibri","tag":"light"},
    {"id":77,"name":"Meadow","bg":"#e8f5e8","card_bg":"#d5edd5","header_bg":"#c0e0c0","title_color":"#0a2a0a","body_color":"#1a4a1a","muted_color":"#406840","accent":"#2e7d32","accent2":"#1b5e20","font_head":"Trebuchet MS","font_body":"Calibri Light","tag":"light"},
    {"id":78,"name":"Volcanic","bg":"#1a0a00","card_bg":"#2d1200","header_bg":"#220d00","title_color":"#ff8844","body_color":"#ff6622","muted_color":"#cc4400","accent":"#ff2200","accent2":"#cc0000","font_head":"Calibri","font_body":"Calibri","tag":"dark"},
    {"id":79,"name":"Coral Reef","bg":"#001a20","card_bg":"#002835","header_bg":"#00202a","title_color":"#ff9977","body_color":"#ff7755","muted_color":"#cc5533","accent":"#ff4422","accent2":"#ff6644","font_head":"Trebuchet MS","font_body":"Calibri Light","tag":"dark"},
    {"id":80,"name":"Misty Morning","bg":"#d8e8f0","card_bg":"#e8f4fc","header_bg":"#c8d8e8","title_color":"#102030","body_color":"#2a4060","muted_color":"#607090","accent":"#2080c0","accent2":"#1060a0","font_head":"Calibri","font_body":"Calibri Light","tag":"light"},

    # ── TECH / DIGITAL ───────────────────────────────────────────────────────
    {"id":81,"name":"Terminal Green","bg":"#000a00","card_bg":"#001400","header_bg":"#000f00","title_color":"#00ff00","body_color":"#00cc00","muted_color":"#008800","accent":"#00ff44","accent2":"#00cc33","font_head":"Courier New","font_body":"Courier New","tag":"dark"},
    {"id":82,"name":"Matrix","bg":"#000800","card_bg":"#001000","header_bg":"#000c00","title_color":"#00ff00","body_color":"#00cc00","muted_color":"#007700","accent":"#00ff88","accent2":"#00dd66","font_head":"Courier New","font_body":"Courier New","tag":"dark"},
    {"id":83,"name":"Hologram","bg":"#000f1a","card_bg":"#001828","header_bg":"#001220","title_color":"#00eeff","body_color":"#00aacc","muted_color":"#005588","accent":"#00ddff","accent2":"#00bbee","font_head":"Trebuchet MS","font_body":"Calibri Light","tag":"dark"},
    {"id":84,"name":"Cyberpunk Yellow","bg":"#0a0a00","card_bg":"#151500","header_bg":"#101000","title_color":"#ffff00","body_color":"#cccc00","muted_color":"#888800","accent":"#ffdd00","accent2":"#ccaa00","font_head":"Trebuchet MS","font_body":"Courier New","tag":"neon"},
    {"id":85,"name":"Circuit Board","bg":"#001a0a","card_bg":"#002815","header_bg":"#002010","title_color":"#00ff88","body_color":"#00cc66","muted_color":"#008844","accent":"#00ffcc","accent2":"#00ddaa","font_head":"Courier New","font_body":"Courier New","tag":"neon"},
    {"id":86,"name":"Data Stream","bg":"#040014","card_bg":"#080022","header_bg":"#06001a","title_color":"#4488ff","body_color":"#2266dd","muted_color":"#1144aa","accent":"#66aaff","accent2":"#4488ee","font_head":"Calibri","font_body":"Courier New","tag":"dark"},
    {"id":87,"name":"Pixel Art","bg":"#1a1a2e","card_bg":"#16213e","header_bg":"#0f3460","title_color":"#e94560","body_color":"#e0e0e0","muted_color":"#a0a0a0","accent":"#e94560","accent2":"#c0364a","font_head":"Courier New","font_body":"Calibri","tag":"retro"},
    {"id":88,"name":"Synthwave","bg":"#0d0221","card_bg":"#1a0435","header_bg":"#14032a","title_color":"#ff7eee","body_color":"#df73ff","muted_color":"#9940cc","accent":"#08f7fe","accent2":"#09fbd3","font_head":"Trebuchet MS","font_body":"Calibri Light","tag":"retro"},
    {"id":89,"name":"Quantum","bg":"#000d1a","card_bg":"#001528","header_bg":"#001020","title_color":"#88ccff","body_color":"#4499dd","muted_color":"#2266aa","accent":"#00aaff","accent2":"#0088dd","font_head":"Calibri","font_body":"Calibri Light","tag":"dark"},
    {"id":90,"name":"Neural Net","bg":"#0a000f","card_bg":"#14001e","header_bg":"#0f0018","title_color":"#dd88ff","body_color":"#aa55dd","muted_color":"#7733aa","accent":"#cc44ff","accent2":"#aa22ee","font_head":"Trebuchet MS","font_body":"Calibri","tag":"dark"},

    # ── LUXURY / PREMIUM ─────────────────────────────────────────────────────
    {"id":91,"name":"Black Gold","bg":"#0a0800","card_bg":"#151000","header_bg":"#100c00","title_color":"#ffd700","body_color":"#c8a800","muted_color":"#886e00","accent":"#ffd700","accent2":"#c8a000","font_head":"Calibri","font_body":"Calibri Light","tag":"dark"},
    {"id":92,"name":"Platinum","bg":"#1a1a20","card_bg":"#252530","header_bg":"#1e1e28","title_color":"#e8e8f0","body_color":"#c0c0d0","muted_color":"#808090","accent":"#c0c0c0","accent2":"#a0a0b0","font_head":"Trebuchet MS","font_body":"Calibri Light","tag":"dark"},
    {"id":93,"name":"Rose Gold","bg":"#1a0a0f","card_bg":"#2a1018","header_bg":"#220d14","title_color":"#ffddcc","body_color":"#e8b090","muted_color":"#c07050","accent":"#e8927c","accent2":"#c87060","font_head":"Calibri","font_body":"Calibri Light","tag":"dark"},
    {"id":94,"name":"Obsidian","bg":"#080808","card_bg":"#111111","header_bg":"#0d0d0d","title_color":"#ffffff","body_color":"#dddddd","muted_color":"#888888","accent":"#ff7820","accent2":"#dd5500","font_head":"Trebuchet MS","font_body":"Calibri","tag":"dark"},
    {"id":95,"name":"Sapphire","bg":"#000a2a","card_bg":"#001040","header_bg":"#000d35","title_color":"#c8d8ff","body_color":"#8899cc","muted_color":"#445588","accent":"#4466ff","accent2":"#2244dd","font_head":"Calibri","font_body":"Calibri Light","tag":"dark"},
    {"id":96,"name":"Malachite","bg":"#001a10","card_bg":"#002818","header_bg":"#002014","title_color":"#ccffee","body_color":"#88ddbb","muted_color":"#449977","accent":"#00cc88","accent2":"#00aa66","font_head":"Trebuchet MS","font_body":"Calibri","tag":"dark"},
    {"id":97,"name":"Amethyst","bg":"#150020","card_bg":"#200030","header_bg":"#1a0028","title_color":"#e8ccff","body_color":"#c099ee","muted_color":"#8855bb","accent":"#aa55ff","accent2":"#8833ee","font_head":"Calibri","font_body":"Calibri Light","tag":"dark"},
    {"id":98,"name":"Ivory Tower","bg":"#fafaf0","card_bg":"#ffffff","header_bg":"#f5f5e5","title_color":"#1a1a0a","body_color":"#3a3a2a","muted_color":"#7a7a6a","accent":"#8b7355","accent2":"#6b5535","font_head":"Trebuchet MS","font_body":"Calibri Light","tag":"light"},
    {"id":99,"name":"Onyx","bg":"#0f0f0f","card_bg":"#1a1a1a","header_bg":"#141414","title_color":"#f0f0f0","body_color":"#d0d0d0","muted_color":"#808080","accent":"#ff8c00","accent2":"#e07000","font_head":"Calibri","font_body":"Calibri","tag":"dark"},
    {"id":100,"name":"Diamond","bg":"#f0f8ff","card_bg":"#e8f4ff","header_bg":"#deeeff","title_color":"#0a1a2a","body_color":"#1a3a5a","muted_color":"#5070a0","accent":"#0055aa","accent2":"#003388","font_head":"Trebuchet MS","font_body":"Calibri Light","tag":"light"},

    # ── EXTRA BONUS THEMES ───────────────────────────────────────────────────
    {"id":101,"name":"Neon Tokyo","bg":"#05001a","card_bg":"#0a0030","header_bg":"#070025","title_color":"#ff00aa","body_color":"#cc0088","muted_color":"#880055","accent":"#00eeff","accent2":"#00ccdd","font_head":"Trebuchet MS","font_body":"Calibri Light","tag":"neon"},
    {"id":102,"name":"Miami Vice","bg":"#000a1a","card_bg":"#001528","header_bg":"#001020","title_color":"#ff6ec7","body_color":"#ff8ed7","muted_color":"#cc55a0","accent":"#44ffdd","accent2":"#22ddbb","font_head":"Calibri","font_body":"Calibri Light","tag":"retro"},
    {"id":103,"name":"Nordic Frost","bg":"#e8f0f8","card_bg":"#f0f8ff","header_bg":"#dde8f2","title_color":"#0a1830","body_color":"#2a3050","muted_color":"#607090","accent":"#2a6090","accent2":"#1a4870","font_head":"Trebuchet MS","font_body":"Calibri","tag":"light"},
    {"id":104,"name":"Terracotta","bg":"#f5e8e0","card_bg":"#fff0e8","header_bg":"#ede0d5","title_color":"#2a0a00","body_color":"#5a2a10","muted_color":"#a06040","accent":"#b85042","accent2":"#963830","font_head":"Calibri","font_body":"Calibri Light","tag":"light"},
    {"id":105,"name":"Spearmint","bg":"#002820","card_bg":"#003828","header_bg":"#003020","title_color":"#aaffd0","body_color":"#77ddaa","muted_color":"#449977","accent":"#00ff99","accent2":"#00dd77","font_head":"Trebuchet MS","font_body":"Calibri","tag":"dark"},
]


def get_theme_by_id(theme_id: int) -> dict:
    for t in THEMES:
        if t["id"] == theme_id:
            return t
    return THEMES[0]


def get_random_theme() -> dict:
    return random.choice(THEMES)


def get_themes_list() -> list:
    return [{"id": t["id"], "name": t["name"], "tag": t["tag"],
             "accent": t["accent"], "bg": t["bg"]} for t in THEMES]