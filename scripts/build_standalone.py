"""Bundle the offline app into one HTML file, without any calendar data."""
from pathlib import Path

root = Path(__file__).resolve().parent.parent
html = (root / "index.html").read_text(encoding="utf-8")
for filename in ("planner.js", "local-calendar.js", "app.js"):
    source = (root / filename).read_text(encoding="utf-8").replace("</script", "<\\/script")
    html = html.replace(f'<script src="{filename}" defer></script>', f"<script>\n{source}\n</script>")
destination = root / "dist" / "Focus_Planner_v0.7.html"
destination.parent.mkdir(exist_ok=True)
destination.write_text(html, encoding="utf-8")
print(destination)
