import html
import json
import re
from pathlib import Path

ROOT = Path(r"C:\Users\HZJ\Desktop\Git\work")
IDS = ["7404299369547173170", "7406177617952460070", "7408391769232018727"]
manifest = {}
metadata = {}
for video_id in IDS:
    path = ROOT / f"jingxuan_{video_id}.html"
    text = path.read_text(encoding="utf-8")
    ld_match = re.search(r'<script[^>]+type="application/ld\+json"[^>]*>(.*?)</script>', text, re.S)
    if not ld_match:
        raise RuntimeError(f"No JSON-LD in {path}")
    ld = json.loads(html.unescape(ld_match.group(1)))
    raw_urls = re.findall(r'\\"main_url\\":\\"(.*?)\\"', text)
    urls = [json.loads('"' + raw.replace(r'\"', '"') + '"') for raw in raw_urls]
    metadata[video_id] = {
        "name": ld.get("name"),
        "description": ld.get("description"),
        "uploadDate": ld.get("uploadDate"),
        "duration": ld.get("duration"),
        "author": (ld.get("author") or {}).get("name"),
        "url_count": len(urls),
    }
    manifest[video_id] = urls

(ROOT / "cases_5_7_metadata.json").write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
(ROOT / "cases_5_7_urls.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps(metadata, ensure_ascii=False, indent=2))
