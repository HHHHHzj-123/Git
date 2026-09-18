import html
import json
import re
from pathlib import Path


ROOT = Path(r"C:\Users\HZJ\Desktop\Git\work")
IDS = [
    "7411032592763407656",
    "7413956278709390618",
    "7413563369455619337",
    "7418870264194206985",
]


def one(pattern, text):
    match = re.search(pattern, text)
    return html.unescape(match.group(1)) if match else ""


metadata = []
for video_id in IDS:
    page = ROOT / f"jingxuan_{video_id}.html"
    text = page.read_text(encoding="utf-8")
    escaped_urls = re.findall(r'\\"main_url\\":\\"(.*?)\\"', text)
    urls = [json.loads('"' + value.replace(r'\"', '"') + '"') for value in escaped_urls]
    item = {
        "id": video_id,
        "title": one(r'name="og:title" content="(.*?)"', text).removesuffix("-抖音精选"),
        "duration": one(r'name="op:video:duration" content="(.*?)"', text),
        "date": one(r'name="op:video:release_date" content="(.*?)"', text),
        "url_count": len(urls),
        "download_url": urls[0] if urls else "",
    }
    metadata.append(item)

(ROOT / "cases_8_11_metadata.json").write_text(
    json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8"
)
print(json.dumps(metadata, ensure_ascii=False, indent=2))
