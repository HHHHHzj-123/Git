import glob
import json
import os
import re
import sys
import urllib.request

manifest = {}
for path in glob.glob(r"C:\Users\HZJ\Desktop\Git\work\jingxuan_*.html"):
    text = open(path, encoding="utf-8").read()
    # The page embeds another JSON document inside a JavaScript string, so the
    # property quotes and URL slashes are escaped one extra time.
    matches = re.findall(r'\\"main_url\\":\\"(.*?)\\"', text)
    urls = [json.loads('"' + value.replace(r'\"', '"') + '"') for value in matches]
    url = urls[0] if urls else "NONE"
    print(os.path.basename(path), url, sep="\t")
    manifest[os.path.basename(path)] = urls
    if "--download" in sys.argv and url != "NONE":
        video_id = re.search(r"(\d+)\.html$", path).group(1)
        out = os.path.join(os.path.dirname(path), f"video_{video_id}.mp4")
        request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(request, timeout=90) as response, open(out, "wb") as target:
            while chunk := response.read(1024 * 1024):
                target.write(chunk)
        print("saved", out, os.path.getsize(out), sep="\t")

if "--manifest" in sys.argv:
    manifest_path = r"C:\Users\HZJ\Desktop\Git\work\jingxuan_urls.json"
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    print("manifest", manifest_path, sep="\t")
