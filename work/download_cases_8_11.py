import json
import subprocess
from pathlib import Path


ROOT = Path(r"C:\Users\HZJ\Desktop\Git\work")
items = json.loads((ROOT / "cases_8_11_metadata.json").read_text(encoding="utf-8"))
for item in items:
    output = ROOT / f"video_{item['id']}.mp4"
    subprocess.run(
        [
            "curl.exe", "-L", "--fail", "--silent", "--show-error",
            "--retry", "3", "--output", str(output), item["download_url"],
        ],
        check=True,
    )
    print(item["id"], output.stat().st_size, flush=True)
