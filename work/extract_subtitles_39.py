import sys
from pathlib import Path

ROOT = Path(r"C:\Users\HZJ\Desktop\Git\work")
sys.path.insert(0, str(ROOT / "stt_pkg_accessible"))

import av
from PIL import Image, ImageDraw


VIDEO = ROOT / "video_7538263260135427382.mp4"
OUT = ROOT / "subtitles_7538263260135427382"
OUT.mkdir(exist_ok=True)

container = av.open(str(VIDEO))
stream = container.streams.video[0]
duration = float(stream.duration * stream.time_base)
timestamps = [index * 2.5 for index in range(int(duration / 2.5) + 1)]
tiles = []

for second in timestamps:
    container.seek(int(second / stream.time_base), stream=stream, any_frame=False, backward=True)
    image = None
    for frame in container.decode(stream):
        current = float(frame.pts * stream.time_base)
        if current + 0.12 >= second:
            image = frame.to_image().convert("RGB")
            break
    if image is None:
        continue
    crop = image.crop((40, 960, image.width - 40, 1450))
    crop.thumbnail((660, 300), Image.Resampling.LANCZOS)
    tile = Image.new("RGB", (700, 340), "white")
    tile.paste(crop, ((700 - crop.width) // 2, 34))
    draw = ImageDraw.Draw(tile)
    minute = int(second) // 60
    sec = int(second) % 60
    draw.text((10, 8), f"{minute:02d}:{sec:02d}.{int((second % 1) * 10)}", fill="black")
    tiles.append(tile)

for offset in range(0, len(tiles), 12):
    batch = tiles[offset: offset + 12]
    sheet = Image.new("RGB", (1400, 2040), "#E5E7EB")
    for index, tile in enumerate(batch):
        x = (index % 2) * 700
        y = (index // 2) * 340
        sheet.paste(tile, (x, y))
    sheet.save(OUT / f"subtitle_{offset // 12 + 1:02d}.jpg", quality=96)

print(f"duration={duration:.2f}\ttiles={len(tiles)}\tsheets={(len(tiles) + 11) // 12}")
