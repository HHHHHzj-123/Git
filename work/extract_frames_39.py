import sys
from pathlib import Path

ROOT = Path(r"C:\Users\HZJ\Desktop\Git\work")
sys.path.insert(0, str(ROOT / "stt_pkg_accessible"))

import av
from PIL import Image, ImageDraw, ImageFont


VIDEO = ROOT / "video_7538263260135427382.mp4"
OUT = ROOT / "frames_7538263260135427382"
OUT.mkdir(exist_ok=True)

container = av.open(str(VIDEO))
stream = container.streams.video[0]
duration = float(stream.duration * stream.time_base)
timestamps = list(range(0, int(duration) + 1, 5))
frames = []

for second in timestamps:
    container.seek(int(second / stream.time_base), stream=stream, any_frame=False, backward=True)
    frame_image = None
    for frame in container.decode(stream):
        current = float(frame.pts * stream.time_base)
        if current + 0.25 >= second:
            frame_image = frame.to_image().convert("RGB")
            break
    if frame_image is None:
        continue
    frame_image.thumbnail((480, 270), Image.Resampling.LANCZOS)
    tile = Image.new("RGB", (500, 310), "white")
    tile.paste(frame_image, ((500 - frame_image.width) // 2, 24))
    draw = ImageDraw.Draw(tile)
    draw.text((10, 5), f"{second // 60:02d}:{second % 60:02d}", fill="black")
    frames.append(tile)

for page_no in range(0, len(frames), 12):
    batch = frames[page_no: page_no + 12]
    sheet = Image.new("RGB", (1500, 1240), "#DDDDDD")
    for index, tile in enumerate(batch):
        x = (index % 3) * 500
        y = (index // 3) * 310
        sheet.paste(tile, (x, y))
    sheet.save(OUT / f"contact_{page_no // 12 + 1:02d}.jpg", quality=94)

print(f"duration={duration:.2f}\tframes={len(frames)}\tsheets={(len(frames) + 11) // 12}")
