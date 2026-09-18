import json
import os
import sys
from pathlib import Path

ROOT = Path(r"C:\Users\HZJ\Desktop\Git\work")
sys.path.insert(0, str(ROOT / "stt_pkg_accessible"))

from faster_whisper import WhisperModel

model = WhisperModel(
    "small",
    device="cpu",
    compute_type="int8",
    download_root=str(ROOT / "whisper_models"),
)

requested = set(sys.argv[1:])
inputs = (
    list(ROOT.glob("video_*.mp4"))
    + list(ROOT.glob("audio_*.m4a"))
    + list(ROOT.glob("audio_*.mp4"))
    + list(ROOT.glob("audio_*.mp3"))
)
for video in sorted(inputs):
    if video.stat().st_size == 0:
        continue
    video_id = video.stem.split("_", 1)[1]
    if (ROOT / f"audio_{video_id}.m4a").exists() and video.name.startswith("video_"):
        continue
    if requested and video_id not in requested:
        continue
    segments, info = model.transcribe(
        str(video),
        language="zh",
        beam_size=1,
        vad_filter=False,
        condition_on_previous_text=True,
    )
    rows = []
    lines = []
    for segment in segments:
        text = segment.text.strip()
        rows.append({"start": segment.start, "end": segment.end, "text": text})
        lines.append(f"[{segment.start:07.2f} --> {segment.end:07.2f}] {text}")
    (ROOT / f"transcript_{video_id}.txt").write_text("\n".join(lines), encoding="utf-8")
    (ROOT / f"transcript_{video_id}.json").write_text(
        json.dumps({"language": info.language, "segments": rows}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(video_id, len(rows), info.language, sep="\t", flush=True)
