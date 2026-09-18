import json
import sys
from pathlib import Path

ROOT = Path(r"C:\Users\HZJ\Desktop\Git\work")
sys.path.insert(0, str(ROOT / "stt_pkg"))
from faster_whisper import WhisperModel

video_id = sys.argv[1]
start = float(sys.argv[2])
end = float(sys.argv[3])
model = WhisperModel("small", device="cpu", compute_type="int8", download_root=str(ROOT / "whisper_models"))
segments, _ = model.transcribe(
    str(ROOT / f"video_{video_id}.mp4"),
    language="zh",
    beam_size=5,
    vad_filter=False,
    clip_timestamps=f"{start},{end}",
    condition_on_previous_text=False,
)
rows = []
for segment in segments:
    text = segment.text.strip()
    rows.append({"start": segment.start, "end": segment.end, "text": text})
(ROOT / f"transcript_{video_id}_tail.txt").write_text(
    "\n".join(f"[{r['start']:07.2f} --> {r['end']:07.2f}] {r['text']}" for r in rows), encoding="utf-8"
)
print(video_id, len(rows), flush=True)
