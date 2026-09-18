import sys
from pathlib import Path

ROOT = Path(r"C:\Users\HZJ\Desktop\Git\work")
sys.path.insert(0, str(ROOT / "stt_pkg"))
from faster_whisper.audio import decode_audio

for video_id in ("7404299369547173170", "7406177617952460070", "7408391769232018727"):
    audio = decode_audio(str(ROOT / f"video_{video_id}.mp4"), sampling_rate=16000)
    print(video_id, f"{len(audio) / 16000:.3f}", sep="\t")
