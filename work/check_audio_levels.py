import sys
from pathlib import Path

ROOT = Path(r"C:\Users\HZJ\Desktop\Git\work")
sys.path.insert(0, str(ROOT / "stt_pkg"))

import numpy as np
from faster_whisper.audio import decode_audio

for path in sorted(ROOT.glob("video_*.mp4")):
    audio = decode_audio(str(path), sampling_rate=16000)
    print(path.name, "samples", len(audio), "seconds", len(audio) / 16000)
    for start in range(0, len(audio), 160000):
        chunk = audio[start:start + 160000]
        rms = float(np.sqrt(np.mean(chunk * chunk))) if len(chunk) else 0
        peak = float(np.max(np.abs(chunk))) if len(chunk) else 0
        print(f"{start/16000:6.1f}-{(start+len(chunk))/16000:6.1f} rms={rms:.6f} peak={peak:.6f}")
