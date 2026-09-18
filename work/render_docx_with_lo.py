import os
import runpy
import sys

os.environ["PATH"] = r"C:\Program Files\LibreOffice\program" + os.pathsep + os.environ.get("PATH", "")
renderer = r"C:\Users\HZJ\.codex\plugins\cache\openai-primary-runtime\documents\26.909.12148\skills\documents\render_docx.py"
sys.argv = [renderer, *sys.argv[1:]]
runpy.run_path(renderer, run_name="__main__")
