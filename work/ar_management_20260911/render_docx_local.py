from pathlib import Path

source = Path(r"C:\Users\HZJ\.codex\plugins\cache\openai-primary-runtime\documents\26.909.12148\skills\documents\render_docx.py")
code = source.read_text(encoding="utf-8")
code = code.replace(
    'soffice = shutil.which(executable_name)\n    if soffice is None:',
    'soffice = shutil.which(executable_name)\n    if soffice is None and sys.platform == "win32":\n        candidate = r"C:\\Program Files\\LibreOffice\\program\\soffice.exe"\n        soffice = candidate if os.path.exists(candidate) else None\n    if soffice is None:'
)
exec(compile(code, str(source), "exec"), {"__name__": "__main__", "__file__": str(source)})
