from pathlib import Path
app_path = Path(__file__).parent / "app.py"
exec(app_path.read_text(encoding="utf-8"), {"__name__": "__main__", "__file__": str(app_path)})
