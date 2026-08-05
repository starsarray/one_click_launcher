import runpy
import tempfile
from pathlib import Path


app = runpy.run_path(Path(__file__).with_name("one_click_launcher.pyw"))
matches = app["window_matches_item"]
matches.__globals__["get_process_image_path"] = lambda _pid: r"C:\ImageGlass.exe"

with tempfile.TemporaryDirectory() as directory:
    first = Path(directory, "first.jpg")
    second = Path(directory, "second.jpg")
    first.touch()
    second.touch()
    window = {"pid": 1, "title": "first.jpg | ImageGlass"}
    assert matches(window, str(first), first.stem, r"C:\ImageGlass.exe")
    assert not matches(window, str(second), second.stem, r"C:\ImageGlass.exe")
