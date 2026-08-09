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
    assert matches(window, str(second), second.stem, r"C:\ImageGlass.exe")


events = []
group = {
    "id": "test",
    "items": [
        {"path": "first.jpg", "window": {"exe": r"C:\ImageGlass.exe"}},
        {"path": "second.jpg", "window": {"exe": r"C:\ImageGlass.exe"}},
        {"path": "notes.txt", "window": {"exe": r"C:\Notepad.exe"}},
        {"path": "notes2.txt", "window": {"exe": r"C:\Notepad.exe"}},
    ],
}
launch_group = app["launch_group"]
launch_group.__globals__["load_config"] = lambda: {"groups": [group]}
launch_group.__globals__["time"].sleep = lambda seconds: events.append(f"delay:{seconds}")


def launch(item, threads, claimed_windows):
    events.append(f"launch:{item['path']}")


launch_group.__globals__["launch_item"] = launch
assert launch_group("test", wait_restore=False) == 0
assert events == ["launch:first.jpg", "delay:0.05", "launch:second.jpg", "launch:notes.txt", "delay:0.05", "launch:notes2.txt"]

restore = app["restore_item_window"]
restore.__globals__["get_visible_window_infos"] = lambda **_kwargs: [{"hwnd": 1, "pid": 1}, {"hwnd": 2, "pid": 2}]
restore.__globals__["window_matches_item"] = lambda *_args: True
placed = []
restore.__globals__["set_window_rect"] = lambda hwnd, *_args: placed.append(hwnd)
shared = (set(), app["threading"].Lock())
item = {"path": __file__, "window": {"exe": r"C:\ImageGlass.exe", "rect": {"x": 0, "y": 0, "width": 100, "height": 100}}}
restore(item, set(), shared, launched_pid=2)
restore(item, set(), shared, launched_pid=1)
assert placed == [2, 1]
