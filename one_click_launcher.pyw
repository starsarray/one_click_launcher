import argparse
import ctypes
from ctypes import wintypes
import json
import os
import re
import subprocess
import sys
import threading
import time
import uuid
from pathlib import Path
from urllib.parse import unquote
from tkinter import (
    BOTH,
    END,
    LEFT,
    RIGHT,
    SINGLE,
    VERTICAL,
    W,
    X,
    Y,
    Button,
    Entry,
    Frame,
    Label,
    Listbox,
    Menu,
    Scrollbar,
    StringVar,
    Tk,
    filedialog,
    messagebox,
    simpledialog,
    ttk,
)


APP_NAME = "一键启动"
APP_VERSION = "v0.4.34-demo"
CONFIG_FILE = "launcher_config.json"
OFFICE_CAPTURE_SPECS = (
    ({"wps", "kwps", "et", "ket", "wpp", "kwpp"}, "KWPS.Application", "Documents"),
    ({"wps", "kwps", "et", "ket", "wpp", "kwpp"}, "KET.Application", "Workbooks"),
    ({"wps", "kwps", "et", "ket", "wpp", "kwpp"}, "KWPP.Application", "Presentations"),
    ({"wps", "kwps", "et", "ket", "wpp", "kwpp", "wpspdf", "kpdf"}, "KPDF.Application", "Documents"),
    ({"winword"}, "Word.Application", "Documents"),
    ({"excel"}, "Excel.Application", "Workbooks"),
    ({"powerpnt"}, "PowerPoint.Application", "Presentations"),
)


def default_config() -> dict:
    return {"groups": [], "rule_groups": [], "active_rule_group_ids": []}


class GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", wintypes.DWORD),
        ("Data2", wintypes.WORD),
        ("Data3", wintypes.WORD),
        ("Data4", wintypes.BYTE * 8),
    ]

    def __init__(self, value: str):
        guid = uuid.UUID(value)
        super().__init__(
            guid.time_low,
            guid.time_mid,
            guid.time_hi_version,
            (wintypes.BYTE * 8)(*guid.bytes[8:]),
        )


class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


class WINDOWPLACEMENT(ctypes.Structure):
    _fields_ = [
        ("length", ctypes.c_uint),
        ("flags", ctypes.c_uint),
        ("showCmd", ctypes.c_uint),
        ("ptMinPosition", POINT),
        ("ptMaxPosition", POINT),
        ("rcNormalPosition", wintypes.RECT),
    ]


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def app_entry_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve()
    return Path(__file__).resolve()


def config_path() -> Path:
    return app_dir() / CONFIG_FILE


def shortcut_trash_dir() -> Path:
    path = app_dir() / ".shortcut_trash"
    path.mkdir(exist_ok=True)
    return path


def notify_shell_path(path: Path, event: int) -> None:
    ctypes.windll.shell32.SHChangeNotify(event, 0x2005, str(path), None)
    refresh_desktop_view()


def refresh_desktop_view() -> None:
    try:
        ctypes.windll.shell32.SHChangeNotify(0x00001000, 0x2005, str(get_desktop_dir()), None)
        user32 = ctypes.windll.user32
        handles = []

        def add_desktop_list(hwnd):
            shell = user32.FindWindowExW(hwnd, 0, "SHELLDLL_DefView", None)
            if shell:
                listview = user32.FindWindowExW(shell, 0, "SysListView32", None)
                if listview:
                    handles.append(listview)

        add_desktop_list(user32.FindWindowW("Progman", None))
        worker = 0
        while True:
            worker = user32.FindWindowExW(0, worker, "WorkerW", None)
            if not worker:
                break
            add_desktop_list(worker)
        for hwnd in handles:
            user32.InvalidateRect(hwnd, None, True)
            user32.UpdateWindow(hwnd)
            user32.PostMessageW(hwnd, 0x0111, 28931, 0)
    except Exception:
        pass


def load_config() -> dict:
    path = config_path()
    if not path.exists():
        return default_config()
    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return default_config()
    data.setdefault("groups", [])
    data.setdefault("rule_groups", [])
    if "active_rule_group_ids" not in data:
        old_id = data.get("active_rule_group_id")
        data["active_rule_group_ids"] = [old_id] if old_id else []
    data["active_rule_group_ids"] = list(dict.fromkeys(group_id for group_id in data.get("active_rule_group_ids", []) if group_id))
    data.pop("active_rule_group_id", None)
    return data


def save_config(data: dict) -> None:
    with config_path().open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def is_url(target: str) -> bool:
    if re.match(r"^[a-zA-Z]:[\\/]", target or ""):
        return False
    return re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*:", target or "") is not None


def expand_target(target: str) -> str:
    return os.path.expandvars(os.path.expanduser(target.strip()))


def launch_item(item: dict, restore_threads: list[threading.Thread] | None = None) -> str | None:
    target = expand_target(item.get("path", ""))
    if not target:
        return "空路径"

    args = item.get("args", "").strip()
    if not is_url(target) and not Path(target).exists():
        return f"不存在：{target}"

    window = item.get("window") or {}
    before_hwnds = {window["hwnd"] for window in get_visible_window_infos()} if valid_window_rect(window.get("rect") or {}) else set()
    try:
        if args:
            os.startfile(target, arguments=args)
        else:
            os.startfile(target)
    except Exception as exc:
        return f"{target}：{exc}"
    if before_hwnds:
        thread = threading.Thread(target=restore_item_window, args=(item, before_hwnds), daemon=True)
        thread.start()
        if restore_threads is not None:
            restore_threads.append(thread)
    return None


def restore_item_window(item: dict, before_hwnds: set[int] | None = None) -> None:
    window = item.get("window") or {}
    rect = window.get("rect") or {}
    if not valid_window_rect(rect):
        return

    target = expand_target(item.get("path", ""))
    target_stem = Path(target).stem.lower()
    exe_path = os.path.normcase(window.get("exe", ""))
    before_hwnds = before_hwnds or set()

    for attempt in range(40):
        for info in get_visible_window_infos():
            title = (info.get("title") or "").lower()
            matched = (target_stem and target_stem in title) or (exe_path and os.path.normcase(get_process_image_path(info["pid"])) == exe_path)
            if not matched:
                continue
            if info["hwnd"] not in before_hwnds:
                set_window_rect(info["hwnd"], rect, bool(window.get("maximized")))
                return
        time.sleep(0.2)


def find_group(data: dict, key: str) -> dict | None:
    for group in data.get("groups", []):
        if group.get("id") == key or group.get("name") == key:
            return group
    return None


def find_rule_group(data: dict, key: str) -> dict | None:
    for group in data.get("rule_groups", []):
        if group.get("id") == key or group.get("name") == key:
            return group
    return None


def normalize_rule_path(path: str) -> str:
    value = expand_target(path or "").strip().strip('"')
    if is_url(value):
        return value.rstrip("/").lower()
    try:
        return os.path.normcase(str(Path(value).resolve()))
    except Exception:
        return os.path.normcase(value)


def active_rule_group_ids(data: dict) -> list[str]:
    ids = data.get("active_rule_group_ids", [])
    if not ids and data.get("active_rule_group_id"):
        ids = [data.get("active_rule_group_id")]
    return list(dict.fromkeys(group_id for group_id in ids if group_id))


def active_rule_groups(data: dict) -> list[dict]:
    groups = []
    for group_id in active_rule_group_ids(data):
        group = find_rule_group(data, group_id)
        if group:
            groups.append(group)
    return groups


def active_rule_names(data: dict) -> list[str]:
    return [group.get("name", "未命名规则组") for group in active_rule_groups(data)]


def item_blocked_by_active_rules(data: dict, item: dict) -> bool:
    groups = active_rule_groups(data)
    if not groups:
        return False

    item_path = normalize_rule_path(item.get("path", ""))
    item_args = (item.get("args", "") or "").strip()
    merged_rules = set()
    for group in groups:
        for rule in group.get("rules", []):
            if not rule.get("enabled", True):
                continue
            rule_path = normalize_rule_path(rule.get("path", ""))
            rule_args = (rule.get("args", "") or "").strip()
            if not rule_path:
                continue
            merged_rules.add((rule_path, rule_args))

    for rule_path, rule_args in merged_rules:
        if rule_path != item_path:
            continue
        if not rule_args or rule_args == item_args:
            return True
    return False


def launch_group(group_key: str, show_done: bool = False, wait_restore: bool = True) -> int:
    data = load_config()
    group = find_group(data, group_key)
    if not group:
        messagebox.showerror(APP_NAME, f"找不到工作组：{group_key}")
        return 1

    failures = []
    restore_threads = []
    for item in group.get("items", []):
        if item.get("enabled", True):
            error = launch_item(item, restore_threads)
            if error:
                failures.append(error)
    if wait_restore:
        for thread in restore_threads:
            thread.join()

    if failures:
        messagebox.showwarning(APP_NAME, "部分项目启动失败：\n\n" + "\n".join(failures))
        return 2

    if show_done:
        messagebox.showinfo(APP_NAME, f"已启动：{group.get('name', '')}")
    return 0


def get_window_text(hwnd: int) -> str:
    length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
    if length <= 0:
        return ""
    buffer = ctypes.create_unicode_buffer(length + 1)
    ctypes.windll.user32.GetWindowTextW(hwnd, buffer, length + 1)
    return buffer.value


def get_foreground_window_info() -> dict:
    hwnd = ctypes.windll.user32.GetForegroundWindow()
    if not hwnd:
        raise RuntimeError("没有找到当前窗口。")

    pid = wintypes.DWORD()
    ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    if not pid.value:
        raise RuntimeError("没有读取到窗口进程。")

    return {"hwnd": hwnd, "pid": pid.value, "title": get_window_text(hwnd), "window": get_window_state(hwnd)}


def get_window_state(hwnd: int) -> dict:
    user32 = ctypes.windll.user32
    rect = wintypes.RECT()
    maximized = bool(user32.IsZoomed(hwnd))
    if user32.IsIconic(hwnd):
        placement = WINDOWPLACEMENT()
        placement.length = ctypes.sizeof(WINDOWPLACEMENT)
        if not user32.GetWindowPlacement(hwnd, ctypes.byref(placement)):
            return {}
        rect = placement.rcNormalPosition
        maximized = bool(placement.flags & 0x0002)
    elif not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
        return {}
    state = {
        "rect": {
            "x": int(rect.left),
            "y": int(rect.top),
            "width": int(rect.right - rect.left),
            "height": int(rect.bottom - rect.top),
        },
        "maximized": maximized,
    }
    return state if valid_window_rect(state["rect"]) else {}


def valid_window_rect(rect: dict) -> bool:
    return (
        all(key in rect for key in ("x", "y", "width", "height"))
        and int(rect["width"]) > 0
        and int(rect["height"]) > 0
        and int(rect["x"]) > -10000
        and int(rect["y"]) > -10000
    )


def set_window_rect(hwnd: int, rect: dict, maximized: bool = False) -> None:
    user32 = ctypes.windll.user32
    user32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_uint]
    user32.ShowWindow(hwnd, 9)
    user32.SetWindowPos(hwnd, 0, int(rect["x"]), int(rect["y"]), int(rect["width"]), int(rect["height"]), 0x0014)
    if maximized:
        user32.ShowWindow(hwnd, 3)


def get_window_pid(hwnd: int) -> int:
    pid = wintypes.DWORD()
    ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return pid.value


def get_visible_window_infos() -> list[dict]:
    windows = []
    own_pid = os.getpid()
    enum_proc_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    def callback(hwnd, _lparam):
        if not ctypes.windll.user32.IsWindowVisible(hwnd):
            return True
        title = get_window_text(hwnd).strip()
        if not title:
            return True
        pid = get_window_pid(hwnd)
        if not pid or pid == own_pid:
            return True
        windows.append({"hwnd": hwnd, "pid": pid, "title": title, "window": get_window_state(hwnd)})
        return True

    ctypes.windll.user32.EnumWindows(enum_proc_type(callback), 0)
    return windows


def get_process_image_path(pid: int) -> str:
    process_query_limited_information = 0x1000
    handle = ctypes.windll.kernel32.OpenProcess(process_query_limited_information, False, pid)
    if not handle:
        return ""
    try:
        size = wintypes.DWORD(32768)
        buffer = ctypes.create_unicode_buffer(size.value)
        ok = ctypes.windll.kernel32.QueryFullProcessImageNameW(handle, 0, buffer, ctypes.byref(size))
        return buffer.value if ok else ""
    finally:
        ctypes.windll.kernel32.CloseHandle(handle)


def get_process_command_line(pid: int) -> str:
    script = (
        "$p = Get-CimInstance Win32_Process -Filter \"ProcessId="
        + str(pid)
        + "\"; if ($p) { $p.CommandLine }"
    )
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", script],
            capture_output=True,
            text=True,
            timeout=5,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:
        return ""
    return result.stdout.strip()


def get_process_command_lines(pids: list[int]) -> dict[int, str]:
    ids = sorted({int(pid) for pid in pids if pid})
    if not ids:
        return {}

    filter_text = " OR ".join(f"ProcessId = {pid}" for pid in ids)
    script = (
        'Get-CimInstance Win32_Process -Filter "'
        + filter_text
        + '" | Select-Object ProcessId,CommandLine | ConvertTo-Json -Compress'
    )
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", script],
            capture_output=True,
            text=True,
            timeout=8,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:
        return {}

    output = result.stdout.strip()
    if not output:
        return {}
    try:
        records = json.loads(output)
    except Exception:
        return {}
    if isinstance(records, dict):
        records = [records]

    command_lines = {}
    for record in records:
        try:
            pid = int(record.get("ProcessId"))
        except Exception:
            continue
        command_lines[pid] = record.get("CommandLine") or ""
    return command_lines


def split_command_line(command_line: str) -> list[str]:
    if not command_line:
        return []

    argc = ctypes.c_int()
    ctypes.windll.shell32.CommandLineToArgvW.argtypes = [wintypes.LPCWSTR, ctypes.POINTER(ctypes.c_int)]
    ctypes.windll.shell32.CommandLineToArgvW.restype = ctypes.POINTER(ctypes.c_wchar_p)
    argv = ctypes.windll.shell32.CommandLineToArgvW(command_line, ctypes.byref(argc))
    if not argv:
        return []
    try:
        return [argv[i] for i in range(argc.value)]
    finally:
        ctypes.windll.kernel32.LocalFree(argv)


def normalize_capture_value(value: str) -> str:
    value = (value or "").strip().strip('"')
    if value.lower().startswith("file:///"):
        value = unquote(value[8:])
        if re.match(r"^/[a-zA-Z]:/", value):
            value = value[1:]
        value = value.replace("/", "\\")
    return os.path.expandvars(os.path.expanduser(value))


def iter_capture_values(args: list[str]):
    for arg in args:
        if not arg:
            continue
        yield arg
        if "=" in arg:
            yield arg.split("=", 1)[1]


def choose_captured_target(exe_path: str, command_line: str) -> tuple[str, str, str]:
    args = split_command_line(command_line)
    exe_resolved = str(Path(exe_path).resolve()) if exe_path else ""

    for value in iter_capture_values(args[1:]):
        candidate = normalize_capture_value(value)
        if is_url(candidate):
            return candidate, "", "网址"
        if candidate and Path(candidate).exists():
            try:
                if exe_resolved and str(Path(candidate).resolve()).lower() == exe_resolved.lower():
                    continue
            except Exception:
                pass
            return candidate, "", "文件"

    if exe_path:
        return exe_path, subprocess.list2cmdline(args[1:]), "软件"
    raise RuntimeError("没有识别到可启动的文件或软件。")


def iter_com_items(collection):
    try:
        count = int(collection.Count)
        for index in range(1, count + 1):
            yield collection.Item(index)
        return
    except Exception:
        pass

    try:
        yield from collection
    except Exception:
        return


def get_open_office_document_paths(exe_path: str) -> list[str]:
    exe_name = Path(exe_path).stem.lower() if exe_path else ""
    specs = [spec for spec in OFFICE_CAPTURE_SPECS if exe_name in spec[0]]
    if not specs:
        return []

    try:
        import pythoncom
        import win32com.client
    except Exception:
        return []

    paths = []
    seen = set()
    pythoncom.CoInitialize()
    try:
        for _names, prog_id, collection_name in specs:
            try:
                app = win32com.client.GetActiveObject(prog_id)
                collection = getattr(app, collection_name)
            except Exception:
                continue
            for document in iter_com_items(collection):
                try:
                    path = str(document.FullName)
                except Exception:
                    continue
                if not path or not Path(path).exists():
                    continue
                key = str(Path(path).resolve()).lower()
                if key in seen:
                    continue
                seen.add(key)
                paths.append(path)
    finally:
        pythoncom.CoUninitialize()
    return paths


def get_explorer_folder_path(hwnd: int) -> str:
    try:
        import pythoncom
        import win32com.client
    except Exception:
        return ""

    pythoncom.CoInitialize()
    try:
        shell = win32com.client.Dispatch("Shell.Application")
        for window in shell.Windows():
            try:
                if int(window.HWND) != int(hwnd):
                    continue
                path = normalize_capture_value(str(window.LocationURL or ""))
                return path if path and Path(path).is_dir() else ""
            except Exception:
                continue
    finally:
        pythoncom.CoUninitialize()
    return ""


def capture_window_launch_items(info: dict, command_line: str | None = None) -> tuple[list[dict], list[str]]:
    exe_path = get_process_image_path(info["pid"])
    window = {"title": info.get("title", ""), "exe": exe_path, **(info.get("window") or {})}
    if not valid_window_rect(window.get("rect") or {}):
        window = {}
    if Path(exe_path).name.lower() == "explorer.exe":
        folder_path = get_explorer_folder_path(info["hwnd"])
        if not folder_path:
            raise RuntimeError("跳过后台资源管理器。")
        item = {"path": folder_path, "args": "", "enabled": True, **({"window": window} if window else {})}
        return [item], [f"文件夹：{Path(folder_path).name or folder_path}"]
    office_paths = get_open_office_document_paths(exe_path)
    if office_paths:
        items = [{"path": path, "args": "", "enabled": True, **({"window": window} if window else {})} for path in office_paths]
        return items, [f"文件：{Path(path).name}" for path in office_paths]

    command_line = command_line if command_line is not None else get_process_command_line(info["pid"])
    target, args, target_type = choose_captured_target(exe_path, command_line)
    item = {"path": target, "args": args, "enabled": True, **({"window": window} if window else {})}
    title = info.get("title") or Path(target).name or target
    return [item], [f"{target_type}：{title}"]


def capture_window_launch_item(info: dict, command_line: str | None = None) -> tuple[dict, str]:
    items, statuses = capture_window_launch_items(info, command_line)
    return items[0], statuses[0]


def should_skip_capture_item(item: dict) -> bool:
    return False


def capture_foreground_launch_item() -> tuple[dict, str]:
    items, statuses = capture_foreground_launch_items()
    return items[0], statuses[0]


def capture_foreground_launch_items() -> tuple[list[dict], list[str]]:
    info = get_foreground_window_info()
    if info["pid"] == os.getpid():
        raise RuntimeError("捕捉到的是一键启动窗口，请点捕捉后切换到目标软件。")

    items, statuses = capture_window_launch_items(info)
    return items, [f"已捕捉{status}" for status in statuses]


def capture_all_visible_launch_items() -> tuple[list[dict], list[str]]:
    windows = get_visible_window_infos()
    if not windows:
        raise RuntimeError("没有找到可捕捉的窗口。")

    command_lines = get_process_command_lines([window["pid"] for window in windows])
    items = []
    statuses = []
    seen = set()

    for window in windows:
        try:
            window_items, window_statuses = capture_window_launch_items(window, command_lines.get(window["pid"], ""))
        except Exception:
            continue
        for item, status in zip(window_items, window_statuses):
            if should_skip_capture_item(item):
                continue
            key = (item.get("path", "").lower(), item.get("args", "").lower())
            if key in seen:
                continue
            seen.add(key)
            items.append(item)
            statuses.append(status)

    if not items:
        raise RuntimeError("没有识别到可启动的文件或软件。")
    return items, statuses


def get_desktop_dir() -> Path:
    folder_id_desktop = GUID("B4BFCC3A-DB2C-424C-B029-7FE99A87C641")
    path_ptr = wintypes.LPWSTR()
    try:
        ctypes.windll.shell32.SHGetKnownFolderPath(
            ctypes.byref(folder_id_desktop),
            0,
            None,
            ctypes.byref(path_ptr),
        )
        if path_ptr.value:
            return Path(path_ptr.value)
    except Exception:
        pass
    finally:
        if path_ptr:
            ctypes.windll.ole32.CoTaskMemFree(path_ptr)
    return Path.home() / "Desktop"


def find_pythonw() -> str:
    exe = Path(sys.executable)
    pythonw = exe.with_name("pythonw.exe")
    return str(pythonw if pythonw.exists() else exe)


def shortcut_target() -> str:
    return str(app_entry_path() if getattr(sys, "frozen", False) else find_pythonw())


def safe_shortcut_name(name: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]', "_", name).strip()
    return cleaned or "未命名工作组"


def check_hresult(hr: int, action: str) -> None:
    if hr < 0:
        raise OSError(ctypes.c_long(hr).value, action)


def com_method(ptr: ctypes.c_void_p, index: int, restype, *argtypes):
    vtable = ctypes.cast(ptr, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
    return ctypes.WINFUNCTYPE(restype, ctypes.c_void_p, *argtypes)(vtable[index])


def create_shortcut_with_com(shortcut_path: Path, target: str, args: str, workdir: str, description: str) -> None:
    clsid_shell_link = GUID("00021401-0000-0000-C000-000000000046")
    iid_shell_link = GUID("000214F9-0000-0000-C000-000000000046")
    iid_persist_file = GUID("0000010B-0000-0000-C000-000000000046")
    clsctx_inproc_server = 1
    sw_shownormal = 1
    rpc_e_changed_mode = -2147417850

    ole32 = ctypes.oledll.ole32
    initialized = False
    shell_link = ctypes.c_void_p()
    persist_file = ctypes.c_void_p()

    hr = ole32.CoInitialize(None)
    if hr >= 0:
        initialized = True
    elif hr != rpc_e_changed_mode:
        check_hresult(hr, "初始化 COM")

    try:
        hr = ole32.CoCreateInstance(
            ctypes.byref(clsid_shell_link),
            None,
            clsctx_inproc_server,
            ctypes.byref(iid_shell_link),
            ctypes.byref(shell_link),
        )
        check_hresult(hr, "创建快捷方式对象")

        set_description = com_method(shell_link, 7, ctypes.c_long, wintypes.LPCWSTR)
        set_workdir = com_method(shell_link, 9, ctypes.c_long, wintypes.LPCWSTR)
        set_args = com_method(shell_link, 11, ctypes.c_long, wintypes.LPCWSTR)
        set_show_cmd = com_method(shell_link, 15, ctypes.c_long, ctypes.c_int)
        set_icon = com_method(shell_link, 17, ctypes.c_long, wintypes.LPCWSTR, ctypes.c_int)
        set_path = com_method(shell_link, 20, ctypes.c_long, wintypes.LPCWSTR)

        check_hresult(set_path(shell_link, target), "设置快捷方式目标")
        check_hresult(set_args(shell_link, args), "设置快捷方式参数")
        check_hresult(set_workdir(shell_link, workdir), "设置工作目录")
        check_hresult(set_description(shell_link, description), "设置描述")
        check_hresult(set_icon(shell_link, str(Path(os.environ["SystemRoot"]) / "System32" / "shell32.dll"), 167), "设置图标")
        check_hresult(set_show_cmd(shell_link, sw_shownormal), "设置显示方式")

        query_interface = com_method(
            shell_link,
            0,
            ctypes.c_long,
            ctypes.POINTER(GUID),
            ctypes.POINTER(ctypes.c_void_p),
        )
        check_hresult(query_interface(shell_link, ctypes.byref(iid_persist_file), ctypes.byref(persist_file)), "获取保存接口")

        save = com_method(persist_file, 6, ctypes.c_long, wintypes.LPCWSTR, wintypes.BOOL)
        check_hresult(save(persist_file, str(shortcut_path), True), "保存快捷方式")
    finally:
        if persist_file:
            com_method(persist_file, 2, ctypes.c_ulong) (persist_file)
        if shell_link:
            com_method(shell_link, 2, ctypes.c_ulong) (shell_link)
        if initialized:
            ole32.CoUninitialize()


def create_shortcut_with_powershell(shortcut_path: Path, target: str, args: str, workdir: str, description: str) -> None:
    script = """
$wsh = New-Object -ComObject WScript.Shell
$shortcut = $wsh.CreateShortcut($env:LA_SHORTCUT)
$shortcut.TargetPath = $env:LA_TARGET
$shortcut.Arguments = $env:LA_ARGS
$shortcut.WorkingDirectory = $env:LA_WORKDIR
$shortcut.Description = $env:LA_DESC
$shortcut.IconLocation = "$env:SystemRoot\\System32\\shell32.dll,167"
$shortcut.Save()
"""
    env = os.environ.copy()
    env.update(
        {
            "LA_SHORTCUT": str(shortcut_path),
            "LA_TARGET": target,
            "LA_ARGS": args,
            "LA_WORKDIR": workdir,
            "LA_DESC": description,
        }
    )
    subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        env=env,
        capture_output=True,
        text=True,
        timeout=20,
        check=True,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )


def create_shortcut(shortcut_path: Path, target: str, args: str, workdir: str, description: str) -> None:
    try:
        create_shortcut_with_com(shortcut_path, target, args, workdir, description)
    except Exception:
        create_shortcut_with_powershell(shortcut_path, target, args, workdir, description)


class LauncherApp:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title(f"{APP_NAME} {APP_VERSION}")
        self.data = load_config()
        geometry = self.data.get("app_geometry", "")
        self.root.geometry(geometry if re.match(r"^\d+x\d+([+-]\d+){0,2}$", geometry) else "860x520")
        self.root.minsize(700, 460)
        if self.data.get("app_window_state") == "zoomed":
            self.root.after(0, lambda: self.root.state("zoomed"))
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.selected_group_id = None
        self.selected_rule_group_id = None
        self.status = StringVar(value="配置已自动保存在本机。")

        self.build_ui()
        self.refresh_groups()
        self.refresh_rule_groups()

    def close(self) -> None:
        state = self.root.state()
        if state == "normal":
            self.data["app_geometry"] = self.root.geometry()
        self.data["app_window_state"] = state
        save_config(self.data)
        self.root.destroy()

    def setup_style(self) -> None:
        style = ttk.Style(self.root)
        for theme in ("vista", "xpnative", "clam"):
            if theme in style.theme_names():
                style.theme_use(theme)
                break
        style.configure("TButton", padding=(12, 6))
        style.configure("Primary.TButton", padding=(14, 6))
        style.configure("Section.TLabel", font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 11, "bold"))
        style.configure("Treeview", rowheight=30, font=("Microsoft YaHei UI", 10))
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("Status.TLabel", padding=(8, 4))

    def build_ui(self) -> None:
        self.root.option_add("*Font", "{Microsoft YaHei UI} 10")
        self.setup_style()

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=BOTH, expand=True)
        launch_tab = ttk.Frame(self.notebook)
        rules_tab = ttk.Frame(self.notebook)
        self.notebook.add(launch_tab, text="启动工作组")
        self.notebook.add(rules_tab, text="捕捉规则")

        main = ttk.Frame(launch_tab, padding=14)
        main.pack(fill=BOTH, expand=True)
        main.columnconfigure(0, weight=0, minsize=210)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)

        left = ttk.Frame(main)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 14))

        ttk.Label(left, text="工作组", anchor=W, style="Section.TLabel").pack(fill=X)
        group_wrap = ttk.Frame(left)
        group_wrap.pack(fill=BOTH, expand=True, pady=(6, 8))
        group_scroll = ttk.Scrollbar(group_wrap, orient=VERTICAL)
        self.group_list = Listbox(
            group_wrap,
            height=5,
            exportselection=False,
            activestyle="dotbox",
            selectmode=SINGLE,
            borderwidth=0,
            highlightthickness=1,
            highlightbackground="#d7dce5",
            highlightcolor="#8aa4d6",
            relief="flat",
            bg="#ffffff",
            selectbackground="#2563eb",
            selectforeground="#ffffff",
            font=("Microsoft YaHei UI", 10),
        )
        self.group_list.pack(side=LEFT, fill=BOTH, expand=True)
        group_scroll.pack(side=RIGHT, fill=Y)
        self.group_list.config(yscrollcommand=group_scroll.set)
        group_scroll.config(command=self.group_list.yview)
        self.group_list.bind("<<ListboxSelect>>", self.on_group_select)

        ttk.Button(left, text="新建", command=self.add_group).pack(fill=X, pady=2)
        ttk.Button(left, text="重命名", command=self.rename_group).pack(fill=X, pady=2)
        ttk.Button(left, text="删除", command=self.delete_group).pack(fill=X, pady=2)
        ttk.Separator(left).pack(fill=X, pady=10)
        ttk.Button(left, text="创建本软件快捷方式", command=self.create_manager_shortcut).pack(fill=X, pady=2)

        right = ttk.Frame(main)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(2, weight=1)

        top = ttk.Frame(right)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)
        ttk.Label(top, text="当前工作组：", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        self.group_name = StringVar(value="未选择")
        ttk.Label(top, textvariable=self.group_name, anchor=W).grid(row=0, column=1, sticky="ew")
        ttk.Button(top, text="清空当前工作组", command=self.clear_current_group_items).grid(row=0, column=2, sticky="e", padx=(8, 0))
        ttk.Button(top, text="创建桌面快捷方式", command=self.create_group_shortcut).grid(row=0, column=3, sticky="e", padx=(8, 0))
        ttk.Button(top, text="运行", command=self.run_selected_group, style="Primary.TButton").grid(row=0, column=4, sticky="e", padx=(8, 0))

        ttk.Label(right, text="启动项目", anchor=W, style="Section.TLabel").grid(row=1, column=0, sticky="ew", pady=(14, 0))
        item_wrap = ttk.Frame(right)
        item_wrap.grid(row=2, column=0, sticky="nsew", pady=(6, 8))
        item_wrap.columnconfigure(0, weight=1)
        item_wrap.rowconfigure(0, weight=1)
        item_scroll = ttk.Scrollbar(item_wrap, orient=VERTICAL)
        self.item_tree = ttk.Treeview(
            item_wrap,
            columns=("delete", "name", "path"),
            show="headings",
            selectmode="browse",
            height=5,
        )
        self.item_tree.heading("delete", text="操作")
        self.item_tree.heading("name", text="名称")
        self.item_tree.heading("path", text="路径 / 参数")
        self.item_tree.column("delete", width=68, minwidth=58, stretch=False, anchor="center")
        self.item_tree.column("name", width=210, minwidth=110, stretch=True, anchor=W)
        self.item_tree.column("path", width=470, minwidth=220, stretch=True, anchor=W)
        self.item_tree.grid(row=0, column=0, sticky="nsew")
        item_scroll.grid(row=0, column=1, sticky="ns")
        self.item_tree.config(yscrollcommand=item_scroll.set)
        item_scroll.config(command=self.item_tree.yview)
        self.item_tree.bind("<Button-1>", self.on_item_tree_press)
        self.item_tree.bind("<ButtonRelease-1>", self.on_item_tree_release)
        self.item_tree.bind("<Double-Button-1>", self.on_item_tree_double_click)
        self.item_tree.bind("<Configure>", self.resize_item_columns)

        buttons = ttk.Frame(right)
        buttons.grid(row=3, column=0, sticky="ew")
        for column in range(3):
            buttons.columnconfigure(column, weight=1, uniform="actions")
        ttk.Button(buttons, text="添加文件/软件", command=self.add_files).grid(row=0, column=0, sticky="ew", padx=(0, 4), pady=(0, 6))
        ttk.Button(buttons, text="添加文件夹", command=self.add_folder).grid(row=0, column=1, sticky="ew", padx=4, pady=(0, 6))
        ttk.Button(buttons, text="添加路径/网址", command=self.add_path).grid(row=0, column=2, sticky="ew", padx=(4, 0), pady=(0, 6))
        ttk.Button(buttons, text="设置参数", command=self.set_args).grid(row=1, column=0, sticky="ew", padx=(0, 4))
        self.capture_button = ttk.Button(buttons, text="捕捉当前窗口", command=self.capture_current_window)
        self.capture_button.grid(row=1, column=1, sticky="ew", padx=4)
        self.capture_all_button = ttk.Button(buttons, text="捕捉全部窗口", command=self.capture_all_windows)
        self.capture_all_button.grid(row=1, column=2, sticky="ew", padx=(4, 0))

        status = ttk.Label(self.root, textvariable=self.status, anchor=W, relief="sunken", style="Status.TLabel")
        status.pack(fill=X, side="bottom")

        menu = Menu(self.root)
        self.root.config(menu=menu)
        app_menu = Menu(menu, tearoff=False)
        menu.add_cascade(label="操作", menu=app_menu)
        app_menu.add_command(label="创建一键启动桌面快捷方式", command=self.create_manager_shortcut)
        app_menu.add_command(label="打开程序目录", command=lambda: os.startfile(str(app_dir())))
        menu.add_command(label="关于", command=self.show_about)

        self.build_rules_ui(rules_tab)

    def build_rules_ui(self, parent) -> None:
        main = ttk.Frame(parent, padding=14)
        main.pack(fill=BOTH, expand=True)
        main.columnconfigure(0, weight=0, minsize=210)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)

        left = ttk.Frame(main)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 14))

        ttk.Label(left, text="规则组", anchor=W, style="Section.TLabel").pack(fill=X)
        rule_group_wrap = ttk.Frame(left)
        rule_group_wrap.pack(fill=BOTH, expand=True, pady=(6, 8))
        rule_group_scroll = ttk.Scrollbar(rule_group_wrap, orient=VERTICAL)
        self.rule_group_list = Listbox(
            rule_group_wrap,
            height=5,
            exportselection=False,
            activestyle="dotbox",
            selectmode=SINGLE,
            borderwidth=0,
            highlightthickness=1,
            highlightbackground="#d7dce5",
            highlightcolor="#8aa4d6",
            relief="flat",
            bg="#ffffff",
            selectbackground="#2563eb",
            selectforeground="#ffffff",
            font=("Microsoft YaHei UI", 10),
        )
        self.rule_group_list.pack(side=LEFT, fill=BOTH, expand=True)
        rule_group_scroll.pack(side=RIGHT, fill=Y)
        self.rule_group_list.config(yscrollcommand=rule_group_scroll.set)
        rule_group_scroll.config(command=self.rule_group_list.yview)
        self.rule_group_list.bind("<<ListboxSelect>>", self.on_rule_group_select)

        ttk.Button(left, text="新建", command=self.add_rule_group).pack(fill=X, pady=2)
        ttk.Button(left, text="重命名", command=self.rename_rule_group).pack(fill=X, pady=2)
        ttk.Button(left, text="删除", command=self.delete_rule_group).pack(fill=X, pady=2)
        ttk.Separator(left).pack(fill=X, pady=10)
        ttk.Button(left, text="启用当前规则组", command=self.enable_current_rule_group).pack(fill=X, pady=2)
        ttk.Button(left, text="停用当前规则组", command=self.disable_rule_group).pack(fill=X, pady=2)
        ttk.Button(left, text="停用全部规则组", command=self.disable_all_rule_groups).pack(fill=X, pady=2)

        right = ttk.Frame(main)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(2, weight=1)

        top = ttk.Frame(right)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)
        ttk.Label(top, text="当前规则组：", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        self.rule_group_name = StringVar(value="未选择")
        ttk.Label(top, textvariable=self.rule_group_name, anchor=W).grid(row=0, column=1, sticky="ew")
        self.active_rule_name = StringVar(value="启用规则组：无")
        ttk.Label(top, textvariable=self.active_rule_name, anchor=W).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 0))

        ttk.Label(right, text="不会被捕捉的路径", anchor=W, style="Section.TLabel").grid(row=1, column=0, sticky="ew", pady=(14, 0))
        rule_wrap = ttk.Frame(right)
        rule_wrap.grid(row=2, column=0, sticky="nsew", pady=(6, 8))
        rule_wrap.columnconfigure(0, weight=1)
        rule_wrap.rowconfigure(0, weight=1)
        rule_scroll = ttk.Scrollbar(rule_wrap, orient=VERTICAL)
        self.rule_tree = ttk.Treeview(
            rule_wrap,
            columns=("delete", "name", "path"),
            show="headings",
            selectmode="browse",
            height=5,
        )
        self.rule_tree.heading("delete", text="操作")
        self.rule_tree.heading("name", text="名称")
        self.rule_tree.heading("path", text="路径 / 参数")
        self.rule_tree.column("delete", width=68, minwidth=58, stretch=False, anchor="center")
        self.rule_tree.column("name", width=210, minwidth=110, stretch=True, anchor=W)
        self.rule_tree.column("path", width=470, minwidth=220, stretch=True, anchor=W)
        self.rule_tree.grid(row=0, column=0, sticky="nsew")
        rule_scroll.grid(row=0, column=1, sticky="ns")
        self.rule_tree.config(yscrollcommand=rule_scroll.set)
        rule_scroll.config(command=self.rule_tree.yview)
        self.rule_tree.bind("<Button-1>", self.on_rule_tree_press)
        self.rule_tree.bind("<ButtonRelease-1>", self.on_rule_tree_release)
        self.rule_tree.bind("<Configure>", self.resize_rule_columns)

        buttons = ttk.Frame(right)
        buttons.grid(row=3, column=0, sticky="ew")
        for column in range(3):
            buttons.columnconfigure(column, weight=1, uniform="rule_actions")
        ttk.Button(buttons, text="添加路径/网址", command=self.add_rule_path).grid(row=0, column=0, sticky="ew", padx=(0, 4))
        self.capture_rule_button = ttk.Button(buttons, text="捕捉当前窗口", command=self.capture_current_rule)
        self.capture_rule_button.grid(row=0, column=1, sticky="ew", padx=4)
        self.capture_all_rules_button = ttk.Button(buttons, text="捕捉全部窗口", command=self.capture_all_rules)
        self.capture_all_rules_button.grid(row=0, column=2, sticky="ew", padx=(4, 0))

    def current_group(self) -> dict | None:
        if not self.selected_group_id:
            return None
        return find_group(self.data, self.selected_group_id)

    def current_rule_group(self) -> dict | None:
        if not self.selected_rule_group_id:
            return None
        return find_rule_group(self.data, self.selected_rule_group_id)

    def selected_index(self, listbox: Listbox) -> int | None:
        selection = listbox.curselection()
        return selection[0] if selection else None

    def refresh_groups(self) -> None:
        self.group_list.delete(0, END)
        groups = self.data.get("groups", [])
        for group in groups:
            self.group_list.insert(END, group.get("name", "未命名工作组"))

        if not groups:
            self.selected_group_id = None
            self.group_name.set("未选择")
            self.refresh_items()
            return

        index = 0
        for i, group in enumerate(groups):
            if group.get("id") == self.selected_group_id:
                index = i
                break
        self.group_list.selection_clear(0, END)
        self.group_list.selection_set(index)
        self.group_list.activate(index)
        self.selected_group_id = groups[index].get("id")
        self.group_name.set(groups[index].get("name", "未命名工作组"))
        self.refresh_items()

    def refresh_rule_groups(self) -> None:
        if not hasattr(self, "rule_group_list"):
            return
        self.rule_group_list.delete(0, END)
        groups = self.data.get("rule_groups", [])
        active_ids = set(active_rule_group_ids(self.data))
        for group in groups:
            name = group.get("name", "未命名规则组")
            if group.get("id") in active_ids:
                name += "  [启用]"
            self.rule_group_list.insert(END, name)

        names = active_rule_names(self.data)
        self.active_rule_name.set(f"启用规则组：{', '.join(names)}" if names else "启用规则组：无")

        if not groups:
            self.selected_rule_group_id = None
            self.rule_group_name.set("未选择")
            self.refresh_rules()
            return

        index = 0
        for i, group in enumerate(groups):
            if group.get("id") == self.selected_rule_group_id:
                index = i
                break
        self.rule_group_list.selection_clear(0, END)
        self.rule_group_list.selection_set(index)
        self.rule_group_list.activate(index)
        self.selected_rule_group_id = groups[index].get("id")
        self.rule_group_name.set(groups[index].get("name", "未命名规则组"))
        self.refresh_rules()

    def refresh_rules(self) -> None:
        if not hasattr(self, "rule_tree"):
            return
        for row in self.rule_tree.get_children():
            self.rule_tree.delete(row)
        group = self.current_rule_group()
        if not group:
            return
        for index, rule in enumerate(group.get("rules", [])):
            path = rule.get("path", "")
            args = rule.get("args", "").strip()
            name = Path(path).name or path or "未命名规则"
            path_text = path
            if args:
                path_text += f"    参数：{args}"
            self.rule_tree.insert("", END, iid=str(index), values=("❌", name, path_text))

    def on_rule_group_select(self, _event=None) -> None:
        index = self.selected_index(self.rule_group_list)
        if index is None:
            return
        groups = self.data.get("rule_groups", [])
        if index >= len(groups):
            return
        group = groups[index]
        self.selected_rule_group_id = group.get("id")
        self.rule_group_name.set(group.get("name", "未命名规则组"))
        self.refresh_rules()

    def add_rule_group(self) -> None:
        name = simpledialog.askstring(APP_NAME, "规则组名称：", parent=self.root)
        if not name:
            return
        group = {"id": uuid.uuid4().hex, "name": name.strip(), "rules": []}
        self.data.setdefault("rule_groups", []).append(group)
        self.selected_rule_group_id = group["id"]
        save_config(self.data)
        self.refresh_rule_groups()
        self.status.set(f"已新建规则组：{group['name']}")

    def rename_rule_group(self) -> None:
        group = self.current_rule_group()
        if not group:
            messagebox.showinfo(APP_NAME, "请先选择一个规则组。")
            return
        name = simpledialog.askstring(APP_NAME, "新的规则组名称：", initialvalue=group.get("name", ""), parent=self.root)
        if not name:
            return
        group["name"] = name.strip()
        save_config(self.data)
        self.refresh_rule_groups()
        self.status.set("已重命名规则组。")

    def delete_rule_group(self) -> None:
        group = self.current_rule_group()
        if not group:
            return
        active_ids = [group_id for group_id in active_rule_group_ids(self.data) if group_id != group.get("id")]
        self.data["active_rule_group_ids"] = active_ids
        self.data["rule_groups"] = [g for g in self.data.get("rule_groups", []) if g.get("id") != group.get("id")]
        self.selected_rule_group_id = None
        save_config(self.data)
        self.refresh_rule_groups()
        self.status.set("已删除规则组。")

    def enable_current_rule_group(self) -> None:
        group = self.current_rule_group()
        if not group:
            messagebox.showinfo(APP_NAME, "请先选择一个规则组。")
            return
        active_ids = active_rule_group_ids(self.data)
        if group.get("id") not in active_ids:
            active_ids.append(group.get("id"))
        self.data["active_rule_group_ids"] = active_ids
        save_config(self.data)
        self.refresh_rule_groups()
        self.status.set(f"已启用规则组：{group.get('name', '')}")

    def disable_rule_group(self) -> None:
        group = self.current_rule_group()
        if not group:
            messagebox.showinfo(APP_NAME, "请先选择一个规则组。")
            return
        self.data["active_rule_group_ids"] = [group_id for group_id in active_rule_group_ids(self.data) if group_id != group.get("id")]
        save_config(self.data)
        self.refresh_rule_groups()
        self.status.set(f"已停用规则组：{group.get('name', '')}")

    def disable_all_rule_groups(self) -> None:
        self.data["active_rule_group_ids"] = []
        save_config(self.data)
        self.refresh_rule_groups()
        self.status.set("已停用全部捕捉规则。")

    def refresh_items(self) -> None:
        for row in self.item_tree.get_children():
            self.item_tree.delete(row)
        group = self.current_group()
        if not group:
            return
        for index, item in enumerate(group.get("items", [])):
            path = item.get("path", "")
            args = item.get("args", "").strip()
            name = Path(path).name or path or "未命名项目"
            path_text = path
            if args:
                path_text += f"    参数：{args}"
            self.item_tree.insert("", END, iid=str(index), values=("❌", name, path_text))

    def on_item_tree_press(self, event) -> str | None:
        if self.item_tree.identify_column(event.x) != "#1":
            return None
        return "break"

    def on_item_tree_release(self, event) -> str | None:
        if self.item_tree.identify_column(event.x) != "#1":
            return None
        return self.delete_item_from_event(event)

    def on_item_tree_double_click(self, event) -> str | None:
        if self.item_tree.identify_column(event.x) == "#1":
            return "break"
        self.run_selected_item()
        return "break"

    def delete_item_from_event(self, event) -> str | None:
        row_id = self.item_tree.identify_row(event.y)
        if not row_id:
            return "break"
        try:
            index = int(row_id)
        except ValueError:
            return None
        self.delete_item_at_index(index)
        self.root.update_idletasks()
        self.item_tree.event_generate("<Motion>")
        return "break"

    def resize_item_columns(self, _event=None) -> None:
        total_width = max(self.item_tree.winfo_width(), 360)
        delete_width = 68
        available = max(total_width - delete_width - 24, 260)
        name_width = max(120, min(260, int(available * 0.32)))
        path_width = max(180, available - name_width)
        self.item_tree.column("delete", width=delete_width)
        self.item_tree.column("name", width=name_width)
        self.item_tree.column("path", width=path_width)

    def on_rule_tree_press(self, event) -> str | None:
        if self.rule_tree.identify_column(event.x) != "#1":
            return None
        return "break"

    def on_rule_tree_release(self, event) -> str | None:
        if self.rule_tree.identify_column(event.x) != "#1":
            return None
        row_id = self.rule_tree.identify_row(event.y)
        if not row_id:
            return "break"
        try:
            index = int(row_id)
        except ValueError:
            return None
        self.delete_rule_at_index(index)
        self.root.update_idletasks()
        self.rule_tree.event_generate("<Motion>")
        return "break"

    def resize_rule_columns(self, _event=None) -> None:
        total_width = max(self.rule_tree.winfo_width(), 360)
        delete_width = 68
        available = max(total_width - delete_width - 24, 260)
        name_width = max(120, min(260, int(available * 0.32)))
        path_width = max(180, available - name_width)
        self.rule_tree.column("delete", width=delete_width)
        self.rule_tree.column("name", width=name_width)
        self.rule_tree.column("path", width=path_width)

    def on_group_select(self, _event=None) -> None:
        index = self.selected_index(self.group_list)
        if index is None:
            return
        groups = self.data.get("groups", [])
        if index >= len(groups):
            return
        group = groups[index]
        self.selected_group_id = group.get("id")
        self.group_name.set(group.get("name", "未命名工作组"))
        self.refresh_items()

    def add_group(self) -> None:
        name = simpledialog.askstring(APP_NAME, "工作组名称：", parent=self.root)
        if not name:
            return
        group = {"id": uuid.uuid4().hex, "name": name.strip(), "items": []}
        self.data.setdefault("groups", []).append(group)
        self.selected_group_id = group["id"]
        save_config(self.data)
        self.refresh_groups()
        self.status.set(f"已新建工作组：{group['name']}")

    def rename_group(self) -> None:
        group = self.current_group()
        if not group:
            messagebox.showinfo(APP_NAME, "请先选择一个工作组。")
            return
        name = simpledialog.askstring(APP_NAME, "新的工作组名称：", initialvalue=group.get("name", ""), parent=self.root)
        if not name:
            return
        group["name"] = name.strip()
        save_config(self.data)
        self.refresh_groups()
        self.status.set("已重命名。")

    def delete_group(self) -> None:
        group = self.current_group()
        if not group:
            return
        shortcut_name = f"{safe_shortcut_name(group.get('name', '未命名工作组'))}.lnk"
        shortcut_path = group.get("shortcut_path", "")
        self.data["groups"] = [g for g in self.data.get("groups", []) if g.get("id") != group.get("id")]
        self.selected_group_id = None
        save_config(self.data)
        self.refresh_groups()
        self.status.set("已删除工作组。")
        threading.Thread(target=self.delete_group_shortcut_worker, args=(shortcut_path, shortcut_name), daemon=True).start()

    def delete_group_shortcut_worker(self, shortcut_path: str, shortcut_name: str) -> None:
        try:
            shortcut = Path(shortcut_path) if shortcut_path else get_desktop_dir() / shortcut_name
            trashed = shortcut_trash_dir() / f"{uuid.uuid4().hex}.lnk"
            try:
                os.replace(shortcut, trashed)
            except FileNotFoundError:
                return
            notify_shell_path(shortcut, 0x00000004)
            trashed.unlink(missing_ok=True)
        except Exception as exc:
            self.root.after(0, lambda: messagebox.showwarning(APP_NAME, f"快捷方式删除失败：{exc}"))

    def require_group(self) -> dict | None:
        group = self.current_group()
        if not group:
            messagebox.showinfo(APP_NAME, "请先新建或选择一个工作组。")
        return group

    def require_rule_group(self) -> dict | None:
        group = self.current_rule_group()
        if not group:
            messagebox.showinfo(APP_NAME, "请先新建或选择一个规则组。")
        return group

    def add_files(self) -> None:
        group = self.require_group()
        if not group:
            return
        paths = filedialog.askopenfilenames(title="选择要启动的文件或软件")
        if not paths:
            return
        group.setdefault("items", []).extend({"path": path, "args": "", "enabled": True} for path in paths)
        save_config(self.data)
        self.refresh_items()
        self.status.set(f"已添加 {len(paths)} 个项目。")

    def add_folder(self) -> None:
        group = self.require_group()
        if not group:
            return
        path = filedialog.askdirectory(title="选择要打开的文件夹")
        if not path:
            return
        group.setdefault("items", []).append({"path": path, "args": "", "enabled": True})
        save_config(self.data)
        self.refresh_items()
        self.status.set("已添加文件夹。")

    def add_path(self) -> None:
        group = self.require_group()
        if not group:
            return
        path = simpledialog.askstring(APP_NAME, "输入文件、软件、文件夹路径，或网址：", parent=self.root)
        if not path:
            return
        group.setdefault("items", []).append({"path": path.strip(), "args": "", "enabled": True})
        save_config(self.data)
        self.refresh_items()
        self.status.set("已添加路径。")

    def add_rule_path(self) -> None:
        group = self.require_rule_group()
        if not group:
            return
        path = simpledialog.askstring(APP_NAME, "输入不希望被捕捉的文件、软件、文件夹路径，或网址：", parent=self.root)
        if not path:
            return
        self.add_rules_to_group(group.get("id"), [{"path": path.strip(), "args": "", "enabled": True}])

    def add_rules_to_group(self, group_id: str, rules: list[dict]) -> int:
        group = find_rule_group(self.data, group_id)
        if not group:
            return 0
        existing = {
            (normalize_rule_path(rule.get("path", "")), (rule.get("args", "") or "").strip())
            for rule in group.get("rules", [])
        }
        added = 0
        for rule in rules:
            key = (normalize_rule_path(rule.get("path", "")), (rule.get("args", "") or "").strip())
            if not key[0] or key in existing:
                continue
            group.setdefault("rules", []).append(
                {
                    "path": rule.get("path", ""),
                    "args": rule.get("args", ""),
                    "enabled": rule.get("enabled", True),
                }
            )
            existing.add(key)
            added += 1
        if added:
            self.selected_rule_group_id = group_id
            save_config(self.data)
            self.refresh_rule_groups()
        return added

    def capture_current_rule(self) -> None:
        group = self.require_rule_group()
        if not group:
            return
        group_id = group.get("id")
        self.set_capture_buttons_state("disabled")
        self.status.set("3 秒内切换到要作为规则的窗口。")
        self.root.after(3000, lambda: self.start_rule_capture_worker(group_id))

    def capture_all_rules(self) -> None:
        group = self.require_rule_group()
        if not group:
            return
        group_id = group.get("id")
        self.set_capture_buttons_state("disabled")
        self.status.set("正在捕捉当前桌面上的可见窗口作为规则...")
        thread = threading.Thread(target=self.capture_all_rules_worker, args=(group_id,), daemon=True)
        thread.start()

    def start_rule_capture_worker(self, group_id: str) -> None:
        thread = threading.Thread(target=self.capture_rule_worker, args=(group_id,), daemon=True)
        thread.start()

    def capture_rule_worker(self, group_id: str) -> None:
        try:
            items, statuses = capture_foreground_launch_items()
        except Exception as exc:
            self.root.after(0, lambda: self.finish_rule_capture(group_id, [], str(exc)))
            return
        self.root.after(0, lambda: self.finish_rule_capture(group_id, items, f"已捕捉 {len(statuses)} 条规则候选。"))

    def capture_all_rules_worker(self, group_id: str) -> None:
        try:
            items, statuses = capture_all_visible_launch_items()
        except Exception as exc:
            self.root.after(0, lambda: self.finish_rule_capture(group_id, [], str(exc)))
            return
        self.root.after(0, lambda: self.finish_rule_capture(group_id, items, f"已捕捉 {len(items)} 条规则候选。"))

    def finish_rule_capture(self, group_id: str, rules: list[dict], status: str) -> None:
        self.set_capture_buttons_state("normal")
        if not rules:
            self.status.set("规则捕捉失败。")
            messagebox.showwarning(APP_NAME, status)
            return
        added = self.add_rules_to_group(group_id, rules)
        self.status.set(f"已添加 {added} 条捕捉规则。" if added else "没有新增规则，可能已存在。")

    def delete_rule_at_index(self, index: int) -> None:
        group = self.current_rule_group()
        if not group:
            return
        rules = group.get("rules", [])
        if index < 0 or index >= len(rules):
            return
        del rules[index]
        save_config(self.data)
        self.refresh_rules()
        self.status.set("已删除规则。")

    def capture_current_window(self) -> None:
        group = self.require_group()
        if not group:
            return
        group_id = group.get("id")
        self.set_capture_buttons_state("disabled")
        self.status.set("3 秒内切换到要捕捉的窗口，一键启动会优先保存它正在打开的文件。")
        self.root.after(3000, lambda: self.start_capture_worker(group_id))

    def capture_all_windows(self) -> None:
        group = self.require_group()
        if not group:
            return
        group_id = group.get("id")
        self.set_capture_buttons_state("disabled")
        self.status.set("正在捕捉当前桌面上的可见窗口...")
        thread = threading.Thread(target=self.capture_all_worker, args=(group_id,), daemon=True)
        thread.start()

    def set_capture_buttons_state(self, state: str) -> None:
        self.capture_button.config(state=state)
        self.capture_all_button.config(state=state)
        if hasattr(self, "capture_rule_button"):
            self.capture_rule_button.config(state=state)
        if hasattr(self, "capture_all_rules_button"):
            self.capture_all_rules_button.config(state=state)

    def start_capture_worker(self, group_id: str) -> None:
        thread = threading.Thread(target=self.capture_worker, args=(group_id,), daemon=True)
        thread.start()

    def capture_worker(self, group_id: str) -> None:
        try:
            items, statuses = capture_foreground_launch_items()
        except Exception as exc:
            self.root.after(0, lambda: self.finish_capture(group_id, None, str(exc)))
            return
        if len(items) == 1:
            self.root.after(0, lambda: self.finish_capture(group_id, items[0], statuses[0]))
        else:
            self.root.after(0, lambda: self.finish_capture_all(group_id, items, statuses))

    def capture_all_worker(self, group_id: str) -> None:
        try:
            items, statuses = capture_all_visible_launch_items()
        except Exception as exc:
            self.root.after(0, lambda: self.finish_capture_all(group_id, [], str(exc)))
            return
        self.root.after(0, lambda: self.finish_capture_all(group_id, items, statuses))

    def finish_capture(self, group_id: str, item: dict | None, status: str) -> None:
        self.set_capture_buttons_state("normal")
        if not item:
            self.status.set("捕捉失败。")
            messagebox.showwarning(APP_NAME, status)
            return

        group = find_group(self.data, group_id)
        if not group:
            self.status.set("捕捉成功，但原工作组不存在。")
            return

        if item_blocked_by_active_rules(self.data, item):
            self.status.set("已按启用规则组跳过该捕捉项目。")
            return

        group.setdefault("items", []).append(item)
        self.selected_group_id = group_id
        save_config(self.data)
        self.refresh_groups()
        self.status.set(status)

    def finish_capture_all(self, group_id: str, items: list[dict], status) -> None:
        self.set_capture_buttons_state("normal")
        if not items:
            self.status.set("捕捉失败。")
            messagebox.showwarning(APP_NAME, str(status))
            return

        group = find_group(self.data, group_id)
        if not group:
            self.status.set("捕捉成功，但原工作组不存在。")
            return

        allowed_items = [item for item in items if not item_blocked_by_active_rules(self.data, item)]
        skipped = len(items) - len(allowed_items)
        if not allowed_items:
            self.status.set(f"已按启用规则组跳过 {skipped} 个捕捉项目。")
            return

        group.setdefault("items", []).extend(allowed_items)
        self.selected_group_id = group_id
        save_config(self.data)
        self.refresh_groups()
        suffix = f"，跳过 {skipped} 个" if skipped else ""
        self.status.set(f"已捕捉 {len(allowed_items)} 个窗口项目{suffix}，可按需删除多余项。")

    def selected_item(self) -> tuple[dict, int] | tuple[None, None]:
        group = self.current_group()
        selection = self.item_tree.selection()
        if not group or not selection:
            return None, None
        try:
            index = int(selection[0])
        except ValueError:
            return None, None
        items = group.get("items", [])
        if index >= len(items):
            return None, None
        return items[index], index

    def set_args(self) -> None:
        item, _index = self.selected_item()
        if not item:
            messagebox.showinfo(APP_NAME, "请先选择一个启动项目。")
            return
        args = simpledialog.askstring(APP_NAME, "启动参数（不需要可留空）：", initialvalue=item.get("args", ""), parent=self.root)
        if args is None:
            return
        item["args"] = args.strip()
        save_config(self.data)
        self.refresh_items()
        self.status.set("已保存参数。")

    def delete_item(self) -> None:
        group = self.current_group()
        item, index = self.selected_item()
        if not group or item is None or index is None:
            return
        self.delete_item_at_index(index)

    def delete_item_at_index(self, index: int) -> None:
        group = self.current_group()
        if not group:
            return
        items = group.get("items", [])
        if index < 0 or index >= len(items):
            return
        del group["items"][index]
        save_config(self.data)
        self.refresh_items()
        self.status.set("已删除项目。")

    def clear_current_group_items(self) -> None:
        group = self.current_group()
        if not group:
            messagebox.showinfo(APP_NAME, "请先选择一个工作组。")
            return
        if not group.get("items"):
            self.status.set("当前工作组已经是空的。")
            return
        group["items"] = []
        save_config(self.data)
        self.refresh_items()
        self.status.set("已清空当前工作组。")

    def run_selected_item(self) -> None:
        item, _index = self.selected_item()
        if not item:
            return
        error = launch_item(item)
        if error:
            messagebox.showwarning(APP_NAME, error)

    def run_selected_group(self) -> None:
        group = self.current_group()
        if not group:
            messagebox.showinfo(APP_NAME, "请先选择一个工作组。")
            return
        launch_group(group.get("id"), wait_restore=False)

    def shortcut_args_for_group(self, group: dict) -> str:
        if getattr(sys, "frozen", False):
            return f'--run {group.get("id")}'
        return f'"{app_entry_path()}" --run {group.get("id")}'

    def shortcut_args_for_app(self) -> str:
        if getattr(sys, "frozen", False):
            return ""
        return f'"{app_entry_path()}"'

    def create_group_shortcut(self) -> None:
        group = self.current_group()
        if not group:
            messagebox.showinfo(APP_NAME, "请先选择一个工作组。")
            return
        desktop = get_desktop_dir()
        shortcut = desktop / f"{safe_shortcut_name(group.get('name', '未命名工作组'))}.lnk"
        try:
            create_shortcut(
                shortcut,
                shortcut_target(),
                self.shortcut_args_for_group(group),
                str(app_dir()),
                f"启动工作组：{group.get('name', '')}",
            )
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"创建失败：{exc}")
            return
        group["shortcut_path"] = str(shortcut)
        save_config(self.data)
        notify_shell_path(shortcut, 0x00000002)
        self.status.set(f"已创建桌面快捷方式：{shortcut.name}")

    def create_manager_shortcut(self) -> None:
        desktop = get_desktop_dir()
        shortcut = desktop / f"{APP_NAME}.lnk"
        try:
            create_shortcut(
                shortcut,
                shortcut_target(),
                self.shortcut_args_for_app(),
                str(app_dir()),
                APP_NAME,
            )
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"创建失败：{exc}")
            return
        notify_shell_path(shortcut, 0x00000002)
        self.status.set(f"已创建桌面快捷方式：{shortcut.name}")

    def show_about(self) -> None:
        messagebox.showinfo(APP_NAME, f"{APP_NAME}\n版本：{APP_VERSION}\n作者：时光的星阵")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=f"{APP_NAME} {APP_VERSION}")
    parser.add_argument("--run", help="直接启动指定工作组 ID 或名称")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.run:
        root = Tk()
        root.withdraw()
        code = launch_group(args.run)
        root.destroy()
        return code

    root = Tk()
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    LauncherApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
