import argparse
import ctypes
from ctypes import wintypes
import json
import os
import re
import subprocess
import sys
import uuid
from pathlib import Path
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
CONFIG_FILE = "launcher_config.json"


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


def load_config() -> dict:
    path = config_path()
    if not path.exists():
        return {"groups": []}
    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return {"groups": []}
    data.setdefault("groups", [])
    return data


def save_config(data: dict) -> None:
    with config_path().open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def is_url(target: str) -> bool:
    return re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*:", target or "") is not None


def expand_target(target: str) -> str:
    return os.path.expandvars(os.path.expanduser(target.strip()))


def launch_item(item: dict) -> str | None:
    target = expand_target(item.get("path", ""))
    if not target:
        return "空路径"

    args = item.get("args", "").strip()
    if not is_url(target) and not Path(target).exists():
        return f"不存在：{target}"

    try:
        if args:
            os.startfile(target, arguments=args)
        else:
            os.startfile(target)
    except Exception as exc:
        return f"{target}：{exc}"
    return None


def find_group(data: dict, key: str) -> dict | None:
    for group in data.get("groups", []):
        if group.get("id") == key or group.get("name") == key:
            return group
    return None


def launch_group(group_key: str, show_done: bool = False) -> int:
    data = load_config()
    group = find_group(data, group_key)
    if not group:
        messagebox.showerror(APP_NAME, f"找不到工作组：{group_key}")
        return 1

    failures = []
    for item in group.get("items", []):
        if item.get("enabled", True):
            error = launch_item(item)
            if error:
                failures.append(error)

    if failures:
        messagebox.showwarning(APP_NAME, "部分项目启动失败：\n\n" + "\n".join(failures))
        return 2

    if show_done:
        messagebox.showinfo(APP_NAME, f"已启动：{group.get('name', '')}")
    return 0


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
        self.root.title(APP_NAME)
        self.root.geometry("860x520")
        self.root.minsize(760, 460)
        self.data = load_config()
        self.selected_group_id = None
        self.status = StringVar(value="配置已自动保存在本机。")

        self.build_ui()
        self.refresh_groups()

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

        main = ttk.Frame(self.root, padding=14)
        main.pack(fill=BOTH, expand=True)

        left = ttk.Frame(main, width=240)
        left.pack(side=LEFT, fill=Y)
        left.pack_propagate(False)

        ttk.Label(left, text="工作组", anchor=W, style="Section.TLabel").pack(fill=X)
        group_wrap = ttk.Frame(left)
        group_wrap.pack(fill=BOTH, expand=True, pady=(6, 8))
        group_scroll = ttk.Scrollbar(group_wrap, orient=VERTICAL)
        self.group_list = Listbox(
            group_wrap,
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
        right.pack(side=RIGHT, fill=BOTH, expand=True, padx=(14, 0))

        top = ttk.Frame(right)
        top.pack(fill=X)
        ttk.Label(top, text="当前工作组：", style="Title.TLabel").pack(side=LEFT)
        self.group_name = StringVar(value="未选择")
        ttk.Label(top, textvariable=self.group_name, anchor=W).pack(side=LEFT, fill=X, expand=True)
        ttk.Button(top, text="运行", command=self.run_selected_group, style="Primary.TButton").pack(side=RIGHT, padx=(8, 0))
        ttk.Button(top, text="创建桌面快捷方式", command=self.create_group_shortcut).pack(side=RIGHT)

        ttk.Label(right, text="启动项目", anchor=W, style="Section.TLabel").pack(fill=X, pady=(14, 0))
        item_wrap = ttk.Frame(right)
        item_wrap.pack(fill=BOTH, expand=True, pady=(6, 8))
        item_scroll = ttk.Scrollbar(item_wrap, orient=VERTICAL)
        self.item_tree = ttk.Treeview(
            item_wrap,
            columns=("name", "path"),
            show="headings",
            selectmode="browse",
        )
        self.item_tree.heading("name", text="名称")
        self.item_tree.heading("path", text="路径 / 参数")
        self.item_tree.column("name", width=210, minwidth=130, stretch=False, anchor=W)
        self.item_tree.column("path", width=470, minwidth=220, stretch=True, anchor=W)
        self.item_tree.pack(side=LEFT, fill=BOTH, expand=True)
        item_scroll.pack(side=RIGHT, fill=Y)
        self.item_tree.config(yscrollcommand=item_scroll.set)
        item_scroll.config(command=self.item_tree.yview)
        self.item_tree.bind("<Double-Button-1>", lambda _event: self.run_selected_item())

        buttons = ttk.Frame(right)
        buttons.pack(fill=X)
        ttk.Button(buttons, text="添加文件/软件", command=self.add_files).pack(side=LEFT, padx=(0, 6))
        ttk.Button(buttons, text="添加文件夹", command=self.add_folder).pack(side=LEFT, padx=6)
        ttk.Button(buttons, text="添加路径/网址", command=self.add_path).pack(side=LEFT, padx=6)
        ttk.Button(buttons, text="设置参数", command=self.set_args).pack(side=LEFT, padx=6)
        ttk.Button(buttons, text="删除项目", command=self.delete_item).pack(side=LEFT, padx=6)

        status = ttk.Label(self.root, textvariable=self.status, anchor=W, relief="sunken", style="Status.TLabel")
        status.pack(fill=X, side="bottom")

        menu = Menu(self.root)
        self.root.config(menu=menu)
        app_menu = Menu(menu, tearoff=False)
        menu.add_cascade(label="操作", menu=app_menu)
        app_menu.add_command(label="创建一键启动桌面快捷方式", command=self.create_manager_shortcut)
        app_menu.add_command(label="打开程序目录", command=lambda: os.startfile(str(app_dir())))

    def current_group(self) -> dict | None:
        if not self.selected_group_id:
            return None
        return find_group(self.data, self.selected_group_id)

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
            self.item_tree.insert("", END, iid=str(index), values=(name, path_text))

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
        if not messagebox.askyesno(APP_NAME, f"删除工作组“{group.get('name', '')}”？"):
            return
        self.data["groups"] = [g for g in self.data.get("groups", []) if g.get("id") != group.get("id")]
        self.selected_group_id = None
        save_config(self.data)
        self.refresh_groups()
        self.status.set("已删除工作组。")

    def require_group(self) -> dict | None:
        group = self.current_group()
        if not group:
            messagebox.showinfo(APP_NAME, "请先新建或选择一个工作组。")
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
        del group["items"][index]
        save_config(self.data)
        self.refresh_items()
        self.status.set("已删除项目。")

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
        launch_group(group.get("id"), show_done=True)

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
        if shortcut.exists() and not messagebox.askyesno(APP_NAME, f"快捷方式“{shortcut.name}”已存在，覆盖吗？"):
            return
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
        self.status.set(f"已创建桌面快捷方式：{shortcut.name}")

    def create_manager_shortcut(self) -> None:
        desktop = get_desktop_dir()
        shortcut = desktop / f"{APP_NAME}.lnk"
        if shortcut.exists() and not messagebox.askyesno(APP_NAME, f"快捷方式“{shortcut.name}”已存在，覆盖吗？"):
            return
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
        self.status.set(f"已创建桌面快捷方式：{shortcut.name}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=APP_NAME)
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
