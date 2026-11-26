"""仅在交互选择时初始化 Tk；不创建图形主界面或输出文件。"""

from __future__ import annotations

from pathlib import Path

from .paths import application_root


class FileDialogs:
    """在一次选择流程中复用隐藏的 Tk 根，并在离开上下文时销毁。"""

    def __init__(self, app_root: str | Path | None = None) -> None:
        self.app_root = Path(app_root if app_root is not None else application_root())
        self._root = None
        self._filedialog = None

    def __enter__(self) -> FileDialogs:
        if self._root is not None:
            raise RuntimeError("文件选择对话框已打开。")
        # 批处理只导入本模块，不应要求安装 Tk 或具有显示环境。
        import tkinter
        from tkinter import filedialog

        self._filedialog = filedialog
        self._root = tkinter.Tk()
        try:
            self._root.withdraw()
            self._root.update()
        except BaseException:
            self._destroy()
            raise
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        self._destroy()

    def _destroy(self) -> None:
        root = self._root
        self._root = None
        self._filedialog = None
        if root is not None:
            root.destroy()

    def _require_open(self) -> None:
        if self._root is None:
            raise RuntimeError("请在 with FileDialogs() 中选择文件。")

    def select_template(self) -> str:
        self._require_open()
        initial_directory = self.app_root
        for name in ("示例数据_v0.6", "示例数据_v0.5", "示例数据_v0.4", "示例数据_v0.3", "示例数据_v0.2", "示例数据"):
            candidate = self.app_root / name
            if candidate.is_dir():
                initial_directory = candidate
                break
        return self._filedialog.askopenfilename(
            parent=self._root,
            title="选择汇总模板 (.xlsx)",
            initialdir=str(initial_directory),
            filetypes=(("Excel 工作簿", "*.xlsx"),),
        ) or ""

    def select_sources(self, template_path: str | Path) -> tuple[str, ...]:
        self._require_open()
        template_directory = Path(template_path).parent
        initial_directory = template_directory / "正常数据"
        if not initial_directory.is_dir():
            initial_directory = template_directory
        selected = self._filedialog.askopenfilenames(
            parent=self._root,
            title="选择源文件（可多选 .xlsx）",
            initialdir=str(initial_directory),
            filetypes=(("Excel 工作簿", "*.xlsx"),),
        )
        if not selected:
            return ()
        if isinstance(selected, str):
            # 兼容 Tk 未启用对象返回时的 Tcl 路径列表，不能按空格拆分。
            return tuple(self._root.tk.splitlist(selected))
        return tuple(selected)

    def select_output_directory(self) -> str:
        self._require_open()
        return self._filedialog.askdirectory(
            parent=self._root,
            title="选择输出目录",
            initialdir=str(self.app_root),
            mustexist=True,
        ) or ""


def check_dialogs() -> None:
    """打包自检：仅初始化、更新、销毁隐藏根，不弹出选择框。"""
    with FileDialogs():
        pass
