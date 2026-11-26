import argparse
from pathlib import Path
import shutil
import subprocess
import sys
from zipfile import ZIP_DEFLATED, ZipFile

from summary_tool import APP_NAME, __version__


def main():
    root = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description="打包精简 CLI 与标准文件选择框，不含 pandas/NumPy。")
    parser.add_argument("--dist-dir", type=Path, default=root / "dist")
    args = parser.parse_args()
    dist = args.dist_dir.resolve()
    name = f"{APP_NAME}_v{__version__}"
    executable = dist / f"{name}.exe"
    archive_path = dist / f"{APP_NAME}.zip"
    from summary_tool.configuration import ConfigurationError, load_configuration
    config = root / "config.xlsx"
    try:
        load_configuration(config)
    except ConfigurationError as exc:
        parser.error(str(exc))
    manual = root / "使用说明.docx"
    if not manual.is_file():
        parser.error("缺少使用说明.docx，请将Word使用说明放在项目根目录。")

    from PyInstaller.utils.win32.versioninfo import (
        FixedFileInfo, StringFileInfo, StringStruct, StringTable,
        VarFileInfo, VarStruct, VSVersionInfo,
    )
    build_dir = root / ".build"
    build_dir.mkdir(exist_ok=True)
    version_path = build_dir / "version_info.txt"
    version_number = tuple(map(int, __version__.split("."))) + (0,)
    version_info = VSVersionInfo(
        ffi=FixedFileInfo(filevers=version_number, prodvers=version_number),
        kids=[
            StringFileInfo([StringTable("080404B0", [
                StringStruct("ProductName", APP_NAME),
                StringStruct("FileDescription", APP_NAME),
                StringStruct("FileVersion", __version__),
                StringStruct("ProductVersion", __version__),
                StringStruct("OriginalFilename", executable.name),
            ])]),
            VarFileInfo([VarStruct("Translation", [0x0804, 1200])]),
        ],
    )
    version_path.write_text(str(version_info), encoding="utf-8")
    command = [
        sys.executable, "-m", "PyInstaller", "--onefile", "--console", "--noupx", "--noconfirm",
        "--name", name, "--distpath", str(dist),
        "--workpath", str(build_dir / "pyinstaller"),
        "--specpath", str(build_dir / "spec"),
        "--version-file", str(version_path),
        "--collect-all", "python_calamine", "--hidden-import", "tkinter.filedialog",
        "--exclude-module", "pandas", "--exclude-module", "numpy",
        "--exclude-module", "matplotlib", "--exclude-module", "PIL", str(root / "app.py"),
    ]
    subprocess.run(command, cwd=root, check=True)
    shutil.copy2(config, dist / "config.xlsx")
    load_configuration(dist / "config.xlsx")
    shutil.copy2(manual, dist / manual.name)
    payload = (executable, dist / "config.xlsx", dist / manual.name)
    temporary_archive = archive_path.with_suffix(".tmp.zip")
    try:
        with ZipFile(temporary_archive, "w", compression=ZIP_DEFLATED, compresslevel=6) as archive:
            for path in payload:
                archive.write(path, path.name)
        temporary_archive.replace(archive_path)
    finally:
        temporary_archive.unlink(missing_ok=True)
    (dist / "使用说明.md").unlink(missing_ok=True)
    print(f"程序：{executable}（{executable.stat().st_size / 1024 / 1024:.1f} MiB）")
    print(f"压缩包：{archive_path}（{archive_path.stat().st_size / 1024 / 1024:.1f} MiB）")


if __name__ == "__main__":
    main()
