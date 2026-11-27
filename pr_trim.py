# -*- coding: utf-8 -*-
"""
PR瘦身 v2.0 — Premiere Pro 缓存清理与项目优化工具
=====================================================
解决 Premiere Pro 在高端硬件（RTX 4070 + 32GB RAM）上仍然卡顿的痛点。
根本原因：媒体缓存堆积、预览文件膨胀、项目碎片化。

四大功能模块：
  1. 缓存扫描 — 自动检测 PR 媒体缓存目录，按类型统计文件数与总体积
  2. 项目文件分析 — 解析 .prproj 项目文件，检测断链（丢失）媒体文件
  3. 大文件扫描 — 递归扫描项目文件夹，找出超大视频文件并可归档
  4. 一键优化 — 批量清理选中的缓存，保留最近 5 个自动保存

v2.0 新增：安全的 dry-run 预览、跨平台回收站、详细报告（JSON/文本）、CLI 模式与彩蛋。

依赖：仅使用 Python 标准库（os / shutil / xml / tkinter / datetime / pathlib / subprocess / gzip）。
未使用任何第三方包，无需 pip install。
更新：2026-11-27 (v2.0)
"""

import argparse
import json
import os
import re
import shutil
import gzip
import sys
import platform
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime
from collections import defaultdict
import subprocess
import threading

# v2.0：无图形环境时 CLI 仍可用
try:
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog
except ImportError:
    tk = None


# ============================================================================
# 全局常量 & 配置
# ============================================================================

# 窗口配色方案（深色主题 + 青色强调）
COLOR_BG = "#1e1e2e"             # 主背景深色
COLOR_CARD = "#2a2a3c"           # 卡片/子面板背景
COLOR_ACCENT = "#06b6d4"         # 青色强调色（按钮、标题）
COLOR_ACCENT_HOVER = "#0891b2"   # 青色深色变体（悬停）
COLOR_TEXT = "#e0e0e0"           # 正文文本色
COLOR_TEXT_SECONDARY = "#a0a0b0" # 次要文本色
COLOR_GREEN = "#4ade80"          # 绿色 — <100MB
COLOR_ORANGE = "#fb923c"         # 橙色 — 100MB~1GB
COLOR_RED = "#f87171"            # 红色 — >1GB
COLOR_ENTRY_BG = "#3a3a4c"       # 输入框背景
COLOR_SUCCESS_BG = "#14532d"     # 深绿色背景（成功标记）
COLOR_DANGER_BG = "#7f1d1d"     # 深红色背景（错误标记）

# 缓存目录常见路径模板
PR_CACHE_PATHS = {
    "media_cache_db": [
        # Adobe 公共媒体缓存数据库目录
        lambda: Path(os.environ.get("LOCALAPPDATA", "")) / "Adobe" / "Common" / "Media Cache",
    ],
    "media_cache_files": [
        # Adobe 公共媒体缓存文件目录（.cfa / .pek）
        lambda: Path(os.environ.get("LOCALAPPDATA", "")) / "Adobe" / "Common" / "Media Cache Files",
    ],
    "auto_save": [
        # Premiere Pro 自动保存目录（版本号子目录）
        lambda: Path.home() / "Documents" / "Adobe" / "Premiere Pro",
    ],
}


# ============================================================================
# 工具函数
# ============================================================================

def format_size(size_bytes: int) -> str:
    """将字节数格式化为人类可读的字符串"""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.2f} GB"


def get_size_color(size_bytes: int) -> str:
    """根据文件大小返回对应颜色（<100MB 绿 / 100MB~1GB 橙 / >1GB 红）"""
    mb = size_bytes / (1024 * 1024)
    if mb < 100:
        return COLOR_GREEN
    elif mb < 1024:
        return COLOR_ORANGE
    else:
        return COLOR_RED


def detect_platform() -> str:
    """v2.0：返回 'windows' / 'macos' / 'linux'，用于多平台分支"""
    system = platform.system().lower()
    if system == 'darwin':
        return 'macos'
    if system == 'windows':
        return 'windows'
    return 'linux'


def send_to_recycle_bin(file_path: str) -> bool:
    """
    v2.0：跨平台将文件/目录移入回收站。
      - Windows：PowerShell + Microsoft.VisualBasic.FileIO
      - macOS：osascript 告知 Finder 删除
      - Linux：gio trash / trash-put，均无则移动到 ~/.local/share/Trash/files
    返回 True 表示成功，False 表示失败。
    """
    if not (os.path.isfile(file_path) or os.path.isdir(file_path)):
        return False
    current = detect_platform()

    try:
        if current == 'windows':
            method = "DeleteFile" if os.path.isfile(file_path) else "DeleteDirectory"
            ps_script = (
                f"Add-Type -AssemblyName Microsoft.VisualBasic; "
                f"[Microsoft.VisualBasic.FileIO.FileSystem]::{method}("
                f"'{file_path}', 'OnlyErrorDialogue', 'SendToRecycleBin')"
            )
            result = subprocess.run(
                ["powershell", "-NoProfile", "-Command", ps_script],
                capture_output=True, text=True, timeout=30)
            return result.returncode == 0

        if current == 'macos':
            script = f'tell application "Finder" to delete POSIX file "{file_path}"'
            result = subprocess.run(["osascript", "-e", script],
                                    capture_output=True, text=True, timeout=30)
            return result.returncode == 0

        # Linux：优先使用标准 trash 命令
        for cmd in (["gio", "trash", file_path], ["trash-put", file_path]):
            if shutil.which(cmd[0]):
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
                if result.returncode == 0:
                    return True
        # 兜底：移动到 XDG 回收站目录
        trash_dir = Path(os.environ.get(
            "XDG_DATA_HOME", str(Path.home() / ".local" / "share"))) / "Trash" / "files"
        trash_dir.mkdir(parents=True, exist_ok=True)
        src = Path(file_path)
        dest = trash_dir / src.name
        counter = 1
        while dest.exists():
            dest = trash_dir / f"{src.name}.{counter}"
            counter += 1
        shutil.move(str(src), str(dest))
        return True
    except Exception:
        return False


def safe_path(path_resolver) -> Path | None:
    """安全解析路径：捕获异常，返回 Path 或 None"""
    try:
        p = path_resolver()
        if isinstance(p, Path) and p.exists():
            return p
    except Exception:
        pass
    return None


def scan_directory_recursive(
    directory: Path,
    extensions: set | None = None,
) -> tuple[list[Path], int]:
    """
    递归扫描目录，返回 (文件路径列表, 总大小字节数)。
    可选参数 extensions 用于过滤文件后缀（如 {'.cfa', '.pek'}）。
    """
    files = []
    total_size = 0
    if not directory or not directory.exists():
        return files, total_size
    try:
        for entry in directory.rglob("*"):
            if entry.is_file():
                # 后缀过滤
                if extensions and entry.suffix.lower() not in extensions:
                    continue
                try:
                    size = entry.stat().st_size
                except OSError:
                    size = 0
                files.append(entry)
                total_size += size
    except PermissionError:
        pass
    return files, total_size


def extract_media_paths_from_prproj(prproj_path: Path) -> list[str]:
    """
    解析 .prproj 项目文件（gzip 压缩的 XML），提取所有引用的媒体文件路径。
    Premiere Pro 项目 XML 中的路径可能出现在以下位置：
      - <pathurl>file:///D:/path/to/file.mp4</pathurl>
      - ActualMediaFilePath 元素
      - path 属性
    返回去重后的绝对路径列表。
    """
    paths_set = set()
    try:
        # .prproj 文件是 gzip 压缩的 XML
        with gzip.open(prproj_path, "rb") as f:
            raw_xml = f.read()

        # 尝试解析 XML
        try:
            root = ET.fromstring(raw_xml)
        except ET.ParseError:
            # 某些版本的 .prproj 可能在末尾有额外数据，尝试截断
            # 查找最后一个有效标签
            return list(paths_set)

        # 命名空间（PR 使用 Adobe 命名空间）
        ns_prefix = ""
        if "}" in root.tag:
            ns_prefix = root.tag.split("}")[0] + "}"

        # 遍历所有元素，提取路径信息
        for elem in root.iter():
            tag_local = elem.tag.replace(ns_prefix, "") if ns_prefix else elem.tag

            # 检查元素的文本内容（如 <pathurl>file:///...</pathurl>）
            if tag_local.lower() in ("pathurl", "actualmediafilepath", "fullpath"):
                text = (elem.text or "").strip()
                if text:
                    normalized = _normalize_path(text)
                    if normalized:
                        paths_set.add(normalized)

            # 检查属性中的路径
            for attr_name in ("pathurl", "path", "filepath", "sourcepath"):
                attr_val = elem.get(attr_name, "").strip()
                if attr_val:
                    normalized = _normalize_path(attr_val)
                    if normalized:
                        paths_set.add(normalized)

        # 如果以上没有提取到路径，用正则表达式从原始 XML 中提取
        if not paths_set:
            # 匹配 file:/// 格式的 URI
            uri_matches = re.findall(
                r'file:///([^"<>\s]+?)(?:"|<|>|\s|$)', raw_xml.decode("utf-8", errors="ignore")
            )
            for m in uri_matches:
                normalized = _normalize_path(m)
                if normalized:
                    paths_set.add(normalized)

            # 匹配 Windows 绝对路径（如 D:\path\to\file.mp4）
            path_matches = re.findall(
                r'([A-Za-z]:\\[^"<>\n\r]+?\.\w{2,5})', raw_xml.decode("utf-8", errors="ignore")
            )
            for m in path_matches:
                normalized = _normalize_path(m)
                if normalized:
                    paths_set.add(normalized)

    except (gzip.BadGzipFile, OSError, Exception):
        pass

    return sorted(paths_set)


def _normalize_path(raw: str) -> str | None:
    """
    标准化路径字符串：
    - 去除 file:// 前缀
    - 将正斜杠转为反斜杠
    - 过滤明显非文件路径的字符串
    """
    text = raw.strip()
    if not text:
        return None
    # 移除 file:// 前缀
    if text.startswith("file:///"):
        text = text[8:]
    elif text.startswith("file://"):
        text = text[7:]
    # 统一为反斜杠
    text = text.replace("/", "\\")
    # URL 解码（处理 %20 等）
    try:
        from urllib.parse import unquote
        text = unquote(text)
    except Exception:
        pass
    # 过滤：必须是 Windows 绝对路径（含盘符）
    if re.match(r'^[A-Za-z]:\\', text):
        # 过滤明显不是媒体文件的扩展名
        return text
    return None


def estimate_project_complexity(prproj_path: Path) -> str:
    """
    根据 .prproj 文件大小估算项目复杂度。
      文件大小 < 1MB   → 简单项目
      1MB ~ 5MB        → 中等项目
      5MB ~ 20MB       → 复杂项目
      > 20MB           → 大型项目
    """
    try:
        size_mb = prproj_path.stat().st_size / (1024 * 1024)
    except OSError:
        return "未知"
    if size_mb < 1:
        return "简单项目"
    elif size_mb < 5:
        return "中等项目"
    elif size_mb < 20:
        return "复杂项目"
    else:
        return "大型项目"


# ============================================================================
# v2.0 新增：安全 dry-run / 清理计划 / 详细报告
# ============================================================================

# 绝不允许清理的路径片段（防止误删系统/用户目录）
FORBIDDEN_PATH_PARTS = ('\\windows\\', '/windows/', '\\system32\\', '/system/',
                        '\\program files\\', '/etc/', '/usr/', '/home/', '/root/')


def is_safe_to_delete(path: Path, allowed_roots: list[Path] | None = None) -> tuple[bool, str]:
    """
    v2.0：删除前的安全校验，返回 (是否安全, 原因)。
    规则：
      1. 路径必须存在且是普通文件
      2. 拒绝系统关键目录
      3. 拒绝隐藏的符号链接（避免跟随链接删到别处）
      4. 若提供 allowed_roots，路径必须位于其中之一
    """
    try:
        path = Path(path).resolve()
    except OSError as e:
        return False, f"无法解析路径: {e}"
    if not path.exists():
        return False, "文件不存在"
    if not path.is_file():
        return False, "不是普通文件"
    if path.is_symlink():
        return False, "是符号链接，已跳过"
    lowered = str(path).lower().replace('/', '\\')
    if any(part.replace('/', '\\') in lowered for part in FORBIDDEN_PATH_PARTS):
        return False, "位于系统关键目录，已跳过"
    if allowed_roots:
        roots = [r.resolve() for r in allowed_roots if r.exists()]
        if roots and not any(str(path).lower().startswith(str(root).lower()) for root in roots):
            return False, "不在允许的缓存目录内，已跳过"
    return True, ""


def build_cleanup_plan(scan_results: dict, categories: list[tuple],
                       selected_keys: set[str] | None = None,
                       keep_recent: int = 5) -> dict:
    """
    v2.0：根据缓存扫描结果构建清理计划（纯函数，不删除任何文件）。
    返回 dict：{key: {'files': [...], 'kept': [...], 'skipped': [...], 'size': int}}
    """
    plan = {}
    for key, display_name, _ in categories:
        if selected_keys is not None and key not in selected_keys:
            continue
        data = scan_results.get(key, {"files": [], "size": 0})
        files = list(data.get("files", []))
        entry = {'name': display_name, 'files': [], 'kept': [], 'skipped': [], 'size': 0}

        if key == 'auto_save' and keep_recent > 0:
            files.sort(key=lambda f: f.stat().st_mtime if f.exists() else 0, reverse=True)
            entry['kept'] = files[:keep_recent]
            candidates = files[keep_recent:]
        else:
            candidates = files

        for f in candidates:
            safe, reason = is_safe_to_delete(f)
            if safe:
                entry['files'].append(f)
                try:
                    entry['size'] += f.stat().st_size
                except OSError:
                    pass
            else:
                entry['skipped'].append({'path': str(f), 'reason': reason})
        plan[key] = entry
    return plan


def plan_summary(plan: dict) -> tuple[int, int]:
    """统计清理计划的文件数与总字节数"""
    count = sum(len(e['files']) for e in plan.values())
    size = sum(e['size'] for e in plan.values())
    return count, size


def format_cleanup_plan(plan: dict, dry_run: bool = True) -> str:
    """v2.0：生成人类可读的清理计划文本（dry-run 与确认框共用）"""
    count, size = plan_summary(plan)
    title = "===== 干跑预览（未删除任何文件） =====" if dry_run else "===== 即将执行 ====="
    lines = [title, ""]
    for key, entry in plan.items():
        lines.append(f"[{entry['name']}]")
        lines.append(f"  待清理：{len(entry['files'])} 个（{format_size(entry['size'])}）")
        if entry['kept']:
            lines.append(f"  保留：{len(entry['kept'])} 个最新文件")
        if entry['skipped']:
            lines.append(f"  安全跳过：{len(entry['skipped'])} 个")
            for item in entry['skipped'][:5]:
                lines.append(f"    - {Path(item['path']).name}: {item['reason']}")
        lines.append("")
    lines.append(f"总计：{count} 个文件，约 {format_size(size)}")
    if dry_run:
        lines.append("提示：这是 dry-run 预览，磁盘未做任何更改。")
    else:
        lines.append("所有删除均移入回收站，可随时恢复。")
    return '\n'.join(lines)


def execute_cleanup_plan(plan: dict, dry_run: bool = True) -> dict:
    """
    v2.0：执行清理计划。dry_run=True 时只生成详细报告，绝不删除文件。
    返回详细报告 dict，可导出 JSON 或显示到 GUI。
    """
    report = {
        'tool': 'PRTrim', 'version': '2.0',
        'generated_at': datetime.now().isoformat(timespec='seconds'),
        'platform': detect_platform(),
        'dry_run': dry_run,
        'categories': [], 'deleted_files': 0, 'deleted_size': 0,
        'failed': [], 'skipped': [],
    }
    for key, entry in plan.items():
        cat = {'key': key, 'name': entry['name'], 'planned': len(entry['files']),
               'planned_size': entry['size'], 'deleted': 0, 'failed': [], 'skipped': entry['skipped']}
        report['skipped'].extend(entry['skipped'])
        if not dry_run:
            for f in entry['files']:
                try:
                    if send_to_recycle_bin(str(f)):
                        cat['deleted'] += 1
                        report['deleted_files'] += 1
                        try:
                            report['deleted_size'] += f.stat().st_size
                        except OSError:
                            pass
                    else:
                        cat['failed'].append(str(f))
                        report['failed'].append(str(f))
                except Exception as e:
                    cat['failed'].append(f"{f}: {e}")
                    report['failed'].append(f"{f}: {e}")
        report['categories'].append(cat)
    report['planned_files'], report['planned_size'] = plan_summary(plan)
    return report


def format_cleanup_report(report: dict) -> str:
    """v2.0：格式化详细清理报告（供 GUI 与 CLI 复用）"""
    mode = "干跑预览" if report['dry_run'] else "清理报告"
    lines = [f"===== PR瘦身 v2.0 {mode} =====", ""]
    lines.append(f"时间：{report['generated_at']}    平台：{report['platform']}")
    lines.append(f"计划清理：{report['planned_files']} 个（{format_size(report['planned_size'])}）")
    if not report['dry_run']:
        lines.append(f"实际清理：{report['deleted_files']} 个（{format_size(report['deleted_size'])}）")
    lines.append("")
    for cat in report['categories']:
        lines.append(f"[{cat['name']}] 计划 {cat['planned']} 个 / "
                     f"{format_size(cat['planned_size'])}" +
                     (f"，实际 {cat['deleted']} 个" if not report['dry_run'] else ""))
        if cat['skipped']:
            lines.append(f"  安全跳过 {len(cat['skipped'])} 个")
        if cat['failed']:
            lines.append(f"  失败 {len(cat['failed'])} 个")
            for f in cat['failed'][:10]:
                lines.append(f"    - {f}")
    if report['dry_run']:
        lines.append("\n本次为 dry-run，未对磁盘做任何更改。")
    elif report['failed']:
        lines.append(f"\n失败 {len(report['failed'])} 个，请检查上方的路径。")
    else:
        lines.append("\n全部完成！文件已移入回收站，可随时恢复。")
    return '\n'.join(lines)


def export_cleanup_report(report: dict, output_path: str) -> str:
    """v2.0：导出详细报告为 JSON"""
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    # Path 对象不可直接 JSON 序列化，统一转换为字符串
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(report, f, ensure_ascii=False, indent=2, default=str)
    return output_path


def print_easter_egg():
    """v2.0：--easter 隐藏彩蛋 —— SOLCAT ASCII 署名"""
    art = r"""
  ____             _   _       _   _
 / ___|  ___  _ __| |_(_) __ _| |_| |__
 \___ \ / _ \| '__| __| |/ _` | __| '_ \
  ___) | (_) | |  | |_| | (_| | |_| | |
 |____/ \___/|_|   \__|_|\__,_|\__|_| |_|
"""
    print(art)
    print("  SOLCAT EASTER EGG PROTOCOL v2.0 — 你找到了隐藏入口")
    print("  提示：↑↑↓↓←→←→BA · 连点标题 5 次 · 输入 solcat · Ctrl+Shift+T")
    print("  清理彩蛋：保留最近 5 个自动保存——第 42 个才是真正的秘密 ;)")


# ============================================================================
# 主应用程序类
# ============================================================================

class PRTrimApp:
    """PR瘦身 v2.0 主窗口应用程序"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PR瘦身 v2.0")
        self.root.geometry("540x540")
        self.root.resizable(False, False)
        self.root.configure(bg=COLOR_BG)

        # ---- 自定义 ttk 样式 ----
        self._setup_styles()

        # ---- 数据存储 ----
        # 缓存扫描结果 {类别名: {"files": [Path,...], "size": int}}
        self.cache_scan_results: dict = {}
        # 各类别复选框变量 {类别名: tk.BooleanVar}
        self.cache_checkbox_vars: dict = {}
        # 项目分析结果
        self.broken_links: list[str] = []  # 断链文件路径列表
        # 大文件扫描结果
        self.large_files: list[dict] = []  # [{"path": Path, "size": int, "mtime": float}, ...]

        # ---- 构建 UI ----
        self._build_scrollable_ui()

    # ------------------------------------------------------------------
    # 样式设置
    # ------------------------------------------------------------------

    def _setup_styles(self):
        """配置 ttk 控件样式以匹配深色主题"""
        style = ttk.Style()
        style.theme_use("clam")

        # 通用配置
        style.configure(".", background=COLOR_BG, foreground=COLOR_TEXT, font=("微软雅黑", 9))
        style.configure("TLabel", background=COLOR_BG, foreground=COLOR_TEXT)
        style.configure("TCheckbutton", background=COLOR_BG, foreground=COLOR_TEXT)
        style.map("TCheckbutton", background=[("active", COLOR_BG)])

        # 按钮样式
        style.configure(
            "Accent.TButton",
            background=COLOR_ACCENT,
            foreground="#ffffff",
            borderwidth=0,
            font=("微软雅黑", 9, "bold"),
        )
        style.map(
            "Accent.TButton",
            background=[("active", COLOR_ACCENT_HOVER), ("disabled", "#555566")],
            foreground=[("disabled", "#888899")],
        )

        style.configure(
            "Secondary.TButton",
            background="#444466",
            foreground=COLOR_TEXT,
            borderwidth=0,
            font=("微软雅黑", 9),
        )
        style.map(
            "Secondary.TButton",
            background=[("active", "#555577"), ("disabled", "#3a3a4c")],
        )

        # 标签框样式
        style.configure("Card.TLabelframe", background=COLOR_CARD, foreground=COLOR_ACCENT)
        style.configure(
            "Card.TLabelframe.Label",
            background=COLOR_CARD,
            foreground=COLOR_ACCENT,
            font=("微软雅黑", 10, "bold"),
        )

        # 大清理按钮
        style.configure(
            "BigAccent.TButton",
            background=COLOR_ACCENT,
            foreground="#ffffff",
            borderwidth=0,
            font=("微软雅黑", 12, "bold"),
            padding=(20, 10),
        )

    # ------------------------------------------------------------------
    # 滚动 UI 构建
    # ------------------------------------------------------------------

    def _build_scrollable_ui(self):
        """构建可滚动的主界面容器"""
        # 画布 + 滚动条
        canvas = tk.Canvas(self.root, bg=COLOR_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.root, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 主内容框架（放置于画布内）
        self.main_frame = tk.Frame(canvas, bg=COLOR_BG)
        canvas_window = canvas.create_window((0, 0), window=self.main_frame, anchor=tk.NW)

        # 绑定滚动事件
        def _on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        self.main_frame.bind("<Configure>", _on_frame_configure)

        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.bind("<Configure>", _on_canvas_configure)

        # 鼠标滚轮支持
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel, add="+")

        # 构建四个功能区域
        self._build_area_cache_scan()
        self._build_area_project_analysis()
        self._build_area_large_files()
        self._build_area_cleanup()

    # ==================================================================
    # 区域一：缓存扫描
    # ==================================================================

    def _build_area_cache_scan(self):
        """构建缓存扫描区域"""
        frame = ttk.LabelFrame(self.main_frame, text="区域一：缓存扫描", style="Card.TLabelframe")
        frame.pack(fill=tk.X, padx=8, pady=(8, 4))

        # 目录信息行
        info_frame = tk.Frame(frame, bg=COLOR_CARD)
        info_frame.pack(fill=tk.X, padx=6, pady=(4, 2))
        tk.Label(
            info_frame,
            text="自动检测 Adobe 公共媒体缓存目录及 PR 自动保存目录",
            bg=COLOR_CARD,
            fg=COLOR_TEXT_SECONDARY,
            font=("微软雅黑", 8),
        ).pack(anchor=tk.W)

        # 扫描按钮行
        btn_frame = tk.Frame(frame, bg=COLOR_CARD)
        btn_frame.pack(fill=tk.X, padx=6, pady=(2, 6))

        self.btn_scan_cache = ttk.Button(
            btn_frame,
            text="扫描缓存",
            command=self._on_scan_cache,
            style="Accent.TButton",
        )
        self.btn_scan_cache.pack(side=tk.LEFT, padx=(0, 8))

        self.lbl_scan_status = tk.Label(
            btn_frame,
            text="就绪",
            bg=COLOR_CARD,
            fg=COLOR_TEXT_SECONDARY,
            font=("微软雅黑", 8),
        )
        self.lbl_scan_status.pack(side=tk.LEFT)

        # 扫描结果行（4 个类别，每个一行）
        self.result_frame = tk.Frame(frame, bg=COLOR_CARD)
        self.result_frame.pack(fill=tk.X, padx=6, pady=(0, 6))

        # 类别定义：(内部key, 显示名称, 说明)
        self.cache_categories = [
            ("media_files", "媒体缓存文件", ".cfa 音频整合缓存"),
            ("media_db", "媒体缓存数据库", "媒体缓存索引数据库"),
            ("auto_save", "自动保存文件", ".prproj 自动保存"),
            ("peak_files", "峰值渲染文件", ".pek 波形峰值缓存"),
        ]

        for key, display_name, desc in self.cache_categories:
            cat_frame = tk.Frame(self.result_frame, bg=COLOR_CARD)
            cat_frame.pack(fill=tk.X, pady=1)

            var = tk.BooleanVar(value=True)
            self.cache_checkbox_vars[key] = var
            cb = tk.Checkbutton(
                cat_frame,
                variable=var,
                bg=COLOR_CARD,
                fg=COLOR_TEXT,
                selectcolor=COLOR_CARD,
                activebackground=COLOR_CARD,
                activeforeground=COLOR_TEXT,
                font=("微软雅黑", 9),
            )
            cb.pack(side=tk.LEFT)

            # 标签使用 Label 方便后续更新颜色
            lbl = tk.Label(
                cat_frame,
                text=f"{display_name}（{desc}）",
                bg=COLOR_CARD,
                fg=COLOR_TEXT_SECONDARY,
                font=("微软雅黑", 8),
                anchor=tk.W,
            )
            lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

            setattr(self, f"lbl_{key}", lbl)
            setattr(self, f"cb_{key}", cb)

    # ------------------------------------------------------------------
    # 缓存扫描逻辑
    # ------------------------------------------------------------------

    def _on_scan_cache(self):
        """点击「扫描缓存」按钮：启动后台线程执行扫描"""
        self.btn_scan_cache.configure(state="disabled", text="扫描中...")
        self.lbl_scan_status.configure(text="正在扫描缓存目录...", fg=COLOR_ORANGE)
        # 清空旧结果
        for key, _, _ in self.cache_categories:
            lbl = getattr(self, f"lbl_{key}", None)
            if lbl:
                lbl.configure(text="扫描中...", fg=COLOR_TEXT_SECONDARY)

        thread = threading.Thread(target=self._scan_cache_thread, daemon=True)
        thread.start()

    def _scan_cache_thread(self):
        """后台线程：递归扫描 PR 缓存目录"""
        results = {}

        # ----- 媒体缓存文件 (.cfa) -----
        mc_files_dir = safe_path(PR_CACHE_PATHS["media_cache_files"][0])
        if mc_files_dir:
            cfa_files, cfa_size = scan_directory_recursive(mc_files_dir, {".cfa"})
            results["media_files"] = {"files": cfa_files, "size": cfa_size}
        else:
            results["media_files"] = {"files": [], "size": 0}

        # ----- 峰值渲染文件 (.pek) -----
        if mc_files_dir:
            pek_files, pek_size = scan_directory_recursive(mc_files_dir, {".pek"})
            results["peak_files"] = {"files": pek_files, "size": pek_size}
        else:
            results["peak_files"] = {"files": [], "size": 0}

        # ----- 媒体缓存数据库 (.db) -----
        mc_db_dir = safe_path(PR_CACHE_PATHS["media_cache_db"][0])
        if mc_db_dir:
            db_files, db_size = scan_directory_recursive(mc_db_dir, {".db"})
            results["media_db"] = {"files": db_files, "size": db_size}
        else:
            results["media_db"] = {"files": [], "size": 0}

        # ----- 自动保存文件 (.prproj) -----
        auto_save_base = safe_path(PR_CACHE_PATHS["auto_save"][0])
        auto_save_files = []
        auto_save_size = 0
        if auto_save_base:
            for subdir in auto_save_base.iterdir():
                if subdir.is_dir():
                    as_dir = subdir / "Adobe Premiere Pro Auto-Save"
                    if as_dir.exists():
                        f_list, f_size = scan_directory_recursive(as_dir, {".prproj"})
                        auto_save_files.extend(f_list)
                        auto_save_size += f_size
        results["auto_save"] = {"files": auto_save_files, "size": auto_save_size}

        self.cache_scan_results = results

        # 回到主线程更新 UI
        self.root.after(0, self._update_cache_ui)

    def _update_cache_ui(self):
        """主线程：更新缓存扫描结果的 UI 显示"""
        results = self.cache_scan_results
        if not results:
            self.lbl_scan_status.configure(text="扫描完成（未找到缓存目录）", fg=COLOR_TEXT_SECONDARY)
            self.btn_scan_cache.configure(state="normal", text="扫描缓存")
            return

        for key, display_name, desc in self.cache_categories:
            data = results.get(key, {"files": [], "size": 0})
            files = data["files"]
            total_size = data["size"]
            count = len(files)
            size_str = format_size(total_size)
            color = get_size_color(total_size) if total_size > 0 else COLOR_TEXT_SECONDARY

            lbl = getattr(self, f"lbl_{key}", None)
            if lbl:
                lbl.configure(
                    text=f"{display_name}：{count} 个文件，共 {size_str}",
                    fg=color,
                )

        self.lbl_scan_status.configure(text="扫描完成", fg=COLOR_GREEN)
        self.btn_scan_cache.configure(state="normal", text="扫描缓存")

    # ==================================================================
    # 区域二：项目文件分析
    # ==================================================================

    def _build_project_analysis(self):
        """构建项目文件分析区域"""
        frame = ttk.LabelFrame(self.main_frame, text="区域二：项目文件分析", style="Card.TLabelframe")
        frame.pack(fill=tk.X, padx=8, pady=4)

        # 按钮行
        btn_frame = tk.Frame(frame, bg=COLOR_CARD)
        btn_frame.pack(fill=tk.X, padx=6, pady=(4, 4))

        ttk.Button(
            btn_frame,
            text="选择项目文件",
            command=self._on_select_project,
            style="Accent.TButton",
        ).pack(side=tk.LEFT, padx=(0, 8))

        self.btn_copy_links = ttk.Button(
            btn_frame,
            text="复制断链列表",
            command=self._on_copy_broken_links,
            style="Secondary.TButton",
            state="disabled",
        )
        self.btn_copy_links.pack(side=tk.LEFT)

        # 项目信息显示行
        self.lbl_project_info = tk.Label(
            frame,
            text="尚未选择项目文件",
            bg=COLOR_CARD,
            fg=COLOR_TEXT_SECONDARY,
            font=("微软雅黑", 8),
            anchor=tk.W,
        )
        self.lbl_project_info.pack(fill=tk.X, padx=6, pady=(0, 2))

        # 断链文件显示区域
        self.broken_text = tk.Text(
            frame,
            height=4,
            bg=COLOR_ENTRY_BG,
            fg=COLOR_RED,
            font=("Consolas", 8),
            wrap=tk.WORD,
            state="disabled",
            relief=tk.FLAT,
            padx=4,
            pady=4,
        )
        self.broken_text.pack(fill=tk.X, padx=6, pady=(0, 6))

    def _on_select_project(self):
        """选择 .prproj 文件并分析"""
        file_path = filedialog.askopenfilename(
            title="选择 Premiere Pro 项目文件",
            filetypes=[("PR 项目文件", "*.prproj"), ("所有文件", "*.*")],
        )
        if not file_path:
            return

        pr_path = Path(file_path)
        # 文件大小与复杂度
        try:
            size_bytes = pr_path.stat().st_size
        except OSError:
            messagebox.showerror("错误", "无法读取项目文件")
            return

        complexity = estimate_project_complexity(pr_path)
        self.lbl_project_info.configure(
            text=f"项目文件：{pr_path.name}  |  大小：{format_size(size_bytes)}  |  预估复杂度：{complexity}",
            fg=COLOR_TEXT,
        )

        # 提取媒体文件路径
        media_paths = extract_media_paths_from_prproj(pr_path)

        # 检测断链文件
        self.broken_links = [p for p in media_paths if not os.path.exists(p)]

        # 更新断链显示
        self.broken_text.configure(state="normal")
        self.broken_text.delete("1.0", tk.END)
        if self.broken_links:
            self.broken_text.insert("1.0", f"断链文件（{len(self.broken_links)} 个）：\n")
            for bl in self.broken_links:
                self.broken_text.insert(tk.END, f"  ✕ {bl}\n")
            self.btn_copy_links.configure(state="normal")
            self.lbl_project_info.configure(
                text=f"{self.lbl_project_info.cget('text')}  |  引用媒体：{len(media_paths)} 个  |  断链：{len(self.broken_links)} 个",
                fg=COLOR_ORANGE,
            )
        else:
            if media_paths:
                self.broken_text.insert(
                    "1.0",
                    f"未发现断链文件（共引用 {len(media_paths)} 个媒体文件，全部可访问）",
                )
            else:
                self.broken_text.insert("1.0", "未能从项目文件中提取到媒体路径信息。")
            self.btn_copy_links.configure(state="disabled")
        self.broken_text.configure(state="disabled")

    def _on_copy_broken_links(self):
        """一键复制断链文件列表到剪贴板"""
        if not self.broken_links:
            messagebox.showinfo("提示", "没有断链文件可复制")
            return
        text = "\n".join(self.broken_links)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        messagebox.showinfo("已复制", f"已将 {len(self.broken_links)} 个断链文件路径复制到剪贴板")

    # ==================================================================
    # 区域三：大文件扫描
    # ==================================================================

    def _build_area_large_files(self):
        """构建大文件扫描区域"""
        frame = ttk.LabelFrame(self.main_frame, text="区域三：大文件扫描（>100MB）", style="Card.TLabelframe")
        frame.pack(fill=tk.X, padx=8, pady=4)

        # 按钮行
        btn_frame = tk.Frame(frame, bg=COLOR_CARD)
        btn_frame.pack(fill=tk.X, padx=6, pady=(4, 2))

        ttk.Button(
            btn_frame,
            text="选择项目文件夹",
            command=self._on_select_project_folder,
            style="Secondary.TButton",
        ).pack(side=tk.LEFT, padx=(0, 6))

        self.btn_scan_large = ttk.Button(
            btn_frame,
            text="扫描大文件",
            command=self._on_scan_large_files,
            style="Accent.TButton",
            state="disabled",
        )
        self.btn_scan_large.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_move_archive = ttk.Button(
            btn_frame,
            text="移至存档",
            command=self._on_move_to_archive,
            style="Secondary.TButton",
            state="disabled",
        )
        self.btn_move_archive.pack(side=tk.LEFT)

        self.lbl_large_status = tk.Label(
            btn_frame,
            text="",
            bg=COLOR_CARD,
            fg=COLOR_TEXT_SECONDARY,
            font=("微软雅黑", 8),
        )
        self.lbl_large_status.pack(side=tk.LEFT, padx=(8, 0))

        # 大文件列表（支持多选）
        list_frame = tk.Frame(frame, bg=COLOR_CARD)
        list_frame.pack(fill=tk.X, padx=6, pady=(2, 6))

        scrollbar_y = tk.Scrollbar(list_frame, orient=tk.VERTICAL)
        scrollbar_x = tk.Scrollbar(list_frame, orient=tk.HORIZONTAL)

        self.listbox_large = tk.Listbox(
            list_frame,
            height=6,
            bg=COLOR_ENTRY_BG,
            fg=COLOR_TEXT,
            font=("Consolas", 8),
            selectmode=tk.MULTIPLE,
            selectbackground=COLOR_ACCENT,
            selectforeground="#ffffff",
            relief=tk.FLAT,
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set,
        )
        scrollbar_y.config(command=self.listbox_large.yview)
        scrollbar_x.config(command=self.listbox_large.xview)

        self.listbox_large.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        self.project_folder: Path | None = None

    def _on_select_project_folder(self):
        """选择项目文件夹"""
        folder = filedialog.askdirectory(title="选择 PR 项目所在的文件夹")
        if folder:
            self.project_folder = Path(folder)
            self.btn_scan_large.configure(state="normal")
            self.lbl_large_status.configure(
                text=f"已选：{self.project_folder.name}",
                fg=COLOR_TEXT_SECONDARY,
            )

    def _on_scan_large_files(self):
        """扫描项目文件夹中 >100MB 的视频文件"""
        if not self.project_folder:
            return

        self.btn_scan_large.configure(state="disabled", text="扫描中...")
        self.lbl_large_status.configure(text="正在扫描大文件...", fg=COLOR_ORANGE)
        self.listbox_large.delete(0, tk.END)

        thread = threading.Thread(target=self._scan_large_files_thread, daemon=True)
        thread.start()

    def _scan_large_files_thread(self):
        """后台线程：递归扫描 >100MB 的视频文件"""
        # 常见视频文件后缀
        video_exts = {
            ".mp4", ".mov", ".avi", ".mkv", ".wmv", ".flv", ".webm",
            ".m4v", ".mpg", ".mpeg", ".mts", ".m2ts", ".ts", ".r3d",
            ".braw", ".ari", ".arx", ".mxf",
        }
        threshold = 100 * 1024 * 1024  # 100MB

        results = []
        try:
            for entry in self.project_folder.rglob("*"):
                if entry.is_file() and entry.suffix.lower() in video_exts:
                    try:
                        stat = entry.stat()
                        if stat.st_size > threshold:
                            results.append({
                                "path": entry,
                                "size": stat.st_size,
                                "mtime": stat.st_mtime,
                            })
                    except OSError:
                        pass
        except PermissionError:
            pass

        # 按大小降序排列
        results.sort(key=lambda x: x["size"], reverse=True)
        self.large_files = results

        self.root.after(0, self._update_large_files_ui)

    def _update_large_files_ui(self):
        """主线程：更新大文件列表 UI"""
        self.listbox_large.delete(0, tk.END)

        if not self.large_files:
            self.listbox_large.insert(tk.END, "未发现 >100MB 的视频文件")
            self.lbl_large_status.configure(text="扫描完成，未发现大文件", fg=COLOR_GREEN)
        else:
            total_size = sum(f["size"] for f in self.large_files)
            for f in self.large_files:
                mtime_str = datetime.fromtimestamp(f["mtime"]).strftime("%Y-%m-%d %H:%M")
                # 截断长路径以便显示
                path_str = str(f["path"])
                if len(path_str) > 60:
                    path_str = "..." + path_str[-57:]
                display = f"{f['path'].name:30s} | {format_size(f['size']):>8s} | {mtime_str} | {path_str}"
                self.listbox_large.insert(tk.END, display)

            self.lbl_large_status.configure(
                text=f"共 {len(self.large_files)} 个大文件，合计 {format_size(total_size)}",
                fg=COLOR_ORANGE if total_size > 1024 * 1024 * 1024 else COLOR_GREEN,
            )

        self.btn_scan_large.configure(state="normal", text="扫描大文件")
        self.btn_move_archive.configure(state="normal" if self.large_files else "disabled")

    def _on_move_to_archive(self):
        """将列表框中多选的大文件移动到 _archive 子目录"""
        selected_indices = self.listbox_large.curselection()
        if not selected_indices:
            messagebox.showinfo("提示", "请先在列表中选择要移动的文件")
            return

        if not self.project_folder:
            return

        archive_dir = self.project_folder / "_archive"
        archive_dir.mkdir(exist_ok=True)

        moved_count = 0
        failed = []
        for idx in selected_indices:
            if idx >= len(self.large_files):
                continue
            file_info = self.large_files[idx]
            src = file_info["path"]
            dst = archive_dir / src.name

            # 处理重名
            if dst.exists():
                stem = dst.stem
                suffix = dst.suffix
                counter = 1
                while dst.exists():
                    dst = archive_dir / f"{stem}_{counter}{suffix}"
                    counter += 1

            try:
                shutil.move(str(src), str(dst))
                moved_count += 1
                # 更新记录
                file_info["path"] = dst
            except OSError as e:
                failed.append(f"{src.name}: {e}")

        # 刷新列表
        self._update_large_files_ui()
        msg = f"已移动 {moved_count} 个文件到 {archive_dir}"
        if failed:
            msg += f"\n失败 {len(failed)} 个：\n" + "\n".join(failed)
        messagebox.showinfo("移动完成", msg)

    # ==================================================================
    # 区域四：一键清理
    # ==================================================================

    def _build_area_cleanup(self):
        """构建一键清理区域"""
        frame = ttk.LabelFrame(self.main_frame, text="区域四：一键优化", style="Card.TLabelframe")
        frame.pack(fill=tk.X, padx=8, pady=(4, 8))

        # 大按钮
        self.btn_cleanup = ttk.Button(
            frame,
            text="一键优化",
            command=self._on_cleanup,
            style="BigAccent.TButton",
        )
        self.btn_cleanup.pack(pady=(12, 4))

        # v2.0：dry-run 与报告导出按钮
        btn_row = tk.Frame(frame, bg=COLOR_CARD)
        btn_row.pack(pady=(0, 4))
        ttk.Button(btn_row, text="预览清理（dry-run）",
                   command=self._on_dry_run).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="导出上次报告",
                   command=self._on_export_report).pack(side=tk.LEFT)
        self.last_report = None

        # 说明文字
        tk.Label(
            frame,
            text="清理所有勾选的缓存文件，自动保存保留最近 5 个；支持跨平台回收站",
            bg=COLOR_CARD,
            fg=COLOR_TEXT_SECONDARY,
            font=("微软雅黑", 8),
        ).pack(pady=(0, 8))

    def _on_dry_run(self):
        """v2.0：安全的 dry-run —— 生成详细清理计划，绝不删除文件"""
        if not self.cache_scan_results:
            messagebox.showinfo("提示", "请先执行「扫描缓存」再进行预览")
            return
        selected = {key for key, _, _ in self.cache_categories
                    if self.cache_checkbox_vars.get(key) and
                    self.cache_checkbox_vars[key].get()}
        plan = build_cleanup_plan(self.cache_scan_results, self.cache_categories,
                                  selected_keys=selected)
        text = format_cleanup_plan(plan, dry_run=True)
        self.last_report = execute_cleanup_plan(plan, dry_run=True)
        messagebox.showinfo("干跑预览（未删除任何文件）", text)

    def _on_export_report(self):
        """v2.0：导出最近一次报告为 JSON"""
        if not self.last_report:
            messagebox.showinfo("提示", "请先执行干跑预览或清理操作")
            return
        path = filedialog.asksaveasfilename(
            title="导出详细报告", defaultextension=".json",
            filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")])
        if path:
            out = export_cleanup_report(self.last_report, path)
            messagebox.showinfo("导出成功", f"报告已保存: {out}")

    def _on_cleanup(self):
        """一键清理：确认后执行清理并输出报告"""
        if not self.cache_scan_results:
            messagebox.showinfo("提示", "请先执行「扫描缓存」再进行清理")
            return

        # ----- 构建清理预览 -----
        preview_lines = ["即将执行以下清理操作：\n"]

        total_files = 0
        total_size = 0
        cleanup_plan = {}  # key -> [file_paths]

        for key, display_name, _ in self.cache_categories:
            var = self.cache_checkbox_vars.get(key)
            if var and var.get():
                data = self.cache_scan_results.get(key, {"files": [], "size": 0})
                files = data["files"]
                size = data["size"]

                if key == "auto_save":
                    # 自动保存：保留最近 5 个
                    if len(files) > 5:
                        # 按修改时间排序，保留最新的 5 个
                        files_sorted = sorted(files, key=lambda f: f.stat().st_mtime, reverse=True)
                        to_delete = files_sorted[5:]
                        kept = files_sorted[:5]
                        del_size = sum(f.stat().st_size for f in to_delete)
                        preview_lines.append(
                            f"  [{display_name}] 保留 {len(kept)} 个最新，删除 {len(to_delete)} 个（{format_size(del_size)}）"
                        )
                        cleanup_plan[key] = to_delete
                        total_files += len(to_delete)
                        total_size += del_size
                    else:
                        preview_lines.append(
                            f"  [{display_name}] 仅 {len(files)} 个自动保存，无需清理"
                        )
                        cleanup_plan[key] = []
                else:
                    if files:
                        preview_lines.append(
                            f"  [{display_name}] 删除 {len(files)} 个文件（{format_size(size)}）"
                        )
                        cleanup_plan[key] = files
                        total_files += len(files)
                        total_size += size
                    else:
                        preview_lines.append(f"  [{display_name}] 无缓存文件")
                        cleanup_plan[key] = []

        if total_files == 0:
            messagebox.showinfo("提示", "没有需要清理的文件")
            return

        preview_lines.append(f"\n总计将删除 {total_files} 个文件，释放约 {format_size(total_size)} 空间")
        preview_lines.append("\n所有删除操作将使用回收站方式执行，可随时恢复。")

        # 弹出确认对话框
        confirm = messagebox.askyesno(
            "确认清理",
            "\n".join(preview_lines),
        )
        if not confirm:
            return

        # ----- 执行清理（后台线程）-----
        self.btn_cleanup.configure(state="disabled", text="清理中...")
        thread = threading.Thread(
            target=self._execute_cleanup_thread,
            args=(cleanup_plan, total_files, total_size),
            daemon=True,
        )
        thread.start()

    def _execute_cleanup_thread(self, cleanup_plan: dict, total_files: int, total_size: int):
        """后台线程：逐项执行清理"""
        deleted_files = 0
        deleted_size = 0
        failed_files = []

        for key, files in cleanup_plan.items():
            for f in files:
                try:
                    if send_to_recycle_bin(str(f)):
                        deleted_files += 1
                        try:
                            deleted_size += f.stat().st_size
                        except OSError:
                            pass
                    else:
                        failed_files.append(str(f))
                except Exception as e:
                    failed_files.append(f"{f}: {e}")

        # 回到主线程显示报告
        self.root.after(
            0,
            self._show_cleanup_report,
            deleted_files,
            deleted_size,
            failed_files,
            total_files,
            total_size,
        )

    def _show_cleanup_report(
        self,
        deleted: int,
        deleted_size: int,
        failed: list,
        planned_files: int,
        planned_size: int,
    ):
        """主线程：显示清理完成报告（v2.0：详细报告，可导出）"""
        self.btn_cleanup.configure(state="normal", text="一键优化")
        report = {
            'tool': 'PRTrim', 'version': '2.0',
            'generated_at': datetime.now().isoformat(timespec='seconds'),
            'platform': detect_platform(), 'dry_run': False,
            'categories': [], 'planned_files': planned_files, 'planned_size': planned_size,
            'deleted_files': deleted, 'deleted_size': deleted_size, 'failed': failed, 'skipped': [],
        }
        self.last_report = report
        text = format_cleanup_report(report)
        lines = [text]
        if detect_platform() == 'windows':
            lines.append("\n提示：如需恢复，请打开 Windows 回收站。")
        elif detect_platform() == 'macos':
            lines.append("\n提示：如需恢复，请打开 macOS 废纸篓。")
        else:
            lines.append("\n提示：如需恢复，请查看回收站目录 ~/.local/share/Trash。")
        messagebox.showinfo("清理报告", '\n'.join(lines))


# ============================================================================
# 程序入口
# ============================================================================

def build_parser():
    """v2.0：构建命令行参数解析器"""
    parser = argparse.ArgumentParser(
        prog='pr_trim.py',
        description='PR瘦身 v2.0 — Premiere Pro 缓存清理与项目优化（安全的 dry-run + 详细报告）',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='示例:\n'
               '  pr_trim.py --dry-run                # 只预览清理计划，不删除\n'
               '  pr_trim.py --dry-run --report plan.json\n'
               '  pr_trim.py --keep-recent 10 --report result.json\n'
               '  pr_trim.py --easter\n'
    )
    parser.add_argument('--dry-run', action='store_true',
                        help='安全预览：只生成清理计划与报告，不删除任何文件')
    parser.add_argument('--keep-recent', type=int, default=5,
                        help='自动保存保留最近 N 个（默认 5，0 表示全部清理）')
    parser.add_argument('--report', dest='report_path', help='导出详细报告为 JSON')
    parser.add_argument('--cache-root', action='append', default=[],
                        help='自定义缓存根目录（可多次指定）')
    parser.add_argument('--gui', action='store_true', help='启动图形界面')
    parser.add_argument('--easter', action='store_true', help='隐藏彩蛋（嘘……）')
    return parser


def run_cli(argv=None):
    """v2.0：命令行入口，返回退出码"""
    args = build_parser().parse_args(argv)

    if args.easter:
        print_easter_egg()
        return 0

    if args.gui or (not any([args.dry_run, args.report_path, args.cache_root])):
        if tk is None:
            print("错误：当前环境无 tkinter，请使用 --dry-run 等命令行参数", file=sys.stderr)
            return 2
        main()
        return 0

    # 收集缓存目录：默认路径 + 自定义路径
    cache_results = {}
    categories = []
    allowed_roots = [Path(p) for p in args.cache_root if p]
    for name, resolvers in PR_CACHE_PATHS.items():
        roots = []
        for resolver in resolvers:
            p = safe_path(resolver)
            if p:
                roots.append(p)
        if not roots:
            continue
        all_files, total = [], 0
        for root in roots:
            files, size = scan_directory_recursive(root)
            all_files.extend(files)
            total += size
        display = {'media_cache_db': '媒体缓存数据库', 'media_cache_files': '媒体缓存文件',
                   'auto_save': '自动保存'}.get(name, name)
        cache_results[name] = {'files': all_files, 'size': total}
        categories.append((name, display, None))

    # 自定义缓存根目录作为独立分类扫描，避免与默认目录重复计数
    if allowed_roots:
        all_files, total = [], 0
        for root in allowed_roots:
            files, size = scan_directory_recursive(root)
            all_files.extend(files)
            total += size
        if all_files:
            cache_results['custom'] = {'files': all_files, 'size': total}
            categories.append(('custom', '自定义缓存目录', None))

    if not cache_results:
        print("未发现可清理的 PR 缓存目录（Windows 默认路径不存在）。"
              "可用 --cache-root 指定目录。", file=sys.stderr)
        return 1

    plan = build_cleanup_plan(cache_results, categories,
                              keep_recent=max(0, args.keep_recent))
    print(format_cleanup_plan(plan, dry_run=True))
    report = execute_cleanup_plan(plan, dry_run=True)

    if args.report_path:
        out = export_cleanup_report(report, args.report_path)
        print(f"\n详细报告已导出: {out}")
    return 0


def main():
    """启动 PR瘦身 v2.0 应用程序"""
    root = tk.Tk()
    # 设置窗口图标（如果有的话）
    try:
        root.iconbitmap(default="")
    except Exception:
        pass
    app = PRTrimApp(root)
    root.mainloop()


if __name__ == "__main__":
    sys.exit(run_cli())
