# -*- coding: utf-8 -*-
"""
PR瘦身 v1.0 — Premiere Pro 缓存清理与项目优化工具
=====================================================
解决 Premiere Pro 在高端硬件（RTX 4070 + 32GB RAM）上仍然卡顿的痛点。
根本原因：媒体缓存堆积、预览文件膨胀、项目碎片化。

四大功能模块：
  1. 缓存扫描 — 自动检测 PR 媒体缓存目录，按类型统计文件数与总体积
  2. 项目文件分析 — 解析 .prproj 项目文件，检测断链（丢失）媒体文件
  3. 大文件扫描 — 递归扫描项目文件夹，找出超大视频文件并可归档
  4. 一键优化 — 批量清理选中的缓存，保留最近 5 个自动保存

依赖：仅使用 Python 标准库（os / shutil / xml / tkinter / datetime / pathlib / subprocess / gzip）。
未使用任何第三方包，无需 pip install。
"""

import os
import re
import shutil
import gzip
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from datetime import datetime
from collections import defaultdict
import subprocess
import threading


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


def send_to_recycle_bin(file_path: str) -> bool:
    """
    通过 PowerShell 调用 .NET Framework 方法将文件/目录移入 Windows 回收站。
    使用 Microsoft.VisualBasic.FileIO.FileSystem 提供可靠的回收站操作。
    返回 True 表示成功，False 表示失败。
    """
    try:
        # 判断路径类型，选择对应方法
        if os.path.isfile(file_path):
            method = "DeleteFile"
        elif os.path.isdir(file_path):
            method = "DeleteDirectory"
        else:
            return False

        # 构建 PowerShell 命令：加载 VB 程序集后调用删除方法
        ps_script = (
            f"Add-Type -AssemblyName Microsoft.VisualBasic; "
            f"[Microsoft.VisualBasic.FileIO.FileSystem]::{method}("
            f"'{file_path}', 'OnlyErrorDialogue', 'SendToRecycleBin')"
        )
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_script],
            capture_output=True,
            text=True,
            timeout=30,
        )
        return result.returncode == 0
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
# 主应用程序类
# ============================================================================

class PRTrimApp:
    """PR瘦身 v1.0 主窗口应用程序"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PR瘦身 v1.0")
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

        # 说明文字
        tk.Label(
            frame,
            text="清理所有勾选的缓存文件，自动保存保留最近 5 个",
            bg=COLOR_CARD,
            fg=COLOR_TEXT_SECONDARY,
            font=("微软雅黑", 8),
        ).pack(pady=(0, 8))

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
        """主线程：显示清理完成报告"""
        self.btn_cleanup.configure(state="normal", text="一键优化")

        lines = ["===== PR瘦身 v1.0 清理报告 =====", ""]
        lines.append(f"已清理文件数：{deleted} 个")
        lines.append(f"释放空间：{format_size(deleted_size)}")
        lines.append(f"计划清理：{planned_files} 个文件（{format_size(planned_size)}）")
        if failed:
            lines.append(f"\n失败 {len(failed)} 个：")
            for f_path in failed[:20]:  # 最多显示 20 个失败项
                lines.append(f"  - {f_path}")
            if len(failed) > 20:
                lines.append(f"  ... 还有 {len(failed) - 20} 个")
        else:
            lines.append("\n全部清理成功！所有文件已移入回收站。")
        lines.append("\n提示：如需恢复，请打开 Windows 回收站。")

        messagebox.showinfo("清理报告", "\n".join(lines))


# ============================================================================
# 程序入口
# ============================================================================

def main():
    """启动 PR瘦身 v1.0 应用程序"""
    root = tk.Tk()
    # 设置窗口图标（如果有的话）
    try:
        root.iconbitmap(default="")
    except Exception:
        pass
    app = PRTrimApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
