# PRTrim

Premiere Pro 缓存清理与项目优化工具（v2.0，Windows / macOS / Linux，Python + tkinter）。

## 功能

- **缓存扫描**：自动检测 PR 媒体缓存目录，按类型统计文件数与体积
- **项目文件分析**：解析 `.prproj`，检测断链（丢失）媒体文件
- **大文件扫描**：递归扫描项目文件夹，找出 >100MB 的视频文件并可归档
- **一键优化**：批量清理选中缓存，自动保留最近 5 个自动保存
- **v2.0 新增**：安全的 dry-run、跨平台回收站、详细报告（JSON/文本）、CLI 模式

## 安装与运行

```bash
install.bat                  # 或 pip install -r requirements.txt
python pr_trim.py            # 图形界面
```

## v2.0 更新说明

### 1. 更安全的 dry-run
操作前先看计划，磁盘零改动：

```bash
python pr_trim.py --dry-run
python pr_trim.py --dry-run --cache-root /path/to/cache --keep-recent 10 --report plan.json
```

GUI 中新增「预览清理（dry-run）」按钮。`build_cleanup_plan()` 是纯函数，
只生成计划不删除任何文件；`execute_cleanup_plan(dry_run=True)` 仅产出报告。

### 2. 删除前安全校验
`is_safe_to_delete()` 拒绝：系统关键目录、非普通文件、符号链接、
以及不在允许缓存根目录下的路径。被拒绝的文件会列入报告的「安全跳过」清单。

### 3. 多平台回收站
- Windows：PowerShell + `Microsoft.VisualBasic.FileIO`（保留 v1 行为）
- macOS：`osascript` 交给 Finder 废纸篓
- Linux：优先 `gio trash` / `trash-put`，否则移动到 `~/.local/share/Trash/files`

### 4. 详细报告
`--report result.json` 导出含时间、平台、dry-run 标志、各分类计划/实际删除数、
失败与安全跳过明细的 JSON。GUI 新增「导出上次报告」按钮。

### 5. 命令行模式
无需 GUI 即可扫描自定义缓存目录（`--cache-root` 可多次指定）并输出清理计划。

## 彩蛋秘籍

| 入口 | 触发方式 | 效果 |
| --- | --- | --- |
| `--easter` | `python pr_trim.py --easter` | 输出 SOLCAT ASCII 署名与隐藏提示 |
| 清理彩蛋 | 保留最近 5 个自动保存 | 第 42 个才是真正的秘密 ;) |

通用秘籍（GUI/其他 SOLCAT 项目）：↑↑↓↓←→←→BA、连点标题 5 次、输入 `solcat`、Ctrl+Shift+T。

## 说明

`.prproj` 实际是 gzip 压缩的 XML，工具会直接解压解析并提取媒体路径，
不需要启动 Premiere Pro。仅读取，不修改工程文件。
所有实际删除均走回收站，可随时恢复；dry-run 模式下不会删除任何文件。
