# PRTrim

Premiere Pro 缓存清理与项目优化工具（Windows / Python + tkinter）。

## 功能

- **缓存扫描**：自动检测 PR 媒体缓存目录，按类型统计文件数与体积
- **项目文件分析**：解析 `.prproj`，检测断链（丢失）媒体文件
- **大文件扫描**：递归扫描项目文件夹，找出 >100MB 的视频文件并可归档
- **一键优化**：批量清理选中缓存，自动保留最近 5 个
- 所有删除走**回收站**，可随时恢复

## 安装与运行

```bash
install.bat                  # 或 pip install -r requirements.txt
python pr_trim.py
```

## 说明

`.prproj` 实际是 gzip 压缩的 XML，工具会直接解压解析并提取媒体路径，
不需要启动 Premiere Pro。仅读取，不修改工程文件。
