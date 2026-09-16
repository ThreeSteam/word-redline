# Word Redline 2.0 — Windows & macOS

拖入新旧文件，确认配对后批量执行 Word 原生比较。结果保存在各新版目录，使用 ` - redline` 后缀，同名自动编号。

## 下载

- **Windows**：`Word-Redline-2.0-Windows-Portable.zip`，完整解压后运行 `Word Redline.exe`，保留同目录 PS1 文件。
- **macOS**：`Word-Redline-2.0-macOS-Universal.zip`，解压后运行 `Word Redline.app`；支持 macOS 13+、Apple Silicon 和 Intel。

两个平台都需要可正常使用的 Microsoft Word 桌面版。

## 验证状态

- Windows：已验证三组实际 Word 比较、增删修订、配对拦截、迁移路径及源文件不变。
- macOS：已由 macOS CI 编译并执行配对测试；**Word for Mac 的实际比较尚未实机验证**，作为预览版提供。
- Mac App 未进行 Developer ID 签名或 Apple 公证，可能显示系统安全提示。

本次双平台发布标记为预览发布，以反映 Mac 版的验证状态。详细说明见仓库 `docs/`。
