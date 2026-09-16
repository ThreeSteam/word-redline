# Word Redline 2.0 for macOS

## 下载与启动

在仓库 **Releases** 下载 `Word-Redline-2.0-macOS-Universal.zip`，完整解压后将 `Word Redline.app` 放到“应用程序”或其他位置运行。

- 需要 macOS 13+ 和可正常使用的 Microsoft Word for Mac 桌面版。
- App 同时包含 Apple Silicon（arm64）和 Intel（x86_64）代码，运行已构建版本不需要 Python 或 Xcode。
- 首次比较时可能需要允许 Word Redline 控制 Microsoft Word，以及允许 Word 访问文件。拒绝后可在“系统设置 → 隐私与安全性 → 自动化”检查权限。
- App 使用临时签名，未进行 Apple Developer ID 签名或公证。macOS 可能阻止直接打开；请在核对来源后按系统安全提示处理，或从源码构建。

## 操作与保存

1. 左侧拖入旧版，右侧拖入新版；两侧可放多份文件，也可点击“添加文件”。
2. 根据文件名生成配对建议，在下拉菜单调整对应旧版，取消勾选不需要的行。
3. 确认后点击“开始比对”，完成后点击该行“打开结果”。

仅拖入文件不会启动 Word。配对识别投资协议 / 增资协议、股东协议、章程、股权转让等关键词，并参考文件名相似度；候选不明确需手动选择。不能重复使用同一旧版，也不能让文件与自己比较。

每组结果在对应新版目录，命名为 `新版文件名 - redline.docx`；重名自动递增为 `新版文件名 - redline (2).docx` 等，不覆盖已有文件。

源文件会复制到本机临时目录，Word 仅打开副本执行比较。请先保存最新修改；比较期间请勿切换或编辑 Word 文档，以免干扰结果识别。

成功后清理该组临时副本；失败或超时保留临时目录，并在错误中显示位置，方便检查。临时副本也包含文档内容。

## 平台差异

- Mac 通过 AppleScript 控制 Word；Windows 使用独立 Word 实例和 COM。
- “停止后续比对”等待当前组结束后停止，不强行退出用户 Word。
- Word 比对错误或超时会停止后续组，避免隐藏授权提示或未结束的命令影响下一组。处理后只勾选需要重试的组即可。
- 系统及 Word 权限提示由用户处理，不能静默绕过。

## 验证状态

Mac 版为预览版。GitHub Actions 执行实际 macOS 编译、Universal 架构检查与文件名配对测试。构建机没有已授权的 Word for Mac，**实际 Word 比较与保存尚待实机验证**。

正式用于工作前，请先用两份测试文件验证，检查修订结果。若 Word 不支持脚本中的命令，界面会显示错误；请提供错误文字及 Word 版本，以便定位。

## 从源码构建

在装有 Xcode Command Line Tools 的 Mac 上，从仓库根目录运行：

```bash
bash macos/build.sh
```

App 与 Universal ZIP 输出到 `macos/build/`，无第三方运行库依赖。

## 接口参考

- [Microsoft Word 比较接口](https://learn.microsoft.com/en-us/office/vba/api/word.application.comparedocuments)
- [Microsoft 社区：Mac 比较接口历史兼容性说明](https://learn.microsoft.com/en-us/answers/questions/4906054/application-comparedocuments-method-in-office-mac)
- [Apple Mac Automation Scripting Guide](https://developer.apple.com/library/archive/documentation/LanguagesUtilities/Conceptual/MacAutomationScriptingGuide/)

Mac 运行时通过已安装 Word 的脚本字典解析命令，与 Windows 的接口不同。
