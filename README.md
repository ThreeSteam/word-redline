# Word Redline 2.0

**Windows / macOS 双平台 Word 文档比对工具。** 拖入新旧文件，核对配对后点击开始，调用 Microsoft Word 原生比较功能，自动保存带修订标记的 Redline 文档。

## 下载

**[前往 Releases 下载两个平台版本](https://github.com/ThreeSteam/word-redline-windows/releases)**

| 平台 | 下载文件 | 运行要求 | 验证状态 |
| --- | --- | --- | --- |
| Windows | `Word-Redline-2.0-Windows-Portable.zip` | Windows 10/11 + Word 桌面版 | 已验证实际 Word 批量比对 |
| macOS | `Word-Redline-2.0-macOS-Universal.zip` | macOS 13+ + Word for Mac | 预览版；实际 Word 比对待实机验证 |

绿色版无需安装本工具；完整解压后运行 EXE 或 App。Mac 包支持 Apple Silicon 和 Intel。**两版都不包含 Office，必须安装 Microsoft Word。**

## 使用方法

1. 左侧拖入一组旧版，右侧拖入一组新版。
2. 检查配对预览，可手动选择对应旧版，不需要的行取消勾选。
3. 点击 **开始比对**。仅拖入文件不会自动执行。
4. 结果在每份新版目录，命名为 `新版文件名 - redline.docx`；重名自动编号。

配对根据投资协议 / 增资协议、股东协议、章程等关键词及文件名相似度生成建议，**不是正文语义识别**。候选不明确时需手动选择。输入支持 `.docx`、`.doc`、`.docm`，输出统一为 `.docx`。

## 说明文档

- [Windows 使用说明](docs/windows.md)
- [macOS 使用说明、授权与验证范围](docs/macos.md)
- [2.0 双平台发布说明](docs/release-2.0.md)

## 仓库结构

```text
windows/
  src/Redline.cs       Windows Forms 界面、配对与批量调度
  portable/           可直接使用的 Windows 绿色版
  build.ps1           Windows 编译脚本
macos/
  Sources/            SwiftUI 界面、配对与任务执行
  Resources/          Word AppleScript 与 App 配置
  build.sh            Universal App 编译和打包
docs/                 分平台说明与发布说明
.github/workflows/    双平台构建、测试和发布
```

仓库沿用最初确认的 `word-redline-windows` 名称，内部已按双平台组织。平台特有的 Word 连接与界面实现分别维护，用户流程、配对规则与命名保持一致。

## 从源码构建

Windows PowerShell：

```powershell
./windows/build.ps1
```

Mac（需要 Xcode Command Line Tools）：

```bash
bash macos/build.sh
```

推送 `main` 运行双平台构建及配对测试；推送 `v*` 标签在构建通过后发布两个 ZIP。

## 数据处理与限制

- 本工具在本地处理，不上传文档到比对服务。
- 源文件请先保存最新修改，新版目录需有写权限。
- Word 授权、密码或文件访问提示由用户处理。
- Mac 当前组完成后才停止后续任务，不强行退出用户的 Word。
- Mac 实际 Word 比较尚待实机验证，且未进行 Apple Developer ID 签名或公证。构建及配对测试通过不等于 Word 全流程已验证。

上传内容只包含程序、源码与说明，不包含实际工作文档、测试协议或本地任务日志。
