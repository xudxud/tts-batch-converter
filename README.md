# TTS批量转换工具

一个使用 Tauri 2、React 和 Rust 开发的 Windows Excel 批量文字转语音工具。读取 Excel 中的文件名和文字内容，逐行生成独立 MP3 文件。

## 功能

- 读取 `.xlsx` 第一个工作表，A 列作为文件名，B 列作为朗读内容。
- 支持普通话、台湾国语、儿童感、美式英语和英式英语音色。
- 支持中英文混合朗读、在线试听和语速调整。
- 支持逐行勾选及 Excel 行号范围，例如 `2-10,15,20-25`。
- 可只重新生成修改过的行。
- 覆盖同名文件前确认，生成失败时保留原 MP3。
- 网络或服务失败可单条重试，也可在全部完成后批量重试。
- 任务状态持久保存，重新打开应用后可继续处理失败项。
- 支持 Excel `000`、`0000` 等数字格式的前导零文件名。

## 使用

1. 下载 Release 中的 Windows 安装包或便携版。
2. 选择 Excel 文件和 MP3 输出文件夹。
3. 如果 Excel 第一行是表头，勾选“第一行是表头”。
4. 选择音色并试听。
5. 勾选需要生成的行，点击“开始生成”。

详细说明见 [`使用说明.md`](使用说明.md)。

## Excel 格式

| A 列：MP3 文件名 | B 列：朗读内容 |
| --- | --- |
| 001 | 欢迎使用批量文字转语音工具。 |
| product-A | This is product A. |

程序将生成 `001.mp3` 和 `product-A.mp3`。

## 开发

### 环境

- Node.js 20+
- Rust stable MSVC
- Microsoft C++ Build Tools
- Microsoft Edge WebView2

### 命令

```powershell
npm install
npm run tauri dev
```

检查和构建：

```powershell
npm run build
cargo test --manifest-path src-tauri/Cargo.toml
cargo clippy --manifest-path src-tauri/Cargo.toml -- -D warnings
npm run tauri build
```

## 技术栈

- Tauri 2
- React 19 + TypeScript
- Rust
- `calamine` 读取 Excel
- `msedge-tts` 调用 Microsoft Edge 在线朗读服务

## 注意

- 语音生成需要联网。
- Edge 在线朗读服务没有公开固定额度或可用性承诺。
- 当前 Windows 发布文件未进行代码签名，首次运行时可能出现 Microsoft Defender SmartScreen 提示。
- 仓库当前未附带开源许可证。
