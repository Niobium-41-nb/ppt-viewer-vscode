# PPT浏览器 VS Code 插件

[![Visual Studio Marketplace Version](https://img.shields.io/visual-studio-marketplace/v/Niobium-41-nb.ppt-viewer-vscode?label=VS%20Code%20Marketplace)](https://marketplace.visualstudio.com/items?itemName=Niobium-41-nb.ppt-viewer-vscode)
[![GitHub stars](https://img.shields.io/github/stars/Niobium-41-nb/ppt-viewer-vscode?style=social)](https://github.com/Niobium-41-nb/ppt-viewer-vscode)
[![MIT License](https://img.shields.io/github/license/Niobium-41-nb/ppt-viewer-vscode)](LICENSE)

## 简介

PPT浏览器是一款 VS Code 扩展插件，允许你在 VS Code 内部直接浏览 PPT（.pptx）文件，无需离开编辑器。

---

## 功能特性

- 通过命令面板一键选择并预览本地 PPTX 文件
- 支持 PPTX 幻灯片内容的解析与渲染（基于 jszip + pptxjs）
- 支持幻灯片切换（pptxjs 内置）
- 纯本地预览，无需上传文件，安全高效

---

## 安装方法

### 方式一：市场安装

1. 在 [VS Code 扩展市场](https://marketplace.visualstudio.com/items?itemName=Niobium-41-nb.ppt-viewer-vscode) 搜索并安装“PPT浏览器”插件。

### 方式二：本地安装

1. 克隆本仓库：
   ```sh
   git clone https://github.com/Niobium-41-nb/ppt-viewer-vscode.git
   cd ppt-viewer-vscode
   npm install
   npm run compile
   ```
2. 打包扩展：
   ```sh
   vsce package
   ```
3. 在 VS Code 终端执行：
   ```sh
   code --install-extension ppt----0.0.1.vsix
   ```

---

## 使用方法

1. 打开 VS Code。
2. 按下 `Ctrl+Shift+P`，打开命令面板。
3. 输入并选择“浏览PPT文件”命令。
4. 在弹出的文件选择框中，选择你要浏览的 `.pptx` 文件。
5. 选中后，会自动弹出一个 Webview 面板，直接在 VS Code 内浏览 PPT 内容，并支持翻页。

---

## 效果预览

> ![插件效果预览](images/demo.gif)

---

## 技术实现

- Webview 内集成 [jszip](https://stuk.github.io/jszip/) 和 [pptxjs](https://gitbrent.github.io/PptxGenJS/) 实现 PPTX 文件解析与渲染。
- 纯本地预览，无需上传文件。

---

## 常见问题

- **Q: 支持哪些 PPT 格式？**
  - 目前仅支持 .pptx 格式。
- **Q: PPT 内容是否会上传？**
  - 不会，所有解析和渲染均在本地完成。
- **Q: 如何卸载/升级插件？**
  - 可在 VS Code 扩展管理器中直接操作。

---

## 贡献

欢迎 issue、PR 和建议！

---

## 许可证

本项目采用 [MIT License](LICENSE) 开源协议，欢迎自由使用和二次开发。
