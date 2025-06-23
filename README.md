# PPT浏览器 VS Code 插件

本插件允许你在 VS Code 内部直接浏览 PPT（.pptx）文件，无需离开编辑器。

## 功能特性
- 通过命令面板一键选择并预览本地 PPTX 文件
- 支持 PPTX 幻灯片内容的解析与渲染（基于 jszip + pptxjs）
- 支持幻灯片切换（pptxjs 内置）

## 安装方法
1. 克隆本仓库：
   ```sh
   git clone https://github.com/Niobium-41-nb/ppt-viewer-vscode.git
   ```
2. 在 VS Code 中打开本项目文件夹。
3. 运行 `npm install` 安装依赖。
4. 按 F5 启动扩展开发主机进行调试。

## 使用方法
1. 在 VS Code 命令面板（Ctrl+Shift+P）输入并选择“浏览PPT文件”。
2. 选择本地的 .pptx 文件。
3. 即可在 Webview 面板中浏览 PPT 内容。

## 效果预览
（可自行添加截图）

## 技术实现
- Webview 内集成 [jszip](https://stuk.github.io/jszip/) 和 [pptxjs](https://gitbrent.github.io/PptxGenJS/) 实现 PPTX 文件解析与渲染。
- 纯本地预览，无需上传文件。

## 贡献
欢迎 issue、PR 和建议！

## 许可协议
MIT License
