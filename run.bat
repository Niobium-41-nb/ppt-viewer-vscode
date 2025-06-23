@echo off
REM 自动卸载旧插件（如未发布可跳过此步）
REM code --uninstall-extension publisher.ppt-viewer-vscode

REM 检查 .vsix 文件是否存在
if not exist "ppt---0.0.1.vsix" (
    echo 未找到 ppt---0.0.1.vsix，请先用 vsce package 生成！
    pause
    exit /b
)

REM 自动安装最新插件
code --install-extension ppt---0.0.1.vsix

REM 启动 VS Code
code .
echo 插件已自动安装完成！
pause