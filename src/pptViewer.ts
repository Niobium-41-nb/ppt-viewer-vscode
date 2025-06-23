import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';

export function activate(context: vscode.ExtensionContext) {
	console.log('PPT浏览器扩展已激活');

	const disposable = vscode.commands.registerCommand('ppt---.browsePPT', async () => {
		const fileUri = await vscode.window.showOpenDialog({
			canSelectMany: false,
			filters: { 'PowerPoint': ['pptx'] },
			openLabel: '选择PPT文件'
		});
		if (!fileUri || fileUri.length === 0) {
			return;
		}

		const pptxPath = fileUri[0].fsPath;
		const panel = vscode.window.createWebviewPanel(
			'pptViewer',
			'PPT浏览器',
			vscode.ViewColumn.One,
			{ enableScripts: true, localResourceRoots: [vscode.Uri.file(path.dirname(pptxPath))] }
		);

		const pptxData = fs.readFileSync(pptxPath).toString('base64');
		panel.webview.html = getWebviewContent(pptxData);
	});

	context.subscriptions.push(disposable);
}

function getWebviewContent(pptxBase64: string): string {
	return `<!DOCTYPE html>
<html lang="zh-cn">
<head>
	<meta charset="UTF-8">
	<title>PPT浏览器</title>
	<script src="https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js"></script>
	<script src="https://cdn.jsdelivr.net/npm/pptxjs@3.1.5/dist/pptxjs.min.js"></script>
	<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/pptxjs@3.1.5/dist/pptxjs.min.css" />
	<style>body{margin:0;padding:0;}#ppt-container{width:100vw;height:90vh;overflow:auto;}</style>
</head>
<body>
	<div id="ppt-container">正在加载PPT...</div>
	<script>
	(function(){
		const pptxBase64 = "${pptxBase64}";
		const pptArrayBuffer = Uint8Array.from(atob(pptxBase64), c => c.charCodeAt(0)).buffer;
		const pptContainer = document.getElementById('ppt-container');
		window.PPTX = window.PPTX || {};
		PPTX.render = function(){
			window.PPTXJS.render({
				fileContent: pptArrayBuffer,
				container: pptContainer,
				slideMode: true
			});
		};
		if(window.PPTXJS){
			PPTX.render();
		}else{
			window.addEventListener('PPTXJSReady', PPTX.render);
		}
	})();
	</script>
</body>
</html>`;
}

export function deactivate() {}
