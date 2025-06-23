import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';

export function activate(context: vscode.ExtensionContext) {
	console.log('PPT浏览器扩展已激活');

	// 注册 .pptx 文件的自定义编辑器
	context.subscriptions.push(
		vscode.window.registerCustomEditorProvider(
			'ppt-viewer-vscode.pptxViewer',
			new PPTXCustomEditorProvider(),
			{ supportsMultipleEditorsPerDocument: false }
		)
	);

	// 保留命令方式，兼容老用户
	const disposable = vscode.commands.registerCommand('ppt---.browsePPT', async () => {
		const fileUri = await vscode.window.showOpenDialog({
			canSelectMany: false,
			filters: { 'PowerPoint': ['pptx'] },
			openLabel: '选择PPT文件'
		});
		if (!fileUri || fileUri.length === 0) {
			return;
		}
		vscode.commands.executeCommand('vscode.openWith', fileUri[0], 'ppt-viewer-vscode.pptxViewer');
	});
	context.subscriptions.push(disposable);
}

class PPTXCustomEditorProvider implements vscode.CustomReadonlyEditorProvider {
	async openCustomDocument(uri: vscode.Uri): Promise<vscode.CustomDocument> {
		return { uri, dispose: () => {} };
	}

	async resolveCustomEditor(
		document: vscode.CustomDocument,
		webviewPanel: vscode.WebviewPanel
	): Promise<void> {
		const pptxPath = document.uri.fsPath;
		const pptxData = fs.readFileSync(pptxPath).toString('base64');
		webviewPanel.webview.options = {
			enableScripts: true,
			localResourceRoots: [vscode.Uri.file(path.dirname(pptxPath))]
		};
		webviewPanel.webview.html = getWebviewContent(pptxData);
	}
}

function getWebviewContent(pptxBase64: string): string {
	// 构造本地资源 URI
	const jszipUri = vscode.Uri.joinPath(vscode.Uri.file(__dirname), '../media/jszip.min.js');
	const pptxjsUri = vscode.Uri.joinPath(vscode.Uri.file(__dirname), '../media/pptxjs.min.js');
	const pptxcssUri = vscode.Uri.joinPath(vscode.Uri.file(__dirname), '../media/pptxjs.min.css');
	return `<!DOCTYPE html>
<html lang="zh-cn">
<head>
	<meta charset="UTF-8">
	<title>PPT浏览器</title>
	<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; script-src 'unsafe-eval' 'unsafe-inline' vscode-resource:; style-src vscode-resource: 'unsafe-inline';">
	<script src="${jszipUri}"></script>
	<script src="${pptxjsUri}"></script>
	<link rel="stylesheet" href="${pptxcssUri}" />
	<style>body{margin:0;padding:0;}#ppt-container{width:100vw;height:90vh;overflow:auto;}</style>
</head>
<body>
	<div id="ppt-container">正在加载PPT...</div>
	<script>
	(function(){
		try {
			const pptxBase64 = "${pptxBase64}";
			const byteCharacters = atob(pptxBase64);
			const byteNumbers = new Array(byteCharacters.length);
			for (let i = 0; i < byteCharacters.length; i++) {
				byteNumbers[i] = byteCharacters.charCodeAt(i);
			}
			const byteArray = new Uint8Array(byteNumbers);
			const pptContainer = document.getElementById('ppt-container');
			window.PPTX = window.PPTX || {};
			PPTX.render = function(){
				try {
					window.PPTXJS.render({
						fileContent: byteArray.buffer,
						container: pptContainer,
						slideMode: true
					});
				} catch(e) {
					pptContainer.innerText = 'PPT 渲染失败: ' + e;
					console.error('PPT 渲染失败', e);
				}
			};
			if(window.PPTXJS){
				PPTX.render();
			}else{
				window.addEventListener('PPTXJSReady', PPTX.render);
			}
		} catch(e) {
			const pptContainer = document.getElementById('ppt-container');
			pptContainer.innerText = 'PPT 加载失败: ' + e;
			console.error('PPT 加载失败', e);
		}
	})();
	</script>
</body>
</html>`;
}

export function deactivate() {}
