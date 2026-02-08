// WASMモジュールの読み込み
let wasm;
let converter;

async function initWasm() {
    try {
        // wasm-packで生成されたモジュールを読み込む
        wasm = await import('./pkg/wasm_document_converter.js');
        await wasm.default();
        converter = new wasm.WasmConverter();

        const version = wasm.getVersion();
        console.log('WASM Converter initialized:', version);
        showStatus(`初期化完了: ${version}`, 'success');
        updateFontList();
    } catch (e) {
        console.error('WASM initialization failed:', e);
        showStatus('WASMモジュールの読み込みに失敗しました。「./build.sh build」を実行してビルドしてください。', 'error');
    }
}

// UI要素
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const fileSize = document.getElementById('fileSize');
const removeFile = document.getElementById('removeFile');
const convertBtn = document.getElementById('convertBtn');
const outputFormat = document.getElementById('outputFormat');
const dpiInput = document.getElementById('dpiInput');
const dpiGroup = document.getElementById('dpiGroup');
const progressBar = document.getElementById('progressBar');
const status = document.getElementById('status');

// フォント管理UI
const fontDropZone = document.getElementById('fontDropZone');
const fontFileInput = document.getElementById('fontFileInput');
const fontList = document.getElementById('fontList');
const fontUrl = document.getElementById('fontUrl');
const fontUrlName = document.getElementById('fontUrlName');
const loadFontUrl = document.getElementById('loadFontUrl');

let selectedFile = null;

// フォントリスト更新
function updateFontList() {
    if (!converter) return;
    try {
        const fonts = JSON.parse(converter.listFonts());
        const hasAny = converter.hasAnyFont();
        const extCount = converter.externalFontCount();
        let html = '';
        if (fonts.length > 0) {
            html += '<div style="font-size: 0.85rem; margin-bottom: 0.5rem;">';
            html += `<strong>読み込み済みフォント (${fonts.length}件, 外部: ${extCount}件):</strong>`;
            html += '</div>';
            html += '<div style="display: flex; flex-wrap: wrap; gap: 0.5rem;">';
            fonts.forEach(name => {
                html += `<span style="background: #e2e8f0; padding: 0.25rem 0.5rem; border-radius: 4px; font-size: 0.8rem;">${name}</span>`;
            });
            html += '</div>';
        } else {
            html = '<p style="font-size: 0.85rem; color: var(--text-muted);">フォント未読み込み（外部フォントを追加すると描画品質が向上します）</p>';
        }
        fontList.innerHTML = html;
    } catch (e) {
        console.error('Font list error:', e);
    }
}

// フォントファイル読み込み
fontDropZone.addEventListener('click', () => fontFileInput.click());
fontDropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    fontDropZone.classList.add('dragover');
});
fontDropZone.addEventListener('dragleave', () => {
    fontDropZone.classList.remove('dragover');
});
fontDropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    fontDropZone.classList.remove('dragover');
    for (const file of e.dataTransfer.files) {
        loadFontFile(file);
    }
});
fontFileInput.addEventListener('change', (e) => {
    for (const file of e.target.files) {
        loadFontFile(file);
    }
    fontFileInput.value = '';
});

async function loadFontFile(file) {
    if (!converter) {
        showStatus('WASMモジュールの初期化を待ってください。', 'error');
        return;
    }
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['ttf', 'otf', 'woff', 'woff2'].includes(ext)) {
        showStatus(`フォントファイル形式が不正です: ${file.name}`, 'error');
        return;
    }
    try {
        const buffer = await file.arrayBuffer();
        const data = new Uint8Array(buffer);
        const name = file.name.replace(/\.[^/.]+$/, '');
        converter.addFont(name, data);
        showStatus(`✅ フォント「${name}」を読み込みました（${formatFileSize(data.length)}）`, 'success');
        updateFontList();
    } catch (e) {
        showStatus(`フォント読み込みエラー: ${e}`, 'error');
    }
}

// URLからフォント読み込み
loadFontUrl.addEventListener('click', async () => {
    const url = fontUrl.value.trim();
    const name = fontUrlName.value.trim() || url.split('/').pop().replace(/\.[^/.]+$/, '');
    if (!url) {
        showStatus('フォントURLを入力してください。', 'error');
        return;
    }
    if (!converter) {
        showStatus('WASMモジュールの初期化を待ってください。', 'error');
        return;
    }
    try {
        showStatus(`フォント「${name}」をダウンロード中...`, 'info');
        const resp = await fetch(url);
        if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
        const buffer = await resp.arrayBuffer();
        const data = new Uint8Array(buffer);
        converter.addFont(name, data);
        showStatus(`✅ フォント「${name}」をURLから読み込みました（${formatFileSize(data.length)}）`, 'success');
        fontUrl.value = '';
        fontUrlName.value = '';
        updateFontList();
    } catch (e) {
        showStatus(`フォントダウンロードエラー: ${e}`, 'error');
    }
});

// ドラッグ＆ドロップ
dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
});
dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('dragover');
});
dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    if (e.dataTransfer.files.length > 0) {
        handleFile(e.dataTransfer.files[0]);
    }
});
fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFile(e.target.files[0]);
    }
});

// ファイル選択処理
function handleFile(file) {
    selectedFile = file;
    fileName.textContent = file.name;
    fileSize.textContent = formatFileSize(file.size);
    fileInfo.style.display = 'flex';
    dropZone.style.display = 'none';
    convertBtn.disabled = false;
    hideStatus();
}

removeFile.addEventListener('click', () => {
    selectedFile = null;
    fileInfo.style.display = 'none';
    dropZone.style.display = 'block';
    convertBtn.disabled = true;
    fileInput.value = '';
    hideStatus();
});

// 出力形式変更
outputFormat.addEventListener('change', () => {
    dpiGroup.style.display = outputFormat.value === 'images_zip' ? 'block' : 'none';
});
dpiGroup.style.display = 'none';

// 変換実行
convertBtn.addEventListener('click', async () => {
    if (!selectedFile || !converter) {
        showStatus('ファイルを選択するか、WASMモジュールの読み込みを待ってください。', 'error');
        return;
    }

    convertBtn.disabled = true;
    progressBar.classList.add('active');
    showStatus('変換中...', 'info');

    try {
        const arrayBuffer = await selectedFile.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);
        const format = outputFormat.value;
        let result;

        if (format === 'pdf') {
            result = converter.convertToPdf(selectedFile.name, data);
        } else {
            const dpi = parseFloat(dpiInput.value) || 150;
            result = converter.convertToImagesZip(selectedFile.name, data, dpi);
        }

        // ダウンロード
        const blob = new Blob([result], {
            type: format === 'pdf' ? 'application/pdf' : 'application/zip'
        });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        const baseName = selectedFile.name.replace(/\.[^/.]+$/, '');
        a.href = url;
        a.download = format === 'pdf' ? `${baseName}.pdf` : `${baseName}_pages.zip`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        const sizeStr = formatFileSize(result.length);
        showStatus(`✅ 変換完了！（${sizeStr}）ダウンロードが開始されます。`, 'success');
    } catch (e) {
        console.error('Conversion failed:', e);
        showStatus(`❌ 変換エラー: ${e}`, 'error');
    } finally {
        convertBtn.disabled = false;
        progressBar.classList.remove('active');
    }
});

function showStatus(msg, type) {
    status.textContent = msg;
    status.className = `status show ${type}`;
}

function hideStatus() {
    status.className = 'status';
}

function formatFileSize(bytes) {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

// 初期化
initWasm();
