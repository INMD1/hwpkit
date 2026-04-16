import { Pipeline } from 'hwpkit';
import type { AnyNode, DocRoot } from 'hwpkit';

// ─── DOM refs ───────────────────────────────────────────────
const inputEl    = document.getElementById('input')    as HTMLTextAreaElement;
const srcFmtEl   = document.getElementById('src-fmt')  as HTMLSelectElement;
const dstFmtEl   = document.getElementById('dst-fmt')  as HTMLSelectElement;
const runBtn     = document.getElementById('run-btn')!;
const inspectBtn = document.getElementById('inspect-btn')!;
const clearBtn   = document.getElementById('clear-btn')!;
const fileInput  = document.getElementById('file-input') as HTMLInputElement;
const outputText = document.getElementById('output-text')!;
const outputHtml = document.getElementById('output-html')!;
const outputTree = document.getElementById('output-tree')!;
const warnBar    = document.getElementById('warn-bar')!;
const statusEl   = document.getElementById('status')!;
const timingEl   = document.getElementById('timing')!;
const inputInfo  = document.getElementById('input-info')!;
const outputInfo = document.getElementById('output-info')!;

let rawBytes: Uint8Array | null = null;
let currentTab: 'text' | 'html' | 'tree' = 'text';

// ─── Demo 파일 로드 ──────────────────────────────────────────
async function loadDemoFile(path: string, fmt: string, label: string) {
  setStatus(`${label} 로드 중...`, false);
  try {
    const res = await fetch(path);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    rawBytes = new Uint8Array(await res.arrayBuffer());
    srcFmtEl.value = fmt;
    inputEl.value = `[데모 파일 로드됨: ${label} (${fmtBytes(rawBytes.length)})]`;
    inputEl.style.color = '#a78bfa';
    updateInputInfo();
    setStatus(`${label} 로드 완료`, false);
  } catch (e: any) {
    setStatus(`데모 파일 로드 실패: ${e.message}`, true);
  }
}

document.getElementById('demo-hwpx-btn')!.addEventListener('click', () =>
  loadDemoFile('/demo.hwpx', 'hwpx', 'demo.hwpx'));
document.getElementById('demo-hwp-btn')!.addEventListener('click', () =>
  loadDemoFile('/demo.hwp', 'hwp', 'demo.hwp'));

// ─── 파일 업로드 ─────────────────────────────────────────────
fileInput.addEventListener('change', async () => {
  const file = fileInput.files?.[0];
  if (!file) return;

  rawBytes = new Uint8Array(await file.arrayBuffer());
  const ext = file.name.split('.').pop()?.toLowerCase() ?? 'md';
  const fmt = ext === 'hwpx' ? 'hwpx' : ext === 'docx' ? 'docx' : ext === 'hwp' ? 'hwp' : 'md';
  srcFmtEl.value = fmt;

  inputEl.value = `[파일 로드됨: ${file.name} (${fmtBytes(rawBytes.length)})]`;
  inputEl.style.color = '#60a5fa';
  updateInputInfo();
  setStatus(`파일 로드: ${file.name}`, false);
});

inputEl.addEventListener('input', () => {
  rawBytes = null;
  inputEl.style.color = '';
  updateInputInfo();
});

// ─── 변환 ────────────────────────────────────────────────────
runBtn.addEventListener('click', async () => { await runConvert(); });
inspectBtn.addEventListener('click', async () => { await runInspect(); });

clearBtn.addEventListener('click', () => {
  inputEl.value = '';
  inputEl.style.color = '';
  rawBytes = null;
  clearOutput();
  statusEl.textContent = '초기화됨';
  statusEl.className = '';
  timingEl.textContent = '';
  updateInputInfo();
});

function clearOutput() {
  outputText.textContent = '';
  (outputHtml as HTMLIFrameElement).srcdoc = '';
  outputTree.innerHTML = '';
  outputInfo.textContent = '';
  warnBar.classList.remove('visible');
}

async function getInput(): Promise<{ data: Uint8Array | string; fmt: string }> {
  const fmt = srcFmtEl.value;
  if (rawBytes) return { data: rawBytes, fmt };
  return { data: inputEl.value, fmt: 'md' };
}

async function runConvert() {
  const t0 = performance.now();
  setStatus('변환 중...', false);
  clearOutput();

  try {
    const { data, fmt } = await getInput();
    const srcFmt = rawBytes ? fmt : 'md';
    const dstFmt = dstFmtEl.value;

    const pipeline = await Pipeline.openAsync(
      typeof data === 'string' ? data : data,
      srcFmt,
    );

    const result = await pipeline.to(dstFmt);
    const ms = (performance.now() - t0).toFixed(1);

    showWarns(result.warns ?? []);

    if (!result.ok) {
      showText(`오류: ${result.error}`);
      setStatus(`실패: ${result.error}`, true);
      timingEl.textContent = `${ms}ms`;
      return;
    }

    outputInfo.textContent = `${fmtBytes(result.data.length)}`;

    if (dstFmt === 'html') {
      const html = new TextDecoder().decode(result.data);
      showHtml(html);
    } else if (dstFmt === 'md') {
      showText(new TextDecoder().decode(result.data));
    } else {
      // Binary format: offer download
      const blob = new Blob([result.data]);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `output.${dstFmt}`;
      a.click();
      URL.revokeObjectURL(url);
      showText(`[${dstFmt.toUpperCase()} 파일 다운로드 완료]\n크기: ${fmtBytes(result.data.length)}\n\n다시 다운로드하려면 '변환 실행'을 눌러주세요.`);
    }

    setStatus(`변환 완료 (${srcFmt} → ${dstFmt})`, false);
    timingEl.textContent = `${ms}ms`;
  } catch (e: any) {
    const ms = (performance.now() - t0).toFixed(1);
    showText(`예외: ${e?.message ?? String(e)}\n\n${e?.stack ?? ''}`);
    setStatus('예외 발생', true);
    timingEl.textContent = `${ms}ms`;
  }
}

async function runInspect() {
  const t0 = performance.now();
  setStatus('검사 중...', false);
  clearOutput();

  try {
    const { data, fmt } = await getInput();
    const srcFmt = rawBytes ? fmt : 'md';

    const pipeline = await Pipeline.openAsync(
      typeof data === 'string' ? data : data,
      srcFmt,
    );

    const result = await pipeline.inspect();
    const ms = (performance.now() - t0).toFixed(1);

    showWarns(result.warns ?? []);

    if (!result.ok) {
      showText(`오류: ${result.error}`);
      setStatus(`실패: ${result.error}`, true);
      timingEl.textContent = `${ms}ms`;
      return;
    }

    renderTree(result.data);
    switchTab('tree');
    setStatus('DocRoot 검사 완료', false);
    timingEl.textContent = `${ms}ms`;
  } catch (e: any) {
    const ms = (performance.now() - t0).toFixed(1);
    showText(`예외: ${e?.message ?? String(e)}`);
    setStatus('예외 발생', true);
    timingEl.textContent = `${ms}ms`;
  }
}

// ─── 트리 렌더링 ──────────────────────────────────────────────
function renderTree(root: DocRoot) {
  outputTree.innerHTML = '';
  outputTree.appendChild(buildTreeEl(root as unknown as AnyNode));
}

function buildTreeEl(node: AnyNode): HTMLElement {
  const wrap = document.createElement('div');
  wrap.className = 'tree-node';

  const tag = document.createElement('span');
  tag.className = 'tree-tag';
  tag.textContent = `<${node.tag}`;
  wrap.appendChild(tag);

  if ('props' in node && node.props && Object.keys(node.props).length > 0) {
    const filtered = Object.entries(node.props).filter(([, v]) => v !== undefined && v !== null);
    if (filtered.length > 0) {
      const prop = document.createElement('span');
      prop.className = 'tree-prop';
      prop.textContent = ' ' + filtered.map(([k, v]) => `${k}=${JSON.stringify(v)}`).join(' ');
      wrap.appendChild(prop);
    }
  }

  const closeTag = document.createElement('span');
  closeTag.className = 'tree-tag';

  if (node.tag === 'txt') {
    const ct = document.createElement('span');
    ct.className = 'tree-content';
    ct.textContent = ` "${(node as any).text ?? (node as any).content ?? ''}"`;
    wrap.appendChild(ct);
    closeTag.textContent = ' />';
    wrap.appendChild(closeTag);
    return wrap;
  }

  if ('kids' in node && Array.isArray(node.kids) && node.kids.length > 0) {
    closeTag.textContent = '>';
    wrap.appendChild(closeTag);

    const children = document.createElement('div');
    children.className = 'tree-children';
    for (const child of node.kids as AnyNode[]) {
      children.appendChild(buildTreeEl(child));
    }
    wrap.appendChild(children);

    const endTag = document.createElement('div');
    endTag.className = 'tree-tag tree-node';
    endTag.textContent = `</${node.tag}>`;
    wrap.appendChild(endTag);
  } else {
    closeTag.textContent = ' />';
    wrap.appendChild(closeTag);
  }

  return wrap;
}

// ─── UI 헬퍼 ─────────────────────────────────────────────────
function showText(text: string) {
  outputText.textContent = text;
  if (currentTab !== 'text') switchTab('text');
}

function showHtml(html: string) {
  const iframe = outputHtml as HTMLIFrameElement;
  iframe.srcdoc = html;
  switchTab('html');
}

function showWarns(warns: string[]) {
  if (warns.length === 0) {
    warnBar.classList.remove('visible');
    return;
  }
  warnBar.textContent = warns.map(w => `⚠ ${w}`).join('\n');
  warnBar.classList.add('visible');
}

function setStatus(msg: string, isErr: boolean) {
  statusEl.textContent = msg;
  statusEl.className = isErr ? 'status-err' : 'status-ok';
}

function updateInputInfo() {
  const fmt = srcFmtEl.value;
  const bytes = rawBytes ? rawBytes.length : new TextEncoder().encode(inputEl.value).length;
  inputInfo.textContent = `${fmt} · ${fmtBytes(bytes)}`;
}

function fmtBytes(n: number): string {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / 1024 / 1024).toFixed(2)} MB`;
}

// ─── 탭 전환 ─────────────────────────────────────────────────
(window as any).switchTab = function (tab: 'text' | 'html' | 'tree') {
  currentTab = tab;
  outputText.classList.toggle('active', tab === 'text');
  outputHtml.classList.toggle('active', tab === 'html');
  outputTree.classList.toggle('active', tab === 'tree');
  document.getElementById('tab-text')!.classList.toggle('active', tab === 'text');
  document.getElementById('tab-html')!.classList.toggle('active', tab === 'html');
  document.getElementById('tab-tree')!.classList.toggle('active', tab === 'tree');
};

srcFmtEl.addEventListener('change', updateInputInfo);
updateInputInfo();
