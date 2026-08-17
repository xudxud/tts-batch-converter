import { useEffect, useRef, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import { listen } from "@tauri-apps/api/event";
import { confirm, open } from "@tauri-apps/plugin-dialog";
import "./App.css";

type RowStatus = "pending" | "running" | "success" | "failed" | "skipped" | "invalid";

interface ExcelRow {
  rowNumber: number;
  name: string;
  text: string;
  valid: boolean;
  error?: string;
  retryable: boolean;
  selected: boolean;
  exists: boolean;
  status: RowStatus;
  lastError?: string;
}

interface ExcelPreview {
  sheetName: string;
  rows: ExcelRow[];
  validCount: number;
  invalidCount: number;
}

interface Voice {
  label: string;
  id: string;
  locale: string;
  group: string;
  rate?: string;
}

interface VoiceSetting {
  id: string;
  locale: string;
  rate: string;
}

interface ProgressEvent {
  rowNumber: number;
  name: string;
  status: RowStatus;
  error?: string;
  retryable: boolean;
  processed: number;
  total: number;
}

interface BatchSummary {
  success: number;
  failed: number;
  skipped: number;
  cancelled: boolean;
}

const voices: Voice[] = [
  { group: "普通话", label: "晓晓 · 自然温暖", id: "zh-CN-XiaoxiaoNeural", locale: "zh-CN" },
  { group: "普通话", label: "晓伊 · 活泼明亮", id: "zh-CN-XiaoyiNeural", locale: "zh-CN" },
  { group: "普通话", label: "云希 · 年轻自然", id: "zh-CN-YunxiNeural", locale: "zh-CN" },
  { group: "普通话", label: "云健 · 成熟有力", id: "zh-CN-YunjianNeural", locale: "zh-CN" },
  { group: "普通话", label: "云扬 · 专业播报", id: "zh-CN-YunyangNeural", locale: "zh-CN" },
  { group: "普通话", label: "云夏 · 亲切活泼", id: "zh-CN-YunxiaNeural", locale: "zh-CN" },
  { group: "台湾国语", label: "曉臻 · 柔和关怀", id: "zh-TW-HsiaoChenNeural", locale: "zh-TW" },
  { group: "台湾国语", label: "曉雨 · 清晰明亮", id: "zh-TW-HsiaoYuNeural", locale: "zh-TW" },
  { group: "台湾国语", label: "雲哲 · 温和自然", id: "zh-TW-YunJheNeural", locale: "zh-TW" },
  { group: "儿童感", label: "云夏 · 普通话男声", id: "zh-CN-YunxiaNeural", locale: "zh-CN", rate: "+5%" },
  { group: "儿童感", label: "Ana · 美式女声", id: "en-US-AnaNeural", locale: "en-US" },
  { group: "儿童感", label: "Maisie · 英式女声", id: "en-GB-MaisieNeural", locale: "en-GB" },
  { group: "美式英语", label: "Jenny · 亲切自然", id: "en-US-JennyNeural", locale: "en-US" },
  { group: "美式英语", label: "Aria · 清晰明亮", id: "en-US-AriaNeural", locale: "en-US" },
  { group: "美式英语", label: "Ava · 友好温柔", id: "en-US-AvaNeural", locale: "en-US" },
  { group: "美式英语", label: "Michelle · 成熟自信", id: "en-US-MichelleNeural", locale: "en-US" },
  { group: "美式英语", label: "Guy · 轻松友好", id: "en-US-GuyNeural", locale: "en-US" },
  { group: "美式英语", label: "Andrew · 温暖自信", id: "en-US-AndrewNeural", locale: "en-US" },
  { group: "美式英语", label: "Brian · 沉稳亲切", id: "en-US-BrianNeural", locale: "en-US" },
  { group: "美式英语", label: "Christopher · 低沉温暖", id: "en-US-ChristopherNeural", locale: "en-US" },
  { group: "英式英语", label: "Sonia · 轻柔", id: "en-GB-SoniaNeural", locale: "en-GB" },
  { group: "英式英语", label: "Libby · 清晰明亮", id: "en-GB-LibbyNeural", locale: "en-GB" },
  { group: "英式英语", label: "Ryan · 明快自然", id: "en-GB-RyanNeural", locale: "en-GB" },
  { group: "英式英语", label: "Thomas · 自然沉稳", id: "en-GB-ThomasNeural", locale: "en-GB" },
];

const statusText: Record<RowStatus, string> = {
  pending: "待生成",
  running: "生成中",
  success: "已完成",
  failed: "失败",
  skipped: "已跳过",
  invalid: "数据错误",
};

function basename(path: string) {
  return path.split(/[\\/]/).pop() || path;
}

function defaultOutput(path: string) {
  const slash = Math.max(path.lastIndexOf("\\"), path.lastIndexOf("/"));
  const folder = slash >= 0 ? path.slice(0, slash + 1) : "";
  const filename = path.slice(slash + 1).replace(/\.xlsx$/i, "");
  return `${folder}${filename}_MP3`;
}

function parseRanges(value: string): Set<number> {
  const result = new Set<number>();
  if (!value.trim()) return result;
  for (const part of value.split(/[，,]/)) {
    const token = part.trim();
    const single = token.match(/^\d+$/);
    const range = token.match(/^(\d+)\s*-\s*(\d+)$/);
    if (single) result.add(Number(token));
    else if (range) {
      const start = Number(range[1]);
      const end = Number(range[2]);
      if (start > end || end - start > 100000) throw new Error(`无效范围：${token}`);
      for (let row = start; row <= end; row += 1) result.add(row);
    } else throw new Error(`无法识别：${token}`);
  }
  return result;
}

function App() {
  const [excelPath, setExcelPath] = useState("");
  const [outputPath, setOutputPath] = useState("");
  const [rows, setRows] = useState<ExcelRow[]>([]);
  const [sheetName, setSheetName] = useState("");
  const [voiceIndex, setVoiceIndex] = useState(2);
  const [rate, setRate] = useState(0);
  const [rangeText, setRangeText] = useState("");
  const [filter, setFilter] = useState<"all" | RowStatus>("all");
  const [firstRowIsHeader, setFirstRowIsHeader] = useState(() => localStorage.getItem("firstRowIsHeader") === "true");
  const [previewText, setPreviewText] = useState("这是一段语音试听文本。This is a voice preview.");
  const [previewing, setPreviewing] = useState(false);
  const [running, setRunning] = useState(false);
  const [processed, setProcessed] = useState(0);
  const [total, setTotal] = useState(0);
  const [notice, setNotice] = useState("选择一个 Excel 文件开始");
  const audioRef = useRef<HTMLAudioElement>(null);

  const selectedVoice = voices[voiceIndex] || voices[2];
  const selectedRows = rows.filter((row) => row.selected && row.valid);
  const failedRows = rows.filter((row) => row.status === "failed" && row.retryable);
  const visibleRows = filter === "all" ? rows : rows.filter((row) => row.status === filter);
  const successCount = rows.filter((row) => row.status === "success").length;
  const skippedCount = rows.filter((row) => row.status === "skipped").length;

  useEffect(() => {
    const setup = async () => {
      const unlisten = await listen<ProgressEvent>("tts-progress", ({ payload }) => {
        setRows((current) => current.map((row) => row.rowNumber === payload.rowNumber
          ? { ...row, status: payload.status, lastError: payload.error, retryable: payload.retryable }
          : row));
        setProcessed(payload.processed);
        setTotal(payload.total);
        setNotice(payload.status === "running" ? `正在生成 ${payload.name}.mp3` : `已处理 ${payload.processed} / ${payload.total}`);
      });
      return unlisten;
    };
    let cleanup: (() => void) | undefined;
    setup().then((value) => { cleanup = value; });
    return () => cleanup?.();
  }, []);

  useEffect(() => {
    const savedExcel = localStorage.getItem("excelPath") || "";
    const savedOutput = localStorage.getItem("outputPath") || "";
    if (savedExcel && savedOutput) {
      setExcelPath(savedExcel);
      setOutputPath(savedOutput);
      loadExcel(savedExcel, savedOutput);
    }
  }, []);

  function voiceSetting(): VoiceSetting {
    return { id: selectedVoice.id, locale: selectedVoice.locale, rate: `${rate >= 0 ? "+" : ""}${rate}%` };
  }

  async function chooseExcel() {
    const selected = await open({ multiple: false, filters: [{ name: "Excel 工作簿", extensions: ["xlsx"] }] });
    if (!selected) return;
    const path = String(selected);
    const output = outputPath || defaultOutput(path);
    setExcelPath(path);
    setOutputPath(output);
    localStorage.setItem("excelPath", path);
    localStorage.setItem("outputPath", output);
    await loadExcel(path, output);
  }

  async function chooseOutput() {
    const selected = await open({ directory: true, multiple: false });
    if (!selected) return;
    const output = String(selected);
    setOutputPath(output);
    localStorage.setItem("outputPath", output);
    if (excelPath) await loadExcel(excelPath, output);
  }

  async function loadExcel(path = excelPath, output = outputPath, header = firstRowIsHeader) {
    try {
      setNotice("正在读取 Excel...");
      const preview = await invoke<ExcelPreview>("load_excel", {
        path,
        outputPath: output || null,
        firstRowIsHeader: header,
      });
      setRows(preview.rows);
      setSheetName(preview.sheetName);
      setProcessed(0);
      setTotal(preview.validCount);
      setNotice(`已读取 ${preview.validCount} 条有效数据${preview.invalidCount ? `，${preview.invalidCount} 条需修正` : ""}`);
    } catch (error) {
      setNotice(String(error));
      await confirm(String(error), { title: "读取失败", kind: "error", okLabel: "知道了", cancelLabel: "关闭" });
    }
  }

  function setSelection(mode: "all" | "none" | "invert" | "failed") {
    setRows((current) => current.map((row) => ({
      ...row,
      selected: row.valid && (mode === "all" || (mode === "invert" ? !row.selected : mode === "failed" && row.status === "failed")),
    })));
  }

  function applyRange() {
    try {
      const wanted = parseRanges(rangeText);
      setRows((current) => current.map((row) => ({ ...row, selected: row.valid && wanted.has(row.rowNumber) })));
      setNotice(`已按 Excel 行号选择 ${[...wanted].length} 个位置`);
    } catch (error) {
      setNotice(String(error));
    }
  }

  async function previewVoice() {
    if (!previewText.trim()) return setNotice("请先输入试听文字");
    setPreviewing(true);
    setNotice("正在生成试听...");
    try {
      const source = await invoke<string>("preview_speech", { request: { text: previewText, voice: voiceSetting() } });
      if (audioRef.current) {
        audioRef.current.src = source;
        await audioRef.current.play();
      }
      setNotice(`正在试听：${selectedVoice.label}`);
    } catch (error) {
      setNotice(`试听失败：${String(error)}`);
    } finally {
      setPreviewing(false);
    }
  }

  function stopPreview() {
    audioRef.current?.pause();
    if (audioRef.current) audioRef.current.currentTime = 0;
    setNotice("试听已停止");
  }

  async function startBatch(targetRows = selectedRows) {
    if (!excelPath || !outputPath) return setNotice("请选择 Excel 和输出文件夹");
    if (!targetRows.length) return setNotice("请至少选择一条有效数据");
    const overwriteCount = targetRows.filter((row) => row.exists || row.status === "success").length;
    let overwrite = false;
    if (overwriteCount) {
      overwrite = await confirm(`选中的行中有 ${overwriteCount} 个同名 MP3 已存在。\n\n确认后仅覆盖这些选中项，生成失败时原文件会保留。`, {
        title: "确认覆盖", kind: "warning", okLabel: "确认并生成", cancelLabel: "取消",
      });
      if (!overwrite) return;
    }
    setRunning(true);
    setProcessed(0);
    setTotal(targetRows.length);
    setNotice(`准备生成 ${targetRows.length} 条语音`);
    try {
      const summary = await invoke<BatchSummary>("run_batch", {
        request: { excelPath, outputPath, rows: targetRows, snapshotRows: rows, voice: voiceSetting(), overwrite },
      });
      setNotice(summary.cancelled
        ? `任务已停止：成功 ${summary.success}，失败 ${summary.failed}`
        : `转换完成：成功 ${summary.success}，失败 ${summary.failed}，跳过 ${summary.skipped}`);
    } catch (error) {
      setNotice(`转换失败：${String(error)}`);
    } finally {
      setRunning(false);
    }
  }

  async function retryFailed() {
    if (!failedRows.length) return;
    setRows((current) => current.map((row) => ({ ...row, selected: row.status === "failed" })));
    await startBatch(failedRows);
  }

  async function stopBatch() {
    await invoke("cancel_batch");
    setNotice("将在当前文件完成后停止...");
  }

  return (
    <div className="app-shell">
      <audio ref={audioRef} />
      <header className="topbar">
        <div className="brand-mark"><img src="/app-icon.svg" alt="" /></div>
        <div className="brand-copy"><h1>TTS批量转换工具</h1><p>Excel 批量文字转语音</p></div>
        <div className={`network-pill ${running ? "active" : ""}`}><i />{running ? "正在生成" : "在线语音"}</div>
      </header>

      <main>
        <section className="file-strip card">
          <button className="file-target" onClick={chooseExcel} disabled={running}>
            <span className="file-icon">X</span><span><b>{excelPath ? basename(excelPath) : "选择 Excel 文件"}</b><small>{excelPath || "支持 .xlsx，A 列文件名，B 列文字"}</small></span>
          </button>
          <span className="flow-arrow">→</span>
          <button className="file-target" onClick={chooseOutput} disabled={running}>
            <span className="folder-icon">⌁</span><span><b>{outputPath ? basename(outputPath) : "选择输出文件夹"}</b><small>{outputPath || "MP3 将保存到这里"}</small></span>
          </button>
          <button className="ghost-button" onClick={() => loadExcel()} disabled={!excelPath || running}>重新读取</button>
          <label className="header-toggle">
            <input
              type="checkbox"
              checked={firstRowIsHeader}
              disabled={running}
              onChange={(event) => {
                const checked = event.target.checked;
                setFirstRowIsHeader(checked);
                localStorage.setItem("firstRowIsHeader", String(checked));
                if (excelPath) loadExcel(excelPath, outputPath, checked);
              }}
            />
            第一行是表头
          </label>
        </section>

        <div className="workspace-grid">
          <aside className="settings card">
            <div className="section-title"><span>01</span><div><h3>声音设置</h3><p>选择最适合内容的声音</p></div></div>
            <label>音色</label>
            <select value={voiceIndex} onChange={(event) => {
              const index = Number(event.target.value);
              const voice = voices[index];
              setVoiceIndex(index);
              if (voice?.rate) setRate(Number.parseInt(voice.rate));
            }} disabled={running}>
              {[...new Set(voices.map((voice) => voice.group))].map((group) => (
                <optgroup key={group} label={group}>{voices.map((voice, index) => ({ voice, index })).filter((item) => item.voice.group === group).map(({ voice, index }) => (
                  <option key={`${group}-${voice.id}`} value={index}>{voice.label}</option>
                ))}</optgroup>
              ))}
            </select>
            <div className="voice-meta"><span>{selectedVoice.group}</span><span>{selectedVoice.locale}</span><span>{selectedVoice.id.includes("Xiao") || selectedVoice.id.includes("Hsiao") || ["Ana", "Maisie", "Jenny", "Aria", "Ava", "Michelle", "Sonia", "Libby"].some((name) => selectedVoice.id.includes(name)) ? "女声" : "男声"}</span></div>

            <div className="range-label"><label>语速</label><b>{rate >= 0 ? "+" : ""}{rate}%</b></div>
            <input type="range" min="-30" max="30" step="5" value={rate} onChange={(event) => setRate(Number(event.target.value))} disabled={running} />
            <div className="range-ticks"><span>慢</span><span>自然</span><span>快</span></div>

            <label>试听文字</label>
            <textarea value={previewText} onChange={(event) => setPreviewText(event.target.value)} rows={4} disabled={running} />
            <div className="preview-actions">
              <button className="secondary-button" onClick={previewVoice} disabled={previewing || running}>{previewing ? "生成中..." : "▶ 生成并试听"}</button>
              <button className="icon-button" onClick={stopPreview} title="停止试听">■</button>
            </div>
            <p className="fine-print">Edge 在线语音服务，需保持网络连接。纯 Rust 版本当前仅支持语速调节。</p>
          </aside>

          <section className="data-panel card">
            <div className="data-head">
              <div className="section-title"><span>02</span><div><h3>选择生成内容</h3><p>{sheetName ? `工作表：${sheetName}` : "读取 Excel 后在这里预览"}</p></div></div>
              <div className="count-badge"><b>{selectedRows.length}</b> 条已选</div>
            </div>

            <div className="selection-tools">
              <div className="button-cluster">
                <button onClick={() => setSelection("all")}>全选</button><button onClick={() => setSelection("none")}>取消</button><button onClick={() => setSelection("invert")}>反选</button>
              </div>
              <div className="range-input"><span>Excel 行号</span><input value={rangeText} onChange={(event) => setRangeText(event.target.value)} placeholder="例如 2-10, 15, 20-25" onKeyDown={(event) => event.key === "Enter" && applyRange()} /><button onClick={applyRange}>应用</button></div>
            </div>

            <div className="filter-row">
              {(["all", "pending", "success", "failed", "skipped", "invalid"] as const).map((item) => (
                <button key={item} className={filter === item ? "selected" : ""} onClick={() => setFilter(item)}>{item === "all" ? "全部" : statusText[item]}</button>
              ))}
            </div>

            <div className="table-wrap">
              <table>
                <thead><tr><th className="check-col"></th><th>行</th><th>MP3 文件名</th><th>文字内容</th><th>状态</th><th></th></tr></thead>
                <tbody>{visibleRows.length ? visibleRows.map((row) => (
                  <tr key={`${row.rowNumber}-${row.name}`} className={!row.valid ? "invalid-row" : ""}>
                    <td><input type="checkbox" checked={row.selected} disabled={!row.valid || running} onChange={() => setRows((current) => current.map((item) => item.rowNumber === row.rowNumber ? { ...item, selected: !item.selected } : item))} /></td>
                    <td className="row-number">{row.rowNumber}</td>
                    <td><strong>{row.name || "未填写"}</strong><small>.mp3</small></td>
                    <td className="text-cell" title={row.text}>{row.text || row.error}</td>
                    <td><span className={`status status-${row.status}`}><i />{statusText[row.status]}</span>{(row.lastError || row.error) && <small className="error-text" title={row.lastError || row.error}>{row.lastError || row.error}</small>}</td>
                    <td>{row.status === "failed" && row.retryable && <button className="retry-one" onClick={() => startBatch([row])} disabled={running}>重试</button>}</td>
                  </tr>
                )) : <tr><td colSpan={6} className="empty-state"><b>还没有数据</b><span>选择 Excel 文件后，将在这里预览每一行</span></td></tr>}</tbody>
              </table>
            </div>
          </section>
        </div>

        <section className="run-panel card">
          <div className="progress-copy"><div className="progress-top"><b>{notice}</b><span>{total ? Math.round((processed / total) * 100) : 0}%</span></div><div className="progress-track"><i style={{ width: `${total ? (processed / total) * 100 : 0}%` }} /></div><div className="stats"><span><i className="green" />完成 {successCount}</span><span><i className="red" />失败 {failedRows.length}</span><span><i className="gray" />跳过 {skippedCount}</span></div></div>
          <div className="run-actions">
            {failedRows.length > 0 && <button className="retry-button" onClick={retryFailed} disabled={running}>↻ 重试失败项 ({failedRows.length})</button>}
            {running ? <button className="stop-button" onClick={stopBatch}>停止任务</button> : <button className="primary-button" onClick={() => startBatch()} disabled={!selectedRows.length}>开始生成 <span>→</span></button>}
            <button className="open-button" onClick={() => invoke("open_output_folder", { path: outputPath })} disabled={!outputPath}>打开目录</button>
          </div>
        </section>
      </main>
    </div>
  );
}

export default App;
