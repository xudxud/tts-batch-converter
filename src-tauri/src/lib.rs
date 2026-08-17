use std::{
    collections::{HashMap, HashSet},
    fs::{self, File},
    io::Read,
    path::{Path, PathBuf},
    sync::{
        atomic::{AtomicBool, Ordering},
        Arc, OnceLock,
    },
    time::Duration,
};

use base64::{engine::general_purpose::STANDARD as BASE64, Engine as _};
use calamine::{open_workbook, Data, Reader, Xlsx};
use chrono::Utc;
use msedge_tts::{
    tts::{client::connect, SpeechConfig},
    voice::get_voices_list,
};
use quick_xml::{events::Event, Reader as XmlReader};
use serde::{Deserialize, Serialize};
use tauri::{AppHandle, Emitter, State};
use zip::ZipArchive;

static VOICE_CACHE: OnceLock<Vec<msedge_tts::voice::Voice>> = OnceLock::new();

const TASK_FILE: &str = ".excel-tts-task.json";
const REQUEST_TIMEOUT_SECS: u64 = 45;

#[derive(Default)]
struct AppState {
    cancelled: Arc<AtomicBool>,
    running: AtomicBool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ExcelRow {
    row_number: u32,
    name: String,
    text: String,
    valid: bool,
    error: Option<String>,
    selected: bool,
    exists: bool,
    status: String,
    last_error: Option<String>,
    #[serde(default)]
    retryable: bool,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct ExcelPreview {
    sheet_name: String,
    rows: Vec<ExcelRow>,
    valid_count: usize,
    invalid_count: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct VoiceSetting {
    id: String,
    locale: String,
    rate: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct BatchRequest {
    excel_path: String,
    output_path: String,
    rows: Vec<ExcelRow>,
    snapshot_rows: Vec<ExcelRow>,
    voice: VoiceSetting,
    overwrite: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct TaskRecord {
    version: u8,
    excel_path: String,
    output_path: String,
    voice: VoiceSetting,
    updated_at: String,
    rows: Vec<ExcelRow>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct ProgressEvent {
    row_number: u32,
    name: String,
    status: String,
    error: Option<String>,
    retryable: bool,
    processed: usize,
    total: usize,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct BatchSummary {
    success: usize,
    failed: usize,
    skipped: usize,
    cancelled: bool,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct PreviewRequest {
    text: String,
    voice: VoiceSetting,
}

fn data_to_string(value: Option<&Data>) -> String {
    match value {
        None | Some(Data::Empty) => String::new(),
        Some(Data::String(value)) => value.clone(),
        Some(Data::Int(value)) => value.to_string(),
        Some(Data::Float(value)) if value.fract() == 0.0 => (*value as i64).to_string(),
        Some(Data::Float(value)) => value.to_string(),
        Some(Data::Bool(value)) => value.to_string(),
        Some(value) => value.to_string(),
    }
}

fn zip_text(archive: &mut ZipArchive<File>, name: &str) -> Option<String> {
    let mut file = archive.by_name(name).ok()?;
    let mut text = String::new();
    file.read_to_string(&mut text).ok()?;
    Some(text)
}

fn first_sheet_path(archive: &mut ZipArchive<File>) -> Option<String> {
    let workbook = zip_text(archive, "xl/workbook.xml")?;
    let mut reader = XmlReader::from_str(&workbook);
    let mut relation_id = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(event)) | Ok(Event::Empty(event))
                if event.local_name().as_ref() == b"sheet" =>
            {
                relation_id = event.attributes().flatten().find_map(|attribute| {
                    (attribute.key.local_name().as_ref() == b"id")
                        .then(|| String::from_utf8_lossy(attribute.value.as_ref()).into_owned())
                });
                break;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    let relation_id = relation_id?;
    let relationships = zip_text(archive, "xl/_rels/workbook.xml.rels")?;
    let mut reader = XmlReader::from_str(&relationships);
    loop {
        match reader.read_event() {
            Ok(Event::Start(event)) | Ok(Event::Empty(event))
                if event.local_name().as_ref() == b"Relationship" =>
            {
                let mut id = None;
                let mut target = None;
                for attribute in event.attributes().flatten() {
                    match attribute.key.local_name().as_ref() {
                        b"Id" => {
                            id =
                                Some(String::from_utf8_lossy(attribute.value.as_ref()).into_owned())
                        }
                        b"Target" => {
                            target =
                                Some(String::from_utf8_lossy(attribute.value.as_ref()).into_owned())
                        }
                        _ => {}
                    }
                }
                if id.as_deref() == Some(&relation_id) {
                    let target = target?;
                    return Some(if target.starts_with('/') {
                        target.trim_start_matches('/').to_string()
                    } else {
                        format!("xl/{}", target.trim_start_matches("../"))
                    });
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    None
}

fn zero_padded_names(path: &str) -> HashMap<u32, String> {
    let Ok(file) = File::open(path) else {
        return HashMap::new();
    };
    let Ok(mut archive) = ZipArchive::new(file) else {
        return HashMap::new();
    };
    let Some(styles) = zip_text(&mut archive, "xl/styles.xml") else {
        return HashMap::new();
    };
    let Some(sheet_path) = first_sheet_path(&mut archive) else {
        return HashMap::new();
    };
    let Some(sheet) = zip_text(&mut archive, &sheet_path) else {
        return HashMap::new();
    };

    let mut formats: HashMap<u32, String> = HashMap::new();
    let mut style_formats = Vec::new();
    let mut in_cell_xfs = false;
    let mut reader = XmlReader::from_str(&styles);
    loop {
        match reader.read_event() {
            Ok(Event::Start(event)) if event.local_name().as_ref() == b"cellXfs" => {
                in_cell_xfs = true
            }
            Ok(Event::End(event)) if event.local_name().as_ref() == b"cellXfs" => {
                in_cell_xfs = false
            }
            Ok(Event::Start(event)) | Ok(Event::Empty(event))
                if event.local_name().as_ref() == b"numFmt" =>
            {
                let mut id = None;
                let mut code = None;
                for attribute in event.attributes().flatten() {
                    match attribute.key.local_name().as_ref() {
                        b"numFmtId" => {
                            id = String::from_utf8_lossy(attribute.value.as_ref())
                                .parse()
                                .ok()
                        }
                        b"formatCode" => {
                            code =
                                Some(String::from_utf8_lossy(attribute.value.as_ref()).into_owned())
                        }
                        _ => {}
                    }
                }
                if let (Some(id), Some(code)) = (id, code) {
                    formats.insert(id, code);
                }
            }
            Ok(Event::Start(event)) | Ok(Event::Empty(event))
                if in_cell_xfs && event.local_name().as_ref() == b"xf" =>
            {
                let id = event
                    .attributes()
                    .flatten()
                    .find_map(|attribute| {
                        (attribute.key.local_name().as_ref() == b"numFmtId")
                            .then(|| {
                                String::from_utf8_lossy(attribute.value.as_ref())
                                    .parse::<u32>()
                                    .ok()
                            })
                            .flatten()
                    })
                    .unwrap_or(0);
                style_formats.push(id);
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    let mut result = HashMap::new();
    let mut reader = XmlReader::from_str(&sheet);
    let mut current: Option<(u32, usize)> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(event)) if event.local_name().as_ref() == b"c" => {
                let mut cell_ref = None;
                let mut style = 0usize;
                for attribute in event.attributes().flatten() {
                    match attribute.key.local_name().as_ref() {
                        b"r" => {
                            cell_ref =
                                Some(String::from_utf8_lossy(attribute.value.as_ref()).into_owned())
                        }
                        b"s" => {
                            style = String::from_utf8_lossy(attribute.value.as_ref())
                                .parse()
                                .unwrap_or(0)
                        }
                        _ => {}
                    }
                }
                current = cell_ref.and_then(|reference| {
                    reference
                        .strip_prefix('A')?
                        .parse::<u32>()
                        .ok()
                        .map(|row| (row, style))
                });
            }
            Ok(Event::Text(text)) if current.is_some() => {
                let (row, style) = current.unwrap();
                let format_id = style_formats.get(style).copied().unwrap_or(0);
                if let Some(format) = formats
                    .get(&format_id)
                    .filter(|value| !value.is_empty() && value.chars().all(|c| c == '0'))
                {
                    let raw = String::from_utf8_lossy(text.as_ref());
                    if let Ok(number) = raw.parse::<f64>() {
                        if number.fract() == 0.0 {
                            result.insert(
                                row,
                                format!("{:0width$}", number as i64, width = format.len()),
                            );
                        }
                    }
                }
            }
            Ok(Event::End(event)) if event.local_name().as_ref() == b"c" => current = None,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    result
}

fn filename_error(name: &str) -> Option<String> {
    if name.is_empty() {
        return Some("序号为空".into());
    }
    if name.ends_with(' ') || name.ends_with('.') {
        return Some("文件名不能以空格或句点结尾".into());
    }
    if name
        .chars()
        .any(|c| matches!(c, '\\' | '/' | ':' | '*' | '?' | '"' | '<' | '>' | '|'))
    {
        return Some("文件名含有 Windows 不允许的字符".into());
    }
    let stem = name
        .split('.')
        .next()
        .unwrap_or_default()
        .to_ascii_uppercase();
    let reserved = matches!(stem.as_str(), "CON" | "PRN" | "AUX" | "NUL")
        || (stem.len() == 4
            && (stem.starts_with("COM") || stem.starts_with("LPT"))
            && stem[3..].parse::<u8>().is_ok_and(|n| (1..=9).contains(&n)));
    if reserved {
        return Some("文件名是 Windows 保留名称".into());
    }
    if name.chars().count() > 200 {
        return Some("文件名过长（最多 200 个字符）".into());
    }
    None
}

fn task_path(output_path: &str) -> PathBuf {
    Path::new(output_path).join(TASK_FILE)
}

fn save_task(record: &TaskRecord) -> Result<(), String> {
    fs::create_dir_all(&record.output_path).map_err(|e| e.to_string())?;
    let path = task_path(&record.output_path);
    let temp = path.with_extension("json.tmp");
    let bytes = serde_json::to_vec_pretty(record).map_err(|e| e.to_string())?;
    fs::write(&temp, bytes).map_err(|e| e.to_string())?;
    fs::rename(&temp, &path).map_err(|e| e.to_string())
}

fn apply_saved_status(rows: &mut [ExcelRow], output_path: &str) {
    let Ok(bytes) = fs::read(task_path(output_path)) else {
        return;
    };
    let Ok(saved) = serde_json::from_slice::<TaskRecord>(&bytes) else {
        return;
    };
    for row in rows {
        if let Some(previous) = saved.rows.iter().find(|item| {
            item.row_number == row.row_number && item.name == row.name && item.text == row.text
        }) {
            row.status = previous.status.clone();
            row.last_error = previous.last_error.clone();
            row.retryable = previous.retryable;
        }
    }
}

#[tauri::command]
fn load_excel(
    path: String,
    output_path: Option<String>,
    first_row_is_header: bool,
) -> Result<ExcelPreview, String> {
    let padded_names = zero_padded_names(&path);
    let mut workbook: Xlsx<_> = open_workbook(&path).map_err(|e| format!("无法打开 Excel：{e}"))?;
    let sheet_name = workbook
        .sheet_names()
        .first()
        .cloned()
        .ok_or_else(|| "Excel 中没有工作表".to_string())?;
    let range = workbook
        .worksheet_range(&sheet_name)
        .map_err(|e| format!("无法读取第一个工作表：{e}"))?;

    let mut rows = Vec::new();
    let mut seen = HashSet::new();
    for (index, row) in range
        .rows()
        .enumerate()
        .skip(usize::from(first_row_is_header))
    {
        let row_number = index as u32 + 1;
        let name = padded_names
            .get(&row_number)
            .cloned()
            .unwrap_or_else(|| data_to_string(row.first()));
        let text = data_to_string(row.get(1));
        if name.is_empty() && text.trim().is_empty() {
            continue;
        }
        let mut error = filename_error(&name);
        if error.is_none() && text.trim().is_empty() {
            error = Some("文字内容为空".into());
        }
        if error.is_none() && !seen.insert(name.to_lowercase()) {
            error = Some("序号重复".into());
        }
        let valid = error.is_none();
        let exists = output_path
            .as_ref()
            .is_some_and(|output| Path::new(output).join(format!("{name}.mp3")).exists());
        rows.push(ExcelRow {
            row_number,
            name,
            text,
            valid,
            error,
            selected: valid,
            exists,
            status: if valid { "pending" } else { "invalid" }.into(),
            last_error: None,
            retryable: false,
        });
    }
    if let Some(output) = &output_path {
        apply_saved_status(&mut rows, output);
    }
    let valid_count = rows.iter().filter(|row| row.valid).count();
    Ok(ExcelPreview {
        sheet_name,
        invalid_count: rows.len() - valid_count,
        valid_count,
        rows,
    })
}

async fn synthesize(text: &str, voice: &VoiceSetting) -> Result<Vec<u8>, String> {
    let text = text.to_string();
    let voice = voice.clone();
    let audio = tokio::time::timeout(
        Duration::from_secs(REQUEST_TIMEOUT_SECS),
        tokio::task::spawn_blocking(move || {
            if VOICE_CACHE.get().is_none() {
                let voices = get_voices_list().map_err(|e| format!("获取音色失败：{e}"))?;
                let _ = VOICE_CACHE.set(voices);
            }
            let voices = VOICE_CACHE
                .get()
                .ok_or_else(|| "在线音色列表不可用".to_string())?;
            let selected = voices
                .iter()
                .find(|item| item.short_name.as_deref() == Some(&voice.id))
                .ok_or_else(|| format!("在线音色列表中找不到 {}", voice.id))?;
            let mut config = SpeechConfig::from(selected);
            config.rate = voice.rate.trim_end_matches('%').parse().unwrap_or(0);
            let mut client = connect().map_err(|e| format!("连接语音服务失败：{e}"))?;
            client
                .synthesize(&text, &config)
                .map(|result| result.audio_bytes)
                .map_err(|e| format!("语音生成失败：{e}"))
        }),
    )
    .await
    .map_err(|_| format!("请求超时（超过 {REQUEST_TIMEOUT_SECS} 秒）"))?
    .map_err(|e| format!("语音任务异常：{e}"))??;
    if audio.is_empty() {
        Err("服务未返回音频".into())
    } else {
        Ok(audio)
    }
}

#[tauri::command]
async fn preview_speech(request: PreviewRequest) -> Result<String, String> {
    if request.text.trim().is_empty() {
        return Err("请输入试听文字".into());
    }
    let audio = synthesize(request.text.trim(), &request.voice).await?;
    Ok(format!("data:audio/mpeg;base64,{}", BASE64.encode(audio)))
}

#[tauri::command]
fn load_saved_task(output_path: String) -> Result<Option<TaskRecord>, String> {
    let path = task_path(&output_path);
    if !path.exists() {
        return Ok(None);
    }
    let bytes = fs::read(path).map_err(|e| e.to_string())?;
    serde_json::from_slice(&bytes)
        .map(Some)
        .map_err(|e| format!("任务记录已损坏：{e}"))
}

#[tauri::command]
fn cancel_batch(state: State<'_, AppState>) {
    state.cancelled.store(true, Ordering::Relaxed);
}

#[tauri::command]
fn open_output_folder(path: String) -> Result<(), String> {
    fs::create_dir_all(&path).map_err(|e| e.to_string())?;
    std::process::Command::new("explorer")
        .arg(&path)
        .spawn()
        .map(|_| ())
        .map_err(|e| e.to_string())
}

#[tauri::command]
async fn run_batch(
    app: AppHandle,
    state: State<'_, AppState>,
    request: BatchRequest,
) -> Result<BatchSummary, String> {
    if state.running.swap(true, Ordering::SeqCst) {
        return Err("已有转换任务正在运行".into());
    }
    state.cancelled.store(false, Ordering::Relaxed);

    let result = run_batch_inner(&app, &state.cancelled, request).await;
    state.running.store(false, Ordering::SeqCst);
    result
}

async fn run_batch_inner(
    app: &AppHandle,
    cancelled: &AtomicBool,
    request: BatchRequest,
) -> Result<BatchSummary, String> {
    let output = PathBuf::from(&request.output_path);
    fs::create_dir_all(&output).map_err(|e| format!("无法创建输出文件夹：{e}"))?;
    let work_rows = request.rows;
    let total = work_rows.len();
    let mut success = 0;
    let mut failed = 0;
    let mut skipped = 0;
    let mut record = TaskRecord {
        version: 1,
        excel_path: request.excel_path,
        output_path: request.output_path,
        voice: request.voice.clone(),
        updated_at: Utc::now().to_rfc3339(),
        rows: request.snapshot_rows,
    };
    save_task(&record)?;

    for (index, row) in work_rows.into_iter().enumerate() {
        if cancelled.load(Ordering::Relaxed) {
            break;
        }
        let record_index = record
            .rows
            .iter()
            .position(|item| item.row_number == row.row_number && item.name == row.name)
            .ok_or_else(|| format!("任务记录中找不到 Excel 第 {} 行", row.row_number))?;
        let target = output.join(format!("{}.mp3", row.name));
        if target.exists() && !request.overwrite {
            skipped += 1;
            record.rows[record_index].status = "skipped".into();
            record.rows[record_index].last_error = None;
            record.rows[record_index].retryable = false;
        } else {
            record.rows[record_index].status = "running".into();
            record.rows[record_index].last_error = None;
            record.rows[record_index].retryable = false;
            let _ = app.emit(
                "tts-progress",
                ProgressEvent {
                    row_number: row.row_number,
                    name: row.name.clone(),
                    status: "running".into(),
                    error: None,
                    retryable: false,
                    processed: index,
                    total,
                },
            );
            let temp = output.join(format!(".{}.{}.tmp", row.name, row.row_number));
            match synthesize(&row.text, &request.voice).await {
                Ok(audio) => {
                    let backup = output.join(format!(".{}.{}.backup", row.name, row.row_number));
                    let write_result = fs::write(&temp, audio).and_then(|_| {
                        let had_target = target.exists();
                        if had_target {
                            let _ = fs::remove_file(&backup);
                            fs::rename(&target, &backup)?;
                        }
                        match fs::rename(&temp, &target) {
                            Ok(()) => {
                                let _ = fs::remove_file(&backup);
                                Ok(())
                            }
                            Err(error) => {
                                if had_target {
                                    let _ = fs::rename(&backup, &target);
                                }
                                Err(error)
                            }
                        }
                    });
                    if let Err(error) = write_result {
                        let _ = fs::remove_file(&temp);
                        failed += 1;
                        record.rows[record_index].status = "failed".into();
                        record.rows[record_index].last_error =
                            Some(format!("写入文件失败：{error}"));
                        record.rows[record_index].retryable = false;
                    } else {
                        success += 1;
                        record.rows[record_index].status = "success".into();
                        record.rows[record_index].last_error = None;
                        record.rows[record_index].retryable = false;
                    }
                }
                Err(error) => {
                    let _ = fs::remove_file(&temp);
                    failed += 1;
                    record.rows[record_index].status = "failed".into();
                    record.rows[record_index].last_error = Some(error);
                    record.rows[record_index].retryable = true;
                }
            }
        }

        record.updated_at = Utc::now().to_rfc3339();
        save_task(&record)?;
        let current = &record.rows[record_index];
        let _ = app.emit(
            "tts-progress",
            ProgressEvent {
                row_number: current.row_number,
                name: current.name.clone(),
                status: current.status.clone(),
                error: current.last_error.clone(),
                retryable: current.retryable,
                processed: index + 1,
                total,
            },
        );
    }

    Ok(BatchSummary {
        success,
        failed,
        skipped,
        cancelled: cancelled.load(Ordering::Relaxed),
    })
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .manage(AppState::default())
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .plugin(tauri_plugin_opener::init())
        .invoke_handler(tauri::generate_handler![
            load_excel,
            preview_speech,
            run_batch,
            cancel_batch,
            load_saved_task,
            open_output_folder
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

#[cfg(test)]
mod tests {
    use super::*;
    use calamine::{open_workbook, Reader, Xlsx};
    use std::io::Write;
    use zip::{write::SimpleFileOptions, ZipWriter};

    #[test]
    fn validates_windows_filenames() {
        assert!(filename_error("A-001").is_none());
        assert!(filename_error("产品 01").is_none());
        assert!(filename_error("bad:name").is_some());
        assert!(filename_error("CON").is_some());
        assert!(filename_error("name.").is_some());
    }

    #[test]
    fn restores_zero_padded_excel_names() {
        let path = std::env::temp_dir().join("excel_tts_zero_padding.xlsx");
        let file = File::create(&path).unwrap();
        let mut zip = ZipWriter::new(file);
        let options = SimpleFileOptions::default();
        let files = [
            (
                "[Content_Types].xml",
                r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#,
            ),
            (
                "_rels/.rels",
                r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            ),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#,
            ),
            (
                "xl/styles.xml",
                r#"<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="000"/></numFmts><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="164" applyNumberFormat="1"/></cellXfs></styleSheet>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>序号</t></is></c><c r="B1" t="inlineStr"><is><t>文字</t></is></c></row><row r="2"><c r="A2" s="1"><v>1</v></c><c r="B2" t="inlineStr"><is><t>测试文字</t></is></c></row></sheetData></worksheet>"#,
            ),
        ];
        for (name, content) in files {
            zip.start_file(name, options).unwrap();
            zip.write_all(content.as_bytes()).unwrap();
        }
        zip.finish().unwrap();

        let names = zero_padded_names(path.to_str().unwrap());
        assert_eq!(names.get(&2).map(String::as_str), Some("001"));
        let mut workbook: Xlsx<_> = open_workbook(&path).unwrap();
        assert_eq!(
            workbook.worksheet_range_at(0).unwrap().unwrap().get((1, 0)),
            Some(&Data::Float(1.0))
        );
        let all_rows = load_excel(path.to_string_lossy().into_owned(), None, false).unwrap();
        assert_eq!(all_rows.rows[0].row_number, 1);
        assert_eq!(all_rows.rows[0].name, "序号");
        let without_header = load_excel(path.to_string_lossy().into_owned(), None, true).unwrap();
        assert_eq!(without_header.rows.len(), 1);
        assert_eq!(without_header.rows[0].row_number, 2);
        assert_eq!(without_header.rows[0].name, "001");
        fs::remove_file(path).unwrap();
    }

    #[tokio::test]
    #[ignore = "requires Edge TTS network access"]
    async fn generates_real_edge_tts_audio() {
        let audio = synthesize(
            "你好，this is a Rust text to speech test.",
            &VoiceSetting {
                id: "zh-CN-YunxiNeural".into(),
                locale: "zh-CN".into(),
                rate: "+0%".into(),
            },
        )
        .await
        .unwrap();
        assert!(audio.len() > 1_000);
        assert!(audio.starts_with(b"ID3") || audio.starts_with(&[0xff, 0xf3]));
    }
}
