#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use regex::Regex;
use serde::Deserialize;
use serde::Serialize;
use std::fs;
use std::os::windows::process::CommandExt;
use std::path::{Path, PathBuf};
use std::process::{Command, Output};
use std::time::{SystemTime, UNIX_EPOCH};

const INTERNAL_DOCX_TO_MD_SCRIPT: &str = "convert_docx_to_md.py";
const INTERNAL_MD_TO_DOCX_SCRIPT: &str = "convert_md_to_docx.py";
const CREATE_NO_WINDOW: u32 = 0x08000000;

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ConvertRequest {
    input_path: String,
    output_path: String,
    math: String,
    extract_images: bool,
    split_sections: bool,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct OpenFolderRequest {
    path: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct StyleSpec {
    zh_font: String,
    en_font: String,
    font_size_pt: f64,
    line_spacing_mode: String,
    line_spacing_value: f64,
    align: String,
    #[serde(default)]
    advanced_override: Option<AdvancedSettings>,
    // Legacy v1/v2 fields kept for compatibility.
    #[serde(default)]
    bold: Option<bool>,
    #[serde(default)]
    italic: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct AdvancedSettings {
    #[serde(default = "default_spacing_setting_pt")]
    before: SpacingSetting,
    #[serde(default = "default_spacing_setting_pt")]
    after: SpacingSetting,
    // Legacy v1/v2/v3 compatibility.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    before_pt: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    after_pt: Option<f64>,
    first_line_indent_chars: f64,
    bold: bool,
    italic: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct SpacingSetting {
    mode: String,
    value: f64,
}

fn default_advanced_settings() -> AdvancedSettings {
    AdvancedSettings {
        before: default_spacing_setting_pt(),
        after: default_spacing_setting_pt(),
        before_pt: None,
        after_pt: None,
        first_line_indent_chars: 0.0,
        bold: false,
        italic: false,
    }
}

fn default_spacing_setting_pt() -> SpacingSetting {
    SpacingSetting {
        mode: "pt".to_string(),
        value: 0.0,
    }
}

fn default_heading3_style() -> StyleSpec {
    StyleSpec {
        zh_font: "宋体".to_string(),
        en_font: "Times New Roman".to_string(),
        font_size_pt: 12.0,
        line_spacing_mode: "multiple".to_string(),
        line_spacing_value: 1.5,
        align: "left".to_string(),
        advanced_override: Some(AdvancedSettings {
            before: default_spacing_setting_pt(),
            after: default_spacing_setting_pt(),
            before_pt: None,
            after_pt: None,
            first_line_indent_chars: 0.0,
            bold: true,
            italic: false,
        }),
        bold: None,
        italic: None,
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct MdToDocxStyleConfig {
    title: StyleSpec,
    abstract_zh: StyleSpec,
    abstract_en: StyleSpec,
    heading1: StyleSpec,
    heading2: StyleSpec,
    #[serde(default = "default_heading3_style")]
    heading3: StyleSpec,
    figure_caption: StyleSpec,
    table_caption: StyleSpec,
    body: StyleSpec,
    #[serde(default = "default_advanced_settings")]
    advanced_defaults: AdvancedSettings,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct StylePresetFile {
    version: u32,
    style_config: MdToDocxStyleConfig,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ConvertMdToDocxRequest {
    input_path: String,
    output_path: String,
    style_config: MdToDocxStyleConfig,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct SaveStylePresetRequest {
    path: String,
    style_config: MdToDocxStyleConfig,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct LoadStylePresetRequest {
    path: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct ConvertResponse {
    success: bool,
    exit_code: i32,
    stdout: String,
    stderr: String,
    output_path: String,
    markdown: String,
    split_output_dir: Option<String>,
    split_file_count: Option<usize>,
    assets_dir: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct ConvertMdToDocxResponse {
    success: bool,
    exit_code: i32,
    stdout: String,
    stderr: String,
    output_path: String,
}

fn migrate_advanced_settings_compat(settings: &mut AdvancedSettings) {
    if let Some(before_pt) = settings.before_pt {
        settings.before.mode = "pt".to_string();
        settings.before.value = before_pt;
    }
    if let Some(after_pt) = settings.after_pt {
        settings.after.mode = "pt".to_string();
        settings.after.value = after_pt;
    }
}

fn migrate_style_spec_compat(spec: &mut StyleSpec) {
    if let Some(advanced) = spec.advanced_override.as_mut() {
        migrate_advanced_settings_compat(advanced);
    }
}

fn migrate_style_config_compat(mut config: MdToDocxStyleConfig) -> MdToDocxStyleConfig {
    migrate_advanced_settings_compat(&mut config.advanced_defaults);
    migrate_style_spec_compat(&mut config.title);
    migrate_style_spec_compat(&mut config.abstract_zh);
    migrate_style_spec_compat(&mut config.abstract_en);
    migrate_style_spec_compat(&mut config.heading1);
    migrate_style_spec_compat(&mut config.heading2);
    migrate_style_spec_compat(&mut config.heading3);
    migrate_style_spec_compat(&mut config.figure_caption);
    migrate_style_spec_compat(&mut config.table_caption);
    migrate_style_spec_compat(&mut config.body);
    config
}

fn find_internal_converter_script(script_name: &str) -> Result<PathBuf, String> {
    let current_dir = std::env::current_dir().map_err(|e| format!("Cannot read current dir: {e}"))?;
    let exe_dir = std::env::current_exe()
        .ok()
        .and_then(|p| p.parent().map(|p| p.to_path_buf()));

    let mut candidates = vec![
        current_dir.join("converter").join(script_name),
        current_dir.join("..").join("converter").join(script_name),
    ];

    if let Some(exe_dir) = exe_dir {
        candidates.push(exe_dir.join("converter").join(script_name));
        candidates.push(
            exe_dir
                .join("..")
                .join("..")
                .join("..")
                .join("converter")
                .join(script_name),
        );
    }

    for candidate in candidates {
        if candidate.exists() {
            return Ok(candidate);
        }
    }

    Err(format!(
        "Cannot locate internal converter script (converter/{script_name})"
    ))
}

fn find_docx_to_md_script() -> Result<PathBuf, String> {
    find_internal_converter_script(INTERNAL_DOCX_TO_MD_SCRIPT)
}

fn ensure_parent(path: &str) -> Result<(), String> {
    let path = Path::new(path);
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)
            .map_err(|e| format!("Failed to create dir {}: {e}", parent.display()))?;
    }
    Ok(())
}

fn non_conflicting_output_path(path: &Path) -> PathBuf {
    if !path.exists() {
        return path.to_path_buf();
    }

    let parent = path.parent().unwrap_or_else(|| Path::new("."));
    let stem = path
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("output")
        .to_string();
    let ext = path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| format!(".{e}"))
        .unwrap_or_default();

    for i in 1..10_000 {
        let candidate = parent.join(format!("{stem} ({i}){ext}"));
        if !candidate.exists() {
            return candidate;
        }
    }

    parent.join(format!("{stem}_copy{ext}"))
}

fn build_python_candidates() -> Vec<(String, Vec<String>)> {
    let mut candidates: Vec<(String, Vec<String>)> = Vec::new();

    if let Ok(executable) = std::env::var("PYTHON_EXECUTABLE") {
        if !executable.trim().is_empty() {
            candidates.push((executable, vec![]));
        }
    }

    candidates.push(("python".to_string(), vec![]));
    candidates.push(("python3".to_string(), vec![]));
    candidates.push(("py".to_string(), vec!["-3".to_string()]));
    candidates.push(("py".to_string(), vec![]));

    candidates
}

fn run_python_script_with_fallback(script_path: &Path, script_args: &[String]) -> Result<Output, String> {
    let mut attempted: Vec<String> = Vec::new();
    let mut errors: Vec<String> = Vec::new();

    for (program, prefix_args) in build_python_candidates() {
        let mut command = Command::new(&program);
        command.creation_flags(CREATE_NO_WINDOW);
        command.env("PYTHONIOENCODING", "utf-8");
        command.env("PYTHONUTF8", "1");
        for arg in &prefix_args {
            command.arg(arg);
        }

        command.arg(script_path);
        for arg in script_args {
            command.arg(arg);
        }

        let mut cmd_display = program.clone();
        for arg in &prefix_args {
            cmd_display.push(' ');
            cmd_display.push_str(arg);
        }
        attempted.push(cmd_display);

        match command.output() {
            Ok(output) => return Ok(output),
            Err(err) => errors.push(format!("{program}: {err}")),
        }
    }

    Err(format!(
        "Python runtime not found. Tried: {}. Install Python 3 or set PYTHON_EXECUTABLE. Last errors: {}",
        attempted.join(", "),
        errors.join(" | ")
    ))
}

fn build_docx_to_md_args(
    input_path: &Path,
    request: &ConvertRequest,
    image_dir: Option<&str>,
    assets_dir: &Path,
) -> Vec<String> {
    let mut args: Vec<String> = Vec::new();
    args.push("--input".to_string());
    args.push(input_path.to_string_lossy().to_string());
    args.push("--output".to_string());
    args.push(request.output_path.clone());
    args.push("--math".to_string());
    args.push(request.math.clone());
    args.push("--assets-dir".to_string());
    args.push(assets_dir.to_string_lossy().to_string());

    if let Some(dir) = image_dir {
        args.push("--extract-images".to_string());
        args.push("--image-dir".to_string());
        args.push(dir.to_string());
    }

    args
}

fn run_converter_with_python_fallback(
    script_path: &Path,
    input_path: &Path,
    request: &ConvertRequest,
    image_dir: Option<&str>,
    assets_dir: &Path,
) -> Result<Output, String> {
    let args = build_docx_to_md_args(input_path, request, image_dir, assets_dir);
    run_python_script_with_fallback(script_path, &args)
}

fn extension_lower(path: &Path) -> String {
    path.extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase()
}

fn build_temp_path(prefix: &str, suffix: &str) -> Result<PathBuf, String> {
    let millis = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map_err(|e| format!("System time error: {e}"))?
        .as_millis();
    Ok(std::env::temp_dir().join(format!(
        "{}_{}_{}.{}",
        prefix,
        std::process::id(),
        millis,
        suffix
    )))
}

fn build_temp_docx_path(input_path: &Path) -> Result<PathBuf, String> {
    let stem = input_path
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("input")
        .replace(['\\', '/', ':', '*', '?', '"', '<', '>', '|'], "_");

    let millis = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map_err(|e| format!("System time error: {e}"))?
        .as_millis();

    let filename = format!("docxmd_{}_{}_{}.docx", stem, std::process::id(), millis);
    Ok(std::env::temp_dir().join(filename))
}

fn build_assets_dir(output_md_path: &Path) -> PathBuf {
    let parent = output_md_path.parent().unwrap_or_else(|| Path::new("."));
    let stem = output_md_path
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("output")
        .replace(['\\', '/', ':', '*', '?', '"', '<', '>', '|'], "_");
    parent.join(format!("{stem}_assets"))
}

fn build_image_dir(assets_dir: &Path) -> PathBuf {
    assets_dir.join("images")
}

fn ps_single_quote(value: &str) -> String {
    value.replace('\'', "''")
}

fn convert_doc_to_docx_with_word(input_path: &Path, output_path: &Path) -> Result<(), String> {
    let in_path = ps_single_quote(&input_path.to_string_lossy());
    let out_path = ps_single_quote(&output_path.to_string_lossy());

    let script = format!(
        r#"$ErrorActionPreference = 'Stop'
$in = '{in_path}'
$out = '{out_path}'
$word = $null
$doc = $null
try {{
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $doc = $word.Documents.Open($in)
  $wdFormatDocumentDefault = 16
  $doc.SaveAs([ref]$out, [ref]$wdFormatDocumentDefault)
}} catch {{
  Write-Error $_
  exit 1
}} finally {{
  if ($doc -ne $null) {{ try {{ $doc.Close() }} catch {{}} }}
  if ($word -ne $null) {{ try {{ $word.Quit() }} catch {{}} }}
}}"#
    );

    let output = Command::new("powershell")
        .creation_flags(CREATE_NO_WINDOW)
        .arg("-NoProfile")
        .arg("-NonInteractive")
        .arg("-ExecutionPolicy")
        .arg("Bypass")
        .arg("-Command")
        .arg(script)
        .output()
        .map_err(|e| format!("Word automation command failed to start: {e}"))?;

    if output.status.success() && output_path.exists() {
        return Ok(());
    }

    let stderr = String::from_utf8_lossy(&output.stderr).trim().to_string();
    let stdout = String::from_utf8_lossy(&output.stdout).trim().to_string();
    Err(format!("Word COM conversion failed. stdout: {stdout}; stderr: {stderr}"))
}

fn convert_doc_to_docx_with_soffice(input_path: &Path, output_path: &Path) -> Result<(), String> {
    let out_dir = output_path
        .parent()
        .ok_or_else(|| "Cannot resolve temp output directory".to_string())?;

    let stem = input_path
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("converted");
    let generated = out_dir.join(format!("{stem}.docx"));

    let candidates = vec![
        "soffice".to_string(),
        r"C:\Program Files\LibreOffice\program\soffice.exe".to_string(),
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe".to_string(),
    ];

    let mut errors: Vec<String> = Vec::new();

    for program in candidates {
        let result = Command::new(&program)
            .creation_flags(CREATE_NO_WINDOW)
            .arg("--headless")
            .arg("--convert-to")
            .arg("docx")
            .arg("--outdir")
            .arg(out_dir)
            .arg(input_path)
            .output();

        match result {
            Ok(output) => {
                if output.status.success() && generated.exists() {
                    if generated != output_path {
                        fs::copy(&generated, output_path).map_err(|e| {
                            format!("LibreOffice created DOCX but copy to temp path failed: {e}")
                        })?;
                        let _ = fs::remove_file(&generated);
                    }
                    if output_path.exists() {
                        return Ok(());
                    }
                }

                let stderr = String::from_utf8_lossy(&output.stderr).trim().to_string();
                let stdout = String::from_utf8_lossy(&output.stdout).trim().to_string();
                errors.push(format!("{program}: stdout={stdout}; stderr={stderr}"));
            }
            Err(err) => {
                errors.push(format!("{program}: {err}"));
            }
        }
    }

    Err(format!(
        "LibreOffice conversion failed. Attempts: {}",
        errors.join(" | ")
    ))
}

fn prepare_input_docx(input_path: &Path) -> Result<(PathBuf, Option<PathBuf>), String> {
    if !input_path.exists() {
        return Err(format!("Input file not found: {}", input_path.display()));
    }

    match extension_lower(input_path).as_str() {
        "docx" => Ok((input_path.to_path_buf(), None)),
        "doc" => {
            let temp_docx = build_temp_docx_path(input_path)?;

            match convert_doc_to_docx_with_word(input_path, &temp_docx) {
                Ok(()) => Ok((temp_docx.clone(), Some(temp_docx))),
                Err(word_err) => match convert_doc_to_docx_with_soffice(input_path, &temp_docx) {
                    Ok(()) => Ok((temp_docx.clone(), Some(temp_docx))),
                    Err(soffice_err) => Err(format!(
                        "Cannot convert .doc to .docx automatically. Please install Microsoft Word or LibreOffice, then retry. Word error: {word_err}. LibreOffice error: {soffice_err}"
                    )),
                },
            }
        }
        _ => Err("Input file must be .doc or .docx".to_string()),
    }
}

fn normalize_markdown_output(raw: &str) -> String {
    let image_word_re = Regex::new(r"(?i)^\s*!image\s*$").unwrap();
    let def_number_re = Regex::new(r"瀹氫箟\s+(\d+)").unwrap();
    let number_paren_re = Regex::new(r"(\d+)\s*([锛塡)])").unwrap();
    let multi_space_re = Regex::new(r"[ \t]{2,}").unwrap();
    let cjk_gap_re = Regex::new(r"([\u3400-\u9fff])\s+([\u3400-\u9fff])").unwrap();

    let normalized_source = raw.replace("\r\n", "\n").replace('\r', "\n");
    let mut cleaned_lines: Vec<String> = Vec::new();

    for source_line in normalized_source.lines() {
        let mut line = source_line.trim_end().to_string();

        if image_word_re.is_match(line.trim()) {
            continue;
        }

        loop {
            let next = cjk_gap_re.replace_all(&line, "$1$2").into_owned();
            if next == line {
                break;
            }
            line = next;
        }

        line = def_number_re.replace_all(&line, "瀹氫箟$1").into_owned();
        line = number_paren_re.replace_all(&line, "$1$2").into_owned();
        line = multi_space_re.replace_all(&line, " ").into_owned();
        line = line.trim().to_string();

        if line.is_empty() {
            cleaned_lines.push(String::new());
            continue;
        }

        cleaned_lines.push(line);
    }

    let mut compact: Vec<String> = Vec::new();
    let mut previous_blank = true;
    for line in cleaned_lines {
        if line.is_empty() {
            if !previous_blank {
                compact.push(String::new());
            }
            previous_blank = true;
        } else {
            compact.push(line);
            previous_blank = false;
        }
    }

    while compact.last().is_some_and(|line| line.is_empty()) {
        compact.pop();
    }

    if compact.is_empty() {
        String::new()
    } else {
        format!("{}\n", compact.join("\n"))
    }
}

fn trim_blank_edges(lines: &[String]) -> Vec<String> {
    if lines.is_empty() {
        return Vec::new();
    }
    let mut start = 0usize;
    let mut end = lines.len();

    while start < end && lines[start].trim().is_empty() {
        start += 1;
    }
    while end > start && lines[end - 1].trim().is_empty() {
        end -= 1;
    }
    lines[start..end].to_vec()
}

fn sanitize_filename(input: &str) -> String {
    let mut value = input
        .replace(['\\', '/', ':', '*', '?', '"', '<', '>', '|'], "_")
        .replace('\t', " ");
    value = Regex::new(r"\s+")
        .unwrap()
        .replace_all(&value, " ")
        .to_string()
        .trim()
        .to_string();
    if value.is_empty() {
        "section".to_string()
    } else {
        value.chars().take(60).collect()
    }
}

fn chapter_title_from_line(line: &str) -> Option<String> {
    let trimmed = line.trim();
    if !trimmed.starts_with("# ") || trimmed.starts_with("##") {
        return None;
    }
    let title = trimmed.trim_start_matches("#").trim();
    if title.is_empty() {
        None
    } else {
        Some(title.to_string())
    }
}

fn rewrite_split_image_paths(lines: &[String]) -> Vec<String> {
    let image_re = Regex::new(r"!\[([^\]]*)\]\(([^)\n]+)\)").unwrap();
    lines.iter()
        .map(|line| {
            image_re
                .replace_all(line, |caps: &regex::Captures| {
                    let alt = caps.get(1).map(|m| m.as_str()).unwrap_or_default();
                    let target = caps.get(2).map(|m| m.as_str().trim()).unwrap_or_default();

                    if target.is_empty()
                        || target.starts_with("http://")
                        || target.starts_with("https://")
                        || target.starts_with("data:")
                        || target.starts_with("../")
                        || Path::new(target).is_absolute()
                    {
                        return caps.get(0).map(|m| m.as_str()).unwrap_or_default().to_string();
                    }

                    let rewritten = if let Some(rest) = target.strip_prefix("./") {
                        format!("../{rest}")
                    } else {
                        format!("../{target}")
                    };
                    format!("![{alt}]({})", rewritten.replace('\\', "/"))
                })
                .into_owned()
        })
        .collect()
}

fn split_markdown_sections(markdown: &str, output_path: &Path) -> Result<Option<(PathBuf, usize)>, String> {
    let lines: Vec<String> = markdown
        .replace("\r\n", "\n")
        .replace('\r', "\n")
        .split('\n')
        .map(|line| line.to_string())
        .collect();

    let mut chapter_starts: Vec<(usize, String)> = Vec::new();
    for (index, line) in lines.iter().enumerate() {
        if let Some(title) = chapter_title_from_line(line) {
            chapter_starts.push((index, title));
        }
    }

    if chapter_starts.is_empty() {
        return Ok(None);
    }

    let mut section_dir = output_path
        .parent()
        .unwrap_or_else(|| Path::new("."))
        .to_path_buf();
    let stem = output_path
        .file_stem()
        .and_then(|value| value.to_str())
        .unwrap_or("output");
    section_dir.push(format!("{stem}_sections"));
    fs::create_dir_all(&section_dir)
        .map_err(|e| format!("Failed to create section output dir {}: {e}", section_dir.display()))?;

    if let Ok(existing) = fs::read_dir(&section_dir) {
        for item in existing.flatten() {
            let path = item.path();
            if path
                .extension()
                .and_then(|v| v.to_str())
                .unwrap_or_default()
                .eq_ignore_ascii_case("md")
            {
                let _ = fs::remove_file(path);
            }
        }
    }

    let first_body_index = if chapter_starts.len() > 1 {
        chapter_starts[1].0
    } else {
        lines.len()
    };
    let first_part = rewrite_split_image_paths(&trim_blank_edges(&lines[..first_body_index]));

    let mut written = 0usize;
    if !first_part.is_empty() {
        let first_file = section_dir.join("00_前置部分.md");
        fs::write(&first_file, format!("{}\n", first_part.join("\n")))
            .map_err(|e| format!("Failed writing {}: {e}", first_file.display()))?;
        written += 1;
    }

    let mut chapter_index = 1usize;
    for (i, (start, title)) in chapter_starts.iter().enumerate().skip(1) {
        let end = if i + 1 < chapter_starts.len() {
            chapter_starts[i + 1].0
        } else {
            lines.len()
        };
        let block = rewrite_split_image_paths(&trim_blank_edges(&lines[*start..end]));
        if block.is_empty() {
            continue;
        }

        let file_name = format!("{chapter_index:02}_{}.md", sanitize_filename(title));
        let chapter_file = section_dir.join(file_name);
        fs::write(&chapter_file, format!("{}\n", block.join("\n")))
            .map_err(|e| format!("Failed writing {}: {e}", chapter_file.display()))?;
        chapter_index += 1;
        written += 1;
    }

    if written == 0 {
        Ok(None)
    } else {
        Ok(Some((section_dir, written)))
    }
}

fn parse_report_field_string(value: &serde_json::Value, field: &str) -> Option<String> {
    value
        .get(field)
        .and_then(|item| item.as_str())
        .map(|s| s.to_string())
}

#[tauri::command]
fn convert_docx(mut request: ConvertRequest) -> Result<ConvertResponse, String> {
    ensure_parent(&request.output_path)?;
    let resolved_output = non_conflicting_output_path(Path::new(&request.output_path));
    request.output_path = resolved_output.to_string_lossy().to_string();

    let input_path = PathBuf::from(&request.input_path);
    let output_path_buf = PathBuf::from(&request.output_path);
    let assets_dir = build_assets_dir(&output_path_buf);
    let image_dir = if request.extract_images {
        let image_dir = build_image_dir(&assets_dir);
        fs::create_dir_all(&image_dir)
            .map_err(|e| format!("Failed to create image dir {}: {e}", image_dir.display()))?;
        Some(image_dir.to_string_lossy().to_string())
    } else {
        None
    };

    let (prepared_input, temp_docx) = prepare_input_docx(&input_path)?;

    let script = find_docx_to_md_script()?;
    let run_result = run_converter_with_python_fallback(
        &script,
        &prepared_input,
        &request,
        image_dir.as_deref(),
        &assets_dir,
    );

    if let Some(tmp) = temp_docx {
        let _ = fs::remove_file(tmp);
    }

    let output = run_result?;

    let exit_code = output.status.code().unwrap_or(2);
    let stdout = String::from_utf8_lossy(&output.stdout).trim().to_string();
    let stderr = String::from_utf8_lossy(&output.stderr).trim().to_string();
    let raw_markdown = fs::read_to_string(&request.output_path).unwrap_or_default();
    let markdown = normalize_markdown_output(&raw_markdown);
    if markdown != raw_markdown {
        fs::write(&request.output_path, &markdown)
            .map_err(|e| format!("Failed to write normalized markdown {}: {e}", request.output_path))?;
    }
    let report_value = serde_json::from_str::<serde_json::Value>(&stdout).ok();
    let assets_dir_reported = report_value
        .as_ref()
        .and_then(|v| parse_report_field_string(v, "assetsDir"))
        .or_else(|| {
            if request.extract_images {
                Some(assets_dir.to_string_lossy().to_string())
            } else {
                None
            }
        });

    let split_result = if request.split_sections {
        split_markdown_sections(&markdown, Path::new(&request.output_path))?
    } else {
        None
    };
    let (split_output_dir, split_file_count) = match split_result {
        Some((dir, count)) => (Some(dir.to_string_lossy().to_string()), Some(count)),
        None => (None, None),
    };

    Ok(ConvertResponse {
        success: exit_code == 0,
        exit_code,
        stdout,
        stderr,
        output_path: request.output_path.clone(),
        markdown,
        split_output_dir,
        split_file_count,
        assets_dir: assets_dir_reported,
    })
}

#[tauri::command]
fn convert_md_to_docx(mut request: ConvertMdToDocxRequest) -> Result<ConvertMdToDocxResponse, String> {
    ensure_parent(&request.output_path)?;
    let resolved_output = non_conflicting_output_path(Path::new(&request.output_path));
    request.output_path = resolved_output.to_string_lossy().to_string();

    let input_path = PathBuf::from(&request.input_path);
    if !input_path.exists() {
        return Err(format!("Input markdown not found: {}", input_path.display()));
    }
    if !extension_lower(&input_path).eq("md") {
        return Err("Input file must be .md".to_string());
    }

    let script = find_internal_converter_script(INTERNAL_MD_TO_DOCX_SCRIPT)?;
    let temp_style_path = build_temp_path("docxmd_style", "json")?;
    let style_json = serde_json::to_string_pretty(&request.style_config)
        .map_err(|e| format!("Failed to serialize style config: {e}"))?;

    fs::write(&temp_style_path, style_json)
        .map_err(|e| format!("Failed to write temp style config: {e}"))?;

    let args = vec![
        "--input".to_string(),
        request.input_path.clone(),
        "--output".to_string(),
        request.output_path.clone(),
        "--style".to_string(),
        temp_style_path.to_string_lossy().to_string(),
    ];

    let run_result = run_python_script_with_fallback(&script, &args);
    let _ = fs::remove_file(&temp_style_path);

    let output = run_result?;

    let exit_code = output.status.code().unwrap_or(2);
    let stdout = String::from_utf8_lossy(&output.stdout).trim().to_string();
    let stderr = String::from_utf8_lossy(&output.stderr).trim().to_string();

    Ok(ConvertMdToDocxResponse {
        success: exit_code == 0,
        exit_code,
        stdout,
        stderr,
        output_path: request.output_path.clone(),
    })
}

#[tauri::command]
fn save_style_preset(request: SaveStylePresetRequest) -> Result<(), String> {
    ensure_parent(&request.path)?;
    let payload = StylePresetFile {
        version: 5,
        style_config: request.style_config,
    };
    let json = serde_json::to_string_pretty(&payload)
        .map_err(|e| format!("Failed to serialize style preset: {e}"))?;
    fs::write(&request.path, json)
        .map_err(|e| format!("Failed to write style preset {}: {e}", request.path))?;
    Ok(())
}

#[tauri::command]
fn load_style_preset(request: LoadStylePresetRequest) -> Result<MdToDocxStyleConfig, String> {
    let text = fs::read_to_string(&request.path)
        .map_err(|e| format!("Failed to read style preset {}: {e}", request.path))?;
    let value: serde_json::Value =
        serde_json::from_str(&text).map_err(|e| format!("Invalid style preset JSON: {e}"))?;

    if value.get("styleConfig").is_some() {
        let preset: StylePresetFile = serde_json::from_value(value)
            .map_err(|e| format!("Invalid style preset JSON: {e}"))?;
        return Ok(migrate_style_config_compat(preset.style_config));
    }

    serde_json::from_str::<MdToDocxStyleConfig>(&text)
        .map(migrate_style_config_compat)
        .map_err(|e| format!("Invalid style preset JSON: {e}"))
}

#[tauri::command]
fn open_folder_for_path(request: OpenFolderRequest) -> Result<(), String> {
    if request.path.trim().is_empty() {
        return Err("Path is empty".to_string());
    }
    let target = PathBuf::from(&request.path);

    let folder = if target.exists() {
        if target.is_dir() {
            target
        } else {
            target
                .parent()
                .map(Path::to_path_buf)
                .ok_or_else(|| "Cannot resolve parent folder".to_string())?
        }
    } else {
        target
            .parent()
            .map(Path::to_path_buf)
            .ok_or_else(|| "Cannot resolve parent folder".to_string())?
    };

    if !folder.exists() {
        return Err(format!("Folder not found: {}", folder.display()));
    }

    Command::new("explorer")
        .arg(folder)
        .creation_flags(CREATE_NO_WINDOW)
        .spawn()
        .map_err(|e| format!("Failed to open folder: {e}"))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::{chapter_title_from_line, split_markdown_sections};
    use std::fs;
    use std::time::{SystemTime, UNIX_EPOCH};

    #[test]
    fn chapter_title_should_parse_top_level_headings() {
        assert_eq!(chapter_title_from_line("# 论文题目"), Some("论文题目".to_string()));
        assert_eq!(chapter_title_from_line("# 1 引言"), Some("1 引言".to_string()));
        assert_eq!(chapter_title_from_line("# 4 结果分析"), Some("4 结果分析".to_string()));
    }

    #[test]
    fn chapter_title_should_skip_non_top_level_lines() {
        assert_eq!(chapter_title_from_line("1 引言"), None);
        assert_eq!(chapter_title_from_line("2.2 集成框架"), None);
        assert_eq!(chapter_title_from_line("## 3.1.4 子章节"), None);
        assert_eq!(chapter_title_from_line("### 2024年数据增长明显"), None);
        assert_eq!(chapter_title_from_line("- # 1 引言"), None);
    }

    #[test]
    fn split_sections_should_create_output_without_crash() {
        let millis = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_millis();
        let base = std::env::temp_dir().join(format!(
            "docxmd_split_test_{}_{}",
            std::process::id(),
            millis
        ));
        fs::create_dir_all(&base).expect("create temp test dir");

        let output_path = base.join("paper.md");
        let markdown = "# 论文题目\n\n## 摘要\n这是摘要。\n\n![image](paper_assets/images/fig1.png)\n\n# 1 引言\n引言内容。\n\n# 2 方法\n方法内容。\n";
        let result = split_markdown_sections(markdown, &output_path).expect("split markdown");
        let (section_dir, count) = result.expect("split result");

        assert!(section_dir.exists());
        assert!(count >= 2);
        let preface = fs::read_to_string(section_dir.join("00_前置部分.md")).expect("read preface");
        assert!(preface.contains("# 论文题目"));
        assert!(preface.contains("../paper_assets/images/fig1.png"));

        let _ = fs::remove_dir_all(&base);
    }
}

fn main() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![
            convert_docx,
            convert_md_to_docx,
            save_style_preset,
            load_style_preset,
            open_folder_for_path
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

