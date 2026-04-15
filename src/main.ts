import { invoke } from "@tauri-apps/api/core";
import { open, save } from "@tauri-apps/plugin-dialog";
import "./style.css";

type Direction = "docxToMd" | "mdToDocx";
type Locale = "zh" | "en";
type LineSpacingMode = "multiple" | "fixed";
type AlignMode = "left" | "center" | "right" | "justify";
type SpacingMode = "pt" | "lines";

type ConversionResponse = {
  success: boolean;
  exitCode: number;
  stdout: string;
  stderr: string;
  outputPath: string;
  markdown: string;
  splitOutputDir?: string | null;
  splitFileCount?: number | null;
  assetsDir?: string | null;
};

type ConvertRequest = {
  inputPath: string;
  outputPath: string;
  math: "latex";
  extractImages: boolean;
  splitSections: boolean;
};

type StyleAdvancedSettings = {
  before: {
    mode: SpacingMode;
    value: number;
  };
  after: {
    mode: SpacingMode;
    value: number;
  };
  firstLineIndentChars: number;
  bold: boolean;
  italic: boolean;
};

type MdToDocxStyleBlock = {
  zhFont: string;
  enFont: string;
  fontSizePt: number;
  lineSpacingMode: LineSpacingMode;
  lineSpacingValue: number;
  align: AlignMode;
  advancedOverride?: StyleAdvancedSettings | null;
  // Legacy fields kept for preset compatibility (v1/v2).
  bold?: boolean;
  italic?: boolean;
};

type TablePreset = "threeLine" | "tableGrid" | "table";

type MdToDocxTableSettings = {
  tablePreset: TablePreset;
  headerBold: boolean;
  textStyle: MdToDocxStyleBlock;
  // Compatibility field for old presets; not exposed in UI.
  applyTextStyle: boolean;
};

type MdToDocxStyleConfig = {
  title: MdToDocxStyleBlock;
  abstractZh: MdToDocxStyleBlock;
  abstractEn: MdToDocxStyleBlock;
  heading1: MdToDocxStyleBlock;
  heading2: MdToDocxStyleBlock;
  heading3: MdToDocxStyleBlock;
  figureCaption: MdToDocxStyleBlock;
  tableCaption: MdToDocxStyleBlock;
  body: MdToDocxStyleBlock;
  tableSettings: MdToDocxTableSettings;
  advancedDefaults: StyleAdvancedSettings;
};

type ConvertMdToDocxRequest = {
  inputPath: string;
  outputPath: string;
  styleConfig: MdToDocxStyleConfig;
};

type ConvertMdToDocxResponse = {
  success: boolean;
  exitCode: number;
  stdout: string;
  stderr: string;
  outputPath: string;
};

const ZH_FONT_OPTIONS = ["宋体", "黑体", "仿宋", "楷体", "微软雅黑"];
const EN_FONT_OPTIONS = ["Times New Roman", "Arial", "Calibri", "Cambria", "Georgia"];
let currentLocale: Locale = "zh";

const tr = (zh: string, en: string) => (currentLocale === "zh" ? zh : en);

const DEFAULT_ADVANCED_SETTINGS: StyleAdvancedSettings = {
  before: {
    mode: "pt",
    value: 0,
  },
  after: {
    mode: "pt",
    value: 0,
  },
  firstLineIndentChars: 0,
  bold: false,
  italic: false,
};

const DEFAULT_TABLE_TEXT_STYLE: MdToDocxStyleBlock = {
  zhFont: "宋体",
  enFont: "Times New Roman",
  fontSizePt: 12,
  lineSpacingMode: "multiple",
  lineSpacingValue: 1.0,
  align: "center",
  advancedOverride: { ...DEFAULT_ADVANCED_SETTINGS },
};

const DEFAULT_TABLE_SETTINGS: MdToDocxTableSettings = {
  tablePreset: "threeLine",
  headerBold: false,
  textStyle: { ...DEFAULT_TABLE_TEXT_STYLE },
  applyTextStyle: true,
};

const COMPAT_TABLE_SETTINGS: MdToDocxTableSettings = {
  tablePreset: "tableGrid",
  headerBold: false,
  textStyle: { ...DEFAULT_TABLE_TEXT_STYLE },
  applyTextStyle: false,
};

const DEFAULT_STYLE_CONFIG: MdToDocxStyleConfig = {
  title: {
    zhFont: "宋体",
    enFont: "Times New Roman",
    fontSizePt: 16,
    lineSpacingMode: "multiple",
    lineSpacingValue: 1.5,
    align: "center",
    advancedOverride: { ...DEFAULT_ADVANCED_SETTINGS, bold: true },
  },
  abstractZh: {
    zhFont: "宋体",
    enFont: "Times New Roman",
    fontSizePt: 12,
    lineSpacingMode: "multiple",
    lineSpacingValue: 1.5,
    align: "justify",
    advancedOverride: null,
  },
  abstractEn: {
    zhFont: "宋体",
    enFont: "Times New Roman",
    fontSizePt: 12,
    lineSpacingMode: "multiple",
    lineSpacingValue: 1.5,
    align: "justify",
    advancedOverride: null,
  },
  heading1: {
    zhFont: "宋体",
    enFont: "Times New Roman",
    fontSizePt: 14,
    lineSpacingMode: "multiple",
    lineSpacingValue: 1.5,
    align: "left",
    advancedOverride: { ...DEFAULT_ADVANCED_SETTINGS, bold: true },
  },
  heading2: {
    zhFont: "宋体",
    enFont: "Times New Roman",
    fontSizePt: 12,
    lineSpacingMode: "multiple",
    lineSpacingValue: 1.5,
    align: "left",
    advancedOverride: { ...DEFAULT_ADVANCED_SETTINGS, bold: true },
  },
  heading3: {
    zhFont: "宋体",
    enFont: "Times New Roman",
    fontSizePt: 12,
    lineSpacingMode: "multiple",
    lineSpacingValue: 1.5,
    align: "left",
    advancedOverride: { ...DEFAULT_ADVANCED_SETTINGS, bold: true },
  },
  figureCaption: {
    zhFont: "宋体",
    enFont: "Times New Roman",
    fontSizePt: 11,
    lineSpacingMode: "multiple",
    lineSpacingValue: 1.2,
    align: "center",
    advancedOverride: null,
  },
  tableCaption: {
    zhFont: "宋体",
    enFont: "Times New Roman",
    fontSizePt: 11,
    lineSpacingMode: "multiple",
    lineSpacingValue: 1.2,
    align: "center",
    advancedOverride: null,
  },
  body: {
    zhFont: "宋体",
    enFont: "Times New Roman",
    fontSizePt: 12,
    lineSpacingMode: "multiple",
    lineSpacingValue: 1.5,
    align: "justify",
    advancedOverride: null,
  },
  tableSettings: { ...DEFAULT_TABLE_SETTINGS },
  advancedDefaults: { ...DEFAULT_ADVANCED_SETTINGS },
};

type StyleSectionKey = Exclude<keyof MdToDocxStyleConfig, "advancedDefaults" | "tableSettings">;

const STYLE_SECTIONS: StyleSectionKey[] = [
  "title",
  "abstractZh",
  "abstractEn",
  "heading1",
  "heading2",
  "heading3",
  "figureCaption",
  "tableCaption",
  "body",
];
const COLLAPSIBLE_STYLE_KEYS: StyleSectionKey[] = [
  "title",
  "abstractZh",
  "abstractEn",
  "heading1",
  "heading2",
  "heading3",
  "figureCaption",
  "tableCaption",
  "body",
];

const getStyleSectionLabel = (key: StyleSectionKey) => {
  switch (key) {
    case "title":
      return tr("论文题目", "Paper Title");
    case "abstractZh":
      return tr("中文摘要", "Chinese Abstract");
    case "abstractEn":
      return tr("英文摘要", "English Abstract");
    case "heading1":
      return tr("一级标题", "Heading 1");
    case "heading2":
      return tr("二级标题", "Heading 2");
    case "heading3":
      return tr("三级标题", "Heading 3");
    case "figureCaption":
      return tr("图标题", "Figure Caption");
    case "tableCaption":
      return tr("表标题", "Table Caption");
    case "body":
      return tr("正文", "Body Text");
    default:
      return key;
  }
};

const getTablePresetLabel = (preset: TablePreset) => {
  switch (preset) {
    case "threeLine":
      return tr("三线表", "Three-Line");
    case "table":
      return tr("基础表", "Basic Table");
    case "tableGrid":
    default:
      return tr("网格表", "Grid Table");
  }
};

const app = document.querySelector<HTMLDivElement>("#app");
if (!app) {
  throw new Error("#app not found");
}

const styleInputId = (key: StyleSectionKey, field: string) => `style-${key}-${field}`;
const isCollapsibleSection = (key: StyleSectionKey) => COLLAPSIBLE_STYLE_KEYS.includes(key);

const renderSelectOptions = (options: string[], selected: string) => {
  const allOptions = options.includes(selected) ? options : [selected, ...options];
  return allOptions
    .map((option) => `<option value="${option}" ${option === selected ? "selected" : ""}>${option}</option>`)
    .join("");
};

const renderStyleCard = (
  key: StyleSectionKey,
  label: string,
  block: MdToDocxStyleBlock,
) => {
  const advancedValues = block.advancedOverride ?? DEFAULT_ADVANCED_SETTINGS;

  const fields = `
    <article class="style-card" data-style-key="${key}">
    <div class="style-grid-fields">
      <label>${tr("中文字体", "Chinese Font")}
        <select id="${styleInputId(key, "zhFont")}">
          ${renderSelectOptions(ZH_FONT_OPTIONS, block.zhFont)}
        </select>
      </label>
      <label>${tr("英文字体", "English Font")}
        <select id="${styleInputId(key, "enFont")}">
          ${renderSelectOptions(EN_FONT_OPTIONS, block.enFont)}
        </select>
      </label>
      <label>${tr("字号(pt)", "Font Size (pt)")}<input id="${styleInputId(key, "fontSizePt")}" type="number" min="1" step="0.5" value="${block.fontSizePt}" /></label>
      <label>${tr("行距模式", "Line Spacing Mode")}
        <select id="${styleInputId(key, "lineSpacingMode")}">
          <option value="multiple" ${block.lineSpacingMode === "multiple" ? "selected" : ""}>${tr("倍数", "Multiple")}</option>
          <option value="fixed" ${block.lineSpacingMode === "fixed" ? "selected" : ""}>${tr("固定值(pt)", "Fixed (pt)")}</option>
        </select>
      </label>
      <label>${tr("行距值", "Line Spacing Value")}<input id="${styleInputId(key, "lineSpacingValue")}" type="number" min="0.1" step="0.1" value="${block.lineSpacingValue}" /></label>
      <label>${tr("对齐", "Alignment")}
        <select id="${styleInputId(key, "align")}">
          <option value="left" ${block.align === "left" ? "selected" : ""}>${tr("左对齐", "Left")}</option>
          <option value="center" ${block.align === "center" ? "selected" : ""}>${tr("居中", "Center")}</option>
          <option value="right" ${block.align === "right" ? "selected" : ""}>${tr("右对齐", "Right")}</option>
          <option value="justify" ${block.align === "justify" ? "selected" : ""}>${tr("两端对齐", "Justify")}</option>
        </select>
      </label>
      <label>${tr("段前模式", "Before Paragraph Mode")}
        <select id="${styleInputId(key, "beforeMode")}">
          <option value="pt" ${advancedValues.before.mode === "pt" ? "selected" : ""}>${tr("固定值(pt)", "Fixed (pt)")}</option>
          <option value="lines" ${advancedValues.before.mode === "lines" ? "selected" : ""}>${tr("行", "Lines")}</option>
        </select>
      </label>
      <label>${tr("段前值", "Before Paragraph Value")}<input id="${styleInputId(key, "beforeValue")}" type="number" min="0" step="0.1" value="${advancedValues.before.value}" /></label>
      <label>${tr("段后模式", "After Paragraph Mode")}
        <select id="${styleInputId(key, "afterMode")}">
          <option value="pt" ${advancedValues.after.mode === "pt" ? "selected" : ""}>${tr("固定值(pt)", "Fixed (pt)")}</option>
          <option value="lines" ${advancedValues.after.mode === "lines" ? "selected" : ""}>${tr("行", "Lines")}</option>
        </select>
      </label>
      <label>${tr("段后值", "After Paragraph Value")}<input id="${styleInputId(key, "afterValue")}" type="number" min="0" step="0.1" value="${advancedValues.after.value}" /></label>
      <label>${tr("首行缩进(字符)", "First Line Indent (chars)")}<input id="${styleInputId(key, "firstLineIndentChars")}" type="number" min="0" step="0.5" value="${advancedValues.firstLineIndentChars}" /></label>
      <div class="style-inline-checks">
        <label class="compact-check"><input id="${styleInputId(key, "bold")}" type="checkbox" ${advancedValues.bold ? "checked" : ""} />${tr("加粗", "Bold")}</label>
        <label class="compact-check"><input id="${styleInputId(key, "italic")}" type="checkbox" ${advancedValues.italic ? "checked" : ""} />${tr("斜体", "Italic")}</label>
      </div>
    </div>
  </article>
`;

  if (isCollapsibleSection(key)) {
    return `
    <details class="style-section" data-style-key="${key}">
      <summary>${label}</summary>
      ${fields}
    </details>
`;
  }

  return `
  <section class="style-section style-section-static" data-style-key="${key}">
    <h3>${label}</h3>
    ${fields}
  </section>
`;
};

const renderStyleCards = (config: MdToDocxStyleConfig) =>
  STYLE_SECTIONS.map((key) => renderStyleCard(key, getStyleSectionLabel(key), config[key])).join("\n");

const renderTableSettingsCard = (settings: MdToDocxTableSettings) => {
  const textStyle = settings.textStyle.advancedOverride
    ? settings.textStyle
    : { ...settings.textStyle, advancedOverride: DEFAULT_ADVANCED_SETTINGS };
  const advancedValues = textStyle.advancedOverride ?? DEFAULT_ADVANCED_SETTINGS;

  return `
  <details class="style-section" data-style-key="table-settings">
    <summary>${tr("表格样式", "Table Style")}</summary>
    <article class="style-card" data-style-key="table-settings-body">
      <div class="style-grid-fields">
        <label>${tr("表格预设", "Table Preset")}
          <select id="table-tablePreset">
            ${(["threeLine", "tableGrid", "table"] as TablePreset[])
              .map((preset) => `<option value="${preset}" ${settings.tablePreset === preset ? "selected" : ""}>${getTablePresetLabel(preset)}</option>`)
              .join("")}
          </select>
        </label>
        <div class="style-inline-checks">
          <label class="compact-check"><input id="table-headerBold" type="checkbox" ${settings.headerBold ? "checked" : ""} />${tr("首行加粗", "Header Bold")}</label>
        </div>
        <label>${tr("中文字体", "Chinese Font")}
          <select id="table-zhFont">
            ${renderSelectOptions(ZH_FONT_OPTIONS, textStyle.zhFont)}
          </select>
        </label>
        <label>${tr("英文字体", "English Font")}
          <select id="table-enFont">
            ${renderSelectOptions(EN_FONT_OPTIONS, textStyle.enFont)}
          </select>
        </label>
        <label>${tr("字号(pt)", "Font Size (pt)")}<input id="table-fontSizePt" type="number" min="1" step="0.5" value="${textStyle.fontSizePt}" /></label>
        <label>${tr("行距模式", "Line Spacing Mode")}
          <select id="table-lineSpacingMode">
            <option value="multiple" ${textStyle.lineSpacingMode === "multiple" ? "selected" : ""}>${tr("倍数", "Multiple")}</option>
            <option value="fixed" ${textStyle.lineSpacingMode === "fixed" ? "selected" : ""}>${tr("固定值(pt)", "Fixed (pt)")}</option>
          </select>
        </label>
        <label>${tr("行距值", "Line Spacing Value")}<input id="table-lineSpacingValue" type="number" min="0.1" step="0.1" value="${textStyle.lineSpacingValue}" /></label>
        <label>${tr("对齐", "Alignment")}
          <select id="table-align">
            <option value="left" ${textStyle.align === "left" ? "selected" : ""}>${tr("左对齐", "Left")}</option>
            <option value="center" ${textStyle.align === "center" ? "selected" : ""}>${tr("居中", "Center")}</option>
            <option value="right" ${textStyle.align === "right" ? "selected" : ""}>${tr("右对齐", "Right")}</option>
            <option value="justify" ${textStyle.align === "justify" ? "selected" : ""}>${tr("两端对齐", "Justify")}</option>
          </select>
        </label>
        <label>${tr("段前模式", "Before Paragraph Mode")}
          <select id="table-beforeMode">
            <option value="pt" ${advancedValues.before.mode === "pt" ? "selected" : ""}>${tr("固定值(pt)", "Fixed (pt)")}</option>
            <option value="lines" ${advancedValues.before.mode === "lines" ? "selected" : ""}>${tr("行", "Lines")}</option>
          </select>
        </label>
        <label>${tr("段前值", "Before Paragraph Value")}<input id="table-beforeValue" type="number" min="0" step="0.1" value="${advancedValues.before.value}" /></label>
        <label>${tr("段后模式", "After Paragraph Mode")}
          <select id="table-afterMode">
            <option value="pt" ${advancedValues.after.mode === "pt" ? "selected" : ""}>${tr("固定值(pt)", "Fixed (pt)")}</option>
            <option value="lines" ${advancedValues.after.mode === "lines" ? "selected" : ""}>${tr("行", "Lines")}</option>
          </select>
        </label>
        <label>${tr("段后值", "After Paragraph Value")}<input id="table-afterValue" type="number" min="0" step="0.1" value="${advancedValues.after.value}" /></label>
        <label>${tr("首行缩进(字符)", "First Line Indent (chars)")}<input id="table-firstLineIndentChars" type="number" min="0" step="0.5" value="${advancedValues.firstLineIndentChars}" /></label>
        <div class="style-inline-checks">
          <label class="compact-check"><input id="table-bold" type="checkbox" ${advancedValues.bold ? "checked" : ""} />${tr("加粗", "Bold")}</label>
          <label class="compact-check"><input id="table-italic" type="checkbox" ${advancedValues.italic ? "checked" : ""} />${tr("斜体", "Italic")}</label>
        </div>
      </div>
    </article>
  </details>
`;
};

const renderStylePanelCards = (config: MdToDocxStyleConfig) =>
  `${renderStyleCards(config)}\n${renderTableSettingsCard(config.tableSettings)}`;

app.innerHTML = `
  <main class="app-shell">
    <header class="hero">
      <div class="hero-head">
        <div>
          <h1 id="heroTitle">文档格式转换工具</h1>
          <p id="heroDesc">支持 DOC/DOCX 与 Markdown 双向转换，并可自定义论文样式导出 DOCX。</p>
        </div>
        <button id="langToggle" type="button" class="btn btn-soft lang-toggle">中文 / EN</button>
      </div>
    </header>

    <section class="content-grid">
      <article class="panel controls-panel">
        <div class="panel-head">
          <h2 id="controlsTitle">转换设置</h2>
          <p id="controlsDesc">先选择转换方向，再配置输入输出和参数。</p>
        </div>

        <div id="directionGroup" class="mode-switch" role="radiogroup" aria-label="转换方向">
          <label class="mode-pill">
            <input type="radio" name="direction" value="docxToMd" checked />
            <span id="docxToMdLabel">DOC/DOCX → MD</span>
          </label>
          <label class="mode-pill">
            <input type="radio" name="direction" value="mdToDocx" />
            <span id="mdToDocxLabel">MD → DOCX</span>
          </label>
        </div>

        <section id="docxModePanel" class="mode-panel">
          <label id="docInputLabel" class="field-label" for="inputPath">输入文件</label>
          <div class="field-row">
            <input id="inputPath" type="text" placeholder="选择 .doc 或 .docx 文件" readonly />
            <button id="pickInput" class="btn">选择</button>
          </div>

          <label id="docOutputLabel" class="field-label" for="outputPath">输出 Markdown</label>
          <div class="field-row">
            <input id="outputPath" type="text" placeholder="默认与源文件同目录" readonly />
            <button id="pickOutput" class="btn">另存为</button>
          </div>

          <label class="checkbox-row">
            <input id="extractImages" type="checkbox" />
            <span id="extractImagesText">导出图片（自动创建源目录下图片文件夹）</span>
          </label>

          <label class="checkbox-row">
            <input id="splitSections" type="checkbox" />
            <span id="splitSectionsText">按章节拆分为多个 md 文件</span>
          </label>

          <div class="actions-row">
            <button id="openDocOutputDir" type="button" class="btn btn-soft">打开输出文件夹</button>
            <button id="run" class="btn btn-primary run">转换为MD</button>
          </div>
        </section>

        <section id="mdModePanel" class="mode-panel hidden">
          <label id="mdInputLabel" class="field-label" for="mdInputPath">输入 Markdown</label>
          <div class="field-row">
            <input id="mdInputPath" type="text" placeholder="选择 .md 文件" readonly />
            <button id="pickMdInput" class="btn">选择</button>
          </div>

          <label id="docxOutputLabel" class="field-label" for="docxOutputPath">输出 DOCX</label>
          <div class="field-row">
            <input id="docxOutputPath" type="text" placeholder="默认与源文件同目录" readonly />
            <button id="pickDocxOutput" class="btn">另存为</button>
          </div>

          <div class="style-toolbar">
            <button id="toggleStylePanel" type="button" class="btn btn-soft">展开样式设置</button>
            <div class="preset-toolbar">
              <button id="resetStylePreset" type="button" class="btn btn-soft">恢复默认样式</button>
              <button id="loadStylePreset" type="button" class="btn btn-soft">加载样式预设</button>
              <button id="saveStylePreset" type="button" class="btn btn-soft">保存样式预设</button>
            </div>
          </div>

          <div id="stylePanel" class="style-panel hidden">
            <div id="styleCards" class="style-cards">${renderStylePanelCards(DEFAULT_STYLE_CONFIG)}</div>
          </div>

          <div class="actions-row actions-row-md">
            <button id="openDocxOutputDir" type="button" class="btn btn-soft">打开输出文件夹</button>
            <button id="runMdToDocx" class="btn btn-primary run">转换为Docx</button>
          </div>
        </section>

        <p id="status" class="status-pill" data-state="idle">等待操作</p>
      </article>

      <article class="panel preview-panel">
        <div class="panel-head">
          <h2 id="previewTitle">Markdown 预览</h2>
          <p id="previewDesc">显示当前转换结果文本与日志。</p>
        </div>
        <pre id="outputView" class="markdown">暂无内容</pre>
      </article>
    </section>
  </main>
`;

const docxModePanel = document.querySelector<HTMLElement>("#docxModePanel")!;
const mdModePanel = document.querySelector<HTMLElement>("#mdModePanel")!;
const outputViewEl = document.querySelector<HTMLPreElement>("#outputView")!;
const previewTitleEl = document.querySelector<HTMLHeadingElement>("#previewTitle")!;
const previewDescEl = document.querySelector<HTMLParagraphElement>("#previewDesc")!;

const inputPathEl = document.querySelector<HTMLInputElement>("#inputPath")!;
const outputPathEl = document.querySelector<HTMLInputElement>("#outputPath")!;
const extractImagesEl = document.querySelector<HTMLInputElement>("#extractImages")!;
const splitSectionsEl = document.querySelector<HTMLInputElement>("#splitSections")!;
const mdInputPathEl = document.querySelector<HTMLInputElement>("#mdInputPath")!;
const docxOutputPathEl = document.querySelector<HTMLInputElement>("#docxOutputPath")!;

const statusEl = document.querySelector<HTMLParagraphElement>("#status")!;

const pickInputButton = document.querySelector<HTMLButtonElement>("#pickInput")!;
const pickOutputButton = document.querySelector<HTMLButtonElement>("#pickOutput")!;
const openDocOutputDirButton = document.querySelector<HTMLButtonElement>("#openDocOutputDir")!;
const runButton = document.querySelector<HTMLButtonElement>("#run")!;

const pickMdInputButton = document.querySelector<HTMLButtonElement>("#pickMdInput")!;
const pickDocxOutputButton = document.querySelector<HTMLButtonElement>("#pickDocxOutput")!;
const openDocxOutputDirButton = document.querySelector<HTMLButtonElement>("#openDocxOutputDir")!;
const runMdToDocxButton = document.querySelector<HTMLButtonElement>("#runMdToDocx")!;
const resetStylePresetButton = document.querySelector<HTMLButtonElement>("#resetStylePreset")!;
const loadStylePresetButton = document.querySelector<HTMLButtonElement>("#loadStylePreset")!;
const saveStylePresetButton = document.querySelector<HTMLButtonElement>("#saveStylePreset")!;
const toggleStylePanelButton = document.querySelector<HTMLButtonElement>("#toggleStylePanel")!;
const langToggleButton = document.querySelector<HTMLButtonElement>("#langToggle")!;
const styleCardsEl = document.querySelector<HTMLDivElement>("#styleCards")!;
const stylePanelEl = document.querySelector<HTMLElement>("#stylePanel")!;
const directionGroupEl = document.querySelector<HTMLElement>("#directionGroup")!;

const directionRadios = Array.from(
  document.querySelectorAll<HTMLInputElement>('input[name="direction"]'),
);

let currentDirection: Direction = "docxToMd";
let tableApplyTextStyle = DEFAULT_STYLE_CONFIG.tableSettings.applyTextStyle;

const setStatus = (message: string, state: "idle" | "info" | "success" | "error" = "info") => {
  statusEl.textContent = message;
  statusEl.dataset.state = state;
};

const getIdleText = () => tr("等待操作", "Waiting for action");
const getNoContentText = () => tr("暂无内容", "No content yet");
const getPreviewTitle = (direction: Direction) =>
  direction === "docxToMd" ? tr("Markdown 预览", "Markdown Preview") : tr("DOCX 导出日志", "DOCX Export Logs");
const getPreviewDescription = (direction: Direction) =>
  direction === "docxToMd"
    ? tr("显示当前转换后的 Markdown 文本。", "Show converted Markdown content.")
    : tr("显示导出结果与日志。", "Show export results and logs.");

const setStylePanelExpanded = (expanded: boolean) => {
  stylePanelEl.classList.toggle("hidden", !expanded);
  toggleStylePanelButton.textContent = expanded
    ? tr("样式 ▲", "Style ▲")
    : tr("样式 ▼", "Style ▼");
};

const rerenderStyleCards = () => {
  let preservedConfig: MdToDocxStyleConfig = DEFAULT_STYLE_CONFIG;
  try {
    preservedConfig = readStyleConfigFromForm();
  } catch {
    preservedConfig = DEFAULT_STYLE_CONFIG;
  }

  styleCardsEl.innerHTML = renderStylePanelCards(preservedConfig);
  applyStyleConfigToForm(preservedConfig);
};

const applyLocaleToUI = () => {
  const setText = (selector: string, text: string) => {
    const el = document.querySelector<HTMLElement>(selector);
    if (el) {
      el.textContent = text;
    }
  };

  setText("#heroTitle", tr("文档格式转换工具", "Document Format Converter"));
  setText(
    "#heroDesc",
    tr(
      "支持 DOC/DOCX 与 Markdown 双向转换，并可自定义论文样式导出 DOCX。",
      "Supports DOC/DOCX and Markdown conversion in both directions, with customizable paper style export to DOCX.",
    ),
  );

  setText("#controlsTitle", tr("转换设置", "Conversion Settings"));
  setText("#controlsDesc", tr("先选择转换方向，再配置输入输出和参数。", "Pick a direction, then configure input/output and options."));
  directionGroupEl.setAttribute("aria-label", tr("转换方向", "Conversion Direction"));

  setText("#docxToMdLabel", "DOC/DOCX → MD");
  setText("#mdToDocxLabel", "MD → DOCX");

  setText("#docInputLabel", tr("输入文件", "Input File"));
  setText("#docOutputLabel", tr("输出 Markdown", "Output Markdown"));
  setText("#extractImagesText", tr("导出图片（自动创建源目录下图片文件夹）", "Export images (auto-create image folder beside source file)"));
  setText("#splitSectionsText", tr("按章节拆分为多个 md 文件", "Split output into multiple Markdown files by sections"));

  inputPathEl.placeholder = tr("选择 .doc 或 .docx 文件", "Choose a .doc or .docx file");
  outputPathEl.placeholder = tr("默认与源文件同目录", "Default: same folder as source");

  pickInputButton.textContent = tr("选择", "Choose");
  pickOutputButton.textContent = tr("另存为", "Save As");
  openDocOutputDirButton.textContent = tr("打开输出文件夹", "Open Output Folder");
  runButton.textContent = tr("转换为MD", "Convert to MD");

  setText("#mdInputLabel", tr("输入 Markdown", "Input Markdown"));
  setText("#docxOutputLabel", tr("输出 DOCX", "Output DOCX"));
  mdInputPathEl.placeholder = tr("选择 .md 文件", "Choose a .md file");
  docxOutputPathEl.placeholder = tr("默认与源文件同目录", "Default: same folder as source");

  pickMdInputButton.textContent = tr("选择", "Choose");
  pickDocxOutputButton.textContent = tr("另存为", "Save As");
  openDocxOutputDirButton.textContent = tr("打开输出文件夹", "Open Output Folder");
  runMdToDocxButton.textContent = tr("转换为Docx", "Convert to DOCX");
  resetStylePresetButton.textContent = tr("恢复默认样式", "Reset Style");
  loadStylePresetButton.textContent = tr("加载样式预设", "Load Preset");
  saveStylePresetButton.textContent = tr("保存样式预设", "Save Preset");

  langToggleButton.textContent = currentLocale === "zh" ? "中文 / EN" : "EN / 中文";

  setText("#previewTitle", getPreviewTitle(currentDirection));
  setText("#previewDesc", getPreviewDescription(currentDirection));
  if (outputViewEl.textContent === "暂无内容" || outputViewEl.textContent === "No content yet") {
    outputViewEl.textContent = getNoContentText();
  }

  rerenderStyleCards();
  setStylePanelExpanded(!stylePanelEl.classList.contains("hidden"));
};

const splitPath = (fullPath: string) => {
  const lastSep = Math.max(fullPath.lastIndexOf("\\"), fullPath.lastIndexOf("/"));
  if (lastSep < 0) {
    return { dir: ".", file: fullPath, sep: "/" };
  }
  return {
    dir: fullPath.slice(0, lastSep),
    file: fullPath.slice(lastSep + 1),
    sep: fullPath.includes("\\") ? "\\" : "/",
  };
};

const stemFromPath = (path: string) => {
  const { file } = splitPath(path);
  const dot = file.lastIndexOf(".");
  return dot > 0 ? file.slice(0, dot) : file;
};

const defaultOutputFromDocInput = (inputPath: string) => {
  const { dir, sep } = splitPath(inputPath);
  const stem = stemFromPath(inputPath);
  return `${dir}${sep}${stem}.md`;
};

const defaultImageDirFromMarkdownOutput = (markdownPath: string) => {
  const { dir, sep } = splitPath(markdownPath);
  const stem = stemFromPath(markdownPath);
  return `${dir}${sep}${stem}_assets${sep}images`;
};

const defaultImageDirFromInput = (inputPath: string) =>
  defaultImageDirFromMarkdownOutput(defaultOutputFromDocInput(inputPath));

const defaultDocxOutputFromMd = (inputPath: string) => {
  const { dir, sep } = splitPath(inputPath);
  const stem = stemFromPath(inputPath);
  return `${dir}${sep}${stem}.docx`;
};

const switchDirection = (direction: Direction) => {
  currentDirection = direction;
  const isDocToMd = direction === "docxToMd";
  docxModePanel.classList.toggle("hidden", !isDocToMd);
  mdModePanel.classList.toggle("hidden", isDocToMd);

  previewTitleEl.textContent = getPreviewTitle(direction);
  previewDescEl.textContent = getPreviewDescription(direction);
  outputViewEl.textContent = getNoContentText();
  setStatus(getIdleText(), "idle");
  if (!isDocToMd) {
    setStylePanelExpanded(false);
  }
};

const openFolderForPath = async (path: string, successLabel?: string) => {
  if (!path) {
    setStatus(tr("请先设置输出路径", "Please set output path first"), "error");
    return;
  }
  try {
    await invoke("open_folder_for_path", {
      request: { path },
    });
    setStatus(successLabel ?? tr("已打开输出文件夹", "Output folder opened"), "success");
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    setStatus(`${tr("打开失败：", "Open failed: ")}${message}`, "error");
  }
};

const requireElement = <T extends HTMLElement>(selector: string) => {
  const element = document.querySelector<T>(selector);
  if (!element) {
    throw new Error(`Missing element: ${selector}`);
  }
  return element;
};

const parsePositiveNumber = (rawValue: string, fieldLabel: string) => {
  const value = Number.parseFloat(rawValue);
  if (!Number.isFinite(value) || value <= 0) {
    throw new Error(`${fieldLabel}${tr(" 必须是正数", " must be a positive number")}`);
  }
  return value;
};

const parseNonNegativeNumber = (rawValue: string, fieldLabel: string) => {
  const value = Number.parseFloat(rawValue);
  if (!Number.isFinite(value) || value < 0) {
    throw new Error(`${fieldLabel}${tr(" 不能为负数", " cannot be negative")}`);
  }
  return value;
};

const ensureSelectHasValue = (select: HTMLSelectElement, value: string) => {
  const exists = Array.from(select.options).some((option) => option.value === value);
  if (!exists) {
    const option = document.createElement("option");
    option.value = value;
    option.text = value;
    select.appendChild(option);
  }
  select.value = value;
};

const parseSpacingMode = (value: unknown): SpacingMode => (value === "lines" ? "lines" : "pt");
const parseTablePreset = (value: unknown): TablePreset => {
  if (value === "threeLine" || value === "table" || value === "tableGrid") {
    return value;
  }
  return "tableGrid";
};

const normalizeAdvancedSettings = (
  raw: unknown,
  fallback: StyleAdvancedSettings,
): StyleAdvancedSettings => {
  const value = (raw ?? {}) as Record<string, unknown>;
  const beforeRaw = value.before as Record<string, unknown> | undefined;
  const afterRaw = value.after as Record<string, unknown> | undefined;
  const legacyBeforePt = value.beforePt;
  const legacyAfterPt = value.afterPt;

  const normalizedBeforeValue = Number.isFinite(Number(beforeRaw?.value))
    ? Number(beforeRaw?.value)
    : Number.isFinite(Number(legacyBeforePt))
      ? Number(legacyBeforePt)
      : fallback.before.value;
  const normalizedAfterValue = Number.isFinite(Number(afterRaw?.value))
    ? Number(afterRaw?.value)
    : Number.isFinite(Number(legacyAfterPt))
      ? Number(legacyAfterPt)
      : fallback.after.value;

  return {
    before: {
      mode: parseSpacingMode(beforeRaw?.mode),
      value: normalizedBeforeValue,
    },
    after: {
      mode: parseSpacingMode(afterRaw?.mode),
      value: normalizedAfterValue,
    },
    firstLineIndentChars: Number.isFinite(Number(value.firstLineIndentChars))
      ? Number(value.firstLineIndentChars)
      : fallback.firstLineIndentChars,
    bold: typeof value.bold === "boolean" ? value.bold : fallback.bold,
    italic: typeof value.italic === "boolean" ? value.italic : fallback.italic,
  };
};

const normalizeStyleBlock = (
  raw: unknown,
  fallback: MdToDocxStyleBlock,
  defaults: StyleAdvancedSettings,
): MdToDocxStyleBlock => {
  const withCompat = (raw ?? fallback) as MdToDocxStyleBlock;

  let advancedOverride = normalizeAdvancedSettings(withCompat.advancedOverride, defaults);
  if (
    (!withCompat.advancedOverride || typeof withCompat.advancedOverride !== "object") &&
    (typeof withCompat.bold === "boolean" || typeof withCompat.italic === "boolean")
  ) {
    advancedOverride = normalizeAdvancedSettings(
      { bold: withCompat.bold, italic: withCompat.italic },
      defaults,
    );
  }

  return {
    zhFont: withCompat.zhFont ?? fallback.zhFont,
    enFont: withCompat.enFont ?? fallback.enFont,
    fontSizePt: Number.isFinite(Number(withCompat.fontSizePt))
      ? Number(withCompat.fontSizePt)
      : fallback.fontSizePt,
    lineSpacingMode: withCompat.lineSpacingMode === "fixed" ? "fixed" : "multiple",
    lineSpacingValue: Number.isFinite(Number(withCompat.lineSpacingValue))
      ? Number(withCompat.lineSpacingValue)
      : fallback.lineSpacingValue,
    align:
      withCompat.align === "center" ||
      withCompat.align === "right" ||
      withCompat.align === "justify"
        ? withCompat.align
        : "left",
    advancedOverride,
  };
};

const normalizeStyleConfig = (raw: MdToDocxStyleConfig): MdToDocxStyleConfig => {
  const base = DEFAULT_STYLE_CONFIG;
  const defaults = normalizeAdvancedSettings(
    (raw as unknown as { advancedDefaults?: unknown }).advancedDefaults,
    base.advancedDefaults,
  );

  const normalizedBlock = (key: StyleSectionKey): MdToDocxStyleBlock =>
    normalizeStyleBlock(
      (raw as unknown as Partial<MdToDocxStyleConfig>)[key],
      base[key],
      defaults,
    );

  const rawTableSettings = (raw as unknown as { tableSettings?: unknown }).tableSettings;
  let normalizedTableSettings = { ...COMPAT_TABLE_SETTINGS };
  if (rawTableSettings && typeof rawTableSettings === "object") {
    const value = rawTableSettings as Partial<MdToDocxTableSettings>;
    normalizedTableSettings = {
      tablePreset: parseTablePreset(value.tablePreset),
      headerBold: Boolean(value.headerBold),
      textStyle: normalizeStyleBlock(value.textStyle, base.tableSettings.textStyle, defaults),
      applyTextStyle: typeof value.applyTextStyle === "boolean"
        ? value.applyTextStyle
        : true,
    };
  } else {
    normalizedTableSettings.textStyle = normalizeStyleBlock(
      base.tableSettings.textStyle,
      base.tableSettings.textStyle,
      defaults,
    );
  }

  return {
    title: normalizedBlock("title"),
    abstractZh: normalizedBlock("abstractZh"),
    abstractEn: normalizedBlock("abstractEn"),
    heading1: normalizedBlock("heading1"),
    heading2: normalizedBlock("heading2"),
    heading3: normalizedBlock("heading3"),
    figureCaption: normalizedBlock("figureCaption"),
    tableCaption: normalizedBlock("tableCaption"),
    body: normalizedBlock("body"),
    tableSettings: normalizedTableSettings,
    advancedDefaults: defaults,
  };
};

const readStyleBlock = (key: StyleSectionKey, label: string): MdToDocxStyleBlock => {
  const zhFont = requireElement<HTMLSelectElement>(`#${styleInputId(key, "zhFont")}`).value.trim();
  const enFont = requireElement<HTMLSelectElement>(`#${styleInputId(key, "enFont")}`).value.trim();
  const fontSizePt = parsePositiveNumber(
    requireElement<HTMLInputElement>(`#${styleInputId(key, "fontSizePt")}`).value,
    `${label}${tr(" 字号", " font size")}`,
  );
  const lineSpacingMode = requireElement<HTMLSelectElement>(
    `#${styleInputId(key, "lineSpacingMode")}`,
  ).value as LineSpacingMode;
  const lineSpacingValue = parsePositiveNumber(
    requireElement<HTMLInputElement>(`#${styleInputId(key, "lineSpacingValue")}`).value,
    `${label}${tr(" 行距值", " line spacing value")}`,
  );
  const align = requireElement<HTMLSelectElement>(`#${styleInputId(key, "align")}`).value as AlignMode;

  if (!zhFont) {
    throw new Error(`${label}${tr(" 中文字体不能为空", " Chinese font cannot be empty")}`);
  }
  if (!enFont) {
    throw new Error(`${label}${tr(" 英文字体不能为空", " English font cannot be empty")}`);
  }

  const beforeMode = parseSpacingMode(
    requireElement<HTMLSelectElement>(`#${styleInputId(key, "beforeMode")}`).value,
  );
  const beforeValue = parseNonNegativeNumber(
    requireElement<HTMLInputElement>(`#${styleInputId(key, "beforeValue")}`).value,
    `${label}${tr(" 段前值", " before paragraph value")}`,
  );
  const afterMode = parseSpacingMode(
    requireElement<HTMLSelectElement>(`#${styleInputId(key, "afterMode")}`).value,
  );
  const afterValue = parseNonNegativeNumber(
    requireElement<HTMLInputElement>(`#${styleInputId(key, "afterValue")}`).value,
    `${label}${tr(" 段后值", " after paragraph value")}`,
  );

  const advancedOverride = {
    before: {
      mode: beforeMode,
      value: beforeValue,
    },
    after: {
      mode: afterMode,
      value: afterValue,
    },
    firstLineIndentChars: parseNonNegativeNumber(
      requireElement<HTMLInputElement>(`#${styleInputId(key, "firstLineIndentChars")}`).value,
      `${label}${tr(" 首行缩进", " first line indent")}`,
    ),
    bold: requireElement<HTMLInputElement>(`#${styleInputId(key, "bold")}`).checked,
    italic: requireElement<HTMLInputElement>(`#${styleInputId(key, "italic")}`).checked,
  };

  return {
    zhFont,
    enFont,
    fontSizePt,
    lineSpacingMode,
    lineSpacingValue,
    align,
    advancedOverride,
  };
};

const readTableTextStyleFromForm = (): MdToDocxStyleBlock => {
  const zhFont = requireElement<HTMLSelectElement>("#table-zhFont").value.trim();
  const enFont = requireElement<HTMLSelectElement>("#table-enFont").value.trim();
  const fontSizePt = parsePositiveNumber(
    requireElement<HTMLInputElement>("#table-fontSizePt").value,
    tr("表格 字号", "Table font size"),
  );
  const lineSpacingMode = requireElement<HTMLSelectElement>("#table-lineSpacingMode")
    .value as LineSpacingMode;
  const lineSpacingValue = parsePositiveNumber(
    requireElement<HTMLInputElement>("#table-lineSpacingValue").value,
    tr("表格 行距值", "Table line spacing value"),
  );
  const align = requireElement<HTMLSelectElement>("#table-align").value as AlignMode;

  if (!zhFont) {
    throw new Error(tr("表格 中文字体不能为空", "Table Chinese font cannot be empty"));
  }
  if (!enFont) {
    throw new Error(tr("表格 英文字体不能为空", "Table English font cannot be empty"));
  }

  const beforeMode = parseSpacingMode(requireElement<HTMLSelectElement>("#table-beforeMode").value);
  const beforeValue = parseNonNegativeNumber(
    requireElement<HTMLInputElement>("#table-beforeValue").value,
    tr("表格 段前值", "Table before paragraph value"),
  );
  const afterMode = parseSpacingMode(requireElement<HTMLSelectElement>("#table-afterMode").value);
  const afterValue = parseNonNegativeNumber(
    requireElement<HTMLInputElement>("#table-afterValue").value,
    tr("表格 段后值", "Table after paragraph value"),
  );

  return {
    zhFont,
    enFont,
    fontSizePt,
    lineSpacingMode,
    lineSpacingValue,
    align,
    advancedOverride: {
      before: {
        mode: beforeMode,
        value: beforeValue,
      },
      after: {
        mode: afterMode,
        value: afterValue,
      },
      firstLineIndentChars: parseNonNegativeNumber(
        requireElement<HTMLInputElement>("#table-firstLineIndentChars").value,
        tr("表格 首行缩进", "Table first line indent"),
      ),
      bold: requireElement<HTMLInputElement>("#table-bold").checked,
      italic: requireElement<HTMLInputElement>("#table-italic").checked,
    },
  };
};

const readTableSettingsFromForm = (): MdToDocxTableSettings => {
  const tablePreset = parseTablePreset(
    requireElement<HTMLSelectElement>("#table-tablePreset").value,
  );
  const headerBold = requireElement<HTMLInputElement>("#table-headerBold").checked;
  return {
    tablePreset,
    headerBold,
    textStyle: readTableTextStyleFromForm(),
    applyTextStyle: tableApplyTextStyle,
  };
};

const readStyleConfigFromForm = (): MdToDocxStyleConfig => {
  const values: Partial<MdToDocxStyleConfig> = {
    advancedDefaults: { ...DEFAULT_ADVANCED_SETTINGS },
  };

  for (const key of STYLE_SECTIONS) {
    values[key] = readStyleBlock(key, getStyleSectionLabel(key));
  }
  values.tableSettings = readTableSettingsFromForm();
  return values as MdToDocxStyleConfig;
};

const applyStyleBlock = (
  key: StyleSectionKey,
  block: MdToDocxStyleBlock,
) => {
  ensureSelectHasValue(requireElement<HTMLSelectElement>(`#${styleInputId(key, "zhFont")}`), block.zhFont);
  ensureSelectHasValue(requireElement<HTMLSelectElement>(`#${styleInputId(key, "enFont")}`), block.enFont);
  requireElement<HTMLInputElement>(`#${styleInputId(key, "fontSizePt")}`).value = String(block.fontSizePt);
  requireElement<HTMLSelectElement>(`#${styleInputId(key, "lineSpacingMode")}`).value = block.lineSpacingMode;
  requireElement<HTMLInputElement>(`#${styleInputId(key, "lineSpacingValue")}`).value = String(block.lineSpacingValue);
  requireElement<HTMLSelectElement>(`#${styleInputId(key, "align")}`).value = block.align;

  const advanced = block.advancedOverride ?? DEFAULT_ADVANCED_SETTINGS;
  requireElement<HTMLSelectElement>(`#${styleInputId(key, "beforeMode")}`).value = advanced.before.mode;
  requireElement<HTMLInputElement>(`#${styleInputId(key, "beforeValue")}`).value = String(advanced.before.value);
  requireElement<HTMLSelectElement>(`#${styleInputId(key, "afterMode")}`).value = advanced.after.mode;
  requireElement<HTMLInputElement>(`#${styleInputId(key, "afterValue")}`).value = String(advanced.after.value);
  requireElement<HTMLInputElement>(`#${styleInputId(key, "firstLineIndentChars")}`).value = String(
    advanced.firstLineIndentChars,
  );
  requireElement<HTMLInputElement>(`#${styleInputId(key, "bold")}`).checked = advanced.bold;
  requireElement<HTMLInputElement>(`#${styleInputId(key, "italic")}`).checked = advanced.italic;
};

const applyTableSettingsToForm = (settings: MdToDocxTableSettings) => {
  requireElement<HTMLSelectElement>("#table-tablePreset").value = settings.tablePreset;
  requireElement<HTMLInputElement>("#table-headerBold").checked = settings.headerBold;

  const block = settings.textStyle;
  ensureSelectHasValue(requireElement<HTMLSelectElement>("#table-zhFont"), block.zhFont);
  ensureSelectHasValue(requireElement<HTMLSelectElement>("#table-enFont"), block.enFont);
  requireElement<HTMLInputElement>("#table-fontSizePt").value = String(block.fontSizePt);
  requireElement<HTMLSelectElement>("#table-lineSpacingMode").value = block.lineSpacingMode;
  requireElement<HTMLInputElement>("#table-lineSpacingValue").value = String(block.lineSpacingValue);
  requireElement<HTMLSelectElement>("#table-align").value = block.align;

  const advanced = block.advancedOverride ?? DEFAULT_ADVANCED_SETTINGS;
  requireElement<HTMLSelectElement>("#table-beforeMode").value = advanced.before.mode;
  requireElement<HTMLInputElement>("#table-beforeValue").value = String(advanced.before.value);
  requireElement<HTMLSelectElement>("#table-afterMode").value = advanced.after.mode;
  requireElement<HTMLInputElement>("#table-afterValue").value = String(advanced.after.value);
  requireElement<HTMLInputElement>("#table-firstLineIndentChars").value = String(
    advanced.firstLineIndentChars,
  );
  requireElement<HTMLInputElement>("#table-bold").checked = advanced.bold;
  requireElement<HTMLInputElement>("#table-italic").checked = advanced.italic;
  tableApplyTextStyle = settings.applyTextStyle;
};

const applyStyleConfigToForm = (config: MdToDocxStyleConfig) => {
  const normalized = normalizeStyleConfig(config);

  for (const key of STYLE_SECTIONS) {
    applyStyleBlock(key, normalized[key]);
  }
  applyTableSettingsToForm(normalized.tableSettings);
};

pickInputButton.addEventListener("click", async () => {
  const selected = await open({
    multiple: false,
    filters: [{ name: "Word", extensions: ["doc", "docx"] }],
  });
  if (typeof selected === "string") {
    inputPathEl.value = selected;
    outputPathEl.value = defaultOutputFromDocInput(selected);
    setStatus(tr("已选择输入文件", "Input file selected"), "success");
  }
});

pickOutputButton.addEventListener("click", async () => {
  const defaultPath = inputPathEl.value ? defaultOutputFromDocInput(inputPathEl.value) : "output.md";
  const selected = await save({
    filters: [{ name: "Markdown", extensions: ["md"] }],
    defaultPath,
  });
  if (typeof selected === "string") {
    outputPathEl.value = selected;
    setStatus(tr("已设置输出路径", "Output path set"), "success");
  }
});

openDocOutputDirButton.addEventListener("click", async () => {
  if (!outputPathEl.value && inputPathEl.value) {
    outputPathEl.value = defaultOutputFromDocInput(inputPathEl.value);
  }
  await openFolderForPath(outputPathEl.value);
});

extractImagesEl.addEventListener("change", () => {
  if (extractImagesEl.checked && inputPathEl.value) {
    const markdownPath = outputPathEl.value || defaultOutputFromDocInput(inputPathEl.value);
    setStatus(
      `${tr("图片将导出到 ", "Images will be exported to ")}${defaultImageDirFromMarkdownOutput(markdownPath)}`,
      "info",
    );
  }
});

runButton.addEventListener("click", async () => {
  if (!inputPathEl.value) {
    setStatus(tr("请先选择输入文件", "Please select an input file first"), "error");
    return;
  }

  if (!outputPathEl.value) {
    outputPathEl.value = defaultOutputFromDocInput(inputPathEl.value);
  }

  const request: ConvertRequest = {
    inputPath: inputPathEl.value,
    outputPath: outputPathEl.value,
    extractImages: extractImagesEl.checked,
    splitSections: splitSectionsEl.checked,
    math: "latex",
  };

  runButton.disabled = true;
  setStatus(tr("转换中...", "Converting..."), "info");

  try {
    const response = await invoke<ConversionResponse>("convert_docx", { request });

    const originalOutput = outputPathEl.value;
    if (response.outputPath) {
      outputPathEl.value = response.outputPath;
    }
    outputViewEl.textContent = response.markdown || tr("无输出内容", "No output content");

    if (response.success) {
      const messages: string[] = [
        tr("导出成功", "Export succeeded"),
        `${tr("文件：", "File: ")}${response.outputPath || outputPathEl.value}`,
      ];
      if (response.outputPath && originalOutput && response.outputPath !== originalOutput) {
        messages.push(tr("检测到重名，已自动改名保存", "Name conflict detected; auto-renamed and saved"));
      }
      if (extractImagesEl.checked) {
        const markdownPath = response.outputPath || outputPathEl.value;
        messages.push(`${tr("图片目录：", "Images folder: ")}${defaultImageDirFromMarkdownOutput(markdownPath)}`);
      }
      if (splitSectionsEl.checked && response.splitOutputDir && response.splitFileCount) {
        messages.push(
          currentLocale === "zh"
            ? `章节拆分：${response.splitFileCount} 个文件`
            : `Section split: ${response.splitFileCount} files`,
        );
      }
      setStatus(messages.join(" | "), "success");
    } else {
      const fallback = response.stderr || tr(`转换失败（退出码 ${response.exitCode}）`, `Conversion failed (exit code ${response.exitCode})`);
      setStatus(fallback, "error");
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    setStatus(`${tr("调用失败：", "Invocation failed: ")}${message}`, "error");
  } finally {
    runButton.disabled = false;
  }
});

pickMdInputButton.addEventListener("click", async () => {
  const selected = await open({
    multiple: false,
    filters: [{ name: "Markdown", extensions: ["md"] }],
  });
  if (typeof selected === "string") {
    mdInputPathEl.value = selected;
    docxOutputPathEl.value = defaultDocxOutputFromMd(selected);
    setStatus(tr("已选择 Markdown 文件", "Markdown file selected"), "success");
  }
});

pickDocxOutputButton.addEventListener("click", async () => {
  const defaultPath = mdInputPathEl.value ? defaultDocxOutputFromMd(mdInputPathEl.value) : "output.docx";
  const selected = await save({
    filters: [{ name: "Word", extensions: ["docx"] }],
    defaultPath,
  });
  if (typeof selected === "string") {
    docxOutputPathEl.value = selected;
    setStatus(tr("已设置 DOCX 输出路径", "DOCX output path set"), "success");
  }
});

openDocxOutputDirButton.addEventListener("click", async () => {
  if (!docxOutputPathEl.value && mdInputPathEl.value) {
    docxOutputPathEl.value = defaultDocxOutputFromMd(mdInputPathEl.value);
  }
  await openFolderForPath(docxOutputPathEl.value);
});

toggleStylePanelButton.addEventListener("click", () => {
  const expanded = stylePanelEl.classList.contains("hidden");
  setStylePanelExpanded(expanded);
});

resetStylePresetButton.addEventListener("click", () => {
  applyStyleConfigToForm(DEFAULT_STYLE_CONFIG);
  setStatus(tr("已恢复默认样式", "Default style restored"), "success");
});

loadStylePresetButton.addEventListener("click", async () => {
  try {
    const selected = await open({
      multiple: false,
      filters: [{ name: "JSON", extensions: ["json"] }],
    });
    if (typeof selected !== "string") {
      return;
    }

    const config = await invoke<MdToDocxStyleConfig>("load_style_preset", {
      request: { path: selected },
    });
    applyStyleConfigToForm(config);
    setStatus(tr("已加载样式预设", "Style preset loaded"), "success");
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    setStatus(`${tr("加载预设失败：", "Failed to load preset: ")}${message}`, "error");
  }
});

saveStylePresetButton.addEventListener("click", async () => {
  try {
    const config = readStyleConfigFromForm();
    const defaultPath = mdInputPathEl.value
      ? `${splitPath(mdInputPathEl.value).dir}${splitPath(mdInputPathEl.value).sep}${stemFromPath(mdInputPathEl.value)}_style_preset.json`
      : "md_docx_style_preset.json";

    const selected = await save({
      filters: [{ name: "JSON", extensions: ["json"] }],
      defaultPath,
    });

    if (typeof selected !== "string") {
      return;
    }

    await invoke("save_style_preset", {
      request: {
        path: selected,
        styleConfig: config,
      },
    });
    setStatus(tr("样式预设已保存", "Style preset saved"), "success");
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    setStatus(`${tr("保存预设失败：", "Failed to save preset: ")}${message}`, "error");
  }
});

runMdToDocxButton.addEventListener("click", async () => {
  if (!mdInputPathEl.value) {
    setStatus(tr("请先选择 Markdown 输入文件", "Please select a Markdown input file first"), "error");
    return;
  }

  if (!docxOutputPathEl.value) {
    docxOutputPathEl.value = defaultDocxOutputFromMd(mdInputPathEl.value);
  }

  let styleConfig: MdToDocxStyleConfig;
  try {
    styleConfig = readStyleConfigFromForm();
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    setStatus(`${tr("样式配置无效：", "Invalid style configuration: ")}${message}`, "error");
    return;
  }

  const request: ConvertMdToDocxRequest = {
    inputPath: mdInputPathEl.value,
    outputPath: docxOutputPathEl.value,
    styleConfig,
  };

  runMdToDocxButton.disabled = true;
  setStatus(tr("导出 DOCX 中...", "Exporting DOCX..."), "info");

  try {
    const originalOutput = docxOutputPathEl.value;
    const response = await invoke<ConvertMdToDocxResponse>("convert_md_to_docx", { request });
    if (response.outputPath) {
      docxOutputPathEl.value = response.outputPath;
    }

    const logs = [`${tr("输出文件：", "Output file: ")}${response.outputPath}`];
    if (response.stdout) {
      logs.push("\n[stdout]\n" + response.stdout);
    }
    if (response.stderr) {
      logs.push("\n[stderr]\n" + response.stderr);
    }
    outputViewEl.textContent = logs.join("\n");

    if (response.success) {
      if (response.outputPath && originalOutput && response.outputPath !== originalOutput) {
        setStatus(
          tr("DOCX 导出成功（检测到重名，已自动改名保存）", "DOCX export succeeded (name conflict detected; auto-renamed and saved)"),
          "success",
        );
      } else {
        setStatus(tr("DOCX 导出成功", "DOCX export succeeded"), "success");
      }
    } else {
      setStatus(response.stderr || tr(`导出失败（退出码 ${response.exitCode}）`, `Export failed (exit code ${response.exitCode})`), "error");
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    setStatus(`${tr("调用失败：", "Invocation failed: ")}${message}`, "error");
  } finally {
    runMdToDocxButton.disabled = false;
  }
});

for (const radio of directionRadios) {
  radio.addEventListener("change", () => {
    if (radio.checked) {
      switchDirection(radio.value as Direction);
    }
  });
}

langToggleButton.addEventListener("click", () => {
  currentLocale = currentLocale === "zh" ? "en" : "zh";
  applyLocaleToUI();
  switchDirection(currentDirection);
  setStatus(
    currentLocale === "zh" ? "已切换到中文界面" : "Switched to English UI",
    "success",
  );
});

setStylePanelExpanded(false);
applyLocaleToUI();
switchDirection(currentDirection);
