import { invoke } from "@tauri-apps/api/core";
import { open, save } from "@tauri-apps/plugin-dialog";
import "./style.css";

type Direction = "docxToMd" | "mdToDocx";
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
  advancedDefaults: { ...DEFAULT_ADVANCED_SETTINGS },
};

type StyleSectionKey = Exclude<keyof MdToDocxStyleConfig, "advancedDefaults">;

const STYLE_SECTIONS: Array<{ key: StyleSectionKey; label: string }> = [
  { key: "title", label: "论文题目" },
  { key: "abstractZh", label: "中文摘要" },
  { key: "abstractEn", label: "英文摘要" },
  { key: "heading1", label: "一级标题" },
  { key: "heading2", label: "二级标题" },
  { key: "heading3", label: "三级标题" },
  { key: "figureCaption", label: "图标题" },
  { key: "tableCaption", label: "表标题" },
  { key: "body", label: "正文" },
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
      <label>中文字体
        <select id="${styleInputId(key, "zhFont")}">
          ${renderSelectOptions(ZH_FONT_OPTIONS, block.zhFont)}
        </select>
      </label>
      <label>英文字体
        <select id="${styleInputId(key, "enFont")}">
          ${renderSelectOptions(EN_FONT_OPTIONS, block.enFont)}
        </select>
      </label>
      <label>字号(pt)<input id="${styleInputId(key, "fontSizePt")}" type="number" min="1" step="0.5" value="${block.fontSizePt}" /></label>
      <label>行距模式
        <select id="${styleInputId(key, "lineSpacingMode")}">
          <option value="multiple" ${block.lineSpacingMode === "multiple" ? "selected" : ""}>倍数</option>
          <option value="fixed" ${block.lineSpacingMode === "fixed" ? "selected" : ""}>固定值(pt)</option>
        </select>
      </label>
      <label>行距值<input id="${styleInputId(key, "lineSpacingValue")}" type="number" min="0.1" step="0.1" value="${block.lineSpacingValue}" /></label>
      <label>对齐
        <select id="${styleInputId(key, "align")}">
          <option value="left" ${block.align === "left" ? "selected" : ""}>左对齐</option>
          <option value="center" ${block.align === "center" ? "selected" : ""}>居中</option>
          <option value="right" ${block.align === "right" ? "selected" : ""}>右对齐</option>
          <option value="justify" ${block.align === "justify" ? "selected" : ""}>两端对齐</option>
        </select>
      </label>
      <label>段前模式
        <select id="${styleInputId(key, "beforeMode")}">
          <option value="pt" ${advancedValues.before.mode === "pt" ? "selected" : ""}>固定值(pt)</option>
          <option value="lines" ${advancedValues.before.mode === "lines" ? "selected" : ""}>行</option>
        </select>
      </label>
      <label>段前值<input id="${styleInputId(key, "beforeValue")}" type="number" min="0" step="0.1" value="${advancedValues.before.value}" /></label>
      <label>段后模式
        <select id="${styleInputId(key, "afterMode")}">
          <option value="pt" ${advancedValues.after.mode === "pt" ? "selected" : ""}>固定值(pt)</option>
          <option value="lines" ${advancedValues.after.mode === "lines" ? "selected" : ""}>行</option>
        </select>
      </label>
      <label>段后值<input id="${styleInputId(key, "afterValue")}" type="number" min="0" step="0.1" value="${advancedValues.after.value}" /></label>
      <label>首行缩进(字符)<input id="${styleInputId(key, "firstLineIndentChars")}" type="number" min="0" step="0.5" value="${advancedValues.firstLineIndentChars}" /></label>
      <div class="style-inline-checks">
        <label class="compact-check"><input id="${styleInputId(key, "bold")}" type="checkbox" ${advancedValues.bold ? "checked" : ""} />加粗</label>
        <label class="compact-check"><input id="${styleInputId(key, "italic")}" type="checkbox" ${advancedValues.italic ? "checked" : ""} />斜体</label>
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
  STYLE_SECTIONS.map((item) => renderStyleCard(item.key, item.label, config[item.key])).join("\n");

app.innerHTML = `
  <main class="app-shell">
    <header class="hero">
      <h1>文档格式转换工具</h1>
      <p>支持 DOC/DOCX 与 Markdown 双向转换，并可自定义论文样式导出 DOCX。</p>
    </header>

    <section class="content-grid">
      <article class="panel controls-panel">
        <div class="panel-head">
          <h2>转换设置</h2>
          <p>先选择转换方向，再配置输入输出和参数。</p>
        </div>

        <div class="mode-switch" role="radiogroup" aria-label="转换方向">
          <label class="mode-pill">
            <input type="radio" name="direction" value="docxToMd" checked />
            <span>DOC/DOCX → MD</span>
          </label>
          <label class="mode-pill">
            <input type="radio" name="direction" value="mdToDocx" />
            <span>MD → DOCX</span>
          </label>
        </div>

        <section id="docxModePanel" class="mode-panel">
          <label class="field-label" for="inputPath">输入文件</label>
          <div class="field-row">
            <input id="inputPath" type="text" placeholder="选择 .doc 或 .docx 文件" readonly />
            <button id="pickInput" class="btn">选择</button>
          </div>

          <label class="field-label" for="outputPath">输出 Markdown</label>
          <div class="field-row">
            <input id="outputPath" type="text" placeholder="默认与源文件同目录" readonly />
            <button id="pickOutput" class="btn">另存为</button>
          </div>

          <label class="checkbox-row">
            <input id="extractImages" type="checkbox" />
            <span>导出图片（自动创建源目录下图片文件夹）</span>
          </label>

          <label class="checkbox-row">
            <input id="splitSections" type="checkbox" />
            <span>按章节拆分为多个 md 文件</span>
          </label>

          <div class="actions-row">
            <button id="openDocOutputDir" type="button" class="btn btn-soft">打开输出文件夹</button>
            <button id="run" class="btn btn-primary run">转换为MD</button>
          </div>
        </section>

        <section id="mdModePanel" class="mode-panel hidden">
          <label class="field-label" for="mdInputPath">输入 Markdown</label>
          <div class="field-row">
            <input id="mdInputPath" type="text" placeholder="选择 .md 文件" readonly />
            <button id="pickMdInput" class="btn">选择</button>
          </div>

          <label class="field-label" for="docxOutputPath">输出 DOCX</label>
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
            <div class="style-cards">${renderStyleCards(DEFAULT_STYLE_CONFIG)}</div>
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
const stylePanelEl = document.querySelector<HTMLElement>("#stylePanel")!;

const directionRadios = Array.from(
  document.querySelectorAll<HTMLInputElement>('input[name="direction"]'),
);

let currentDirection: Direction = "docxToMd";

const setStatus = (message: string, state: "idle" | "info" | "success" | "error" = "info") => {
  statusEl.textContent = message;
  statusEl.dataset.state = state;
};

const setStylePanelExpanded = (expanded: boolean) => {
  stylePanelEl.classList.toggle("hidden", !expanded);
  toggleStylePanelButton.textContent = expanded ? "收起样式设置 ▲" : "展开样式设置 ▼";
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

  previewTitleEl.textContent = isDocToMd ? "Markdown 预览" : "DOCX 导出日志";
  previewDescEl.textContent = isDocToMd
    ? "显示当前转换后的 Markdown 文本。"
    : "显示导出结果与日志。";
  outputViewEl.textContent = "暂无内容";
  setStatus("等待操作", "idle");
  if (!isDocToMd) {
    setStylePanelExpanded(false);
  }
};

const openFolderForPath = async (path: string, successLabel = "已打开输出文件夹") => {
  if (!path) {
    setStatus("请先设置输出路径", "error");
    return;
  }
  try {
    await invoke("open_folder_for_path", {
      request: { path },
    });
    setStatus(successLabel, "success");
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    setStatus(`打开失败：${message}`, "error");
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
    throw new Error(`${fieldLabel} 必须是正数`);
  }
  return value;
};

const parseNonNegativeNumber = (rawValue: string, fieldLabel: string) => {
  const value = Number.parseFloat(rawValue);
  if (!Number.isFinite(value) || value < 0) {
    throw new Error(`${fieldLabel} 不能为负数`);
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

const normalizeStyleConfig = (raw: MdToDocxStyleConfig): MdToDocxStyleConfig => {
  const base = DEFAULT_STYLE_CONFIG;
  const defaults = normalizeAdvancedSettings(
    (raw as unknown as { advancedDefaults?: unknown }).advancedDefaults,
    base.advancedDefaults,
  );

  const normalizedBlock = (key: StyleSectionKey): MdToDocxStyleBlock => {
    const source = (raw as unknown as Partial<MdToDocxStyleConfig>)[key] ?? base[key];
    const withCompat = source as MdToDocxStyleBlock;

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
      zhFont: withCompat.zhFont ?? base[key].zhFont,
      enFont: withCompat.enFont ?? base[key].enFont,
      fontSizePt: Number.isFinite(Number(withCompat.fontSizePt))
        ? Number(withCompat.fontSizePt)
        : base[key].fontSizePt,
      lineSpacingMode: withCompat.lineSpacingMode === "fixed" ? "fixed" : "multiple",
      lineSpacingValue: Number.isFinite(Number(withCompat.lineSpacingValue))
        ? Number(withCompat.lineSpacingValue)
        : base[key].lineSpacingValue,
      align:
        withCompat.align === "center" ||
        withCompat.align === "right" ||
        withCompat.align === "justify"
          ? withCompat.align
          : "left",
      advancedOverride,
    };
  };

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
    advancedDefaults: defaults,
  };
};

const readStyleBlock = (key: StyleSectionKey, label: string): MdToDocxStyleBlock => {
  const zhFont = requireElement<HTMLSelectElement>(`#${styleInputId(key, "zhFont")}`).value.trim();
  const enFont = requireElement<HTMLSelectElement>(`#${styleInputId(key, "enFont")}`).value.trim();
  const fontSizePt = parsePositiveNumber(
    requireElement<HTMLInputElement>(`#${styleInputId(key, "fontSizePt")}`).value,
    `${label} 字号`,
  );
  const lineSpacingMode = requireElement<HTMLSelectElement>(
    `#${styleInputId(key, "lineSpacingMode")}`,
  ).value as LineSpacingMode;
  const lineSpacingValue = parsePositiveNumber(
    requireElement<HTMLInputElement>(`#${styleInputId(key, "lineSpacingValue")}`).value,
    `${label} 行距值`,
  );
  const align = requireElement<HTMLSelectElement>(`#${styleInputId(key, "align")}`).value as AlignMode;

  if (!zhFont) {
    throw new Error(`${label} 中文字体不能为空`);
  }
  if (!enFont) {
    throw new Error(`${label} 英文字体不能为空`);
  }

  const beforeMode = parseSpacingMode(
    requireElement<HTMLSelectElement>(`#${styleInputId(key, "beforeMode")}`).value,
  );
  const beforeValue = parseNonNegativeNumber(
    requireElement<HTMLInputElement>(`#${styleInputId(key, "beforeValue")}`).value,
    `${label} 段前值`,
  );
  const afterMode = parseSpacingMode(
    requireElement<HTMLSelectElement>(`#${styleInputId(key, "afterMode")}`).value,
  );
  const afterValue = parseNonNegativeNumber(
    requireElement<HTMLInputElement>(`#${styleInputId(key, "afterValue")}`).value,
    `${label} 段后值`,
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
      `${label} 首行缩进`,
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

const readStyleConfigFromForm = (): MdToDocxStyleConfig => {
  const values: Partial<MdToDocxStyleConfig> = {
    advancedDefaults: { ...DEFAULT_ADVANCED_SETTINGS },
  };

  for (const item of STYLE_SECTIONS) {
    values[item.key] = readStyleBlock(item.key, item.label);
  }
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

const applyStyleConfigToForm = (config: MdToDocxStyleConfig) => {
  const normalized = normalizeStyleConfig(config);

  for (const item of STYLE_SECTIONS) {
    applyStyleBlock(item.key, normalized[item.key]);
  }
};

pickInputButton.addEventListener("click", async () => {
  const selected = await open({
    multiple: false,
    filters: [{ name: "Word", extensions: ["doc", "docx"] }],
  });
  if (typeof selected === "string") {
    inputPathEl.value = selected;
    outputPathEl.value = defaultOutputFromDocInput(selected);
    setStatus("已选择输入文件", "success");
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
    setStatus("已设置输出路径", "success");
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
    setStatus(`图片将导出到 ${defaultImageDirFromMarkdownOutput(markdownPath)}`, "info");
  }
});

runButton.addEventListener("click", async () => {
  if (!inputPathEl.value) {
    setStatus("请先选择输入文件", "error");
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
  setStatus("转换中...", "info");

  try {
    const response = await invoke<ConversionResponse>("convert_docx", { request });

    const originalOutput = outputPathEl.value;
    if (response.outputPath) {
      outputPathEl.value = response.outputPath;
    }
    outputViewEl.textContent = response.markdown || "无输出内容";

    if (response.success) {
      const messages: string[] = ["导出成功", `文件：${response.outputPath || outputPathEl.value}`];
      if (response.outputPath && originalOutput && response.outputPath !== originalOutput) {
        messages.push("检测到重名，已自动改名保存");
      }
      if (extractImagesEl.checked) {
        const markdownPath = response.outputPath || outputPathEl.value;
        messages.push(`图片目录：${defaultImageDirFromMarkdownOutput(markdownPath)}`);
      }
      if (splitSectionsEl.checked && response.splitOutputDir && response.splitFileCount) {
        messages.push(`章节拆分：${response.splitFileCount} 个文件`);
      }
      setStatus(messages.join(" | "), "success");
    } else {
      const fallback = response.stderr || `转换失败（退出码 ${response.exitCode}）`;
      setStatus(fallback, "error");
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    setStatus(`调用失败：${message}`, "error");
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
    setStatus("已选择 Markdown 文件", "success");
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
    setStatus("已设置 DOCX 输出路径", "success");
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
  setStatus("已恢复默认样式", "success");
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
    setStatus("已加载样式预设", "success");
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    setStatus(`加载预设失败：${message}`, "error");
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
    setStatus("样式预设已保存", "success");
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    setStatus(`保存预设失败：${message}`, "error");
  }
});

runMdToDocxButton.addEventListener("click", async () => {
  if (!mdInputPathEl.value) {
    setStatus("请先选择 Markdown 输入文件", "error");
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
    setStatus(`样式配置无效：${message}`, "error");
    return;
  }

  const request: ConvertMdToDocxRequest = {
    inputPath: mdInputPathEl.value,
    outputPath: docxOutputPathEl.value,
    styleConfig,
  };

  runMdToDocxButton.disabled = true;
  setStatus("导出 DOCX 中...", "info");

  try {
    const originalOutput = docxOutputPathEl.value;
    const response = await invoke<ConvertMdToDocxResponse>("convert_md_to_docx", { request });
    if (response.outputPath) {
      docxOutputPathEl.value = response.outputPath;
    }

    const logs = [`输出文件：${response.outputPath}`];
    if (response.stdout) {
      logs.push("\n[stdout]\n" + response.stdout);
    }
    if (response.stderr) {
      logs.push("\n[stderr]\n" + response.stderr);
    }
    outputViewEl.textContent = logs.join("\n");

    if (response.success) {
      if (response.outputPath && originalOutput && response.outputPath !== originalOutput) {
        setStatus("DOCX 导出成功（检测到重名，已自动改名保存）", "success");
      } else {
        setStatus("DOCX 导出成功", "success");
      }
    } else {
      setStatus(response.stderr || `导出失败（退出码 ${response.exitCode}）`, "error");
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    setStatus(`调用失败：${message}`, "error");
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

setStylePanelExpanded(false);
switchDirection(currentDirection);
