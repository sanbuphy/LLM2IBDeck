#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..");
const PRESENTATIONS_SKILL = "/Users/sanbu/.codex/plugins/cache/openai-primary-runtime/presentations/26.601.10930/skills/presentations";
const BUILD_SCRIPT = path.join(PRESENTATIONS_SKILL, "scripts", "build_artifact_deck.mjs");
const NODE = "/Users/sanbu/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node";
const PYTHON = "/Users/sanbu/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3";

function parseArgs(argv) {
  const args = {};
  for (let i = 0; i < argv.length; i += 1) {
    const key = argv[i];
    if (!key.startsWith("--")) throw new Error(`Unexpected arg: ${key}`);
    const value = argv[i + 1];
    if (!value || value.startsWith("--")) {
      args[key.slice(2)] = true;
    } else {
      args[key.slice(2)] = value;
      i += 1;
    }
  }
  return args;
}

const THEMES = {
  "mckinsey-inspired": {
    theme: "mckinsey-inspired",
    label: "McKinsey-inspired",
    primary: "#003A70",
    secondary: "#00A3E0",
    accent: "#1F77B4",
    soft: "#EEF5FA",
    ink: "#202124",
    muted: "#5F6368",
    footer: "BankerDeck | public-reference calibrated | McKinsey-inspired",
    title: "Digital infrastructure investment priorities",
    subtitle: "Executive fact base and implications",
    claim: "Power-secure platforms are best positioned to capture AI-driven demand",
    panel: "Power availability is emerging as the gating constraint.",
    testsTitle: "Four tests separate scalable platforms from one-off development risk",
    chartTitle: "Capital intensity rises as facilities shift toward AI-ready specifications",
    source: "Source: Public industry sources; BankerDeck synthesis.",
    pageStyle: "consulting",
    points: [
      "Prioritize contracted capacity and utility relationships",
      "Underwrite higher capex only with demand visibility",
      "Separate platform scarcity from project-level risk",
    ],
    tests: [
      ["Power-secure", "Verified grid capacity and utility path", "Constrains near-term supply"],
      ["Customer-backed", "Anchor demand and pre-lease evidence", "Reduces absorption risk"],
      ["Capex-disciplined", "Repeatable design and cost controls", "Protects return profile"],
      ["Exit-ready", "Platform scarcity and buyer universe", "Supports terminal value"],
    ],
  },
  "goldman-inspired": {
    theme: "goldman-inspired",
    label: "Goldman-inspired",
    primary: "#0B1F3A",
    secondary: "#C9A227",
    accent: "#335C81",
    soft: "#E9F1FA",
    ink: "#172033",
    muted: "#637083",
    footer: "BankerDeck | public-reference calibrated | Goldman-inspired",
    title: "Digital infrastructure financing perspectives",
    subtitle: "Capital markets diligence materials",
    claim: "Risk / reward increasingly depends on contracted power access and capex control",
    panel: "Sponsors should distinguish scarcity premium from development execution risk.",
    testsTitle: "Platform diligence focuses on power, customer backing, capex discipline and exit depth",
    chartTitle: "Higher AI-ready specification raises capital intensity and financing sensitivity",
    source: "Source: Company filings, public market data and BankerDeck analysis.",
    pageStyle: "capital-markets",
    points: [
      "Frame valuation around power-secure capacity and delivery milestones",
      "Stress-test leverage capacity under delayed energization and cost inflation",
      "Position buyer universe around contracted growth and platform scarcity",
    ],
    tests: [
      ["Power access", "Utility path, queue position, contracted capacity", "Key gating item"],
      ["Customer demand", "Anchor leases, take-or-pay terms, renewal evidence", "Revenue visibility"],
      ["Capex control", "Repeatable design, procurement, contingency", "Return protection"],
      ["Exit options", "Strategic, infra, REIT and sponsor buyer depth", "Terminal support"],
    ],
  },
  "cicc-inspired": {
    theme: "cicc-inspired",
    label: "CICC-inspired",
    primary: "#B20D1E",
    secondary: "#B88900",
    accent: "#8A1E2D",
    soft: "#F5F1EA",
    ink: "#1F1F1F",
    muted: "#5B5148",
    footer: "BankerDeck | 公开资料校准 | CICC-inspired",
    title: "算力基础设施投资主线",
    subtitle: "行业研究框架与投资含义",
    claim: "电力资源、客户锁定与资本开支管控共同决定平台价值弹性",
    panel: "核心观点：AI 需求提升供给约束，具备电力确定性的项目更具稀缺性。",
    testsTitle: "筛选优质平台需重点关注四类验证指标",
    chartTitle: "AI-ready 规格提升资本开支强度，回报测算需纳入敏感性分析",
    source: "资料来源：公开资料，BankerDeck 整理。注：指数为示意。",
    pageStyle: "research-cn",
    points: [
      "优先关注电力指标明确、并网路径清晰的项目",
      "结合预租约和锚定客户验证需求确定性",
      "通过成本、工期和融资敏感性评估回报韧性",
    ],
    tests: [
      ["电力确定性", "并网容量、能耗指标、审批进度", "决定供给释放"],
      ["客户锁定", "锚定客户、预租约、合同期限", "提升收入可见度"],
      ["资本开支", "单位造价、采购周期、预备费", "影响 IRR 弹性"],
      ["退出空间", "产业买方、基础设施基金、REITs", "支撑估值中枢"],
    ],
  },
};

function createSlideModules(config) {
  const cfg = JSON.stringify(config, null, 2);
  return [
    {
      file: "slide-01.mjs",
      code: `
const cfg = ${cfg};

export async function slide01(presentation, ctx) {
  const slide = presentation.slides.add();
  ctx.addShape(slide, { left: 0, top: 0, width: 1280, height: 720, fill: "#FFFFFF", line: ctx.line() });
  if (cfg.pageStyle === "capital-markets") {
    ctx.addShape(slide, { left: 0, top: 0, width: 1280, height: 128, fill: cfg.primary, line: ctx.line() });
    ctx.addShape(slide, { left: 0, top: 128, width: 1280, height: 5, fill: cfg.secondary, line: ctx.line() });
    ctx.addText(slide, { left: 72, top: 44, width: 860, height: 42, text: "BANKERDECK CAPITAL MARKETS", fontSize: 13, bold: true, color: "#FFFFFF", typeface: "Arial" });
    ctx.addText(slide, { left: 74, top: 212, width: 820, height: 104, text: cfg.title, fontSize: 33, bold: true, color: cfg.primary, typeface: "Arial" });
    ctx.addText(slide, { left: 76, top: 342, width: 620, height: 36, text: cfg.subtitle, fontSize: 16, color: cfg.muted, typeface: "Arial" });
    ctx.addShape(slide, { left: 76, top: 432, width: 920, height: 1.5, fill: "#B7C1CF", line: ctx.line() });
    ctx.addText(slide, { left: 76, top: 642, width: 640, height: 18, text: cfg.footer, fontSize: 8.5, color: cfg.muted, typeface: "Arial" });
  } else if (cfg.pageStyle === "research-cn") {
    ctx.addShape(slide, { left: 0, top: 0, width: 1280, height: 96, fill: cfg.primary, line: ctx.line() });
    ctx.addShape(slide, { left: 0, top: 96, width: 1280, height: 7, fill: cfg.secondary, line: ctx.line() });
    ctx.addText(slide, { left: 74, top: 38, width: 760, height: 32, text: "BankerDeck 行业研究", fontSize: 15, bold: true, color: "#FFFFFF", typeface: "Arial" });
    ctx.addText(slide, { left: 72, top: 190, width: 760, height: 88, text: cfg.title, fontSize: 34, bold: true, color: cfg.ink, typeface: "Arial" });
    ctx.addShape(slide, { left: 72, top: 308, width: 560, height: 2, fill: cfg.secondary, line: ctx.line() });
    ctx.addText(slide, { left: 72, top: 332, width: 720, height: 38, text: cfg.subtitle, fontSize: 17, color: cfg.muted, typeface: "Arial" });
    ctx.addShape(slide, { left: 850, top: 188, width: 260, height: 156, fill: cfg.soft, line: { fill: "#D9CDB8", width: 1, style: "solid" } });
    ctx.addText(slide, { left: 880, top: 224, width: 202, height: 60, text: "核心观点\\n资料来源明确\\n表格优先", fontSize: 16, bold: true, color: cfg.primary, typeface: "Arial", align: "center" });
    ctx.addText(slide, { left: 72, top: 642, width: 780, height: 18, text: cfg.footer, fontSize: 8.5, color: cfg.muted, typeface: "Arial" });
  } else {
    ctx.addShape(slide, { left: 72, top: 132, width: 8, height: 320, fill: cfg.secondary, line: ctx.line() });
    ctx.addText(slide, { left: 112, top: 154, width: 760, height: 120, text: cfg.title, fontSize: 38, bold: true, color: cfg.primary, typeface: "Arial" });
    ctx.addText(slide, { left: 114, top: 288, width: 680, height: 40, text: cfg.subtitle, fontSize: 18, color: cfg.muted, typeface: "Arial" });
    ctx.addShape(slide, { left: 114, top: 372, width: 260, height: 4, fill: cfg.secondary, line: ctx.line() });
    ctx.addText(slide, { left: 114, top: 640, width: 560, height: 20, text: cfg.footer, fontSize: 9, color: "#6B7280", typeface: "Arial" });
  }
  return slide;
}
`,
    },
    {
      file: "slide-02.mjs",
      code: `
const cfg = ${cfg};

export async function slide02(presentation, ctx) {
  const slide = presentation.slides.add();
  ctx.addShape(slide, { left: 0, top: 0, width: 1280, height: 720, fill: "#FFFFFF", line: ctx.line() });
  if (cfg.pageStyle === "capital-markets") {
    ctx.addShape(slide, { left: 0, top: 0, width: 1280, height: 54, fill: cfg.primary, line: ctx.line() });
    ctx.addText(slide, { left: 66, top: 18, width: 1020, height: 22, text: cfg.claim, fontSize: 20, bold: true, color: "#FFFFFF", typeface: "Arial" });
    ctx.addShape(slide, { left: 66, top: 92, width: 1050, height: 1.5, fill: "#C7CED8", line: ctx.line() });
  } else if (cfg.pageStyle === "research-cn") {
    ctx.addShape(slide, { left: 0, top: 0, width: 1280, height: 42, fill: cfg.primary, line: ctx.line() });
    ctx.addText(slide, { left: 64, top: 64, width: 1020, height: 50, text: cfg.claim, fontSize: 22, bold: true, color: cfg.ink, typeface: "Arial" });
    ctx.addShape(slide, { left: 64, top: 118, width: 1090, height: 3, fill: cfg.secondary, line: ctx.line() });
  } else {
    ctx.addShape(slide, { left: 64, top: 36, width: 8, height: 58, fill: cfg.secondary, line: ctx.line() });
    ctx.addText(slide, { left: 92, top: 36, width: 1020, height: 62, text: cfg.claim, fontSize: 25, bold: true, color: cfg.primary, typeface: "Arial" });
  }
  ctx.addShape(slide, { left: 92, top: 138, width: 360, height: 390, fill: cfg.pageStyle === "research-cn" ? cfg.soft : "#FFFFFF", line: { fill: "#CFD7E2", width: 1, style: "solid" } });
  ctx.addText(slide, { left: 122, top: 174, width: 300, height: 118, text: cfg.panel, fontSize: cfg.pageStyle === "research-cn" ? 18 : 21, bold: true, color: cfg.primary, typeface: "Arial" });
  const points = cfg.points;
  points.forEach((point, i) => {
    const y = 146 + i * 116;
    ctx.addShape(slide, { left: 520, top: y, width: 34, height: 34, fill: i === 0 ? cfg.secondary : cfg.primary, line: ctx.line() });
    ctx.addText(slide, { left: 520, top: y + 4, width: 34, height: 24, text: String(i + 1), fontSize: 14, bold: true, color: "#FFFFFF", align: "center", typeface: "Arial" });
    ctx.addText(slide, { left: 578, top: y - 4, width: 540, height: 56, text: point, fontSize: cfg.pageStyle === "research-cn" ? 15 : 17, color: cfg.ink, typeface: "Arial" });
  });
  ctx.addText(slide, { left: 92, top: 632, width: 850, height: 18, text: cfg.source, fontSize: 9, color: "#6B7280", typeface: "Arial" });
  return slide;
}
`,
    },
    {
      file: "slide-03.mjs",
      code: `
const cfg = ${cfg};

export async function slide03(presentation, ctx) {
  const slide = presentation.slides.add();
  ctx.addShape(slide, { left: 0, top: 0, width: 1280, height: 720, fill: "#FFFFFF", line: ctx.line() });
  if (cfg.pageStyle === "capital-markets") {
    ctx.addShape(slide, { left: 0, top: 0, width: 1280, height: 48, fill: cfg.primary, line: ctx.line() });
    ctx.addText(slide, { left: 64, top: 68, width: 1040, height: 40, text: cfg.testsTitle, fontSize: 21, bold: true, color: cfg.primary, typeface: "Arial" });
    ctx.addShape(slide, { left: 64, top: 116, width: 1090, height: 2, fill: cfg.secondary, line: ctx.line() });
  } else {
    ctx.addText(slide, { left: 64, top: 34, width: 1020, height: 58, text: cfg.testsTitle, fontSize: cfg.pageStyle === "research-cn" ? 22 : 24, bold: true, color: cfg.primary, typeface: "Arial" });
    ctx.addShape(slide, { left: 64, top: 104, width: 1090, height: 2, fill: cfg.pageStyle === "research-cn" ? cfg.secondary : "#CFD7E2", line: ctx.line() });
  }
  const cards = cfg.tests;
  cards.forEach((card, i) => {
    const x = 70 + (i % 2) * 560;
    const y = 150 + Math.floor(i / 2) * 205;
    const fill = cfg.pageStyle === "research-cn" ? (i % 2 === 0 ? "#FFFFFF" : cfg.soft) : "#FFFFFF";
    ctx.addShape(slide, { left: x, top: y, width: 500, height: 150, fill, line: { fill: "#CFD7E2", width: 1, style: "solid" } });
    ctx.addShape(slide, { left: x, top: y, width: 7, height: 150, fill: i === 0 ? cfg.secondary : cfg.primary, line: ctx.line() });
    ctx.addText(slide, { left: x + 28, top: y + 22, width: 420, height: 26, text: card[0], fontSize: 18, bold: true, color: cfg.primary, typeface: "Arial" });
    ctx.addText(slide, { left: x + 28, top: y + 58, width: 420, height: 34, text: card[1], fontSize: 13, color: cfg.ink, typeface: "Arial" });
    ctx.addText(slide, { left: x + 28, top: y + 104, width: 420, height: 24, text: card[2], fontSize: 11, color: cfg.muted, typeface: "Arial" });
  });
  ctx.addText(slide, { left: 64, top: 632, width: 860, height: 18, text: cfg.source, fontSize: 9, color: "#6B7280", typeface: "Arial" });
  return slide;
}
`,
    },
    {
      file: "slide-04.mjs",
      code: `
const cfg = ${cfg};

export async function slide04(presentation, ctx) {
  const slide = presentation.slides.add();
  ctx.addShape(slide, { left: 0, top: 0, width: 1280, height: 720, fill: "#FFFFFF", line: ctx.line() });
  ctx.addText(slide, { left: 64, top: 34, width: 1020, height: 58, text: cfg.chartTitle, fontSize: cfg.pageStyle === "research-cn" ? 21 : 24, bold: true, color: cfg.primary, typeface: "Arial" });
  ctx.addShape(slide, { left: 64, top: 104, width: 1090, height: 2, fill: cfg.pageStyle === "research-cn" ? cfg.secondary : "#CFD7E2", line: ctx.line() });
  ctx.addShape(slide, { left: 86, top: 150, width: 270, height: 390, fill: cfg.pageStyle === "research-cn" ? cfg.soft : "#FFFFFF", line: { fill: "#CFD7E2", width: 1, style: "solid" } });
  ctx.addText(slide, { left: 112, top: 184, width: 218, height: 70, text: "260", fontSize: 52, bold: true, color: cfg.primary, align: "center", typeface: "Arial" });
  ctx.addText(slide, { left: 112, top: 256, width: 218, height: 44, text: cfg.pageStyle === "research-cn" ? "AI-ready 规格下电力密度指数" : "power density index at AI-ready specification", fontSize: 13, color: cfg.muted, align: "center", typeface: "Arial" });
  const cats = ["Legacy colo", "Hyperscale", "AI-ready"];
  const capex = [100, 135, 180];
  const power = [100, 160, 260];
  cats.forEach((cat, i) => {
    const baseX = 470 + i * 180;
    const yBase = 540;
    const h1 = capex[i] * 1.15;
    const h2 = power[i] * 1.15;
    ctx.addShape(slide, { left: baseX, top: yBase - h1, width: 54, height: h1, fill: cfg.primary, line: ctx.line() });
    ctx.addShape(slide, { left: baseX + 62, top: yBase - h2, width: 54, height: h2, fill: cfg.secondary, line: ctx.line() });
    ctx.addText(slide, { left: baseX - 20, top: 554, width: 150, height: 26, text: cfg.pageStyle === "research-cn" ? ["传统机柜", "云计算", "AI-ready"][i] : cat, fontSize: 10, color: cfg.muted, align: "center", typeface: "Arial" });
  });
  ctx.addText(slide, { left: 468, top: 144, width: 500, height: 20, text: cfg.pageStyle === "research-cn" ? "指数化示意" : "Indexed illustrative view", fontSize: 11, color: cfg.muted, typeface: "Arial" });
  ctx.addText(slide, { left: 64, top: 632, width: 860, height: 18, text: cfg.source, fontSize: 9, color: "#6B7280", typeface: "Arial" });
  return slide;
}
`,
    },
  ];
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const requestedTheme = args.theme || "all";
  const themeNames = requestedTheme === "all" ? Object.keys(THEMES) : [requestedTheme];
  const unknown = themeNames.filter((themeName) => !THEMES[themeName]);
  if (unknown.length) {
    throw new Error(`Unknown theme: ${unknown.join(", ")}. Expected one of: ${Object.keys(THEMES).join(", ")}, all`);
  }

  const rootWorkspace = path.resolve(args.workspace || path.join(ROOT, "temp", "real-deck"));
  await fs.rm(rootWorkspace, { recursive: true, force: true });
  const manifests = [];
  for (const themeName of themeNames) {
    const config = THEMES[themeName];
    const workspace = path.join(rootWorkspace, themeName);
    const slidesDir = path.join(workspace, "slides");
    const outputDir = path.join(workspace, "output");
    const previewDir = path.join(workspace, "preview");
    const layoutDir = path.join(workspace, "layout");
    const contactSheet = path.join(workspace, "contact-sheet.png");
    const out = path.resolve(args.output && themeNames.length === 1
      ? args.output
      : path.join(outputDir, `${themeName}-real-demo.pptx`));
    const slideModules = createSlideModules(config);

    await fs.rm(workspace, { recursive: true, force: true });
    await fs.mkdir(slidesDir, { recursive: true });
    for (const mod of slideModules) {
      await fs.writeFile(path.join(slidesDir, mod.file), mod.code.trimStart(), "utf8");
    }

    const result = spawnSync(
      NODE,
      [
        BUILD_SCRIPT,
        "--workspace", workspace,
        "--slides-dir", slidesDir,
        "--out", out,
        "--preview-dir", previewDir,
        "--layout-dir", layoutDir,
        "--contact-sheet", contactSheet,
        "--slide-count", String(slideModules.length),
        "--scale", "1",
      ],
      {
        cwd: ROOT,
        encoding: "utf8",
        env: { ...process.env, PYTHON },
      },
    );
    if (result.status !== 0) {
      console.error(result.stdout);
      console.error(result.stderr);
      process.exit(result.status || 1);
    }
    manifests.push(JSON.parse(result.stdout));
  }
  console.log(JSON.stringify({ generated: manifests }, null, 2));
}

main().catch((error) => {
  console.error(error.stack || error.message || String(error));
  process.exit(1);
});
