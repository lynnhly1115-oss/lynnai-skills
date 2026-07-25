import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

function parseArgs(argv) {
  const result = {};
  for (let i = 2; i < argv.length; i += 1) {
    const key = argv[i];
    const value = argv[i + 1];
    if (!key.startsWith("--") || !value) throw new Error(`Invalid argument near ${key}`);
    result[key.slice(2)] = value;
    i += 1;
  }
  return result;
}

function validateAndNormalize(data) {
  if (!data || typeof data !== "object") throw new Error("Input JSON must be an object.");
  if (!data.projectTitle || typeof data.projectTitle !== "string") throw new Error("projectTitle is required.");
  if (!Array.isArray(data.episodes) || data.episodes.length === 0) throw new Error("episodes must be a non-empty array.");

  const allowedCategories = new Set(["人物", "场景", "道具"]);
  const allowedTypes = new Set(["基础", "变体"]);
  const categoryOrder = { 人物: 0, 场景: 1, 道具: 2 };
  const knownNames = new Set();

  const episodes = data.episodes.map((episodeBlock, episodeIndex) => {
    if (!episodeBlock.episode || !Array.isArray(episodeBlock.assets)) {
      throw new Error(`Episode ${episodeIndex + 1} must contain episode and assets.`);
    }
    const normalizedAssets = episodeBlock.assets.map((asset, assetIndex) => {
      for (const key of ["scene", "category", "name", "assetType", "description"]) {
        if (typeof asset[key] !== "string" || asset[key].trim() === "") {
          throw new Error(`${episodeBlock.episode} asset ${assetIndex + 1}: ${key} is required.`);
        }
      }
      if (!allowedCategories.has(asset.category)) throw new Error(`Invalid category: ${asset.category}`);
      if (!allowedTypes.has(asset.assetType)) throw new Error(`Invalid assetType: ${asset.assetType}`);
      const parent = typeof asset.parent === "string" ? asset.parent.trim() : "";
      if (asset.assetType === "变体" && !parent) throw new Error(`Variant ${asset.name} requires parent.`);
      if (asset.assetType === "基础" && parent) throw new Error(`Base asset ${asset.name} must not have parent.`);
      if (knownNames.has(asset.name)) throw new Error(`Duplicate asset name: ${asset.name}`);
      if (asset.assetType === "变体" && !knownNames.has(parent)) {
        throw new Error(`Variant ${asset.name} references missing or later parent: ${parent}`);
      }
      knownNames.add(asset.name);
      return {
        scene: asset.scene.trim(),
        category: asset.category,
        name: asset.name.trim(),
        assetType: asset.assetType,
        parent,
        description: asset.description.trim(),
        sourceOrder: assetIndex,
      };
    });
    normalizedAssets.sort((a, b) => {
      const categoryDiff = categoryOrder[a.category] - categoryOrder[b.category];
      if (categoryDiff !== 0) return categoryDiff;
      const sceneDiff = a.scene.localeCompare(b.scene, "zh-CN", { numeric: true });
      return sceneDiff !== 0 ? sceneDiff : a.sourceOrder - b.sourceOrder;
    });
    return { episode: episodeBlock.episode.trim(), assets: normalizedAssets };
  });
  return { projectTitle: data.projectTitle.trim(), episodes };
}

function workbookTitle(projectTitle) {
  return `${projectTitle} AI真人剧资产统计表`;
}

const args = parseArgs(process.argv);
if (!args.input || !args.output) {
  throw new Error("Usage: --input assets.json --output output.xlsx [--preview-dir previews]");
}

const raw = JSON.parse(await fs.readFile(args.input, "utf8"));
const data = validateAndNormalize(raw);
const workbook = Workbook.create();
const sheet = workbook.worksheets.add("资产统计表");
sheet.showGridLines = false;

sheet.mergeCells("A1:M1");
sheet.getRange("A1").values = [[workbookTitle(data.projectTitle)]];
sheet.mergeCells("A2:M2");
sheet.getRange("A2").values = [["统计规则：每项资产只在首次出现的集数中列出；后续复用不重复。服装、年龄、伪装、受伤、损毁与功能/特效变化按变体统计。"]];
sheet.mergeCells("A3:M3");
sheet.getRange("A3").values = [["图片位：人物和戏用道具使用图片1-3；场景使用图片1-6。场景陈设并入场景说明，纯背景群众不单独统计。"]];

const headers = ["集数", "场次", "类别", "资产名称", "基础/变体", "母资产", "资产说明", "图片1", "图片2", "图片3", "图片4", "图片5", "图片6"];
sheet.getRange("A4:M4").values = [headers];

const rows = [];
const meta = [];
for (const episodeBlock of data.episodes) {
  for (const asset of episodeBlock.assets) {
    rows.push([
      episodeBlock.episode,
      asset.scene,
      asset.category,
      asset.name,
      asset.assetType,
      asset.parent,
      asset.description,
      null,
      null,
      null,
      null,
      null,
      null,
    ]);
    meta.push({ kind: "asset", category: asset.category, assetType: asset.assetType });
  }
  rows.push([episodeBlock.episode, "", "小计", "本集小计", "", "", null, null, null, null, null, null, null]);
  meta.push({ kind: "subtotal" });
}
rows.push(["全剧", "", "总计", "累计唯一资产", "", "", null, null, null, null, null, null, null]);
meta.push({ kind: "grandTotal" });

const startRow = 5;
const endRow = startRow + rows.length - 1;
sheet.getRange(`A${startRow}:M${endRow}`).values = rows;

for (let i = 0; i < meta.length; i += 1) {
  const row = startRow + i;
  if (meta[i].kind === "subtotal") {
    sheet.getRange(`G${row}`).formulas = [[
      `="新增人物 "&COUNTIFS($A$${startRow}:$A$${endRow},A${row},$C$${startRow}:$C$${endRow},"人物")&"｜新增场景 "&COUNTIFS($A$${startRow}:$A$${endRow},A${row},$C$${startRow}:$C$${endRow},"场景")&"｜新增道具 "&COUNTIFS($A$${startRow}:$A$${endRow},A${row},$C$${startRow}:$C$${endRow},"道具")&"｜合计 "&(COUNTIFS($A$${startRow}:$A$${endRow},A${row},$C$${startRow}:$C$${endRow},"人物")+COUNTIFS($A$${startRow}:$A$${endRow},A${row},$C$${startRow}:$C$${endRow},"场景")+COUNTIFS($A$${startRow}:$A$${endRow},A${row},$C$${startRow}:$C$${endRow},"道具"))`,
    ]];
  } else if (meta[i].kind === "grandTotal") {
    sheet.getRange(`G${row}`).formulas = [[
      `="累计人物 "&COUNTIF($C$${startRow}:$C$${endRow},"人物")&"｜累计场景 "&COUNTIF($C$${startRow}:$C$${endRow},"场景")&"｜累计道具 "&COUNTIF($C$${startRow}:$C$${endRow},"道具")&"｜总计 "&(COUNTIF($C$${startRow}:$C$${endRow},"人物")+COUNTIF($C$${startRow}:$C$${endRow},"场景")+COUNTIF($C$${startRow}:$C$${endRow},"道具"))`,
    ]];
  }
}

const table = sheet.tables.add(`A4:M${endRow}`, true, "AssetInventoryTable");
table.style = "TableStyleLight1";
table.showFilterButton = true;

sheet.getRange("A1:M1").format = {
  fill: "#351719",
  font: { bold: true, color: "#F1D39B", size: 18 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  rowHeight: 34,
};
sheet.getRange("A2:M2").format = {
  fill: "#F3E4C2",
  font: { color: "#5A2426", size: 10 },
  horizontalAlignment: "left",
  verticalAlignment: "center",
  wrapText: true,
  rowHeight: 34,
};
sheet.getRange("A3:M3").format = {
  fill: "#FBF7EF",
  font: { color: "#4C4A46", size: 10 },
  horizontalAlignment: "left",
  verticalAlignment: "center",
  wrapText: true,
  rowHeight: 30,
};
sheet.getRange("A4:M4").format = {
  fill: "#7A2E2E",
  font: { bold: true, color: "#FFFFFF", size: 10 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
  borders: { preset: "all", style: "thin", color: "#D9C7B8" },
  rowHeight: 34,
};
sheet.getRange(`A${startRow}:M${endRow}`).format = {
  font: { color: "#2F2A27", size: 10 },
  verticalAlignment: "center",
  wrapText: true,
  borders: { preset: "all", style: "thin", color: "#D9D1C7" },
};
sheet.getRange(`A${startRow}:C${endRow}`).format.horizontalAlignment = "center";
sheet.getRange(`D${startRow}:G${endRow}`).format.horizontalAlignment = "left";
sheet.getRange(`E${startRow}:E${endRow}`).format.horizontalAlignment = "center";

const fills = { 人物: "#F7E8E3", 场景: "#F5E9CB", 道具: "#E9F0E8" };
for (let i = 0; i < meta.length; i += 1) {
  const row = startRow + i;
  const item = meta[i];
  if (item.kind === "asset") {
    sheet.getRange(`A${row}:G${row}`).format.fill = fills[item.category];
    sheet.getRange(`H${row}:M${row}`).format.fill = "#FFFFFF";
    if (item.category !== "场景") sheet.getRange(`K${row}:M${row}`).format.fill = "#ECEBE8";
    if (item.assetType === "变体") {
      sheet.getRange(`E${row}:F${row}`).format.fill = "#EED9A7";
      sheet.getRange(`E${row}:F${row}`).format.font = { bold: true, color: "#6B3C1F" };
      sheet.getRange(`E${row}`).format.horizontalAlignment = "center";
    }
    sheet.getRange(`A${row}:M${row}`).format.rowHeight = 90;
  } else if (item.kind === "subtotal") {
    sheet.getRange(`A${row}:M${row}`).format = {
      fill: "#E3C996",
      font: { bold: true, color: "#4B2A20" },
      verticalAlignment: "center",
      wrapText: true,
      borders: { preset: "all", style: "thin", color: "#B99762" },
      rowHeight: 28,
    };
    sheet.getRange(`A${row}:F${row}`).format.horizontalAlignment = "center";
  } else {
    sheet.getRange(`A${row}:M${row}`).format = {
      fill: "#351719",
      font: { bold: true, color: "#F1D39B" },
      verticalAlignment: "center",
      wrapText: true,
      borders: { preset: "all", style: "medium", color: "#351719" },
      rowHeight: 32,
    };
    sheet.getRange(`A${row}:F${row}`).format.horizontalAlignment = "center";
  }
}

const widths = { A: 9, B: 9, C: 9, D: 32, E: 12, F: 32, G: 48, H: 20, I: 20, J: 20, K: 20, L: 20, M: 20 };
for (const [column, width] of Object.entries(widths)) {
  sheet.getRange(`${column}1:${column}${endRow}`).format.columnWidth = width;
}
sheet.freezePanes.freezeRows(4);
sheet.freezePanes.freezeColumns(7);

await fs.mkdir(path.dirname(args.output), { recursive: true });
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(args.output);

if (args["preview-dir"]) {
  await fs.mkdir(args["preview-dir"], { recursive: true });
  const segmentSize = 35;
  for (let from = 1, segment = 1; from <= endRow; from += segmentSize, segment += 1) {
    const to = Math.min(endRow, from + segmentSize - 1);
    const preview = await workbook.render({ sheetName: "资产统计表", range: `A${from}:M${to}`, scale: 1, format: "png" });
    await fs.writeFile(path.join(args["preview-dir"], `preview-${String(segment).padStart(2, "0")}.png`), new Uint8Array(await preview.arrayBuffer()));
  }
}

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});
console.log(JSON.stringify({ output: args.output, rows: endRow, formulaErrorScan: errors.ndjson }));
