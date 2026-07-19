import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, "..");
const workbook = Workbook.create();
const sheet = workbook.worksheets.add("项目计划");
sheet.showGridLines = false;

sheet.getRange("A1:G1").merge();
sheet.getRange("A1").values = [["市场活动排期项目"]];
sheet.getRange("A1:G1").format = {
  fill: "#172554",
  font: { color: "#FFFFFF", bold: true, size: 18 },
  verticalAlignment: "center",
};
sheet.getRange("A1:G1").format.rowHeight = 34;

sheet.getRange("A2:G2").merge();
sheet.getRange("A2").values = [[
  "公开脱敏示例｜每行一条任务事实；日期为真实日期单元格；空卡片会触发默认分组 warning。",
]];
sheet.getRange("A2:G2").format = {
  fill: "#EFF6FF",
  font: { color: "#475569", italic: true, size: 10 },
  wrapText: true,
  verticalAlignment: "center",
};
sheet.getRange("A2:G2").format.rowHeight = 30;

const headers = [["模块", "卡片", "任务", "负责人", "开始日期", "截止日期", "备注"]];
sheet.getRange("A4:G4").values = headers;
sheet.getRange("A4:G4").format = {
  fill: "#2563EB",
  font: { color: "#FFFFFF", bold: true },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: {
    top: { style: "thin", color: "#1D4ED8" },
    bottom: { style: "thin", color: "#1D4ED8" },
    left: { style: "thin", color: "#1D4ED8" },
    right: { style: "thin", color: "#1D4ED8" },
  },
};
sheet.getRange("A4:G4").format.rowHeight = 26;

const rows = [
  ["筹备", "内容准备", "确认活动主题", "产品负责人", new Date("2026-07-01T00:00:00Z"), new Date("2026-07-02T00:00:00Z"), "与业务目标保持一致"],
  ["筹备", "内容准备", "完成活动主文案", "内容负责人", new Date("2026-07-02T00:00:00Z"), new Date("2026-07-04T00:00:00Z"), "首轮评审前完成"],
  ["筹备", "渠道准备", "整理渠道清单", "运营负责人", new Date("2026-07-02T00:00:00Z"), new Date("2026-07-05T00:00:00Z"), ""],
  ["设计", "视觉物料", "输出主视觉方案", "设计负责人", new Date("2026-07-03T00:00:00Z"), new Date("2026-07-06T00:00:00Z"), "包含桌面与移动端"],
  ["设计", "视觉物料", "适配渠道尺寸", "设计负责人", new Date("2026-07-06T00:00:00Z"), new Date("2026-07-08T00:00:00Z"), ""],
  ["开发", "页面搭建", "完成活动落地页", "开发负责人", new Date("2026-07-05T00:00:00Z"), new Date("2026-07-10T00:00:00Z"), "完成埋点自测"],
  ["上线", "发布检查", "核对渠道与链接", "运营负责人", new Date("2026-07-11T00:00:00Z"), new Date("2026-07-12T00:00:00Z"), "上线前逐项确认"],
  ["上线", "", "汇总活动数据", "数据负责人", new Date("2026-07-13T00:00:00Z"), new Date("2026-07-15T00:00:00Z"), "示例：空卡片触发默认分组"],
];
sheet.getRange("A5:G12").values = rows;
sheet.getRange("A5:G12").format = {
  font: { color: "#24324A", size: 10 },
  verticalAlignment: "center",
  borders: {
    top: { style: "hair", color: "#DCE3EF" },
    bottom: { style: "hair", color: "#DCE3EF" },
    left: { style: "hair", color: "#DCE3EF" },
    right: { style: "hair", color: "#DCE3EF" },
  },
};
sheet.getRange("A5:G12").format.rowHeight = 24;
sheet.getRange("E5:F12").setNumberFormat("yyyy-mm-dd");
sheet.getRange("A6:G6").format.fill = "#F8FAFC";
sheet.getRange("A8:G8").format.fill = "#F8FAFC";
sheet.getRange("A10:G10").format.fill = "#F8FAFC";
sheet.getRange("A12:G12").format.fill = "#FFF7ED";

sheet.getRange("A14:G14").merge();
sheet.getRange("A14").values = [[
  "验证命令：python src/excel_to_project_demo.py demo/sample_project_plan.xlsx --output demo/generated_project.json",
]];
sheet.getRange("A14:G14").format = {
  fill: "#ECFDF5",
  font: { color: "#047857", size: 10 },
  wrapText: true,
};

sheet.getRange("A1:A14").format.columnWidth = 13;
sheet.getRange("B1:B14").format.columnWidth = 16;
sheet.getRange("C1:C14").format.columnWidth = 25;
sheet.getRange("D1:D14").format.columnWidth = 16;
sheet.getRange("E1:F14").format.columnWidth = 14;
sheet.getRange("G1:G14").format.columnWidth = 26;
sheet.freezePanes.freezeRows(4);

sheet.tables.add("A4:G12", true, "ProjectPlanTable");

const inspection = await workbook.inspect({
  kind: "sheet,table,region",
  sheetId: "项目计划",
  range: "A1:G14",
  maxChars: 4500,
  tableMaxRows: 14,
  tableMaxCols: 7,
});
console.log(inspection.ndjson);

const demoDir = path.join(root, "demo");
const previewDir = path.join(root, "output", "spreadsheet");
await fs.mkdir(demoDir, { recursive: true });
await fs.mkdir(previewDir, { recursive: true });

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(path.join(demoDir, "sample_project_plan.xlsx"));

const preview = await workbook.render({
  sheetName: "项目计划",
  range: "A1:G14",
  scale: 1.4,
  format: "png",
});
await fs.writeFile(
  path.join(previewDir, "sample_project_plan.png"),
  new Uint8Array(await preview.arrayBuffer()),
);

console.log("Created demo/sample_project_plan.xlsx");
console.log("Rendered output/spreadsheet/sample_project_plan.png");
