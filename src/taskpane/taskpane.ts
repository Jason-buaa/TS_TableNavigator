/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("setup").onclick = setup;
    document.getElementById("apply-color-scale-format").onclick = applyColorScaleFormat;
    document.getElementById("list-conditional-formats").onclick = listConditionalFormats;
    document.getElementById("apply-preset-format").onclick = applyPresetFormat;
    document.getElementById("apply-databar-format").onclick = applyDataBarFormat;
    document.getElementById("apply-icon-set-format").onclick = applyIconSetFormat;
    document.getElementById("apply-text-format").onclick = applyTextFormat;
    document.getElementById("apply-cell-value-format").onclick = applyCellValueFormat;
    document.getElementById("apply-top-bottom-format").onclick = applyTopBottomFormat;
    document.getElementById("apply-custom-format").onclick = applyCustomFormat;
    document.getElementById("clear-all-conditional-formats").onclick = clearAllConditionalFormats;
    document.getElementById("save-all-conditional-formats").onclick = saveConditionalFormats;
  }
});
let savedConditionalFormats = [];
async function setup() {
  await Excel.run(async (context) => {
    context.workbook.worksheets.getItemOrNullObject("Sample").delete();
    const sheet = context.workbook.worksheets.add("Sample");

    queueCommandsToCreateTemperatureTable(sheet);
    queueCommandsToCreateSalesTable(sheet);
    queueCommandsToCreateProjectTable(sheet);
    queueCommandsToCreateProfitLossTable(sheet);

    let format = sheet.getRange().format;
    format.autofitColumns();
    format.autofitRows();

    sheet.activate();
    await context.sync();
  });
}

function queueCommandsToCreateTemperatureTable(sheet: Excel.Worksheet) {
  let temperatureTable = sheet.tables.add("A1:M1", true);
  temperatureTable.name = "TemperatureTable";
  temperatureTable.getHeaderRowRange().values = [
    ["Category", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
  ];
  temperatureTable.rows.add(null, [
    ["Avg High", 40, 38, 44, 45, 51, 56, 67, 72, 79, 59, 45, 41],
    ["Avg Low", 34, 33, 38, 41, 45, 48, 51, 55, 54, 45, 41, 38],
    ["Record High", 61, 69, 79, 83, 95, 97, 100, 101, 94, 87, 72, 66],
    ["Record Low", 0, 2, 9, 24, 28, 32, 36, 39, 35, 21, 12, 4],
  ]);
}

function queueCommandsToCreateSalesTable(sheet: Excel.Worksheet) {
  let salesTable = sheet.tables.add("A7:E7", true);
  salesTable.name = "SalesTable";
  salesTable.getHeaderRowRange().values = [["Sales Team", "Qtr1", "Qtr2", "Qtr3", "Qtr4"]];
  salesTable.rows.add(null, [
    ["Asian Team 1", 500, 700, 654, 234],
    ["Asian Team 2", 400, 323, 276, 345],
    ["Asian Team 3", 1200, 876, 845, 456],
    ["Euro Team 1", 600, 500, 854, 567],
    ["Euro Team 2", 5001, 2232, 4763, 678],
    ["Euro Team 3", 130, 776, 104, 789],
  ]);
}

function queueCommandsToCreateProjectTable(sheet: Excel.Worksheet) {
  let projectTable = sheet.tables.add("A15:D15", true);
  projectTable.name = "ProjectTable";
  projectTable.getHeaderRowRange().values = [["Project", "Alpha", "Beta", "Ship"]];
  projectTable.rows.add(null, [
    ["Project 1", "Complete", "Ongoing", "On Schedule"],
    ["Project 2", "Complete", "Complete", "On Schedule"],
    ["Project 3", "Ongoing", "Not Started", "Delayed"],
  ]);
}

function queueCommandsToCreateProfitLossTable(sheet: Excel.Worksheet) {
  let profitLossTable = sheet.tables.add("A20:E20", true);
  profitLossTable.name = "ProfitLossTable";
  profitLossTable.getHeaderRowRange().values = [["Company", "2013", "2014", "2015", "2016"]];
  profitLossTable.rows.add(null, [
    ["Contoso", 256.0, -55.31, 68.9, -82.13],
    ["Fabrikam", 454.0, 75.29, -88.88, 781.87],
    ["Northwind", -858.21, 35.33, 49.01, 112.68],
  ]);
  profitLossTable.getDataBodyRange().numberFormat = [["$#,##0.00"]];
}

async function listConditionalFormats() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const worksheetRange = sheet.getRange();
    worksheetRange.conditionalFormats.load("type");

    await context.sync();

    let cfRangePairs: { cf: Excel.ConditionalFormat; range: Excel.Range }[] = [];
    worksheetRange.conditionalFormats.items.forEach((item) => {
      const cfRange = item.getRange();
      cfRange.load("address");
      cfRangePairs.push({
        cf: item,
        range: cfRange,
      });
    });

    await context.sync();

    if (cfRangePairs.length > 0) {
      cfRangePairs.forEach(async (pair) => {
        console.log("条件格式类型:", pair.cf.type);
        if (pair.cf.type == "Custom") {
          pair.cf.custom.rule.load("formula");
          pair.cf.custom.format.font.load("color");
          await context.sync();
          console.log(pair.cf.custom.rule.formula);
          console.log(pair.cf.custom.format.font.color);
        }
        console.log("应用范围:", pair.range.address);
      });
    } else {
      console.log("未应用任何条件格式。");
    }
  });
}

async function applyColorScaleFormat() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("B2:M5");
    const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
    const criteria = {
      minimum: { formula: null, type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: "blue" },
      midpoint: { formula: "50", type: Excel.ConditionalFormatColorCriterionType.percent, color: "yellow" },
      maximum: { formula: null, type: Excel.ConditionalFormatColorCriterionType.highestValue, color: "red" },
    };
    conditionalFormat.colorScale.criteria = criteria;

    await context.sync();
  });
}

async function applyPresetFormat() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("B2:M5");
    const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.presetCriteria);
    conditionalFormat.preset.format.font.color = "white";
    conditionalFormat.preset.rule = { criterion: Excel.ConditionalFormatPresetCriterion.oneStdDevAboveAverage };

    await context.sync();
  });
}

async function applyDataBarFormat() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("B8:E13");
    const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
    conditionalFormat.dataBar.barDirection = Excel.ConditionalDataBarDirection.leftToRight;

    await context.sync();
  });
}

async function applyIconSetFormat() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("B8:E13");
    const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
    const iconSetCF = conditionalFormat.iconSet;
    iconSetCF.style = Excel.IconSet.threeTriangles;

    /*
          The iconSetCF.criteria array is automatically prepopulated with
          criterion elements whose properties have been given default settings.
          You can't write to each property of a criterion directly. Instead,
          replace the whole criteria object.

          With a "three*" icon set style, such as "threeTriangles", the third
          element in the criteria array (criteria[2]) defines the "top" icon;
          e.g., a green triangle. The second (criteria[1]) defines the "middle"
          icon. The first (criteria[0]) defines the "low" icon, but it
          can often be left empty as the following object shows, because every
          cell that does not match the other two criteria always gets the low
          icon.            
      */
    iconSetCF.criteria = [
      {} as any,
      {
        type: Excel.ConditionalFormatIconRuleType.number,
        operator: Excel.ConditionalIconCriterionOperator.greaterThanOrEqual,
        formula: "=700",
      },
      {
        type: Excel.ConditionalFormatIconRuleType.number,
        operator: Excel.ConditionalIconCriterionOperator.greaterThanOrEqual,
        formula: "=1000",
      },
    ];

    await context.sync();
  });
}

async function applyTextFormat() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("B16:D18");
    const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
    conditionalFormat.textComparison.format.font.color = "red";
    conditionalFormat.textComparison.rule = { operator: Excel.ConditionalTextOperator.contains, text: "Delayed" };

    await context.sync();
  });
}

async function applyCellValueFormat() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("B21:E23");
    const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
    conditionalFormat.cellValue.format.font.color = "red";
    conditionalFormat.cellValue.rule = { formula1: "=0", operator: "LessThan" };

    await context.sync();
  });
}

async function applyTopBottomFormat() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("B21:E23");
    const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
    conditionalFormat.topBottom.format.fill.color = "green";
    conditionalFormat.topBottom.rule = { rank: 1, type: "TopItems" };

    await context.sync();
  });
}

async function applyCustomFormat() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("B8:E13");
    const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormat.custom.rule.formula = '=IF(B8>INDIRECT("RC[-1]",0),TRUE)';
    conditionalFormat.custom.format.font.color = "green";

    await context.sync();
  });
}
async function clearAllConditionalFormats() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange();
    range.conditionalFormats.clearAll();

    await context.sync();
  });
}
// 保存当前工作表所有条件格式
async function saveConditionalFormats() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const worksheetRange = sheet.getRange();
    worksheetRange.conditionalFormats.load("type");

    await context.sync();

    savedConditionalFormats = [];
    worksheetRange.conditionalFormats.items.forEach((item) => {
      let savedCF = {
        type: item.type,
        rangeAddress: item.getRange().address,
        criteria: [],
        format: null,
        rule: null,
      };

      switch (item.type) {
        case "ContainsText":
        case "CellValue":
        case "TopBottom":
        case "Custom":
          savedCF.format = item.custom.format;
          savedCF.rule = item.custom.rule;
          break;
      }

      savedConditionalFormats.push(savedCF);
    });

    console.log("保存的条件格式信息:");
    console.log(savedConditionalFormats);
  });
}
