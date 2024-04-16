/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
//require("dotenv").config();

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("setup").onclick = setup;
    document.getElementById("enable-highlighter").onclick = enableCellHighlight;
    document.getElementById("disable-highlighter").onclick = disableCellHighlight;
  }
});

//TODO
//2024.4.13:To remove the handler by code.
let eventResult;
async function disableCellHighlight() {
  await Excel.run(eventResult.context, async (context) => {
    let workbook = context.workbook;
    let selectedSheet = workbook.worksheets.getActiveWorksheet();
    selectedSheet.getRange().style = Excel.BuiltInStyle.normal;
    eventResult.remove();
    await context.sync();
  });
}
async function enableCellHighlight() {
  await Excel.run(async (context) => {
    let workbook = context.workbook;
    let selectedSheet = workbook.worksheets.getActiveWorksheet();
    eventResult = selectedSheet.onSelectionChanged.add(CellHighlightHandler);
    await context.sync();
  });
}
async function CellHighlightHandler(event) {
  await Excel.run(async (context) => {
    let workbook = context.workbook;
    let selectedSheet = workbook.worksheets.getActiveWorksheet();
    let selection = workbook.getSelectedRange();
    selection.load("rowIndex,columnIndex");
    await context.sync();
    // Assuming 'rowIndex' and 'columnIndex' are the row and column index of the selected cell
    let rowIndex = selection.rowIndex;
    console.log(`=ROW()= + ${selection.rowIndex + 1})`);
    console.log("Address of current selection: " + event.address);
    let columnIndex = selection.columnIndex;
    // Convert column index to letter
    let colLetter = String.fromCharCode(65 + columnIndex); // 65 is the ASCII value for 'A'
    selectedSheet.getRange().style = Excel.BuiltInStyle.normal;
    // Apply the style to the entire row and column
    selectedSheet.getRange(rowIndex + 1 + ":" + (rowIndex + 1)).style = Excel.BuiltInStyle.neutral;
    selectedSheet.getRange(colLetter + ":" + colLetter).style = Excel.BuiltInStyle.neutral;
  });
}

async function addNewStyle() {
  await Excel.run(async (context) => {
    let styles = context.workbook.styles;

    // Add a new style to the style collection.
    // Styles is in the Home tab ribbon.
    styles.add("Highlighter");

    let newStyle = styles.getItem("Highlighter");

    // The "Diagonal Orientation Style" properties.
    newStyle.includeFont = true;
    newStyle.fill.color = "green";
    await context.sync();

    console.log("Successfully added a new style with Highlighter to the Home tab ribbon.");
  });
}

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
    applyColorScaleFormat();
    applyPresetFormat();
    applyDataBarFormat();
    applyIconSetFormat();
    applyTextFormat();
    applyCellValueFormat();
    applyTopBottomFormat();
    applyCustomFormat();
    //addNewStyle();
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
