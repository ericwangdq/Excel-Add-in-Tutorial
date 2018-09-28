/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import "babel-polyfill";

(function() {
  Office.initialize = function(reason) {
    $(document).ready(function() {
      if (!Office.context.requirements.isSetSupported("ExcelApi", 1.8)) {
        console.log(
          "Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office."
        );
      }

      $("#create-table").click(createDataTable);
      $("#create-pivot-table").click(createPivotTable);
      $("#add-rows-columns").click(addRowsAndColumns);
      $("#create-chart").click(createChart);
      $("#open-dialog").click(openDialog);
    });
  };

  function createTable() {
    Excel.run(function(context) {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      const expensesTable = currentWorksheet.tables.add(
        "A1:D1",
        true /*hasHeaders*/
      );
      expensesTable.name = "ExpensesTable";

      expensesTable.getHeaderRowRange().values = [
        ["Date", "Merchant", "Category", "Amount"]
      ];

      expensesTable.rows.add(null /*add at the end*/, [
        ["1/1/2017", "The Phone Company", "Communications", "120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
        ["1/11/2017", "Bellows College", "Education", "350.1"],
        ["1/15/2017", "Trey Research", "Other", "135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
      ]);

      expensesTable.columns.getItemAt(3).getRange().numberFormat = [
        ["€#,##0.00"]
      ];
      expensesTable.getRange().format.autofitColumns();
      expensesTable.getRange().format.autofitRows();

      return context.sync();
    }).catch(function(error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  }

  function createDataTable() {
    Excel.run(function(context) {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      const expensesTable = currentWorksheet.tables.add(
        "A1:E1",
        true /*hasHeaders*/
      );
      expensesTable.name = "ExpensesTable";

      expensesTable.getHeaderRowRange().values = [
        [
          "Farm",
          "Type",
          "Classification",
          "Crates Sold at Farm",
          "Crates Sold Wholesale"
        ]
      ];

      expensesTable.rows.add(null /*add at the end*/, [
        ["A Farms", "Lime", "Organic", "300", "2000"],
        ["A Farms", "Lemon", "Organic", "250", "1800"],
        ["A Farms", "Orange", "Organic", "300", "2200"],
        ["B Farms", "Lime", "Conventional", "80", "1000"],
        ["B Farms", "Lemon", "Conventional", "75", "1230"],
        ["B Farms", "Orange", "Conventional", "25", "800"]
      ]);

      // expensesTable.columns.getItemAt(3).getRange().numberFormat = [
      //   ["€#,##0.00"]
      // ];
      expensesTable.getRange().format.autofitColumns();
      expensesTable.getRange().format.autofitRows();

      return context.sync();
    }).catch(function(error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  }

  async function createPivotTable() {
    await Excel.run(async context => {
      // creating a PivotTable named "Farm Sales" created on the current worksheet at cell A22 with data from the range A1:E21
      console.log("create pivot table");
      // creating a PivotTable named "Farm Sales" created on the current worksheet at cell A22 with data from the range A1:E7
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      const rangeToAnalyze = currentWorksheet.getRange("A1:E7");
      const rangeToPlacePivot = currentWorksheet.getRange("A11");
      // const pivotTable = currentWorksheet.pivotTables.add(
      //   "Farm Sales",
      //   "A1:E7",
      //   "A8"
      // );

      // const pivotTable = currentWorksheet.pivotTables.add("Farm Sales");
      currentWorksheet.pivotTables.add(
        "Farm Sales",
        rangeToAnalyze,
        rangeToPlacePivot
      );
      await context.sync();
    });
  }

  async function addRowsAndColumns() {
    await Excel.run(async context => {
      console.log("addRowsAndColumns");
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      const pivotTable = currentWorksheet.pivotTables.getItem("Farm Sales");
      // pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
      // pivotTable.rowHierarchies.add(
      //   pivotTable.hierarchies.getItem("Classification")
      // );
      // pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

      // "Farm" and "Type" are the hierarchies on which the aggregation is based
      pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
      pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));

      // "Crates Sold at Farm" and "Crates Sold Wholesale" are the heirarchies that will have their data aggregated (summed in this case)
      pivotTable.dataHierarchies.add(
        pivotTable.hierarchies.getItem("Crates Sold at Farm")
      );
      pivotTable.dataHierarchies.add(
        pivotTable.hierarchies.getItem("Crates Sold Wholesale")
      );

      await context.sync();
    });
  }

  function createChart() {
    console.log("create chart");
    Excel.run(function(context) {
      // TODO1: Queue commands to get the range of data to be charted.
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
      const dataRange = expensesTable.getDataBodyRange();

      // TODO2: Queue command to create the chart and define its type.
      let chart = currentWorksheet.charts.add(
        "ColumnClustered",
        dataRange,
        "auto"
      );

      // TODO3: Queue commands to position and format the chart.
      chart.setPosition("A15", "F30");
      chart.title.text = "Expenses";
      chart.legend.position = "right";
      chart.legend.format.fill.setSolidColor("white");
      chart.dataLabels.format.font.size = 15;
      chart.dataLabels.format.font.color = "black";
      chart.series.getItemAt(0).name = "Value in €";

      return context.sync();
    }).catch(function(error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  }

  let dialog = null;
  function openDialog() {
    Office.context.ui.displayDialogAsync(
      "https://localhost:3000/popup.html",
      { height: 45, width: 55 },
      function(result) {
        dialog = result.value;
        dialog.addEventHandler(
          Microsoft.Office.WebExtension.EventType.DialogMessageReceived,
          processMessage
        );
      }
    );

    function processMessage(arg) {
      $("#user-name").text(arg.message);
      dialog.close();
    }
  }
})();
