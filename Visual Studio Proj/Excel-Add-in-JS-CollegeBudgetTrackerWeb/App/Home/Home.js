/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
	See full license at the bottom of this file. */

/// <reference path="../App.js" />

(function () {
	"use strict";

	// The initialize function must be run each time a new page is loaded
	Office.initialize = function (reason) {
		$(document).ready(function () {
		    app.initialize();

		    $('#add-expense').click(addExpense);
		    $('#add-income').click(addIncome);

		    // If not using Excel 2016, return
			if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
			    app.showNotification("Need Office 2016 or greater", "Sorry, this add-in only works with newer versions of Excel.");
			    return;
			}

            // Wire up the Dropdown control
			if ($.fn.Dropdown) {
			    $('.ms-Dropdown').Dropdown();
			}

            // Wire up the Pivot control
			if ($.fn.Pivot) {
			    $('.ms-Pivot').Pivot();
			}
			
			createBudgetAnalyzer();
		});
	};

	function createBudgetAnalyzer() {

		// Run a batch operation against the Excel object model
		Excel.run(function (ctx) {

			// Create a proxy object for the active worksheet
			var dashboardSheet = ctx.workbook.worksheets.getActiveWorksheet();

			// Queue a command to clear the contents before inserting data
			dashboardSheet.getUsedRange().clear();

			// Queue a command to rename the sheet
			dashboardSheet.name = "Dashboard";

			// Queue commands to set the title and format it
			var title = "College Budget Tracker";
			var range = dashboardSheet.getRange("A1");
			range.values = title;
			range.format.font.name = "Rockwell";
			range.format.font.size = 22.5;


			// Queue commands to add the Expenses table at the bottom of the sheet with sample data
			var expenseTable = ctx.workbook.tables.add('Dashboard!A117:C117', true);
			expenseTable.name = "expenseTable";

			expenseTable.getHeaderRowRange().values = [["Description", "Cost", "Category"]];
			var tableRows = expenseTable.rows;

			tableRows.add(null, [["Rent", "$600", "Housing"]]);
			tableRows.add(null, [["Movie Club", "$75", "Entertainment"]]);
			tableRows.add(null, [["Food", "$450", "Food"]]);
			tableRows.add(null, [["Car", "$150", "Transportation"]]);
			tableRows.add(null, [["Tuition", "$800", "School costs"]]);
			tableRows.add(null, [["Books", "$150", "School costs"]]);
			tableRows.add(null, [["Gift", "$100", "Other"]]);
			tableRows.add(null, [["Loan", "$250", "Loans/Payments"]]);

			// Queue commands to set the title for the Expenses table and format it
			var expenseTableTitle = "Monthly Expenses";
			var range = dashboardSheet.getRange("A116");
			range.values = expenseTableTitle;
			range.format.font.name = "Rockwell";
			range.format.font.size = 18;

			// Queue commands to add the Income table at the bottom of the sheet with sample data
			var incomeTable = ctx.workbook.tables.add('Dashboard!F117:H117', true);
			incomeTable.name = "incomeTable";

			incomeTable.getHeaderRowRange().values = [["Description", "Amount", "Category"]];
			var tableRows = incomeTable.rows;

			tableRows.add(null, [["Wages", "$2500", "Wages"]]);
			tableRows.add(null, [["Parents", "$700", "Assistance from parents"]]);
			tableRows.add(null, [["Gift", "$100", "Other"]]);
			tableRows.add(null, [["Bank interest", "$250", "From savings"]]);
			tableRows.add(null, [["Scholarship", "$500", "Financial aid"]]);

			// Queue commands to set the title for the Expenses table and format it
			var incomeTableTitle = "Monthly Income";
			var range = dashboardSheet.getRange("F116");
			range.values = incomeTableTitle;
			range.format.font.name = "Rockwell";
			range.format.font.size = 18;


			// Queue commands to create the summary section at the top right
			var summaryValues = [["Percentage of income spent", "=D4/D3"],
								  ["Income", '=SUM(G117:G217)'],
								  ["Expenses", '=SUM(B117:B217)'],
								  ["Balance", "=D3-D4"]];

			// Set the number format before setting the values
			dashboardSheet.getRange("D2:D2").numberFormat = "0.00%";
			dashboardSheet.getRange("D3:D5").numberFormat = "$#";
			dashboardSheet.getRange("C2:D5").values = summaryValues;

			dashboardSheet.getRange("C2:D2").format.font.size = 18;
			dashboardSheet.getRange("C2:D2").format.font.color = "red";
			dashboardSheet.getRange("C2:D5").format.font.name = "Rockwell";
			dashboardSheet.getRange("C3:D5").format.font.size = 10;
			dashboardSheet.getRange("C2:D5").format.borders.getItem("InsideHorizontal").style = "Continuous";
			dashboardSheet.getRange("C2:D5").format.borders.getItem('EdgeBottom').style = 'Continuous';
			dashboardSheet.getRange("C2:D5").format.borders.getItem('EdgeTop').style = 'Continuous';
			dashboardSheet.getRange("C5:D5").format.font.size = 13;
			dashboardSheet.getRange("C5:D5").format.font.name = "Rockwell";


			// Queue commands to create the Money In section
			var moneyInValues = [["Money coming in", ""],
								 ["Category", "Amount"],
								 ["Wages", '=IFERROR(SUMIFS(G117:G217,H117:H217,C10),"")'],
								 ["Financial aid", '=IFERROR(SUMIFS(G117:G217,H117:H217,C11),"")'],
								 ["From savings", '=IFERROR(SUMIFS(G117:G217,H117:H217,C12),"")'],
								 ["Assistance from parents", '=IFERROR(SUMIFS(G117:G217,H117:H217,C13),"")'],
								 ["Other", '=IFERROR(SUMIFS(G117:G217,H117:H217,C14),"")'],
								 ["Total", "=sum(D10:D14)"]];

			// Set the number format before setting the values
			dashboardSheet.getRange("D10:D15").numberFormat = "$#";
			dashboardSheet.getRange("C8:D15").values = moneyInValues;
			dashboardSheet.getRange("C8:D8").format.font.size = 18;
			dashboardSheet.getRange("C8:D8").format.font.color = "red";
			dashboardSheet.getRange("C8:D15").format.font.name = "Rockwell";
			dashboardSheet.getRange("C9:D9").format.font.size = 13;
			dashboardSheet.getRange("C10:D14").format.font.size = 10;
			dashboardSheet.getRange("C8:D15").format.borders.getItem("InsideHorizontal").style = "Continuous";
			dashboardSheet.getRange("C8:D15").format.borders.getItem('EdgeBottom').style = 'Continuous';
			dashboardSheet.getRange("C8:D15").format.borders.getItem('EdgeTop').style = 'Continuous';
			dashboardSheet.getRange("C15:D15").format.font.size = 13;
			dashboardSheet.getRange("C15:D15").format.font.name = "Rockwell";

			// Queue commands to create the Money Out section
			var moneyOutValues = [["Money going out", ""],
								  ["Category", "Cost"],
								  ["School costs", '=IFERROR(SUMIFS(B117:B217,C117:C217,C20),"")'],
								  ["Entertainment", '=IFERROR(SUMIFS(B117:B217,C117:C217,C21),"")'],
								  ["Food", '=IFERROR(SUMIFS(B117:B217,C117:C217,C22),"")'],
								["Housing", '=IFERROR(SUMIFS(B117:B217,C117:C217,C23),"")'],
								  ["Transportation", '=IFERROR(SUMIFS(B117:B217,C117:C217,C24),"")'],
								  ["Loans/Payments", '=IFERROR(SUMIFS(B117:B217,C117:C217,C25),"")'],
								  ["Other", '=IFERROR(SUMIFS(B117:B217,C117:C217,C26),"")'],
								  ["Total", "=sum(D20:D26)"]];
			
			// Set the number format before setting the values
			dashboardSheet.getRange("D19:D27").numberFormat = "$#";
			dashboardSheet.getRange("C18:D27").values = moneyOutValues;
			dashboardSheet.getRange("C18:D18").format.font.size = 18;
			dashboardSheet.getRange("C18:D18").format.font.color = "red";
			dashboardSheet.getRange("C18:D27").format.font.name = "Rockwell";
			dashboardSheet.getRange("C19:D19").format.font.size = 13;;
			dashboardSheet.getRange("C20:D26").format.font.size = 10;
			dashboardSheet.getRange("C18:D27").format.borders.getItem("InsideHorizontal").style = "Continuous";
			dashboardSheet.getRange("C18:D27").format.borders.getItem('EdgeBottom').style = 'Continuous';
			dashboardSheet.getRange("C18:D27").format.borders.getItem('EdgeTop').style = 'Continuous';
			dashboardSheet.getRange("C27:D27").format.font.size = 13;
			dashboardSheet.getRange("C27:D27").format.font.name = "Rockwell";

		    // Queue commands to autofit rows and columns in a sheet
			dashboardSheet.getUsedRange().getEntireColumn().format.autofitColumns();
			dashboardSheet.getUsedRange().getEntireRow().format.autofitRows();

			// Queue commands to create the income chart
			var incomeChartDataRange = dashboardSheet.getRange("C10:D14");
			var chart = dashboardSheet.charts.add("doughnut", incomeChartDataRange, "auto");
			chart.setPosition("A3", "A13");
			chart.title.text = "Income";
			chart.title.format.font.size = 15;
			chart.title.format.font.color = "red";
			chart.legend.position = "left";
			chart.legend.format.font.name = "Trebuchet MS (Body)";
			chart.legend.format.font.size = 8;
			chart.dataLabels.showPercentage = true;
			chart.dataLabels.format.font.size = 8;
			chart.dataLabels.format.font.color = "white";
			var points = chart.series.getItemAt(0).points;
			points.getItemAt(0).format.fill.setSolidColor("#ff3300");
			points.getItemAt(1).format.fill.setSolidColor("#00cccc");
			points.getItemAt(2).format.fill.setSolidColor("#bf6514");
			points.getItemAt(3).format.fill.setSolidColor("#2be6c2");
			points.getItemAt(4).format.fill.setSolidColor("#993cf3");


			// Queue commands to create the expenses chart
			var expenseChartDataRange = dashboardSheet.getRange("C20:D25");
			var expenseChart = dashboardSheet.charts.add("doughnut", expenseChartDataRange, "auto");
			expenseChart.setPosition("A16", "A26");
			expenseChart.title.text = "Expenses";
			expenseChart.title.format.font.size = 15;
			expenseChart.title.format.font.color = "red";
			expenseChart.legend.position = "left";
			expenseChart.legend.format.font.name = "Trebuchet MS (Body)";
			expenseChart.legend.format.font.size = 8;
			expenseChart.dataLabels.showPercentage = true;
			expenseChart.dataLabels.format.font.size = 8;
			expenseChart.dataLabels.format.font.color = "white";
			var points = expenseChart.series.getItemAt(0).points;
			points.getItemAt(0).format.fill.setSolidColor("#ff3300");
			points.getItemAt(1).format.fill.setSolidColor("#00cccc");
			points.getItemAt(2).format.fill.setSolidColor("#bf6514");
			points.getItemAt(3).format.fill.setSolidColor("#2be6c2");
			points.getItemAt(4).format.fill.setSolidColor("#993cf3");

			// Run the queued-up commands, and return a promise to indicate task completion
			return ctx.sync();
		})
		.catch(function (error) {
			// Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
			app.showNotification("Error: " + error);
			console.log("Error: " + error);
			if (error instanceof OfficeExtension.Error) {
				console.log("Debug info: " + JSON.stringify(error.debugInfo));
			}
		});
	}

	function addExpense() {

		// Run a batch operation against the Excel object model
		Excel.run(function (ctx) {

			// Create a proxy object for the expense table rows
			var tableRows = ctx.workbook.tables.getItem('expenseTable').rows;
			tableRows.add(null, [[$("#expense-description").val(), $("#expense-cost").val(), $("#expense-category").val()]]);

			// Run the queued-up commands, and return a promise to indicate task completion
			return ctx.sync();
		})
		.catch(function (error) {
			// Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
			app.showNotification("Error: " + error);
			console.log("Error: " + error);
			if (error instanceof OfficeExtension.Error) {
				console.log("Debug info: " + JSON.stringify(error.debugInfo));
			}
		});
	}


	function addIncome() {
		// Run a batch operation against the Excel object model
		Excel.run(function (ctx) {

			// Create a proxy object for the expense table rows
			var tableRows = ctx.workbook.tables.getItem('incomeTable').rows;

			// Run the queued-up commands, and return a promise to indicate task completion
			tableRows.add(null, [[$("#income-description").val(), $("#income-amount").val(), $("#income-category").val()]]);

			// Run the queued-up commands, and return a promise to indicate task completion
			return ctx.sync();
		})
		.catch(function (error) {
			// Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
			app.showNotification("Error: " + error);
			console.log("Error: " + error);
			if (error instanceof OfficeExtension.Error) {
				console.log("Debug info: " + JSON.stringify(error.debugInfo));
			}
		});
	}
})();

/* 
Excel-Add-in-JS-CollegeBudgetTracker, https://github.com/OfficeDev/Excel-Add-in-JS-CollegeBudgetTracker

Copyright (c) Microsoft Corporation

All rights reserved.

MIT License:

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
associated documentation files (the "Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the
following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial
portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT
LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT
SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE
USE OR OTHER DEALINGS IN THE SOFTWARE.
*/