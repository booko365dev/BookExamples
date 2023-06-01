
//gavdcodebegin 001
function main(workbook: ExcelScript.Workbook) {
    let mySheet = workbook.getActiveWorksheet();
    console.log(mySheet.getName())
}
//gavdcodeend 001

//gavdcodebegin 002
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("Sheet1");
    let wsCollection = workbook.getWorksheets();
    let wsByPosition = workbook.getWorksheets()[0];
    let wsCurrent = workbook.getActiveWorksheet();
    let wsFirst = workbook.getFirstWorksheet();
    let wsLast = workbook.getLastWorksheet();
}
//gavdcodeend 002

//gavdcodebegin 003
function main(workbook: ExcelScript.Workbook) {
    let wsCollection = workbook.getWorksheets();
    console.log(wsCollection.length);
}
//gavdcodeend 003

//gavdcodebegin 004
function main(workbook: ExcelScript.Workbook) {
    let wsCollection = workbook.getWorksheets();
    for (let items = 0; items < wsCollection.length; items++) {
        console.log(wsCollection[items].getName());
    };
}
//gavdcodeend 004

//gavdcodebegin 005
function main(workbook: ExcelScript.Workbook) {
    let wsCollection = workbook.getWorksheets();
    wsCollection.forEach(item => {
        console.log(item.getName());
    });
}
//gavdcodeend 005

//gavdcodebegin 006
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("Sheet1");
    if (wsByName != undefined) {
        console.log("Worksheet found");
        return;
    } else {
        console.log("Worksheet not found");
    };
}
//gavdcodeend 006

//gavdcodebegin 007
function main(workbook: ExcelScript.Workbook) {
    workbook.addWorksheet();
    workbook.addWorksheet("MyWorksheet");
}
//gavdcodeend 007

//gavdcodebegin 008
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("Sheet1");
    let wsNewSheet01 = wsByName.copy();
    let wsForPosition = workbook.getWorksheet("MyWorksheet");
    let wsNewSheet02 = wsByName.copy(ExcelScript.WorksheetPositionType.after,
                                                                        wsForPosition);
}
//gavdcodeend 008

//gavdcodebegin 009
function main(workbook: ExcelScript.Workbook) {
    let wsByPosition = workbook.getWorksheets()[0];
    console.log(wsByPosition.getName());
    wsByPosition.setName("SheetNewName");
    console.log(wsByPosition.getName());
}
//gavdcodeend 009

//gavdcodebegin 010
function main(workbook: ExcelScript.Workbook) {
    let wsByPosition = workbook.getWorksheets()[0];
    console.log(wsByPosition.getPosition());
    wsByPosition.setPosition(1);
    console.log(wsByPosition.getPosition());
}
//gavdcodeend 010

//gavdcodebegin 011
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");
    console.log(wsByName.getVisibility());
    wsByName.setVisibility(ExcelScript.SheetVisibility.hidden);
    console.log(wsByName.getVisibility());
    wsByName.setVisibility(ExcelScript.SheetVisibility.veryHidden);
    console.log(wsByName.getVisibility());
    wsByName.setVisibility(ExcelScript.SheetVisibility.visible);
    console.log(wsByName.getVisibility());
}
//gavdcodeend 011

//gavdcodebegin 012
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");
    let isProtected = wsByName.getProtection().getProtected()
    console.log(isProtected);
}
//gavdcodeend 012

//gavdcodebegin 013
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");
    console.log(wsByName.getProtection().getOptions().allowAutoFilter);
    console.log(wsByName.getProtection().getOptions().allowDeleteColumns);
    console.log(wsByName.getProtection().getOptions().allowDeleteRows);
    console.log(wsByName.getProtection().getOptions().allowEditObjects);
    console.log(wsByName.getProtection().getOptions().allowEditScenarios);
    console.log(wsByName.getProtection().getOptions().allowFormatCells);
    console.log(wsByName.getProtection().getOptions().allowFormatColumns);
    console.log(wsByName.getProtection().getOptions().allowFormatRows);
    console.log(wsByName.getProtection().getOptions().allowInsertColumns);
    console.log(wsByName.getProtection().getOptions().allowInsertHyperlinks);
    console.log(wsByName.getProtection().getOptions().allowInsertRows);
    console.log(wsByName.getProtection().getOptions().allowPivotTables);
    console.log(wsByName.getProtection().getOptions().allowSort);
    console.log(wsByName.getProtection().getOptions().selectionMode);
}
//gavdcodeend 013

//gavdcodebegin 014
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");
    wsByName.getProtection().protect();
    wsByName.getProtection().unprotect();
}
//gavdcodeend 014

//gavdcodebegin 015
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");
    wsByName.getProtection().protect({}, "myPW");
    wsByName.getProtection().unprotect("myPW");
}
//gavdcodeend 015

//gavdcodebegin 016
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");
    wsByName.getProtection().protect({
        allowAutoFilter: false,
        allowDeleteRows: false,
        allowDeleteColumns: false,
        allowEditObjects: false,
        allowEditScenarios: false,
        allowFormatCells: false,
        allowFormatColumns: false,
        allowFormatRows: false,
        allowInsertColumns: false,
        allowInsertHyperlinks: false,
        allowInsertRows: false,
        allowPivotTables: false,
        allowSort: false
    });
}
//gavdcodeend 016

//gavdcodebegin 017
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");
    wsByName.activate();
}
//gavdcodeend 017

//gavdcodebegin 018
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    wsByName.getRange("A1:D4").setValues([
        [1, 2, 3, 4],
        [5, 6, 7, 8],
        [9, 10, 11, 12],
        [13, 14, 15, 16]]);

    console.log(wsByName.getRange("A1:D4").getValues());
}
//gavdcodeend 018

//gavdcodebegin 019
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    console.log(wsByName.getRange("A1:D4").getValues());
    console.log(wsByName.getRange("B2").getValues());
}
//gavdcodeend 019

//gavdcodebegin 020
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    wsByName.getRange("B2").setValues([[111]]);

    console.log(wsByName.getRange("A1:D4").getValues());
}
//gavdcodeend 020

//gavdcodebegin 021
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    wsByName.getRange("B2:C3").clear(ExcelScript.ClearApplyTo.formats);
    wsByName.getRange("B2:C3").clear();

    console.log(wsByName.getRange("A1:D4").getValues());
}
//gavdcodeend 021

//gavdcodebegin 022
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    wsByName.getRange("A1").delete(ExcelScript.DeleteShiftDirection.left);
    wsByName.getRange("A1").delete(ExcelScript.DeleteShiftDirection.up);

    console.log(wsByName.getRange("A1:D4").getValues());
}
//gavdcodeend 022

//gavdcodebegin 023
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    console.log(wsByName.getRange("A1:D4").getRowCount());
    console.log(wsByName.getRange("A1:D4").getColumnCount());
}
//gavdcodeend 023

//gavdcodebegin 024
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    let a1Value = wsByName.getRange("A1");
    a1Value.setValue(101);

    let b1Value = wsByName.getRange("B1")
    b1Value.setFormula("=(10*A1)");

    console.log('B1 - Formula: ${b1Value.getFormula()} - Value: ${b1Value.getValue()}');
}
//gavdcodeend 024

//gavdcodebegin 025
function main(workbook: ExcelScript.Workbook) {
    let fileName = workbook.getName();
    console.log(fileName);
}
//gavdcodeend 025

//gavdcodebegin 026
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    let myRange = wsByName.getRange("A1:D4")
    let newTable = wsByName.addTable(myRange, true);
}
//gavdcodeend 026

//gavdcodebegin 027
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    console.log(wsByName.getTables());
}
//gavdcodeend 027

//gavdcodebegin 028
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    let myTable = wsByName.getTables()[0];
    myTable.setName("TableOne");
}
//gavdcodeend 028

//gavdcodebegin 029
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    let myTable = wsByName.getTables()[0];
    myTable.addRow(2, [11, 22, 33, 44]);
}
//gavdcodeend 029

//gavdcodebegin 030
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    let myTable = wsByName.getTables()[0];
    myTable.addColumn(1, [55, 66, 77, 88, 99], 'zz');
}
//gavdcodeend 030

//gavdcodebegin 031
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    let myTable = wsByName.getTables()[0];
    console.log(myTable.getColumn(1));
    console.log(myTable.getColumn('zz'));
    console.log(myTable.getColumnById(2));
    console.log(myTable.getColumnByName('b'));
    console.log(myTable.getRowCount());
}
//gavdcodeend 031

//gavdcodebegin 032
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    let myTable = wsByName.getTables()[0];
    myTable.getColumn('zz').delete();
    myTable.deleteRowsAt(2, 1);
}
//gavdcodeend 032

//gavdcodebegin 033
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    let myChart = wsByName.addChart(ExcelScript.ChartType.columnClustered, selectedSheet.getRange("B7:D8"));
}
//gavdcodeend 033

//gavdcodebegin 034
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    let myChart = wsByName.addChart(ExcelScript.ChartType.columnClustered, selectedSheet.getRange("B7:D8"));
    myChart.setPosition("A1");
    myChart.getTitle().setText("My Chart");
    myChart.getTitle().setTextOrientation(180);
    myChart.Name("MyChart");
}
//gavdcodeend 034

//gavdcodebegin 035
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    let myChart = wsByName.getChart("MyChart");
    myChart.delete();
}
//gavdcodeend 035

//gavdcodebegin 036
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    console.log(wsByName.getCharts());
}
//gavdcodeend 036
