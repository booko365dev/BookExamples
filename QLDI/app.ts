
//gavdcodebegin 01
function main(workbook: ExcelScript.Workbook) {
    let mySheet = workbook.getActiveWorksheet();
    console.log(mySheet.getName())
}
//gavdcodeend 01

//gavdcodebegin 02
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("Sheet1");
    let wsCollection = workbook.getWorksheets();
    let wsByPosition = workbook.getWorksheets()[0];
    let wsCurrent = workbook.getActiveWorksheet();
    let wsFirst = workbook.getFirstWorksheet();
    let wsLast = workbook.getLastWorksheet();
}
//gavdcodeend 02

//gavdcodebegin 03
function main(workbook: ExcelScript.Workbook) {
    let wsCollection = workbook.getWorksheets();
    console.log(wsCollection.length);
}
//gavdcodeend 03

//gavdcodebegin 04
function main(workbook: ExcelScript.Workbook) {
    let wsCollection = workbook.getWorksheets();
    for (let items = 0; items < wsCollection.length; items++) {
        console.log(wsCollection[items].getName());
    };
}
//gavdcodeend 04

//gavdcodebegin 05
function main(workbook: ExcelScript.Workbook) {
    let wsCollection = workbook.getWorksheets();
    wsCollection.forEach(item => {
        console.log(item.getName());
    });
}
//gavdcodeend 05

//gavdcodebegin 06
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("Sheet1");
    if (wsByName != undefined) {
        console.log("Worksheet found");
        return;
    } else {
        console.log("Worksheet not found");
    };
}
//gavdcodeend 06

//gavdcodebegin 07
function main(workbook: ExcelScript.Workbook) {
    workbook.addWorksheet();
    workbook.addWorksheet("MyWorksheet");
}
//gavdcodeend 07

//gavdcodebegin 08
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("Sheet1");
    let wsNewSheet01 = wsByName.copy();
    let wsForPosition = workbook.getWorksheet("MyWorksheet");
    let wsNewSheet02 = wsByName.copy(ExcelScript.WorksheetPositionType.after,
                                                                        wsForPosition);
}
//gavdcodeend 08

//gavdcodebegin 09
function main(workbook: ExcelScript.Workbook) {
    let wsByPosition = workbook.getWorksheets()[0];
    console.log(wsByPosition.getName());
    wsByPosition.setName("SheetNewName");
    console.log(wsByPosition.getName());
}
//gavdcodeend 09

//gavdcodebegin 10
function main(workbook: ExcelScript.Workbook) {
    let wsByPosition = workbook.getWorksheets()[0];
    console.log(wsByPosition.getPosition());
    wsByPosition.setPosition(1);
    console.log(wsByPosition.getPosition());
}
//gavdcodeend 10

//gavdcodebegin 11
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
//gavdcodeend 11

//gavdcodebegin 12
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");
    let isProtected = wsByName.getProtection().getProtected()
    console.log(isProtected);
}
//gavdcodeend 12

//gavdcodebegin 13
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
//gavdcodeend 13

//gavdcodebegin 14
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");
    wsByName.getProtection().protect();
    wsByName.getProtection().unprotect();
}
//gavdcodeend 14

//gavdcodebegin 15
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");
    wsByName.getProtection().protect({}, "myPW");
    wsByName.getProtection().unprotect("myPW");
}
//gavdcodeend 15

//gavdcodebegin 16
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
//gavdcodeend 16

//gavdcodebegin 17
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");
    wsByName.activate();
}
//gavdcodeend 17

//gavdcodebegin 18
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    wsByName.getRange("A1:D4").setValues([
        [1, 2, 3, 4],
        [5, 6, 7, 8],
        [9, 10, 11, 12],
        [13, 14, 15, 16]]);

    console.log(wsByName.getRange("A1:D4").getValues());
}
//gavdcodeend 18

//gavdcodebegin 19
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    console.log(wsByName.getRange("A1:D4").getValues());
    console.log(wsByName.getRange("B2").getValues());
}
//gavdcodeend 19

//gavdcodebegin 20
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    wsByName.getRange("B2").setValues([[111]]);

    console.log(wsByName.getRange("A1:D4").getValues());
}
//gavdcodeend 20

//gavdcodebegin 21
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    wsByName.getRange("B2:C3").clear(ExcelScript.ClearApplyTo.formats);
    wsByName.getRange("B2:C3").clear();

    console.log(wsByName.getRange("A1:D4").getValues());
}
//gavdcodeend 21

//gavdcodebegin 22
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    wsByName.getRange("A1").delete(ExcelScript.DeleteShiftDirection.left);
    wsByName.getRange("A1").delete(ExcelScript.DeleteShiftDirection.up);

    console.log(wsByName.getRange("A1:D4").getValues());
}
//gavdcodeend 22

//gavdcodebegin 23
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    console.log(wsByName.getRange("A1:D4").getRowCount());
    console.log(wsByName.getRange("A1:D4").getColumnCount());
}
//gavdcodeend 23

//gavdcodebegin 24
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    let a1Value = wsByName.getRange("A1");
    a1Value.setValue(101);

    let b1Value = wsByName.getRange("B1")
    b1Value.setFormula("=(10*A1)");

    console.log('B1 - Formula: ${b1Value.getFormula()} - Value: ${b1Value.getValue()}');
}
//gavdcodeend 24

//gavdcodebegin 25
function main(workbook: ExcelScript.Workbook) {
    let fileName = workbook.getName();
    console.log(fileName);
}
//gavdcodeend 25

//gavdcodebegin 26
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    let myRange = wsByName.getRange("A1:D4")
    let newTable = wsByName.addTable(myRange, true);
}
//gavdcodeend 26

//gavdcodebegin 27
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    console.log(wsByName.getTables());
}
//gavdcodeend 27

//gavdcodebegin 28
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    let myTable = wsByName.getTables()[0];
    myTable.setName("TableOne");
}
//gavdcodeend 28

//gavdcodebegin 29
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    let myTable = wsByName.getTables()[0];
    myTable.addRow(2, [11, 22, 33, 44]);
}
//gavdcodeend 29

//gavdcodebegin 30
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    let myTable = wsByName.getTables()[0];
    myTable.addColumn(1, [55, 66, 77, 88, 99], 'zz');
}
//gavdcodeend 30

//gavdcodebegin 31
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    let myTable = wsByName.getTables()[0];
    console.log(myTable.getColumn(1));
    console.log(myTable.getColumn('zz'));
    console.log(myTable.getColumnById(2));
    console.log(myTable.getColumnByName('b'));
    console.log(myTable.getRowCount());
}
//gavdcodeend 31

//gavdcodebegin 32
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    let myTable = wsByName.getTables()[0];
    myTable.getColumn('zz').delete();
    myTable.deleteRowsAt(2, 1);
}
//gavdcodeend 32

//gavdcodebegin 33
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    let myChart = wsByName.addChart(ExcelScript.ChartType.columnClustered, selectedSheet.getRange("B7:D8"));
}
//gavdcodeend 33

//gavdcodebegin 34
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    let myChart = wsByName.addChart(ExcelScript.ChartType.columnClustered, selectedSheet.getRange("B7:D8"));
    myChart.setPosition("A1");
    myChart.getTitle().setText("My Chart");
    myChart.getTitle().setTextOrientation(180);
    myChart.Name("MyChart");
}
//gavdcodeend 34

//gavdcodebegin 35
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    let myChart = wsByName.getChart("MyChart");
    myChart.delete();
}
//gavdcodeend 35

//gavdcodebegin 36
function main(workbook: ExcelScript.Workbook) {
    let wsByName = workbook.getWorksheet("MyWorksheet");

    console.log(wsByName.getCharts());
}
//gavdcodeend 36
