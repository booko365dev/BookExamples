//gavdcodebegin 02
(function () {
    "use strict"

    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#btnShowText').text("Copy selected text")
            $('#btnShowText').click(ShowText)
                
            WriteSomeData()
        });
    };

    function WriteSomeData() {
        var cellValues = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ]

        Excel.run(function (myContext) {
            var mySheet = myContext.workbook.worksheets.getActiveWorksheet()
            mySheet.getRange("B3:D5").values = cellValues

            return myContext.sync()
        })
            .catch(CathMyError)
    }

    function ShowText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (myResult) {
                if (myResult.status === Office.AsyncResultStatus.Succeeded) {
                    WriteToBox("The text is: '" + myResult.value + "'")
                } else {
                    WriteToBox("Error: " + myResult.error.message)
                }
            })
    }

    function WriteToBox(TextToWrite) {
        document.getElementById('myTextArea').innerText += TextToWrite
    }

    function CathMyError(error) {
        showNotification("Error:", error)
        console.log("Error: " + error)
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo))
        }
    }
})()
//gavdcodeend 02
