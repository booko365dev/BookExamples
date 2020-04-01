//gavdcodebegin 02
(function () {
    "use strict"

    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#btnShowText').text("Copy selected text")
            $('#btnShowText').click(ShowText)

            WriteSomeText()
        })
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

    function WriteSomeText() {
        Word.run(function (context) {
            var docBody = context.document.body
            docBody.clear()
            docBody.insertText(
                "To Whom It May Concern",
                Word.InsertLocation.end)

            return context.sync()
        })
            .catch(CathMyError)
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
