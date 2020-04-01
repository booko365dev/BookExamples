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

    function WriteSomeText() {
        Office.context.document.setSelectedDataAsync("To Whom It May Concern",
            function (asyncResult) {
                var error = asyncResult.error;
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.log(error.name + ": " + error.message)
                }
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
})()
//gavdcodeend 02
