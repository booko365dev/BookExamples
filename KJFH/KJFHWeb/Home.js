
(function () {
    "use strict"

    //gavdcodebegin 002
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#btnGetPlaceHolder').text("Get PlaceHolder")
            $('#btnGetPlaceHolder').click(GetPlaceHolder)
        })
    }
    //gavdcodeend 002

    //gavdcodebegin 003
    function GetPlaceHolder() {
        var ServiceUrl = "https://jsonplaceholder.typicode.com/posts/11"

        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (resultReadPlaceHolder) {
                if (resultReadPlaceHolder.status == Office.AsyncResultStatus.Succeeded) {
                    CallGetAsync(ServiceUrl + resultReadPlaceHolder.value.trim()).then(
                        function (resultPlaceHolder) {

                            var values = [
                                [resultPlaceHolder.id],
                                [resultPlaceHolder.title],
                                [resultPlaceHolder.body]
                            ]

                            Excel.run(function (myContext) {
                                var mySheet = myContext.workbook.worksheets.
                                                                    getActiveWorksheet()
                                var myRange = mySheet.getRange("C3:C5")
                                myRange.values = values
                                myRange.format.autofitColumns()

                                return myContext.sync()
                            })
                                .catch(handleMyErrors)
                        })
                }
                else {
                    alert(resultPlaceHolder.error.message)
                }
            }
        )
    }
    //gavdcodeend 003

    //gavdcodebegin 004
    async function CallGetAsync(urlToCall) {
        var urlResponse = await fetch(urlToCall)
        var responseData = await urlResponse.json()
        return responseData
    }

    function handleMyErrors(error) {
        console.log("Error: " + error)
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo))
        }
    }
    //gavdcodeend 004
})()
