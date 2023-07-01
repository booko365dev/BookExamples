
(function () {
    "use strict"

    //gavdcodebegin 002
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#btnGetDateTime').text("Get DateTime")
            $('#btnGetDateTime').click(GetDateTime)
        });
    };
    //gavdcodeend 002

    //gavdcodebegin 003
    function GetDateTime() {
        var ServiceUrl = "http://date.jsontest.com"

        CallGetAsync(ServiceUrl).then(
            function (resultDateTime) {
                WriteToSlide(
                    resultDateTime.date + " - " +
                    resultDateTime.time
                )
            })
    }

    function WriteToSlide(DateAndTime) {
        Office.context.document.setSelectedDataAsync(DateAndTime,
            function (asyncResult) {
                var error = asyncResult.error
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.log(error.name + ": " + error.message)
                }
            })
    }
    //gavdcodeend 003

    //gavdcodebegin 004
    async function CallGetAsync(urlToCall) {
        var urlResponse = await fetch(urlToCall)
        var responseData = await urlResponse.json()
        return responseData
    }
    //gavdcodeend 004
})();
