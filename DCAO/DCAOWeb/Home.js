
(function () {
    "use strict"

    //gavdcodebegin 02
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#btnGetDateTime').text("Get DateTime")
            $('#btnGetDateTime').click(GetDateTime)
        });
    };
    //gavdcodeend 02

    //gavdcodebegin 03
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
    //gavdcodeend 03

    //gavdcodebegin 04
    async function CallGetAsync(urlToCall) {
        var urlResponse = await fetch(urlToCall)
        var responseData = await urlResponse.json()
        return responseData
    }
    //gavdcodeend 04
})();
