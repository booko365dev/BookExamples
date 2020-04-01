//gavdcodebegin 02
(function () {
    "use strict";

    Office.initialize = function (reason) {
    $(document).ready(function () {
        $('#btnCallRest').text("Call REST Service")
        $('#btnCallRest').click(CallRest)
    });
    };

    function CallRest() {
        var ServiceUrl = "https://reqres.in/api/users/2"

        CallGetAsync(ServiceUrl).then(
            function (resultFromService) {
                var resultString = resultFromService.data.first_name + " " +
                    resultFromService.data.last_name + " - " +
                    resultFromService.data.email
                $('#resultRest').text(resultString);
            })
    }

    async function CallGetAsync(urlToCall) {
        var urlResponse = await fetch(urlToCall)
        var responseData = await urlResponse.json()
        return responseData
    }
})();
//gavdcodeend 02