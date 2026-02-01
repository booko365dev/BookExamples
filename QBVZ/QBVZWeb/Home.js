
(function () {
    "use strict"

//gavdcodebegin 002
// Solution Deprecated by Microsoft
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#btnGetCity').text("Get City")
            $('#btnGetCity').click(GetCity)
        })
    }
//gavdcodeend 002

//gavdcodebegin 003
    // Solution Deprecated by Microsoft
    function GetCity() {
        var ServiceUrl = "http://ziptasticapi.com/"

        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (resultReadZip) {
                if (resultReadZip.status == Office.AsyncResultStatus.Succeeded) {
                    CallGetAsync(ServiceUrl + resultReadZip.value.trim()).then(
                        function (resultCity) {
                            var cityString = resultCity.city + " - " +
                                resultCity.state + " - " +
                                resultCity.country

                            Office.context.document.setSelectedDataAsync(cityString,
                                function (resultInsertCity) {
                                    if (resultInsertCity.status ==
                                        Office.AsyncResultStatus.Failed) {
                                        alert(resultInsertCity.error.message);
                                    }
                                })
                        })
                }
                else {
                    alert(resultReadZip.error.message)
                }
            }
        )
    }
//gavdcodeend 003

//gavdcodebegin 004
    // Solution Deprecated by Microsoft
    async function CallGetAsync(urlToCall) {
        var urlResponse = await fetch(urlToCall)
        var responseData = await urlResponse.json()
        return responseData
    }
//gavdcodeend 004

})()
