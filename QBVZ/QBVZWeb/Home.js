
(function () {
    "use strict"

//gavdcodebegin 02
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#btnGetCity').text("Get City")
            $('#btnGetCity').click(GetCity)
        })
    }
//gavdcodeend 02

//gavdcodebegin 03
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
//gavdcodeend 03

//gavdcodebegin 04
    async function CallGetAsync(urlToCall) {
        var urlResponse = await fetch(urlToCall)
        var responseData = await urlResponse.json()
        return responseData
    }
//gavdcodeend 04

})()
