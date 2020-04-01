<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="MyPHClientAppPart.aspx.cs" Inherits="XNYYWeb.Pages.MyPHClientAppPart" %>

<!DOCTYPE html>

<html>
<head>
    <title></title>
    <script type="text/javascript">
        // Set the style of the client web part page to be consistent with the host web.
        (function () {
            'use strict';

            var hostUrl = '';
            var link = document.createElement('link');
            link.setAttribute('rel', 'stylesheet');
            if (document.URL.indexOf('?') != -1) {
                var params = document.URL.split('?')[1].split('&');
                for (var i = 0; i < params.length; i++) {
                    var p = decodeURIComponent(params[i]);
                    if (/^SPHostUrl=/i.test(p)) {
                        hostUrl = p.split('=')[1];
                        link.setAttribute('href', hostUrl + '/_layouts/15/defaultcss.ashx');
                        break;
                    }
                }
            }
            if (hostUrl == '') {
                link.setAttribute('href', '/_layouts/15/1033/styles/themable/corev15.css');
            }
            document.head.appendChild(link);
        })();
    </script>
</head>
<!--gavdcodebegin 01-->
<body>
    <div id="content">
        <p>String property value: <span id="strProperty"></span></p>
    </div>

    <script lang="javascript">
        "use strict";

        var myParams = document.URL.split("?")[1].split("&amp;");
        var strProperty;

        for (var i = 0; i < myParams.length; i = i + 1) {
            var oneParam = myParams[i].split("=");
            if (oneParam[0] == "strProperty") {
                strProperty = decodeURIComponent(oneParam[1]);
            }
        }

        document.getElementById("strProperty").innerText = strProperty;
    </script>
</body>
<!--gavdcodeend 01-->
</html>
