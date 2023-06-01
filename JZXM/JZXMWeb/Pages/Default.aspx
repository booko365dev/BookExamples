<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="JZXMWeb.Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<!--gavdcodebegin 001-->
<head runat="server">

     <title>Chrome control host page</title>
     <script 
         type="text/javascript" 
         src="../Scripts/jquery-1.9.1.min.js">
     </script>      
     <script type="text/javascript">
         var hostUrl =
             decodeURIComponent(
                 getQueryStringParameter("SPHostUrl")
             );

         var scriptBase = hostUrl + "/_layouts/15/";
         $.getScript(scriptBase + "SP.UI.Controls.js")

         function getQueryStringParameter(paramToRetrieve) {
             var myParams = document.URL.split("?")[1].split("&amp;");
             for (var i = 0; i < myParams.length; i = i + 1) {
                 var oneParam = myParams[i].split("=");
                 if (oneParam[0] == paramToRetrieve) {
                     var myReturn = oneParam[1].replace("&SPLanguage", "");
                     return myReturn;
                 }
             }
         }
     </script>

</head>
<!--gavdcodeend 001-->
<!--gavdcodebegin 002-->
<body>

     <div 
         id="chrome_ctrl_container"
         data-ms-control="SP.UI.Controls.Navigation"  
         data-ms-options=
             '{  
                 "appHelpPageUrl" : "Help.html",
                 "appTitle" : "Chrome control in my Add-in",
                 "settingsLinks" : [
                     {
                         "linkUrl" : "Account.html",
                         "displayName" : "Account settings"
                     },
                     {
                         "linkUrl" : "Contact.html",
                         "displayName" : "Contact us"
                     }
                 ]
              }'>
     </div>

    <form id="form1" runat="server">
    <div>
    
    </div>
    </form>
</body>
<!--gavdcodeend 002-->
</html>
