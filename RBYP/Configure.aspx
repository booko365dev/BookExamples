<!--gavdcodebegin 05-->
<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Configure.aspx.cs" 
    Inherits="RBYP.Configure" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script src="https://statics.teams.microsoft.com/sdk/v1.0/js/MicrosoftTeams.min.js" 
        type="text/javascript"></script>
    <script src="GenerateAppScripts.js" type="text/javascript"></script>
    <link rel="stylesheet" href="GenerateThemes.css" type="text/css" /> 
</head>
<body class="theme-light">
    <form id="form1" runat="server">
        <label for="tabChoice">Select the tab you would like to see: </label>
        <select id="tabChoice" name="tabChoice">
            <option value="" selected="selected">(Select a tab)</option>
            <option value="GenerateGuid.aspx">Generate GUID</option>
            <option value="GenerateRandomString.aspx">Generate Random String</option>
        </select>
    </form>
</body>
</html>
<!--gavdcodeend 05-->

