<!--gavdcodebegin 003-->
<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="GenerateRandomString.aspx.cs" 
    Inherits="RBYP.GenerateRandomString" %>

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
        <div class="font-semibold font-title"><h2>Generate a random string</h2></div>
        <div>
            <p>
                <asp:Button ID="btnGenerateRandomString" runat="server" Text="Generate" 
                    OnClick="btnGenerateRandomString_Click" />
            </p>
            <p class="surface">
                <asp:Label ID="lblRandomString" runat="server" Text=""></asp:Label>
            </p>
        </div>
    </form>
</body>
</html>
<!--gavdcodeend 003-->
