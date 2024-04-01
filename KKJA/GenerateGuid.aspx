<!--gavdcodebegin 001-->
<!-- ATTENTION: Replaced by LZWD -->
<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="GenerateGuid.aspx.cs" 
    Inherits="KKJA.GenerateGuid" %>

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
        <div class="surface font-semibold font-title"><h2>Generate a new GUID</h2></div>
        <div>
            <p>
                <asp:Button ID="btnGenerateGuid" runat="server" Text="Generate" 
                    OnClick="btnGenerateGuid_Click" />
            </p>
            <p class="surface">
                <asp:Label ID="lblNewGuid" runat="server" Text=""></asp:Label><br />
                <asp:Label ID="lblContextInfo" runat="server" Text=""></asp:Label>
            </p>
        </div>
    </form>
</body>
</html>
<!--gavdcodeend 001-->