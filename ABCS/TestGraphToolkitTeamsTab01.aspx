<!--gavdcodebegin 005-->
<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="TestGraphToolkitTeamsTab01.aspx.cs" Inherits="Test_TeamsTabGraphToolkit.TestGraphToolkitTeamsTab01" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script src="GenerateAppScripts.js" type="text/javascript"></script> 
    <link rel="stylesheet" href="GenerateThemes.css" type="text/css" />

    <script src="https://unpkg.com/@microsoft/teams-js/dist/MicrosoftTeams.min.js" 
        crossorigin="anonymous"></script>
    <script src="https://unpkg.com/@microsoft/mgt/dist/bundle/mgt-loader.js"></script>

</head>
<body class="theme-light">

    <mgt-teams-provider
      client-id="3c99d9e2-4db0-4c90-888c-fa6b4a9fef1f"
      auth-popup-url="/auth.aspx">
    </mgt-teams-provider>

    <form id="form1" runat="server">
        <div class="surface font-semibold font-title"><h2>Using Graph Toolkit</h2></div>

        <mgt-login></mgt-login>
        <mgt-todo></mgt-todo>

    </form>
</body>
</html>
<!--gavdcodeend 005-->
