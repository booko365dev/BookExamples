<!--gavdcodebegin 006-->
<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="auth.aspx.cs" Inherits="Test_TeamsTabGraphToolkit.auth" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    
    <script src="https://unpkg.com/@microsoft/teams-js/dist/MicrosoftTeams.min.js" 
        crossorigin="anonymous"></script>
    <script src="https://unpkg.com/@microsoft/mgt/dist/bundle/mgt-loader.js"></script>

</head>
<body>

    <script>
      mgt.TeamsProvider.handleAuth();
    </script>

</body>
</html>
<!--gavdcodeend 006-->
