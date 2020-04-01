<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="LPPBWeb.Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<!--gavdcodebegin 01-->
<body>
    <form id="form1" runat="server">
    <div>

        <h3>User</h3>
        <asp:Label runat="server" ID="lblUser" />

        <h3>Other Users</h3>
        <asp:ListView ID="lstOtherUsers" runat="server">
            <ItemTemplate>
                <asp:Label ID="oneUser" runat="server"
                    Text="<%# Container.DataItem.ToString() %>">
                </asp:Label><br />
            </ItemTemplate>
        </asp:ListView>

        <h3>Lists</h3>
        <asp:ListView ID="lstLists" runat="server">
            <ItemTemplate>
                <asp:Label ID="oneList" runat="server"
                    Text="<%# Container.DataItem.ToString() %>">
                </asp:Label><br />
            </ItemTemplate>
        </asp:ListView>

    </div>
    </form>
</body>
<!--gavdcodeend 01-->
</html>
