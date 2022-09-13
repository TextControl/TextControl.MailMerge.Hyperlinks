<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="index.aspx.cs" Inherits="TextControl.MailMerge.Hyperlinks.index" %>

<%@ Register assembly="TXDocumentServer, Version=31.0.1700.500, Culture=neutral, PublicKeyToken=6b83fe9a75cfb638" namespace="TXTextControl.DocumentServer" tagprefix="cc1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Button ID="Button1" runat="server" Text="Merge" OnClick="Button1_Click" />
    
    </div>
    </form>
</body>
</html>
