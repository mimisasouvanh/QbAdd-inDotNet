﻿<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Close.aspx.cs" Inherits="QbAdd_inDotNetWeb.Close" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Close</title>
    <script src="//appsforoffice.microsoft.com/lib/beta/hosted/office.js" type="text/javascript"></script>
    <script src="//ajax.aspnetcdn.com/ajax/jQuery/jquery-1.9.1.min.js" type="text/javascript"></script>

    <script>
        Office.initialize = function (reason) {
            var message = {};
            message.token = "<%=HttpContext.Current.Session["accessToken"].ToString()%>";
            message.secret = "<%=HttpContext.Current.Session["accessTokenSecret"].ToString()%>";
            Office.context.ui.messageParent(JSON.stringify(message));
        }

    </script>
</head>
<body>
    <form id="form1" runat="server">

    </form>
</body>
</html>
