﻿<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Home.aspx.cs" Inherits="QbAdd_inDotNetWeb.Home" %>

<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title>Excel Addin with Commands Sample</title>

    <script src="/Scripts/jquery-1.9.1.js" type="text/javascript"></script>
    <script src="/Scripts/FabricUI/MessageBanner.js" type="text/javascript"></script>
    <script src="//appsforoffice.microsoft.com/lib/beta/hosted/office.js" type="text/javascript"></script>
    <script src="//appcenter.intuit.com/Content/IA/intuit.ipp.anywhere-1.3.1.js" type="text/javascript"></script>

    <!-- To enable offline debugging using a local reference to Office.js, use:                        -->
    <!-- <script src="/Scripts/Office/MicrosoftAjax.js" type="text/javascript"></script>  -->
    <!-- <script src="/Scripts/Office/1.1/office.js" type="text/javascript"></script>  -->

    <link href="Home.css" rel="stylesheet" type="text/css" />
    <script src="Home.js" type="text/javascript"></script>

    <!-- For the Office UI Fabric, go to http://aka.ms/office-ui-fabric to learn more. -->
    <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/2.1.0/fabric.min.css">
    <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/2.1.0/fabric.components.min.css">

    <!-- To enable the offline use of Office UI Fabric, use: -->
    <!-- link rel="stylesheet" href="/Content/fabric.min.css" -->
    <!-- link rel="stylesheet" href="/Content/fabric.components.min.css" -->
</head>
<body>
    <div class="padding">
        <div id="welcomePanel" class="ms-u-fadeIn200" style="padding:50px">
            <h1 class="first-run-title ms-font-l ms-fontWeight-light">Welcome!</h1>
            <br />
            <h1 class="first-run-title ms-font-l ms-fontWeight-light">This add-in connects Excel with your Quickbooks Online account, pleasse sign-in to get started.</h1><br />
            <a id="btnSignIn" class="intuitPlatformConnectButton" style="cursor:pointer">Connect with QuickBooks</a>
        </div>

        <div id="actionsPanel" class="ms-u-fadeIn200" style="display: none">
            <h1 class="first-run-title ms-font-l ms-fontWeight-light">Please select an action below:</h1>
            <br />

            <button id="btnGetAccounts" class="ms-Button ms-Button--command"><span class="ms-font-xxl ms-Button-icon"><i class="ms-Icon ms-Icon--people"></i></span><span class="ms-Button-label">Get Accounts</span></button><br />
            <br />
            <button id="btnGetPurchases" class="ms-Button ms-Button--command"><span class="ms-font-xxl ms-Button-icon"><i class="ms-Icon ms-Icon--bag"></i></span><span class="ms-Button-label">Get Expenses</span></button><br />
            <br />
            <button id="btnCreateReport" class="ms-Button ms-Button--command"><span class="ms-font-xxl ms-Button-icon"><i class="ms-Icon ms-Icon--graph"></i></span><span class="ms-Button-label">Spending Report</span></button><br />
            <br />
                    <div style="position: absolute; bottom: 50px;">
            <button id="btnSignOut" class="ms-Button ms-Button--command"><span class="ms-font-xxl ms-Button-icon"><i class="ms-Icon ms-Icon--oofReply"></i></span><span class="ms-Button-label">Sign Out</span></button>
        </div>
        </div>


    </div>
</body>
</html>

