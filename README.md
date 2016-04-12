# QuickBooks Excel Add-in with .Net

Microsoft Office Excel Add-in that integrates with QuickBooks.

## Table of Contents
* [Change History](#change-history)
* [Prerequisites](#prerequisites)
* [Azure client application registration](#azure-client-application-registration)
* [Configure the project](#configure-the-project)
* [Run the project](#run-the-project)
* [Understand the code](#understand-the-code)
* [Connect to Office 365](#connect-to-office-365)
* [Questions and comments](#questions-and-comments)
* [Additional resources](#additional-resources)

## Change History
//*Date and description each time significant change is made. Example below:*//

June 18, 2025:
* Created first version of readme template.

## Prerequisites

* A QuickBooks developer account
* Visual Studio 2015
* Office Developer Tools for Visual Studio

## Configure the project
//*Document how to add client ids, redirects, command line configuration commands, or whatever other steps to correctly open and get the sample ready to run.*//
Go to https://developer.intuit.com/ and sign up for a developer account.
In the upper right hand corner, choose My Apps and select an app or click Create new app. 
Once the app is selected, choose **Development** | **Keys** to copy the **App Token**, **OAuth Consumer Key**, and **OAuth Consumer Secret**.
Download or clone the project.
Open the solution file QbAdd-inDotNet.sln in Visual Studio.
In the Web.config file, insert the values for `ConsumerKey` and `ConsumerSecret`, like this.

```
<appSettings>
    <!-- QuickBooks Settings -->
    <add key="ConsumerKey" value="<your OAuth Consumer Key>" />
    <add key="ConsumerSecret" value="<your OAuth Consumer Secret>" />
    <add key="RealmId" value="123145709687887" />
    <add key="OauthLink" value="https://oauth.intuit.com/oauth/v1" />
    <add key="AuthorizeUrl" value="https://workplace.intuit.com/Connect/Begin" />
    <add key="RequestTokenUrl" value="https://oauth.intuit.com/oauth/v1/get_request_token" />
    <add key="AccessTokenUrl" value="https://oauth.intuit.com/oauth/v1/get_access_token" />
    <add key="ServiceContext.BaseUrl.Qbo" value="https://sandbox-quickbooks.api.intuit.com/" />
    <add key="DeepLink" value="sandbox.qbo.intuit.com" />
  </appSettings>
```

## Run the project
//*Document how to use the sample once it is running. How to log in, and get it to do something interesting.*//
## Understand the code
//*Optional. Provide information about what code files the developer should study to understand how they can use it in their own code, or how the sample works. 
Note that you can use deep links to the code samples or specific line numbers using this [linking mechanism](http://stackoverflow.com/questions/23821235/how-to-link-to-specific-line-number-on-github). *//
## Questions and comments
//*Use the following boilerplate. Fill in your sample name below, and the correct link to your issues section*//
We'd love to get your feedback on the *YOUR SAMPLE NAME* sample. You can send your feedback to us in the *Issues* section of this repository. //br//
Questions about Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Make sure that your questions are tagged with [Office365] and [API].
## Additional resources
//*Provide links to other samples or relevant documentation. Some common ones listed below.*//

* [Office 365 APIs documentation](http://msdn.microsoft.com/office/office365/howto/platform-development-overview)
* [Microsoft Office 365 API Tools](https://visualstudiogallery.msdn.microsoft.com/a15b85e6-69a7-4fdf-adda-a38066bb5155)
* [Office Dev Center](http://dev.office.com/)
* [Office 365 APIs starter projects and code samples](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)
## Copyright
Copyright (c) 2016 Microsoft. All rights reserved.