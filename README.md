# Microsoft Graph Connect Sample for AngularJS (REST)

## Table of contents

* [Introduction](#introduction)
* [Prerequisites](#prerequisites)
* [Register the application](#register-the-application)
* [Build and run the sample](#build-and-run-the-sample)
* [Additional resources](#additional-resources)

## Introduction

This sample shows how to connect an AngularJS app to a Microsoft work or school (Azure Active Directory) or personal (Microsoft) account  using the Microsoft Graph API to send an email. In addition, the sample uses the Office Fabric UI for styling and formatting the user experience.  We also have an [Angular connect sample](https://github.com/microsoftgraph/angular-connect-sample) that uses that [Microsoft Graph JavaScript SDK](https://github.com/microsoftgraph/msgraph-sdk-javascript).

![Microsoft Graph Connect sample screenshot](./README_assets/screenshot.png)

This sample uses the [Microsoft Authentication Library Preview for JavaScript (msal.js)](https://github.com/AzureAD/microsoft-authentication-library-for-js) to get an access token.

## Prerequisites

To use the Microsoft Graph Connect sample for AngularJS, you need the following:
* [Node.js](https://nodejs.org/). Node is required to run the sample on a development server and to install dependencies. 

* Either a [Microsoft account](https://www.outlook.com) or [Office 365 for business account](https://msdn.microsoft.com/en-us/office/office365/howto/setup-development-environment#bk_Office365Account)

## Register the application

1. Sign into the [App Registration Portal](https://apps.dev.microsoft.com/) using either your personal or work or school account.

2. Choose **Add an app**.

3. Enter a name for the app, and choose **Create application**. 
	
   The registration page displays, listing the properties of your app.

4. Copy the Application Id. This is the unique identifier for your app. 

5. Under **Platforms**, choose **Add Platform**.

6. Choose **Web**.

7. Make sure the **Allow Implicit Flow** check box is selected, and enter *http://localhost:8080* as the Redirect URI. 

8. Choose **Save**.


## Build and run the sample

1. Using your favorite IDE, open **config.js** in *public/scripts*.

2. Replace the **clientID** placeholder value with the application ID of your registered Azure application.

3. In a command prompt, run the following command in the root directory. This installs project dependencies.

  ```
npm install
  ```
  
4. Run `npm start` to start the development server.

5. Navigate to `http://localhost:8080` in your web browser.

6. Choose the **Connect** button.

7. Sign in with your personal or work or school account and grant the requested permissions.

## Additional resources

- [Other Microsoft Graph Connect samples](https://github.com/MicrosoftGraph?utf8=%E2%9C%93&query=-Connect)
- [Microsoft Graph](http://graph.microsoft.io)





