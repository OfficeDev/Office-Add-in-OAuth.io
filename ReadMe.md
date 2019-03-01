# [ARCHIVED] Office Add-in that uses the OAuth.io Service to get Authorization to External Services

> **Note:** This repo is archived and no longer actively maintained. Security vulnerabilities may exist in the project, or its dependencies. If you plan to reuse or run any code from this repo, be sure to perform appropriate security checks on the code or dependencies first. Do not use this project as the starting point of a production Office Add-in. Always start your production code by using the Office/SharePoint development workload in Visual Studio, or the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), and follow security best practices as you develop the add-in.

The OAuth.io service simplifies the process of getting OAuth 2.0 access tokens from popular online services such as Facebook and Google. This sample shows how to use it in an Office add-in. 

## Table of Contents
* [Change History](#change-history)
* [Prerequisites](#prerequisites)
* [Configure the project](#configure-the-project)
* [Download the OAuth.io JavaScript Library](#download-the-oauth-io-javascript-library)
* [Create a Google project](#create-a-google-project)
* [Create a Facebook app](#create-a-facebook-app)
* [Create an OAuth.io app and configure it to use Google and Facebook](#create-an-OAuth-io-app-and-configure-it-to-use-google-and-facebook)
* [Add the public key to the sample code](#add-the-public-key-to-the-sample-code)
* [Deploy the add-in](#deploy-the-add-in)
* [Run the project](#run-the-project)
* [Start the add-in](#start-the-add-in)
* [Test the add-in](#test-the-add-in)
* [Questions and comments](#questions-and-comments)
* [Additional resources](#additional-resources)

## Change History

January 5, 2017:

* Initial version.

## Prerequisites

* An account with [OAuth.io](https://oauth.io/home)
* Developer accounts with [Google](https://developers.google.com/) and [Facebook](https://developers.facebook.com/).
* Word 2016 for Windows, build 16.0.6727.1000 or later.
* [Node and npm](https://nodejs.org/en/) The project is configured to use npm as both a package manager and a task runner. It is also configured to use the Lite Server as the web server that will host the add-in during development, so you can have the add-in up and running quickly. You are welcome to use another task runner or web server.
* [Git Bash](https://git-scm.com/downloads) (Or another git client.)

## Configure the project

In the folder where you want to put the project, run the following commands in the git bash shell:

1. ```git clone {URL of this repo}``` to clone this repo to your local machine.
2. ```npm install``` to install all of the dependencies itemized in the package.json file.
3. ```bash gen-cert.sh``` to create the certificate needed to run this sample. 

Set the certificate to be a trusted root authority. On a Windows machine, these are the steps:

1. In the repo folder on your local computer, double-click ca.crt, and select **Install Certificate**. 
2. Select **Local Machine** and select **Next** to continue. 
3. Select **Place all certificates in the following store** and then select **Browse**.
4. Select **Trusted Root Certification Authorities** and then select **OK**. 
5. Select **Next** and then **Finish**. 

## Download the OAuth.io JavaScript Library

2. In the project, create a subfolder called OAuth.io under the Scripts folder.
3. Download the OAuth.io JavaScript library from [oauth-js](https://github.com/oauth-io/oauth-js).
2. From the \dist folder of the library, copy either oauth.js or oauth.min.js to the \Scripts\OAuth.io folder of the project. 
3. If you chose oauth.min.js in the preceding step, open the file \Scripts\popupRedirect.js and change the line `url: 'Scripts/OAuth.io/oauth.js',` to `url: 'Scripts/OAuth.io/oauth.min.js',`.

## Create a Google project

1. On your [Google developer console](https://console.developers.google.com/apis/dashboard), create a project named **Office Add-in**.
2. Enable the Google Plus API for the project. See Google's help for details.
3. Create app credentials for the project and choose the **OAuth client ID** as the credential type, and **Web application** as the application type. Use "Office-Add-in-OAuth.io" as the client **Name** (and the product name, if you are prompted to specify one).
4. Set the **Authorized redirect URIs** to ```https://oauth.io/auth```.
5. Make a copy of the client secret and client ID that it is assigned.

## Create a Facebook app

1. On your [Facebook developer dashboard](https://developers.facebook.com/), create an app named **Office Add-in**. See Facebook help for details.
3. Make a copy of the client secret and client ID that it is assigned. 
4. Specify ```oauth.io``` in the **App Domains** box and specify ```https://oauth.io/``` as the **Site URL**. 

## Create an OAuth.io app and configure it to use Google and Facebook

1. In your OAuth.io dashboard create an app and make a note of the **public key** that OAuth.io assigns to it.
3. If it is not already there, add `localhost` to the whitelisted URLs for the app.
2. Add **Google Plus** as a provider (also called an "Integrated API") to the app. Fill in the client ID and secret you got when you created the Google project, and then specify `profile` as the **scope**. If the configuration of the provider -- Integrated API -- doesn't save in IE, try reopening the OAuth.io dashboard in another browser.)
3. Repeat the previous step for **Facebook**, but specify `user_about_me` as the **scope**.

## Add the public key to the sample code

The API `OAuth.initialize` is called in the file popup.js. Find the line and pass the pubic key you got in the previous section as a string to this function.

## Deploy the add-in

Now you need to let Microsoft Word know where to find the add-in.

1. Create a network share, or [share a folder to the network](https://technet.microsoft.com/en-us/library/cc770880.aspx).
2. Place a copy of the Office-Add-in-OAuth.io.xml manifest file, from the root of the project, into the shared folder.
3. Launch Word and open a document.
4. Choose the **File** tab, and then choose **Options**.
5. Choose **Trust Center**, and then choose the **Trust Center Settings** button.
6. Choose **Trusted Add-ins Catalogs**.
7. In the **Catalog Url** field, enter the network path to the folder share that contains Office-Add-in-OAuth.io.xml, and then choose **Add Catalog**.
8. Select the **Show in Menu** check box, and then choose **OK**.
9. A message is displayed to inform you that your settings will be applied the next time you start Microsoft Office. Close Word.

## Run the project

1. Open a node command window in the folder of the project and run ```npm start``` to start the web service. Leave the command window open.
2. Open Internet Explorer or Edge and enter ```https://localhost:3000``` in the address box. If you do not receive any warnings about the certificate, close the browser and continue with the section below titled **Start the add-in**. If you do receive a warning that the certificate is not trusted, continue with the following steps:
3. The browser gives you a link to open the page despite the warning. Open it.
4. After the page opens, there will be a red certificate error in the address bar. Double click the error.
5. Select **View Certificate**.
5. Select **Install Certificate**.
4. Select **Local Machine** and select **Next** to continue. 
3. Select **Place all certificates in the following store** and then select **Browse**.
4. Select **Trusted Root Certification Authorities** and then select **OK**. 
5. Select **Next** and then **Finish**.
6. Close the browser.

## Start the add-in

1. Restart Word and open a Word document.
2. On the **Insert** tab in Word 2016, choose **My Add-ins**.
3. Select the **SHARED FOLDER** tab.
4. Choose **Auth with OAuthIO**, and then select **OK**.
5. If add-in commands are supported by your version of Word, the UI will inform you that the add-in was loaded.
6. On the Home ribbon is a new group called **OAuth.io** with a button labeled **Show** and an icon. Click that button to open the add-in.

 > Note: The add-in will load in a task pane if add-in commands are not supported by your version of Word.

## Test the add-in

1. Click the **Get Facebook Name** button.
2. A popup will open and you will be prompted to sign in with Facebook (unless you already are).
3. After you sign in, your Facebook user name is inserted into the Word document.
4. Repeat the above steps with the **Get Google Name** button.

## Questions and comments

We'd love to get your feedback about this sample. You can send your feedback to us in the *Issues* section of this repository.

Questions about Microsoft Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). If your question is about the Office JavaScript APIs, make sure that your questions are tagged with [office-js] and [API].

## Additional resources

* [Office add-in documentation](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Office Dev Center](http://dev.office.com/)
* More Office Add-in samples at [OfficeDev on Github](https://github.com/officedev)

## Copyright
Copyright (c) 2017 Microsoft Corporation. All rights reserved.

