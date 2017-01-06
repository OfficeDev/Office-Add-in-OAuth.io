// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/* 
    This file provides the functionality for the main task pane page. 
*/

/// <reference path="./App.js" />

(function () {
    "use strict";

    var dialog;

    // Office.initialize must be assigned to a function on every page in which there is a 
	// call to something in the Office JS library. Put initialization code in it.
    Office.initialize = function () {

        $(document).ready(function () {
            // Initialize error message banner.
            app.initialize();

            // Enable the buttons
            $(".popupButton").prop("disabled", false);

            $("#facebookButton").click(function () {
                showLoginPopup('facebook');
            });

            $("#googleButton").click(function () {
                showLoginPopup('google_plus');
            });
        });       
    };

    // This handler responds to the success or failure message that the pop-up dialog receives from the identity provider
    // and access token provider.
    function processMessage(arg) {
        var messageFromPopupDialog = JSON.parse(arg.message);

        if (messageFromPopupDialog.outcome === "success") {
            // The provider's ticket has been stored.
            dialog.close();
            getUserData(messageFromPopupDialog.provider, messageFromPopupDialog.ticket);
        } else {
            // Something went wrong with authentication or the authorization of the web application,
            // either with OAuth.io or with the provider.
            dialog.close();
            app.showNotification("User authentication and application authorization", "Unable to successfully authenticate user or authorize application: " + messageFromPopupDialog.error);
        }
    }

    // Use the Office dialog API to open a pop-up and display the sign-in page for the identity provider.
    function showLoginPopup(provider) {
        
        // Provider name needs to be shared between task pane and popup window.
        localStorage.removeItem("provider");
        localStorage.setItem("provider", provider);

        // Create the popup URL and open it.        
        var fullUrl = location.protocol + '//' + location.hostname + (location.port ? ':' + location.port : '') + '/popup.html';

        // height and width are percentages of the size of the screen.
        Office.context.ui.displayDialogAsync(fullUrl,
                {height: 65, width: 40, requireHTTPS: true}, 
                function (result) {
                    dialog = result.value;
                    dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
                });
    }

    function getUserData(provider, oAuthIOTicket) {
            try {
                var endpoint = getProviderEndPoint(provider);

                // Call the provider and insert the returned data into the document.
                $.get(endpoint + oAuthIOTicket.access_token,
                    function (data) { insertDataInDocument(provider, data); }
                );
            }
            catch(err) {
                app.showNotification(err.message);
            }
    }

    function insertDataInDocument(provider, data) {

        var parsedData = getParsedProviderData(provider, data);
        
        Word.run(function (context) {

            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue commands to insert text into the end of the Word document body.
            body.insertText(parsedData, "End");

            // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
            return context.sync();
        })
        .catch(this.errorHandler);
    }

    function getParsedProviderData(provider, data) {

            // Each provider has its own structure for the data it returns.
            switch(provider) {
                case "facebook":
                    return data.name;
                case "google_plus":
                    return data.displayName;
                default:
                    return "Error: Cannot find provider named: " + provider;
            }
    }

    function getProviderEndPoint(provider) {

            // Each provider has its own REST API URL pattern.
            switch(provider) {
                case "facebook":
                    return "https://graph.facebook.com/v2.7/me?access_token=";
                case "google_plus":
                    return "https://www.googleapis.com/plus/v1/people/me?access_token="
                default:
                    return "Error: Cannot find provider named: " + provider;
            }
    }
}());
