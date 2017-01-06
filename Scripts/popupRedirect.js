// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

/* This file provides the functionality for the page that opens in the popup 
   after the user has logged in with a provider 
*/

(function () {
    "use strict";

    // Office.initialize must be called on every page where Office JavaScript is 
    // called. Other initialization code should go inside it.
    Office.initialize = function() {
        
        // Loading the OAuth.io library with a <script> tag on the popupRedirect
        // page blocks further page rendering. So we must load it programmatically
        // after the page has completely loaded.
        $(document).ready(function () {  
            try {
                jQuery.ajax({
                    url: 'Scripts/OAuth.io/oauth.js',
                    dataType: 'script',
                    success: getProviderTicket,
                    async: true
                }); 
            }
            catch(err) {
                console.log(err.message);
            }
        });

    }

    function getProviderTicket() {

        // Tell the OAuth.io app what provider to call.
        var provider = localStorage.getItem("provider");
        OAuth.callback(provider)
        .done(function (result) {

            // Create the outcome message and send it to the task pane.
            var messageObject = {outcome: "success", provider: provider, ticket: result};            
            var jsonMessage = JSON.stringify(messageObject);

            // Tell the task pane about the outcome.
            Office.context.ui.messageParent(jsonMessage);
        })
        .fail(function (err) {

            // Create the outcome message and send it to the task pane.
            var messageObject = {outcome: "failure", error: err.message};            
            var jsonMessage = JSON.stringify(messageObject);

            // Tell the task pane about the outcome.
            Office.context.ui.messageParent(jsonMessage); 
        });
    }
}());    
 