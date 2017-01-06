// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

/* 
   This file provides the functionality for the page that opens in the popup. 
*/

(function () {
    "use strict";

    // Initialize the OAuth.io app with its ID.
    OAuth.initialize('{OAuth.io Public Key}');
    try {
        // Tell the OAuth.io app which provider to call and where to send the 
        // provider's ticket.
        var provider = localStorage.getItem("provider");
        OAuth.redirect(provider, 'https://localhost:3000/popupRedirect.html');
    }
    catch(err) {
        console.log(err.message);
    }
}());    