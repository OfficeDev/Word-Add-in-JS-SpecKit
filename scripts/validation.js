/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// The initialize function must be run each time a new page is loaded
(function() {
    Office.initialize = function(reason) {
        //If you need to initialize something you can do so here.
    };
})();

/************************
 * Validation functions
 * We are calling these from the add-in commands on the ribbon. This file is
 * loaded when functionfile.html is loaded. functionfile.html is referenced in
 * the manifest.
 *************************/

/**
 * Get the array of words in the document by using the Word 1.3 splitTextRanges
 * API. Get the blacklist words kept in localStorage, and then compare the two
 * lists to determine if there are any blacklisted words in the document.
 * A search may be a better scenario for finding blacklist words. We're using
 * splitTextRanges for demo purposes.
 **/
function validateAgainstBlacklist() {

    // Get all of the words in the document.
    // Let's get the list of words.
    Word.run(function(context) {
        var body = context.document.body;

        // We'll get all the words in document. We're identifying
        // words by a space and paragraph delimiter. We may want to consider paging
        // through the results when we load the ranges into the words
        // variable.
        var delimiters = [' ', '\r', ',', ':', ';', '[', ']'];

        // Queue a command to get all the words in the document.
        // The words in the document are contained in a collection of Range objects.
        var range = body.getRange();
        var words = range.split(delimiters, true, true, true);

        // Queue a command to load the ranges that represent words.
        context.load(words, 'text');

        // Synchronize the document state by executing the queued command to
        // insert the template into the current document
        // and return a promise to indicate task completion.
        return context.sync().then(function() {
            var wordsText = [];

            // Extract the words from the Range objects.
            for (var i = 0; i < words.items.length; i++) {

                // This removes ranges that contain zero length strings.
                // This is caused when there are two delimiters adjacent to each other.
                if (words.items[i].text !== '') {
                    wordsText.push(words.items[i].text);
                }
            }

            // Get the blacklist words from local storage. The blacklist was added
            // in sample.js, getBlackList()
            var blacklist = JSON.parse(localStorage.getItem('badwordcache'));

            // Performs a case-sensitive comparison to provide the intersections of words
            // found in the document and in the blacklist.
            var arrayOfFoundBlackListWords = intersect_safe(wordsText, blacklist.badwords);

            var notifyPromise = notifyFoundWords(arrayOfFoundBlackListWords);
            return notifyPromise;
        });
    }).catch(function(error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

/**
 * Provided by atk on StackOverflow. Gives us the intersection of bad words.
 * stackoverflow.com/questions/1885557/simplest-code-for-array-intersection-in-javascript
 */
function intersect_safe(a, b) {

    a.sort();
    b.sort();
    var ai = 0, bi = 0;
    var result = [];

    while (ai < a.length && bi < b.length) {
        if (a[ai] < b[bi]) { ai++; }
        else if (a[ai] > b[bi]) { bi++; }
        else /* they're equal */ {
            result.push(a[ai]);
            ai++;
            bi++;
        }
    }
    return result;
}

/**
 * Add the words to local storage for access from the dialog.
 * Open the dialog to provide notification of found words.
 */
function notifyFoundWords(arrayOfFoundBlackListWords) {
    return Q.Promise(function(resolve, reject) {

        // Put the array of bad words into local storage so that we can
        // access them from the dialog.
        localStorage.setItem('badwords', JSON.stringify(arrayOfFoundBlackListWords));

        var url = getCurrentUrl() + 'dialog';

        // Call the dialog API.
        Office.context.ui.displayDialogAsync(url, {
            height: 30,
            width: 20            
        }, function(asyncResult) {
            dialog = asyncResult.value;

            // Listen for events that occur on the dialog object. This gives us
            // an entry point for actions after the dialog closes, or if the
            // dialog sends a message with messageParent().
            dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogEventReceived, function(event) {

                // See the ErrorCodeManager in the OfficeJS library for error codes.
                // Error 12002 indicates that the dialog page can't be found.
                if (event.error === 12002) {
                    reject(null);
                }
                else {
                    resolve();
                }
            });
        });
    });
}

/**
 * The dialog object populated by the dialogCallback. The dialog object must
 * be in the same scope as the dialog callback and event handler for dialog events.
 */
var dialog;

/**
 * The callback for the displayDialogAsync call to open the dialog UI.
 */
function dialogCallback(asyncResult) {

    dialog = asyncResult.value;

    // Listen for events that occur on the dialog object. This gives us
    // an entry point for actions after the dialog closes, or if the
    // dialog sends a message with messageParent().
    dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogEventReceived, processEvent);

    if (asyncResult.message === 'success') {
        console.log('Success message received in dialogCallback');
    } else {
        console.log('Failed message received in dialogCallback: ' + JSON.stringify(asyncResult));
    }
}

/**
 * Handle the events that occur on the dialog object.
 */
function processEvent(args) {
    if (args.type === 'dialogEventReceived') {
        // The dialog has been closed.
    }

    if (args.type === 'dialogMessageReceived') {
        // The dialog object sent a message to the add-in.
    }
}

/**
 * Get the base URL for the call.
 */
function getCurrentUrl() {

    // Note that calls to the function file are at a different location.
    if (location.href.indexOf('functionfile') > 0) {
        return location.href.slice(0, location.href.indexOf('functionfile'));
    } else {
        return location.href.slice(0, location.href.indexOf('?'));
    }
}
