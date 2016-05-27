/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
(function () {
    "use strict";

    // The initialize function is run each time a page of the add-in is loaded into the task pane.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // Use this to check whether the new API is supported in the Word client.
            // The compareLocationWith and splitTextRanges method calls will be
            // in the 1.3 requirement set. The 1.3 requirement set check is not
            // implemented in preview. We are checking against the 1.2 API
            // requirement set check because this is the minimum requirement set
            // that supports the new functionality in Word. This is true as long as you
            // have the March 2016 update of Word 2016 for Windows.
            // Update this to target the correct set after 1.3 is generally available.
            if (Office.context.requirements.isSetSupported("WordApi", 1.2)) {

                // Turn on and off logging to the console for WordJS calls.
                OfficeExtension.Utility._logEnabled = false;

                // Initialize stylized fabric UI for text fields.
                $(".ms-TextField").TextField();

                // You'll want to test function file code before you put it in a
                // HTML function file because there isn't a well-defined way to
                // test and debug code that is referenced by, or contained in, the
                // function file.
                //$('#testFunctionFileCode').click(onTestFunctionFileCode);

                // This is in authorCustomXml.js.
                $('#updateAuthor').click(updateAuthor);

                // This is in boilerplate.js.
                $('#btnAddBoilerplateParagraph').click(addBoilerplateParagraph);
                $('#btnAddBoilerplateSentence').click(addBoilerplateSentence);
                $('#boilerplateDropdown').change(selectBoilerplate);
                $('#btnSaveBoilerplate').click(saveBoilerplate)

                /**************************************************************/
                /* Default actions to happen when the task pane loads.
                 **************************************************************/

                // Insert the spec template when the task pane is loaded. This would
                // be a great place to use the dialog API to confirm whether you
                // want to reload the template. As this code is now, the template
                // will reload if you refresh the task pane.
                fetchSpecTemplate();

                // Gets the author's name and loads it into the task pane UI.
                // This is in authorCustomXml.js.
                loadAuthorName();
            }
            else {
                // Just letting you know that this code will not work with your version of Word.
                console.log('This add-in requires Word 2016 or greater.');
            }
        });
    };

    /**************************************************************************/
    /** Default task pane load actions.
    /**************************************************************************/

    /**
     * Fetch the template from the service.
    */
    function fetchSpecTemplate() {

        // Form the URL to fetch the template. We need to remove the host
        // information by slicing off the host information beginning at
        // ?_host_Info. See server.js for the gettemplate route.
        var getTemplateUrl = getCurrentUrl() + 'gettemplate';

        // Fetch the template and then insert it into the current document's body.
        var urlPromise = httpGetAsync(getTemplateUrl);
        urlPromise.then(insertFileIntoBody)
            .catch(function (err) {
                console.log('Error: ' + err);
            });
    }

    /**
     * Inserts the template document into the current document body and
     * initializes the current document with boilerplate stored
     * at the service.
     *
     * @param templateBase64 The base64 encoded template document.
     **/
    function insertFileIntoBody(templateBase64) {
        // Entry point for accessing the Word object model.
        return Word.run(function (context) {

            // Queue a command to insert the template into the current document.
            var body = context.document.body;
            body.insertFileFromBase64(templateBase64, Word.InsertLocation.replace);

            // Synchronize the document state by executing the queued command to
            // insert the template into the current document,
            // call getBlackList/getBoilerplate to,
            // and return a promise to indicate task completion.
            return context.sync()
                .then(getBlackList)
                .then(getBoilerplate)
                .then(context.sync);
        }).catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
    }

    /**
     * Gets a JSON document that contains a blacklist of words.
     */
    function getBlackList() {

        // Check whether the cache of bad words exists. Add the list to localStorage.
        // We aren't updating the cache.
        if (!localStorage.getItem('badwordcache')) {

            // Form the URL to get the bad word list. Need to remove the host information by slicing
            // off the host information beginning at ?_host_Info. See server.js for the route
            // that is blacklist.
            var getBlacklistUrl = getCurrentUrl() + 'blacklist';

            // Call the service to get the blacklist, and then cache it in
            // localStorage, then return the promise.
            return httpGetAsync(getBlacklistUrl)
                .then(function (json) {
                    localStorage.setItem('badwordcache', json);
                });
        } else {
            return Q.fcall(function () {
                return 'The bad words cache exists.';
            });
        }
    }

    /**
     * Gets the boilerplate text. This demonstrates how to initialize the
     * current document with boilerplate stored in a service.
     */
    function getBoilerplate() {

        // Form the URL to get the boilerplate from the service. Need to remove
        // the host information by slicing off the host information beginning
        // at ? _host_Info. See server.js for the route that is boilerplate.
        var getBoilerplateUrl = getCurrentUrl() + 'boilerplate';

        // Call the service to get the boilerplate and add it to localStorage,
        // then return the promise.
        return httpGetAsync(getBoilerplateUrl)
            .then(function (json) {

                // Put the array of bad words into local storage so that we can
                // access them.
                localStorage.setItem('boilerplate', json);
                var boilerplate = JSON.parse(json);

                var elements = boilerplate.elements;

                // Add the boilerplate names to the drop down.
                for (var i = 0; i < elements.length; i++) {
                    $('#boilerplateDropdown').append(new Option(elements[i].name,
                        elements[i].name));
                }

                // Initialize stylized fabric UI for dropdown and call the dropdown
                // function to populate dropdown with values. You need to call this
                // when you update contents of a dropdown.
                $(".ms-Dropdown").Dropdown();
            });
    }

    /**
     * Save the current boilerplate state with the services.
     **/
    function saveBoilerplate() {

        // Form the URL to post the boilerplate to the service. Need to remove
        // the host information by slicing off the host information beginning
        // at ? _host_Info. See server.js for the route which is boilerplate.
        var postBoilerplateUrl = getCurrentUrl() + 'boilerplate';

        var boilerplate = localStorage.getItem('boilerplate');

        httpPostAsync(postBoilerplateUrl, boilerplate)
            .then(function (value) {
                console.log(value);
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
            });
    }

    /**
     * GET helper to call a service.
     *
     * @param url {string} The URL of the service.
     */
    function httpGetAsync(url) {

        return Q.Promise(function (resolve, reject) {
            var request = new XMLHttpRequest();

            request.open("GET", url, true);
            request.onload = onload;
            request.onerror = onerror;

            function onload() {
                if (request.status === 200) {
                    resolve(request.responseText);
                } else {
                    reject(new Error("Status code: " + request.status));
                }
            }

            function onerror() {
                reject(new Error('Status code: ' + request.status));
            }
            request.send();
        });
    }

    /**
     * POST helper to call a service.
     *
     * @param url {string} The URL of the service.
     * @param payload {string} The JSON payload.
     */
    function httpPostAsync(url, payload) {
        var request = new XMLHttpRequest();
        var deferred = Q.defer();

        request.open("POST", url, true);
        request.setRequestHeader("Content-type", "application/json");

        request.onload = function () {
            if (request.status === 200) {
                deferred.resolve("Successful boilerplate save.");
            } else {
                deferred.reject(new Error('Status code: ' + request.status));
            }
        };

        request.onerror = function () {
            deferred.reject(new Error('Status code: ' + request.status));
        };

        request.send(payload);

        return deferred.promise;
    }

    /**************************************************************************/
    /** Test area for function file code. You'll want to use a function like
     *  this to test your function file code before you add it to the function
     *  file. This let's you attach a debugger and make sure the code works
     *  before adding it to the function file where you won't be able to use the
     *  debugger.
     **************************************************************************/

    // Temp function for testing add-in command
    // function onTestFunctionFileCode() {
    //     validateAgainstBlacklist();
    // }

})();
