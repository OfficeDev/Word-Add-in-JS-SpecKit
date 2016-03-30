/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/**************************************************************************/
/** Get and set the author from a custom XML part.
 **************************************************************************/

/**
 * Gets the author's name from the document's core properties and loads the
 * name into the task pane UI. This function is run when the task pane loads.
 **/
function loadAuthorName() {

    // Get the XML node in core.xml that contains the author setting.
    getCoreXml(function(authorCustomXmlNode) {

        // Set the author's name in the document, which means this updates the
        // custom XML part called core.xml.
        authorCustomXmlNode.getTextAsync(function(getTextAsyncResult) {
            if (getTextAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
                // Let's clear the default label.
                var authorLabel = document.getElementById('authorLabel');
                authorLabel.innerText = '';

                // This gets the task pane input element associated with the author.
                var author = document.getElementById('author');

                // Put the author's name in the task pane.
                author.value = getTextAsyncResult.value;

                console.log('Updated the author property in the UI');
            }
        });
    });
}

/**
 * Sets the author's name in the core document properties. A content control
 * in the template is bound to that property so that content control's text
 * value is updated when the author's name is updated and submitted in the task pane.
 **/
function updateAuthor() {

    // This gets the input element associated with the author.
    var author = document.getElementById('author');

    // Get the XML node in core.xml that contains the author setting.
    getCoreXml(function(authorCustomXmlNode) {

        // Set the author's name in the document, which means this updates the
        // custom XML part called core.xml.
        authorCustomXmlNode.setTextAsync(author.value, function(setTextAsyncResult) {
            if (setTextAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
                console.log('Updated the author property in the document.');
            }
        });
    });
}

/**
 * Helper to get the author property from core document properties (core.xml).
 **/
function getCoreXml(callback) {
    // Get the built-in core properties XML part by using its ID. This results in a call to Word.
    // You can get the ID of a custom XML part by renaming the Word document's extension to
    // .zip and look at the custom XML part information.
    Office.context.document.customXmlParts.getByIdAsync("{6C3C8BC8-F283-45AE-878A-BAB7291924A1}", function(getByIdAsyncResult) {

        // Access the XML part.
        var xmlPart = getByIdAsyncResult.value;

        // Add namespaces to the namespace manager. These two calls result in two calls to Word.
        xmlPart.namespaceManager.addNamespaceAsync('cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties', function() {
            xmlPart.namespaceManager.addNamespaceAsync('dc', 'http://purl.org/dc/elements/1.1/', function() {

                // Get XML nodes by using an Xpath expression. This results in a call to Word.
                xmlPart.getNodesAsync("/cp:coreProperties/dc:creator", function(getNodesAsyncResult) {

                    // Get the first node returned by using the Xpath expression.
                    var customXmlNode = getNodesAsyncResult.value[0];

                    // Provide the CustomXmlNode object.
                    // https://msdn.microsoft.com/EN-US/library/office/fp142260.aspx
                    callback(customXmlNode);
                });
            });
        });
    });
}