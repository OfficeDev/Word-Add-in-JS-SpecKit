/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/**
 * Shows how to get the sentence of the current selection. This example shows
 * how to get the current paragraph of the selection, get all of the sentences
 * in the paragraph, and then compare each sentence to the original selection
 * to determine which sentence contains the selection. This is essentially
 * like expanding a range to the bounds of the sentence that contains the range.
 **/
function addBoilerplateSentence() {

    Word.run(function(context) {

        // This is the range of the sentence we want to save. You can just
        // add the cursor to the sentence and we'll figure out the bounds of the sentence.
        var originalRange = context.document.getSelection();

        // Get the paragraphs collection of the current selection.
        var paragraphs = originalRange.paragraphs;

        // Queue a command to load the paragraph collection of the selection.
        context.load(paragraphs, 'text');
        context.load(originalRange, 'text');

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync()
            .then(function() {

                if (paragraphs.items.length === 1) {
                    // The range exists within a single paragraph. Now we just
                    // need to get the paragraph.

                    // This essentially gets the paragraph where this range exists. We don't have an
                    // expand to paragraph, but this approach gives the same results.
                    var paragraph = paragraphs.items[0];

                    // Queue a command to get all of the sentences in the paragraph.
                    // We'll include the delimiters because we want a complete sentence.
                    var ranges = paragraph.splitTextRanges(['.'], false, true);

                    // Queue a command to load the sentences and their text.
                    context.load(ranges, 'text');

                    // Synchronize the document state by executing the queued commands,
                    // and return a promise to indicate task completion. We're passing
                    // the array of sentence ranges to the next promise.
                    return context.sync(ranges);
                }
                else {
                    // You've selected zero or more than one paragraph.
                    console.log('Please select a single paragraph');
                    return;
                }
            }).then(function(sentences) {

                if (sentences) {
                    var callbacklist = [];

                    // Compare our original selection range with the sentences
                    // returned by splitTextRanges. We call this function for
                    // each sentence.
                    function getRangeLocation(count) {
                        var rangeLocation = sentences.items[count].compareLocationWith(originalRange);
                        return context.sync().then(function() {
                            if (rangeLocation.value === Word.LocationRelation.contains) {
                                // Here we save the sentence text as boilerplate in
                                // local storage. The text is saved without formatting.
                                var sentenceName = document.getElementById('inputAddBoilerplateSentence').value;
                                saveBoilerplate(sentenceName, sentences.items[count].text, 'sentence');
                            }
                        });
                    }

                    // Add each call to getRangeLocation to an array of callbacks.
                    // We're doing this so that we can maintain the index of each
                    // sentence we want to call.
                    for (var r = 0; r < sentences.items.length; r++) {
                        callbacklist.push(
                            (function(r) {
                                return function() {
                                    getRangeLocation(r);
                                }
                            })(r)
                        )
                    }

                    // Call each instance of getRangeLocation that we added to
                    // the array of callbacks in callbacklist.
                    for (var callback in callbacklist) {
                        callbacklist[callback].call(this);
                    }
                }
            })
    }).catch(function(error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

/**
 * Shows how to get the paragraph of a range assuming that the range does not
 * cross paragraph boundaries. This is essentially like expanding a range to
 * the bounds of the paragraph that contains the range.
 **/
function addBoilerplateParagraph() {

    Word.run(function(context) {

        // Get the paragraphs collection of the current selection.
        var paragraphs = context.document.getSelection().paragraphs;

        // Queue a command to load the paragraph collection of the selection.
        context.load(paragraphs, 'text');

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync()
            .then(function() {

                if (paragraphs.items.length === 1) {
                    // The range exists within a single paragraph. Now we just
                    // need to get the paragraph.

                    // This essentially gets the paragraph where this range exists. We don't have an
                    // expand to paragraph, but this approach gives the same results.
                    var paragraph = paragraphs.items[0];

                    // Here we save the paragraph text as boilerplate. The text
                    // is saved without formatting.
                    var paragraphName = document.getElementById('inputAddBoilerplateParagraph').value;
                    saveBoilerplate(paragraphName, paragraph.text, 'paragraph');
                }
                else {
                    // You've selected zero or more than one paragraph.
                    console.log('Please select a single paragraph');
                }
            })
    }).catch(function(error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

/**
 * Save the selected content to the boilerplate storage.
 * @param boilerplateName The name of the content provided in the task pane.
 * @param boilerplateText The plain text contents of the selected content.
 * @param boilerplateType The type of boilerplate. Currently, section is the only type that matters.
 **/
function saveBoilerplate(boilerplateName, boilerplateText, boilerplateType) {

    // Update localstorage with the new boilerplate.
    var boilerplate = JSON.parse(localStorage.getItem('boilerplate'));
    var element = { name: boilerplateName, type: boilerplateType, text: boilerplateText };

    boilerplate.elements.push(element);
    localStorage.setItem('boilerplate', JSON.stringify(boilerplate));

    // Update the dropdown with the new sentence name.
    $('#boilerplateDropdown').append(new Option(element.name,
        element.name));

    // Initialize stylized fabric UI for dropdown and call the dropdown
    // function to populate dropdown with values. You need to call this
    // when you update contents of a dropdown.
    $(".ms-Dropdown").Dropdown('refresh');
}

/**
 * Finds the selected boilerplate from local storage then insert boilerplate
 * into the document.
 **/
function selectBoilerplate() {

    var boilerplateSelection = this.value;
    var boilerplate = JSON.parse(localStorage.getItem('boilerplate'));

    for (var i = 0; i < boilerplate.elements.length; i++) {

        if (boilerplate.elements[i].name === boilerplateSelection) {
            // We have the boilerplate.
            var text = boilerplate.elements[i].text;
            var type = boilerplate.elements[i].type;

            insertBoilerplate(text, type);
            return;
        }
    }
}

/**
 * Inserts the boilerplate into the document.
 *
 * @param text The text to insert into the document.
 * @param type The type of boilerplate. This can be a paragraph, sentence, or section.
 **/
function insertBoilerplate(text, type) {
    Word.run(function(context) {

        // Queue a command to get the current selection.
        var selection = context.document.getSelection();

        if (type === 'section') {
            // Queue a command to insert text that is styled with Heading 1.
            var range = selection.insertText(text, Word.InsertLocation.end);
            range.style = 'Heading 1';
        }
        else {
            // Queue a command to insert text into the document.
            selection.insertText(text, Word.InsertLocation.end);
        }

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync();

    }).catch(function(error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}