/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/**
 * Shows how to get the sentence of the current selection.
 **/
function addBoilerplateSentence() {

    Word.run(function (context) {

        // This is the range of the sentence we want to save. You can just
        // add the cursor to the sentence.
        var sentence = context.document.getSelection().getTextRanges(['.'], true).first;

        // Queue a command to load the sentence text.
        context.load(sentence, 'text');

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync()
            .then(function () {

                // Here we save the sentence text as boilerplate in
                // local storage. The text is saved without formatting.
                var sentenceName = document.getElementById('inputAddBoilerplateSentence').value;
                saveBoilerplate(sentenceName, sentence.text, 'sentence');

            })
    }).catch(function (error) {
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

    Word.run(function (context) {

        // Get the paragraphs collection of the current selection.
        var paragraphs = context.document.getSelection().paragraphs;

        // Queue a command to load the paragraph collection of the selection.
        context.load(paragraphs, 'text');

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync()
            .then(function () {

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
    }).catch(function (error) {
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
    Word.run(function (context) {

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

    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}