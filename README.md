---
topic: sample
products:
- office-word
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 3/24/2016 12:45:01 PM
---
# Word Add-in JavaScript SpecKit

Learn how you can create an add-in that captures and inserts boilerplate text, and how you can implement a simple document validation process.

## Table of Contents
* [Change History](#change-history)
* [Prerequisites](#prerequisites)
* [Configure the project](#configure-the-project)
* [Run the project](#run-the-project)
* [Understand the code](#understand-the-code)
* [Questions and comments](#questions-and-comments)
* [Additional resources](#additional-resources)

## Change History

March 31, 2016:
* Initial sample version.

## Prerequisites

* Word 2016 for Windows, build 16.0.6727.1000 or later.
* [Node and npm](https://nodejs.org/en/)
* [Git Bash](https://git-scm.com/downloads) - You should use a later build as earlier builds can show an error when generating the certificates.

## Configure the project

Run the following commands from your Bash shell at the root of this project:

1. Clone this repo to your local machine.
2. ```npm install``` to install all of the dependencies in package.json.
3. ```bash gen-cert.sh``` to create the certificates needed to run this sample. Then in the repo on your local machine, double-click ca.crt, and select **Install Certificate**. Select **Local Machine** and select **Next** to continue. Select **Place all certificates in the following store** and then select **Browse**.  Select **Trusted Root Certification Authorities** and then select **OK**. Select **Next** and then **Finish** and now your certificate authority cert has been added to your certificate store.
4. ```npm start``` to start the service.

You've deployed this sample add-in at this point. Now you need to let Microsoft Word know where to find the add-in.

1. Create a network share, or [share a folder to the network](https://technet.microsoft.com/en-us/library/cc770880.aspx) and place the [word-add-in-javascript-speckit-manifest.xml](word-add-in-javascript-speckit-manifest.xml) manifest file in it.
3. Launch Word and open a document.
4. Choose the **File** tab, and then choose **Options**.
5. Choose **Trust Center**, and then choose the **Trust Center Settings** button.
6. Choose **Trusted Add-ins Catalogs**.
7. In the **Catalog Url** field, enter the network path to the folder share that contains word-add-in-javascript-speckit-manifest.xml, and then choose **Add Catalog**.
8. Select the **Show in Menu** check box, and then choose **OK**.
9. A message is displayed to inform you that your settings will be applied the next time you start Microsoft Office. Close and restart Word.

## Run the project

1. Open a Word document.
2. On the **Insert** tab in Word 2016, choose **My Add-ins**.
3. Select the **SHARED FOLDER** tab.
4. Choose **Word SpecKit add-in**, and then select **OK**.
5. If add-in commands are supported by your version of Word, the UI will inform you that the add-in was loaded.

### Ribbon UI
On the Ribbon, you can:
* Select **SpecKit add-in** tab to launch the add-in in the UI.
* Select **Insert spec template** to launch the task pane and insert a spec template into the document.
* Use the validation buttons in the ribbon or right-click the context menu to validate the document against a blacklist of words.

 > Note: The add-in will load in a task pane if add-in commands are not supported by your version of Word.

### Task pane UI
On the task pane, you can:
* Save a sentence by putting the cursor in a sentence, give it a name in the field above **Add sentence to boilerplate* in the task pane, and select **Add sentence to boilerplate**. You can do the same for paragraphs.
* Saving sentences and paragraphs will also make the boilerplate available in the **Insert boilerplate** dropdown.
* Place the cursor in the document. Select a boilerplate text from the drop down and the boilerplate text will get inserted into the document.
* Change the *Author* property of the document by changing the author name and selecting the **Update author name** button. This will update both the document property and the contents of a bound content control.

## Understand the code

This sample uses the 1.2 [requirement set](http://dev.office.com/reference/add-ins/office-add-in-requirement-sets?product=word) during the preview period but will require the 1.3 requirement set once that requirement set is generally available.

### Task pane

The task pane functionality is set up in sample.js. sample.js contains the following functionality:

* Set up the UI and event handlers.
* Fetch the spec template from a service and insert it into the document.
* Load a blacklist that contains words that are used to validate the document. These words are considered bad words for the purpose of this sample.
* Load a default boilerplate from a service and cache them in local storage.
* Skeleton code for testing function file code. You'll want to develop your add-in command code in the task pane before moving it into a function file because you can't attach a debugger to the function file.
* Load the default author's name from the document properties into the task pane. This shows how you can access and change a custom XML part in a document.
* Post the boilerplate to the service.

### Document validation and the Dialog API

validation.js contains the code to validate the document against a blacklist of words. The validateContentAgainstBlacklist() method uses the new splitTextRanges method to split the document into ranges based on delimiters. The delimiters in this function identify words in the document. We identify the intersection of words in the document and the blacklist and pass those results to local storage. Then we use the displayDialogAsync method to open a dialog (dialog.html). The dialog gets the validation results from local storage and displays the results.

### Boilerplate text functionality

boilerplate.js contains examples of how you can save boilerplate text to local storage, update a Fabric dropdown with entries that correspond to saved boilerplate, and insert boilerplate selected from a dropdown. This file contains examples of:
* splitTextRanges (new for the WordApi 1.3 requirement set) - this API will be replaced by split() in a future release.
* compareLocationWith (new for the WordApi 1.3 requirement set)
* Update the Fabric dropdown with the new entries.
* Insert boilerplate text into the document.

### Custom XML binding to core document properties

authorCustomXml.js contains methods for getting and setting the default document properties.

* Load the author property into the task pane when the task pane loads. Notice that the document also contains the value of the author property. This is because the template contains a content control that is bound to this document property. This enables you to set default values in the document based on the contents of a custom XML part.
* Update the author property from the task pane. This will update the document property and the bound content control in the document.

### Add-in commands

The add-in command declarations are located in word-add-in-javascript-speckit-manifest.xml. This sample shows how to create add-in commands in the ribbon and in a context menu.

## Questions and comments

We'd love to get your feedback about the Word SpecKit sample. You can send your feedback to us in the *Issues* section of this repository.

Questions about Microsoft Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). Make sure that your questions are tagged with [office-js] and [API].

## Additional resources

* [Office add-in documentation](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Office Dev Center](http://dev.office.com/)
* [Office 365 APIs starter projects and code samples](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)

## Copyright
Copyright (c) 2016 Microsoft Corporation. All rights reserved.



This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
