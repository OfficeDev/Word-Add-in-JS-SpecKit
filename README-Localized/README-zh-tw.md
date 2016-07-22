# Word 增益集 JavaScript SpecKit

了解如何建立增益集，擷取並插入重複使用文字，以及如何實作簡單的文件驗證程序。

## 目錄
* [變更歷程記錄](#change-history)
* [必要條件](#prerequisites)
* [設定專案](#configure-the-project)
* [執行專案](#run-the-project)
* [瞭解程式碼](#understand-the-code)
* [問題和建議](#questions-and-comments)
* [其他資源](#additional-resources)

## 變更歷程記錄

2016 年 3 月 31 日：
* 初始範例版本。

## 必要條件

* Word 2016 for Windows，組建 16.0.6727.1000 或更新版本。
* [節點和 npm](https://nodejs.org/en/)
* [Git Bash](https://git-scm.com/downloads) - 您應該使用較新的組建，因為較早的組建在建立憑證時會顯示錯誤。

## 設定專案

在這個專案的根目錄，從您的 Bash 殼層執行下列命令︰

1. 複製此儲存機制到本機電腦。
2. ```npm install``` 安裝 package.json 中的所有相依性。
3. ```bash gen-cert.sh``` 建立執行這個範例所需的憑證。然後在您本機電腦上的儲存機制中，連按兩下 ca.crt，然後選取 [安裝憑證]<e />。選取 [本機電腦]<e />，然後選取 [下一步]<e /> 以繼續。選取 [將所有憑證放入以下的存放區]<e />，然後選取 [瀏覽]<e />。選取 [信任的根憑證授權]<e />，然後選取 [確定]<e />。選取 [下一步]<e />，然後選取 [完成]<e />，現在您的憑證授權單位憑證已新增至您的憑證存放區。
4. ```npm start``` 啟動服務。

您已經在這個時候部署此範例增益集。現在，您需要讓 Microsoft Word 知道哪裡可以找到此增益集。

1. 建立網路共用，或[共用網路資料夾](https://technet.microsoft.com/zh-tw/library/cc770880.aspx)，並將 [word-add-in-javascript-speckit-manifest.xml](word-add-in-javascript-speckit-manifest.xml) 資訊清單檔案放置在其中。
3. 啟動 Word 並開啟一個文件。
4. 選擇 [檔案]<e /> 索引標籤，然後選擇 [選項]<e />。
5. 選擇 [信任中心]<e />，然後選擇 [信任中心設定]<e /> 按鈕。
6. 選擇 [受信任的增益集目錄]<e />。
7. 在 [目錄 URL]<e /> 欄位中，輸入包含 word-add-in-javascript-speckit-manifest.xml 的資料夾共用的網路路徑，然後選擇 [新增目錄]<e />。
8. 選取 [顯示於功能表中]<e /> 核取方塊，然後選擇 [確定]<e />。
9. 接著會顯示訊息，通知您下次啟動 Office 時就會套用您的設定。關閉並重新啟動 Microsoft Word。

## 執行專案

1. 開啟 Word 文件。
2. 在 Word 2016 的 [插入]<e /> 索引標籤上，選擇 [我的增益集]<e />。
3. 選取 [共用資料夾]<e /> 索引標籤。
4. 選擇 [Word SpecKit 增益集]<e />，然後選取 [確定]<e />。
5. 如果您的 Word 版本支援增益集命令，UI 會通知您已載入增益集。

### 功能區 UI
在功能區上，您可以：
* 選取 [SpecKit 增益集]<e /> 索引標籤，在 UI 中啟動增益集。
* 選取 [插入規格範本]<e /> 以啟動工作窗格，並且將規格範本插入文件。
* 使用功能區中的驗證按鈕，或以滑鼠右鍵按一下內容功能表，根據文字的黑名單驗證文件。

 > 附註：如果您的 Word 版本不支援增益集命令，增益集會載入工作窗格。

### 工作窗格 UI
在工作窗格上，您可以：
* 將游標放在一個句子中以儲存句子、在上述欄位中為它指定名稱 **在工作窗格中新增句子至重複使用*，然後選取**新增句子至重複使用**。您可以對段落執行相同動作。
* 儲存句子與段落也會讓重複使用可在 [插入重複使用]<e /> 下拉式清單中使用。
* 將游標置於文件中。從下拉式清單選取重複使用文字，重複使用文字會插入文件。
* 變更文件的 [作者]<b /> 屬性，方法是變更作者名稱，並選取 [更新作者名稱]<e /> 按鈕。這樣會更新文件屬性和繫結內容控制項的內容。

## 瞭解程式碼

這個範例在預覽期間使用 1.2 [需求集](http://dev.office.com/reference/add-ins/office-add-in-requirement-sets?product=word)，但是當需求集正式運作時，需要 1.3 需求集。

### Task pane

工作窗格的功能是在 sample.js 中設定。sample.js 包含下列功能︰

* 設定 UI 和事件處理常式。
* 從服務擷取規格範本，並將它插入文件。
* 載入包含用於驗證文件的文字的黑名單。這些文字會針對這個範例的目的被視為不良字眼。
* 從服務載入預設的重複使用並且在本機存放區中快取它們。
* 測試函數檔案程式碼的基本架構程式碼。您想要在將增益集命令程式碼移至函數檔案之前在工作窗格中開發它，因為您無法將偵錯工具附加至函數檔案。
* 從文件屬性將預設作者的名稱載入工作窗格。顯示如何存取和變更文件中的自訂 XML 組件。
* 發佈重複使用至服務。

### 文件驗證和對話方塊 API

validation.js 包含要根據文字黑名單驗證文件的程式碼。ValidateContentAgainstBlacklist() 方法會使用新的 splitTextRanges 方法，根據分隔符號將文件分成幾個範圍。這個函數中的分隔符號會識別文件中的文字。我們識別文件和黑名單中文字的交集，並將這些結果傳遞至本機存放區。然後我們會使用 displayDialogAsync 方法來開啟對話方塊 (dialog.html)。對話方塊會從本機存放區取得驗證結果，並顯示結果。

### 重複使用文字功能

boilerplate.js 包含如何將重複使用文字儲存至本機存放區、使用對應至重複使用的項目更新 Fabric 下拉式清單，以及從下拉式清單中插入選取的重複使用的範例。這個檔案包含的範例︰
* splitTextRanges (WordApi 1.3 需求集的新項目) - 這個 API 在將來會被 split() 取代。
* compareLocationWith (WordApi 1.3 需求集的新項目)
* 使用新的項目來更新 Fabric 下拉式清單。
* 將重複使用文字插入文件。

### 自訂 XML 繫結至核心文件屬性

authorCustomXml.js 包含用於取得和設定預設文件屬性的方法。

* 載入工作窗格時，將作者屬性載入工作窗格。請注意，文件也包含作者屬性的值。這是因為範本會包含繫結至這個文件屬性的內容控制項。這可讓您根據自訂 XML 組件的內容，設定文件中的預設值。
* 從工作窗格更新作者屬性。這樣會更新文件屬性和文件中的繫結內容控制項。

### 增益集命令

增益集命令宣告位於 word-add-in-javascript-speckit-manifest.xml。這個範例會示範如何在功能區中和在內容功能表中建立增益集命令。

## 問題和建議

我們很樂於收到您對於 Word SpecKit 範例的意見反應。您可以在此儲存機制的*問題*區段中，將您的意見反應傳送給我們。

請在 [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API) 提出有關 Microsoft Office 365 開發的一般問題。務必以 [office-js] 和 [API] 標記您的問題。

## 其他資源

* [Office 增益集文件](https://msdn.microsoft.com/zh-tw/library/office/jj220060.aspx)
* [Office 開發人員中心](http://dev.office.com/)
* [Office 365 API 入門專案和程式碼範例](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)

## 著作權
Copyright (c) 2016 Microsoft Corporation 著作權所有，並保留一切權利。


