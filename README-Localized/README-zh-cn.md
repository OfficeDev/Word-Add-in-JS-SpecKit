# <a name="word-add-in-javascript-speckit"></a>Word 外接程序 JavaScript SpecKit

了解如何创建捕获和插入样本文字的外接程序，以及如何实施一个简单的文档验证过程。

## <a name="table-of-contents"></a>目录
* [修订记录](#change-history)
* [先决条件](#prerequisites)
* [配置项目](#configure-the-project)
* [运行项目](#run-the-project)
* [了解代码](#understand-the-code)
* [问题和意见](#questions-and-comments)
* [其他资源](#additional-resources)

## <a name="change-history"></a>修订记录

2016 年 3 月 31 日：
* 初始示例版本。

## <a name="prerequisites"></a>先决条件

* Word 2016 for Windows，内部版本 16.0.6727.1000 或更高版本。
* [Node 和 npm](https://nodejs.org/en/)
* [Git Bash](https://git-scm.com/downloads) - 应使用较高的内部版本，因为较低的内部版本可能会在生成证书时显示错误。

## <a name="configure-the-project"></a>配置项目

在此项目的根目录处从你的 Bash shell 中运行以下命令：

1. 将此存储库克隆到本地计算机中。
2. ```npm install``` 可安装 package.json 中的所有依赖项。
3. ```bash gen-cert.sh``` 可创建运行此示例所需的证书。然后在本地计算机的存储库中，双击 ca.crt，然后选择“**安装证书**”。选择“**本地计算机**”，然后选择“**下一步**”以继续。选择“**将所有的证书都放入下列存储**”，然后选择“**浏览**”。选择“**受信任的根证书颁发机构**”，然后选择“**确定**”。选择“**下一步**”，然后选择“**完成**”，现在你的证书颁发机构证书已添加到你的证书存储中。
4. ```npm start``` 可启动服务。

此时，已部署此示例加载项。现在，你需要让 Microsoft Word 知道在哪里可以找到该外接程序。

1. 创建网络共享，或[将文件夹共享到网络](https://technet.microsoft.com/en-us/library/cc770880.aspx)，并将 [word-add-in-javascript-speckit-manifest.xml](word-add-in-javascript-speckit-manifest.xml) 清单文件放入其中。
3. 启动 Word，然后打开一个文档。
4. 选择**文件**选项卡，然后选择**选项**。
5. 选择**信任中心**，然后选择**信任中心设置**按钮。
6. 选择“受信任的加载项目录”****。
7. 在“目录 URL”****字段中，输入包含 word-add-in-javascript-speckit-manifest.xml 的文件夹共享的网络路径，再选择“添加目录”****。
8. 选中“在菜单中显示”****复选框，再选择“确定”****。
9. 随后会出现一条消息，告知你下次启动 Microsoft Office 时将应用你的设置。关闭并重新启动 Word。

## <a name="run-the-project"></a>运行项目

1. 打开一个 Word 文档。
2. 在 Word 2016 的“插入”****选项卡中，选择“我的加载项”****。
3. 选择“共享文件夹”****选项卡。
4. 依次选择“Word SpecKit 加载项”****和“确定”****。
5. 如果你的 Word 版本支持外接程序命令，UI 将通知你加载了外接程序。

### <a name="ribbon-ui"></a>功能区用户界面
在功能区上，你可以：
* 选择“SpecKit 加载项”****选项卡，在 UI 中启动加载项。
* 选择“插入规范模板”****，以启动任务窗格，并将规范模板插入文档。
* 使用功能区中的验证按钮，或右键单击关联菜单，即可验证文档中是否有列入黑名单的字词。

 > 注意：如果你的 Word 版本不支持外接程序命令，则外接程序将在任务窗格中加载。

### <a name="task-pane-ui"></a>任务窗格 UI
在任务窗格上，可以：
* 通过将光标放入句中来保存句子，在任务窗格中“将句子添加到样本”***上方的字段中命名句子，再选择“将句子添加到样本”****。可以对段落执行相同的操作。
* 保存句子和段落也会让样本出现在“插入样本”****下拉列表中。
* 将光标放在文档中。从下拉列表中选择样本文字，即会将选择的样本文字插入到文档中。
* 通过更改作者姓名并选择“*更新作者姓名*”按钮来更改文档的“**作者**”属性。这将更新文档属性和绑定内容控件的内容。

## <a name="understand-the-code"></a>了解代码

此示例在预览期间使用 1.2 [要求集](http://dev.office.com/reference/add-ins/office-add-in-requirement-sets?product=word)，但在 1.3 要求集公开发布后将需要该要求集。

### <a name="task-pane"></a>任务窗格

在 sample.js 中设置任务窗格功能。sample.js 包含以下功能：

* 设置 UI 和事件处理程序。
* 从服务获取规范模板并将其插入到文档中。
* 加载包含用于验证文档的单词的黑名单。这些词被认为是对本示例不好的单词。
* 从服务加载默认样本并将其缓存在本地存储区中。
* 用于测试函数文件代码的主干代码。你将希望在任务窗格中开发自己的外接程序命令代码，之后将其移到函数文件中，因为你无法将调试器附加到函数文件中。
* 将来自文档属性的默认作者的姓名加载到任务窗格中。这将显示如何访问和更改文档中的自定义 XML 部件。
* 将样本发布到服务。

### <a name="document-validation-and-the-dialog-api"></a>文档验证和对话框 API

validation.js 包含针对单词的黑名单验证文档的代码。validateContentAgainstBlacklist() 方法使用新的 splitTextRanges 方法将文档拆分为基于分隔符的范围。此函数中的分隔符可标识文档中的单词。我们可以标识文档中插入的单词和黑名单，并将这些结果传递到本地存储区。然后，我们可以使用 displayDialogAsync 方法打开一个对话框 (dialog.html)。该对话框将从本地存储区中获得验证结果并显示结果。

### <a name="boilerplate-text-functionality"></a>样本文字功能

boilerplate.js 包含一些示例，说明如何将样本文字保存到本地存储区，如何使用与保存的样本相对应的条目更新结构下拉列表，以及如何插入从下拉列表中选择的样本。此文件包含以下内容的示例：
* splitTextRanges（WordApi 1.3 要求集的新内容）- 在将来的版本中此 API 将由 split() 替代。
* compareLocationWith（WordApi 1.3 要求集的新内容）
* 使用新条目更新结构下拉列表。
* 将样本文字插入到文档中。

### <a name="custom-xml-binding-to-core-document-properties"></a>绑定到核心文档属性的自定义 XML

authorCustomXml.js 包含用于获取和设置默认文档属性的方法。

* 加载任务窗格时，将作者属性加载到任务窗格中。请注意，文档还包含作者属性的值。这是因为模板包含绑定到此文档属性的内容控件。这使你可以在基于自定义的 XML 部件的文档设置默认值。
* 从任务窗格中更新作者属性。这将更新文档属性和文档中的绑定内容控件。

### <a name="add-in-commands"></a>外接程序命令

外接程序命令声明位于 word-add-in-javascript-speckit-manifest.xml 中。此示例演示如何在功能区中和上下文菜单中创建外接程序命令。

## <a name="questions-and-comments"></a>问题和意见

我们乐意倾听你对 Word SpecKit 示例的相关反馈。你可以在该存储库中的“*问题*”部分将你的反馈发送给我们。

与 Microsoft Office 365 开发相关的问题一般应发布到 [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API)。确保你的问题使用了 [office-js] 和 [API] 标记。

## <a name="additional-resources"></a>其他资源

* 
  [Office 外接程序文档](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Office 开发人员中心](http://dev.office.com/)
* [Office 365 API 入门项目和代码示例](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)

## <a name="copyright"></a>版权
版权所有 (c) 2016 Microsoft Corporation。保留所有权利。



此项目已采用 [Microsoft 开放源代码行为准则](https://opensource.microsoft.com/codeofconduct/)。有关详细信息，请参阅[行为准则 FAQ](https://opensource.microsoft.com/codeofconduct/faq/)。如有其他任何问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。
