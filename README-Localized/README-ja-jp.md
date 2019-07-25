---
page_type: sample
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
# <a name="word-add-in-javascript-speckit"></a>Word アドインの JavaScript SpecKit

定型句をキャプチャして挿入するアドインの作成方法について説明します。また、簡単なドキュメントの検証プロセスを実装する方法について説明します。

## <a name="table-of-contents"></a>目次
* [変更履歴](#change-history)
* [前提条件](#prerequisites)
* [プロジェクトを構成する](#configure-the-project)
* [プロジェクトを実行する](#run-the-project)
* [コードを理解する](#understand-the-code)
* [質問とコメント](#questions-and-comments)
* [その他のリソース](#additional-resources)

## <a name="change-history"></a>変更履歴

2016 年 3 月 31 日
* サンプルの初期バージョン。

## <a name="prerequisites"></a>前提条件

* Windows 用の Word 2016 (16.0.6727.1000 以降のビルド)。
* [Node と npm](https://nodejs.org/en/)
* [Git Bash](https://git-scm.com/downloads) -証明書を生成するときに、以前のビルドではエラーが発生する可能性があるため、最新のビルドを使用する必要があります。

## <a name="configure-the-project"></a>プロジェクトを構成する

このプロジェクトのルートで Bash シェルから以下のコマンドを実行します。

1. ローカル コンピューターにこのリポジトリのクローンを作成します。
2. ```npm install``` で package.json にすべての依存関係をインストールします。
3. ```bash gen-cert.sh``` このサンプルを実行するために必要な証明書を作成します。ローカル コンピューターにあるリポジトリで、ca.crt をダブルクリックし、**[証明書のインストール]** を選択します。**[ローカル コンピューター]** を選択して、**[次へ]** 選択して続行します。**[証明書をすべて次のストアに配置する]** を選択してから **[参照]** を選択します。**[信頼されたルート証明機関]** を選択して、**[OK]** を選択します。**[次へ]**、**[完了]** の順に選択すると、証明書ストアに証明書機関の証明書が追加されます。
4. ```npm start``` でサービスを開始します。

この時点でこのサンプル アドインは展開されました。次に、Microsoft Word がアドインを検索する場所を認識できるようにする必要があります。

1. ネットワーク共有を作成するか、[ネットワークでフォルダーを共有し](https://technet.microsoft.com/en-us/library/cc770880.aspx)、そのフォルダーに [word-add-in-javascript-speckit-manifest.xml](word-add-in-javascript-speckit-manifest.xml) マニフェスト ファイルを配置します。
3. Word を起動し、ドキュメントを開きます。
4. [**ファイル**] タブを選択し、[**オプション**] を選択します。
5. [**セキュリティ センター**] を選択し、[**セキュリティ センターの設定**] ボタンを選択します。
6. **[信頼されているアドイン カタログ]** を選択します。
7. **[カタログの URL]** フィールドに、word-add-in-javascript-speckit-manifest.xml があるフォルダー共有へのネットワーク パスを入力して、**[カタログの追加]** を選択します。
8. **[メニューに表示する]** チェック ボックスをオンにして、**[OK]** を選択します。
9. これらの設定が Microsoft Office を次回起動したときに適用されることを示すメッセージが表示されます。Word を終了して、再起動します。

## <a name="run-the-project"></a>プロジェクトを実行する

1. Word 文書を開きます。
2. Word 2016 の**[挿入]** タブで、**[マイ アドイン]** を選択します。
3. **[共有フォルダー]** タブを選択します。
4. **[Word SpecKit アドイン]** を選択して、**[OK]** を選択します。
5. ご使用の Word バージョンでアドイン コマンドがサポートされている場合、UI によってアドインが読み込まれたことが通知されます。

### <a name="ribbon-ui"></a>リボン UI
リボンで、次の操作を実行できます。
* **[SpecKit アドイン]** タブを選択して、UI でアドインを起動します。
* **[仕様テンプレートを挿入]** を選択して作業ウィンドウを起動し、仕様テンプレートをドキュメントに挿入します。
* リボンで検証ボタンを使用するか、コンテキスト メニューを右クリックしてブラックリストの単語に照らし合わせてドキュメントを検証します。

 > 注: アドイン コマンドが Word バージョンによってサポートされていない場合は、アドインが作業ウィンドウに読み込まれます。

### <a name="task-pane-ui"></a>作業ウィンドウの UI
作業ウィンドウで次の操作を実行できます。
* 文にカーソルを配置して文を保存し、作業ウィンドウの **[文を定型に追加する]* の上にあるフィールドで名前を付けて、**[文を定型に追加する]** を選択します。段落に対しても同じ操作を実行できます。
* 文や段落を保存すると、**[定型を挿入]** ドロップダウン リストで利用できる定型になります。
* ドキュメントにカーソルを配置します。ドロップダウンから定型句を選択すると、定型句がドキュメントに挿入されます。
* 作成者名を変更して *[作成者名の更新]* ボタンを選択することにより、ドキュメントの **[作成者]** プロパティを変更します。この操作により、ドキュメントのプロパティとバインドされたコンテンツ コントロールの内容の両方が更新されます。

## <a name="understand-the-code"></a>コードを理解する

このサンプルでは、プレビュー期間中は 1.2 [要件セット](http://dev.office.com/reference/add-ins/office-add-in-requirement-sets?product=word)が使用されます。ただし、1.3 要件セットが一般的に使用可能になれば、1.3 要件セットが必要になります。

### <a name="task-pane"></a>作業ウィンドウ

作業ウィンドウの機能は、sample.js で設定されます。sample.js には、以下の機能が含まれます。

* UI とイベント ハンドラーを設定します。
* サービスから仕様テンプレートをフェッチし、それをドキュメントに挿入します。
* ドキュメントを検証するために使用される単語が含まれているブラックリストを読み込みます。これらの単語は、このサンプルの目的に対して不適切な単語と見なされます。
* サービスから既定の定型を読み込み、ローカル ストレージにそれらをキャッシュします。
* 関数ファイル コードをテストするためのスケルトン コード。関数ファイルにデバッガーをアタッチすることができないため、アドイン コマンド コードを関数ファイルに移動する前に、作業ウィンドウでそれらのコードを開発することをお勧めします。
* 作業ウィンドウにドキュメント プロパティから既定の作成者名をロードします。これにより、ドキュメントのカスタム XML 部分にアクセスして変更する方法が示されます。
* サービスに定型を書き込みます。

### <a name="document-validation-and-the-dialog-api"></a>ドキュメントの検証とダイアログ API

validation.js には、単語のブラックリストに照らし合わせてドキュメントを検証するコードが含まれています。validateContentAgainstBlacklist() メソッドでは、新しい splitTextRanges メソッドを使用して、区切り記号に基づいた範囲にドキュメントを分割します。この関数の区切り記号は、ドキュメント内の単語を識別します。ブラックリストとドキュメント内の単語の共通部分を識別してローカル ストレージにそれらの結果を渡します。displayDialogAsync メソッドを使用して、ダイアログ (dialog.html) を開きます。ダイアログは、ローカル ストレージから検証結果を取得し、その結果を表示します。

### <a name="boilerplate-text-functionality"></a>定型句の機能

boilerplate.js には、ローカル ストレージに定型句を保存する方法、保存した定型に対応するエントリを持つファブリック ドロップダウンを更新する方法、ドロップダウンから選択した定型を挿入する例が含まれています。このファイルには、以下の例が含まれています。
* splitTextRanges (WordApi 1.3 要件セットの新規) - この API は将来のリリースで split() に置き換えられます。
* compareLocationWith (WordApi 1.3 要件セットの新規)
* 新しいエントリで、ファブリック ドロップダウンを更新します。
* 定型句をドキュメントに挿入します。

### <a name="custom-xml-binding-to-core-document-properties"></a>コア ドキュメント プロパティへのカスタム XML バインディング

authorCustomXml.js には、既定のドキュメント プロパティを取得して設定するためのメソッドが含まれています。

* 作業ウィンドウが読み込まれるときに、作成者プロパティを作業ウィンドウに読み込みます。ドキュメントにも作成者プロパティの値が含まれていることに注意してください。これは、テンプレートにこのドキュメント プロパティにバインドされているコンテンツ コントロールが含まれているためです。これにより、カスタム XML 部分の内容に基づいてドキュメントの既定値を設定することができます。
* 作業ウィンドウから作成者プロパティを更新します。この操作により、ドキュメント プロパティとドキュメントのバインドされたコンテンツ コントロールの両方が更新されます。

### <a name="add-in-commands"></a>アドイン コマンド

アドイン コマンドの宣言は、word-add-in-javascript-speckit-manifest.xml にあります。このサンプルでは、リボンやコンテキスト メニューにアドイン コマンドを作成する方法が示されます。

## <a name="questions-and-comments"></a>質問とコメント

Word SpecKit サンプルについて、Microsoft にフィードバックをお寄せください。このリポジトリの「*問題*」セクションにフィードバックを送信できます。

Microsoft Office 365 開発全般の質問につきましては、「[スタック オーバーフロー](http://stackoverflow.com/questions/tagged/office-js+API)」に投稿してください。質問には、必ず [office-js] と [API] のタグを付けてください。

## <a name="additional-resources"></a>その他の技術情報

* [Office アドインのドキュメント](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Office デベロッパー センター](http://dev.office.com/)
* [Office 365 API スタート プロジェクトとコード サンプル](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)

## <a name="copyright"></a>著作権
Copyright (c) 2016 Microsoft Corporation.All rights reserved.



このプロジェクトでは、[Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/) が採用されています。詳細については、「[Code of Conduct の FAQ](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。また、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までお問い合わせください。
