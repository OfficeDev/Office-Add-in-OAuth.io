# <a name="archived-office-add-in-that-uses-the-oauthio-service-to-get-authorization-to-external-services"></a>[アーカイブ済み] OAuth.io サービスを使用して外部サービスへの承認を得る Office アドイン

> **注:** このリポジトリはアーカイブされており、アクティブなメンテナンスは終了しています。 プロジェクトまたはその依存関係にはセキュリティの脆弱性が存在する可能性があります。 このリポジトリのコードを再利用するか、実行することを計画している場合は、必ず最初にコードまたは依存関係に対して適切なセキュリティ チェックを実行してください。 運用 Office アドインの開始点として、このプロジェクトを使わないでください。 運用コードを開始する際は、必ず Visual Studio の Office/SharePoint 開発ワークロードか [Office アドイン用 Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)を使用するようにし、アドインを開発する際はセキュリティのベスト プラクティスに従ってください。

OAuth.io サービスは、Facebook や Google などの人気のあるオンライン サービスから OAuth 2.0 アクセス トークンを取得するプロセスを簡略化します。このサンプルでは、Office アドインでこのサービスを使用する方法を示します。 

## <a name="table-of-contents"></a>目次
* [変更履歴](#change-history)
* [前提条件](#prerequisites)
* [プロジェクトを構成する](#configure-the-project)
* [OAuth.io JavaScript ライブラリをダウンロードする](#download-the-oauth-io-javascript-library)
* [Google プロジェクトを作成する](#create-a-google-project)
* [Facebook アプリを作成する](#create-a-facebook-app)
* [OAuth.io アプリを作成し、Google と Facebook を使用できるように構成する](#create-an-OAuth-io-app-and-configure-it-to-use-google-and-facebook)
* [サンプル コードに公開キーを追加する](#add-the-public-key-to-the-sample-code)
* [アドインを展開する](#deploy-the-add-in)
* [プロジェクトを実行する](#run-the-project)
* [アドインを起動する](#start-the-add-in)
* [アドインをテストする](#test-the-add-in)
* [質問とコメント](#questions-and-comments)
* [その他のリソース](#additional-resources)

## <a name="change-history"></a>変更履歴

2017 年 1 月 5 日:

* 初期バージョン。

## <a name="prerequisites"></a>前提条件

* [OAuth.io](https://oauth.io/home) を使用したアカウント
* [Google](https://developers.google.com/) と [Facebook](https://developers.facebook.com/) を使用した開発者アカウント。
* Windows 用の Word 2016 (16.0.6727.1000 以降のビルド)。
* [Node と npm](https://nodejs.org/en/) プロジェクトはパッケージ マネージャーとタスク ランナーの両方として npm を使用するように構成されます。また、開発中にアドインをホストする Web サーバーとして Lite サーバーを使用するようにも構成されるため、アドインをすばやくオンにして実行することができます。別のタスク ランナーまたは Web サーバーを使用することもできます。
* [Git バッシュ](https://git-scm.com/downloads) (またはその他の Git クライアント。)

## <a name="configure-the-project"></a>プロジェクトを構成する

プロジェクトを配置するフォルダーで、Git バッシュ シェルで次のコマンドを実行します。

1. ローカル コンピューターにこのリポジトリのクローンを作成する ```git clone {URL of this repo}```
2. package.json ファイル内のアイテム化されたすべての依存関係をインストールする ```npm install```。
3. このサンプルを実行するために必要な証明書を作成する ```bash gen-cert.sh```。 

証明書を信頼されたルート機関にするように設定します。Windows コンピューターでの手順は次のとおりです。

1. ローカル コンピューターにあるリポジトリ フォルダーで、ca.crt をダブルクリックし、**[証明書のインストール]** を選択します。 
2. **[ローカル コンピューター]** を選択して、**[次へ]** を選択して続行します。 
3. **[証明書をすべて次のストアに配置する]** を選択してから **[参照]** を選択します。
4. **[信頼されたルート証明機関]** を選択して、**[OK]** を選択します。 
5. **[次へ]**、**[完了]** の順に選択します。 

## <a name="download-the-oauthio-javascript-library"></a>OAuth.io JavaScript ライブラリをダウンロードする

2. プロジェクトの [スクリプト] フォルダーに、OAuth.io という名前のサブフォルダーを作成します。
3. [oauth-js](https://github.com/oauth-io/oauth-js) から OAuth.io JavaScript ライブラリをダウンロードします。
2. ライブラリの \dist フォルダーにある oauth.js または oauth.min.js を、プロジェクトの \Scripts\OAuth.io フォルダーにコピーします。 
3. 前の手順で oauth.min.js を選択した場合は、ファイル \Scripts\popupRedirect.js を開き、行 `url: 'Scripts/OAuth.io/oauth.js',` を `url: 'Scripts/OAuth.io/oauth.min.js',` に変更します。

## <a name="create-a-google-project"></a>Google プロジェクトを作成する

1. [[Google 開発者コンソール]](https://console.developers.google.com/apis/dashboard) で、**Office アドイン**という名前のプロジェクトを作成します。
2. プロジェクトの Google Plus API を有効にします。詳細については、Google のヘルプをご覧ください。
3. プロジェクトのアプリ資格情報を作成し、資格情報の種類で **[OAuth クライアント ID]**、アプリケーションの種類で **[Web アプリケーション]** をそれぞれ選択します。クライアントの **[名前]** には "Office-Add-in-OAuth.io" を使用します (製品名を指定するようにプロンプトが表示された場合も、同じ名前を使用します)。
4. **[承認されているリダイレクト URI]** を ```https://oauth.io/auth``` に設定します。
5. 割り当てられているクライアント シークレットとクライアント ID のコピーを作成します。

## <a name="create-a-facebook-app"></a>Facebook アプリを作成する

1. [[Facebook 開発者ダッシュボード]](https://developers.facebook.com/) で、**Office アドイン**という名前のアプリを作成します。詳細については Facebook のヘルプをご覧ください。
3. 割り当てられているクライアント シークレットとクライアント ID のコピーを作成します。 
4. **[アプリ ドメイン]** ボックスで ```oauth.io``` を指定し、**[サイト URL]** で ```https://oauth.io/``` を指定します。 

## <a name="create-an-oauthio-app-and-configure-it-to-use-google-and-facebook"></a>OAuth.io アプリを作成し、Google と Facebook を使用できるように構成する

1. OAuth.io ダッシュボードでアプリを作成し、OAuth.io で割り当てられた**公開キー**をメモします。
3. 割り当てられていない場合は、アプリのホワイトリスト URL に `localhost` を追加します。
2. プロバイダー ("統合 API" とも呼ばれる) として **Google Plus** をアプリに追加します。Google プロジェクトを作成したときに取得したクライアント ID およびクライアント シークレットを入力してから、`profile` を**スコープ**として指定します。プロバイダー (統合 API) の構成が IE に保存されない場合は、OAuth.io ダッシュボードを別のブラウザーで再度開いてみてください。
3. **Facebook** でも前述の手順を繰り返しますが、**スコープ**には `user_about_me` を指定します。

## <a name="add-the-public-key-to-the-sample-code"></a>サンプル コードに公開キーを追加する

API `OAuth.initialize` はファイル popup.js で呼び出されます。行を検索し、前のセクションで取得した公開キーを、この関数の文字列として渡します。

## <a name="deploy-the-add-in"></a>アドインを展開する

次に、Microsoft Word がアドインを検索する場所を認識できるようにする必要があります。

1. ネットワーク共有を作成するか、[フォルダーをネットワークに共有します](https://technet.microsoft.com/ja-jp/library/cc770880.aspx)。
2. プロジェクトのルートから、Office-Add-in-OAuth.io.xml マニフェスト ファイルのコピーを共有フォルダーに配置します。
3. Word を起動し、ドキュメントを開きます。
4. [**ファイル**] タブを選択し、[**オプション**] を選択します。
5. [**セキュリティ センター**] を選択し、[**セキュリティ センターの設定**] ボタンを選択します。
6. **[信頼されているアドイン カタログ]** を選択します。
7. **[カタログの URL]** フィールドに、Office-Add-in-OAuth.io.xml があるフォルダー共有へのネットワーク パスを入力して、**[カタログの追加]** を選択します。
8. **[メニューに表示する]** チェック ボックスをオンにし、**[OK]** を選択します。
9. これらの設定は Microsoft Office を次回起動したときに適用されることを示すメッセージが表示されます。Word を終了します。

## <a name="run-the-project"></a>プロジェクトを実行する

1. プロジェクトのフォルダー内でノード コマンド ウィンドウを開き、```npm start``` を実行して Web サービスを開始します。コマンド ウィンドウを開いたままにしておきます。
2. Internet Explorer または Microsoft Edge を開いて、```https://localhost:3000``` をアドレス ボックスに入力します。証明書に関する警告が表示されない場合は、ブラウザーを閉じて、「**アドインを起動する**」というタイトルのセクションに進みます。証明書が信頼されていないという警告が表示された場合は、以下の手順に進みます。
3. 警告があっても、ブラウザーにはページを開くためのリンクが表示されます。そのリンクを開きます。
4. ページが開いたら、アドレス バーに赤い証明書エラーが表示されます。エラーをダブルクリックします。
5. **[証明書の表示]** を選択します。
5. **[証明書のインストール]** を選択します。
4. **[ローカル コンピューター]** を選択して、**[次へ]** を選択して続行します。 
3. **[証明書をすべて次のストアに配置する]** を選択してから **[参照]** を選択します。
4. **[信頼されたルート証明機関]** を選択して、**[OK]** を選択します。 
5. **[次へ]**、**[完了]** の順に選択します。
6. ブラウザーを閉じます。

## <a name="start-the-add-in"></a>アドインを起動する

1. Word を再起動して、Word 文書を開きます。
2. Word 2016 の **[挿入]** タブで、**[マイ アドイン]** を選択します。
3. **[共有フォルダー]** タブを選択します。
4. **[OAuthIO で認証]** を選択し、**[OK]** を選択します。
5. ご使用の Word バージョンでアドイン コマンドがサポートされている場合、UI によってアドインが読み込まれたことが通知されます。
6. [ホーム] では、リボンは **OAuth.io** と呼ばれる新しいグループであり、**[表示]** というラベル付きのボタンとアイコンが用意されています。そのボタンをクリックして、アドインを開きます。

 > 注:アドイン コマンドが Word バージョンによってサポートされていない場合は、アドインが作業ウィンドウに読み込まれます。

## <a name="test-the-add-in"></a>アドインをテストする

1. **[Facebook の名前を取得する]** ボタンをクリックします。
2. ポップアップ画面が開き、Facebook を使用してサインインするよう求められます (すでに Facebook を使用している場合を除く)。
3. サインインすると、自分の Facebook ユーザー名が Word 文書に挿入されます。
4. **[Google の名前を取得する]** ボタンを使用して、上記の手順を繰り返します。

## <a name="questions-and-comments"></a>質問とコメント

このサンプルに関するフィードバックをお寄せください。このリポジトリの「*問題*」セクションでフィードバックを送信できます。

Microsoft Office 365 開発全般の質問につきましては、「[スタック オーバーフロー](http://stackoverflow.com/questions/tagged/office-js+API)」に投稿してください。Office JavaScript API に関する質問の場合は、必ず質問に [office-js] と [API] のタグを付けてください。

## <a name="additional-resources"></a>追加リソース

* 
  [Office アドインのドキュメント](https://msdn.microsoft.com/ja-jp/library/office/jj220060.aspx)
* [Office デベロッパー センター](http://dev.office.com/)
* [Github の OfficeDev](https://github.com/officedev) にあるその他の Office アドイン サンプル

## <a name="copyright"></a>著作権
Copyright (c) 2017 Microsoft Corporation.All rights reserved.

