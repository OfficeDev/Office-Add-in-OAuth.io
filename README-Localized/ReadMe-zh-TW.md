# <a name="archived-office-add-in-that-uses-the-oauthio-service-to-get-authorization-to-external-services"></a>[已封存] 使用 OAuth.io 服務以取得外部服務驗證的 Office 增益集

> **注意：** 此 Repo 已封存，不再主動進行維護。 安全性漏洞可能存在專案或其相依專案。 如果您計畫重複使用或執行此Repo 的任何程式碼，請務必對程式碼或其相依程式碼執行適當的安全性檢查。 請勿將此專案當成實際 Office 增益集的起點。 一律使用 Visual Studio 中的 Office/SharePoint 開發工作負載，或是使用 [Office 增益集的 Yeoman 產生器](https://github.com/OfficeDev/generator-office)，開始您的實際程式碼，然後在開發增益集的同時，遵循最佳安全性的作法。

OAuth.io 服務可以簡化從受歡迎的線上服務 (例如 Facebook 和 Google)，取得 OAuth 2.0 存取權杖的程序。這個範例會示範如何在 Office 增益集中使用 OAuth.io 服務。 

## <a name="table-of-contents"></a>目錄
* [變更歷程記錄](#change-history)
* [必要條件](#prerequisites)
* [設定專案](#configure-the-project)
* [下載 OAuth.io JavaScript 程式庫](#download-the-oauth-io-javascript-library)
* [建立 Google 專案](#create-a-google-project)
* [建立 Facebook 應用程式](#create-a-facebook-app)
* [建立 OAuth.io 應用程式，並將其設定為使用 Google 和 Facebook](#create-an-OAuth-io-app-and-configure-it-to-use-google-and-facebook)
* [將公開金鑰新增到範例程式碼中](#add-the-public-key-to-the-sample-code)
* [部署增益集](#deploy-the-add-in)
* [執行專案](#run-the-project)
* [啟動增益集](#start-the-add-in)
* [測試增益集](#test-the-add-in)
* [問題和建議](#questions-and-comments)
* [其他資源](#additional-resources)

## <a name="change-history"></a>變更歷程記錄

2017 年 1 月 5 日：

* 初始版本。

## <a name="prerequisites"></a>必要條件

* 具有 [OAuth.io](https://oauth.io/home) 的帳戶
* 具有 [Google](https://developers.google.com/) 和 [Facebook](https://developers.facebook.com/) 的開發人員帳戶。
* Word 2016 for Windows，組建 16.0.6727.1000 或更新版本。
* [節點和 npm](https://nodejs.org/en/) 專案設定為使用 npm 作為封裝管理員和工作執行器。它也設定為使用精簡版伺服器作為 Web 伺服器，在開發期間裝載增益集，因此您可以快速地啟動並執行增益集。您也可以自由地使用其他工作執行器或 Web 伺服器。
* [就可以給艦隊](https://git-scm.com/downloads) (或其他 git 用戶端。)

## <a name="configure-the-project"></a>設定專案

在您要放置專案的資料夾中，以 git bash shell 執行下列命令︰

1. ```git clone {URL of this repo}``` 可複製此儲存機制到本機電腦。
2. ```npm install``` 可安裝 package.json 檔案中的所有分項相依性。
3. ```bash gen-cert.sh``` 可建立執行這個範例所需的憑證。 

將憑證設定為受信任的根授權。在 Windows 電腦上的步驟如下︰

1. 在您本機電腦上的儲存機制資料夾中，連按兩下 ca.crt，然後選取 [安裝憑證]****。 
2. 選取 [本機電腦]****，然後選取 [下一步]**** 以繼續。 
3. 選取 [將所有憑證放入以下的存放區]****，然後選取 [瀏覽]****。
4. 選取 [信任的根憑證授權]****，然後選取 [確定]****。 
5. 選取 [下一步]****，然後選取 [完成]****。 

## <a name="download-the-oauthio-javascript-library"></a>下載 OAuth.io JavaScript 程式庫

2. 在專案中的「指令碼」資料夾下，建立稱為「OAuth.io」的子資料夾。
3. 從 [oauth js](https://github.com/oauth-io/oauth-js) 下載 OAuth.io JavaScript 程式庫。
2. 從程式庫的 \dist 資料夾中，將 oauth.js 或 oauth.min.js 複製到專案的 \Scripts\OAuth.io 資料夾內。 
3. 如果您在前一個步驟中選擇了 oauth.min.js，請開啟檔案 \Scripts\popupRedirect.js，並變更行 `url: 'Scripts/OAuth.io/oauth.js',` 到 `url: 'Scripts/OAuth.io/oauth.min.js',`。

## <a name="create-a-google-project"></a>建立 Google 專案

1. 在您 [Google 開發人員主控台](https://console.developers.google.com/apis/dashboard) 中，建立名為 [Office 增益集]**** 的專案。
2. 啟用專案的 Google Plus API。如需詳細資訊，請參閱 Google 的說明。
3. 建立專案的應用程式認證，然後認證類型選擇 [OAuth 用戶端 ID]****，而應用程式類型選擇 [Web 應用程式]****。使用「Office-Add-in-OAuth.io」作為用戶端 [名稱]**** (以及產品名稱，如果系統提示您指定一個產品名稱時)。
4. 將 [授權的重新導向 URI]**** 設定為 ```https://oauth.io/auth```。
5. 針對用戶端的密碼及其指派給的用戶端 ID 製作複本。

## <a name="create-a-facebook-app"></a>建立 Facebook 應用程式

1. 在您的 [Facebook 開發人員儀表板](https://developers.facebook.com/) 中，建立名為 [Office 增益集]**** 的應用程式。如需詳細資訊，請參閱 Facebook 的說明。
3. 針對用戶端的密碼及其指派給的用戶端 ID 製作複本。 
4. 在 [應用程式網域]**** 方塊中指定 ```oauth.io```，並將 [網站 URL]**** 指定為 ```https://oauth.io/```。 

## <a name="create-an-oauthio-app-and-configure-it-to-use-google-and-facebook"></a>建立 OAuth.io 應用程式，並將其設定為使用 Google 和 Facebook

1. 在 [OAuth.io 儀表板] 中建立應用程式，並記下 OAuth.io 指派給它的 [公開金鑰]****。
3. 如果 `localhost` 不在允許清單中，請將其新增到應用程式的允許清單 URL 中。
2. 將 **Google Plus** 新增為應用程式的供應商 (也叫做「整合式 API」)。填入用戶端 ID 及建立該 Google 專案時的密碼，然後將 [範圍]**** 指定為 `profile`。如果供應商 (整合式 API) 組態無法儲存在 IE 中，請嘗試在另一個瀏覽器中，重新開啟 [OAuth.io 儀表板]。
3. 針對 **Facebook** 重複上述步驟，但將 [範圍]**** 指定為 `user_about_me`。

## <a name="add-the-public-key-to-the-sample-code"></a>將公開金鑰新增到範例程式碼中

在檔案 popup.js 中會呼叫 API `OAuth.initialize`。找出資料行，並將上一節中所取得的公開金鑰，當作這個函數的字串傳遞。

## <a name="deploy-the-add-in"></a>部署增益集

現在，您需要讓 Microsoft Word 知道哪裡可以找到此增益集。

1. 建立網路共用，或 [在網路上共用資料夾](https://technet.microsoft.com/zh-tw/library/cc770880.aspx)。
2. 將一份 Office-Add-in-OAuth.io.xml 資訊清單檔，從專案的根目錄放入共用資料夾中。
3. 啟動 Word 並開啟一個文件。
4. 選擇 [檔案]**** 索引標籤，然後選擇 [選項]****。
5. 選擇 [信任中心]****，然後選擇 [信任中心設定]**** 按鈕。
6. 選擇 [受信任的增益集目錄]****。
7. 在 [目錄 URL]**** 欄位中，輸入包含 Office-Add-in-OAuth.io.xml 的資料夾共用的網路路徑，然後選擇 [新增目錄]****。
8. 選取 [顯示於功能表中]**** 核取方塊，然後選擇 [確定]****。
9. 接著會顯示訊息，通知您下次啟動 Microsoft Office 時就會套用您的設定。關閉 Word。

## <a name="run-the-project"></a>執行專案

1. 開啟專案的資料夾中節點的命令視窗，並執行 ```npm start``` 以啟動 Web 服務。保留命令視窗開啟。
2. 開啟 Internet Explorer 或 Microsoft Edge，並在網址方塊中輸入 ```https://localhost:3000```。如果您未收到與憑證相關的任何警告，請關閉瀏覽器，並繼續進行下面主題為**啟動增益集**的章節。如果您收到憑證不受信任的警告，請繼續執行下列步驟︰
3. 儘管有警告，瀏覽器還是可以給予您用以開啟頁面的連結。將其開啟。
4. 開啟網頁後，在網址列中會有紅色的憑證錯誤訊息。按兩下錯誤。
5. 選取 [檢視憑證]****。
5. 選取 [安裝憑證]****。
4. 選取 [本機電腦]****，然後選取 [下一步]**** 以繼續。 
3. 選取 [將所有憑證放入以下的存放區]****，然後選取 [瀏覽]****。
4. 選取 [信任的根憑證授權]****，然後選取 [確定]****。 
5. 選取 [下一步]****，然後選取 [完成]****。
6. 關閉瀏覽器。

## <a name="start-the-add-in"></a>啟動增益集

1. 重新啟動 Word，並開啟 Word 文件。
2. 在 Word 2016 的 [插入]**** 索引標籤上，選擇 [我的增益集]****。
3. 選取 [共用資料夾]**** 索引標籤。
4. 選擇 [以 OAuthIO 驗證]****，然後選取 [確定]****。
5. 如果您的 Word 版本支援增益集命令，UI 會通知您已載入增益集。
6. 在 [常用] 功能區中的是稱為 **OAuth.io** 的新群組，具有標示為 [顯示]**** 的按鈕和圖示。按一下該按鈕以開啟增益集。

 > 附註：如果您的 Word 版本不支援增益集命令，增益集會載入工作窗格。

## <a name="test-the-add-in"></a>測試增益集

1. 按一下 [取得 Facebook 名稱]**** 按鈕。
2. 快顯功能表隨即開啟，系統會提示您登入 Facebook (除非您已經登入)。
3. 登入後，系統會將您的 Facebook 使用者名稱插入 Word 文件中。
4. 針對 [取得 Google 名稱]**** 按鈕重複上述步驟。

## <a name="questions-and-comments"></a>問題和建議

我們很樂於收到您對於此範例的意見反應。您可以在此存放庫的 [問題]** 區段中，將您的意見反應傳送給我們。

請在 [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API) 提出有關 Microsoft Office 365 開發的一般問題。如果您的問題是關於 Office JavaScript API，請確定您的問題標記有 [office js] 與 [API]。

## <a name="additional-resources"></a>其他資源

* 
  [Office 增益集文件](https://msdn.microsoft.com/zh-tw/library/office/jj220060.aspx)
* [Office 開發人員中心](http://dev.office.com/)
* 在 [Github 上的 OfficeDev](https://github.com/officedev) 中有更多 Office 增益集範例

## <a name="copyright"></a>著作權
Copyright (c) 2017 Microsoft Corporation.著作權所有，並保留一切權利。

