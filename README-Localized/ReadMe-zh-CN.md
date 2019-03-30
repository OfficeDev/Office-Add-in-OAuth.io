# <a name="archived-office-add-in-that-uses-the-oauthio-service-to-get-authorization-to-external-services"></a>[已存档] Office 外接程序使用 OAuth.io 服务来获取外部服务的授权

> **注意：** 此存储库已存档，不再主动维护。 项目或其依赖项中可能存在安全漏洞。 如果计划从此存储库重用或运行任何代码，请务必首先对代码或依赖项执行适当的安全检查。 请勿将此项目用作生产 Office 外接程序的起点。 始终使用 Visual Studio 中的 Office/SharePoint 开发工作负载或 [Office 外接程序的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)启动生产代码，并在开发外接程序时遵循安全最佳做法。

OAuth.io 服务简化了从热门联机服务（如 Facebook 和 Google） 获取 OAuth 2.0 访问令牌的过程。本示例介绍如何在 Office 外接程序中使用它。本示例介绍如何在 Office 外接程序中使用它。 

## <a name="table-of-contents"></a>目录
* [修订记录](#change-history)
* [先决条件](#prerequisites)
* [配置项目](#configure-the-project)
* [下载 OAuth.io JavaScript 库](#download-the-oauth-io-javascript-library)
* [创建 Google 项目](#create-a-google-project)
* [创建 Facebook 应用](#create-a-facebook-app)
* [创建一个 OAuth.io 应用，并将其配置为使用 Google 和 Facebook](#create-an-OAuth-io-app-and-configure-it-to-use-google-and-facebook)
* [向示例代码添加公钥](#add-the-public-key-to-the-sample-code)
* [部署外接程序](#deploy-the-add-in)
* [运行项目](#run-the-project)
* [启动外接程序](#start-the-add-in)
* [测试外接程序](#test-the-add-in)
* [问题和意见](#questions-and-comments)
* [其他资源](#additional-resources)

## <a name="change-history"></a>更改历史记录

2017 年 1 月 5 日：

* 首版。

## <a name="prerequisites"></a>先决条件

* 可使用 [OAuth.io](https://oauth.io/home) 的帐户
* 可使用 [Google](https://developers.google.com/) 和 [Facebook](https://developers.facebook.com/) 的开发人员帐户。
* Word 2016 for Windows（内部版本 16.0.6727.1000 或更高版本）。
* [节点和 npm](https://nodejs.org/en/) 将项目配置为使用 npm 作为程序包管理器和任务运行程序。还可以配置为将 Lite Server 用作开发期间可托管外接程序的 Web 服务器，以便快速启动并运行外接程序。完全可以使用其他任务运行程序或 Web 服务器。
* [Git Bash](https://git-scm.com/downloads)（或其他 git 客户端）。

## <a name="configure-the-project"></a>配置项目

在要放置项目的文件夹中，于 git bash shell 中运行以下命令：

1. ```git clone {URL of this repo}```（将此存储库克隆到本地计算机。）
2. ```npm install```（安装 package.json 文件中列出明细的所有依赖项。）
3. ```bash gen-cert.sh```（创建要运行此示例所需的证书。） 

将此证书设置为受信任的根证书颁发机构。Windows 计算机上的设置步骤：

1. 在本地计算机的存储库文件夹中，双击 ca.crt，然后选择“**安装证书**”。 
2. 选择“**本地计算机**”，然后选择“**下一步**”以继续。 
3. 选择“**将所有证书放入下列存储**”，然后选择“**浏览**”。
4. 选择“**受信任的根证书颁发机构**”，然后选择“**确定**”。 
5. 依次选择“**下一步**”、“**完成**”。 

## <a name="download-the-oauthio-javascript-library"></a>下载 OAuth.io JavaScript 库

2. 在项目中，在“Scripts”文件夹下创建名为 OAuth.io 的子文件夹。
3. 从 [oauth-js](https://github.com/oauth-io/oauth-js) 下载 OAuth.io JavaScript 库。
2. 在库的 \dist 文件夹中，将 oauth.js 或 oauth.min.js 复制到项目的 \Scripts\OAuth.io 文件夹。 
3. 如果在上述步骤中选择了 oauth.min.js，则打开文件 \Scripts\popupRedirect.js，并将行 `url: 'Scripts/OAuth.io/oauth.js',` 更改为 `url: 'Scripts/OAuth.io/oauth.min.js',`。

## <a name="create-a-google-project"></a>创建 Google 项目

1. 在你的 [Google 开发人员控制台](https://console.developers.google.com/apis/dashboard)上，创建一个名为“**Office 外接程序**”的项目。
2. 为项目启用 Google Plus API。有关详细信息，请参阅 Google 的帮助。
3. 为项目创建应用凭据，并选择“**OAuth 客户端 ID**”作为凭据类型，选择“**Web 应用程序**”作为应用程序类型。将“Office-Add-in-OAuth.io”用作客户端**名称**（如果系统提示指定产品名，则将其也用作产品名）。
4. 将“**授权重定向 URI**”设置为 ```https://oauth.io/auth```。
5. 创建客户端密码和已分配的客户端 ID 的副本。

## <a name="create-a-facebook-app"></a>创建 Facebook 应用

1. 在你的 [Facebook 开发人员仪表板](https://developers.facebook.com/)上，创建一个名为“**Office 外接程序**”的应用。有关详细信息，请参阅 Facebook 帮助。
3. 创建客户端密码和已分配的客户端 ID 的副本。 
4. 在“**应用域**”框中指定 ```oauth.io```，并将 ```https://oauth.io/``` 指定为**网站 URL**。 

## <a name="create-an-oauthio-app-and-configure-it-to-use-google-and-facebook"></a>创建一个 OAuth.io 应用，并将其配置为使用 Google 和 Facebook

1. 在你的 OAuth.io 仪表板中，创建一个应用，并记下 OAuth.io 向其分配的**公钥**。
3. 如果尚不存在，则将 `localhost` 添加到应用的允许列表 URL 中。
2. 将 **Google Plus** 作为提供程序（也称为“集成 API”）添加到应用中。填写在创建 Google 项目时获取的客户端 ID 和密码，然后将 `profile` 指定为**范围**。如果在 IE 中未保存对提供程序（集成 API）的配置，请尝试在其他浏览器中重新打开 OAuth.io 仪表板。
3. 对 **Facebook** 重复上述步骤，但将 `user_about_me` 指定为**范围**。

## <a name="add-the-public-key-to-the-sample-code"></a>将公钥添加到示例代码中

在文件 popup.js 中调用 API `OAuth.initialize`。查找行，并将上一部分中获得的公钥以字符串形式传递给该函数。

## <a name="deploy-the-add-in"></a>部署外接程序

现在需要让 Microsoft Word 知道在哪里可以找到外接程序。

1. 创建网络共享，或[将文件夹共享到网络](https://technet.microsoft.com/zh-cn/library/cc770880.aspx)。
2. 将 Office-Add-in-OAuth.io.xml 清单文件从项目根目录复制到共享文件夹。
3. 启动 Word，然后打开一个文档。
4. 选择**文件**选项卡，然后选择**选项**。
5. 选择**信任中心**，然后选择**信任中心设置**按钮。
6. 选择“**受信任的外接程序目录**”。
7. 在“**目录 URL**”字段中，输入包含 Office-Add-in-OAuth.io.xml 的文件夹共享的网络路径，然后选择“**添加目录**”。
8. 选中“**显示在菜单中**”复选框，然后单击“**确定**”。
9. 随后会出现一条消息，告知你下次启动 Microsoft Office 时将应用你的设置。关闭 Word。

## <a name="run-the-project"></a>运行项目

1. 打开项目文件夹中的节点命令窗口，然后运行 ```npm start``` 来启动 Web 服务。使命令窗口保持打开状态。
2. 打开 Internet Explorer 或 Edge，然后在地址框中输入 ```https://localhost:3000```。如果未收到有关证书的任何警告，则关闭浏览器，然后继续执行下面的“**启动外接程序**”部分。如果看到提示证书不受信任的警告，请继续按以下步骤操作：
3. 除警告外，浏览器还会提供一个可以打开该页面的链接。打开该页面。
4. 打开页面后，地址栏中会有一条显示为红色的证书错误消息。双击此错误。
5. 选择“**查看证书**”。
5. 选择“**安装证书**”。
4. 选择“**本地计算机**”，然后选择“**下一步**”以继续。 
3. 选择“**将所有证书放入下列存储**”，然后选择“**浏览**”。
4. 选择“**受信任的根证书颁发机构**”，然后选择“**确定**”。 
5. 依次选择“**下一步**”、“**完成**”。
6. 关闭浏览器。

## <a name="start-the-add-in"></a>启动外接程序

1. 重新启动 Word 并打开一个 Word 文档。
2. 在 Word 2016 中的“**插入**”选项卡上，选择“**我的外接程序**”。
3. 选择“**共享文件夹**”选项卡。
4. 选择“**使用 OAuthIO 进行身份验证**”，然后选择“**确定**”。
5. 如果 Word 版本支持外接程序命令，UI 将通知你已加载外接程序。
6. 主页功能区上有一个名为“**OAuth.io**”的新组，包含标记为“**显示**”的按钮和一个图标。单击该按钮，打开此外接程序。

 > 注意：如果你的 Word 版本不支持外接程序命令，则外接程序将在任务窗格中加载。

## <a name="test-the-add-in"></a>测试外接程序

1. 单击“**获取 Facebook 名称**”按钮。
2. 弹出窗口将打开，系统将提示你登录到 Facebook（除非你已登录）。
3. 登录后，你的 Facebook 用户名会被插入到 Word 文档。
4. 使用“**获取 Google 名称**”按钮重复以上步骤。

## <a name="questions-and-comments"></a>问题和意见

我们乐意倾听你对此示例的反馈。你可以在此存储库中的“*问题*”部分向我们发送反馈。

与 Microsoft Office 365 开发相关的一般问题应发布到 [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API)。如果你的问题是关于 Office JavaScript API，请务必为问题添加 [office-js] 和 [API].标记。

## <a name="additional-resources"></a>其他资源

* 
  [Office 外接程序文档](https://msdn.microsoft.com/zh-cn/library/office/jj220060.aspx)
* [Office 开发人员中心](http://dev.office.com/)
* 有关更多 Office 外接程序示例，请访问 [Github 上的 OfficeDev](https://github.com/officedev)。

## <a name="copyright"></a>版权信息
版权所有 (c) 2017 Microsoft Corporation。保留所有权利。

