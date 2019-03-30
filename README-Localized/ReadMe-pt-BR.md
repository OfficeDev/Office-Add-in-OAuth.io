# <a name="archived-office-add-in-that-uses-the-oauthio-service-to-get-authorization-to-external-services"></a>[ARQUIVADO] Suplemento do Office que usa o serviço Auth0 para obter autorização para serviços externos

> **Observação:** Este repositório foi arquivado e não é mais ativamente mantido. Podem existir vulnerabilidades de segurança no projeto ou em suas dependências. Caso pretenda reutilizar ou executar qualquer código deste repositório, certifique-se de executar primeiro verificações de segurança apropriadas no código ou nas dependências. Não use esse projeto como ponto de partida de produção de suplementos do Office. Sempre inicie seu código de produção usando a carga de trabalho de desenvolvimento do Office/SharePoint no Visual Studio ou no [gerador do Yeoman para suplementos do Office](https://github.com/OfficeDev/generator-office) e siga as melhores práticas de segurança ao desenvolver o suplemento.

O serviço OAuth.io simplifica o processo de obter tokens de acesso OAuth 2.0 de serviços online populares, como o Facebook e o Google. O exemplo a seguir mostra como usá-lo em um suplemento do Office. 

## <a name="table-of-contents"></a>Sumário
* [Histórico de Alterações](#change-history)
* [Pré-requisitos](#prerequisites)
* [Configurar o projeto](#configure-the-project)
* [Baixar a biblioteca de JavaScript do OAuth.io](#download-the-oauth-io-javascript-library)
* [Criar um projeto do Google](#create-a-google-project)
* [Criar um aplicativo do Facebook](#create-a-facebook-app)
* [Criar um aplicativo do OAuth.io e configurá-lo para usar o Google e o Facebook](#create-an-OAuth-io-app-and-configure-it-to-use-google-and-facebook)
* [Adicionar a chave pública ao código de exemplo](#add-the-public-key-to-the-sample-code)
* [Implantar o suplemento](#deploy-the-add-in)
* [Executar o projeto](#run-the-project)
* [Iniciar o suplemento](#start-the-add-in)
* [Testar o suplemento](#test-the-add-in)
* [Perguntas e comentários](#questions-and-comments)
* [Recursos adicionais](#additional-resources)

## <a name="change-history"></a>Histórico de Alterações

5 de janeiro de 2017:

* Versão inicial.

## <a name="prerequisites"></a>Pré-requisitos

* Uma conta com [OAuth.io](https://oauth.io/home)
* Contas de desenvolvedor com [Google](https://developers.google.com/) e [Facebook](https://developers.facebook.com/).
* Word 2016 para Windows, build 16.0.6727.1000 ou superior.
* [Node e npm](https://nodejs.org/en/) Configuramos o projeto para usar npm como gerenciador de pacotes e executor de tarefas. É possível configurá-lo também para usar o Lite Server como o servidor Web que hospedará o suplemento durante o desenvolvimento, de modo que você possa começar a usar o suplemento rapidamente. Você também pode usar outro executor de tarefas ou servidor Web.
* [Git Bash](https://git-scm.com/downloads) (ou outro cliente Git.)

## <a name="configure-the-project"></a>Configurar o projeto

Na pasta em que deseja armazenar o projeto, execute os seguintes comandos no shell do Git Bash:

1. ```git clone {URL of this repo}``` para clonar este repositório no computador local.
2. ```npm install``` para instalar todas as dependências discriminadas no arquivo package.json.
3. ```bash gen-cert.sh``` para criar os certificados necessários para executar este exemplo. 

Defina o certificado como uma autoridade raiz confiável. Em um computador com Windows, as etapas são as seguintes:

1. Na pasta do repositório do computador local, clique duas vezes em ca.crt e escolha **Instalar Certificado**. 
2. Escolha **Computador Local** e **Avançar** para continuar. 
3. Escolha **Colocar todos os certificados no repositório a seguir** e **Procurar**.
4. Escolha **Autoridades de Certificação Confiáveis** e **OK**. 
5. Escolha **Avançar** e **Concluir**. 

## <a name="download-the-oauthio-javascript-library"></a>Baixar a biblioteca de JavaScript do OAuth.io

2. No projeto, crie uma subpasta chamada OAuth.io sob a pasta Scripts.
3. Baixar a biblioteca de JavaScript do OAuth.io de [oauth-js](https://github.com/oauth-io/oauth-js).
2. Na pasta \dist da biblioteca, copie oauth.js ou oauth.min.js para a pasta \Scripts\OAuth.io do projeto. 
3. Se você escolheu oauth.min.js na etapa anterior, abra o arquivo \Scripts\popupRedirect.js e altere a linha `url: 'Scripts/OAuth.io/oauth.js',` para `url: 'Scripts/OAuth.io/oauth.min.js',`.

## <a name="create-a-google-project"></a>Criar um projeto do Google

1. Em seu [Console de desenvolvedor do Google](https://console.developers.google.com/apis/dashboard), crie um projeto chamado **Suplemento do Office**.
2. Habilite a API do Google Plus para o projeto. Consulte a Ajuda do Google para obter detalhes.
3. Crie credenciais do aplicativo para o projeto e escolha **ID do cliente OAuth** como o tipo de credencial e **Aplicativo Web** como o tipo de aplicativo. Use "Office-Add-in-OAuth.io" como o **Nome** do cliente (e o nome do produto, se você for solicitado a especificar um).
4. Defina as **URIs de redirecionamento autorizadas** como ```https://oauth.io/auth```.
5. Faça uma cópia do segredo do cliente e da ID de cliente que é atribuída.

## <a name="create-a-facebook-app"></a>Criar um aplicativo do Facebook

1. Em seu [Painel do desenvolvedor do Facebook](https://developers.facebook.com/), crie um aplicativo chamado **Suplemento do Office**. Consulte a Ajuda do Facebook para obter detalhes.
3. Faça uma cópia do segredo do cliente e da ID de cliente que é atribuída. 
4. Especifique ```oauth.io``` na caixa **Domínios do Aplicativo** e especifique ```https://oauth.io/``` como a **URL do Site**. 

## <a name="create-an-oauthio-app-and-configure-it-to-use-google-and-facebook"></a>Criar um aplicativo do OAuth.io e configurá-lo para usar o Google e o Facebook

1. No painel do OAuth.io, crie um aplicativo e anote a **chave pública** que OAuth.io atribui a ele.
3. Se ainda não estiver lá, adicione `localhost` às URLs da lista branca para o aplicativo.
2. Adicione **Google Plus** como um provedor (também chamado de "API integrada") para o aplicativo. Preencha a ID do cliente e o segredo obtido quando você criou o projeto do Google e, em seguida, especifique `profile` como o **escopo**. Se a configuração do provedor – API Integrada – não estiver salva no IE, tente reabrir o painel do OAuth.io em outro navegador).
3. Repita a etapa anterior para o **Facebook**, mas especifique `user_about_me` como o **escopo**.

## <a name="add-the-public-key-to-the-sample-code"></a>Adicione a chave pública ao código de amostra

A API `OAuth.initialize` é chamada no arquivo popup.js. Localize a linha e passe a chave pública obtida na seção anterior como uma cadeia de caracteres para esta função.

## <a name="deploy-the-add-in"></a>Implantar o suplemento

Agora, você precisa informar ao Microsoft Word onde encontrar o suplemento.

1. Crie um compartilhamento de rede ou [compartilhe uma pasta na rede](https://technet.microsoft.com/pt-br/library/cc770880.aspx).
2. Coloque uma cópia do arquivo de manifesto Office-Add-in-OAuth.io.xml na raiz do projeto, dentro da pasta compartilhada.
3. Inicie o Word e abra um documento.
4. Escolha a guia **Arquivo** e escolha **Opções**.
5. Escolha **Central de Confiabilidade**, e escolha o botão **Configurações da Central de Confiabilidade**.
6. Escolha **Catálogos de Suplementos Confiáveis**.
7. No campo **URL do Catálogo**, insira o caminho de rede para o compartilhamento de pasta que contém o arquivo Office-Add-in-OAuth.io.xml e escolha **Adicionar catálogo**.
8. Selecione a caixa de seleção **Mostrar no Menu** e, em seguida, escolha **OK**.
9. O sistema exibirá uma mensagem para informar que suas configurações serão aplicadas na próxima vez que você iniciar o Microsoft Office. Feche o Word.

## <a name="run-the-project"></a>Executar o projeto

1. Abra uma janela de comando de nó na pasta do projeto e execute ```npm start``` para iniciar o serviço Web. Deixe a janela de comando aberta.
2. Abra o Internet Explorer ou o Microsoft Edge e insira ```https://localhost:3000``` na caixa de endereço. Se não receber avisos sobre o certificado, feche o navegador e avance para a seção abaixo chamada **Iniciar o suplemento**. Se receber um aviso informando que o certificado não é confiável, vá para as etapas seguintes:
3. Independentemente do aviso, o navegador fornecerá um link para você abrir a página. Abra-a.
4. Após abri-la, o sistema exibirá um Erro de Certificado vermelho na barra de endereços. Clique duas vezes no erro.
5. Escolha **Exibir Certificado**.
5. Escolha **Instalar Certificado**.
4. Escolha **Computador Local** e **Avançar** para continuar. 
3. Escolha **Colocar todos os certificados no repositório a seguir** e **Procurar**.
4. Escolha **Autoridades de Certificação Confiáveis** e **OK**. 
5. Escolha **Avançar** e **Concluir**.
6. Feche o navegador.

## <a name="start-the-add-in"></a>Iniciar o suplemento

1. Inicie novamente o Word e abra um documento.
2. Na guia **Inserir** no Word 2016, escolha **Meus Suplementos**.
3. Escolha a guia **PASTA COMPARTILHADA**.
4. Escolha **Autenticar com OAuthIO**e selecione **OK**.
5. Se os comandos de suplemento forem compatíveis com sua versão do Word, a interface do usuário informará que o suplemento foi carregado.
6. Na Faixa de Opções da Página Inicial, há um novo grupo chamado **OAuth.io** com um botão **Exibir** e um ícone. Clique no botão para abrir o suplemento.

 > Observação: O suplemento será carregado no painel de tarefas se os comandos de suplemento não forem compatíveis com sua versão do Word.

## <a name="test-the-add-in"></a>Testar o suplemento

1. Clique no botão **Obter Nome do Facebook**.
2. Um pop-up será aberto e você será solicitado a entrar com o Facebook (a menos que já esteja lá).
3. Depois de entrar, seu nome de usuário do Facebook será inserido no documento do Word.
4. Repita as etapas acima com o botão **Obter Nome do Google**.

## <a name="questions-and-comments"></a>Perguntas e comentários

Gostaríamos de saber sua opinião sobre este exemplo. Você pode nos enviar comentários na seção *Problemas* deste repositório.

As perguntas sobre o desenvolvimento do Microsoft Office 365 em geral devem ser postadas no [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). Se sua pergunta estiver relacionada às APIs JavaScript para Office, não deixe de marcá-la com as tags [office-js] e [API].

## <a name="additional-resources"></a>Recursos adicionais

* 
  [Documentação dos suplementos do Office](https://msdn.microsoft.com/pt-br/library/office/jj220060.aspx)
* [Centro de Desenvolvimento do Office](http://dev.office.com/)
* Confira outros exemplos de Suplemento do Office em [OfficeDev no Github](https://github.com/officedev)

## <a name="copyright"></a>Direitos autorais
Copyright (c) 2017 Microsoft Corporation. Todos os direitos reservados.

