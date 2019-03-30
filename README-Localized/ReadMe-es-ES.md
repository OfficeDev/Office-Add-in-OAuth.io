# <a name="archived-office-add-in-that-uses-the-oauthio-service-to-get-authorization-to-external-services"></a>[ARCHIVADO] Complemento de Office que usa el servicio OAuth.io para obtener autorización a los servicios externos

> **Nota**: Este repositorio está archivado y ya no se mantiene de manera activa. Las vulnerabilidades de seguridad pueden existir en el proyecto o en sus dependencias. Si tiene previsto volver a usar o a ejecutar algún código de este repositorio, asegúrese primero de realizar las comprobaciones de seguridad adecuadas en el código o las dependencias. No use este proyecto como punto de partida de un complemento de producción de Office. Inicie siempre el código de producción con la carga de trabajo de desarrollo de Office y SharePoint en Visual Studio, o el [generador de Yeoman para complementos de Office](https://github.com/OfficeDev/generator-office), y siga los procedimientos recomendados de seguridad al elaborar el complemento.

El servicio OAuth.io simplifica el proceso de obtención de tokens de acceso OAuth 2.0 de servicios en línea conocidos como Facebook y Google. En este ejemplo se muestra cómo usarlo en un complemento de Office. 

## <a name="table-of-contents"></a>Tabla de contenido
* [Historial de cambios](#change-history)
* [Requisitos previos](#prerequisites)
* [Configurar el proyecto](#configure-the-project)
* [Descargar la biblioteca JavaScript OAuth.io](#download-the-oauth-io-javascript-library)
* [Crear un proyecto de Google](#create-a-google-project)
* [Crear una aplicación de Facebook](#create-a-facebook-app)
* [Crear una aplicación OAuth.io y configurarla para usar Google y Facebook](#create-an-OAuth-io-app-and-configure-it-to-use-google-and-facebook)
* [Agregar la clave pública al código de ejemplo](#add-the-public-key-to-the-sample-code)
* [Implementar el complemento](#deploy-the-add-in)
* [Ejecutar el proyecto](#run-the-project)
* [Iniciar el complemento](#start-the-add-in)
* [Probar el complemento](#test-the-add-in)
* [Preguntas y comentarios](#questions-and-comments)
* [Recursos adicionales](#additional-resources)

## <a name="change-history"></a>Historial de cambios

5 de enero de 2017:

* Versión inicial.

## <a name="prerequisites"></a>Requisitos previos

* Una cuenta con [OAuth.io](https://oauth.io/home)
* Cuentas de desarrollador con [Google](https://developers.google.com/) y [Facebook](https://developers.facebook.com/).
* Word 2016 para Windows, compilación 16.0.6727.1000 o posterior.
* [Nodo y npm](https://nodejs.org/en/) El proyecto está configurado para usar npm como un administrador de paquetes y un ejecutor de tareas. También está configurado para usar Lite Server como el servidor web que hospedará el complemento durante el desarrollo, de manera que tenga el complemento en funcionamiento rápidamente. Puede usar otro ejecutor de tareas o servidor web.
* [Git Bash](https://git-scm.com/downloads) (U otro cliente de Git).

## <a name="configure-the-project"></a>Configurar el proyecto

En la carpeta donde quiera colocar el proyecto, ejecute los siguientes comandos en el shell de Git Bash:

1. ```git clone {URL of this repo}``` para clonar este repositorio en la máquina local.
2. ```npm install``` para instalar todas las dependencias detalladas en el archivo package.json.
3. ```bash gen-cert.sh``` para crear el certificado necesario para ejecutar este ejemplo. 

Establezca el certificado para que sea una entidad de certificación raíz de confianza. En una máquina Windows, siga estos pasos:

1. En la carpeta del repositorio del equipo local, haga doble clic en ca.crt y seleccione **Instalar certificado**. 
2. Seleccione **Máquina local** y **Siguiente** para continuar. 
3. Seleccione **Colocar todos los certificados en el siguiente almacén** y **Examinar**.
4. Seleccione **Entidades de certificación raíz de confianza** y **Aceptar**. 
5. Seleccione **Siguiente** y, después, **Finalizar**. 

## <a name="download-the-oauthio-javascript-library"></a>Descargar la biblioteca JavaScript OAuth.io

2. En el proyecto, cree una subcarpeta denominada OAuth.io en la carpeta Scripts.
3. Descargar la biblioteca JavaScript OAuth.io de [oauth-js](https://github.com/oauth-io/oauth-js).
2. De la carpeta \dist de la biblioteca, copie oauth.js u oauth.min.js en la carpeta \Scripts\OAuth.io del proyecto. 
3. Si ha elegido oauth.min.js en el paso anterior, abra el archivo \Scripts\popupRedirect.js y cambie la línea `url: 'Scripts/OAuth.io/oauth.js',` por `url: 'Scripts/OAuth.io/oauth.min.js',`.

## <a name="create-a-google-project"></a>Crear un proyecto de Google

1. En la [consola del desarrollador de Google](https://console.developers.google.com/apis/dashboard), cree un proyecto denominado **Complemento de Office**.
2. Habilite la API de Google Plus para el proyecto. Vea la ayuda de Google para obtener más información.
3. Cree credenciales de aplicación para el proyecto y elija el **Id. de cliente OAuth** como el tipo de credencial, y **Aplicación web** como el tipo de aplicación. Use "Office-Add-in-OAuth.io" como el **Nombre** del cliente (y el nombre del producto, si se le ha pedido que especifique uno).
4. Establezca los **URI de redireccionamiento autorizados** en ```https://oauth.io/auth```.
5. Realice una copia del secreto de cliente y del Id. de cliente que se ha asignado.

## <a name="create-a-facebook-app"></a>Crear una aplicación de Facebook

1. En el [panel del desarrollador de Facebook](https://developers.facebook.com/), cree una aplicación denominada **Complemento de Office**. Vea la ayuda de Facebook para obtener más información.
3. Realice una copia del secreto de cliente y del Id. de cliente que se ha asignado. 
4. Especifique ```oauth.io``` en el cuadro **Dominios de aplicación** y especifique ```https://oauth.io/``` como la **URL del sitio**. 

## <a name="create-an-oauthio-app-and-configure-it-to-use-google-and-facebook"></a>Crear una aplicación OAuth.io y configurarla para usar Google y Facebook

1. En el panel de OAuth.io, cree una aplicación y anote la **clave pública** que OAuth.io asigna.
3. Si todavía no está ahí, agregue `localhost` a las direcciones URL de la lista de permitidos de la aplicación.
2. Agregue **Google Plus** como un proveedor (también denominado una "API integrada") de la aplicación. Complete el secreto y el Id. de cliente que ha obtenido al crear el proyecto de Google y, después, especifique `profile` como el **ámbito**. Si la configuración del proveedor (API integrada) no se guarda en IE, intente volver a abrir el panel de OAuth.io en otro explorador.
3. Repita el paso anterior para **Facebook**, pero especifique `user_about_me` como el **ámbito**.

## <a name="add-the-public-key-to-the-sample-code"></a>Agregar la clave pública al código de ejemplo

Se llama a la API `OAuth.initialize` en el archivo popup.js. Busque la línea y pase la clave pública que ha obtenido en la sección anterior como una cadena en esta función.

## <a name="deploy-the-add-in"></a>Implementar el complemento

Ahora debe indicarle a Microsoft Word dónde encontrar el complemento.

1. Cree un recurso compartido de red o [comparta una carpeta en la red](https://technet.microsoft.com/es-es/library/cc770880.aspx).
2. Coloque una copia del archivo de manifiesto Office-Add-in-OAuth.io.xml, desde la raíz del proyecto, en la carpeta compartida.
3. Inicie Word y abra un documento.
4. Seleccione la pestaña **Archivo** y haga clic en **Opciones**.
5. Haga clic en **Centro de confianza** y seleccione el botón **Configuración del Centro de confianza**.
6. Seleccione **Catálogos de complementos de confianza**.
7. En el campo **Dirección URL del catálogo**, escriba la ruta de red al recurso compartido de carpeta que contiene Office-Add-in-OAuth.io.xml y, después, pulse **Agregar catálogo**.
8. Active la casilla **Mostrar en menú** y, después, pulse **Aceptar**.
9. Aparecerá un mensaje para informarle que la configuración se aplicará la próxima vez que inicie Microsoft Office. Cierre Word.

## <a name="run-the-project"></a>Ejecutar el proyecto

1. Abra una ventana Comandos de node en la carpeta del proyecto y ejecute ```npm start``` para iniciar el servicio web. Deje abierta la ventana Comandos.
2. Abra Internet Explorer o Microsoft Edge y escriba ```https://localhost:3000``` en el cuadro de dirección. Si no recibe ninguna advertencia sobre el certificado, cierre el explorador y siga con la sección siguiente titulada **Iniciar el complemento**. Si recibe una advertencia que indica que el certificado no es de confianza, siga estos pasos:
3. El explorador le proporciona un vínculo para abrir la página a pesar de la advertencia. Ábralo.
4. Después de que se abra la página, aparecerá un error de certificado rojo en la barra de direcciones. Haga doble clic en el error.
5. Seleccione **Ver certificado**.
5. Seleccione **Instalar certificado**.
4. Seleccione **Máquina local** y **Siguiente** para continuar. 
3. Seleccione **Colocar todos los certificados en el siguiente almacén** y **Examinar**.
4. Seleccione **Entidades de certificación raíz de confianza** y **Aceptar**. 
5. Seleccione **Siguiente** y, después, **Finalizar**.
6. Cierre el explorador.

## <a name="start-the-add-in"></a>Iniciar el complemento

1. Reinicie Word y abra un documento de Word.
2. En la pestaña **Insertar** de Word 2016, elija **Mis complementos**.
3. Seleccione la pestaña **CARPETA COMPARTIDA**.
4. Pulse **Autenticación con OAuthIO** y, después, seleccione **Aceptar**.
5. Si su versión de Word admite los comandos de complemento, la interfaz de usuario le informará de que se ha cargado el complemento.
6. En la cinta de opciones de Inicio es un nuevo grupo denominado **OAuth.io** con un botón llamado **Mostrar** y un icono. Haga clic en ese botón para abrir el complemento.

 > Nota: El complemento se cargará en un panel de tareas si los comandos del complemento no son compatibles con su versión de Word.

## <a name="test-the-add-in"></a>Probar el complemento

1. Haga clic en el botón **Obtener nombre de Facebook**.
2. Se abrirá un elemento emergente y se le pedirá que inicie sesión con Facebook (a no ser que ya lo haya hecho).
3. Después de iniciar sesión, el nombre de usuario de Facebook se inserta en el documento de Word.
4. Repita los pasos anteriores con el botón **Obtener nombre de Google**.

## <a name="questions-and-comments"></a>Preguntas y comentarios

Nos encantaría recibir sus comentarios sobre este ejemplo. Puede enviarnos comentarios a través de la sección *Problemas* de este repositorio.

Las preguntas generales sobre el desarrollo de Microsoft Office 365 deben publicarse en [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). Si su pregunta trata sobre las API de JavaScript para Office, asegúrese de que se etiqueta con [office-js] y [API].

## <a name="additional-resources"></a>Recursos adicionales

* 
  [Documentación de complementos de Office](https://msdn.microsoft.com/es-es/library/office/jj220060.aspx)
* [Centro de desarrollo de Office](http://dev.office.com/)
* Más ejemplos de complementos de Office en [OfficeDev en GitHub](https://github.com/officedev)

## <a name="copyright"></a>Derechos de autor
Copyright (c) 2017 Microsoft Corporation. Todos los derechos reservados.

