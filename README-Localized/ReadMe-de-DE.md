# <a name="archived-office-add-in-that-uses-the-oauthio-service-to-get-authorization-to-external-services"></a>[ARCHIVERT] Office-Add-In, das den OAuth.io-Dienst zur Autorisierung bei externen Diensten nutzt

> **Hinweis:** Dieses Repository wurde archiviert und wird nicht mehr aktualisiert. Das Projekt oder seine Abhängigkeiten enthalten möglicherweise Sicherheitsrisiken. Wenn Sie beabsichtigen, dieses Repository wiederzuverwenden oder darin enthaltenen Code auszuführen, führen Sie bitte zuerst angemessene Sicherheitsprüfungen des Codes oder der Abhängigkeiten durch. Verwenden Sie dieses Projekt nicht als Ausgangspunkt zum Erstellen eines Office-Add-Ins für die Produktionsumgebung. Beginnen Sie die Erstellung von Produktionscode immer mit dem Entwicklungsworkload für Office/SharePoint in Visual Studio oder mit dem [Yeoman-Generator für Office-Add-Ins](https://github.com/OfficeDev/generator-office), und befolgen Sie bei der Entwicklung des Add-Ins die bewährten Methoden für die Sicherheit.

Der Dienst OAuth.io vereinfacht das Abrufen von OAuth 2.0-Zugriffstokens von beliebten Onlinediensten wie Facebook und Google. In diesem Beispiel sehen Sie, wie Sie den Dienst in einem Office-Add-In verwenden können. 

## <a name="table-of-contents"></a>Inhalt
* [Änderungsverlauf](#change-history)
* [Voraussetzungen](#prerequisites)
* [Konfigurieren des Projekts](#configure-the-project)
* [Herunterladen der OAuth.io-JavaScript-Bibliothek](#download-the-oauth-io-javascript-library)
* [Erstellen eines Google-Projekts](#create-a-google-project)
* [Erstellen einer Facebook-App](#create-a-facebook-app)
* [Erstellen einer OAuth.io-App samt Konfiguration für Google und Facebook](#create-an-OAuth-io-app-and-configure-it-to-use-google-and-facebook)
* [Hinzufügen des öffentlichen Schlüssels zum Beispielcode](#add-the-public-key-to-the-sample-code)
* [Bereitstellen des Add-Ins](#deploy-the-add-in)
* [Ausführen des Projekts](#run-the-project)
* [Starten des Add-Ins](#start-the-add-in)
* [Testen des Add-Ins](#test-the-add-in)
* [Fragen und Kommentare](#questions-and-comments)
* [Zusätzliche Ressourcen](#additional-resources)

## <a name="change-history"></a>Änderungsverlauf

5. Januar 2017:

* Ursprüngliche Version

## <a name="prerequisites"></a>Voraussetzungen

* Ein Konto bei [OAuth.io](https://oauth.io/home)
* Entwicklerkonten bei [Google](https://developers.google.com/) und [Facebook](https://developers.facebook.com/)
* Word 2016 für Windows, Build 16.0.6727.1000 oder höher.
* [Node und npm](https://nodejs.org/en/) Das Projekt ist so konfiguriert, dass npm als Paket-Manager und für die Taskausführung verwendet wird. Zudem wird Lite Server als der Webserver verwendet, der das Add-In während der Entwicklung hostet, sodass das Add-In schnell betriebsbereit ist. Sie können jedoch auch eine andere Taskausführung oder einen anderen Webserver verwenden.
* [Git Bash](https://git-scm.com/downloads) (Oder ein anderer Git-Client.)

## <a name="configure-the-project"></a>Konfigurieren des Projekts

Führen Sie in dem Ordner, in dem das Projekt erstellt werden soll, die folgenden Befehle in der Git Bash-Shell aus:

1. ```git clone {URL of this repo}``` zum Klonen dieses Repositorys auf ihrem lokalen Computer.
2. ```npm install``` zum Installieren aller Abhängigkeiten in der Datei „package.json“.
3. ```bash gen-cert.sh``` zum Erstellen des für die Ausführung dieses Beispiels erforderlichen Zertifikats. 

Legen Sie das Zertifikat als vertrauenswürdige Stammzertifizierungsstelle fest. Dies sind die Schritte auf einem Windows-Computer:

1. Doppelklicken Sie im Repository-Ordner auf dem lokalen Computer auf „ca.crt“, und wählen Sie **Zertifikat installieren** aus. 
2. Wählen Sie **Lokaler Computer** aus, und wählen Sie **Weiter**, um den Vorgang fortzusetzen. 
3. Wählen Sie die Option **Alle Zertifikate in folgendem Speicher speichern**, und wählen Sie dann **Durchsuchen**.
4. Wählen Sie **Vertrauenswürdige Stammzertifizierungsstellen** und dann **OK**. 
5. Wählen Sie **Weiter** und dann **Fertig stellen**. 

## <a name="download-the-oauthio-javascript-library"></a>Herunterladen der OAuth.io-JavaScript-Bibliothek

2. Erstellen Sie im Projekt im Ordner „Skripts“ einen Unterordner mit dem Namen „OAuth.io“.
3. Laden Sie die OAuth.io-JavaScript-Bibliothek unter [oauth-js](https://github.com/oauth-io/oauth-js) herunter.
2. Kopieren Sie aus dem Ordner „\dist“ der Bibliothek die Datei „oauth.js“ oder die Datei „oauth.min.js“ in den Projektordner „\Skripts\OAuth.io“. 
3. Wenn Sie im vorherigen Schritt „oauth.min.js“ gewählt haben: Öffnen Sie die Datei „\Skripts\popupRedirect.js“, und ändern Sie die Zeile `url: 'Scripts/OAuth.io/oauth.js',` in `url: 'Scripts/OAuth.io/oauth.min.js',`.

## <a name="create-a-google-project"></a>Erstellen eines Google-Projekts

1. Erstellen Sie in Ihrer [Google-Entwicklerkonsole](https://console.developers.google.com/apis/dashboard) ein Projekt mit dem Namen **Office-Add-In**.
2. Aktivieren Sie die Google Plus-API für das Projekt. Details hierzu finden Sie in der Google-Hilfe.
3. Erstellen Sie App-Anmeldeinformationen für das Projekt. Wählen Sie die **OAuth-Client-ID** als Anmeldeinformationstyp und **Web-Anwendung** als Anwendungstyp aus. Geben Sie unter **Name** für den Client „Office-Add-in-Oauth.io“ ein (sowie den Produktnamen, falls Sie dazu aufgefordert werden).
4. Tragen Sie unter **Authorized redirect URIs** die Adresse ```https://oauth.io/auth``` ein.
5. Kopieren Sie den geheimen Clientschlüssel und die Client-ID, die Ihnen zugewiesen werden.

## <a name="create-a-facebook-app"></a>Erstellen einer Facebook-App

1. Erstellen Sie in Ihrem [Facebook-Entwicklerdashboard](https://developers.facebook.com/) eine App mit dem Namen **Office-Add-In**. Details hierzu finden Sie in der Facebook-Hilfe.
3. Kopieren Sie den geheimen Clientschlüssel und die Client-ID, die Ihnen zugewiesen werden. 
4. Tragen Sie ```oauth.io``` ins Feld **App Domains** und ```https://oauth.io/``` ins Feld **Site URL** ein. 

## <a name="create-an-oauthio-app-and-configure-it-to-use-google-and-facebook"></a>Erstellen einer OAuth.io-App samt Konfiguration für Google und Facebook

1. Erstellen Sie in Ihrem OAuth.io-Dashboard eine App, und notieren Sie sich den **öffentlichen Schlüssel**, den OAuth.io der App zuweist.
3. Fügen Sie `localhost` zur URL-Whitelist der App hinzu, falls dort noch kein entsprechender Eintrag vorhanden ist.
2. Fügen Sie **Google Plus** als Anbieter („Provider“) zur App hinzu. Anbieter werden auch als „integrierte API“ bezeichnet („Integrated API“). Tragen Sie die Client-ID und den geheimen Schlüssel ein, die Sie bei der Erstellung des Google-Projekts erstellt haben. Legen Sie anschließend `profile` als **scope** fest. Falls sich die Konfiguration des Anbieters (der integrierten API) in IE nicht speichern lässt, öffnen Sie das OAuth.io-Dashboard in einem anderen Browser.
3. Wiederholen Sie den vorherigen Schritt für **Facebook**, aber legen Sie `user_about_me` als **scope** fest.

## <a name="add-the-public-key-to-the-sample-code"></a>Hinzufügen des öffentlichen Schlüssels zum Beispielcode

Die API `OAuth.initialize` wird in der Datei „popup.js“ aufgerufen. Suchen Sie die entsprechende Zeile, und übergeben Sie den öffentlichen Schlüssel aus dem vorherigen Abschnitt als Zeichenfolge an diese Funktion.

## <a name="deploy-the-add-in"></a>Bereitstellen des Add-Ins

Jetzt müssen Sie Microsoft Word mitteilen, wo es das Add-In finden kann.

1. Erstellen Sie eine Netzwerkfreigabe, oder [Geben Sie einen Ordner im Netzwerk frei](https://technet.microsoft.com/de-de/library/cc770880.aspx).
2. Kopieren Sie die Manifestdatei „Office-Add-in-OAuth.io.xml“ aus dem Projektstammverzeichnis in den freigegebenen Ordner.
3. Starten Sie Word, und öffnen Sie ein Dokument.
4. Klicken Sie auf die Registerkarte **Datei**, und klicken Sie dann auf **Optionen**.
5. Wählen Sie **Sicherheitscenter** aus, und klicken Sie dann auf die Schaltfläche **Einstellungen für das Sicherheitscenter**.
6. Wählen Sie **Vertrauenswürdige Add-In-Kataloge** aus.
7. Geben Sie in das Feld **Katalog-URL** den Netzwerkpfad zur Ordnerfreigabe ein, unter der die Datei „Office-Add-in-OAuth.io.xml“ liegt, und klicken Sie auf **Katalog hinzufügen**.
8. Aktivieren Sie das Kontrollkästchen **Im Menü anzeigen**, und klicken Sie dann auf **OK**.
9. Eine Meldung wird angezeigt, dass Ihre Einstellungen angewendet werden, wenn Microsoft Office das nächste Mal gestartet wird. Schließen Sie Word.

## <a name="run-the-project"></a>Ausführen des Projekts

1. Öffnen Sie ein node-Befehlsfenster im Ordner des Projekts, und führen Sie ```npm start``` aus, um den Webdienst zu starten. Lassen Sie das Befehlsfenster geöffnet.
2. Öffnen Sie Internet Explorer oder Edge, und geben Sie in das Adressfeld ```https://localhost:3000``` ein. Wenn Sie keine Warnungen über das Zertifikat erhalten, schließen Sie den Browser, und fahren Sie weiter unten mit dem Abschnitt **Starten des Add-Ins** fort. Wenn Sie eine Warnung erhalten, dass das Zertifikat nicht vertrauenswürdig ist, fahren Sie mit den folgenden Schritten fort:
3. Der Browser bietet einen Link, um die Seite trotz der angezeigten Warnung zu öffnen. Öffnen Sie sie.
4. Nach dem Öffnen der Seite wird ein roter Zertifikatfehler in der Adressleiste angezeigt. Doppelklicken Sie auf den Fehler.
5. Wählen Sie **Zertifikat anzeigen**.
5. Wählen Sie **Zertifikat installieren**.
4. Wählen Sie **Lokaler Computer** aus, und wählen Sie **Weiter**, um den Vorgang fortzusetzen. 
3. Wählen Sie die Option **Alle Zertifikate in folgendem Speicher speichern**, und wählen Sie dann **Durchsuchen**.
4. Wählen Sie **Vertrauenswürdige Stammzertifizierungsstellen** und dann **OK**. 
5. Wählen Sie **Weiter** und dann **Fertig stellen**.
6. Schließen Sie den Browser.

## <a name="start-the-add-in"></a>Starten des Add-Ins

1. Starten Sie Word neu, und öffnen Sie ein Word-Dokument.
2. Klicken Sie auf der Registerkarte **Einfügen** in Word 2016 auf **Meine-Add-Ins**.
3. Klicken Sie auf die Registerkarte **FREIGEGEBENER ORDNER**.
4. Wählen Sie **Auth with OAuthIO** und anschließend **OK** aus.
5. Wenn Add-In-Befehle von Ihrer Word-Version unterstützt werden, werden Sie in der Benutzeroberfläche darüber informiert, dass das Add-In geladen wurde.
6. Im Menüband „Start“ finden Sie eine neue Gruppe namens **OAuth.io** mit der Schaltfläche **Show** und einem Symbol. Klicken Sie auf diese Schaltfläche, um das Add-In zu öffnen.

 > Hinweis: Das Add-In wird in einem Aufgabenbereich geladen, wenn Add-In-Befehle von Ihrer Version von Word nicht unterstützt werden.

## <a name="test-the-add-in"></a>Testen des Add-Ins

1. Klicken Sie auf die Schaltfläche **Get Facebook Name**.
2. Ein Popupfenster wird geöffnet, und Sie werden zur Anmeldung bei Facebook aufgefordert (sofern Sie nicht bereits angemeldet sind).
3. Nach der Anmeldung wird Ihr Facebook-Benutzername in das Word-Dokument eingefügt.
4. Wiederholen Sie die oben genannten Schritte mit der Schaltfläche **Get Google Name**.

## <a name="questions-and-comments"></a>Fragen und Kommentare

Wir schätzen Ihr Feedback hinsichtlich dieses Beispiels. Sie können uns Ihr Feedback über den Abschnitt *Probleme* dieses Repositorys senden.

Fragen zur Microsoft Office 365-Entwicklung sollten in [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API) gestellt werden. Wenn Ihre Frage die Office JavaScript-APIs betrifft, sollte die Frage mit [office-js] und [API] kategorisiert sein.

## <a name="additional-resources"></a>Zusätzliche Ressourcen

* 
  [Dokumentation zu Office-Add-Ins](https://msdn.microsoft.com/de-de/library/office/jj220060.aspx)
* [Office Dev Center](http://dev.office.com/)
* Weitere Office-Add-In-Beispiele unter [OfficeDev auf Github](https://github.com/officedev)

## <a name="copyright"></a>Copyright
Copyright (c) 2017 Microsoft Corporation. Alle Rechte vorbehalten.

