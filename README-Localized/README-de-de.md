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
# <a name="word-add-in-javascript-speckit"></a>Word Add-In JavaScript SpecKit

Erfahren Sie, wie Sie ein Add-In erstellen, das Textbausteine erfasst und einfügt und wie Sie einen einfachen Dokumentüberprüfungsprozess implementieren können.

## <a name="table-of-contents"></a>Inhalt
* [Änderungsverlauf](#change-history)
* [Voraussetzungen](#prerequisites)
* [Konfigurieren des Projekts](#configure-the-project)
* [Ausführen des Projekts](#run-the-project)
* [Grundlegendes zum Code](#understand-the-code)
* [Fragen und Kommentare](#questions-and-comments)
* [Zusätzliche Ressourcen](#additional-resources)

## <a name="change-history"></a>Änderungsverlauf

31. März 2016
* Erste Beispielversion

## <a name="prerequisites"></a>Anforderungen

* Word 2016 für Windows, Build 16.0.6727.1000 oder höher.
* [Node und npm](https://nodejs.org/en/)
* [Git Bash](https://git-scm.com/downloads) Sie sollten eine höhere Buildversion verwenden, da bei früheren Buildversionen ein Fehler beim Generieren der Zertifikate auftreten kann.

## <a name="configure-the-project"></a>Konfigurieren des Projekts

Führen Sie folgende Befehle in der Bash-Shell im Stammverzeichnis dieses Projekts aus:

1. Klonen Sie dieses Repository auf ihrem lokalen Computer.
2. ```npm install``` zum Installieren aller Abhängigkeiten in package.json.
3. ```bash gen-cert.sh``` zum Erstellen der für die Ausführung dieses Beispiels erforderlichen Zertifikate. Doppelklicken Sie dann im Repository auf dem lokalen Computer auf „ca.crt“, und wählen Sie **Zertifikat installieren**. Wählen Sie **Lokaler Computer**, und wählen Sie **Weiter**, um den Vorgang fortzusetzen. Wählen Sie die Option **Alle Zertifikate in folgendem Speicher speichern**, und wählen Sie dann **Durchsuchen**.  Wählen Sie **Vertrauenswürdige Stammzertifizierungsstellen**, und wählen Sie dann **OK**. Wählen Sie **Weiter** und dann **Fertig stellen**. Ihre Zertifizierungsstelle wurde nun zum Zertifikatspeicher hinzugefügt.
4. ```npm start``` zum Starten des Diensts.

Sie haben nun dieses Beispiel-Add-In bereitgestellt. Jetzt müssen Sie Microsoft Word mitteilen, wo es das Add-In finden kann.

1. Erstellen Sie eine Netzwerkfreigabe oder [geben Sie einen Ordner im Netzwerk frei](https://technet.microsoft.com/en-us/library/cc770880.aspx), und platzieren Sie die [word-add-in-javascript-speckit-manifest.xml](word-add-in-javascript-speckit-manifest.xml)-Manifestdatei darin.
3. Starten Sie Word, und öffnen Sie ein Dokument.
4. Klicken Sie auf die Registerkarte **Datei**, und klicken Sie dann auf **Optionen**.
5. Wählen Sie **Sicherheitscenter** aus, und klicken Sie dann auf die Schaltfläche **Einstellungen für das Sicherheitscenter**.
6. Klicken Sie auf **Kataloge vertrauenswürdiger Add-Ins**.
7. Geben Sie in das Feld **Katalog-URL** den Netzwerkpfad zur Ordnerfreigabe an, die die Datei „word-add-in-javascript-speckit-manifest.xml“ enthält, und wählen Sie dann **Katalog hinzufügen**.
8. Aktivieren Sie das Kontrollkästchen **Im Menü anzeigen**, und klicken Sie dann auf **OK**.
9. Es wird eine Meldung angezeigt, dass Ihre Einstellungen angewendet werden, wenn Sie Microsoft Office das nächste Mal starten.

## <a name="run-the-project"></a>Ausführen des Projekts

1. Öffnen Sie ein Word-Dokument.
2. Klicken Sie auf der Registerkarte **Einfügen** in Word 2016 auf **Meine-Add-Ins**.
3. Klicken Sie auf die Registerkarte **FREIGEGEBENER ORDNER**.
4. Wählen Sie **Word SpecKit-Add-In**, und wählen Sie dann **OK**.
5. Wenn Add-In-Befehle von Ihrer Word-Version unterstützt werden, werden Sie in der Benutzeroberfläche darüber informiert, dass das Add-In geladen wurde.

### <a name="ribbon-ui"></a>Menüband-Benutzeroberfläche
Im Menüband können Sie folgende Aktionen ausführen:
* Wählen Sie die Registerkarte **SpecKit-Add-In**, um das Add-In in der Benutzeroberfläche zu starten.
* Wählen Sie **Spezifikationsvorlage einfügen**, um den Aufgabenbereich zu starten und eine Spezifikationsvorlage in das Dokument einzufügen.
* Verwenden Sie die Schaltflächen für Validierung im Menüband, oder klicken Sie mit der rechten Maustaste auf das Kontextmenü, um das Dokument anhand einer Blacklist zu validieren.

 > Hinweis: Das Add-In wird in einem Aufgabenbereich geladen, wenn Add-In-Befehle von Ihrer Version von Word nicht unterstützt werden.

### <a name="task-pane-ui"></a>Aufgabenbereich-Benutzeroberfläche
Im Aufgabenbereich können Sie folgende Aktionen ausführen:
* Speichern Sie einen Satz, indem Sie den Cursor in einem Satz platzieren, geben Sie einen Namen in das Feld über **Satz zum Textbaustein hinzufügen* im Aufgabenbereich ein, und wählen Sie **Satz zum Baustein hinzufügen**. Dieselbe Aktion können Sie für Absätze vornehmen.
* Beim Speichern von Sätzen und Absätzen werden auch die Textbausteine in der Dropdownliste ** Textbausteinen einfügen** angezeigt.
* Platzieren Sie den Cursor im Dokument. Wählen Sie einen Textbaustein aus der Dropdownliste aus, und der Textbaustein wird in das Dokument eingefügt.
* Ändern Sie die *Author*-Eigenschaft des Dokuments, indem Sie den Namen des Autors ändern und auf die Schaltfläche **Autorname aktualisieren** klicken. Dadurch werden die Dokumenteigenschaft und der Inhalt eines verbundenen Inhaltssteuerelements aktualisiert.

## <a name="understand-the-code"></a>Grundlegendes zum Code

In diesem Beispiel wird die Version 1.2 des [Anforderungssatzes](http://dev.office.com/reference/add-ins/office-add-in-requirement-sets?product=word) in der Vorschauversion verwendet. Es wird jedoch Version 1.3 benötigt, sobald der Anforderungssatz allgemein verfügbar ist.

### <a name="task-pane"></a>Aufgabenbereich

Die Aufgabenbereichsfunktionen sind in der Datei „sample.js“ eingerichtet, welche die folgenden Funktionen enthält:

* Einrichten der Benutzeroberfläche und der Ereignishandler
* Abrufen der Spezifikationsvorlage von einem Dienst und Einfügen dieser Vorlage in das Dokument
* Laden einer Blacklist mit Wörtern, die für die Validierung des Dokuments verwendet werden. Diese Wörter werden als unzulässige Wörter für die Zwecke dieses Beispiels betrachtet.
* Laden eines Standardtextbausteins aus einem Dienst und Zwischenspeichern von diesem im lokalen Speicher
* Basiscode zum Testen des Funktionsdateicodes. Sie werden ggf. einen Add-In-Befehlscode im Aufgabenbereich entwickeln, bevor Sie diesen in eine Funktionsdatei verschieben, da Sie an die Funktionsdatei keinen Debugger anfügen können.
* Laden des standardmäßigen Autorennamens aus den Dokumenteigenschaften im Aufgabenbereich. Hier wird gezeigt, wie Sie auf eine benutzerdefinierte XML-Komponente in einem Dokument zugreifen und diese ändern.
* Veröffentlichen des Textbausteins im Dienst

### <a name="document-validation-and-the-dialog-api"></a>Dokumentvalidierung und die Dialog-API

Die Datei „validation.js“ enthält den Code zum Validieren des Dokuments anhand einer Blacklist. Die validateContentAgainstBlacklist()-Methode verwendet die neue splitTextRanges-Methode, um das Dokument auf Grundlage von Trennzeichen zu teilen. Die Trennzeichen in dieser Funktion erkennen Wörter im Dokument. Es wird die Schnittmenge der Wörter im Dokument und der Blacklist identifiziert. Diese Ergebnisse werden an den lokalen Speicher übergeben. Anschließend wird die displayDialogAsync-Methode verwendet, um ein Dialogfeld (dialog.html) zu öffnen. Das Dialogfeld ruft die Validierungsergebnisse aus dem lokalen Speicher ab und zeigt die Ergebnisse an.

### <a name="boilerplate-text-functionality"></a>Textbausteinfunktionen

Die Datei „boilerplate.js“ enthält Beispiele dazu, wie Sie Textbausteine im lokalen Speicher speichern, die Fabric-Dropdownliste basierend auf den gespeicherten Textbausteinen aktualisieren, und die in der Dropdownliste ausgewählten Textbausteine einfügen. Diese Datei enthält Beispiele für:
* splitTextRanges (neu für den Anforderungssatz WordApi 1.3). Diese API in zukünftigen Versionen durch split() ersetzt.
* compareLocationWith (neu für den Anforderungssatz WordApi 1.3)
* Aktualisierung der Fabric-Dropdownliste mit den neuen Einträgen
* Einfügen von Textbausteinen in das Dokument

### <a name="custom-xml-binding-to-core-document-properties"></a>Benutzerdefinierte XML-Bindung an die Haupteigenschaften des Dokuments

Die Datei „authorCustomXml.js“ enthält Methoden zum Abrufen und Festlegen der standardmäßigen Dokumenteigenschaften.

* Laden Sie die author-Eigenschaft im Aufgabenbereich, wenn der Aufgabenbereich geladen wird. Beachten Sie, dass das Dokument auch den Wert der author-Eigenschaft enthält. Dies kommt daher, dass die Vorlage ein Inhaltssteuerelement enthält, das an diese Dokumenteigenschaft gebunden ist. Dadurch können Sie die Standardwerte im Dokument basierend auf den Inhalten einer benutzerdefinierten XML-Komponente festlegen.
* Aktualisieren Sie die author-Eigenschaft im Aufgabenbereich. Dadurch werden die Dokumenteigenschaft und der Inhalt des verbundenen Inhaltssteuerelements im Dokument aktualisiert.

### <a name="add-in-commands"></a>Add-In-Befehle

Die Add-In-Befehldeklarationen sind in der Datei „word-add-in-javascript-speckit-manifest.xml“ enthalten. In diesem Beispiel wird gezeigt, wie Sie Add-In-Befehle im Menüband und im Kontextmenü erstellen.

## <a name="questions-and-comments"></a>Fragen und Kommentare

Wir schätzen Ihr Feedback hinsichtlich des Word SpecKit-Beispiels. Sie können uns Ihr Feedback über den Abschnitt *Probleme* dieses Repositorys senden.

Allgemeine Fragen zur Microsoft Office 365-Entwicklung sollten in [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API) gestellt werden. Stellen Sie sicher, dass Ihre Fragen mit [office-js] und [API] markiert sind.

## <a name="additional-resources"></a>Zusätzliche Ressourcen

* [Dokumentation zu Office-Add-Ins](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Office Dev Center](http://dev.office.com/)
* [Office 365 APIs – Startprojekte und Codebeispiele](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)

## <a name="copyright"></a>Copyright
Copyright (c) 2016 Microsoft Corporation. Alle Rechte vorbehalten.



In diesem Projekt wurden die [Microsoft Open Source-Verhaltensregeln](https://opensource.microsoft.com/codeofconduct/) übernommen. Weitere Informationen finden Sie unter [Häufig gestellte Fragen zu Verhaltensregeln](https://opensource.microsoft.com/codeofconduct/faq/), oder richten Sie Ihre Fragen oder Kommentare an [opencode@microsoft.com](mailto:opencode@microsoft.com).
