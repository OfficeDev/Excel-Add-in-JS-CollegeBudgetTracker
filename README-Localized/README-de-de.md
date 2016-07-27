# Aufgabenbereich-Add-In-Beispiel „Studien-Haushaltsplanverfolgung” für Excel 2016

_Gilt für: Excel 2016_

Dieses Aufgabenbereich-Add-In veranschaulicht, wie mithilfe der JavaScript-APIs in Excel 2016 eine Studien-Haushaltsplanverfolgung erstellt werden kann. Es ist in zwei Versionen verfügbar: Code-Editor und Visual Studio.

![Studien-Haushaltsplanverfolgungsbeispiel](../images/CollegeBudgetTracker_tracker.PNG)

## Probieren Sie es aus
### Code-Editor-Version

Am einfachsten können Sie Ihr Add-In bereitstellen und testen, indem Sie das Manifest in eine Netzwerkfreigabe kopieren.

1.  Erstellen Sie einen Ordner in einer Netzwerkfreigabe (zum Beispiel „ \\\MyShare\CollegeBudgetTracker”).  
2.  Kopieren Sie das Manifest (CollegeBudgetTrackerManifest.xml) in eine Netzwerkfreigabe (z.ä B. \\\MyShare\MyManifests).
3.  Fügen Sie den Freigabepfad, unter dem das Manifest enthalten ist, als vertrauenswürdigen App-Katalog in Excel hinzu.

    a. Starten Sie Excel, und öffnen Sie ein leeres Arbeitsblatt.  
    
    b. Klicken Sie auf die Registerkarte **Datei**, und klicken Sie dann auf **Optionen**.
    
    c. Wählen Sie **Trust Center** aus, und klicken Sie dann auf die Schaltfläche **Einstellungen für das Trust Center**.
    
  d. Klicken Sie auf **Vertrauenswürdige Add-in-Kataloge**.
    
  e. Geben Sie im Feld **Katalog-URL** den Pfad zu der in Schritt 3 erstellten Netzwerkfreigabe ein, und klicken Sie auf **Katalog hinzufügen**.
    
   f. Aktivieren Sie das Kontrollkästchen **Im Menü anzeigen**, und wählen Sie dann **OK**. Eine Meldung wird angezeigt, dass Ihre Einstellungen angewendet werden, wenn Office das nächste Mal gestartet wird. 
        
4.  Testen und führen Sie das Add-In aus. 

  a. Klicken Sie auf der Registerkarte **Einfügen** in Excel 2016 auf **Meine-Add-Ins**. 
    
  b. Wählen Sie im Dialogfenster **Office-Add-Ins** die Option **Freigegebener Ordner** aus.
    
  c. Klicken Sie auf der Registerkarte „Start“ auf den Befehl **Studien-Haushaltsplanverfolgung**. The add-in opens in a task pane and creates the college budget tracker in the active sheet as shown in this diagram. 
      
   ![Studien-Haushaltsplanverfolgungsbeispiel](../images/CollegeBudgetTracker_tracker.PNG) 

  d. Fügen Sie einige Ausgaben und Einnahmen mithilfe der Registerkarten **Ausgaben hinzufügen** und **Einnahmen hinzufügen** hinzu, und sehen Sie, wie sich die Daten und die Diagramme dynamisch ändern.
    
      ![College Budget Tracker Sample](../images/CollegeBudgetTracker_taskpane1.PNG) 

Um das Manifest in Ihrem Add-In zu verwenden, müssen Sie das <SourceLocation>-Element der Manifestdatei bearbeiten, damit es auf den Freigabepfad für die Datei „Home.html” zeigt.
    
### Visual Studio-Version
1.  Kopieren Sie das Projekt in einen lokalen Ordner, und öffnen Sie die Datei „Excel-Add-in-JS-CollegeBudgetTracker.sln” in Visual Studio.
2.  Drücken Sie F5, um das Beispiel-Add-In zu erstellen und bereitzustellen. Excel wird gestartet und das Add-In wird in einem Aufgabenbereich rechts neben einem leeren Arbeitsblatt geöffnet, wie in der folgenden Abbildung dargestellt. 
        
  ![Studien-Haushaltsplanverfolgungsbeispiel](../images/CollegeBudgetTracker_tracker.PNG) 

3.  Fügen Sie einige Ausgaben und Einnahmen mithilfe der Registerkarten **Ausgaben hinzufügen** und **Einnahmen hinzufügen** hinzu, und sehen Sie, wie sich die Daten und die Diagramme dynamisch ändern.

  ![Studien-Haushaltsplanverfolgungsbeispiel](../images/CollegeBudgetTracker_taskpane1.PNG) 


### Weitere Informationen

Die Excel-JavaScript-APIs haben viel mehr bei der Entwicklung von Add-Ins zu bieten. Im Folgenden werden nur einige der verfügbaren Ressourcen aufgeführt. 

1.  [Programmierungsübersicht für Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Codeausschnitt-Explorer für Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Codebeispiele zu Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md) 
4.  [JavaScript-API-Referenz zu Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [Erstellen Ihres ersten Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)
