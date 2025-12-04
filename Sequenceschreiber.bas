Attribute VB_Name = "Main"
Option Explicit

Public i As Integer, j  As Integer
Public Const strMethodedaten As String = "L:\Makros\Sequenceschreiber\Daten für Sequenceschreiber.xlsx"

Private Sub Import()
    ' Strings
    Dim strGerät As String
    Dim strMethode As String
    Dim strTopic As String
    Dim strPfad As String
    Dim strDatum As String
    Dim strDatei As String
    Dim strVolleDatei As String, strOperatorTest As String, strName As String
    
    ' Workbooks und Worksheets
    Dim DatenWB As Workbook
    Dim ZWB As Workbook
    Dim QWB As Workbook
    Dim ZWS As Worksheet
    
    ' Ranges
    Dim rngZelle As Range
    
    ' Arrays
    Dim arrQuellKolonne As Variant
    Dim arrExporte() As Variant
    Dim arrEinzelEinwaagen As Variant
    
    ' Integers und Doubles
    Dim intMethodenZeile As Integer
    Dim intOperatorCount As Integer
    Dim j As Integer
    Dim intZeile As Integer
    Dim dblStdEinwaage As Double
    Dim dblSumme As Double

    
    ' Überprüfung, ob Methode ausgewählt wurde
    If wsData.Cells(2, 2) = "Methode" Then
        MsgBox "Bitte Methode wählen. Danke."
        End
    Else
        ' Deaktivierung von Events, Alertmeldungen und Bildschirmaktualisierung
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        
        ' Definition der Variablen
        With wsData
            .Visible = True
            strGerät = .Cells(1, 2)
            strMethode = .Cells(2, 2)
            strTopic = .Cells(3, 2) & IIf(.Cells(10, 2) = "STD", "", "-" & .Cells(10, 2)) & "_"
            strPfad = "L:\UnilabUltimateBatches\ZH_Equipment\"
            strDatum = "ZH_" & Format(.Cells(8, 2), "yyyyMMdd") & "_"
            strDatei = "*.xlsx"
            strVolleDatei = Dir(strPfad & strDatum & strTopic & strDatei)
            If Not strVolleDatei = "" Then strOperatorTest = Split(strVolleDatei, "_")(4)
            .Visible = False
        End With
        
        ' Öffnen der Datenbank und Extrahieren relevanter Daten
        Set DatenWB = Workbooks.Open(strMethodedaten)
        With DatenWB.Sheets(strGerät)
            arrQuellKolonne = .Range(.Cells(2, 2), .Cells(2, Columns.Count).End(xlToLeft))
            intMethodenZeile = .Columns(Application.Match("Methode", arrQuellKolonne, 0) + 1).Find(strMethode).Row
            dblStdEinwaage = .Cells(intMethodenZeile, Application.Match("Standard-einwaage", arrQuellKolonne, 0) + 1)
        End With
        DatenWB.Close (False)
        
        ' Erstellen des Sequencearrays
        i = 0
        Do Until strVolleDatei = ""
            If InStr(strVolleDatei, strOperatorTest) = 0 Then intOperatorCount = intOperatorCount + 1
            ReDim Preserve arrExporte(i)
            arrExporte(i) = strVolleDatei
            i = i + 1
            strVolleDatei = Dir
        Loop
        
        ' Benutzerabfrage, falls mehrere Operatoren vorhanden sind
        If intOperatorCount > 0 Then
            strName = InputBox("Unter dem heutigem Datum sind Dateien von verschiedenen Personen vorhanden. Bitte dein Kürzel eintragen und auf ""OK"" klicken.")
            ' Enfernen der nicht gewolten Sequencen
            For i = UBound(arrExporte) To LBound(arrExporte) Step -1
                If InStr(arrExporte(i), strName) = 0 Then
                    ' Element entfernen, wenn es den gesuchten Inhalt nicht enthält
                    For j = i To UBound(arrExporte) - 1
                        arrExporte(j) = arrExporte(j + 1)
                    Next j
                    ' Redimensionieren des Arrays, um das letzte Element zu entfernen
                    If UBound(arrExporte) > 0 Then
                        ReDim Preserve arrExporte(LBound(arrExporte) To UBound(arrExporte) - 1)
                    Else
                        Erase arrExporte
                    End If
                End If
            Next i
        End If
        
        ' Überprüfung, ob Dateien vorhanden sind und ob der Benutzer vorhanden ist
        If funcIsArrayEmpty(arrExporte) = True Then
            MsgBox "Keine Daten für das Importieren gefunden." & vbCr & "Vergewissere dich bitte, ob ein Batch für diese Methode und Datum existiert und ob dieser den Status Action hat." & vbCr & "Bei Fragen wende dich bitte an den Digital Laboratory Expert." & vbCr & "Danke.", vbCritical, "Keine Daten gefunden."
            wsHauptseite.Protect
            GoTo SaveExit
            End
        End If
        
        ' Sortieren der Sequencen
        defQuickSort arrExporte, LBound(arrExporte), UBound(arrExporte)
        
        ' Durchlaufen der Sequencen und Importieren der Daten
        For i = 0 To UBound(arrExporte)
            Set ZWB = ActiveWorkbook
            Set ZWS = ActiveSheet
            Workbooks.Open strPfad & arrExporte(i)
            Set QWB = ActiveWorkbook
            strName = Split(arrExporte(i), "_")(4)
            With ZWB.Sheets("Hauptseite")
                ' Kopieren und Einfügen der Probenummern und Producktklassen
                Range(Cells(1, 1), Cells(Cells(Rows.Count, 1).End(xlUp).Row, 1)).Copy: .Cells(.Cells(Rows.Count, 2).End(xlUp).Row + 1, 2).PasteSpecial Paste:=xlPasteValues
                Range(Cells(1, 5), Cells(Cells(Rows.Count, 5).End(xlUp).Row, 5)).Copy: .Cells(.Cells(Rows.Count, 5).End(xlUp).Row + 1, 5).PasteSpecial Paste:=xlPasteValues
                '",." Einwaagekorrektur
                .Range(.Cells(1, 2), .Cells(.Cells(Rows.Count, 2).End(xlUp).Row, 2)).Replace What:=",", Replacement:=".", LookAt:=xlPart
                For Each rngZelle In .Range(.Cells(1, 2), .Cells(.Cells(Rows.Count, 2).End(xlUp).Row, 2)).Cells
                    If rngZelle > 50 And IsNumeric(rngZelle) Then rngZelle.value = rngZelle.value / 1000
                Next
                ' Einwaage einfügen
                For intZeile = 1 To Cells(Rows.Count, 3).End(xlUp).Row
                    If Cells(intZeile, 5) Like "*LEATHER*" Then
                        '@TODO Code, der ausgeführt werden soll, wenn strTopic gleich "LEATHER" ist
                        .Cells(.Cells(Rows.Count, 3).End(xlUp).Row + 1, 3) = 0.001
                    Else
                        arrEinzelEinwaagen = Split(Cells(intZeile, 2), "/")
                        For j = 0 To UBound(arrEinzelEinwaagen)
                            dblSumme = dblSumme + CDbl(arrEinzelEinwaagen(j))
                        Next j
                        .Cells(.Cells(Rows.Count, 3).End(xlUp).Row + 1, 3) = dblSumme
                        dblSumme = 0
                    End If
                    .Cells(.Cells(Rows.Count, 4).End(xlUp).Row + 1, 4) = dblStdEinwaage / .Cells(.Cells(Rows.Count, 3).End(xlUp).Row, 3)
                Next intZeile
                QWB.Close Savechanges:=False
                .Cells(4, 9) = strName
            End With
        Next i
    End If

    ' Aktivierung von Events, Alertmeldungen und Bildschirmaktualisierung
SaveExit:
    

    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Private Sub Sequence()

    Dim strPositionMeasage As String: strPositionMeasage = "Die folgenden Kategorien beginnen an diesen Positionen:" & Chr(10) & Chr(10)
    Dim intAnzahlZwischenkalibration As Integer
    Dim objProbe As Object
    Dim objMessung As Object
    Dim Datenbank As clsMethodenLoader
    Dim Kategorie As MessTypen
    Dim dictMaxUsage As Object
    
    Set dictMaxUsage = CreateObject("Scripting.Dictionary")
    Set Datenbank = New clsMethodenLoader
    
    Datenbank.Init wsData, wsHauptseite, wsSequence, strMethodedaten
    
    ' Daten Auslesen
    ' Sequencedaten
    Datenbank.setBatchdaten
    If Datenbank.dictBatchdaten("Methode") = "Methode" Then
        MsgBox "Bitte Methode wählen. Danke.", vbExclamation + vbOKOnly, "Fehler beim Methodeauswahl"
        Exit Sub
    Else
        With Datenbank
            Application.EnableEvents = False
            Application.DisplayAlerts = False
            Application.ScreenUpdating = False
    
            ' Methodenwerte auslesen
            Workbooks.Open strMethodedaten
            .setDatenWorkbook
            .setMethodenZeile
            .setWertePosition
            .setExportordner
            
            ' Methodenwerte auslesen
            .setMethodendaten
            .setBlindwerte
            .setKalibrationen
             
            ' Spetialprobenwerte auslesen
            .setSpezialproben
                    
            ' Probenwerte auslesen
            .setProben
            
            ' Trigger Definieren
            .setTrigger
            .setGanzspalten
            
            ''' Sequence in Collection laden '''
            ' Anfangskalibration
            Call .getBlank
            Call .getKalibration(Volle_Kalibration:=True, Zwischenkalibration:=False)
            Call .getBlank
            
            ' Spezialproben
            If Not .colSpezialproben.Count = 0 Then
                Call .getSpezialproben
                Call .getBlank
            End If
            
            For Each objProbe In .colProben
                ' Probe
                .intZeileSequence = .intZeileSequence + 1
                .colRawSequence.Add objProbe
                .dictMetadaten("Trigger")("CurrentKalibrationTriggerCount") = .dictMetadaten("Trigger")("CurrentKalibrationTriggerCount") + 1
                .dictMetadaten("Trigger")("CurrentBlankTriggerCount") = .dictMetadaten("Trigger")("CurrentBlankTriggerCount") + 1
                
                ' Zwischenkali
                If .dictMetadaten("Trigger")("CurrentKalibrationTriggerCount") = .dictMetadaten("Trigger")("AnzahlProbenZwischenKalibrationen") _
                   And intAnzahlZwischenkalibration < .dictTrigger("MaxKalibration") Then ' diese Zeile kappt unnötige Zwischenkalibrationen am Ende der Sequence
                    intAnzahlZwischenkalibration = intAnzahlZwischenkalibration + 1
                    Call .getBlank
                    Call .getKalibration(Volle_Kalibration:=.dictMethodedaten("ZwischenkaliEinzel_Volle") = "Volle", Zwischenkalibration:=True)
                    Call .getBlank
                End If
                
                ' Zwischenblank
                If .dictMetadaten("Trigger")("CurrentBlankTriggerCount") = .dictMetadaten("Trigger")("AnzahlProbenZwischenBlank") _
                   And .dictMetadaten("Trigger")("CurrentKalibrationTriggerCount") < .dictMetadaten("Trigger")("AnzahlProbenZwischenKalibrationen") Then ' diese Zeile kappt unnötige Zwischenblanks vor den Zwischenkalibrationen
                    Call .getBlank
                End If
                
            Next objProbe
            ' Schlusskalibration
            Call .getBlank
            Call .getKalibration(Volle_Kalibration:=True, Zwischenkalibration:=False)
            Call .getBlank
           
                
            ' Position anlegen
            With dictMaxUsage
                .Add Sample, 1
                .Add Spezialprobe, 1
                .Add Zwischenkalibration, Datenbank.dictMethodedaten("KalWechsel")
                .Add Kalibration, Datenbank.dictMethodedaten("KalWechsel")
                .Add Blank, Datenbank.dictMethodedaten("BlankWechsel")
            End With
            
            .intPosition = .dictBatchdaten("Position")
            
            For Kategorie = Blank To Sample Step -1
                strPositionMeasage = strPositionMeasage + " - " & funcGetMesstypName(Kategorie) & " von " & .intPosition
                Call .setUpdatePosition(KategorieConst:=Kategorie, maxUsage:=dictMaxUsage(Kategorie), UseLevel:=(Kategorie = Kalibration))
                strPositionMeasage = strPositionMeasage + " bis " & .intPosition - 1 & Chr(10)
            Next Kategorie
            
            ' Sortiere die ganze Collection
            Call defSortCollectionByIndex(.colFinalSequence)
        End With
                
        ''' Sequence ins Excel schreiben '''
        With wsSequence
            .Visible = True
            .Cells.ClearContents
            Datenbank.intZeileSequence = 1
            For Each objMessung In Datenbank.colFinalSequence
                Datenbank.intZeileSequence = Datenbank.intZeileSequence + 1
                For prpName = AcquisitionMethode To Wert4
                    If Not Datenbank.dictMetadaten("Kolonnenposition")(funcGetPropertyName(prpName)) = -1 Then _
                       .Cells(Datenbank.intZeileSequence, Datenbank.dictMetadaten("Kolonnenposition")(funcGetPropertyName(prpName))) = funcGetWert(prp:=prpName, Messung:=objMessung)
                Next prpName
            Next objMessung
            
            ' Ganzbatchkolonnen
            For prpName = Sequencename To Sequencename
                If Not Datenbank.dictMetadaten("Kolonnenposition")(funcGetPropertyName(prpName)) = -1 Then _
                   .Range(.Cells(1, Datenbank.dictMetadaten("Kolonnenposition")(funcGetPropertyName(prpName))), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, Datenbank.dictMetadaten("Kolonnenposition")(funcGetPropertyName(prpName)))) = _
                   funcGetWert(prp:=prpName, Ganzspalten:=Datenbank.objGanzspalten)
            Next prpName
        
            ' Sequence in Clipboard überführen oder exportieren
            If Datenbank.strExportordner = "-1\" Then
                .UsedRange.Copy
                Workbooks("Book1").Sheets(1).Cells(1, 1).PasteSpecial xlPasteAll
            ElseIf Not funcIsFileOpen(Datenbank.strExportordner & Datenbank.dictMetadaten("Batchdaten")("Methode") & "_" & Datenbank.dictMetadaten("Batchdaten")("Topic") & ".csv") Then
                ActiveWorkbook.SaveAs filename:=Datenbank.strExportordner & Datenbank.dictMetadaten("Batchdaten")("Methode") & "_" & Datenbank.dictMetadaten("Batchdaten")("Topic"), FileFormat:=xlCSV, Local:=True
            Else
                MsgBox "Der Export ist noch geöffnet und kann daher nicht abgespeichert werden." & vbCrLf & "Bitte schliesse die Datei, bevor du erneut die Sequence exportierst.", Buttons:=vbExclamation + vbOKOnly, Title:="Exportfehler - Datei ist noch geöffnet"
            End If
            .Visible = False
        End With
    End If
    
    Datenbank.dictMetadaten("wbDaten").Close (False)
    
    Set Datenbank = Nothing
    Set dictMaxUsage = Nothing
    
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    'MsgBox strPositionMeasage, vbInformation, "Positionshilfe"
    Debug.Print strPositionMeasage
    
End Sub

Sub Ausdruck()
    
    '--- String-Variablen ---
    Dim EndOrdner As String            'Zielordner
    Dim strGerätename As String        'Name des ausgewählten Geräts
    Dim strMethode As String           'Aktive Methode
    Dim strOperator As String          'Bediener/Benutzername
    Dim strQuellOrdner As String       'Quelldatenordner
    
    '--- Numerische Variablen ---
    Dim c As Long                      'Laufvariable (z. B. Schleifen)
    Dim intAusdruckZeile As Integer    'Zeile für Ausdruck
    Dim intGeräteZeile As Integer      'Zeile für Geräteauswahl
    Dim intSpalte As Integer           'Spaltenindex
    
    '--- Objektvariablen ---
    Dim dictKolonnenposition As Object 'Dictionary für Kolonnenpositionen
    Set dictKolonnenposition = CreateObject("Scripting.Dictionary")
    
    '--- Arrays ---
    Dim arrFelder As Variant           'Feldliste
    arrFelder = Array("Beschriftung", "Sequencename", "Typ", "Position", "Rack", "Level", "Verdünnung")


    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Call Sequence
    ' Kopfzeile
    With Sheets("data")
        strGerätename = .Cells(1, 2)
        strMethode = .Cells(2, 2)
        strOperator = .Cells(4, 2)
    End With
    With Sheets("Ausdruck")
        .Visible = True
        .Activate
        .Unprotect
    End With
    With Sheets("Sequence")
        .Visible = True
        intAusdruckZeile = Columns(2).Find("Name").Row + 1
        Range(Cells(intAusdruckZeile, 2), Cells(Rows.Count, 8)).ClearContents
        Cells(4, 3) = strGerätename
        Cells(5, 3) = strMethode
        Cells(6, 3) = strOperator
        Cells(7, 3) = Now
        
        ' Positionen der Werte definieren (-1 wenn nicht verlangt)
        Workbooks.Open strMethodedaten
        intGeräteZeile = ActiveWorkbook.Sheets(1).Columns(4).Find(Environ("Computername")).Row
        With dictKolonnenposition
            ' Werte für Sequence
            For prpName = AcquisitionMethode To Wert4
                .Add funcGetPropertyName(prpName), ActiveWorkbook.Sheets(1).Cells(intGeräteZeile, prpName).value
            Next prpName
            'Werte für die ganze Sequence
            intSpalte = Wert4
            For prpName = Sequencename To Sequencename
                intSpalte = intSpalte + 1        ' Excel-Spalte direkt fortlaufend
                .Add funcGetPropertyName(prpName), ActiveWorkbook.Sheets(1).Cells(intGeräteZeile, intSpalte).value
            Next prpName
        End With
        ActiveWorkbook.Close Savechanges:=False
        
        ' Werte einfügen, falls verlangt
        i = .Cells(Rows.Count, 1).End(xlUp).Row
        For c = LBound(arrFelder) To UBound(arrFelder)
            If Not dictKolonnenposition(arrFelder(c)) = -1 Then
                .Range( _
                    .Cells(2, dictKolonnenposition(arrFelder(c))), _
                    .Cells(i, dictKolonnenposition(arrFelder(c))) _
                ).Copy
                Cells(intAusdruckZeile, 2 + c).PasteSpecial xlPasteValues
            End If
        Next c
    End With
    
    ' Farben anpassen
    For intSpalte = intAusdruckZeile To Cells(Rows.Count, 2).End(xlUp).Row
        If intSpalte Mod 2 = 0 Then
            Range(Cells(intSpalte, 2), Cells(intSpalte, 8)).Interior.Color = RGB(235, 241, 222)
        End If
    Next intSpalte
    
    ' Ausdruck Exporieren
    Sheets("Ausdruck").Protect
    Sheets("Ausdruck").Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs "L:\Makros\Zwischenspeicher\Sequence Zwischenspeicher\" & Format(Date, "YYMMdd") & "_" & strMethode & "_" & strGerätename & "_" & strOperator & ".xlsx"
    Application.DisplayAlerts = True
    ActiveWorkbook.Close (False)
    
    ' Alte Exporte ins Archiv verschieben
    Dim FSO As Object, Datei As Object
    Dim strOrdner As String, s As String
    strQuellOrdner = "L:\Makros\Zwischenspeicher\Sequence Zwischenspeicher\"
    EndOrdner = "L:\Makros\Zwischenspeicher\Sequence Zwischenspeicher\Archiv\"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    For Each Datei In FSO.GetFolder(strQuellOrdner).Files
        If FSO.FileExists(EndOrdner & Datei.Name) Then FSO.DeleteFile EndOrdner & Datei.Name
        If DateDiff("n", Now, FSO.GetFile(Datei).DateCreated) < -10 Then FSO.MoveFile Source:=Datei, Destination:=EndOrdner
    Next
    On Error GoTo 0
    For Each Datei In FSO.GetFolder(EndOrdner).Files
        If DateDiff("w", Now, FSO.GetFile(Datei).DateCreated) < -2 Then FSO.DeleteFile Datei
    Next
    Set FSO = Nothing
    Sheets("Ausdruck").Visible = False
    Sheets("Sequence").Visible = False
    MsgBox "Export wurde Ausgeführt.", vbInformation, "Done"

    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

Sub Kill()

    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Range(Cells(3, 2), Cells(432, 6)).ClearContents
    Cells(3, 9) = "Methode"
    Cells(3, 10) = "STD"
    Range(Cells(4, 9), Cells(5, 9)).ClearContents
    Cells(5, 10) = 1
    Cells(9, 11) = Date

    With Sheets("Ausdruck")
        .Visible = True
        .Unprotect
        .Range(.Rows(.Columns(2).Find("Name").Row + 1), .Rows(.Cells(.Rows.Count, 2).End(xlUp).Row)).Delete
        .Protect
    End With

    With Sheets("Sequence")
        .Visible = True
        .Cells.ClearContents
        .Visible = False
    End With

    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub


