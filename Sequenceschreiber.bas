Attribute VB_Name = "Main"
Option Explicit

Public i As Integer, j  As Integer
Public Const strMethodedaten As String = "L:\Makros\Sequenceschreiber\Daten für Sequenceschreiber.xlsx"
Public Const lngArbitraryBigNumber As Long = 33550366

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
                ' ",." Einwaagekorrektur
                .Range(.Cells(1, 2), .Cells(.Cells(Rows.Count, 2).End(xlUp).Row, 2)).Replace What:=",", Replacement:=".", LookAt:=xlPart
                For Each rngZelle In .Range(.Cells(1, 2), .Cells(.Cells(Rows.Count, 2).End(xlUp).Row, 2)).Cells
                    If rngZelle > 50 And IsNumeric(rngZelle) Then rngZelle.value = rngZelle.value / 1000
                Next
                ' Einwaage einfügen
                For intZeile = 1 To Cells(Rows.Count, 3).End(xlUp).Row
                    If Cells(intZeile, 5) Like "*LEATHER*" Then
                        Workbooks.Open "L:\Makros\Trockenmasse\Trockenmasse-Original.xlsm"
                        Columns(10).Hidden = False
                        varProbennameTeile = Split(ThisWorkbook.Sheets("BatchEquipmentExport").Cells(intZeile, 1), ".")
                        intDoppelbestimmung = IIf(IsNumeric(varProbennameTeile(UBound(varProbennameTeile) - 1)), 0, 1)
                        Set rngSample = Range(Cells(12, 10), Cells(Cells(Rows.Count, 10).End(xlUp).Row, 10)).Find(varProbennameTeile(0) & "." & varProbennameTeile(1) & "." & Left(varProbennameTeile(1), Len(varProbennameTeile(1) - intDoppelbestimmung)), LookIn:=xlValues)
                        .Cells(.Cells(Rows.Count, 3).End(xlUp).Row + 1, 3) = IIf(rngSample Is Nothing, 0.001, ThisWorkbook.Sheets("BatchEquipmentExport").Cells(intZeile, 2) - (ThisWorkbook.Sheets("BatchEquipmentExport").Cells(intZeile, 2) * Cells(rngSample.Row + 1, 9) / 100))
                        Workbooks("Trockenmasse-Original.xlsm").Close savechanges:=False
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
                QWB.Close savechanges:=False
                .Cells(4, 9) = strName
            End With
        Next i
    End If
    
    ' Asudruck Sequencen übertragen
    With wsAusdruck
        .Visible = xlSheetVisible
        .Unprotect
        For i = 0 To UBound(arrExporte)
            If Not i = 0 Then .Rows(10 + i).EntireRow.Insert
            .Cells(9 + i, 3) = arrExporte(i)
        Next i
        .Protect
        .Visible = xlSheetHidden
    End With
    
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
        End
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
            .setSaveSpaces
            .setMethodendaten
            
            ' Messwerte auslesen
            .setBlindwerte
            .setKalibrationen
            .setSpezialproben
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
                j = .intPosition
                Call .setUpdatePosition(KategorieConst:=Kategorie, maxUsage:=dictMaxUsage(Kategorie), UseLevel:=(Kategorie = Kalibration))
                If Not j = .intPosition Or Kategorie = Kalibration Then strPositionMeasage = strPositionMeasage + " - " & funcGetMesstypName(Kategorie) & " ab " & j & Chr(10)
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
                'Workbooks("Book1").Sheets(1).Cells(1, 1).PasteSpecial xlPasteAll
            ElseIf Not funcIsFileOpen(Datenbank.strExportordner & Datenbank.dictMetadaten("Batchdaten")("Methode") & "_" & Datenbank.dictMetadaten("Batchdaten")("Topic") & ".csv") Then
                ActiveWorkbook.SaveAs filename:=Datenbank.strExportordner & Datenbank.dictMetadaten("Batchdaten")("Methode") & "_" & Datenbank.dictMetadaten("Batchdaten")("Topic"), FileFormat:=xlCSV, Local:=True
            Else
                MsgBox "Der Export ist noch geöffnet und kann daher nicht abgespeichert werden." & vbCrLf & "Bitte schliesse die Datei, bevor du erneut die Sequence exportierst.", Buttons:=vbExclamation + vbOKOnly, Title:="Exportfehler - Datei ist noch geöffnet"
            End If
            .Visible = False
        End With
    End If
    
    'If Datenbank.dictMethodedaten("BlankWechsel") + Datenbank.dictMethodedaten("KalWechsel") > 0 Then MsgBox strPositionMeasage, vbInformation, "Positionshilfe"
    Datenbank.dictMetadaten("wbDaten").Close (False)
    
    Set Datenbank = Nothing
    Set dictMaxUsage = Nothing
    
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

Sub Ausdruck()
    
    '--- String-Variablen ---
    Dim EndOrdner As String            'Zielordner
    Dim strGerätename As String        'Name des ausgewählten Geräts
    Dim strMethode As String           'Aktive Methode
    Dim strOperator As String          'Bediener/Benutzername
    Dim strQuellOrdner As String       'Quelldatenordner
    Dim strCommend As String           'Kommentar für Batchflow
    Dim strUserMail As String              'User für Batchflow
    
    '--- Numerische Variablen ---
    Dim intAusdruckZeile As Integer    'Zeile für Ausdruck
    Dim intGeräteZeile As Integer      'Zeile für Geräteauswahl
    Dim intSpalte As Integer           'Spaltenindex
    
    '--- Arrays ---
    Dim arrFelder As Variant           'Feldliste
    arrFelder = Array("Beschriftung", "Sequencename", "Typ", "Position", "Rack", "Level", "Verdünnung")
    
    '--- Range ---
    Dim rng As Range
    
    '--- Datenbank ---
    Dim Datenbank As clsMethodenLoader
    Set Datenbank = New clsMethodenLoader
    Datenbank.Init wsData, wsHauptseite, wsSequence, strMethodedaten

    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Call Sequence
    ' Kopfzeile
    With wsData
        strGerätename = .Cells(1, 2)
        strMethode = .Cells(2, 2)
        strOperator = .Cells(4, 2)
        If Not .Cells(11, 2) = 0 Then strCommend = .Cells(11, 2)
    End With
    With wsUser
        .Visible = xlSheetVisible
        i = .Cells(.Rows.Count, 1).End(xlUp).Row
        Set rng = .Range(.Cells(2, 1), .Cells(i, 1)).Find(What:=strOperator, LookIn:=xlValues, LookAt:=xlWhole)
        If Not rng Is Nothing Then
            ' Kürzel gefunden -> E-Mail übernehmen
            strUserMail = .Cells(rng.Row, 2).value
        Else
            ' Kürzel nicht gefunden -> InputBox für neue E-Mail
            strUserMail = InputBox("Kürzel nicht gefunden. Bitte gib deine E-Mail ein:", "e-Mail", "@testex.com")
            
            ' Neuen Eintrag in die Liste schreiben
            .Cells(i + 1, 1).value = strOperator
            .Cells(i + 1, 2).value = strUserMail
        End If
        .Visible = xlSheetHidden
    End With
    With wsAusdruck
        .Visible = True
        .Activate
        .Unprotect
    End With
    ' Metadaten Definieren
    Workbooks.Open strMethodedaten
    With Datenbank
        .setBatchdaten
        .setDatenWorkbook
        .setMethodenZeile
        .setWertePosition
        .setSaveSpaces
        .setMethodendaten
            
        ' Messwerte auslesen
        .setBlindwerte
        .setKalibrationen
        .setSpezialproben
        .setProben
    End With
    With wsSequence
        .Visible = True
        intAusdruckZeile = wsAusdruck.Columns(2).Find("Name").Row + 1
        Range(Cells(intAusdruckZeile, 2), Cells(Rows.Count, 8)).ClearContents
        wsAusdruck.Cells(4, 3) = strGerätename
        wsAusdruck.Cells(5, 3) = strMethode
        wsAusdruck.Cells(6, 3) = strOperator
        wsAusdruck.Cells(7, 3) = Now
        wsAusdruck.Cells(8, 3) = Datenbank.strSpeicherort
        
        ' Werte einfügen, falls verlangt
        i = .Cells(Rows.Count, 1).End(xlUp).Row
        For j = LBound(arrFelder) To UBound(arrFelder)
            If Not Datenbank.dictKolonnenposition(arrFelder(j)) = -1 Then
                .Range( _
                    .Cells(2, Datenbank.dictKolonnenposition(arrFelder(j))), _
                    .Cells(i, Datenbank.dictKolonnenposition(arrFelder(j))) _
                ).Copy
                wsAusdruck.Cells(intAusdruckZeile, 2 + j).PasteSpecial xlPasteValues
            End If
        Next j
    End With
    
    
    ' Farben anpassen
    Dim typValue As Variant
    Dim fontColor As Long
    For i = intAusdruckZeile To Cells(Rows.Count, 2).End(xlUp).Row
        Set rng = wsAusdruck.Range(wsAusdruck.Cells(i, 2), wsAusdruck.Cells(i, 8))
    
        ' Hintergrund
        If i Mod 2 = 0 Then
            rng.Interior.Color = RGB(235, 241, 222)
        End If
        
        ' Font-Farbe
        typValue = wsAusdruck.Cells(i, 4).value

        ' Default-Farbe
        fontColor = RGB(0, 0, 0)
        
        Select Case True
            Case typValue = Datenbank.objBlank.Typ
                fontColor = RGB(0, 112, 192)       ' Blau
            Case typValue = Datenbank.colKalibration(1).Typ
                fontColor = RGB(255, 0, 0)         ' Rot
            Case Datenbank.colSpezialproben.Count > 0
                If typValue = Datenbank.colSpezialproben(1).Typ Then fontColor = RGB(0, 176, 80)        ' Grün
        End Select
        
        rng.Font.Color = fontColor
    Next i
    
    ' Ausdruck Exporieren
    wsAusdruck.Protect
    wsAusdruck.Copy
    Application.DisplayAlerts = False
    'ActiveWorkbook.SaveAs "L:\Makros\Zwischenspeicher\Sequence Zwischenspeicher\" & Format(Date, "YYMMdd") & "_" & strMethode & "_" & strGerätename & "_" & strOperator & ".xlsx"
    ActiveWorkbook.SaveAs "https://testex.sharepoint.com/sites/TZHECOLabOrga/Shared Documents/Operation/SequenceExporte/" & Format(Date, "YYMMdd") & "_" & strMethode & "_" & strGerätename & "_" & strOperator & ".xlsx"
    Application.DisplayAlerts = True
    ActiveWorkbook.Close (False)

    Dim http As Object
    Dim url As String
    Dim JSONBody As String

    ' URL deines Flows
    url = "https://default0de4a018140e49e5aa27ff79659f36.0e.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/bf92aecbcc984838ae59193a8881cb0d/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=cMxKL-AXW1Qo_jz118AO-5LWMnIFv-3CsHhl_3WvIjU"
    
    ' JSON aus Excel-Daten zusammenstellen
    JSONBody = "{""title"":""" & Format(Date, "YYMMdd") & "_" & strMethode & "_" & strGerätename & "_" & strOperator & """,""team"":""" & Datenbank.dictMethodedaten("Team") & """,""commend"":""" & strCommend & """,""user"":""" & strUserMail & """}"

    ' HTTP POST Request
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send JSONBody

    If Not http.Status = 200 And Not http.Status = 202 Then
        MsgBox "Beim Versand der Datei ist ein Fehler aufgetreten. Bitte versuche es erneut. Falls das Problem bestehen bleibt, melde dich bitte beim DLE." & Chr(10) & Chr(10) & "Fehler: " & http.Status & " - " & http.responseText, vbCritical, "Versandfehler"
    End If
    
    Datenbank.dictMetadaten("wbDaten").Close savechanges:=False
    wsAusdruck.Visible = xlSheetHidden
    wsUser.Visible = xlSheetHidden
    wsSequence.Visible = xlSheetHidden
    
    MsgBox "Export wurde Ausgeführt." & Chr(10) & "Bitte überprüfe den Batchflow.", vbInformation, "Done"

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
        Do Until IsEmpty(.Cells(10, 3))
            .Rows(10).Delete
        Loop
        If Not IsEmpty(.Cells(13, 2)) Then .Range(.Rows(.Columns(2).Find("Name").Row + 1), .Rows(.Cells(.Rows.Count, 2).End(xlUp).Row)).Delete
        .Protect
        .Visible = False
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


