Attribute VB_Name = "Main"
Option Explicit

Public i As Integer, j  As Integer
Public Const strMethodeData As String = "L:\Makros\Sequenceschreiber\Daten für Sequenceschreiber.xlsx"
Public Const lngArbitraryBigNumber As Long = 33550366

Private Sub Import()
    ' Strings
    Dim strEquipment As String
    Dim strMethod As String
    Dim strTopic As String
    Dim strPath As String
    Dim strDate As String
    Dim strFile As String
    Dim strFullFile As String
    Dim strOperatorTest As String
    Dim strName As String
    
    ' Workbooks und Worksheets
    Dim DataWB As Workbook
    Dim ZWB As Workbook
    Dim QWB As Workbook
    Dim ZWS As Worksheet
    
    ' Ranges
    Dim rngCell As Range
    
    ' Arrays
    Dim arrSourceColumns As Variant
    Dim arrExports() As Variant
    Dim arrIndividualWeighings As Variant
    
    ' Integers und Doubles
    Dim intMethodRow As Integer
    Dim intOperatorCount As Integer
    Dim intRow As Integer
    Dim dblStdWeighings As Double
    Dim dblSum As Double

    
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
            strEquipment = .Cells(1, 2)
            strMethod = .Cells(2, 2)
            strTopic = .Cells(3, 2) & IIf(.Cells(10, 2) = "STD", "", "-" & .Cells(10, 2)) & "_"
            strPath = "L:\UnilabUltimateBatches\ZH_Equipment\"
            strDate = "ZH_" & Format(.Cells(8, 2), "yyyyMMdd") & "_"
            strFile = "*.xlsx"
            strFullFile = Dir(strPath & strDate & strTopic & strFile)
            If Not strFullFile = "" Then strOperatorTest = Split(strFullFile, "_")(4)
            .Visible = False
        End With
        
        ' Öffnen der Datenbank und Extrahieren relevanter Daten
        Set DataWB = Workbooks.Open(strMethodeData)
        With DataWB.Sheets(strEquipment)
            arrSourceColumns = .Range(.Cells(2, 2), .Cells(2, Columns.Count).End(xlToLeft))
            intMethodRow = .Columns(Application.Match("Methode", arrSourceColumns, 0) + 1).Find(strMethod).Row
            dblStdWeighings = .Cells(intMethodRow, Application.Match("Standard-einwaage", arrSourceColumns, 0) + 1)
        End With
        DataWB.Close (False)
        
        ' Erstellen des Sequencearrays
        i = 0
        Do Until strFullFile = ""
            If InStr(strFullFile, strOperatorTest) = 0 Then intOperatorCount = intOperatorCount + 1
            ReDim Preserve arrExports(i)
            arrExports(i) = strFullFile
            i = i + 1
            strFullFile = Dir
        Loop
        
        ' Benutzerabfrage, falls mehrere Operatoren vorhanden sind
        If intOperatorCount > 0 Then
            strName = InputBox("Unter dem heutigem Datum sind Dateien von verschiedenen Personen vorhanden. Bitte dein Kürzel eintragen und auf ""OK"" klicken.")
            ' Enfernen der nicht gewolten Sequencen
            For i = UBound(arrExports) To LBound(arrExports) Step -1
                If InStr(arrExports(i), strName) = 0 Then
                    ' Element entfernen, wenn es den gesuchten Inhalt nicht enthält
                    For j = i To UBound(arrExports) - 1
                        arrExports(j) = arrExports(j + 1)
                    Next j
                    ' Redimensionieren des Arrays, um das letzte Element zu entfernen
                    If UBound(arrExports) > 0 Then
                        ReDim Preserve arrExports(LBound(arrExports) To UBound(arrExports) - 1)
                    Else
                        Erase arrExports
                    End If
                End If
            Next i
        End If
        
        ' Überprüfung, ob Dateien vorhanden sind und ob der Benutzer vorhanden ist
        If funcIsArrayEmpty(arrExports) = True Then
            MsgBox "Keine Daten für das Importieren gefunden." & vbCr & "Vergewissere dich bitte, ob ein Batch für diese Methode und Datum existiert und ob dieser den Status Action hat." & vbCr & "Bei Fragen wende dich bitte an den Digital Laboratory Expert." & vbCr & "Danke.", vbCritical, "Keine Daten gefunden."
            wsMainPage.Protect
            GoTo SaveExit
            End
        End If
        
        ' Sortieren der Sequencen
        defQuickSort arrExports, LBound(arrExports), UBound(arrExports)
        
        ' Durchlaufen der Sequencen und Importieren der Daten
        For i = 0 To UBound(arrExports)
            Set ZWB = ActiveWorkbook
            Set ZWS = ActiveSheet
            Workbooks.Open strPath & arrExports(i)
            Set QWB = ActiveWorkbook
            strName = Split(arrExports(i), "_")(4)
            With ZWB.Sheets("Hauptseite")
                ' Kopieren und Einfügen der Probenummern und Producktklassen
                Range(Cells(1, 1), Cells(Cells(Rows.Count, 1).End(xlUp).Row, 1)).Copy: .Cells(.Cells(Rows.Count, 2).End(xlUp).Row + 1, 2).PasteSpecial Paste:=xlPasteValues
                Range(Cells(1, 5), Cells(Cells(Rows.Count, 5).End(xlUp).Row, 5)).Copy: .Cells(.Cells(Rows.Count, 5).End(xlUp).Row + 1, 5).PasteSpecial Paste:=xlPasteValues
                ' ",." Einwaagekorrektur
                .Range(.Cells(1, 2), .Cells(.Cells(Rows.Count, 2).End(xlUp).Row, 2)).Replace What:=",", Replacement:=".", LookAt:=xlPart
                For Each rngCell In .Range(.Cells(1, 2), .Cells(.Cells(Rows.Count, 2).End(xlUp).Row, 2)).Cells
                    If rngCell > 50 And IsNumeric(rngCell) Then rngCell.value = rngCell.value / 1000
                Next
                ' Einwaage einfügen
                For intRow = 1 To Cells(Rows.Count, 3).End(xlUp).Row
                    If Cells(intRow, 5) Like "*LEATHER*" Then
                        
                        Dim varProbennameTeile As Variant
                        Dim intDoppelbestimmung As Integer
                        Dim rngSample  As Range
                        
                        Workbooks.Open "L:\Makros\Trockenmasse\Trockenmasse-Original.xlsm"
                        Columns(10).Hidden = False
                        varProbennameTeile = Split(ThisWorkbook.Sheets("BatchEquipmentExport").Cells(intRow, 1), ".")
                        intDoppelbestimmung = IIf(IsNumeric(varProbennameTeile(UBound(varProbennameTeile) - 1)), 0, 1)
                        Set rngSample = Range(Cells(12, 10), Cells(Cells(Rows.Count, 10).End(xlUp).Row, 10)).Find(varProbennameTeile(0) & "." & varProbennameTeile(1) & "." & Left(varProbennameTeile(1), Len(varProbennameTeile(1) - intDoppelbestimmung)), LookIn:=xlValues)
                        .Cells(.Cells(Rows.Count, 3).End(xlUp).Row + 1, 3) = IIf(rngSample Is Nothing, 0.001, ThisWorkbook.Sheets("BatchEquipmentExport").Cells(intRow, 2) - (ThisWorkbook.Sheets("BatchEquipmentExport").Cells(intRow, 2) * Cells(rngSample.Row + 1, 9) / 100))
                        Workbooks("Trockenmasse-Original.xlsm").Close savechanges:=False
                    Else
                        arrIndividualWeighings = Split(Cells(intRow, 2), "/")
                        For j = 0 To UBound(arrIndividualWeighings)
                            dblSum = dblSum + CDbl(arrIndividualWeighings(j))
                        Next j
                        .Cells(.Cells(Rows.Count, 3).End(xlUp).Row + 1, 3) = dblSum
                        dblSum = 0
                    End If
                    .Cells(.Cells(Rows.Count, 4).End(xlUp).Row + 1, 4) = dblStdWeighings / .Cells(.Cells(Rows.Count, 3).End(xlUp).Row, 3)
                Next intRow
                QWB.Close savechanges:=False
                .Cells(4, 9) = strName
            End With
        Next i
    End If
    
    ' Asudruck Sequencen übertragen
    With wsAusdruck
        .Visible = xlSheetVisible
        .Unprotect
        For i = 0 To UBound(arrExports)
            If Not i = 0 Then .Rows(10 + i).EntireRow.Insert
            .Cells(9 + i, 3) = arrExports(i)
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
    Dim intNumberIntermediadeCalibration As Integer
    Dim objSample As Object
    Dim objMeasuring As Object
    Dim Database As clsMethodeLoader
    Dim Categorie As MeasurementTypes
    Dim dictMaxUsage As Object
    
    Set dictMaxUsage = CreateObject("Scripting.Dictionary")
    Set Database = New clsMethodeLoader
    
    Database.Init wsData, wsMainPage, wsSequence, strMethodeData
    
    ' Daten Auslesen
    ' Sequencedaten
    Database.setBatchdata
    If Database.dictBatchdata("Methode") = "Methode" Then
        MsgBox "Bitte Methode wählen. Danke.", vbExclamation + vbOKOnly, "Fehler beim Methodeauswahl"
        End
    Else
        With Database
            Application.EnableEvents = False
            Application.DisplayAlerts = False
            Application.ScreenUpdating = False
    
            ' Methodenwerte auslesen
            Workbooks.Open strMethodeData
            .setDataWorkbook
            .setMethodeRow
            .setValuePosition
            .setSaveSpaces
            .setMethodeData
            
            ' Messwerte auslesen
            .setBlankValues
            .setCalibrationValues
            .setSpecialSamleValues
            .setSampleValues
            
            ' Trigger Definieren
            .setTrigger
            .setFullColumns
            
            ''' Sequence in Collection laden '''
            ' Anfangskalibration
            For i = 1 To .dictMetaData("Batchdaten")("AnzahlStartBlanks")
                Call .getBlank
            Next i
            Call .getCalibration(FullCalibration:=True, IntermediateCalibration:=False)
            Call .getBlank
            
            ' Spezialproben
            If Not .colSpecialSamples.Count = 0 Then
                Call .getSpezialproben
                Call .getBlank
            End If
            
            For Each objSample In .colSample
                ' Probe
                .intSequenceRow = .intSequenceRow + 1
                .colRawSequence.Add objSample
                .dictMetaData("Trigger")("CurrentKalibrationTriggerCount") = .dictMetaData("Trigger")("CurrentKalibrationTriggerCount") + 1
                .dictMetaData("Trigger")("CurrentBlankTriggerCount") = .dictMetaData("Trigger")("CurrentBlankTriggerCount") + 1
                
                ' Zwischenkali
                If .dictMetaData("Trigger")("CurrentKalibrationTriggerCount") = .dictMetaData("Trigger")("AnzahlProbenZwischenKalibrationen") _
                   And intNumberIntermediadeCalibration < .dictTrigger("MaxKalibration") Then ' diese Zeile kappt unnötige Zwischenkalibrationen am Ende der Sequence
                    intNumberIntermediadeCalibration = intNumberIntermediadeCalibration + 1
                    Call .getBlank
                    Call .getCalibration(FullCalibration:=.dictMethodeData("ZwischenkaliEinzel_Volle") = "Volle", IntermediateCalibration:=True)
                    Call .getBlank
                End If
                
                ' Zwischenblank
                If .dictMetaData("Trigger")("CurrentBlankTriggerCount") = .dictMetaData("Trigger")("AnzahlProbenZwischenBlank") _
                   And .dictMetaData("Trigger")("CurrentKalibrationTriggerCount") < .dictMetaData("Trigger")("AnzahlProbenZwischenKalibrationen") Then ' diese Zeile kappt unnötige Zwischenblanks vor den Zwischenkalibrationen
                    Call .getBlank
                End If
                
            Next objSample
            ' Schlusskalibration
            Call .getBlank
            Call .getCalibration(FullCalibration:=True, IntermediateCalibration:=False)
            Call .getBlank
           
                
            ' Position anlegen
            With dictMaxUsage
                .Add Sample, 1
                .Add SpezialSample, 1
                .Add IntermediateCalibration, Database.dictMethodeData("KalWechsel")
                .Add Calibration, Database.dictMethodeData("KalWechsel")
                .Add Blank, Database.dictMethodeData("BlankWechsel")
            End With
            
            .intPosition = .dictBatchdata("Position")
            
            For Categorie = Blank To Sample Step -1
                j = .intPosition
                Call .setUpdatePosition(KategorieConst:=Categorie, maxUsage:=dictMaxUsage(Categorie), UseLevel:=(Categorie = Calibration))
                If Not j = .intPosition Or Categorie = Calibration Then strPositionMeasage = strPositionMeasage + " - " & funcGetMeasurementType(Categorie) & " ab " & j & Chr(10)
            Next Categorie
            
            ' Sortiere die ganze Collection
            Call defSortCollectionByIndex(.colFinalSequence)
        End With
                
        ''' Sequence ins Excel schreiben '''
        With wsSequence
            .Visible = True
            .Cells.ClearContents
            Database.intSequenceRow = 1
            For Each objMeasuring In Database.colFinalSequence
                Database.intSequenceRow = Database.intSequenceRow + 1
                For prpName = AcquisitionMethode To Wert4
                    If Not Database.dictMetaData("Kolonnenposition")(funcGetPropertyName(prpName)) = 0 Then _
                       .Cells(Database.intSequenceRow, Database.dictMetaData("Kolonnenposition")(funcGetPropertyName(prpName))) = funcGetValue(prp:=prpName, Messung:=objMeasuring)
                Next prpName
            Next objMeasuring
            
            ' Ganzbatchkolonnen
            For prpName = Sequencename To Sequencename
                If Not Database.dictMetaData("Kolonnenposition")(funcGetPropertyName(prpName)) = 0 Then _
                   .Range(.Cells(1, Database.dictMetaData("Kolonnenposition")(funcGetPropertyName(prpName))), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, Database.dictMetaData("Kolonnenposition")(funcGetPropertyName(prpName)))) = _
                   funcGetValue(prp:=prpName, Ganzspalten:=Database.objFullColumn)
            Next prpName
        
            ' Sequence in Clipboard überführen oder exportieren
            If Database.strExportFolder = "\" Then
                .UsedRange.Copy
                'Workbooks("Book1").Sheets(1).Cells(1, 1).PasteSpecial xlPasteAll
            ElseIf Not funcIsFileOpen(Database.strExportFolder & Database.dictMetaData("Batchdaten")("Methode") & "_" & Database.dictMetaData("Batchdaten")("Topic") & ".csv") Then
                ActiveWorkbook.SaveAs filename:=Database.strExportFolder & Database.dictMetaData("Batchdaten")("Methode") & "_" & Database.dictMetaData("Batchdaten")("Topic"), FileFormat:=xlCSV, Local:=True
            Else
                MsgBox "Der Export ist noch geöffnet und kann daher nicht abgespeichert werden." & vbCrLf & "Bitte schliesse die Datei, bevor du erneut die Sequence exportierst.", Buttons:=vbExclamation + vbOKOnly, Title:="Exportfehler - Datei ist noch geöffnet"
            End If
            .Visible = False
        End With
    End If
    
    'If Datenbank.dictMethodedaten("BlankWechsel") + Datenbank.dictMethodedaten("KalWechsel") > 0 Then MsgBox strPositionMeasage, vbInformation, "Positionshilfe"
    Database.dictMetaData("wbDaten").Close (False)
    
    Set Database = Nothing
    Set dictMaxUsage = Nothing
    
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

Sub Ausdruck()
    
    '--- String-Variablen ---
    Dim strEndFolder As String         'Zielordner
    Dim strEquipmentName As String     'Name des ausgewählten Geräts
    Dim strMethod As String            'Aktive Methode
    Dim strOperator As String          'Bediener/Benutzername
    Dim strSourceFolder As String      'Quelldatenordner
    Dim strCommend As String           'Kommentar für Batchflow
    Dim strUserMail As String          'User für Batchflow
    
    '--- Numerische Variablen ---
    Dim intPrintRow As Integer         'Zeile für Ausdruck
    Dim intEquipmentRow As Integer     'Zeile für Geräteauswahl
    Dim intColumn As Integer           'Spaltenindex
    
    '--- Arrays ---
    Dim arrFields As Variant           'Feldliste
    arrFields = Array("Beschriftung", "Sequencename", "Typ", "Position", "Rack", "Level", "Verdünnung")
    
    '--- Range ---
    Dim rng As Range
    
    '--- Datenbank ---
    Dim Datenbank As clsMethodeLoader
    Set Datenbank = New clsMethodeLoader
    Datenbank.Init wsData, wsMainPage, wsSequence, strMethodeData

    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Call Sequence
    ' Kopfzeile
    With wsData
        strEquipmentName = .Cells(1, 2)
        strMethod = .Cells(2, 2)
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
    Workbooks.Open strMethodeData
    With Datenbank
        .setBatchdata
        .setDataWorkbook
        .setMethodeRow
        .setValuePosition
        .setSaveSpaces
        .setMethodeData
            
        ' Messwerte auslesen
        .setBlankValues
        .setCalibrationValues
        .setSpecialSamleValues
        .setSampleValues
    End With
    With wsSequence
        .Visible = True
        intPrintRow = wsAusdruck.Columns(2).Find("Name").Row + 1
        Range(Cells(intPrintRow, 2), Cells(Rows.Count, 8)).ClearContents
        wsAusdruck.Cells(4, 3) = strEquipmentName
        wsAusdruck.Cells(5, 3) = strMethod
        wsAusdruck.Cells(6, 3) = strOperator
        wsAusdruck.Cells(7, 3) = Now
        wsAusdruck.Cells(8, 3) = Datenbank.strSaveFolder
        
        ' Werte einfügen, falls verlangt
        i = .Cells(Rows.Count, 1).End(xlUp).Row
        For j = LBound(arrFields) To UBound(arrFields)
            If Not Datenbank.dictColumnPosition(arrFields(j)) = 0 Then
                .Range( _
                    .Cells(2, Datenbank.dictColumnPosition(arrFields(j))), _
                    .Cells(i, Datenbank.dictColumnPosition(arrFields(j))) _
                ).Copy
                wsAusdruck.Cells(intPrintRow, 2 + j).PasteSpecial xlPasteValues
            End If
        Next j
    End With
    
    
    ' Farben anpassen
    Dim typValue As Variant
    Dim fontColor As Long
    For i = intPrintRow To Cells(Rows.Count, 2).End(xlUp).Row
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
            Case typValue = Datenbank.colCalibrations(1).Typ
                fontColor = RGB(255, 0, 0)         ' Rot
            Case Datenbank.colSpecialSamples.Count > 0
                If typValue = Datenbank.colSpecialSamples(1).Typ Then fontColor = RGB(0, 176, 80)        ' Grün
        End Select
        
        rng.Font.Color = fontColor
    Next i
    
    ' Ausdruck Exporieren
    wsAusdruck.Protect
    wsAusdruck.Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs "https://testex.sharepoint.com/sites/TZHECOLabOrga/Shared Documents/Operation/SequenceExporte/" & Format(Date, "YYMMdd") & "_" & strMethod & "_" & strEquipmentName & "_" & strOperator & ".xlsx"
    Application.DisplayAlerts = True
    ActiveWorkbook.Close (False)

    Dim http As Object
    Dim url As String
    Dim JSONBody As String

    ' URL deines Flows
    url = "https://default0de4a018140e49e5aa27ff79659f36.0e.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/bf92aecbcc984838ae59193a8881cb0d/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=cMxKL-AXW1Qo_jz118AO-5LWMnIFv-3CsHhl_3WvIjU"
    
    ' JSON aus Excel-Daten zusammenstellen
    JSONBody = "{""title"":""" & Format(Date, "YYMMdd") & "_" & strMethod & "_" & strEquipmentName & "_" & strOperator & """,""team"":""" & Datenbank.dictMethodeData("Team") & """,""commend"":""" & strCommend & """,""user"":""" & strUserMail & """}"

    ' HTTP POST Request
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send JSONBody

    If Not http.Status = 200 And Not http.Status = 202 Then
        MsgBox "Beim Versand der Datei ist ein Fehler aufgetreten. Bitte versuche es erneut. Falls das Problem bestehen bleibt, melde dich bitte beim DLE." & Chr(10) & Chr(10) & "Fehler: " & http.Status & " - " & http.responseText, vbCritical, "Versandfehler"
    End If
    
    Datenbank.dictMetaData("wbDaten").Close savechanges:=False
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


