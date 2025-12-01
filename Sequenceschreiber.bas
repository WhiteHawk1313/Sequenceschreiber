Attribute VB_Name = "Module"
Option Explicit

Public i As Integer, j  As Integer
Public prpName As Properties

' Enum für Eigenschaften
Public Enum Properties
    AcquisitionMethode = 5
    Quantmethode = 6
    Beschriftung = 7
    Einwaage = 8
    Exctraktionsvolumen = 9
    Injektionsvolumen = 10
    Kommentar = 11
    Konzentration = 12
    Position = 13
    Produktklasse = 14
    Rack = 15
    Typ = 16
    Verdünnung = 17
    Level = 18
    Info1 = 19
    Info2 = 20
    Info3 = 21
    Info4 = 22
    Wert1 = 23
    Wert2 = 24
    Wert3 = 25
    Wert4 = 26
    Messkategorie = 27                           ' ab hier Informationen zu den Messugnen, welche nicht auf die Sequence kommen
    Sequencename = 28                            ' ab hier Properties, die erst am Schluss der Sequence eingefügt werden
End Enum

Public Enum MessTypen
    Sample = 0
    Spezialprobe = 1
    Zwischenkalibration = 2
    Kalibration = 3
    Blank = 4
    Ganzspalten = 5
End Enum

Private Sub EmptyCashe()
For i = AcquisitionMethode To Messkategorie
    Debug.Print i
Next i
End Sub

Private Sub Import()
    ' Strings
    Dim strGerät As String
    Dim strMethode As String
    Dim strTopic As String
    Dim strPfad As String
    Dim strDatum As String
    Dim strDatei As String
    Dim strVolleDatei As String, strOperatorTest As String, strName As String
    Const strMethodedaten As String = "L:\Makros\Sequenceschreiber\Daten für Sequenceschreiber.xlsx"
    
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
                        ' Code, der ausgeführt werden soll, wenn strTopic gleich "LEATHER" ist
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
                QWB.Close savechanges:=False
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

Private Sub defSortCollectionByIndex(col As Collection)
    Dim arr() As Variant

    ' Collection in Array kopieren
    ReDim arr(0 To col.Count - 1)
    For i = 1 To col.Count
        Set arr(i - 1) = col.item(i)
    Next i

    ' Array sortieren
    defQuickSort arr, 0, UBound(arr), False

    ' Collection neu aufbauen
    Set col = New Collection
    For i = 0 To UBound(arr)
        col.Add arr(i)
    Next i
End Sub

Private Sub defQuickSort(arr() As Variant, ByVal low As Long, ByVal high As Long, Optional ByVal isStringArray As Boolean = True)
    Dim pivotValue As Long
    Dim tempSwap As Variant
    Dim i As Long, j As Long
    
    If low < high Then
        pivotValue = funcGetValueForSorting(arr((low + high) \ 2), isStringArray)
        i = low
        j = high
        
        Do While i <= j
            Do While funcGetValueForSorting(arr(i), isStringArray) < pivotValue
                i = i + 1
            Loop
            Do While funcGetValueForSorting(arr(j), isStringArray) > pivotValue
                j = j - 1
            Loop
            If i <= j Then
                ' Tausche
                If isStringArray Then
                    tempSwap = arr(i)
                    arr(i) = arr(j)
                    arr(j) = tempSwap
                Else
                    Set tempSwap = arr(i)
                    Set arr(i) = arr(j)
                    Set arr(j) = tempSwap
                End If
                i = i + 1
                j = j - 1
            End If
        Loop
        
        If low < j Then defQuickSort arr, low, j, isStringArray
        If i < high Then defQuickSort arr, i, high, isStringArray
    End If
End Sub

Private Function funcGetValueForSorting(item As Variant, isStringArray As Boolean) As Long
    If isStringArray Then
        funcGetValueForSorting = CLng(Split(item, "_")(3))
    Else
        funcGetValueForSorting = item.Index
    End If
End Function

Private Function funcIsArrayEmpty(arr As Variant) As Boolean
    funcIsArrayEmpty = True
    On Error Resume Next
    funcIsArrayEmpty = (LBound(arr) > UBound(arr))
    On Error GoTo 0
    
End Function

Private Function funcIsOperatorPresent(arr As Variant, strName As String) As Boolean
    On Error Resume Next
    For i = LBound(arr) To UBound(arr)
        If arr(i) Like "*" & strName & "*" Then
            funcIsOperatorPresent = True
            Exit Function
        End If
    Next i
    On Error GoTo 0
End Function

Function funcIsFileOpen(filename As String) As Boolean
    Dim filenum As Integer
    Dim errnum As Integer
    
    On Error Resume Next
    filenum = FreeFile()
    Open filename For Input Lock Read As #filenum
    Close filenum
    errnum = Err
    On Error GoTo 0
    
    ' Überprüfe, ob ein Fehler aufgetreten ist und ob die Datei geöffnet ist
    If errnum = 0 Then
        funcIsFileOpen = False
    Else
        funcIsFileOpen = True
    End If
End Function

Private Sub Sequence()
    ' Strings für Methodeninformationen
    Const strMethodedaten As String = "L:\Makros\Sequenceschreiber\Daten für Sequenceschreiber.xlsx"
    Dim strExportordner As String
    Dim strPositionMeasage As String: strPositionMeasage = "Die folgenden Kategorien beginnen an diesen Positionen:" & Chr(10) & Chr(10)
    
    ' Integer für Zeilenpositionen
    Dim intMethodenZeile As Integer
    Dim intZeileSequence As Integer
    Dim intGeräteZeile As Integer
    Dim intAnzahlZwischenkalibration As Integer
    Dim intSpalte As Integer
    Dim intPosition As Integer
    
    ' Arrays für Daten
    Dim arrQuellKolonne As Variant
    
    ' Range für Loops
    Dim rngProbenRange As Range
    Set rngProbenRange = Range(Cells(3, 2), Cells(Cells(2, 2).End(xlDown).Row, 2))
    
    ' Object und Collection
    Dim dictMetadaten As Object
    Dim dictBatchdaten As Object
    Dim dictMethodedaten As Object
    Dim dictTrigger As Object
    Dim dictKolonnenposition As Object
    Dim objProbe As Object
    Dim objProben As New CWerte
    Dim colProben As Collection
    Dim objKalibration As New CWerte
    Dim colKalibration As Collection
    Dim objBlank As New CWerte
    Dim objSpezialproben As New CWerte
    Dim colSpezialproben As Collection
    Dim objGanzspalten As New CWerte
    Dim colRawSequence As Collection
    Dim colFinalSequence As Collection
    Dim objMessung As Object
    
    Set dictMetadaten = CreateObject("Scripting.Dictionary")
    Set dictBatchdaten = CreateObject("Scripting.Dictionary")
    Set dictMethodedaten = CreateObject("Scripting.Dictionary")
    Set dictTrigger = CreateObject("Scripting.Dictionary")
    Set dictKolonnenposition = CreateObject("Scripting.Dictionary")
    Set colProben = New Collection
    Set colKalibration = New Collection
    Set colSpezialproben = New Collection
    Set colRawSequence = New Collection
    Set colFinalSequence = New Collection
    
    ' Daten Auslesen
    ' Sequencedaten
    With wsData
        .Visible = True
        .Select
        With dictBatchdaten
            .Add "Gerät", Cells(1, 2).value
            .Add "Methode", Cells(2, 2).value
            .Add "Topic", Cells(3, 2).value
            .Add "Operator", Cells(4, 2).value
            .Add "Rack", Cells(5, 2).value
            .Add "Position", Cells(6, 2).value
            .Add "AnzahlMessungen", Cells(7, 2).value
            .Add "Datum", Cells(8, 2).value
            .Add "TotalProben", Cells(9, 2).value
        End With
        wsHauptseite.Select
        .Visible = False
    End With
    Set dictMetadaten("Batchdaten") = dictBatchdaten
    
    If dictBatchdaten("Methode") = "Methode" Then
        MsgBox "Bitte Methode wählen. Danke.", vbExclamation + vbOKOnly, "Fehler beim Methodeauswahl"
        Exit Sub
    Else
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        
        ' Methodenwerte auslesen
        Workbooks.Open strMethodedaten
        Set dictMetadaten("wbDaten") = Workbooks(Dir(strMethodedaten))
        Set dictMetadaten("wsDaten") = dictMetadaten("wbDaten").Sheets(dictBatchdaten("Gerät"))
        
        ' Positionen der Werte definieren (-1 wenn nicht verlangt)
        intGeräteZeile = dictMetadaten("wbDaten").Sheets(1).Columns(4).Find(Environ("Computername")).Row
        With dictKolonnenposition
            ' Werte für Sequence
            For prpName = AcquisitionMethode To Wert4
                .Add funcGetPropertyName(prpName), dictMetadaten("wbDaten").Sheets("Geräte").Cells(intGeräteZeile, prpName).value
            Next prpName
            ' Werte für Informationen zur Messung
            For prpName = Messkategorie To Messkategorie
                .item(funcGetPropertyName(prpName)) = 0
            Next prpName
            'Werte für die ganze Sequence
            intSpalte = Wert4
            For prpName = Sequencename To Sequencename
                intSpalte = intSpalte + 1        ' Excel-Spalte direkt fortlaufend
                .Add funcGetPropertyName(prpName), _
                     dictMetadaten("wbDaten").Sheets("Geräte").Cells(intGeräteZeile, intSpalte).value
            Next prpName
        End With
        Set dictMetadaten("Kolonnenposition") = dictKolonnenposition
        intSpalte = dictMetadaten("wbDaten").Sheets("Geräte").Rows(2).Find("Exportordner").Column
        strExportordner = dictMetadaten("wbDaten").Sheets("Geräte").Cells(intGeräteZeile, intSpalte).value & "\"
        
        ' Methodenwerte auslesen
        With dictMethodedaten
            arrQuellKolonne = dictMetadaten("wsDaten").Range(dictMetadaten("wsDaten").Cells(2, 2), dictMetadaten("wsDaten").Cells(2, dictMetadaten("wsDaten").Columns.Count).End(xlToLeft))
            intMethodenZeile = dictMetadaten("wsDaten").Columns(Application.Match("Methode", arrQuellKolonne, 0) + 1).Find(What:=dictBatchdaten("Methode"), LookAt:=xlWhole).Row
            .Add "Kalibrationsanzahl", (dictMetadaten("wsDaten").Cells(intMethodenZeile, dictMetadaten("wsDaten").Columns.Count).End(xlToLeft).Column - (Application.Match("Lösungsmittel", arrQuellKolonne, 0) + 1)) / 3
            .Add "Spezialbrobenanzahl", WorksheetFunction.CountA(dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Spezialprobe 1 Probe 1 nach Kali", arrQuellKolonne, 0) + 1).Resize(1, 6)) / 2
            .Add "MethodeSTD100", dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Methodenname STD 100 (MUSS GENAU STIMMEN!)", arrQuellKolonne, 0) + 1)
            .Add "MethodeLeder", dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Methodenname Leder (MUSS GENAU STIMMEN!)", arrQuellKolonne, 0) + 1)
            .Add "MethodeECO", dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Methodenname Eco-Pass (MUSS GENAU STIMMEN!)", arrQuellKolonne, 0) + 1)
            .Add "MethodeKalibration", dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Methodenname Kalibration (MUSS GENAU STIMMEN!)", arrQuellKolonne, 0) + 1)
            .Add "Standardeinwaage", dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Standard-einwaage", arrQuellKolonne, 0) + 1)
            .Add "Exctraktionsvolumen", dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Exctraktionsvolumen", arrQuellKolonne, 0) + 1)
            .Add "Injektionsvolumen", dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Injektionsvolumen", arrQuellKolonne, 0) + 1)
            .Add "ProbenTyp", dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Proben Typ", arrQuellKolonne, 0) + 1)
            .Add "Rackname", dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Rackname", arrQuellKolonne, 0) + 1)
            .Add "RackMin", dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Rack Min", arrQuellKolonne, 0) + 1)
            .Add "RackMax", dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Rack Max", arrQuellKolonne, 0) + 1)
            .Add "RackPositionen", dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Maximale Anzahl Position", arrQuellKolonne, 0) + 1)
            .Add "ZwischenkaliEinzel_Volle", dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Einzel/Volle Zwischenkali", arrQuellKolonne, 0) + 1)
            .Add "ZwischenkaliQC_Cal", dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Zwischenkali als QC oder Cal", arrQuellKolonne, 0) + 1)
            .Add "BlankWechsel", dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Blank wechseln nach n Messungen", arrQuellKolonne, 0) + 1)
            .Add "KalWechsel", dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Kali wechseln nach n Messungen", arrQuellKolonne, 0) + 1)
            .Add "ZwischenBlankTrigger", dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Zwischenblank ab X Proben", arrQuellKolonne, 0) + 1)
            .Add "ZwischenKalibartionTrigger", dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Zwischenkali ab X Proben", arrQuellKolonne, 0) + 1)
            .Add "ZwischenKalibartionModus", IIf(dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Einzel/Volle Zwischenkali", arrQuellKolonne, 0) + 1) = "Einzel", 1, .item("Kalibrationsanzahl"))
        End With
        Set dictMetadaten("Methodedaten") = dictMethodedaten
        
        ' Blankwerte auslesen
        For prpName = AcquisitionMethode To Messkategorie
            If Not dictMetadaten("Kolonnenposition")(funcGetPropertyName(prpName)) = -1 Then _
               Call defSetWert(prp:=prpName, Messtyp:=Blank, Metadaten:=dictMetadaten, Blank:=objBlank)
        Next prpName
        
        ' Kalibrationwerte auslesen
        For i = 1 To dictMethodedaten("Kalibrationsanzahl")
            Set objKalibration = New CWerte
            colKalibration.Add objKalibration
            For prpName = AcquisitionMethode To Messkategorie
                If Not dictMetadaten("Kolonnenposition")(funcGetPropertyName(prpName)) = -1 Then _
                   Call defSetWert(prp:=prpName, Messtyp:=Kalibration, Metadaten:=dictMetadaten, Kalibration:=colKalibration, Collectionindex:=i)
            Next prpName
        Next i
         
        ' Spetialprobenwerte auslesen
        For i = 1 To 3
            If Not dictMetadaten("wsDaten").Cells(intMethodenZeile, (i - 1) * 2 + Application.Match("Spezialprobe 1 Probe 1 nach Kali", arrQuellKolonne, 0) + 1) = "" Then
                Set objSpezialproben = New CWerte
                colSpezialproben.Add objSpezialproben
                For prpName = AcquisitionMethode To Messkategorie
                    If Not dictMetadaten("Kolonnenposition")(funcGetPropertyName(prpName)) = -1 Then _
                       Call defSetWert(prp:=prpName, Messtyp:=Spezialprobe, Metadaten:=dictMetadaten, Spezialproben:=colSpezialproben, Kalibration:=colKalibration, Collectionindex:=i)
                Next prpName
            End If
        Next i
                
        ' Probenwerte auslesen
        For i = 1 To dictBatchdaten("TotalProben")
            Set objProben = New CWerte
            colProben.Add objProben
            For prpName = AcquisitionMethode To Messkategorie
                If Not dictMetadaten("Kolonnenposition")(funcGetPropertyName(prpName)) = -1 Then _
                   Call defSetWert(prp:=prpName, Messtyp:=Sample, Metadaten:=dictMetadaten, Probe:=colProben, Kalibration:=colKalibration, Collectionindex:=i)
            Next prpName
        Next i
        
        ' Trigger Definieren
        With dictTrigger
            .Add "MaxKalibration", Int(dictBatchdaten("TotalProben") / dictMethodedaten("ZwischenKalibartionTrigger"))
            .Add "AnzahlProbenZwischenKalibrationen", Int(dictBatchdaten("TotalProben") / (.item("MaxKalibration") + 1))
            .Add "AnzahlProbenZwischenBlank", .item("AnzahlProbenZwischenKalibrationen") \ ((.item("AnzahlProbenZwischenKalibrationen") \ dictMethodedaten("ZwischenBlankTrigger")) + 1)
        End With
        Set dictMetadaten("Trigger") = dictTrigger
        
        For prpName = Sequencename To Sequencename
            If Not dictMetadaten("Kolonnenposition")(funcGetPropertyName(prpName)) = -1 Then _
               Call defSetWert(prp:=prpName, Messtyp:=Ganzspalten, Metadaten:=dictMetadaten, Ganzspalten:=objGanzspalten)
        Next prpName
        
        ''' Sequence in Collection laden '''
        
        ' Anfangskalibration
        Call defInsertBlank(intZeileSequence, dictMetadaten, objBlank, colRawSequence)
        Call defInsertKalibration(intZeileSequence, dictMetadaten, colKalibration, True, False, colRawSequence)
        Call defInsertBlank(intZeileSequence, dictMetadaten, objBlank, colRawSequence)
        
        ' Spezialproben
        If Not colSpezialproben.Count = 0 Then
            Call defInsertSpezialproben(intZeileSequence, dictMetadaten, colSpezialproben, colRawSequence)
            Call defInsertBlank(intZeileSequence, dictMetadaten, objBlank, colRawSequence)
        End If
        
        For Each objProbe In colProben
            ' Probe
            intZeileSequence = intZeileSequence + 1
            colRawSequence.Add objProbe
            dictMetadaten("Trigger")("CurrentKalibrationTriggerCount") = dictMetadaten("Trigger")("CurrentKalibrationTriggerCount") + 1
            dictMetadaten("Trigger")("CurrentBlankTriggerCount") = dictMetadaten("Trigger")("CurrentBlankTriggerCount") + 1
            
            ' Zwischenkali
            If dictMetadaten("Trigger")("CurrentKalibrationTriggerCount") = dictMetadaten("Trigger")("AnzahlProbenZwischenKalibrationen") _
               And intAnzahlZwischenkalibration < dictTrigger("MaxKalibration") Then ' diese Zeile kappt unnötige Zwischenkalibrationen am Ende der Sequence
                intAnzahlZwischenkalibration = intAnzahlZwischenkalibration + 1
                Call defInsertBlank(intZeileSequence, dictMetadaten, objBlank, colRawSequence)
                Call defInsertKalibration(intZeileSequence, dictMetadaten, colKalibration, dictMethodedaten("ZwischenkaliEinzel_Volle") = "Volle", True, colRawSequence)
                Call defInsertBlank(intZeileSequence, dictMetadaten, objBlank, colRawSequence)
            End If
            
            ' Zwischenblank
            If dictMetadaten("Trigger")("CurrentBlankTriggerCount") = dictMetadaten("Trigger")("AnzahlProbenZwischenBlank") _
               And dictMetadaten("Trigger")("CurrentKalibrationTriggerCount") < dictMetadaten("Trigger")("AnzahlProbenZwischenKalibrationen") Then ' diese Zeile kappt unnötige Zwischenblanks vor den Zwischenkalibrationen
                Call defInsertBlank(intZeileSequence, dictMetadaten, objBlank, colRawSequence)
            End If
            
        Next objProbe
        ' Schlusskalibration
        Call defInsertBlank(intZeileSequence, dictMetadaten, objBlank, colRawSequence)
        Call defInsertKalibration(intZeileSequence, dictMetadaten, colKalibration, True, False, colRawSequence)
        Call defInsertBlank(intZeileSequence, dictMetadaten, objBlank, colRawSequence)
       
            
        ' Position anlegen
        Dim Kategorie As MessTypen
        Dim dictUsage As Object
        
        Set dictUsage = CreateObject("Scripting.Dictionary")
        
        intPosition = dictBatchdaten("Position")
        
        For Kategorie = Blank To Sample Step -1
            strPositionMeasage = strPositionMeasage + " - " & funcGetMesstypName(Kategorie) & " von " & intPosition
            dictUsage(Kategorie) = 0
            Select Case Kategorie
            Case Sample: defProcessKategorie colRawSequence, colFinalSequence, Sample, 1, intPosition, dictUsage
            Case Spezialprobe: ' TODO
            Case Zwischenkalibration: defProcessKategorie colRawSequence, colFinalSequence, Zwischenkalibration, dictMethodedaten("KalWechsel"), intPosition, dictUsage
            Case Kalibration: defProcessKategorie colRawSequence, colFinalSequence, Kalibration, dictMethodedaten("KalWechsel"), intPosition, dictUsage, UseLevel:=True
            Case Blank: defProcessKategorie colRawSequence, colFinalSequence, Blank, dictMethodedaten("BlankWechsel"), intPosition, dictUsage
            End Select
            strPositionMeasage = strPositionMeasage + " bis " & intPosition - 1 & Chr(10)
        Next Kategorie
        
        ' Sortiere die ganze Collection
        Call defSortCollectionByIndex(colFinalSequence)
                
        ''' Sequence schreiben '''
        With wsSequence
            .Visible = True
            .Cells.ClearContents
            intZeileSequence = 1
            For Each objMessung In colFinalSequence
                intZeileSequence = intZeileSequence + 1
                For prpName = AcquisitionMethode To Wert4
                    If Not dictMetadaten("Kolonnenposition")(funcGetPropertyName(prpName)) = -1 Then _
                       .Cells(intZeileSequence, dictMetadaten("Kolonnenposition")(funcGetPropertyName(prpName))) = funcGetWert(prp:=prpName, Messung:=objMessung)
                Next prpName
            Next objMessung
            
            ' Ganzbatchkolonnen
            For prpName = Sequencename To Sequencename
                If Not dictMetadaten("Kolonnenposition")(funcGetPropertyName(prpName)) = -1 Then _
                   .Range(.Cells(1, dictMetadaten("Kolonnenposition")(funcGetPropertyName(prpName))), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, dictMetadaten("Kolonnenposition")(funcGetPropertyName(prpName)))) = _
                   funcGetWert(prp:=prpName, Ganzspalten:=objGanzspalten)
            Next prpName
        
            ' Sequence in Clipboard überführen oder exportieren
            If strExportordner = "-1\" Then
                .UsedRange.Copy
                Workbooks("Book1").Sheets(1).Cells(1, 1).PasteSpecial xlPasteAll
            ElseIf Not funcIsFileOpen(strExportordner & dictMetadaten("Batchdaten")("Methode") & "_" & dictMetadaten("Batchdaten")("Topic") & ".csv") Then
                ActiveWorkbook.SaveAs filename:=strExportordner & dictMetadaten("Batchdaten")("Methode") & "_" & dictMetadaten("Batchdaten")("Topic"), FileFormat:=xlCSV, Local:=True
            Else
                MsgBox "Der Export ist noch geöffnet und kann daher nicht abgespeichert werden." & vbCrLf & "Bitte schliesse die Datei, bevor du erneut die Sequence exportierst.", Buttons:=vbExclamation + vbOKOnly, Title:="Exportfehler - Datei ist noch geöffnet"
            End If
            .Visible = False
        End With
    End If
    
    dictMetadaten("wbDaten").Close (False)
    
    Set arrQuellKolonne = Nothing
    Set rngProbenRange = Nothing
    Set dictMetadaten = Nothing
    Set dictBatchdaten = Nothing
    Set dictKolonnenposition = Nothing
    Set dictMethodedaten = Nothing
    Set objBlank = Nothing
    Set objKalibration = Nothing
    Set colKalibration = Nothing
    Set objSpezialproben = Nothing
    Set colSpezialproben = Nothing
    Set dictTrigger = Nothing
    Set objProben = Nothing
    Set colProben = Nothing
    Set colRawSequence = Nothing
    Set colFinalSequence = Nothing
    
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    'MsgBox strPositionMeasage, vbInformation, "Positionshilfe"
    Debug.Print strPositionMeasage
    
End Sub

Function funcCloneObject(orig As CWerte, Index As Integer) As CWerte

    Dim Clone As New CWerte
    Dim strPropertyname As String

    For prpName = AcquisitionMethode To Messkategorie
        strPropertyname = funcGetPropertyName(prpName)
        CallByName Clone, strPropertyname, VbLet, CallByName(orig, strPropertyname, VbGet)
    Next prpName
    Clone.Index = Index

    Set funcCloneObject = Clone

End Function

Private Function funcHasZwischenkalibration(colSequence) As Boolean

    Dim m1 As CWerte, m2 As CWerte, m3 As CWerte
    
    funcHasZwischenkalibration = False
    
    For i = 2 To colSequence.Count - 1
        Set m1 = colSequence(i - 1)
        Set m2 = colSequence(i)
        Set m3 = colSequence(i + 1)
        
        If m1.Messkategorie = Blank _
           And m2.Messkategorie = Kalibration _
           And m3.Messkategorie = Blank Then
           
           funcHasZwischenkalibration = True
           Exit Function
        End If
    Next i
    
End Function

Private Function funcGetMaxPosition(colFinalSequence) As Long
    Dim m As CWerte
    Dim maxPos As Long
    
    maxPos = 0
    For i = 1 To colFinalSequence.Count
        Set m = colFinalSequence(i)
        If m.Position > maxPos Then maxPos = m.Position
    Next i
    
    funcGetMaxPosition = maxPos
End Function


Private Sub defProcessKategorie(colSequence, colFinalSequence, KategorieConst As Long, maxUsage As Long, intPosition As Integer, dictUsage As Object, Optional UseLevel As Boolean = False)
    
    Dim Messung As CWerte
    Dim blnDoProcess As Boolean
    Dim blnIncreaseUsage As Boolean
    Dim blnDidSomething As Boolean
    Dim blnZwischenkalibrationIncreased As Boolean
    Dim Kategorie As MessTypen
    Dim intIncreaseAmount As Integer
    
    blnDidSomething = False
    blnZwischenkalibrationIncreased = dictUsage(Kalibration) = 0
    Kategorie = IIf(KategorieConst = Zwischenkalibration, Kalibration, KategorieConst)
    intIncreaseAmount = 0

    For i = 1 To colSequence.Count
        Set Messung = colSequence(i)
        If Messung.Messkategorie = Kategorie Then
        
            ' Bestimmen, ob Messung relevant ist oder nicht
            Select Case KategorieConst
            Case Kalibration
                blnDoProcess = Not (colSequence(i - 1).Messkategorie = Blank And colSequence(i + 1).Messkategorie = Blank)
                blnIncreaseUsage = colSequence(i + 1).Messkategorie = Blank
                If blnIncreaseUsage And blnDoProcess Then
                    intIncreaseAmount = Messung.Level - 1
                End If
            Case Zwischenkalibration
                blnDoProcess = colSequence(i - 1).Messkategorie = Blank And colSequence(i + 1).Messkategorie = Blank
                blnIncreaseUsage = True
            Case Else
                blnDoProcess = True
                blnIncreaseUsage = True
            End Select
            
            ' Messung in Sequence schreiben
            If blnDoProcess Then
                blnDidSomething = True
                Messung.Position = intPosition + IIf(UseLevel Or (Not blnZwischenkalibrationIncreased And KategorieConst = Zwischenkalibration), Messung.Level - 1, 0)
                colFinalSequence.Add funcCloneObject(Messung, i)
                
                ' Nutzung der Messung erhöhen
                If blnIncreaseUsage Then
                    dictUsage(Kategorie) = dictUsage(Kategorie) + 1
                    
                    ' nächste freie Position und nutzung zurücksetzen
                    If dictUsage(Kategorie) = maxUsage Then
                        If KategorieConst = Zwischenkalibration Then
                            If Not blnZwischenkalibrationIncreased Then
                                intPosition = funcGetMaxPosition(colFinalSequence) + 1
                                blnZwischenkalibrationIncreased = True
                            Else
                                intPosition = intPosition + 1
                            End If
                        Else
                            intPosition = intPosition + intIncreaseAmount + 1
                        End If
                        dictUsage(Kategorie) = 0
                    End If
                End If
            End If
        End If
    Next i
    
    ' Ende der Kategorie Position erhöhen, wenn nicht schon geschehen
    If KategorieConst = Kalibration Then
        If Not dictUsage(Kategorie) = 0 And Not funcHasZwischenkalibration(colSequence) Then intPosition = intPosition + intIncreaseAmount + 1
    Else
        If Not dictUsage(Kategorie) = 0 And blnDidSomething Then intPosition = funcGetMaxPosition(colFinalSequence) + 1
    End If
    
End Sub


Private Sub defInsertBlank(Row As Integer, Metadaten As Object, Blank As CWerte, Sequence As Collection)

    Row = Row + 1
    Sequence.Add Blank
    Metadaten("Trigger")("CurrentBlankTriggerCount") = 0
End Sub

Private Sub defInsertKalibration(Row As Integer, Metadaten As Object, Kalibration As Object, Volle_Kalibration As Boolean, Zwischenkalibration As Boolean, Sequence As Collection)

    Dim intZusatzprobeTrigger As Integer
    Dim blnExtra As Boolean

    If Not Metadaten("Batchdaten")("TotalProben") / Int(Metadaten("Batchdaten")("TotalProben") / Metadaten("Trigger")("AnzahlProbenZwischenKalibrationen")) - Metadaten("Trigger")("AnzahlProbenZwischenKalibrationen") = 0 Then
        ' Berechne, wie viele Proben zwischen zwei Zwischenkalibrationen liegen sollen
        intZusatzprobeTrigger = Int(1 / (Metadaten("Batchdaten")("TotalProben") / Int(Metadaten("Batchdaten")("TotalProben") / Metadaten("Trigger")("AnzahlProbenZwischenKalibrationen")) - Metadaten("Trigger")("AnzahlProbenZwischenKalibrationen")))

        ' Überprüfe, ob eine Extraprobe eingefügt werden soll
        blnExtra = Metadaten("Trigger")("MaxKalibration") Mod intZusatzprobeTrigger = 0
    End If
    
    If Volle_Kalibration Then
        For i = 1 To Kalibration.Count
            Row = Row + 1
            Sequence.Add Kalibration(i)
        Next i
    Else
        Row = Row + 1
        Sequence.Add Kalibration(Round(Kalibration.Count / 2, 0))
    End If
    
    Metadaten("Trigger")("CurrentKalibrationTriggerCount") = IIf(blnExtra = True, -1, 0)
End Sub

Private Sub defInsertSpezialproben(Row As Integer, Metadaten As Object, Spezialproben As Collection, Sequence As Collection)
    
    For i = 1 To Spezialproben.Count
        Row = Row + 1
        Sequence.Add Spezialprobe(i)
    Next i

End Sub

Private Sub defSetWert(prp As Properties, Messtyp As MessTypen, Metadaten As Object, _
                    Optional ByVal Probe As Collection = Nothing, _
                    Optional ByVal Kalibration As Collection = Nothing, _
                    Optional ByVal Blank As Object = Nothing, _
                    Optional ByVal Spezialproben As Collection = Nothing, _
                    Optional ByVal Ganzspalten As Object = Nothing, _
                    Optional ByVal Collectionindex As Integer = -1)
    
    Dim intMethodenZeile As Integer
    Dim arrQuellKolonne As Variant
    
    With Metadaten("wsDaten")
        arrQuellKolonne = .Range(.Cells(2, 2), .Cells(2, Columns.Count).End(xlToLeft))
        intMethodenZeile = .Columns(Application.Match("Methode", arrQuellKolonne, 0) + 1).Find(What:=Metadaten("Batchdaten")("Methode"), LookAt:=xlWhole).Row
        ' Definieren Sie separate Variablen für verschiedene Messarten
        ' Wert für Sample
        If Messtyp = 0 Then
            Select Case prp
            Case AcquisitionMethode: Probe(Collectionindex).AcquisitionMethode = funcGetMethode(wsHauptseite.Cells(Collectionindex + 2, 5), Metadaten)
            Case Quantmethode: Probe(Collectionindex).Quantmethode = .Cells(intMethodenZeile, Application.Match("Quantmethode", arrQuellKolonne, 0) + 1)
            Case Beschriftung: Probe(Collectionindex).Beschriftung = wsHauptseite.Cells(Collectionindex + 2, 2)
            Case Einwaage: Probe(Collectionindex).Einwaage = wsHauptseite.Cells(Collectionindex + 2, 3)
            Case Exctraktionsvolumen: Probe(Collectionindex).Exctraktionsvolumen = .Cells(intMethodenZeile, Application.Match("Exctraktionsvolumen", arrQuellKolonne, 0) + 1)
            Case Injektionsvolumen: Probe(Collectionindex).Injektionsvolumen = .Cells(intMethodenZeile, Application.Match("Injektionsvolumen", arrQuellKolonne, 0) + 1)
            Case Kommentar: Probe(Collectionindex).Kommentar = wsHauptseite.Cells(Collectionindex + 2, 6)
            Case Rack: Probe(Collectionindex).Rack = "Rack" 'Fehlt!
            Case Position: Probe(Collectionindex).Position = IIf(Collectionindex = 1, _
                                                                 Metadaten("Methodedaten")("Spezialbrobenanzahl") + Kalibration(Kalibration.Count).Position + 1, _
                                                                 funcGetPosition(Probe:=Probe, Collectionindex:=Collectionindex, Metadaten:=Metadaten))
            Case Produktklasse: Probe(Collectionindex).Produktklasse = wsHauptseite.Cells(Collectionindex + 2, 5)
            Case Typ: Probe(Collectionindex).Typ = .Cells(intMethodenZeile, Application.Match("Proben Typ", arrQuellKolonne, 0) + 1)
            Case Konzentration: Probe(Collectionindex).Konzentration = 0
            Case Verdünnung: Probe(Collectionindex).Verdünnung = wsHauptseite.Cells(Collectionindex + 2, 4)
            Case Level: Probe(Collectionindex).Level = 0
            Case Info1: Probe(Collectionindex).Info1 = "Sample amount mg or uL"
            Case Info2: Probe(Collectionindex).Info2 = ""
            Case Info3: Probe(Collectionindex).Info3 = ""
            Case Info4: Probe(Collectionindex).Info4 = ""
            Case Wert1: Probe(Collectionindex).Wert1 = 0
            Case Wert2: Probe(Collectionindex).Wert2 = 0
            Case Wert3: Probe(Collectionindex).Wert3 = 0
            Case Wert4: Probe(Collectionindex).Wert4 = 0
            Case Messkategorie: Probe(Collectionindex).Messkategorie = MessTypen.Sample
'            Case Nutzungen: Probe(Collectionindex).Nutzungen = 0
            Case Else: GoTo ErrHandler
            End Select
            
            ' Wert für Kalibration
        ElseIf Messtyp = 3 Then
            Select Case prp
            Case AcquisitionMethode: Kalibration(Collectionindex).AcquisitionMethode = funcGetMethode("CALIBRATION", Metadaten)
            Case Quantmethode: Kalibration(Collectionindex).Quantmethode = .Cells(intMethodenZeile, Application.Match("Quantmethode", arrQuellKolonne, 0) + 1)
            Case Beschriftung: Kalibration(Collectionindex).Beschriftung = .Cells(intMethodenZeile, Application.Match("Kalibration Level " & Collectionindex, arrQuellKolonne, 0) + 1)
            Case Einwaage: Kalibration(Collectionindex).Einwaage = .Cells(intMethodenZeile, Application.Match("Standard-Einwaage", arrQuellKolonne, 0) + 1)
            Case Exctraktionsvolumen: Kalibration(Collectionindex).Exctraktionsvolumen = .Cells(intMethodenZeile, Application.Match("Exctraktionsvolumen", arrQuellKolonne, 0) + 1)
            Case Injektionsvolumen: Kalibration(Collectionindex).Injektionsvolumen = .Cells(intMethodenZeile, Application.Match("Injektionsvolumen", arrQuellKolonne, 0) + 1)
            Case Kommentar: Kalibration(Collectionindex).Kommentar = ""
            Case Konzentration: Kalibration(Collectionindex).Konzentration = 18 'Fehlt!
            Case Position: Kalibration(Collectionindex).Position = .Cells(intMethodenZeile, Application.Match("Position Kalibration Level " & Collectionindex & " (n Positionen nach Lösungsmittel)", arrQuellKolonne, 0) + 1) + 1
            Case Produktklasse: Kalibration(Collectionindex).Produktklasse = ""
            Case Rack: Kalibration(Collectionindex).Rack = "Rack" 'Fehlt!
            Case Typ: Kalibration(Collectionindex).Typ = .Cells(intMethodenZeile, 3 * i + Application.Match("Lösungsmittel", arrQuellKolonne, 0) + 2)
            Case Verdünnung: Kalibration(Collectionindex).Verdünnung = 1
            Case Level: Kalibration(Collectionindex).Level = Collectionindex
            Case Info1: Kalibration(Collectionindex).Info1 = "Sample amount mg or uL"
            Case Info2: Kalibration(Collectionindex).Info2 = ""
            Case Info3: Kalibration(Collectionindex).Info3 = ""
            Case Info4: Kalibration(Collectionindex).Info4 = ""
            Case Wert1: Kalibration(Collectionindex).Wert1 = 0
            Case Wert2: Kalibration(Collectionindex).Wert2 = 0
            Case Wert3: Kalibration(Collectionindex).Wert3 = 0
            Case Wert4: Kalibration(Collectionindex).Wert4 = 0
            Case Messkategorie: Kalibration(Collectionindex).Messkategorie = MessTypen.Kalibration
'            Case Nutzungen: Kalibration(Collectionindex).Nutzungen = 0
            Case Else: GoTo ErrHandler
            End Select
            ' Wert für Blank
        ElseIf Messtyp = 4 Then
            Select Case prp
            Case AcquisitionMethode: Blank.AcquisitionMethode = funcGetMethode("CALIBRATION", Metadaten)
            Case Quantmethode: Blank.Quantmethode = .Cells(intMethodenZeile, Application.Match("Quantmethode", arrQuellKolonne, 0) + 1)
            Case Beschriftung: Blank.Beschriftung = .Cells(intMethodenZeile, Application.Match("Lösungsmittel", arrQuellKolonne, 0) + 1)
            Case Einwaage: Blank.Einwaage = .Cells(intMethodenZeile, Application.Match("Standard-Einwaage", arrQuellKolonne, 0) + 1)
            Case Exctraktionsvolumen: Blank.Exctraktionsvolumen = .Cells(intMethodenZeile, Application.Match("Exctraktionsvolumen", arrQuellKolonne, 0) + 1)
            Case Injektionsvolumen: Blank.Injektionsvolumen = .Cells(intMethodenZeile, Application.Match("Injektionsvolumen", arrQuellKolonne, 0) + 1)
            Case Kommentar: Blank.Kommentar = ""
            Case Rack: Blank.Rack = "Rack"
            Case Position: Blank.Position = Metadaten("Batchdaten")("Position")
            Case Produktklasse: Blank.Produktklasse = ""
            Case Typ: Blank.Typ = "Blank"
            Case Konzentration: Blank.Konzentration = 0
            Case Verdünnung: Blank.Verdünnung = 1
            Case Level: Blank.Level = 0
            Case Info1: Blank.Info1 = "Sample amount mg or uL"
            Case Info2: Blank.Info2 = ""
            Case Info3: Blank.Info3 = ""
            Case Info4: Blank.Info4 = ""
            Case Wert1: Blank.Wert1 = 0
            Case Wert2: Blank.Wert2 = 0
            Case Wert3: Blank.Wert3 = 0
            Case Wert4: Blank.Wert4 = 0
            Case Messkategorie: Blank.Messkategorie = MessTypen.Blank
'            Case Nutzungen: Blank.Nutzungen = 0
            Case Else: GoTo ErrHandler
            End Select
            
            ' Wert für Spezialprobe
        ElseIf Messtyp = 1 Then
            Select Case prp
            Case AcquisitionMethode: Spezialproben(Collectionindex).AcquisitionMethode = funcGetMethode(Metadaten("Batchdaten")("Topic"), Metadaten)
            Case Quantmethode: Blank.Quantmethode = .Cells(intMethodenZeile, Application.Match("Quantmethode", arrQuellKolonne, 0) + 1)
            Case Beschriftung: Spezialproben(Collectionindex).Beschriftung = .Cells(intMethodenZeile, (i - 1) * 2 + Application.Match("Spezialprobe 1 Probe 1 nach Kali", arrQuellKolonne, 0) + 1)
            Case Einwaage: Spezialproben(Collectionindex).Einwaage = .Cells(intMethodenZeile, Application.Match("Standard-Einwaage", arrQuellKolonne, 0) + 1)
            Case Exctraktionsvolumen: Spezialproben(Collectionindex).Exctraktionsvolumen = .Cells(intMethodenZeile, Application.Match("Exctraktionsvolumen", arrQuellKolonne, 0) + 1)
            Case Injektionsvolumen: Spezialproben(Collectionindex).Injektionsvolumen = .Cells(intMethodenZeile, Application.Match("Injektionsvolumen", arrQuellKolonne, 0) + 1)
            Case Kommentar: Spezialproben(Collectionindex).Kommentar = ""
            Case Rack: Spezialproben(Collectionindex).Rack = 38 'Fehlt
            Case Position: Spezialproben(Collectionindex).Position = Kalibration(Kalibration.Count).Position + i
            Case Produktklasse: Spezialproben(Collectionindex).Produktklasse = ""
            Case Typ: Spezialproben(Collectionindex).Typ = .Cells(intMethodenZeile, (i - 1) * 2 + Application.Match("Type für Spezialprobe 1", arrQuellKolonne, 0) + 1)
            Case Konzentration: Spezialproben(Collectionindex).Konzentration = 0 'Fehlt!
            Case Verdünnung: Spezialproben(Collectionindex).Verdünnung = 1
            Case Level: Kalibration(Collectionindex).Level = Null
            Case Info1: Spezialproben(Collectionindex).Info1 = "Sample amount mg or uL"
            Case Info2: Spezialproben(Collectionindex).Info2 = ""
            Case Info3: Spezialproben(Collectionindex).Info3 = ""
            Case Info4: Spezialproben(Collectionindex).Info4 = ""
            Case Wert1: Spezialproben(Collectionindex).Wert1 = 0
            Case Wert2: Spezialproben(Collectionindex).Wert2 = 0
            Case Wert3: Spezialproben(Collectionindex).Wert3 = 0
            Case Wert4: Spezialproben(Collectionindex).Wert4 = 0
            Case Messkategorie: Spezialproben(Collectionindex).Messkategorie = MessTypen.Spezialprobe
'            Case Nutzungen: Spezialproben(Collectionindex).Nutzungen = 0
            Case Else: GoTo ErrHandler
            End Select
            
            ' Wert für Ganzspalten
        ElseIf Messtyp = 5 Then
            Select Case prp
            Case Sequencename: Ganzspalten.Sequencename = Format(Now(), "yymmdd") & "_" & Metadaten("Batchdaten")("Operator") & "_" & funcGetMethode(Metadaten("Batchdaten")("Topic"), Metadaten)
            Case Else: GoTo ErrHandler
            End Select
        End If
    End With
    
    Exit Sub
    
ErrHandler:
    ActiveWorkbook.Close savechanges:=False
    MsgBox "Es gab ein Fehler beim Implementieren eines wertes." & vbCr & "Bitte melde Dich beim Digital Laboratory Expert.", vbCritical, "Fehlender Wert"
    End

End Sub

Private Function funcGetPosition(Probe As Collection, Collectionindex As Integer, Metadaten As Object) As Integer
    If Collectionindex > 1 Then funcGetPosition = Probe(Collectionindex - 1).Position + 1
    If funcGetPosition > Metadaten("Methodedaten")("RackPositionen") Then funcGetPosition = 1
End Function

Private Function funcGetWert(prp As Properties, _
                         Optional ByVal Messung As CWerte = Nothing, _
                         Optional ByVal Ganzspalten As Object = Nothing) As Variant

    Dim obj As Object
    Dim varValue As Variant
    
    ' Bestimmen Sie das entsprechende Objekt basierend auf dem MessTyp
    Set obj = IIf(Ganzspalten Is Nothing, Messung, Ganzspalten)
    
    ' Überprüfen Sie, ob das Objekt gültig ist
    If Not obj Is Nothing Then
        ' Prüfen Sie, ob die angeforderte Eigenschaft vorhanden ist und weisen Sie ihren Wert zu
        Select Case prp
        Case AcquisitionMethode: varValue = obj.AcquisitionMethode
        Case Quantmethode: varValue = obj.Quantmethode
        Case Beschriftung: varValue = obj.Beschriftung
        Case Einwaage: varValue = obj.Einwaage
        Case Exctraktionsvolumen: varValue = obj.Exctraktionsvolumen
        Case Injektionsvolumen: varValue = obj.Injektionsvolumen
        Case Kommentar: varValue = obj.Kommentar
        Case Konzentration: varValue = obj.Konzentration
        Case Position: varValue = obj.Position
        Case Produktklasse: varValue = obj.Produktklasse
        Case Rack: varValue = obj.Rack
        Case Typ: varValue = obj.Typ
        Case Verdünnung: varValue = obj.Verdünnung
        Case Level: varValue = obj.Level
        Case Sequencename: varValue = obj.Sequencename
        Case Info1: varValue = obj.Info1
        Case Info2: varValue = obj.Info2
        Case Info3: varValue = obj.Info3
        Case Info4: varValue = obj.Info4
        Case Wert1: varValue = obj.Wert1
        Case Wert2: varValue = obj.Wert2
        Case Wert3: varValue = obj.Wert3
        Case Wert4: varValue = obj.Wert4
        Case Messkategorie: varValue = obj.Messkategorie
        Case Else:                               ' Aktion für unbekannte Eigenschaft
        End Select

    Else
        varValue = "Unknown"
    End If

    funcGetWert = varValue
    
End Function

Private Function funcGetPropertyName(prp As Properties) As String
    Select Case prp
    Case AcquisitionMethode: funcGetPropertyName = "AcquisitionMethode"
    Case Quantmethode: funcGetPropertyName = "Quantmethode"
    Case Beschriftung: funcGetPropertyName = "Beschriftung"
    Case Einwaage: funcGetPropertyName = "Einwaage"
    Case Exctraktionsvolumen: funcGetPropertyName = "Exctraktionsvolumen"
    Case Injektionsvolumen: funcGetPropertyName = "Injektionsvolumen"
    Case Kommentar: funcGetPropertyName = "Kommentar"
    Case Konzentration: funcGetPropertyName = "Konzentration"
    Case Position: funcGetPropertyName = "Position"
    Case Produktklasse: funcGetPropertyName = "Produktklasse"
    Case Rack: funcGetPropertyName = "Rack"
    Case Typ: funcGetPropertyName = "Typ"
    Case Verdünnung: funcGetPropertyName = "Verdünnung"
    Case Level: funcGetPropertyName = "Level"
    Case Info1: funcGetPropertyName = "Info1"
    Case Info2: funcGetPropertyName = "Info2"
    Case Info3: funcGetPropertyName = "Info3"
    Case Info4: funcGetPropertyName = "Info4"
    Case Wert1: funcGetPropertyName = "Wert1"
    Case Wert2: funcGetPropertyName = "Wert2"
    Case Wert3: funcGetPropertyName = "Wert3"
    Case Wert4: funcGetPropertyName = "Wert4"
    Case Messkategorie: funcGetPropertyName = "Messkategorie"
'    Case Nutzungen: defGetPropertyName = "Nutzungen"
    Case Sequencename: funcGetPropertyName = "Sequencename"
    Case Else: funcGetPropertyName = "Unknown"
    End Select
End Function

Private Function funcGetMesstypName(prp As MessTypen) As String
    Select Case prp
    Case Sample: funcGetMesstypName = "Sample"
    Case Zwischenkalibration: funcGetMesstypName = "Zwischenkalibration"
    Case Kalibration: funcGetMesstypName = "Kalibration"
    Case Blank: funcGetMesstypName = "Blank"
    Case Spezialprobe: funcGetMesstypName = "Spezialprobe"
    Case Ganzspalten: funcGetMesstypName = "Ganzspalten"
    Case Else: funcGetMesstypName = "Unknown"
    End Select
End Function

' Funktion zum Abrufen der Messmethode
Private Function funcGetMethode(strTopic As Variant, Metadaten As Object) As String
    
    Select Case Left(strTopic, 3)
    Case "STD", "STA": funcGetMethode = Metadaten("Methodedaten")("MethodeSTD100")
    Case "L", "LEA": funcGetMethode = Metadaten("Methodedaten")("MethodeLeder")
    Case "ECP", "ECO": funcGetMethode = Metadaten("Methodedaten")("MethodeECO")
    Case "CAL", "CAL": funcGetMethode = Metadaten("Methodedaten")("MethodeKalibration")
    Case Else                                    ' Aktion für unbekannte Eigenschaft
    End Select

End Function

Sub Ausdruck()

    Dim strGC As String, strMethode As String, strOperator As String, strQuellOrdner As String, EndOrdner As String

    Application.EnableEvents = False: Application.DisplayAlerts = False: Application.ScreenUpdating = False

    Cells(24, 8) = "MM"
    Call Sequence
    strGC = Cells(2, 8)
    strMethode = Cells(3, 8)
    strOperator = Cells(4, 8)
    Sheets("Ausdruck").Visible = True
    Sheets("Sequence").Visible = True
    Sheets("Ausdruck").Activate
    Sheets("Ausdruck").Unprotect
    With Sheets("Sequence")
        Range(Cells(Columns(2).Find("Pos.:").Row + 1, 2), Cells(Rows.Count, 6)).ClearContents
        Cells(2, 2) = strGC
        Cells(4, 4) = strMethode
        Cells(5, 4) = strOperator
        Cells(6, 4) = Date
        i = .Cells(Rows.Count, 1).End(xlUp).Row
        .Range(.Cells(1, 1), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 1)).Copy
        Cells(Columns(2).Find("Pos.:").Row + 1, 2).PasteSpecial xlPasteValues
        .Range(.Cells(1, 3), .Cells(.Cells(Rows.Count, 3).End(xlUp).Row, 3)).Copy
        Cells(Columns(2).Find("Pos.:").Row + 1, 3).PasteSpecial xlPasteValues
        .Range(.Cells(1, 7), .Cells(.Cells(Rows.Count, 7).End(xlUp).Row, 7)).Copy
        Cells(Columns(2).Find("Pos.:").Row + 1, 4).PasteSpecial xlPasteValues
        .Range(.Cells(1, 10), .Cells(.Cells(Rows.Count, 10).End(xlUp).Row, 10)).Copy
        Cells(Columns(2).Find("Pos.:").Row + 1, 5).PasteSpecial xlPasteValues
        .Range(.Cells(1, 6), .Cells(.Cells(Rows.Count, 6).End(xlUp).Row, 6)).Copy
        Cells(Columns(2).Find("Pos.:").Row + 1, 6).PasteSpecial xlPasteValues
    End With
    With Range(Cells(Columns(2).Find("Pos.:").Row + 1, 2), Cells(Cells(Rows.Count, 6).End(xlUp).Row, 6))
        .Font.Size = 12
        .Font.Bold = True
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlInsideVertical).Weight = xlThin
        .NumberFormat = "0.000"
        .HorizontalAlignment = xlLeft
    End With
    Sheets("Ausdruck").Protect
    Sheets("Ausdruck").Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs "L:\Makros\Zwischenspeicher\Sequence Zwischenspeicher\" & Format(Date, "YYMMdd") & "_" & strMethode & "_" & strGC & "_" & strOperator & ".xlsx"
    Application.DisplayAlerts = True
    ActiveWorkbook.Close (False)
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
    Cells(24, 8) = Format(Date, "YYMMdd") & "_" & strMethode & "_" & strGC & "_" & strOperator & " " & Format(WorksheetFunction.WorkDay(Date, 2), "dd.MM.YYYY")
    Cells(24, 8).Copy
    MsgBox "Export wurde Ausgeführt.", vbInformation, "Done"

    Application.EnableEvents = True: Application.DisplayAlerts = True: Application.ScreenUpdating = True

End Sub

Sub Kill()

    Application.EnableEvents = False: Application.DisplayAlerts = False: Application.ScreenUpdating = False

    Range(Cells(3, 2), Cells(432, 6)).ClearContents
    Cells(3, 9) = "Methode"
    Cells(3, 10) = "Std100"
    Range(Cells(4, 9), Cells(5, 9)).ClearContents
    Cells(5, 10) = 1
    Cells(9, 11) = Date
    'Sheets("Hauptseite").Unprotect
    'Range(Columns(13), Columns(15)).ClearContents
    'Sheets("Hauptseite").Protect
    'With Sheets("Ausdruck")
    '    .Visible = True
    '    .Unprotect
    '    .Range(.Cells(2, 2), .Cells(2, 6)).ClearContents
    '    .Range(.Cells(4, 4), .Cells(7, 4)).ClearContents
    '    Do Until .Cells(9, 2) = "Pos.:"
    '        .Rows(8).Delete
    '    Loop
    '    If Not .Cells(.Columns(2).Find("Pos.:").Row + 1, 2) = "" Then
    '        .Range(.Cells(.Columns(2).Find("Pos.:").Row + 1, 2), .Cells(.Cells(Rows.Count, 6).End(xlUp).Row, 6)).Borders(xlDiagonalDown).LineStyle = xlNone
    '        .Range(.Cells(.Columns(2).Find("Pos.:").Row + 1, 2), .Cells(.Cells(Rows.Count, 6).End(xlUp).Row, 6)).Borders(xlEdgeLeft).LineStyle = xlNone
    '        .Range(.Cells(.Columns(2).Find("Pos.:").Row + 1, 2), .Cells(.Cells(Rows.Count, 6).End(xlUp).Row, 6)).Borders(xlEdgeBottom).LineStyle = xlNone
    '        .Range(.Cells(.Columns(2).Find("Pos.:").Row + 1, 2), .Cells(.Cells(Rows.Count, 6).End(xlUp).Row, 6)).Borders(xlEdgeRight).LineStyle = xlNone
    '        .Range(.Cells(.Columns(2).Find("Pos.:").Row + 1, 2), .Cells(.Cells(Rows.Count, 6).End(xlUp).Row, 6)).Borders(xlInsideVertical).LineStyle = xlNone
    '        .Range(.Cells(.Columns(2).Find("Pos.:").Row + 1, 2), .Cells(.Cells(Rows.Count, 6).End(xlUp).Row, 6)).ClearContents
    '    End If
    '    .Protect
    '    .Visible = False
    'End With
    With Sheets("Sequence")
        .Visible = True
        .Cells.ClearContents
        .Visible = False
    End With

    Application.EnableEvents = True: Application.DisplayAlerts = True: Application.ScreenUpdating = True

End Sub
