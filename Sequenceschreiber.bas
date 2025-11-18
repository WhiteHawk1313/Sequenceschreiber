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
    Sequencename = 27                            ' ab hier Properties, die erst am Schluss der Sequence eingefügt werden
    Exportordner = 28                            ' Spezialeintrag, nie im einem Loop
End Enum

Enum MessTypen
    Sample = 0
    Kalibration = 1
    Blank = 2
    Spezialprobe = 3
    Ganzspalten = 4
End Enum

Private Sub Import()
    ' Strings
    Dim strGerät As String, strMethode As String, strTopic As String, strPfad As String, strDatum As String, strDatei As String
    Dim strVolleDatei As String, strOperatorTest As String, strName As String
    Const strMethodedaten As String = "L:\Makros\Sequenceschreiber\Daten für Sequenceschreiber.xlsx"
    
    ' Workbooks und Worksheets
    Dim DatenWB As Workbook, ZWB As Workbook, QWB As Workbook
    Dim ZWS As Worksheet
    
    ' Ranges
    Dim rngZelle As Range
    
    ' Arrays
    Dim arrQuellKolonne As Variant, arrExporte() As String, arrEinzelEinwaagen As Variant
    
    ' Integers und Doubles
    Dim intMethodenZeile As Integer, i As Integer, intOperatorCount As Integer, j As Integer, intZeile As Integer
    Dim dblStdEinwaage As Double, dblSumme As Double

    
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
            strTopic = strMethode & IIf(.Cells(3, 2) = "Std100", "", "-" & .Cells(3, 2)) & "_"
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
        If IsArrayEmpty(arrExporte) = True Then
            MsgBox "Keine Daten für das Importieren gefunden." & vbCr & "Vergewissere dich bitte, ob ein Batch für diese Methode und Datum existiert und ob dieser den Status Action hat." & vbCr & "Bei Fragen wende dich bitte an den Digital Laboratory Expert." & vbCr & "Danke.", vbCritical, "Keine Daten gefunden."
            wsHauptseite.Protect
            GoTo SaveExit
            End
        End If
        
        ' Sortieren der Sequencen
        defQuickSortString arrExporte, LBound(arrExporte), UBound(arrExporte)
        
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

Private Sub defQuickSortString(arr() As String, ByVal low As Long, ByVal high As Long)

    Dim pivot As String
    Dim tempSwap As String
    Dim i As Long
    Dim j As Long
    
    On Error GoTo errHandler
    
    If low < high Then
        pivot = arr((low + high) \ 2)
        i = low - 1
        j = high + 1
        
        Do
            Do
                i = i + 1
            Loop While CLng(Split(arr(i), "_")(3)) < CLng(Split(pivot, "_")(3))
            
            Do
                j = j - 1
            Loop While CLng(Split(arr(j), "_")(3)) > CLng(Split(pivot, "_")(3))
            
            If i < j Then
                ' Tausche die Elemente
                tempSwap = arr(i)
                arr(i) = arr(j)
                arr(j) = tempSwap
            End If
        Loop While i < j
        
        defQuickSortString arr, low, j
        defQuickSortString arr, j + 1, high
    End If
    
    On Error GoTo 0
    
    Exit Sub
    
errHandler:
    MsgBox "Beim sortieren der Batches ist ein Fehler aufgeten." & vbCr & "Wende dich bitte an den Digital Laboratory Expert." & vbCr & "Danke.", vbCritical, "Fehler beim Sortieren"

    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    End
    
End Sub

Private Function IsArrayEmpty(arr As Variant) As Boolean
    IsArrayEmpty = True
    On Error Resume Next
    IsArrayEmpty = (LBound(arr) > UBound(arr))
    On Error GoTo 0
    
End Function

Private Function IsOperatorPresent(arr As Variant, strName As String) As Boolean
    On Error Resume Next
    For i = LBound(arr) To UBound(arr)
        If arr(i) Like "*" & strName & "*" Then
            IsOperatorPresent = True
            Exit Function
        End If
    Next i
    On Error GoTo 0
End Function

Function IsFileOpen(filename As String) As Boolean
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
        IsFileOpen = False
    Else
        IsFileOpen = True
    End If
End Function

'Private Sub Import_old()
'
'Dim arrExporte() As String, arrExporteSortiert() As String, arrExportName() As String
'Dim strGC As String, strMethode As String, strTopic As String, strPfad As String, strDatum As String
'Dim strDatei As String, strVolleDatei As String, strOperatorTest As String, strName As String, strShortName As String
'Dim intLaufnummerTest As Integer, intLaufnummemSortiert As Integer, intMethodenZeile As Integer
'Dim intCOperator As Integer, intZeileInBatch As Integer, intQWBC As Integer, intZeile As Integer, intDoppelbestimmung As Integer
'Dim DatenWB As Workbook, ZWB As Workbook, QWB As Workbook
'Dim ZWS As Worksheet, QWS As Worksheet
'Dim rngZelle As Range, rngSample As Range
'Dim dblSumme As Variant, dblStdEinwaage As Double
'Dim arrQuellKolonne As Variant, varEinzelEinwaagen As Variant, varProbennameTeile As Variant
'
'ReDim arrExporteSortiert(0)
'ReDim Preserve arrExporte(0)
'
'Application.EnableEvents = False: Application.DisplayAlerts = False: Application.ScreenUpdating = False
'
'If Cells(3, 8) = "Methode" Then
'    MsgBox ("Bitte Methode wählen. Danke.")
'    End
'Else
'    strGC = Cells(2, 8)
'    strMethode = Cells(3, 8)
'    strTopic = Cells(3, 9) & "_"
'    Workbooks.Open "L:\Makros\Sequenceschreiber\GC\Daten für GC Sequenceschreiber.xlsx"
'    Set DatenWB = Workbooks("Daten für GC Sequenceschreiber.xlsx")
'    ThisWorkbook.Activate
'    With DatenWB.Sheets(strGC)
'        arrQuellKolonne = .Range(.Cells(2, 2), .Cells(2, Columns.Count).End(xlToLeft))
'        intMethodenZeile = .Columns(Application.Match("Methodenname Kalibration*(MUSS GENAU STIMMEN!)", arrQuellKolonne, 0) + 1).Find(strMethode).Row
'        dblStdEinwaage = .Cells(intMethodenZeile, Application.Match("Standard-*einwaage", arrQuellKolonne, 0) + 1)
'    End With
'    DatenWB.Close (False)
'    strPfad = "L:\UnilabUltimateBatches\ZH_Equipment\"
'    strDatum = "ZH_" & Format(Cells(9, 10), "yyyyMMdd") & "_"
'    strDatei = "*.xlsx"
'    strVolleDatei = Dir(strPfad & strDatum & strTopic & strDatei)
'    strOperatorTest = "*" & Right(strVolleDatei, 29) & "*"
'
'    Do Until strVolleDatei = ""
'        If Not strVolleDatei Like strOperatorTest Then intCOperator = intCOperator + 1
'        ReDim Preserve arrExporte(i)
'        ReDim arrExporteSortiert(i)
'        arrExporte(i) = strVolleDatei
'        i = i + 1
'        strVolleDatei = Dir
'    Loop
'    If intCOperator > 0 Then strName = InputBox("Unter dem heutigem Datum sind Dateien von verschiedenen Personen vorhanden. Bitte dein Kürzel eintragen und auf ""OK"" klicken.")
'
'    For i = 0 To UBound(arrExporte)
'        intLaufnummerTest = 1313
'        intLaufnummemSortiert = 1313
'        For j = 0 To UBound(arrExporte)
'            If arrExporte(j) Like "*" & strName & "*.xlsx" Then
'                arrExportName() = Split(arrExporte(j), "_")
'                If arrExportName(UBound(arrExportName) - 2) * 1 < intLaufnummerTest And IsError(Application.Match(arrExporte(j), arrExporteSortiert, 0)) Then
'                    intLaufnummerTest = IIf(IsNumeric(arrExportName(3)), arrExportName(3), 1312)
'                    intLaufnummemSortiert = j
'                End If
'            End If
'        Next
'        If intLaufnummemSortiert = 1313 Then
'            ReDim Preserve arrExporteSortiert(IIf(i - 1 < 0, 0, i - 1))
'            Exit For
'        Else: arrExporteSortiert(i) = arrExporte(intLaufnummemSortiert)
'        End If
'    Next
'
'    If arrExporteSortiert(0) = "" Then
'        MsgBox ("Keine Daten für das Importieren gefunden. Bitte MM oder ein Teamleiter kontaktieren. Danke.")
'        Sheets(1).Protect
'        End
'    End If
'    Sheets("Hauptseite").Unprotect
'    For i = 0 To UBound(arrExporteSortiert)
'        Set ZWB = ActiveWorkbook
'        Set ZWS = ActiveSheet
'        Workbooks.Open strPfad & arrExporteSortiert(i)
'        Set QWB = ActiveWorkbook
'
'        With ZWB.Sheets("Ausdruck")
'            .Visible = True
'            .Unprotect
'            If intQWBC > 0 Then .Rows(intQWBC + 7).Insert Shift:=xlDown
'            .Cells(intQWBC + 7, 4) = QWB.Name
'            intQWBC = intQWBC + 1
'            .Protect
'            .Visible = False
'        End With
'
'        strShortName = Left(Right(QWB.Name, IIf(Right(QWB.Name, 29) Like "_*", 28, 29)), IIf(Right(QWB.Name, 29) Like "_*", 2, 3))
'        Set QWS = QWB.ActiveSheet
'            QWS.Copy after:=ZWS
'            QWB.Close (False)
'        Application.DisplayAlerts = False
'
'        With ThisWorkbook.Sheets("Hauptseite")
'            Range(Cells(1, 1), Cells(Cells(Rows.Count, 1).End(xlUp).Row, 1)).Copy Destination:=.Cells(.Cells(Rows.Count, 2).End(xlUp).Row + 1, 2)
'            Range(Cells(1, 5), Cells(Cells(Rows.Count, 5).End(xlUp).Row, 5)).Copy Destination:=.Cells(.Cells(Rows.Count, 5).End(xlUp).Row + 1, 5)
'
'            Range(Cells(1, 2), Cells(Cells(Rows.Count, 2).End(xlUp).Row, 2)).Replace What:=",", Replacement:=".", LookAt:=xlPart '",." Einwaagekorrektur
'            For Each rngZelle In Range(Cells(1, 2), Cells(Cells(Rows.Count, 2).End(xlUp).Row, 2)).Cells
'                If rngZelle > 50 And IsNumeric(rngZelle) Then Range(rngZelle.Address) = rngZelle / 1000
'            Next
'
'            For intZeile = 1 To Cells(Rows.Count, 2).End(xlUp).Row
'                If Cells(intZeile, 5) Like "*LEATHER*" Then 'Lederproben, Trokenmassekorrektur der Einwaage
'                    Workbooks.Open "L:\Makros\Trockenmasse\Trockenmasse-Original.xlsm"
'                    Columns(10).Hidden = False
'                    varProbennameTeile = Split(ThisWorkbook.Sheets("BatchEquipmentExport").Cells(intZeile, 1), ".")
'                    intDoppelbestimmung = IIf(IsNumeric(varProbennameTeile(UBound(varProbennameTeile) - 1)), 0, 1)
'                    Set rngSample = Range(Cells(12, 10), Cells(Cells(Rows.Count, 10).End(xlUp).Row, 10)).Find(varProbennameTeile(0) & "." & varProbennameTeile(1) & "." & Left(varProbennameTeile(1), Len(varProbennameTeile(1) - intDoppelbestimmung)), LookIn:=xlValues)
'                    If Not rngSample Is Nothing Then
'                        .Cells(.Cells(Rows.Count, 3).End(xlUp).Row + 1, 3) = ThisWorkbook.Sheets("BatchEquipmentExport").Cells(intZeile, 2) - (ThisWorkbook.Sheets("BatchEquipmentExport").Cells(intZeile, 2) * Cells(rngSample.Row + 1, 9) / 100)
'                    Else: .Cells(.Cells(Rows.Count, 3).End(xlUp).Row + 1, 3) = 0.001
'                    End If
'                    Workbooks("Trockenmasse-Original.xlsm").Close savechanges:=False
'                Else
'                    varEinzelEinwaagen = Split(Cells(intZeile, 2), "/")
'                    For j = 0 To UBound(varEinzelEinwaagen): dblSumme = dblSumme + CDbl(varEinzelEinwaagen(j)): Next j
'                    .Cells(.Cells(Rows.Count, 3).End(xlUp).Row + 1, 3) = dblSumme: dblSumme = 0
'                End If
'                .Cells(.Cells(Rows.Count, 4).End(xlUp).Row + 1, 4) = dblStdEinwaage / .Cells(.Cells(Rows.Count, 3).End(xlUp).Row, 3)
'            Next intZeile
'            ActiveSheet.Delete
'        End With
'        Application.DisplayAlerts = True
'    Next i
'    With Sheets("Hauptseite")
'        With .Range(.Cells(3, 2), .Cells(432, 5))
'            .Font.Size = 18
'            .Interior.ColorIndex = 2
'            .Font.Bold = True
'            .Locked = False
'        End With
'        .Cells(4, 8) = strShortName
'        .Protect
'        .Select
'    End With
'End If
'
'Application.EnableEvents = True: Application.DisplayAlerts = True: Application.ScreenUpdating = True
'
'End Sub

Private Sub Sequence()
    ' Strings für Methodeninformationen
    Const strMethodedaten As String = "L:\Makros\Sequenceschreiber\Daten für Sequenceschreiber.xlsx"
    Dim strExportordner As String
    
    ' Integer für Zeilenpositionen
    Dim intMethodenZeile As Integer
    Dim intZeileSequence As Integer
    Dim intGeräteZeile As Integer
    
    ' Arrays für Daten
    Dim arrQuellKolonne As Variant
    
    ' Range für Loops
    Dim rngZelle As Range
    Dim rngProbenRange As Range
    Set rngProbenRange = Range(Cells(3, 2), Cells(Cells(2, 2).End(xlDown).Row, 2))
    
    ' Object und Collection
    Dim dictMetadaten As Object
    Dim dictBatchdaten As Object
    Dim dictMethodedaten As Object
    Dim dictTrigger As Object
    Dim dictKolonnenposition As Object
    Dim objProben As New CWerte
    Dim colProben As Collection
    Dim objKalibration As New CWerte
    Dim colKalibration As Collection
    Dim objBlank As New CWerte
    Dim colBlank As Collection
    Dim objSpezialproben As New CWerte
    Dim colSpezialproben As Collection
    Dim objGanzspalten As New CWerte

    Set dictMetadaten = CreateObject("Scripting.Dictionary")
    Set dictBatchdaten = CreateObject("Scripting.Dictionary")
    Set dictMethodedaten = CreateObject("Scripting.Dictionary")
    Set dictTrigger = CreateObject("Scripting.Dictionary")
    Set dictKolonnenposition = CreateObject("Scripting.Dictionary")
    Set colProben = New Collection
    Set colKalibration = New Collection
    Set colSpezialproben = New Collection
    
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
            For prpName = AcquisitionMethode To Sequencename
                .Add defGetPropertyName(prpName), dictMetadaten("wbDaten").Sheets("Geräte").Cells(intGeräteZeile, prpName).value
            Next prpName
        End With
        Set dictMetadaten("Kolonnenposition") = dictKolonnenposition
        strExportordner = dictMetadaten("wbDaten").Sheets("Geräte").Cells(intGeräteZeile, Exportordner).value & "\"
        
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
            .Add "ZwischenBlankTrigger", dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Zwischenkali ab X Proben", arrQuellKolonne, 0) + 1)
            .Add "ZwischenKalibartionTrigger", dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Zwischenkali ab X Proben", arrQuellKolonne, 0) + 1)
            .Add "ZwischenKalibartionModus", IIf(dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Einzel/Volle Zwischenkali", arrQuellKolonne, 0) + 1) = "Einzel", 1, .Item("Kalibrationsanzahl"))
        End With
        Set dictMetadaten("Methodedaten") = dictMethodedaten
        
        ' Blankwerte auslesen
        For prpName = AcquisitionMethode To Wert4
            If Not dictMetadaten("Kolonnenposition")(defGetPropertyName(prpName)) = -1 Then _
               Call defSetWert(prp:=prpName, MessTyp:=Blank, Metadaten:=dictMetadaten, Blank:=objBlank)
        Next prpName
        
        ' Kalibrationwerte auslesen
        For i = 1 To dictMethodedaten("Kalibrationsanzahl")
            Set objKalibration = New CWerte
            colKalibration.Add objKalibration
            For prpName = AcquisitionMethode To Wert4
                If Not dictMetadaten("Kolonnenposition")(defGetPropertyName(prpName)) = -1 Then _
                   Call defSetWert(prp:=prpName, MessTyp:=Kalibration, Metadaten:=dictMetadaten, Kalibration:=colKalibration, Collectionindex:=i)
            Next prpName
        Next i
         
        ' Spetialprobenwerte auslesen
        For i = 1 To 3
            If Not dictMetadaten("wsDaten").Cells(intMethodenZeile, (i - 1) * 2 + Application.Match("Spezialprobe 1 Probe 1 nach Kali", arrQuellKolonne, 0) + 1) = "" Then
                Set objSpezialproben = New CWerte
                colSpezialproben.Add objSpezialproben
                For prpName = AcquisitionMethode To Wert4
                    If Not dictMetadaten("Kolonnenposition")(defGetPropertyName(prpName)) = -1 Then _
                       Call defSetWert(prp:=prpName, MessTyp:=Spezialprobe, Metadaten:=dictMetadaten, Spezialproben:=colSpezialproben, Kalibration:=colKalibration, Collectionindex:=i)
                Next prpName
            End If
        Next i
                
        ' Probenwerte auslesen
        For i = 1 To dictBatchdaten("TotalProben")
            Set objProben = New CWerte
            colProben.Add objProben
            For prpName = AcquisitionMethode To Wert4
                If Not dictMetadaten("Kolonnenposition")(defGetPropertyName(prpName)) = -1 Then _
                   Call defSetWert(prp:=prpName, MessTyp:=Sample, Metadaten:=dictMetadaten, Probe:=colProben, Kalibration:=colKalibration, Collectionindex:=i)
            Next prpName
        Next i
        
        ' Trigger Definieren
        With dictTrigger
            .Add "MaxKalibration", Int(dictBatchdaten("TotalProben") / dictMethodedaten("ZwischenKalibartionTrigger"))
            .Add "AnzahlProbenZwischenKalibrationen", Int(dictBatchdaten("TotalProben") / (.Item("MaxKalibration") + 1))
            .Add "AnzahlProbenZwischenBlank", .Item("AnzahlProbenZwischenKalibrationen") \ ((.Item("AnzahlProbenZwischenKalibrationen") \ dictMethodedaten("ZwischenBlankTrigger")) + 1)
        End With
        Set dictMetadaten("Trigger") = dictTrigger
        
        For prpName = Sequencename To Sequencename
            If Not dictMetadaten("Kolonnenposition")(defGetPropertyName(prpName)) = -1 Then _
               Call defSetWert(prp:=prpName, MessTyp:=Ganzspalten, Metadaten:=dictMetadaten, Ganzspalten:=objGanzspalten)
        Next prpName
            
        ''' Sequence schreiben '''
        With wsSequence
            .Visible = True
            .Cells.ClearContents
            ' Anfangskalibration
            Call defInsertBlank(intZeileSequence, dictMetadaten, objBlank)
            Call defInsertKalibration(intZeileSequence, dictMetadaten, colKalibration, True)
            Call defInsertBlank(intZeileSequence, dictMetadaten, objBlank)
            
            ' Spezialproben
            If Not colSpezialproben.Count = 0 Then
                Call defInsertSpezialproben(intZeileSequence, dictMetadaten, colSpezialproben)
                Call defInsertBlank(intZeileSequence, dictMetadaten, objBlank)
            End If
            
            For Each rngZelle In rngProbenRange
                ' Probe
                intZeileSequence = intZeileSequence + 1
                For prpName = AcquisitionMethode To Wert4
                    If Not dictMetadaten("Kolonnenposition")(defGetPropertyName(prpName)) = -1 Then _
                       .Cells(intZeileSequence, dictMetadaten("Kolonnenposition")(defGetPropertyName(prpName))) = defGetWert(prp:=prpName, MessTyp:=0, Probe:=colProben, Collectionindex:=rngZelle.Row - rngProbenRange.Row + 1)
                Next prpName
                dictMetadaten("Trigger")("CurrentKalibrationTriggerCount") = dictMetadaten("Trigger")("CurrentKalibrationTriggerCount") + 1
                dictMetadaten("Trigger")("CurrentBlankTriggerCount") = dictMetadaten("Trigger")("CurrentBlankTriggerCount") + 1
                
                ' Zwischenkali
                If dictMetadaten("Trigger")("CurrentKalibrationTriggerCount") = dictMetadaten("Trigger")("AnzahlProbenZwischenKalibrationen") _
                   And rngZelle.Row - rngProbenRange.Row + 2 < dictMetadaten("Batchdaten")("TotalProben") Then ' diese Zeile kappt unnötige Zwischenkalibrationen am Ende der Sequence
                    Call defInsertBlank(intZeileSequence, dictMetadaten, objBlank)
                    Call defInsertKalibration(intZeileSequence, dictMetadaten, colKalibration, dictMethodedaten("ZwischenkaliEinzel_Volle") = "Volle")
                    Call defInsertBlank(intZeileSequence, dictMetadaten, objBlank)
                End If
                
                ' Zwischenblank
                If dictMetadaten("Trigger")("CurrentBlankTriggerCount") = dictMetadaten("Trigger")("AnzahlProbenZwischenBlank") _
                   And dictMetadaten("Trigger")("CurrentKalibrationTriggerCount") + 2 < dictMetadaten("Trigger")("AnzahlProbenZwischenKalibrationen") Then ' diese Zeile kappt unnötige Zwischenblanks vor den Zwischenkalibrationen
                    Call defInsertBlank(intZeileSequence, dictMetadaten, objBlank)
                End If
                
            Next rngZelle
            
            ' Schlusskalibration
            Call defInsertBlank(intZeileSequence, dictMetadaten, objBlank)
            Call defInsertKalibration(intZeileSequence, dictMetadaten, colKalibration, True)
            Call defInsertBlank(intZeileSequence, dictMetadaten, objBlank)
            
            ' Ganzbatchkolonnen
            For prpName = Sequencename To Sequencename
                If Not dictMetadaten("Kolonnenposition")(defGetPropertyName(prpName)) = -1 Then _
                   .Range(.Cells(1, dictMetadaten("Kolonnenposition")(defGetPropertyName(prpName))), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, dictMetadaten("Kolonnenposition")(defGetPropertyName(prpName)))) = _
                   defGetWert(prp:=prpName, MessTyp:=4, Ganzspalten:=objGanzspalten)
            Next prpName
        
            ' Sequence in Clipboard überführen oder exportieren
            If strExportordner = "-1\" Then
                .UsedRange.Copy
            ElseIf Not IsFileOpen(strExportordner & dictMetadaten("Batchdaten")("Methode") & "_" & dictMetadaten("Batchdaten")("Topic") & ".csv") Then
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
    
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Private Sub defInsertBlank(Row As Integer, Metadaten As Object, Blank As Object)

    Row = Row + 1
    For prpName = AcquisitionMethode To Wert4
        If Not Metadaten("Kolonnenposition")(defGetPropertyName(prpName)) = -1 Then _
           wsSequence.Cells(Row, Metadaten("Kolonnenposition")(defGetPropertyName(prpName))) = defGetWert(prp:=prpName, MessTyp:=2, Blank:=Blank)
    Next prpName
    Metadaten("Trigger")("CurrentBlankTriggerCount") = 0
End Sub

Private Sub defInsertKalibration(Row As Integer, Metadaten As Object, Kalibration As Object, Volle_Kalibration As Boolean)

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
            For prpName = AcquisitionMethode To Wert4
                If Not Metadaten("Kolonnenposition")(defGetPropertyName(prpName)) = -1 Then _
                   wsSequence.Cells(Row, Metadaten("Kolonnenposition")(defGetPropertyName(prpName))) = defGetWert(prp:=prpName, MessTyp:=1, Kalibration:=Kalibration, Collectionindex:=i)
            Next prpName
        Next i
    Else
        Row = Row + 1
        For prpName = AcquisitionMethode To Wert4
            If Not Metadaten("Kolonnenposition")(defGetPropertyName(prpName)) = -1 Then _
               wsSequence.Cells(Row, Metadaten("Kolonnenposition")(defGetPropertyName(prpName))) = defGetWert(prp:=prpName, MessTyp:=1, Kalibration:=Kalibration, Collectionindex:=Round(Kalibration.Count / 2, 0))
        Next prpName
    End If
    
    Metadaten("Trigger")("CurrentKalibrationTriggerCount") = IIf(blnExtra = True, -1, 0)
End Sub

Private Sub defInsertSpezialproben(Row As Integer, Metadaten As Object, Spezialproben As Collection)
    
    For i = 1 To Spezialproben.Count
        Row = Row + 1
        For prpName = AcquisitionMethode To Wert4
            If Not Metadaten("Kolonnenposition")(defGetPropertyName(prpName)) = -1 Then _
               wsSequence.Cells(Row, Metadaten("Kolonnenposition")(defGetPropertyName(prpName))) = defGetWert(prp:=prpName, MessTyp:=3, Spezialprobe:=Spezialproben, Collectionindex:=i)
        Next prpName
    Next i

End Sub

Private Sub defSetWert(prp As Properties, MessTyp As MessTypen, Metadaten As Object, _
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
        If MessTyp = 0 Then
            Select Case prp
            Case AcquisitionMethode: Probe(Collectionindex).AcquisitionMethode = defGetMethode(wsHauptseite.Cells(Collectionindex + 2, 5), Metadaten)
            Case Quantmethode: Probe(Collectionindex).Quantmethode = .Cells(intMethodenZeile, Application.Match("Quantmethode", arrQuellKolonne, 0) + 1)
            Case Beschriftung: Probe(Collectionindex).Beschriftung = wsHauptseite.Cells(Collectionindex + 2, 2)
            Case Einwaage: Probe(Collectionindex).Einwaage = wsHauptseite.Cells(Collectionindex + 2, 3)
            Case Exctraktionsvolumen: Probe(Collectionindex).Exctraktionsvolumen = .Cells(intMethodenZeile, Application.Match("Exctraktionsvolumen", arrQuellKolonne, 0) + 1)
            Case Injektionsvolumen: Probe(Collectionindex).Injektionsvolumen = .Cells(intMethodenZeile, Application.Match("Injektionsvolumen", arrQuellKolonne, 0) + 1)
            Case Kommentar: Probe(Collectionindex).Kommentar = wsHauptseite.Cells(Collectionindex + 2, 6)
            Case Rack: Probe(Collectionindex).Rack = "Rack" 'Fehlt!
            Case Position: Probe(Collectionindex).Position = IIf(Collectionindex = 1, _
                                                                 Metadaten("Methodedaten")("Spezialbrobenanzahl") + Kalibration(Kalibration.Count).Position + 1, _
                                                                 defGetPosition(Probe:=Probe, Collectionindex:=Collectionindex, Metadaten:=Metadaten))
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
            Case Else: GoTo errHandler
            End Select
            
            ' Wert für Kalibration
        ElseIf MessTyp = 1 Then
            Select Case prp
            Case AcquisitionMethode: Kalibration(Collectionindex).AcquisitionMethode = defGetMethode(Metadaten("Batchdaten")("Topic"), Metadaten)
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
            Case Else: GoTo errHandler
            End Select
            ' Wert für Blank
        ElseIf MessTyp = 2 Then
            Select Case prp
            Case AcquisitionMethode: Blank.AcquisitionMethode = defGetMethode(Metadaten("Batchdaten")("Topic"), Metadaten)
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
            Case Else: GoTo errHandler
            End Select
            
            ' Wert für Spezialprobe
        ElseIf MessTyp = 3 Then
            Select Case prp
            Case AcquisitionMethode: Spezialproben(Collectionindex).AcquisitionMethode = defGetMethode(Metadaten("Batchdaten")("Topic"), Metadaten)
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
            Case Else: GoTo errHandler
            End Select
            
            ' Wert für Ganzspalten
        ElseIf MessTyp = 4 Then
            Select Case prp
            Case Sequencename: Ganzspalten.Sequencename = Format(Now(), "yymmdd") & "_" & Metadaten("Batchdaten")("Operator") & "_" & defGetMethode(Metadaten("Batchdaten")("Topic"), Metadaten)
            Case Else:                           ' Aktion für unbekannte Eigenschaft
            End Select
        End If
    End With
    
    Exit Sub
    
errHandler:
    ActiveWorkbook.Close savechanges:=False
    MsgBox "Es gab ein Fehler beim Implementieren eines wertes." & vbCr & "Bitte melde Dich beim Digital Laboratory Expert.", vbCritical, "Fehlender Wert"
    End

End Sub

Private Function defGetPosition(Probe As Collection, Collectionindex As Integer, Metadaten As Object) As Integer
    If Collectionindex > 1 Then defGetPosition = Probe(Collectionindex - 1).Position + 1
    If defGetPosition > Metadaten("Methodedaten")("RackPositionen") Then defGetPosition = 1
End Function

Private Function defGetWert(prp As Properties, MessTyp As Integer, _
                         Optional ByVal Probe As Collection = Nothing, _
                         Optional ByVal Kalibration As Collection = Nothing, _
                         Optional ByVal Blank As Object = Nothing, _
                         Optional ByVal Spezialprobe As Collection = Nothing, _
                         Optional ByVal Ganzspalten As Object = Nothing, _
                         Optional ByVal Collectionindex As Integer = -1) As Variant

    Dim obj As Object
    Dim varValue As Variant
    
    ' Bestimmen Sie das entsprechende Objekt basierend auf dem MessTyp
    Select Case MessTyp
    Case 0: Set obj = Probe(Collectionindex)
    Case 1: Set obj = Kalibration(Collectionindex)
    Case 2: Set obj = Blank
    Case 3: Set obj = Spezialprobe(Collectionindex)
    Case 4: Set obj = Ganzspalten
    Case Else: Set obj = Nothing
    End Select
    
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
        Case Else:                               ' Aktion für unbekannte Eigenschaft
        End Select

    Else
        varValue = "Unknown"
    End If

    defGetWert = varValue
    
End Function

Private Function defGetPropertyName(prp As Properties) As String
    Select Case prp
    Case AcquisitionMethode: defGetPropertyName = "AcquisitionMethode"
    Case Quantmethode: defGetPropertyName = "Quantmethode"
    Case Beschriftung: defGetPropertyName = "Beschriftung"
    Case Einwaage: defGetPropertyName = "Einwaage"
    Case Exctraktionsvolumen: defGetPropertyName = "Exctraktionsvolumen"
    Case Injektionsvolumen: defGetPropertyName = "Injektionsvolumen"
    Case Kommentar: defGetPropertyName = "Kommentar"
    Case Konzentration: defGetPropertyName = "Konzentration"
    Case Position: defGetPropertyName = "Position"
    Case Produktklasse: defGetPropertyName = "Produktklasse"
    Case Rack: defGetPropertyName = "Rack"
    Case Typ: defGetPropertyName = "Typ"
    Case Verdünnung: defGetPropertyName = "Verdünnung"
    Case Level: defGetPropertyName = "Level"
    Case Info1: defGetPropertyName = "Info1"
    Case Info2: defGetPropertyName = "Info2"
    Case Info3: defGetPropertyName = "Info3"
    Case Info4: defGetPropertyName = "Info4"
    Case Wert1: defGetPropertyName = "Wert1"
    Case Wert2: defGetPropertyName = "Wert2"
    Case Wert3: defGetPropertyName = "Wert3"
    Case Wert4: defGetPropertyName = "Wert4"
    Case Sequencename: defGetPropertyName = "Sequencename"
    Case Else: defGetPropertyName = "Unknown"
    End Select
End Function

' Funktion zum Abrufen der Messmethode
Private Function defGetMethode(strTopic As Variant, Metadaten As Object) As String
    Select Case Left(strTopic, 3)
    Case "Std", "STA": defGetMethode = Metadaten("Methodedaten")("MethodeSTD100")
    Case "L", "LEA": defGetMethode = Metadaten("Methodedaten")("MethodeLeder")
    Case "ECP", "ECO": defGetMethode = Metadaten("Methodedaten")("MethodeECO")
    Case "CAL", "CAL": defGetMethode = Metadaten("Methodedaten")("MethodeKalibration")
    Case Else                                    ' Aktion für unbekannte Eigenschaft
    End Select

End Function

'Sub Sequence_old()
'
'Dim strGC As String, strMethodeKalibration As String, strMethodeProbe As String, strTopic As String, strOperator As String, strZwischenkalibratinsTyp As String, strBlank As String
'Dim strKalibration() As String, strSpecialProbe(2) As String, strSpecialProbeTyp(2) As String, strKalibrationPosition() As String, strAuswertemethode As String, strAuswertepfad As String, strAuswertemethodeUndPfad As String
'Dim intPosition As Integer, intRack As Integer, intProbenanzahl As Integer, intZeile As Integer, intKalibrationOne As Integer
'Dim intAnzahlRaks As Integer, intRackMax As Integer, intZwischenBlanzTrigger As Integer, intZwischenKalibartionsTrigger As Integer
'Dim intKalibrationswechsel As Integer, intKalibrationsanzahl As Integer, intZwischenkaliEinelOderVoll As Integer, intSpezialproben As Integer
'Dim intZeileHaubtseite As Integer, intZwischenBlankGetPropertyName(prpName)Trigger As Integer, intTrigger As Integer, intZwischenBlankTriggerCount As Integer
'Dim intZwischenKaliTriggerCount As Integer, intMethodenZeile As Integer, intKalibrationPosition() As Integer
'Dim DatenWB As Workbook
'Dim blnJetClean As Boolean
'Dim arrQuellKolonne As Variant
'
'Application.EnableEvents = False: Application.DisplayAlerts = False: Application.ScreenUpdating = False
'
'If Cells(3, 8) = "Methoden" Then
'    MsgBox ("Bitte Methode wählen. Danke."): End
'Else
'    Sheets("Sequence").Visible = True
'    Sheets("Sequence").Range(Sheets("Sequence").Columns(1), Sheets("Sequence").Columns(10)).ClearContents
'    strGC = Cells(2, 8)
'    strMethodeKalibration = Cells(3, 8)
'    strTopic = Cells(3, 9)
'    intPosition = Cells(5, 9)
'    intRack = Columns(14).Find(Cells(5, 8)).Row
'    strOperator = Cells(4, 8)
'    intProbenanzahl = Cells(Rows.Count, 2).End(xlUp).Row - 2
'
'    Workbooks.Open "L:\Makros\Sequenceschreiber\GC\Daten für GC Sequenceschreiber.xlsx"
'    Set DatenWB = Workbooks("Daten für GC Sequenceschreiber.xlsx")
'    blnJetClean = DatenWB.Sheets(1).Cells(DatenWB.Sheets(1).Columns(3).Find(strGC).Row, 5)
'    With DatenWB.Sheets(strGC)
'        arrQuellKolonne = .Range(.Cells(2, 2), .Cells(2, Columns.Count).End(xlToLeft))
'        intMethodenZeile = .Columns(Application.Match("Methodenname Kalibration (MUSS GENAU STIMMEN!)", arrQuellKolonne, 0) + 1).Find(strMethodeKalibration).Row
'
'        strMethodeProbe = .Cells(intMethodenZeile, Application.Match("Methodenname Probe (MUSS GENAU STIMMEN!)", arrQuellKolonne, 0) + 1)
'        strAuswertemethodeUndPfad = .Cells(intMethodenZeile, Application.Match("Auswertemethode und Pfad (MUSS GENAU STIMMEN!)", arrQuellKolonne, 0) + 1)
'        strAuswertemethode = Mid(strAuswertemethodeUndPfad, InStrRev(strAuswertemethodeUndPfad, "\") + 1)
'        strAuswertepfad = Left(strAuswertemethodeUndPfad, InStrRev(strAuswertemethodeUndPfad, "\"))
'        intAnzahlRaks = .Cells(intMethodenZeile, Application.Match("Anzahl Rack", arrQuellKolonne, 0) + 1) + 1
'        intRackMax = .Cells(intMethodenZeile, Application.Match("Max Position", arrQuellKolonne, 0) + 1)
'        intZwischenBlanzTrigger = .Cells(intMethodenZeile, Application.Match("Zwischenblank ab X Proben", arrQuellKolonne, 0) + 1)
'        intZwischenKalibartionsTrigger = .Cells(intMethodenZeile, Application.Match("Zwischenkali ab X Proben", arrQuellKolonne, 0) + 1)
'        intKalibrationswechsel = IIf(IsEmpty(.Cells(intMethodenZeile, Application.Match("Wechsel nach X Einstichen", arrQuellKolonne, 0) + 1)), 100, .Cells(intMethodenZeile, Application.Match("Wechsel nach X Einstichen", arrQuellKolonne, 0) + 1))
'        intKalibrationsanzahl = .Cells(intMethodenZeile, Columns.Count).End(xlToLeft).Column - (Application.Match("Lösungsmittel", arrQuellKolonne, 0) + 1)
'        intZwischenkaliEinelOderVoll = IIf(.Cells(intMethodenZeile, Application.Match("Einzel/Volle Zwischenkali", arrQuellKolonne, 0) + 1) = "Einzel", 1, intKalibrationsanzahl)
'        strZwischenkalibratinsTyp = .Cells(intMethodenZeile, Application.Match("Zwischenkali als QC oder Cal", arrQuellKolonne, 0) + 1)
'        For i = 0 To 2
'            strSpecialProbe(i) = .Cells(intMethodenZeile, i * 2 + Application.Match("Spezialprobe 1 Probe 1 nach Kali", arrQuellKolonne, 0) + 1)
'            strSpecialProbeTyp(i) = .Cells(intMethodenZeile, i * 2 + Application.Match("Type für Spezialprobe 1", arrQuellKolonne, 0) + 1)
'        Next i
'        strBlank = .Cells(intMethodenZeile, Application.Match("Lösungsmittel", arrQuellKolonne, 0) + 1)
'        ReDim strKalibration(1 To intKalibrationsanzahl)
'        For i = 1 To intKalibrationsanzahl: strKalibration(i) = .Cells(intMethodenZeile, i + Application.Match("Kalibration Level 1", arrQuellKolonne, 0)): Next i
'        DatenWB.Close (False)
'    End With
'
'    '********* Sequence schreiben *********'
'
'    With Sheets("Sequence")
'        For i = 1 To wsHauptseite.Cells(17, 11) 'Anzahl Anfangsblanks
'        .Cells(i, 1) = intPosition 'Blank 1
'        .Cells(i, 2) = "DoubleBlank"
'        .Cells(i, 3) = strBlank
'        .Cells(i, 4) = strMethodeKalibration
'        .Cells(i, 7) = Cells(intRack, 14)
'        .Cells(i, 10) = 1
'        Next i
'        intPosition = intPosition + 1
'        If intPosition > intRackMax Then
'            intPosition = 1
'            intRack = intRack + 1
'            If intRack > intAnzahlRaks Then intRack = 2
'        End If
'        ReDim intKalibrationPosition(1 To intKalibrationsanzahl) 'Kal 1
'        ReDim strKalibrationPosition(1 To intKalibrationsanzahl)
'        For intKalibrationOne = 1 To intKalibrationsanzahl
'            .Cells(intKalibrationOne + i - 1, 1) = intPosition
'            .Cells(intKalibrationOne + i - 1, 2) = "Cal"
'            .Cells(intKalibrationOne + i - 1, 3) = strKalibration(intKalibrationOne)
'            .Cells(intKalibrationOne + i - 1, 4) = strMethodeKalibration
'            .Cells(intKalibrationOne + i - 1, 7) = Cells(intRack, 14)
'            .Cells(intKalibrationOne + i - 1, 9) = intKalibrationOne
'            .Cells(intKalibrationOne + i - 1, 10) = 1
'            intKalibrationPosition(intKalibrationOne) = .Cells(intKalibrationOne + i - 1, 1)
'            strKalibrationPosition(intKalibrationOne) = .Cells(intKalibrationOne + i - 1, 7)
'            intPosition = intPosition + 1
'            If intPosition > intRackMax Then
'                intPosition = 1
'                intRack = intRack + 1
'                If intRack > intAnzahlRaks Then intRack = 2
'            End If
'        Next intKalibrationOne
'        .Cells(intKalibrationsanzahl + i, 1) = .Cells(1, 1) 'Blank 2
'        .Cells(intKalibrationsanzahl + i, 2) = "DoubleBlank"
'        .Cells(intKalibrationsanzahl + i, 3) = strBlank
'        .Cells(intKalibrationsanzahl + i, 4) = strMethodeKalibration
'        .Cells(intKalibrationsanzahl + i, 7) = Cells(intRack, 14)
'        .Cells(intKalibrationsanzahl + i, 10) = 1
'        intPosition = intPosition + (WorksheetFunction.RoundUp(Int((intProbenanzahl / intZwischenKalibartionsTrigger) + 2) / intKalibrationswechsel, 0) - 1) * intZwischenkaliEinelOderVoll
'        If intPosition > intRackMax Then
'            intPosition = intPosition - intRackMax
'            intRack = intRack + 1
'            If intRack > intAnzahlRaks Then intRack = 2
'        End If
'        For intSpezialproben = 0 To 2 'Specialproben
'            If Not strSpecialProbe(intSpezialproben) = "" Then
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row + 1, 1) = intPosition
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 2) = strSpecialProbeTyp(intSpezialproben)
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 3) = strSpecialProbe(intSpezialproben)
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 4) = strMethodeProbe
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 7) = Cells(intRack, 14)
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 10) = 1
'                intPosition = intPosition + 1
'                If intPosition > intRackMax Then
'                    intPosition = intPosition - intRackMax
'                    intRack = intRack + 1
'                    If intRack > intAnzahlRaks Then intRack = 2
'                End If
'            End If
'        Next intSpezialproben
'        For intZeileHaubtseite = 3 To 2 + intProbenanzahl
'            .Cells(.Cells(Rows.Count, 1).End(xlUp).Row + 1, 1) = intPosition 'Cells(intZeileHaubtseite, 2)
'            .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 2) = "Sample"
'            .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 3) = Cells(intZeileHaubtseite, 2)
'            .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 4) = strMethodeProbe
'            .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 7) = Cells(intRack, 14)
'            .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 10) = Cells(intZeileHaubtseite, 4)
'            .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 11) = Cells(intZeileHaubtseite, 5)
'            intPosition = intPosition + 1
'            If intPosition > intRackMax Then
'                intPosition = intPosition - intRackMax
'                intRack = intRack + 1
'                If intRack > intAnzahlRaks Then intRack = 2
'            End If
'            intTrigger = intTrigger + 1
'            intZwischenBlankTrigger = intZwischenBlankTrigger + 1 'Zwischenblank
'            If intZwischenBlankTrigger = Int((Int(intProbenanzahl / (1 + Int(intProbenanzahl / intZwischenKalibartionsTrigger)))) / (Int((Int(intProbenanzahl / (1 + Int(intProbenanzahl / intZwischenKalibartionsTrigger)))) / intZwischenBlanzTrigger) + 1)) And Not intZwischenBlankTriggerCount = Int((Int(intProbenanzahl / (1 + Int(intProbenanzahl / intZwischenKalibartionsTrigger)))) / intZwischenBlanzTrigger) Then
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row + 1, 1) = .Cells(1, 1)
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 2) = "DoubleBlank"
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 3) = strBlank
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 4) = strMethodeKalibration
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 7) = .Cells(1, 7)
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 10) = 1
'                intZwischenBlankTrigger = 0
'                intZwischenBlankTriggerCount = intZwischenBlankTriggerCount + 1
'            End If
'            If intTrigger = Int(intProbenanzahl / (1 + Int(intProbenanzahl / intZwischenKalibartionsTrigger))) And Not intZwischenKaliTriggerCount = Int(intProbenanzahl / intZwischenKalibartionsTrigger) Then 'Zwischenkali
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row + 1, 1) = .Cells(1, 1)
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 2) = "DoubleBlank"
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 3) = strBlank
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 4) = strMethodeKalibration
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 7) = .Cells(1, 7)
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 10) = 1
'                If intZwischenkaliEinelOderVoll = 1 Then
'                    .Cells(.Cells(Rows.Count, 1).End(xlUp).Row + 1, 1) = intKalibrationPosition(WorksheetFunction.RoundUp(intKalibrationsanzahl / 2, 0))
'                    .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 2) = strZwischenkalibratinsTyp
'                    .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 3) = strKalibration(WorksheetFunction.RoundUp(intKalibrationsanzahl / 2, 0))
'                    .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 4) = strMethodeKalibration
'                    .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 7) = strKalibrationPosition(WorksheetFunction.RoundUp(intKalibrationsanzahl / 2, 0))
'                    .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 9) = WorksheetFunction.RoundUp(intKalibrationsanzahl / 2, 0)
'                    .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 10) = 1
'                Else
'                    For i = 1 To intKalibrationsanzahl
'                        .Cells(.Cells(Rows.Count, 1).End(xlUp).Row + 1, 1) = intKalibrationPosition(i)
'                        .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 2) = strZwischenkalibratinsTyp
'                        .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 3) = strKalibration(i)
'                        .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 4) = strMethodeKalibration
'                        .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 7) = strKalibrationPosition(i)
'                        .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 9) = i
'                        .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 10) = 1
'                    Next i
'                End If
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row + 1, 1) = .Cells(1, 1)
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 2) = "DoubleBlank"
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 3) = strBlank
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 4) = strMethodeKalibration
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 7) = .Cells(1, 7)
'                .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 10) = 1
'                intZwischenKaliTriggerCount = intZwischenKaliTriggerCount + 1
'                intTrigger = 0
'                intZwischenBlankTrigger = 0
'                intZwischenBlankTriggerCount = 0
'            End If
'        Next intZeileHaubtseite
'        .Cells(.Cells(Rows.Count, 1).End(xlUp).Row + 1, 1) = .Cells(1, 1)
'        .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 2) = "DoubleBlank"
'        .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 3) = strBlank
'        .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 4) = strMethodeKalibration
'        .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 7) = .Cells(1, 7)
'        .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 10) = 1
'        For i = 1 To intKalibrationsanzahl
'            .Cells(.Cells(Rows.Count, 1).End(xlUp).Row + 1, 1) = intKalibrationPosition(i)
'            .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 2) = "Cal"
'            .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 3) = strKalibration(i)
'            .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 4) = strMethodeKalibration
'            .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 7) = strKalibrationPosition(i)
'            .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 9) = i
'            .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 10) = 1
'            intPosition = intPosition + 1
'            If intPosition > intRackMax Then
'                intPosition = intPosition - intRackMax
'                intRack = intRack + 1
'                If intRack > intAnzahlRaks Then intRack = 2
'            End If
'        Next i
'        .Cells(.Cells(Rows.Count, 1).End(xlUp).Row + 1, 1) = .Cells(1, 1)
'        .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 2) = "DoubleBlank"
'        .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 3) = strBlank
'        .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 4) = strMethodeKalibration
'        .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 7) = .Cells(1, 7)
'        .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 10) = 1
'        For i = 1 To .Cells(Rows.Count, 1).End(xlUp).Row
'            .Cells(i, 5) = "D:\MassHunter\GCMS\1\data\" & Format(Date, "YYMMdd") & "_" & strOperator & "_" & strTopic
'            .Cells(i, 6) = Format(Date, "YYMMdd") & "_" & strTopic & "_" & Format(i, "0#")
'        Next i
'        '.Range(.Cells(1, 11), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 11)) = strAuswertepfad
'        '.Range(.Cells(1, 12), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 12)) = strAuswertemethode
'        If blnJetClean = True Then
'            .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 4) = "JetClean_manuell.M"
'            .Cells(1, 4) = "JetClean_manuell.M"
'        End If
'        .Range(.Cells(1, 1), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 12)).Copy
'    End With
'End If
'Sheets("Sequence").Visible = False
'
'Application.EnableEvents = True: Application.DisplayAlerts = True: Application.ScreenUpdating = True
'
'End Sub
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


