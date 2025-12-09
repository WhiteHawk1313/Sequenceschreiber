Attribute VB_Name = "modUtilities"
Option Explicit

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
    DataFile = 29
End Enum

Public Enum MeasurementTypes
    Sample = 0
    SpezialSample = 1
    IntermediateCalibration = 2
    Calibration = 3
    Blank = 4
    FullColumns = 5
End Enum

Public Function funcGetPropertyName(prp As Properties) As String
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
        Case Sequencename: funcGetPropertyName = "Sequencename"
        Case DataFile:  funcGetPropertyName = "Datafile"
        Case Else: funcGetPropertyName = "Unknown"
    End Select
End Function

Public Function funcGetMeasurementType(prp As MeasurementTypes) As String
    Select Case prp
        Case Sample: funcGetMeasurementType = "Sample"
        Case IntermediateCalibration: funcGetMeasurementType = "Zwischenkalibration"
        Case Calibration: funcGetMeasurementType = "Kalibration"
        Case Blank: funcGetMeasurementType = "Blank"
        Case SpezialSample: funcGetMeasurementType = "Spezialprobe"
        Case FullColumns: funcGetMeasurementType = "Ganzspalten"
        Case Else: funcGetMeasurementType = "Unknown"
    End Select
End Function

Public Function funcGetPosition(Probe As Collection, Collectionindex As Integer, MetaData As Object) As Integer

    If Collectionindex > 1 Then funcGetPosition = Probe(Collectionindex - 1).Position + 1
    If funcGetPosition > MetaData("Methodedaten")("RackPositionen") Then funcGetPosition = 1
    
End Function

Public Sub defSetValue(prp As Properties, MeasurementType As MeasurementTypes, Metadaten As Object, _
                    Optional ByVal Probe As Collection = Nothing, _
                    Optional ByVal Kalibration As Collection = Nothing, _
                    Optional ByVal Blank As Object = Nothing, _
                    Optional ByVal Spezialproben As Collection = Nothing, _
                    Optional ByVal Ganzspalten As Object = Nothing, _
                    Optional ByVal Collectionindex As Integer = -1)
    
    Dim intMethodeRow As Integer
    Dim arrSourceColumn As Variant
    
    With Metadaten("wsDaten")
        arrSourceColumn = .Range(.Cells(2, 2), .Cells(2, Columns.Count).End(xlToLeft))
        intMethodeRow = .Columns(Application.Match("Methode", arrSourceColumn, 0) + 1).Find(What:=Metadaten("Batchdaten")("Methode"), LookAt:=xlWhole).Row
        ' Definieren Sie separate Variablen für verschiedene Messarten
        ' Wert für Sample
        If MeasurementType = 0 Then
            Select Case prp
            Case AcquisitionMethode: Probe(Collectionindex).AcquisitionMethode = funcGetMethode(wsMainPage.Cells(Collectionindex + 2, 5), Metadaten)
            Case Quantmethode: Probe(Collectionindex).Quantmethode = .Cells(intMethodeRow, Application.Match("Quantmethode", arrSourceColumn, 0) + 1)
            Case Beschriftung: Probe(Collectionindex).Beschriftung = wsMainPage.Cells(Collectionindex + 2, 2)
            Case Einwaage: Probe(Collectionindex).Einwaage = wsMainPage.Cells(Collectionindex + 2, 3)
            Case Exctraktionsvolumen: Probe(Collectionindex).Exctraktionsvolumen = .Cells(intMethodeRow, Application.Match("Exctraktionsvolumen", arrSourceColumn, 0) + 1)
            Case Injektionsvolumen: Probe(Collectionindex).Injektionsvolumen = .Cells(intMethodeRow, Application.Match("Injektionsvolumen", arrSourceColumn, 0) + 1)
            Case Kommentar: Probe(Collectionindex).Kommentar = wsMainPage.Cells(Collectionindex + 2, 6)
            Case Rack: Probe(Collectionindex).Rack = "Rack" 'Platzhalter für später
            Case Position: Probe(Collectionindex).Position = IIf(Collectionindex = 1, _
                                                                 Metadaten("Methodedaten")("Spezialbrobenanzahl") + Kalibration(Kalibration.Count).Position + 1, _
                                                                 funcGetPosition(Probe:=Probe, Collectionindex:=Collectionindex, MetaData:=Metadaten))
            Case Produktklasse: Probe(Collectionindex).Produktklasse = wsMainPage.Cells(Collectionindex + 2, 5)
            Case Typ: Probe(Collectionindex).Typ = .Cells(intMethodeRow, Application.Match("Proben Typ", arrSourceColumn, 0) + 1)
            Case Konzentration: Probe(Collectionindex).Konzentration = 0
            Case Verdünnung: Probe(Collectionindex).Verdünnung = wsMainPage.Cells(Collectionindex + 2, 4)
            Case Level: Probe(Collectionindex).Level = 0
            Case Info1: Probe(Collectionindex).Info1 = "Sample amount mg or uL"
            Case Info2: Probe(Collectionindex).Info2 = ""
            Case Info3: Probe(Collectionindex).Info3 = ""
            Case Info4: Probe(Collectionindex).Info4 = ""
            Case Wert1: Probe(Collectionindex).Wert1 = 0
            Case Wert2: Probe(Collectionindex).Wert2 = 0
            Case Wert3: Probe(Collectionindex).Wert3 = 0
            Case Wert4: Probe(Collectionindex).Wert4 = 0
            Case Messkategorie: Probe(Collectionindex).Messkategorie = MeasurementTypes.Sample
            Case Else: GoTo ErrHandler
            End Select
            
            ' Wert für Kalibration
        ElseIf MeasurementType = 3 Then
            Select Case prp
            Case AcquisitionMethode: Kalibration(Collectionindex).AcquisitionMethode = funcGetMethode("CALIBRATION", Metadaten)
            Case Quantmethode: Kalibration(Collectionindex).Quantmethode = .Cells(intMethodeRow, Application.Match("Quantmethode", arrSourceColumn, 0) + 1)
            Case Beschriftung: Kalibration(Collectionindex).Beschriftung = .Cells(intMethodeRow, Application.Match("Kalibration Level " & Collectionindex, arrSourceColumn, 0) + 1)
            Case Einwaage: Kalibration(Collectionindex).Einwaage = .Cells(intMethodeRow, Application.Match("Standard-Einwaage", arrSourceColumn, 0) + 1)
            Case Exctraktionsvolumen: Kalibration(Collectionindex).Exctraktionsvolumen = .Cells(intMethodeRow, Application.Match("Exctraktionsvolumen", arrSourceColumn, 0) + 1)
            Case Injektionsvolumen: Kalibration(Collectionindex).Injektionsvolumen = .Cells(intMethodeRow, Application.Match("Injektionsvolumen", arrSourceColumn, 0) + 1)
            Case Kommentar: Kalibration(Collectionindex).Kommentar = ""
            Case Konzentration: Kalibration(Collectionindex).Konzentration = 18 'Fehlt!
            Case Position: Kalibration(Collectionindex).Position = .Cells(intMethodeRow, Application.Match("Position Kalibration Level " & Collectionindex & " (n Positionen nach Lösungsmittel)", arrSourceColumn, 0) + 1) + 1
            Case Produktklasse: Kalibration(Collectionindex).Produktklasse = ""
            Case Rack: Kalibration(Collectionindex).Rack = "Rack" 'Platzhalter für später
            Case Typ: Kalibration(Collectionindex).Typ = .Cells(intMethodeRow, 3 * i + Application.Match("Lösungsmittel", arrSourceColumn, 0) + 2)
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
            Case Messkategorie: Kalibration(Collectionindex).Messkategorie = MeasurementTypes.Calibration
            Case Else: GoTo ErrHandler
            End Select
            ' Wert für Blank
        ElseIf MeasurementType = 4 Then
            Select Case prp
            Case AcquisitionMethode: Blank.AcquisitionMethode = funcGetMethode("CALIBRATION", Metadaten)
            Case Quantmethode: Blank.Quantmethode = .Cells(intMethodeRow, Application.Match("Quantmethode", arrSourceColumn, 0) + 1)
            Case Beschriftung: Blank.Beschriftung = .Cells(intMethodeRow, Application.Match("Lösungsmittel", arrSourceColumn, 0) + 1)
            Case Einwaage: Blank.Einwaage = .Cells(intMethodeRow, Application.Match("Standard-Einwaage", arrSourceColumn, 0) + 1)
            Case Exctraktionsvolumen: Blank.Exctraktionsvolumen = .Cells(intMethodeRow, Application.Match("Exctraktionsvolumen", arrSourceColumn, 0) + 1)
            Case Injektionsvolumen: Blank.Injektionsvolumen = .Cells(intMethodeRow, Application.Match("Injektionsvolumen", arrSourceColumn, 0) + 1)
            Case Kommentar: Blank.Kommentar = ""
            Case Rack: Blank.Rack = "Rack"
            Case Position: Blank.Position = Metadaten("Batchdaten")("Position")
            Case Produktklasse: Blank.Produktklasse = ""
            Case Typ: Blank.Typ = .Cells(intMethodeRow, Application.Match("Lösungsmittel Typ", arrSourceColumn, 0) + 1)
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
            Case Messkategorie: Blank.Messkategorie = MeasurementTypes.Blank
            Case Else: GoTo ErrHandler
            End Select
            
            ' Wert für Spezialprobe
        ElseIf MeasurementType = 1 Then
            Select Case prp
            Case AcquisitionMethode: Spezialproben(Collectionindex).AcquisitionMethode = funcGetMethode(Metadaten("Batchdaten")("Topic"), Metadaten)
            Case Quantmethode: Blank.Quantmethode = .Cells(intMethodeRow, Application.Match("Quantmethode", arrSourceColumn, 0) + 1)
            Case Beschriftung: Spezialproben(Collectionindex).Beschriftung = .Cells(intMethodeRow, (i - 1) * 2 + Application.Match("Spezialprobe 1 Probe 1 nach Kali", arrSourceColumn, 0) + 1)
            Case Einwaage: Spezialproben(Collectionindex).Einwaage = .Cells(intMethodeRow, Application.Match("Standard-Einwaage", arrSourceColumn, 0) + 1)
            Case Exctraktionsvolumen: Spezialproben(Collectionindex).Exctraktionsvolumen = .Cells(intMethodeRow, Application.Match("Exctraktionsvolumen", arrSourceColumn, 0) + 1)
            Case Injektionsvolumen: Spezialproben(Collectionindex).Injektionsvolumen = .Cells(intMethodeRow, Application.Match("Injektionsvolumen", arrSourceColumn, 0) + 1)
            Case Kommentar: Spezialproben(Collectionindex).Kommentar = ""
            Case Rack: Spezialproben(Collectionindex).Rack = "Rack" 'Platzhalter für später
            Case Position: Spezialproben(Collectionindex).Position = Kalibration(Kalibration.Count).Position + i
            Case Produktklasse: Spezialproben(Collectionindex).Produktklasse = ""
            Case Typ: Spezialproben(Collectionindex).Typ = .Cells(intMethodeRow, (i - 1) * 2 + Application.Match("Type für Spezialprobe 1", arrSourceColumn, 0) + 1)
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
            Case Messkategorie: Spezialproben(Collectionindex).Messkategorie = MeasurementTypes.SpezialSample
            Case Else: GoTo ErrHandler
            End Select
            
            ' Wert für Ganzspalten
        ElseIf MeasurementType = 5 Then
            Dim strSaveFolder As String
            Dim strSequenceName As String
            Select Case prp
            Case Sequencename
                strSaveFolder = Metadaten("wbDaten").Sheets("Geräte").Cells(3, Metadaten("wbDaten").Sheets("Geräte").Rows(2).Find("Exportordner").Column + 1).value & "\"
                strSequenceName = Format(Now(), "yymmdd") & "_" & Metadaten("Batchdaten")("Operator") & "_" & Metadaten("Batchdaten")("Topic")
                Ganzspalten.Sequencename = IIf(strSaveFolder = "\", strSequenceName, strSaveFolder & strSequenceName)
            Case DataFile
                strSequenceName = Format(Now(), "yymmdd") & "_" & Metadaten("Batchdaten")("Topic")
                Ganzspalten.DataFile = strSequenceName
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

' Funktion zum Abrufen der Messmethode
Public Function funcGetMethode(strTopic As Variant, Metadaten As Object) As String
    
    Select Case Left(strTopic, 3)
    Case "STD", "STA": funcGetMethode = Metadaten("Methodedaten")("MethodeSTD100")
    Case "L", "LEA": funcGetMethode = Metadaten("Methodedaten")("MethodeLeder")
    Case "ECP", "ECO": funcGetMethode = Metadaten("Methodedaten")("MethodeECO")
    Case "CAL", "CAL": funcGetMethode = Metadaten("Methodedaten")("MethodeKalibration")
    Case Else: funcGetMethode = Metadaten("Methodedaten")("MethodeSTD100")                                    ' Aktion für unbekannte Eigenschaft
    End Select

End Function

Public Function funcGetValue(prp As Properties, _
                         Optional ByVal Messung As clsValues = Nothing, _
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
        Case DataFile: varValue = obj.DataFile
        Case Info1: varValue = obj.Info1
        Case Info2: varValue = obj.Info2
        Case Info3: varValue = obj.Info3
        Case Info4: varValue = obj.Info4
        Case Wert1: varValue = obj.Wert1
        Case Wert2: varValue = obj.Wert2
        Case Wert3: varValue = obj.Wert3
        Case Wert4: varValue = obj.Wert4
        Case Messkategorie: varValue = obj.Messkategorie
        Case Else: varValue = "Somethind didn't Work"                             ' Aktion für unbekannte Eigenschaft
        End Select

    Else
        varValue = "Unknown"
    End If

    funcGetValue = varValue
    
End Function

Public Function funcIsFileOpen(filename As String) As Boolean
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

Public Function funcGetMaxPosition(colFinalSequence) As Long
    Dim m As clsValues
    Dim maxPos As Long
    
    maxPos = 0
    For i = 1 To colFinalSequence.Count
        Set m = colFinalSequence(i)
        If m.Position > maxPos Then maxPos = m.Position
    Next i
    
    funcGetMaxPosition = maxPos
End Function

Public Function funcHasIntermediataCalibration(colSequence) As Boolean

    Dim m1 As clsValues, m2 As clsValues, m3 As clsValues
    
    funcHasIntermediataCalibration = False
    
    For i = 2 To colSequence.Count - 1
        Set m1 = colSequence(i - 1)
        Set m2 = colSequence(i)
        Set m3 = colSequence(i + 1)
        
        If m1.Messkategorie = Blank _
           And m2.Messkategorie = Calibration _
           And m3.Messkategorie = Blank Then
           
           funcHasIntermediataCalibration = True
           Exit Function
        End If
    Next i
    
End Function

Public Function funcCloneObject(orig As clsValues, Index As Integer) As clsValues

    Dim Clone As New clsValues
    Dim strPropertyname As String

    For prpName = AcquisitionMethode To Messkategorie
        strPropertyname = funcGetPropertyName(prpName)
        CallByName Clone, strPropertyname, VbLet, CallByName(orig, strPropertyname, VbGet)
    Next prpName
    Clone.Index = Index

    Set funcCloneObject = Clone

End Function

Public Function funcIsArrayEmpty(arr As Variant) As Boolean

    funcIsArrayEmpty = True
    On Error Resume Next
    funcIsArrayEmpty = (LBound(arr) > UBound(arr))
    On Error GoTo 0
    
End Function

Public Function funcIsOperatorPresent(arr As Variant, strName As String) As Boolean

    On Error Resume Next
    For i = LBound(arr) To UBound(arr)
        If arr(i) Like "*" & strName & "*" Then
            funcIsOperatorPresent = True
            Exit Function
        End If
    Next i
    On Error GoTo 0
    
End Function

Function GetDropdownIndex(dropdownValue As String, listRange As Range) As Long

    Dim cell As Range
    Dim idx As Long
    
    idx = 0
    For Each cell In listRange
        If cell.value = dropdownValue Then
            GetDropdownIndex = idx
            Exit Function
        End If
        idx = idx + 1
    Next cell
    
    ' Wenn nicht gefunden, 0 zurückgeben
    GetDropdownIndex = 0
    
End Function


