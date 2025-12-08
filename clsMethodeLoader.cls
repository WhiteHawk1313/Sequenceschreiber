VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMethodeLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public wsData As Worksheet
Public wsMainPage As Worksheet
Public wsSequence As Worksheet

Public strMethodeData As String
Public strExportFolder As String
Public strSaveFolder As String

Public intEquipmentRow As Integer
Public intMethodeRow As Integer
Public intSequenceRow As Integer
Public intPosition As Integer

Public arrSourceColumn As Variant

Public dictMetaData As Object
Public dictBatchdata As Object
Public dictMethodeData As Object
Public dictTrigger As Object
Public dictColumnPosition As Object
Public dictUsage As Object

Public colSample As Collection
Public colCalibrations As Collection
Public colSpecialSamples As Collection
Public colRawSequence As Collection
Public colFinalSequence As Collection

Public objBlank As clsValues
Public objFullColumn As clsValues

Public Sub Init(ByRef DataWS As Worksheet, ByRef HauptWS As Worksheet, ByRef SequenceWS As Worksheet, ByRef MethodFile As String)

    Set wsData = DataWS
    Set wsMainPage = HauptWS
    Set wsSequence = SequenceWS
    
    strMethodeData = MethodFile
    
    Set dictMetaData = CreateObject("Scripting.Dictionary")
    Set dictBatchdata = CreateObject("Scripting.Dictionary")
    Set dictMethodeData = CreateObject("Scripting.Dictionary")
    Set dictTrigger = CreateObject("Scripting.Dictionary")
    Set dictColumnPosition = CreateObject("Scripting.Dictionary")
    Set dictUsage = CreateObject("Scripting.Dictionary")
    
    Set colSample = New Collection
    Set colCalibrations = New Collection
    Set colSpecialSamples = New Collection
    Set colRawSequence = New Collection
    Set colFinalSequence = New Collection
    
    Set objBlank = New clsValues
    Set objFullColumn = New clsValues
    
End Sub

'---Definitionen---

Public Sub setBatchdata()

    ' Sicherstellen, dass wsData gesetzt ist
    If wsData Is Nothing Then Exit Sub
    
    With dictBatchdata
        .Add "Gerät", wsData.Cells(1, 2).value
        .Add "Methode", wsData.Cells(2, 2).value
        .Add "Topic", wsData.Cells(3, 2).value
        .Add "Operator", wsData.Cells(4, 2).value
        .Add "Rack", wsData.Cells(5, 2).value
        .Add "Position", wsData.Cells(6, 2).value
        .Add "AnzahlMessungen", wsData.Cells(7, 2).value
        .Add "Datum", wsData.Cells(8, 2).value
        .Add "TotalProben", wsData.Cells(9, 2).value
    End With
    
    Set dictMetaData("Batchdaten") = dictBatchdata
    
End Sub
Public Sub setValuePosition()

    Dim intSpalte As Properties
    
    ' Positionen der Werte definieren (-1 wenn nicht verlangt)
    intEquipmentRow = dictMetaData("wbDaten").Sheets(1).Columns(4).Find(Environ("Computername")).Row
    With dictColumnPosition
        ' Werte für Sequence
        For prpName = AcquisitionMethode To Wert4
            .Add funcGetPropertyName(prpName), dictMetaData("wbDaten").Sheets("Geräte").Cells(intEquipmentRow, prpName).value
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
                 dictMetaData("wbDaten").Sheets("Geräte").Cells(intEquipmentRow, intSpalte).value
        Next prpName
    End With
    Set dictMetaData("Kolonnenposition") = dictColumnPosition
        
End Sub

Public Sub setDataWorkbook()

    Set dictMetaData("wbDaten") = Workbooks(Dir(strMethodeData))
    Set dictMetaData("wsDaten") = dictMetaData("wbDaten").Sheets(dictBatchdata("Gerät"))
    
End Sub

Public Sub setSaveSpaces()
    
    Dim intSpalte As Integer
    
    intSpalte = dictMetaData("wbDaten").Sheets("Geräte").Rows(2).Find("Exportordner").Column
    strExportFolder = dictMetaData("wbDaten").Sheets("Geräte").Cells(intEquipmentRow, intSpalte).value & "\"
    strSaveFolder = dictMetaData("wbDaten").Sheets("Geräte").Cells(intEquipmentRow, intSpalte + 1).value & "\"

End Sub

Public Sub setMethodeRow()

    arrSourceColumn = dictMetaData("wsDaten").Range(dictMetaData("wsDaten").Cells(2, 2), dictMetaData("wsDaten").Cells(2, dictMetaData("wsDaten").Columns.Count).End(xlToLeft))
    intMethodeRow = dictMetaData("wsDaten").Columns(Application.Match("Methode", arrSourceColumn, 0) + 1).Find(What:=dictBatchdata("Methode"), LookAt:=xlWhole).Row

End Sub

Public Sub setMethodeData()

    With dictMethodeData
        .Add "Kalibrationsanzahl", (dictMetaData("wsDaten").Cells(intMethodeRow, dictMetaData("wsDaten").Columns.Count).End(xlToLeft).Column - (Application.Match("Lösungsmittel", arrSourceColumn, 0) + 1)) / 3
        .Add "Spezialbrobenanzahl", WorksheetFunction.CountA(dictMetaData("wsDaten").Cells(intMethodeRow, Application.Match("Spezialprobe 1 Probe 1 nach Kali", arrSourceColumn, 0) + 1).Resize(1, 6)) / 2
        .Add "Team", dictMetaData("wsDaten").Cells(intMethodeRow, Application.Match("Team", arrSourceColumn, 0) + 1)
        .Add "MethodeSTD100", dictMetaData("wsDaten").Cells(intMethodeRow, Application.Match("Methodenname STD 100 (MUSS GENAU STIMMEN!)", arrSourceColumn, 0) + 1)
        .Add "MethodeLeder", dictMetaData("wsDaten").Cells(intMethodeRow, Application.Match("Methodenname Leder (MUSS GENAU STIMMEN!)", arrSourceColumn, 0) + 1)
        .Add "MethodeECO", dictMetaData("wsDaten").Cells(intMethodeRow, Application.Match("Methodenname Eco-Pass (MUSS GENAU STIMMEN!)", arrSourceColumn, 0) + 1)
        .Add "MethodeKalibration", dictMetaData("wsDaten").Cells(intMethodeRow, Application.Match("Methodenname Kalibration (MUSS GENAU STIMMEN!)", arrSourceColumn, 0) + 1)
        .Add "Standardeinwaage", dictMetaData("wsDaten").Cells(intMethodeRow, Application.Match("Standard-einwaage", arrSourceColumn, 0) + 1)
        .Add "Exctraktionsvolumen", dictMetaData("wsDaten").Cells(intMethodeRow, Application.Match("Exctraktionsvolumen", arrSourceColumn, 0) + 1)
        .Add "Injektionsvolumen", dictMetaData("wsDaten").Cells(intMethodeRow, Application.Match("Injektionsvolumen", arrSourceColumn, 0) + 1)
        .Add "ProbenTyp", dictMetaData("wsDaten").Cells(intMethodeRow, Application.Match("Proben Typ", arrSourceColumn, 0) + 1)
        .Add "Rackname", dictMetaData("wsDaten").Cells(intMethodeRow, Application.Match("Rackname", arrSourceColumn, 0) + 1)
        .Add "RackMin", dictMetaData("wsDaten").Cells(intMethodeRow, Application.Match("Rack Min", arrSourceColumn, 0) + 1)
        .Add "RackMax", dictMetaData("wsDaten").Cells(intMethodeRow, Application.Match("Rack Max", arrSourceColumn, 0) + 1)
        .Add "RackPositionen", dictMetaData("wsDaten").Cells(intMethodeRow, Application.Match("Maximale Anzahl Position", arrSourceColumn, 0) + 1)
        .Add "ZwischenkaliEinzel_Volle", dictMetaData("wsDaten").Cells(intMethodeRow, Application.Match("Einzel/Volle Zwischenkali", arrSourceColumn, 0) + 1)
        .Add "ZwischenkaliQC_Cal", dictMetaData("wsDaten").Cells(intMethodeRow, Application.Match("Zwischenkali als QC oder Cal", arrSourceColumn, 0) + 1)
        .Add "BlankWechsel", dictMetaData("wsDaten").Cells(intMethodeRow, Application.Match("Blank wechseln nach n Messungen", arrSourceColumn, 0) + 1)
        .Add "KalWechsel", dictMetaData("wsDaten").Cells(intMethodeRow, Application.Match("Kali wechseln nach n Messungen", arrSourceColumn, 0) + 1)
        .Add "ZwischenBlankTrigger", dictMetaData("wsDaten").Cells(intMethodeRow, Application.Match("Zwischenblank ab X Proben", arrSourceColumn, 0) + 1)
        .Add "ZwischenKalibartionTrigger", dictMetaData("wsDaten").Cells(intMethodeRow, Application.Match("Zwischenkali ab X Proben", arrSourceColumn, 0) + 1)
        .Add "ZwischenKalibartionModus", IIf(dictMetaData("wsDaten").Cells(intMethodeRow, Application.Match("Einzel/Volle Zwischenkali", arrSourceColumn, 0) + 1) = "Einzel", 1, .item("Kalibrationsanzahl"))
    End With
    Set dictMetaData("Methodedaten") = dictMethodeData

End Sub

Public Sub setTrigger()

    Dim zKal As Long, zBlank As Long

    zKal = dictMethodeData("ZwischenKalibartionTrigger")
    If zKal <= 0 Then zKal = lngArbitraryBigNumber
    
    zBlank = dictMethodeData("ZwischenBlankTrigger")
    If zBlank <= 0 Then zBlank = lngArbitraryBigNumber

    With dictTrigger
        .Add "MaxKalibration", Int(dictBatchdata("TotalProben") / zKal)
        .Add "AnzahlProbenZwischenKalibrationen", Int(dictBatchdata("TotalProben") / (.item("MaxKalibration") + 1))
        .Add "AnzahlProbenZwischenBlank", .item("AnzahlProbenZwischenKalibrationen") \ ((.item("AnzahlProbenZwischenKalibrationen") \ zBlank) + 1)
    End With

    Set dictMetaData("Trigger") = dictTrigger

End Sub

Public Sub setBlankValues()

    ' Blankwerte auslesen
    For prpName = AcquisitionMethode To Messkategorie
        If Not dictMetaData("Kolonnenposition")(funcGetPropertyName(prpName)) = 0 Then _
           Call defSetValue(prp:=prpName, MeasurementType:=Blank, Metadaten:=dictMetaData, Blank:=objBlank)
    Next prpName
    
End Sub

Public Sub setCalibrationValues()
    
    Dim objKalibration As Object
    
    For i = 1 To dictMethodeData("Kalibrationsanzahl")
        Set objKalibration = New clsValues
        colCalibrations.Add objKalibration
        For prpName = AcquisitionMethode To Messkategorie
            If Not dictMetaData("Kolonnenposition")(funcGetPropertyName(prpName)) = 0 Then _
               Call defSetValue(prp:=prpName, MeasurementType:=Calibration, Metadaten:=dictMetaData, Kalibration:=colCalibrations, Collectionindex:=i)
        Next prpName
    Next i
    
End Sub

Public Sub setSpecialSamleValues()
    
    Dim objSpezialproben As Object
    
    For i = 1 To 3
        If Not dictMetaData("wsDaten").Cells(intMethodeRow, (i - 1) * 2 + Application.Match("Spezialprobe 1 Probe 1 nach Kali", arrSourceColumn, 0) + 1) = "" Then
            Set objSpezialproben = New clsValues
            colSpecialSamples.Add objSpezialproben
            For prpName = AcquisitionMethode To Messkategorie
                If Not dictMetaData("Kolonnenposition")(funcGetPropertyName(prpName)) = 0 Then _
                   Call defSetValue(prp:=prpName, MeasurementType:=SpezialSample, Metadaten:=dictMetaData, Spezialproben:=colSpecialSamples, Kalibration:=colCalibrations, Collectionindex:=i)
            Next prpName
        End If
    Next i
    
End Sub

Public Sub setSampleValues()

    Dim objProben As Object

    For i = 1 To dictBatchdata("TotalProben")
        Set objProben = New clsValues
        colSample.Add objProben
        For prpName = AcquisitionMethode To Messkategorie
            If Not dictMetaData("Kolonnenposition")(funcGetPropertyName(prpName)) = 0 Then _
               Call defSetValue(prp:=prpName, MeasurementType:=Sample, Metadaten:=dictMetaData, Probe:=colSample, Kalibration:=colCalibrations, Collectionindex:=i)
        Next prpName
    Next i

End Sub

Public Sub setFullColumns()

    For prpName = Sequencename To Sequencename
        If Not dictMetaData("Kolonnenposition")(funcGetPropertyName(prpName)) = 0 Then _
           Call defSetValue(prp:=prpName, MeasurementType:=FullColumns, Metadaten:=dictMetaData, Ganzspalten:=objFullColumn)
    Next prpName
        
End Sub

'---Inserts---

Public Sub getBlank()

    intSequenceRow = intSequenceRow + 1
    colRawSequence.Add objBlank
    dictMetaData("Trigger")("CurrentBlankTriggerCount") = 0
    
End Sub

Public Sub getCalibration(FullCalibration As Boolean, IntermediateCalibration As Boolean)

    Dim intZusatzprobeTrigger As Integer
    Dim blnExtra As Boolean

    If Not dictMetaData("Batchdaten")("TotalProben") / Int(dictMetaData("Batchdaten")("TotalProben") / dictMetaData("Trigger")("AnzahlProbenZwischenKalibrationen")) - dictMetaData("Trigger")("AnzahlProbenZwischenKalibrationen") = 0 Then
        ' Berechne, wie viele Proben zwischen zwei Zwischenkalibrationen liegen sollen
        intZusatzprobeTrigger = Int(1 / (dictMetaData("Batchdaten")("TotalProben") / Int(dictMetaData("Batchdaten")("TotalProben") / dictMetaData("Trigger")("AnzahlProbenZwischenKalibrationen")) - dictMetaData("Trigger")("AnzahlProbenZwischenKalibrationen")))

        ' Überprüfe, ob eine Extraprobe eingefügt werden soll
        blnExtra = dictMetaData("Trigger")("MaxKalibration") Mod intZusatzprobeTrigger = 0
    End If
    
    If FullCalibration Then
        For i = 1 To colCalibrations.Count
            intSequenceRow = intSequenceRow + 1
            colRawSequence.Add colCalibrations(i)
        Next i
    Else
        intSequenceRow = intSequenceRow + 1
        colRawSequence.Add colCalibrations(Round(colCalibrations.Count / 2, 0))
    End If
    
    dictMetaData("Trigger")("CurrentKalibrationTriggerCount") = IIf(blnExtra = True, -1, 0)
End Sub

Public Sub getSpezialproben()
    
    For i = 1 To colSpecialSamples.Count
        intSequenceRow = intSequenceRow + 1
        colRawSequence.Add colSpecialSamples(i)
    Next i

End Sub

'---Prozesse für die Positionsvergabe---
Public Sub setUpdatePosition(KategorieConst As Long, maxUsage As Long, UseLevel As Boolean)
    
    Dim Messung As clsValues
    Dim blnDoProcess As Boolean
    Dim blnIncreaseUsage As Boolean
    Dim blnDidSomething As Boolean
    Dim blnZwischenkalibrationIncreased As Boolean
    Dim Kategorie As MeasurementTypes
    Dim intIncreaseAmount As Integer
    
    dictUsage(Kategorie) = 0
    blnDidSomething = False
    blnZwischenkalibrationIncreased = dictUsage(Calibration) = 0
    Kategorie = IIf(KategorieConst = IntermediateCalibration, Calibration, KategorieConst)
    intIncreaseAmount = 0

    For i = 1 To colRawSequence.Count
        Set Messung = colRawSequence(i)
        If Messung.Messkategorie = Kategorie Then
        
            ' Bestimmen, ob Messung relevant ist oder nicht
            Select Case KategorieConst
            Case Calibration
                blnDoProcess = Not (colRawSequence(i - 1).Messkategorie = Blank And colRawSequence(i + 1).Messkategorie = Blank)
                blnIncreaseUsage = colRawSequence(i + 1).Messkategorie = Blank
                If blnIncreaseUsage And blnDoProcess Then
                    intIncreaseAmount = Messung.Level - 1
                End If
            Case IntermediateCalibration
                blnDoProcess = colRawSequence(i - 1).Messkategorie = Blank And colRawSequence(i + 1).Messkategorie = Blank
                blnIncreaseUsage = True
            Case Else
                blnDoProcess = True
                blnIncreaseUsage = True
            End Select
            
            ' Messung in Sequence schreiben
            If blnDoProcess Then
                blnDidSomething = True
                Messung.Position = intPosition + IIf(UseLevel Or (Not blnZwischenkalibrationIncreased And KategorieConst = IntermediateCalibration), Messung.Level - 1, 0)
                colFinalSequence.Add funcCloneObject(Messung, i)
                
                ' Nutzung der Messung erhöhen
                If blnIncreaseUsage Then
                    dictUsage(Kategorie) = dictUsage(Kategorie) + 1
                    
                    ' nächste freie Position und nutzung zurücksetzen
                    If dictUsage(Kategorie) = maxUsage Then
                        If KategorieConst = IntermediateCalibration Then
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
    If KategorieConst = Calibration Then
        If Not dictUsage(Kategorie) = 0 And Not funcHasIntermediataCalibration(colRawSequence) Then intPosition = intPosition + intIncreaseAmount + 1
    Else
        If Not dictUsage(Kategorie) = 0 And blnDidSomething Then intPosition = funcGetMaxPosition(colFinalSequence) + 1
    End If
    
End Sub


