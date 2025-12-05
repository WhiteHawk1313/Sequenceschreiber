VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMethodenLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public wsData As Worksheet
Public wsHauptseite As Worksheet
Public wsSequence As Worksheet

Public strMethodedaten As String
Public strExportordner As String
Public strSpeicherort As String

Public intGeräteZeile As Integer
Public intMethodenZeile As Integer
Public intZeileSequence As Integer
Public intPosition As Integer

Public arrQuellKolonne As Variant

Public dictMetadaten As Object
Public dictBatchdaten As Object
Public dictMethodedaten As Object
Public dictTrigger As Object
Public dictKolonnenposition As Object
Public dictUsage As Object

Public colProben As Collection
Public colKalibration As Collection
Public colSpezialproben As Collection
Public colRawSequence As Collection
Public colFinalSequence As Collection

Public objBlank As clsWerte
Public objGanzspalten As clsWerte

Public Sub Init(ByRef DataWS As Worksheet, ByRef HauptWS As Worksheet, ByRef SequenceWS As Worksheet, ByRef MethodFile As String)

    Set wsData = DataWS
    Set wsHauptseite = HauptWS
    Set wsSequence = SequenceWS
    
    strMethodedaten = MethodFile
    
    Set dictMetadaten = CreateObject("Scripting.Dictionary")
    Set dictBatchdaten = CreateObject("Scripting.Dictionary")
    Set dictMethodedaten = CreateObject("Scripting.Dictionary")
    Set dictTrigger = CreateObject("Scripting.Dictionary")
    Set dictKolonnenposition = CreateObject("Scripting.Dictionary")
    Set dictUsage = CreateObject("Scripting.Dictionary")
    
    Set colProben = New Collection
    Set colKalibration = New Collection
    Set colSpezialproben = New Collection
    Set colRawSequence = New Collection
    Set colFinalSequence = New Collection
    
    Set objBlank = New clsWerte
    Set objGanzspalten = New clsWerte
    
End Sub

'---Definitionen---

Public Sub setBatchdaten()

    ' Sicherstellen, dass wsData gesetzt ist
    If wsData Is Nothing Then Exit Sub
    
    With dictBatchdaten
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
    
    Set dictMetadaten("Batchdaten") = dictBatchdaten
    
End Sub
Public Sub setWertePosition()

    Dim intSpalte As Properties
    
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
        
End Sub

Public Sub setDatenWorkbook()

    Set dictMetadaten("wbDaten") = Workbooks(Dir(strMethodedaten))
    Set dictMetadaten("wsDaten") = dictMetadaten("wbDaten").Sheets(dictBatchdaten("Gerät"))
    
End Sub

Public Sub setSaveSpaces()
    
    Dim intSpalte As Integer
    
    intSpalte = dictMetadaten("wbDaten").Sheets("Geräte").Rows(2).Find("Exportordner").Column
    strExportordner = dictMetadaten("wbDaten").Sheets("Geräte").Cells(intGeräteZeile, intSpalte).value & "\"
    strSpeicherort = dictMetadaten("wbDaten").Sheets("Geräte").Cells(intGeräteZeile, intSpalte + 1).value & "\"

End Sub

Public Sub setMethodenZeile()

    arrQuellKolonne = dictMetadaten("wsDaten").Range(dictMetadaten("wsDaten").Cells(2, 2), dictMetadaten("wsDaten").Cells(2, dictMetadaten("wsDaten").Columns.Count).End(xlToLeft))
    intMethodenZeile = dictMetadaten("wsDaten").Columns(Application.Match("Methode", arrQuellKolonne, 0) + 1).Find(What:=dictBatchdaten("Methode"), LookAt:=xlWhole).Row

End Sub

Public Sub setMethodendaten()

    With dictMethodedaten
        .Add "Kalibrationsanzahl", (dictMetadaten("wsDaten").Cells(intMethodenZeile, dictMetadaten("wsDaten").Columns.Count).End(xlToLeft).Column - (Application.Match("Lösungsmittel", arrQuellKolonne, 0) + 1)) / 3
        .Add "Spezialbrobenanzahl", WorksheetFunction.CountA(dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Spezialprobe 1 Probe 1 nach Kali", arrQuellKolonne, 0) + 1).Resize(1, 6)) / 2
        .Add "Team", dictMetadaten("wsDaten").Cells(intMethodenZeile, Application.Match("Team", arrQuellKolonne, 0) + 1)
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

End Sub

Public Sub setTrigger()

    Dim zKal As Long, zBlank As Long

    zKal = dictMethodedaten("ZwischenKalibartionTrigger")
    If zKal <= 0 Then zKal = lngArbitraryBigNumber
    
    zBlank = dictMethodedaten("ZwischenBlankTrigger")
    If zBlank <= 0 Then zBlank = lngArbitraryBigNumber

    With dictTrigger
        .Add "MaxKalibration", Int(dictBatchdaten("TotalProben") / zKal)
        .Add "AnzahlProbenZwischenKalibrationen", Int(dictBatchdaten("TotalProben") / (.item("MaxKalibration") + 1))
        .Add "AnzahlProbenZwischenBlank", .item("AnzahlProbenZwischenKalibrationen") \ ((.item("AnzahlProbenZwischenKalibrationen") \ zBlank) + 1)
    End With

    Set dictMetadaten("Trigger") = dictTrigger

End Sub

Public Sub setBlindwerte()

    ' Blankwerte auslesen
    For prpName = AcquisitionMethode To Messkategorie
        If Not dictMetadaten("Kolonnenposition")(funcGetPropertyName(prpName)) = -1 Then _
           Call defSetWert(prp:=prpName, Messtyp:=Blank, Metadaten:=dictMetadaten, Blank:=objBlank)
    Next prpName
    
End Sub

Public Sub setKalibrationen()
    
    Dim objKalibration As Object
    
    For i = 1 To dictMethodedaten("Kalibrationsanzahl")
        Set objKalibration = New clsWerte
        colKalibration.Add objKalibration
        For prpName = AcquisitionMethode To Messkategorie
            If Not dictMetadaten("Kolonnenposition")(funcGetPropertyName(prpName)) = -1 Then _
               Call defSetWert(prp:=prpName, Messtyp:=Kalibration, Metadaten:=dictMetadaten, Kalibration:=colKalibration, Collectionindex:=i)
        Next prpName
    Next i
    
End Sub

Public Sub setSpezialproben()
    
    Dim objSpezialproben As Object
    
    For i = 1 To 3
        If Not dictMetadaten("wsDaten").Cells(intMethodenZeile, (i - 1) * 2 + Application.Match("Spezialprobe 1 Probe 1 nach Kali", arrQuellKolonne, 0) + 1) = "" Then
            Set objSpezialproben = New clsWerte
            colSpezialproben.Add objSpezialproben
            For prpName = AcquisitionMethode To Messkategorie
                If Not dictMetadaten("Kolonnenposition")(funcGetPropertyName(prpName)) = -1 Then _
                   Call defSetWert(prp:=prpName, Messtyp:=Spezialprobe, Metadaten:=dictMetadaten, Spezialproben:=colSpezialproben, Kalibration:=colKalibration, Collectionindex:=i)
            Next prpName
        End If
    Next i
    
End Sub

Public Sub setProben()

    Dim objProben As Object

    For i = 1 To dictBatchdaten("TotalProben")
        Set objProben = New clsWerte
        colProben.Add objProben
        For prpName = AcquisitionMethode To Messkategorie
            If Not dictMetadaten("Kolonnenposition")(funcGetPropertyName(prpName)) = -1 Then _
               Call defSetWert(prp:=prpName, Messtyp:=Sample, Metadaten:=dictMetadaten, Probe:=colProben, Kalibration:=colKalibration, Collectionindex:=i)
        Next prpName
    Next i

End Sub

Public Sub setGanzspalten()

    For prpName = Sequencename To Sequencename
        If Not dictMetadaten("Kolonnenposition")(funcGetPropertyName(prpName)) = -1 Then _
           Call defSetWert(prp:=prpName, Messtyp:=Ganzspalten, Metadaten:=dictMetadaten, Ganzspalten:=objGanzspalten)
    Next prpName
        
End Sub

'---Inserts---

Public Sub getBlank()

    intZeileSequence = intZeileSequence + 1
    colRawSequence.Add objBlank
    dictMetadaten("Trigger")("CurrentBlankTriggerCount") = 0
    
End Sub

Public Sub getKalibration(Volle_Kalibration As Boolean, Zwischenkalibration As Boolean)

    Dim intZusatzprobeTrigger As Integer
    Dim blnExtra As Boolean

    If Not dictMetadaten("Batchdaten")("TotalProben") / Int(dictMetadaten("Batchdaten")("TotalProben") / dictMetadaten("Trigger")("AnzahlProbenZwischenKalibrationen")) - dictMetadaten("Trigger")("AnzahlProbenZwischenKalibrationen") = 0 Then
        ' Berechne, wie viele Proben zwischen zwei Zwischenkalibrationen liegen sollen
        intZusatzprobeTrigger = Int(1 / (dictMetadaten("Batchdaten")("TotalProben") / Int(dictMetadaten("Batchdaten")("TotalProben") / dictMetadaten("Trigger")("AnzahlProbenZwischenKalibrationen")) - dictMetadaten("Trigger")("AnzahlProbenZwischenKalibrationen")))

        ' Überprüfe, ob eine Extraprobe eingefügt werden soll
        blnExtra = dictMetadaten("Trigger")("MaxKalibration") Mod intZusatzprobeTrigger = 0
    End If
    
    If Volle_Kalibration Then
        For i = 1 To colKalibration.Count
            intZeileSequence = intZeileSequence + 1
            colRawSequence.Add colKalibration(i)
        Next i
    Else
        intZeileSequence = intZeileSequence + 1
        colRawSequence.Add colKalibration(Round(colKalibration.Count / 2, 0))
    End If
    
    dictMetadaten("Trigger")("CurrentKalibrationTriggerCount") = IIf(blnExtra = True, -1, 0)
End Sub

Public Sub getSpezialproben()
    
    For i = 1 To colSpezialproben.Count
        intZeileSequence = intZeileSequence + 1
        colRawSequence.Add colSpezialproben(i)
    Next i

End Sub

'---Prozesse für die Positionsvergabe---
Public Sub setUpdatePosition(KategorieConst As Long, maxUsage As Long, UseLevel As Boolean)
    
    Dim Messung As clsWerte
    Dim blnDoProcess As Boolean
    Dim blnIncreaseUsage As Boolean
    Dim blnDidSomething As Boolean
    Dim blnZwischenkalibrationIncreased As Boolean
    Dim Kategorie As MessTypen
    Dim intIncreaseAmount As Integer
    
    dictUsage(Kategorie) = 0
    blnDidSomething = False
    blnZwischenkalibrationIncreased = dictUsage(Kalibration) = 0
    Kategorie = IIf(KategorieConst = Zwischenkalibration, Kalibration, KategorieConst)
    intIncreaseAmount = 0

    For i = 1 To colRawSequence.Count
        Set Messung = colRawSequence(i)
        If Messung.Messkategorie = Kategorie Then
        
            ' Bestimmen, ob Messung relevant ist oder nicht
            Select Case KategorieConst
            Case Kalibration
                blnDoProcess = Not (colRawSequence(i - 1).Messkategorie = Blank And colRawSequence(i + 1).Messkategorie = Blank)
                blnIncreaseUsage = colRawSequence(i + 1).Messkategorie = Blank
                If blnIncreaseUsage And blnDoProcess Then
                    intIncreaseAmount = Messung.Level - 1
                End If
            Case Zwischenkalibration
                blnDoProcess = colRawSequence(i - 1).Messkategorie = Blank And colRawSequence(i + 1).Messkategorie = Blank
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
        If Not dictUsage(Kategorie) = 0 And Not funcHasZwischenkalibration(colRawSequence) Then intPosition = intPosition + intIncreaseAmount + 1
    Else
        If Not dictUsage(Kategorie) = 0 And blnDidSomething Then intPosition = funcGetMaxPosition(colFinalSequence) + 1
    End If
    
End Sub


