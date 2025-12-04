VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Workbook_Open()

Application.DisplayAlerts = False: Application.EnableEvents = False

Dim strIC As String
Dim strDroppdownListe As String: strDroppdownListe = "Methode"
Const strMethodedaten = "L:\Makros\Sequenceschreiber\Daten für Sequenceschreiber.xlsx"
Dim intKolonne As Integer, i As Integer, intlastRow As Integer
Dim objGerät As Object
Dim DatenWB As Workbook, WB As Workbook
Dim rngData As Range, rngMethodenZelle As Range

If Dir("C:\TempTest", vbDirectory) = "" Then MkDir "C:\TempTest"
'ThisWorkbook.SaveAs "C:\TempTest\" & ThisWorkbook.Name & "_TEMP"
Set WB = ActiveWorkbook

On Error GoTo ErrHandler

Sheets(1).Unprotect
Sheets(1).Cells(9, 11) = Date
Sheets(1).Cells(3, 9) = "Methode"
Workbooks.Open strMethodedaten
Set DatenWB = Workbooks(Dir(strMethodedaten))
Set objGerät = DatenWB.Sheets(1).Columns(4).Find(What:=Environ("Computername"), LookIn:=xlValues, LookAt:=xlWhole)

If objGerät Is Nothing Then
    MsgBox "Dieser Computer ist nicht bekannt!", vbCritical, "Computer"
Else
    strIC = DatenWB.Sheets(1).Cells(DatenWB.Sheets(1).Columns(4).Find(Environ("Computername")).Row, 3)

    With DatenWB.Sheets(strIC)
        ' Letzte beschriftete Zeile in Spalte B ermitteln
        intlastRow = .Cells(.Rows.Count, 2).End(xlUp).Row

        ' Bereich mit den Daten festlegen
        Set rngData = .Range(.Cells(3, 2), .Cells(intlastRow, 2))

        ' Dropdown-Liste als Textstring für die Methoden erstellen
        For Each rngMethodenZelle In rngData
            strDroppdownListe = strDroppdownListe & "," & rngMethodenZelle.value
        Next rngMethodenZelle
    End With

    With ThisWorkbook.Sheets(1).Cells(3, 9).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=strDroppdownListe
        .InCellDropdown = True
    End With
    
    ' Dropdown-Liste als Textstring für das Datum erstellen
    strDroppdownListe = Empty
    For i = 0 To 6
        strDroppdownListe = strDroppdownListe & "," & Date - i
    Next i
    With ThisWorkbook.Sheets(1).Cells(9, 11).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=strDroppdownListe
        .InCellDropdown = True
    End With

    ThisWorkbook.Sheets(1).Cells(2, 9) = strIC
End If

DatenWB.Close Savechanges:=False
ThisWorkbook.Sheets(1).Protect

On Error GoTo 0
Application.ScreenUpdating = True: Application.EnableEvents = True
End

ErrHandler:
If funcIsFileOpen(DatenWB.Path & DatenWB.Name) Then DatenWB.Close Savechanges:=False
MsgBox "Es ist ein Fehler aufgetreten. Bitte kontaktiere den Digital Laboratory Expert für Unterstützung.", vbExclamation + vbOKOnly, Title:="Error"
Application.ScreenUpdating = True: Application.EnableEvents = True

End Sub


