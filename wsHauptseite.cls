VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsHauptseite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Worksheet_Change(ByVal Target As Range)

Dim strGerät As String
Dim strMethode As String
Dim strRackname As String
Dim strRackteile() As String
Const strMethodedaten As String = "L:\Makros\Sequenceschreiber\Daten für Sequenceschreiber.xlsx"
Dim WB As Workbook
Dim DatenWB As Workbook
Dim intZeile As Integer
Dim intMethodenZeile As Integer
Dim dblStdEinwaage As Double
Dim arrQuellKolonne As Variant

Application.EnableEvents = False
Application.ScreenUpdating = False

On Error GoTo ErrHandler

With Sheets("Data")
    .Visible = True
    strGerät = .Cells(1, 2)
    strMethode = .Cells(2, 2)
    .Visible = False
End With

If Target.Address = "$I$3" And Not Cells(3, 9) = "Methode" Then 'Methode wird geändert
    Sheets(1).Unprotect
    Workbooks.Open strMethodedaten
    Set WB = ThisWorkbook
    Set DatenWB = Workbooks(Dir(strMethodedaten))
    With DatenWB.Sheets(strGerät)
        ThisWorkbook.Sheets(1).Unprotect
        ThisWorkbook.Sheets(1).Columns(13).ClearContents
        arrQuellKolonne = .Range(.Cells(2, 2), .Cells(2, Columns.Count).End(xlToLeft))
        intMethodenZeile = .Columns(Application.Match("Methode", arrQuellKolonne, 0) + 1).Find(strMethode).Row
        ThisWorkbook.Sheets("Data").Visible = True
        ThisWorkbook.Sheets("Data").Cells(3, 2) = .Cells(intMethodenZeile, Application.Match("Topic", arrQuellKolonne, 0) + 1).value
        ThisWorkbook.Sheets("Data").Visible = False
        strRackname = .Cells(intMethodenZeile, Application.Match("Rackname", arrQuellKolonne, 0) + 1).value
        If InStr(1, strRackname, "X", vbTextCompare) = 0 Then
            ReDim strRackListe(0)
            WB.Sheets("Hauptseite").Cells(2, 13) = strRackname
        Else
            strRackteile = Split(strRackname, "X")
            For intZeile = .Cells(intMethodenZeile, Application.Match("Rack Min", arrQuellKolonne, 0) + 1) To .Cells(intMethodenZeile, Application.Match("Rack Max", arrQuellKolonne, 0) + 1)
                WB.Sheets("Hauptseite").Cells(WB.Sheets("Hauptseite").Cells(Rows.Count, 13).End(xlUp).Row + 1, 13) = strRackteile(0) & intZeile & strRackteile(1)
            Next intZeile
        End If
    End With
    DatenWB.Close savechanges:=False
    Cells(5, 9) = Cells(2, 13)
    Sheets(1).Protect
End If

If (Not Intersect(Target, Columns(3)) Is Nothing Or Target.Address = "$I$3") And Not Cells(3, 9) = "Methode" Then 'Einwaage wird geändert
    Workbooks.Open strMethodedaten
    Set WB = ThisWorkbook
    Set DatenWB = Workbooks(Dir(strMethodedaten))
    With DatenWB.Sheets(strGerät)
        arrQuellKolonne = .Range(.Cells(2, 2), .Cells(2, Columns.Count).End(xlToLeft))
        intMethodenZeile = .Columns(Application.Match("Methode", arrQuellKolonne, 0) + 1).Find(strMethode).Row
        dblStdEinwaage = .Cells(intMethodenZeile, Application.Match("Standard-einwaage", arrQuellKolonne, 0) + 1)
    End With
    For intZeile = 3 To Cells(Rows.Count, 2).End(xlUp).Row
        If Not WB.Sheets("Hauptseite").Cells(intZeile, 3) = "" Then WB.Sheets("Hauptseite").Cells(intZeile, 4) = dblStdEinwaage / WB.Sheets("Hauptseite").Cells(intZeile, 3)
    Next
    DatenWB.Close savechanges:=False
End If

ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    If Not DatenWB Is Nothing Then
        If Module.funcIsFileOpen(DatenWB.Path & DatenWB.Name) Then DatenWB.Close savechanges:=False
    End If
    MsgBox "Es ist ein Fehler aufgetreten. Bitte kontaktiere den Digital Laboratory Expert für Unterstützung.", vbExclamation + vbOKOnly, Title:="Fehler"
    Resume ExitHandler
End Sub
