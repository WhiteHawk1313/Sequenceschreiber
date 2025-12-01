VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWerte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----- Klasse für Werte -----
Private ¬prvAcquisitionMethode As String
Private ¬prvQuantmethode As String
Private ¬prvBeschriftung As String
Private ¬prvEinwaage As Double
Private ¬prvExctraktionsvolumen As Double
Private ¬prvInjektionsvolumen As Double
Private ¬prvKommentar As String
Private ¬prvKonzentration As Double
Private ¬prvPosition As Integer
Private ¬prvProduktklasse As String
Private ¬prvRack As String
Private ¬prvTyp As String
Private ¬prvVerdünnung As Double
Private ¬prvLevel As Integer
Private ¬prvSequencename As String
Private ¬prvInfo1 As String
Private ¬prvInfo2 As String
Private ¬prvInfo3 As String
Private ¬prvInfo4 As String
Private ¬prvWert1 As Double
Private ¬prvWert2 As Double
Private ¬prvWert3 As Double
Private ¬prvWert4 As Double
Private ¬prvMesskategorie As MessTypen
Private ¬prvIndex As Integer

' ----- Properties für Werte Information -----
' AcquisitionMethode Property
Public Property Get AcquisitionMethode() As String
    AcquisitionMethode = ¬prvAcquisitionMethode
End Property

Public Property Let AcquisitionMethode(ByVal value As String)
    ¬prvAcquisitionMethode = value
End Property
' Quantmethode Property
Public Property Get Quantmethode() As String
    Quantmethode = ¬prvQuantmethode
End Property

Public Property Let Quantmethode(ByVal value As String)
    ¬prvQuantmethode = value
End Property

' Beschriftung Property
Public Property Get Beschriftung() As String
    Beschriftung = ¬prvBeschriftung
End Property

Public Property Let Beschriftung(ByVal value As String)
    ¬prvBeschriftung = value
End Property

' Einwaage Property
Public Property Get Einwaage() As Double
    Einwaage = ¬prvEinwaage
End Property

Public Property Let Einwaage(ByVal value As Double)
    ¬prvEinwaage = value
End Property

' Exctraktionsvolumen Property
Public Property Get Exctraktionsvolumen() As Double
    Exctraktionsvolumen = ¬prvExctraktionsvolumen
End Property

Public Property Let Exctraktionsvolumen(ByVal value As Double)
    ¬prvExctraktionsvolumen = value
End Property

' Injektionsvolumen Property
Public Property Get Injektionsvolumen() As Double
    Injektionsvolumen = ¬prvInjektionsvolumen
End Property

Public Property Let Injektionsvolumen(ByVal value As Double)
    ¬prvInjektionsvolumen = value
End Property

' Kommentar Property
Public Property Get Kommentar() As String
    Kommentar = ¬prvKommentar
End Property

Public Property Let Kommentar(ByVal value As String)
    ¬prvKommentar = value
End Property

' Konzentration Property
Public Property Get Konzentration() As Double
    Konzentration = ¬prvKonzentration
End Property

Public Property Let Konzentration(ByVal value As Double)
    ¬prvKonzentration = value
End Property

' Position Property
Public Property Get Position() As Integer
    Position = ¬prvPosition
End Property

Public Property Let Position(ByVal value As Integer)
    ¬prvPosition = value
End Property

' Produktklasse Property
Public Property Get Produktklasse() As String
    Produktklasse = ¬prvProduktklasse
End Property

Public Property Let Produktklasse(ByVal value As String)
    ¬prvProduktklasse = value
End Property

' Rack Property
Public Property Get Rack() As String
    Rack = ¬prvRack
End Property

Public Property Let Rack(ByVal value As String)
    ¬prvRack = value
End Property

' Typ Property
Public Property Get Typ() As String
    Typ = ¬prvTyp
End Property

Public Property Let Typ(ByVal value As String)
    ¬prvTyp = value
End Property

' Verdünnung Property
Public Property Get Verdünnung() As Double
    Verdünnung = ¬prvVerdünnung
End Property

Public Property Let Verdünnung(ByVal value As Double)
    ¬prvVerdünnung = value
End Property
' Level Property
Public Property Get Level() As String
    Level = ¬prvLevel
End Property

Public Property Let Level(ByVal value As String)
    ¬prvLevel = value
End Property

' Sequencename Property
Public Property Get Sequencename() As String
    Sequencename = ¬prvSequencename
End Property

Public Property Let Sequencename(ByVal value As String)
    ¬prvSequencename = value
End Property

' Info1 Property
Public Property Get Info1() As String
    Info1 = ¬prvInfo1
End Property

Public Property Let Info1(ByVal value As String)
    ¬prvInfo1 = value
End Property

' Info2 Property
Public Property Get Info2() As String
    Info2 = ¬prvInfo2
End Property

Public Property Let Info2(ByVal value As String)
    ¬prvInfo2 = value
End Property

' Info3 Property
Public Property Get Info3() As String
    Info3 = ¬prvInfo3
End Property

Public Property Let Info3(ByVal value As String)
    ¬prvInfo3 = value
End Property

' Info4 Property
Public Property Get Info4() As String
    Info4 = ¬prvInfo4
End Property

Public Property Let Info4(ByVal value As String)
    ¬prvInfo4 = value
End Property

' Wert1 Property
Public Property Get Wert1() As Double
    Wert1 = ¬prvWert1
End Property

Public Property Let Wert1(ByVal value As Double)
    ¬prvWert1 = value
End Property

' Wert2 Property
Public Property Get Wert2() As Double
    Wert2 = ¬prvWert2
End Property

Public Property Let Wert2(ByVal value As Double)
    ¬prvWert2 = value
End Property

' Wert3 Property
Public Property Get Wert3() As Double
    Wert3 = ¬prvWert3
End Property

Public Property Let Wert3(ByVal value As Double)
    ¬prvWert3 = value
End Property

' Wert4 Property
Public Property Get Wert4() As Double
    Wert4 = ¬prvWert4
End Property

Public Property Let Wert4(ByVal value As Double)
    ¬prvWert4 = value
End Property
' Messkategorie Property
Public Property Get Messkategorie() As MessTypen
    Messkategorie = ¬prvMesskategorie
End Property

Public Property Let Messkategorie(ByVal value As MessTypen)
    ¬prvMesskategorie = value
End Property
' Messkategorie Index
Public Property Get Index() As Integer
    Index = ¬prvIndex
End Property

Public Property Let Index(ByVal value As Integer)
    ¬prvIndex = value
End Property

