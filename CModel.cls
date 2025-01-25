VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarID As Long
Private mvarSifra As Long
Private mvarModel As String
Private mvarDobavljac As String
Private mvarDatum As Date


Public Property Get ID() As Long
ID = mvarID
End Property

Public Property Let ID(iInput As Long)
mvarID = iInput
End Property

Public Property Get Sifra() As Long
Sifra = mvarSifra
End Property

Public Property Let Sifra(iInput As Long)
mvarSifra = iInput
End Property

Public Property Get Model() As String
Model = mvarModel
End Property

Public Property Let Model(iInput As String)
mvarModel = iInput
End Property

Public Property Get Dobavljac() As String
Dobavljac = mvarDobavljac
End Property

Public Property Let Dobavljac(iInput As String)
mvarDobavljac = iInput
End Property

Public Property Get Datum() As Date
Datum = mvarDatum
End Property

Public Property Let Datum(iInput As Date)
mvarDatum = iInput
End Property

Public Property Get IsValid() As Boolean
Dim bTMP As Boolean
bTMP = True
If Val(mvarSifra) < 1 Or Val(mvarSifra) > 99999999 Then bTMP = False
If fnRemoveSpaces(mvarModel) = "" Then bTMP = False
If fnRemoveSpaces(Format(mvarDatum, "d M yyyy")) <> "" Then
    If Not IsDate(mvarDatum) Then bTMP = False
End If
IsValid = bTMP
End Property

Public Property Get Foto() As String
Foto = MyDirectory & "\pic\p" & Val(mvarID) & ".jpg"
End Property


