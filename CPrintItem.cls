VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPrintItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarID As Long
Private mvarSifra As Long
Private mvarModel As String


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

