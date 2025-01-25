VERSION 5.00
Begin VB.Form frmObjekat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Odabir objekta"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox cboListaObjekata 
      Height          =   315
      ItemData        =   "frmObjekat.frx":0000
      Left            =   120
      List            =   "frmObjekat.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Izaberi objekat:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmObjekat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private retVal As String


Public Function Prikazi(frmParent As Form) As String

Me.cboListaObjekata.Clear
Me.cboListaObjekata.AddItem "SVI OBJEKTI"

Dim rs As Recordset
Set rs = dbModeli.OpenRecordset("SELECT DISTINCT PROD, PRODAVNICA FROM ARHIVA")
If rs.RecordCount > 0 Then
    rs.MoveLast
    rs.MoveFirst
    Do Until rs.EOF
        Me.cboListaObjekata.AddItem rs("PROD") & " - " & rs("PRODAVNICA")
        rs.MoveNext
    Loop
End If

Me.cboListaObjekata.ListIndex = 0

retVal = ""
Me.Show 1, frmParent

Prikazi = retVal

End Function

Private Sub cmdOK_Click()
If Me.cboListaObjekata.ListIndex = 0 Then
    retVal = "*"
Else
    retVal = Left(Me.cboListaObjekata.Text, 3)
End If
Me.Hide
End Sub
