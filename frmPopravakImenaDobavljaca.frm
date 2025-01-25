VERSION 5.00
Begin VB.Form frmPopravakImenaDobavljaca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Popravak imena dobavljaèa"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   Icon            =   "frmPopravakImenaDobavljaca.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame2 
      Height          =   35
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   5775
   End
   Begin VB.CommandButton cmdWrite 
      Caption         =   "Promijeni"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdKraj 
      Cancel          =   -1  'True
      Caption         =   "Kraj"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtImeOut 
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      Top             =   360
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   3000
      TabIndex        =   2
      Top             =   0
      Width           =   35
   End
   Begin VB.ComboBox cboImeIn 
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Novo ime:"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Ime u tabeli:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmPopravakImenaDobavljaca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Start()

If rsDobavljaci.RecordCount > 0 Then
    rsDobavljaci.MoveLast
    rsDobavljaci.MoveFirst
Else
    Exit Sub
End If
Do Until rsDobavljaci.EOF
    Me.cboImeIn.AddItem rsDobavljaci("Ime")
    rsDobavljaci.MoveNext
Loop

Me.cboImeIn.ListIndex = 0

Me.Show
End Sub



Private Sub cboImeIn_Click()
Me.txtImeOut = Me.cboImeIn

End Sub

Private Sub cmdKraj_Click()
Unload Me
End Sub

Private Sub cmdWrite_Click()
rsDobavljaci.MoveLast
rsDobavljaci.MoveFirst
Do Until rsDobavljaci.EOF
    If rsDobavljaci("Ime") = Me.cboImeIn Then
        rsDobavljaci.Edit
        rsDobavljaci("Ime") = Me.txtImeOut
        rsDobavljaci.Update
    End If
    rsDobavljaci.MoveNext
Loop

If rsModeli.RecordCount = 0 Then GoTo over
rsModeli.MoveFirst
Do Until rsModeli.EOF
    If rsModeli("Dobavljac") = Me.cboImeIn Then
        rsModeli.Edit
        rsModeli("Dobavljac") = Me.txtImeOut
        rsModeli.Update
    End If
    rsModeli.MoveNext
Loop

Dim k As MSComctlLib.ListItem
For Each k In frmGlavni.ctlListView.ListItems
    If k.SubItems(2) = Me.cboImeIn Then k.SubItems(2) = Me.txtImeOut
Next

over:
Dim i As Integer
i = Me.cboImeIn.ListIndex
Me.cboImeIn.RemoveItem i
Me.cboImeIn.AddItem Me.txtImeOut, i
Me.cboImeIn = Me.txtImeOut

Me.cboImeIn.SetFocus

End Sub

Private Sub txtImeOut_Change()
Me.cmdWrite.Enabled = (Me.cboImeIn <> Me.txtImeOut)
End Sub

