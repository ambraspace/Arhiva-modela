VERSION 5.00
Begin VB.Form frmReplaceString 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zamjena znakovnog niza"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   Icon            =   "frmReplaceString.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   4395
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Kraj"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdOcisti 
      Caption         =   "Oèisti polja"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdZamijeni 
      Caption         =   "Zamijeni"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox chkCitavoPolje 
      Caption         =   "Traži èitavo polje"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox txtNovi 
      Height          =   285
      Left            =   120
      MaxLength       =   30
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox txtStari 
      Height          =   285
      Left            =   120
      MaxLength       =   30
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Zamijeni ga sa:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Traži znakovni niz:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmReplaceString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
Unload Me
End Sub
Public Sub Zamijeni()
Me.txtStari = frmGlavni.ctlListView.SelectedItem.SubItems(1)
Me.Show
End Sub

Private Sub cmdOcisti_Click()
Me.txtNovi = ""
Me.txtStari = ""
Me.chkCitavoPolje = 0
Me.txtStari.SetFocus
End Sub

Private Sub cmdZamijeni_Click()

Select Case Me.chkCitavoPolje
    Case 1
        rsModeli.MoveFirst
        Do Until rsModeli.EOF
            If rsModeli("Model") = Me.txtStari Then
                frmGlavni.ctlListView.ListItems("ID" & rsModeli("ID")).SubItems(1) = Replace(frmGlavni.ctlListView.ListItems("ID" & rsModeli("ID")).SubItems(1), Me.txtStari, Me.txtNovi)
                rsModeli.Edit
                rsModeli("Model") = Replace(rsModeli("Model"), Me.txtStari, Me.txtNovi)
                rsModeli.Update
            End If
            rsModeli.MoveNext
        Loop
    Case 0
        rsModeli.MoveFirst
        Do Until rsModeli.EOF
            If InStr(rsModeli("Model"), Me.txtStari) <> 0 Then
                frmGlavni.ctlListView.ListItems("ID" & rsModeli("ID")).SubItems(1) = Replace(frmGlavni.ctlListView.ListItems("ID" & rsModeli("ID")).SubItems(1), Me.txtStari, Me.txtNovi)
                rsModeli.Edit
                rsModeli("Model") = Replace(rsModeli("Model"), Me.txtStari, Me.txtNovi)
                rsModeli.Update
            End If
            rsModeli.MoveNext
        Loop
End Select

End Sub

