VERSION 5.00
Begin VB.Form frmTrazenje 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pretraživanje modela"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frmTrazenje.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4755
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTacno 
      Caption         =   "Traži taènu vrijednost"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.Frame fraUpDown 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   3000
      TabIndex        =   11
      Top             =   720
      Width           =   1695
      Begin VB.OptionButton optDown 
         Caption         =   "Dolje"
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optUp 
         Caption         =   "Gore"
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.TextBox txtSearchString 
      Height          =   300
      Left            =   120
      TabIndex        =   10
      Top             =   1875
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Odustani"
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Poèisti polja"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdTrazi 
      Caption         =   "Traži!"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   240
      Width           =   1575
   End
   Begin VB.Frame fraNacinTrazenja 
      Caption         =   "Polje traženja"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.OptionButton optSnabdjevac 
         Caption         =   "Option3"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton optModel 
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton optSifra 
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Po dobavljaèu"
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Po modelu"
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Po šifri"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmTrazenje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private POLJE As Integer

Private Sub cmdCancel_Click()
Me.Hide
frmGlavni.Enabled = True
frmGlavni.SetFocus
End Sub
 
Private Sub cmdClear_Click()
Me.optDown = True
Me.chkTacno.Value = 1
Me.optSifra.Value = True
Me.txtSearchString = ""
Me.txtSearchString.SetFocus
End Sub

Private Sub cmdTrazi_Click()
On Error GoTo greska
optSifra_Click
        Dim Pocetak As String
        Pocetak = Mid(frmGlavni.ctlListView.SelectedItem.Key, 3)
        rsModeli.FindFirst "ID=" & Pocetak
        
Select Case POLJE
    Case 1
            If Val(Me.txtSearchString) <> Me.txtSearchString Then
                Beep
                Exit Sub
            End If
        If Me.chkTacno.Value = 1 Then
            If Me.optDown = True Then
                Do Until rsModeli.EOF
                    rsModeli.MoveNext
                    If rsModeli("Sifra") = Me.txtSearchString Then Exit Do
                Loop
            ElseIf Me.optUp = True Then
                Do Until rsModeli.BOF
                    rsModeli.MovePrevious
                    If rsModeli("Sifra") = Me.txtSearchString Then Exit Do
                Loop
            End If
        Else
            If Me.optDown = True Then
                
                Do Until rsModeli.EOF
                    rsModeli.MoveNext
                    If InStr(rsModeli("Sifra"), Me.txtSearchString) <> 0 Then Exit Do
                Loop
                
            ElseIf Me.optUp = True Then
                Do Until rsModeli.BOF
                    rsModeli.MovePrevious
                    If InStr(rsModeli("Sifra"), Me.txtSearchString) <> 0 Then Exit Do
                Loop
            End If
        End If
        
    Case 2
        If Me.chkTacno.Value = 1 Then
            If Me.optDown = True Then
                Do Until rsModeli.EOF
                    rsModeli.MoveNext
                    If rsModeli("Model") = Me.txtSearchString Then Exit Do
                Loop
            ElseIf Me.optUp = True Then
                Do Until rsModeli.BOF
                    rsModeli.MovePrevious
                    If rsModeli("Model") = Me.txtSearchString Then Exit Do
                Loop
            End If
        Else
            If Me.optDown = True Then
                Do Until rsModeli.EOF
                    rsModeli.MoveNext
                    If InStr(rsModeli("Model"), Me.txtSearchString) <> 0 Then Exit Do
                Loop
            ElseIf Me.optUp = True Then
                Do Until rsModeli.BOF
                    rsModeli.MovePrevious
                    If InStr(rsModeli("Model"), Me.txtSearchString) <> 0 Then Exit Do
                Loop
            End If
        End If
        
    Case 3
        If Me.chkTacno.Value = 1 Then
            If Me.optDown = True Then
                Do Until rsModeli.EOF
                    rsModeli.MoveNext
                    If rsModeli("Dobavljac") = Me.txtSearchString Then Exit Do
                Loop
            ElseIf Me.optUp = True Then
                Do Until rsModeli.BOF
                    rsModeli.MovePrevious
                    If rsModeli("Dobavljac") = Me.txtSearchString Then Exit Do
                Loop
            End If
        Else
            If Me.optDown = True Then
                Do Until rsModeli.EOF
                    rsModeli.MoveNext
                    If InStr(rsModeli("Dobavljac"), Me.txtSearchString) <> 0 Then Exit Do
                Loop
            ElseIf Me.optUp = True Then
                Do Until rsModeli.BOF
                    rsModeli.MovePrevious
                    If InStr(rsModeli("Dobavljac"), Me.txtSearchString) <> 0 Then Exit Do
                Loop
            End If
        End If
        
End Select
'If POLJE <> 1 Then GoTo kraj
'        If rsModeli.NoMatch Then
'            Dim a
'            a = MsgBox("Nema više takvih podataka!", vbExclamation + vbOKOnly, "Pretraživanje modela")
'            rsModeli.FindFirst "ID=" & Pocetak
'
'        End If
GoTo kraj

greska:
If Err.Number = 3021 Then
            Dim b
            b = MsgBox("Nema više takvih podataka!", vbExclamation + vbOKOnly, "Pretraživanje modela")
            rsModeli.FindFirst "ID=" & Pocetak
Else
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End If

kraj:
            
            frmGlavni.ctlListView.SelectedItem.Selected = False
            frmGlavni.ctlListView.ListItems("ID" & rsModeli("ID")).Selected = True
            frmGlavni.ctlListView.ListItems("ID" & rsModeli("ID")).EnsureVisible

End Sub


Private Sub optModel_Click()
optSifra_Click
End Sub

Private Sub optSifra_Click()
POLJE = 1
If Me.optModel.Value Then POLJE = 2
If Me.optSnabdjevac.Value Then POLJE = 3
End Sub

Private Sub optSnabdjevac_Click()
optSifra_Click
End Sub

Private Sub txtSearchString_Change()
Me.cmdTrazi.Enabled = Trim(Me.txtSearchString) <> ""
End Sub

