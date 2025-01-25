VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEditModel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "xxx Modela"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   Icon            =   "frmEditModel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDatum 
      Height          =   315
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CheckBox chkBrisati 
      Alignment       =   1  'Right Justify
      Caption         =   "Brisati izvornu fotografiju"
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2400
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog ctlCommDlg 
      Left            =   3600
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Odustani"
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "xxx model"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdOcisti 
      Caption         =   "Poèisti polja"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Frame fraSeparator 
      Height          =   50
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Width           =   7215
   End
   Begin VB.ComboBox txtSnabdjevac 
      DataSource      =   "ctlDataCtrl"
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdSlika 
      Caption         =   "xxx sliku"
      Height          =   375
      Left            =   4500
      TabIndex        =   4
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtModel 
      Height          =   315
      Left            =   1560
      MaxLength       =   30
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtSifra 
      Height          =   315
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image picFoto 
      BorderStyle     =   1  'Fixed Single
      Height          =   2535
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lblDatumNabavke 
      Alignment       =   1  'Right Justify
      Caption         =   "Datum nabavke:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1845
      Width           =   1335
   End
   Begin VB.Label lblSnabdjevac 
      Alignment       =   1  'Right Justify
      Caption         =   "Dobavljaè:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1365
      Width           =   1335
   End
   Begin VB.Label lblModel 
      Alignment       =   1  'Right Justify
      Caption         =   "Model:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   885
      Width           =   1335
   End
   Begin VB.Label lblSifra 
      Alignment       =   1  'Right Justify
      Caption         =   "Šifra:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   405
      Width           =   1335
   End
End
Attribute VB_Name = "frmEditModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const ADD_MODE = 1
Const EDIT_MODE = 2
Private iMode As Integer
Dim bDirty As Boolean


Public Sub DodajModel()
    iMode = ADD_MODE
    bDirty = False
    Me.chkBrisati.Visible = True
    If BrisatiIzvornuFotografiju Then
        Me.chkBrisati.Value = 1
    Else
        Me.chkBrisati.Value = 0
    End If
    
    Me.Caption = "Dodavanje modela"
    Me.cmdSlika.Caption = "Dodaj fotografiju"
    Me.cmdApply.Caption = "Dodaj model"
    Dim sModel As String, sDobavljac As String
    sModel = Me.txtModel
    sDobavljac = Me.txtSnabdjevac
    cmdOcisti_Click
    Me.txtDatum = Format$(PosljednjiDatum, "dd-MM-yy")
    
    Me.txtSnabdjevac.Clear
    DobavljaciCollect
    
    Me.txtModel = sModel
    Me.txtSnabdjevac = sDobavljac

    Me.Show
End Sub

Public Sub IzmijeniModel()
    
    iMode = EDIT_MODE
    bDirty = False
    
    Dim FileImage As New clsJPEGparser
    
    Me.chkBrisati.Visible = False

    Me.txtSnabdjevac.Clear
    DobavljaciCollect

    Me.Caption = "Izmjena modela"
    Me.cmdSlika.Caption = "Izmijeni fotografiju"
    Me.cmdApply.Caption = "Izmijeni model"
    
    If Right(frmGlavni.ctlListView.SelectedItem.SubItems(3), 1) = "." Then
        Me.txtDatum = Format(Left(frmGlavni.ctlListView.SelectedItem.SubItems(3), Len(frmGlavni.ctlListView.SelectedItem.SubItems(3)) - 1), "dd-MM-yy")
    Else
        Me.txtDatum = Format(frmGlavni.ctlListView.SelectedItem.SubItems(3), "dd-MM-yy")
    End If
    Me.txtModel = frmGlavni.ctlListView.SelectedItem.SubItems(1)
    Me.txtSifra = frmGlavni.ctlListView.SelectedItem
    Me.txtSnabdjevac.Text = frmGlavni.ctlListView.SelectedItem.SubItems(2)
    If FileExists(MyDirectory & "\pic\p" & Mid(frmGlavni.ctlListView.SelectedItem.Key, 3) & ".jpg") Then
        Me.ctlCommDlg.Filename = MyDirectory & "\pic\p" & Mid(frmGlavni.ctlListView.SelectedItem.Key, 3) & ".jpg"
    Else
        Me.ctlCommDlg.InitDir = CurrentFolder
    End If
    If FileExists(MyDirectory & "\pic\p" & Mid(frmGlavni.ctlListView.SelectedItem.Key, 3) & ".jpg") Then
        Me.picFoto.Picture = LoadPicture(MyDirectory & "\pic\p" & Mid(frmGlavni.ctlListView.SelectedItem.Key, 3) & ".jpg")
                
            Me.picFoto.Height = 2535
            Me.picFoto.Width = 3735
            Me.picFoto.Top = 120
            Me.picFoto.Left = 3600
    
        FileImage.ParseJpegFile MyDirectory & "\pic\p" & Mid(frmGlavni.ctlListView.SelectedItem.Key, 3) & ".jpg"
        
        If (Me.picFoto.Width / Me.picFoto.Height) > (FileImage.XsizePicture / FileImage.YsizePicture) Then
            Me.picFoto.Width = Int((FileImage.XsizePicture / FileImage.YsizePicture) * 2535)
            Me.picFoto.Left = Round(5467.5 - (Me.picFoto.Width / 2), 0)
        ElseIf (Me.picFoto.Height / Me.picFoto.Width) > (FileImage.YsizePicture / FileImage.XsizePicture) Then
            Me.picFoto.Height = Int((FileImage.YsizePicture / FileImage.XsizePicture) * 3735)
            Me.picFoto.Top = Round(1387.5 - (Me.picFoto.Height / 2), 0)
        End If

    End If

    Me.Show
    
End Sub
    
Private Sub chkBrisati_Click()
BrisatiIzvornuFotografiju = False
If Me.chkBrisati.Value = 1 Then BrisatiIzvornuFotografiju = True
End Sub

Private Sub cmdApply_Click()
    
Dim nd As String
Dim odgovor As Integer
Dim NoviDobavljac As Boolean

If Trim(Me.txtSnabdjevac) = "" Then
    Me.txtSnabdjevac = Trim(Me.txtSnabdjevac)
    Exit Sub
End If
    
NoviDobavljac = True
Dim i As Integer
For i = 1 To Me.txtSnabdjevac.ListCount
    If Me.txtSnabdjevac = Me.txtSnabdjevac.List(i - 1) Then
        NoviDobavljac = False
        Exit For
    End If
Next

If NoviDobavljac Then
    odgovor = MsgBox("Dobavljaè kojeg ste naveli je nepoznat!" & Chr(13) & "Želite li da ga unesete u bazu podataka?", vbYesNo + vbQuestion, "Novi dobavljaè")
    Select Case odgovor
       Case vbYes
            nd = Me.txtSnabdjevac
            rsDobavljaci.AddNew
            rsDobavljaci("Ime") = nd
            rsDobavljaci.Update
            Me.txtSnabdjevac.Clear
            DobavljaciCollect
            Me.txtSnabdjevac = nd
            Me.txtSnabdjevac.SetFocus
            
       Case vbNo
            Me.txtSnabdjevac = ""
            Me.txtSnabdjevac.SetFocus
            Exit Sub
    End Select
End If

    'On Error GoTo greska
    
    Dim id As Long

    Select Case iMode
        Case 1
            If Not IsDate(Me.txtDatum) And Me.txtDatum <> "" Then
                MsgBox "Unesite pravilan datum oblika 'dd-MM-yy'!"
                Exit Sub
            End If
            
            rsModeli.AddNew
            rsModeli("Sifra") = Val(Me.txtSifra)
            rsModeli("Model") = Me.txtModel
            rsModeli("Dobavljac") = Me.txtSnabdjevac
            If Me.txtDatum = "" Then
                rsModeli("Datum nabavke") = Null
            Else
                rsModeli("Datum nabavke") = Format(Me.txtDatum, "yyyy-MM-dd")
            End If
            PosljednjiDatum = Me.txtDatum
            id = rsModeli("ID")
            rsModeli.Update
            If FileExists(Me.ctlCommDlg.Filename) Then
                FileCopy Me.ctlCommDlg.Filename, MyDirectory & "\pic\p" & id & ".jpg"
                If Me.chkBrisati.Value = 1 Then
                    Kill Me.ctlCommDlg.Filename
                End If
            End If
            If InStr(Me.ctlCommDlg.Filename, "\") <> 0 Then
                CurrentFolder = Left(Me.ctlCommDlg.Filename, InStrRev(Me.ctlCommDlg.Filename, "\") - 1)
            End If
        Case 2
            rsModeli.FindFirst "ID=" & Mid(frmGlavni.ctlListView.SelectedItem.Key, 3)
            rsModeli.Edit
            rsModeli("Sifra") = Me.txtSifra
            rsModeli("Model") = Me.txtModel
            rsModeli("Dobavljac") = Me.txtSnabdjevac
            If Trim(Me.txtDatum) = "" Then
                rsModeli("Datum nabavke") = Null
            Else
                rsModeli("Datum nabavke") = Format(Trim(Me.txtDatum), "d. M. yyyy")
            End If
            id = Val(rsModeli("ID"))
            rsModeli.Update
            If FileExists(Me.ctlCommDlg.Filename) Then
                If Me.ctlCommDlg.Filename <> MyDirectory & "\pic\p" & id & ".jpg" Then
                    FileCopy Me.ctlCommDlg.Filename, MyDirectory & "\pic\p" & id & ".jpg"
                End If
            Else
                If FileExists(MyDirectory & "\pic\p" & id & ".jpg") Then
                    Kill MyDirectory & "\pic\p" & id & ".jpg"
                End If
            End If
    End Select
    
    frmGlavni.ValidateStatusBar
    
    Me.Hide
    
    If iMode = ADD_MODE Then
        frmGlavni.ctlListView.ListItems.Add , "ID" & id, Me.txtSifra
    ElseIf iMode = EDIT_MODE Then
        frmGlavni.ctlListView.ListItems("ID" & id).Text = Me.txtSifra
    End If
    frmGlavni.ctlListView.ListItems("ID" & id).SubItems(1) = Me.txtModel
    frmGlavni.ctlListView.ListItems("ID" & id).SubItems(2) = Me.txtSnabdjevac
    frmGlavni.ctlListView.ListItems("ID" & id).SubItems(3) = Format(Trim(Me.txtDatum), "d. M. yyyy.")
    frmGlavni.ctlListView.ListItems("ID" & id).EnsureVisible
    
    frmGlavni.ctlListView.SelectedItem.Selected = False
    frmGlavni.ctlListView.SelectedItem = frmGlavni.ctlListView.ListItems("ID" & id)
    frmGlavni.ctlListView_ItemClick frmGlavni.ctlListView.ListItems("ID" & id)
    GoTo kraj:
    
greska:

If Err.Number = 53 Or Err.Number = 75 Or Err.Number = 70 Then
    Resume Next
Else
    Err.Raise Err.Number
End If

kraj:
frmGlavni.Enabled = True
frmGlavni.SetFocus
End Sub

Private Sub cmdCancel_Click()
Me.Hide
frmGlavni.Enabled = True
frmGlavni.SetFocus
End Sub

Private Sub cmdOcisti_Click()
    Me.txtDatum = ""
    Me.txtModel = ""
    Me.txtSifra = ""
    Me.txtSnabdjevac = ""
    Set Me.picFoto.Picture = Nothing
    Me.ctlCommDlg.Filename = ""
    Me.ctlCommDlg.InitDir = CurrentFolder
End Sub

Private Sub cmdSlika_Click()
    On Error GoTo greska:
    Dim FileImageInfo As New clsJPEGparser
    Dim sTMP As String
    sTMP = Me.ctlCommDlg.Filename
    
    Me.ctlCommDlg.CancelError = True
    Me.ctlCommDlg.DialogTitle = "Naði fotografiju"
    Me.ctlCommDlg.Filter = "JPEG fotografije|*.jpg;*.jpe;*.jpeg"
    Me.ctlCommDlg.Flags = cdlOFNExplorer + cdlOFNFileMustExist + cdlOFNNoReadOnlyReturn + cdlOFNPathMustExist + cdlOFNShareAware + cdlOFNHideReadOnly
    Me.ctlCommDlg.ShowOpen
    If sTMP <> Me.ctlCommDlg.Filename Then bDirty = True
    Me.picFoto.Visible = False
    Me.picFoto.Picture = LoadPicture(Me.ctlCommDlg.Filename)
        
            Me.picFoto.Height = 2535
            Me.picFoto.Width = 3735
            Me.picFoto.Top = 120
            Me.picFoto.Left = 3600
            
        FileImageInfo.ParseJpegFile Me.ctlCommDlg.Filename
        
        If (Me.picFoto.Width / Me.picFoto.Height) > (FileImageInfo.XsizePicture / FileImageInfo.YsizePicture) Then
            Me.picFoto.Width = Int((FileImageInfo.XsizePicture / FileImageInfo.YsizePicture) * 2535)
            Me.picFoto.Left = Round(5467.5 - (Me.picFoto.Width / 2), 0)
        ElseIf (Me.picFoto.Height / Me.picFoto.Width) > (FileImageInfo.YsizePicture / FileImageInfo.XsizePicture) Then
            Me.picFoto.Height = Int((FileImageInfo.YsizePicture / FileImageInfo.XsizePicture) * 3735)
            Me.picFoto.Top = Round(1387.5 - (Me.picFoto.Height / 2), 0)
        End If
Me.picFoto.Visible = True
GoTo kraj

greska:
If Err.Number <> 32755 Then
    Err.Raise Err.Number, Err.Source, Err.Description
End If

kraj:
End Sub

Private Sub txtDatum_Change()
bDirty = True
If Trim(Me.txtDatum) <> "" Then
Me.cmdApply.Enabled = IsDate(Me.txtDatum)
Else
Me.txtDatum.Enabled = True
End If
End Sub

Private Sub txtDatum_GotFocus()
Me.txtDatum.SelStart = 0
Me.txtDatum.SelLength = Len(Me.txtDatum)
End Sub


Private Sub txtDatum_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 Then Exit Sub
If Len(Me.txtDatum) = 2 Or Len(Me.txtDatum) = 5 Then
    Me.txtDatum = Me.txtDatum & "-"
    Me.txtDatum.SelStart = Len(Me.txtDatum)
End If
End Sub

Private Sub txtModel_Change()
bDirty = True
Me.cmdApply.Enabled = Trim(Me.txtModel) <> ""
End Sub

Private Sub txtModel_GotFocus()
Me.txtModel.SelStart = 0
Me.txtModel.SelLength = Len(Me.txtModel)
End Sub

Private Sub txtSifra_Change()

bDirty = True
Me.cmdApply.Enabled = Trim(Me.txtSifra) <> ""
End Sub

Private Sub txtSifra_GotFocus()
Me.txtSifra.SelStart = 0
Me.txtSifra.SelLength = Len(Me.txtSifra)
End Sub

Private Sub DobavljaciCollect()
If rsDobavljaci.RecordCount > 0 Then
    rsDobavljaci.MoveLast
    rsDobavljaci.MoveFirst
End If
Do Until rsDobavljaci.EOF
Me.txtSnabdjevac.AddItem rsDobavljaci("Ime")
rsDobavljaci.MoveNext
Loop
End Sub


Private Sub txtSifra_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtSnabdjevac_Change()
bDirty = True
End Sub

Private Sub txtSnabdjevac_KeyPress(KeyAscii As Integer)
If Len(Me.txtSnabdjevac) = 20 And KeyAscii <> vbKeyBack Then
KeyAscii = 0
End If
End Sub

Private Sub txtSnabdjevac_LostFocus()
Dim nd As String
Dim odgovor As Integer
Dim NoviDobavljac As Boolean

If Trim(Me.txtSnabdjevac) = "" Then
    Me.txtSnabdjevac = Trim(Me.txtSnabdjevac)
    Exit Sub
End If
    
NoviDobavljac = True
Dim i As Integer
For i = 1 To Me.txtSnabdjevac.ListCount
    If Me.txtSnabdjevac = Me.txtSnabdjevac.List(i - 1) Then
        NoviDobavljac = False
        Exit For
    End If
Next

If NoviDobavljac Then
    odgovor = MsgBox("Dobavljaè kojeg ste naveli je nepoznat!" & Chr(13) & "Želite li da ga unesete u bazu podataka?", vbYesNo + vbQuestion, "Novi dobavljaè")
    Select Case odgovor
       Case vbYes
            nd = Me.txtSnabdjevac
            rsDobavljaci.AddNew
            rsDobavljaci("Ime") = nd
            rsDobavljaci.Update
            Me.txtSnabdjevac.Clear
            DobavljaciCollect
            Me.txtSnabdjevac = nd
            Me.txtSnabdjevac.SetFocus
            
       Case vbNo
            Me.txtSnabdjevac = ""
            Me.txtSnabdjevac.SetFocus
    End Select
End If

End Sub

