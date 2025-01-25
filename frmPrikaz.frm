VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrikaz 
   Caption         =   "Fotografija modela"
   ClientHeight    =   4245
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6300
   Icon            =   "frmPrikaz.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar ctlStatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   3945
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   18
            Text            =   "Šifra: "
            TextSave        =   "Šifra: "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1085
            MinWidth        =   71
            Text            =   "Model: "
            TextSave        =   "Model: "
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   71
            Text            =   "Dobavljaè: "
            TextSave        =   "Dobavljaè: "
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2328
            MinWidth        =   71
            Text            =   "Datum nabavke: "
            TextSave        =   "Datum nabavke: "
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   609
            MinWidth        =   2
            Text            =   "ID: "
            TextSave        =   "ID: "
            Object.ToolTipText     =   "Naziv fajla"
         EndProperty
      EndProperty
   End
   Begin VB.Shape shpKvadar 
      BackColor       =   &H00000000&
      Height          =   3135
      Left            =   45
      Top             =   45
      Width           =   4695
   End
   Begin VB.Image imgFoto 
      Appearance      =   0  'Flat
      Height          =   2415
      Left            =   45
      Stretch         =   -1  'True
      Top             =   45
      Width           =   3855
   End
   Begin VB.Menu mnuClose 
      Caption         =   "Zatvori (ESC)"
   End
End
Attribute VB_Name = "frmPrikaz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SlikaFile As String


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
Unload Me
End If
End Sub

Public Sub Form_Resize()
Dim slika As New clsJPEGparser
Me.imgFoto.Visible = False
On Error Resume Next
Me.imgFoto.Left = 45
Me.imgFoto.Top = 45
If Me.height < 4000 Then Me.height = 4000
If Me.width < 5000 Then Me.width = 5000
Me.shpKvadar.width = Me.ScaleWidth - (2 * Me.shpKvadar.Left)
Me.shpKvadar.height = Me.ScaleHeight - (2 * Me.shpKvadar.Top) - Me.ctlStatusBar.height
Me.imgFoto.width = Me.ScaleWidth - (2 * Me.imgFoto.Left)
Me.imgFoto.height = Me.ScaleHeight - (2 * Me.imgFoto.Top) - Me.ctlStatusBar.height
If FileExists(SlikaFile) Then
    slika.ParseJpegFile SlikaFile
    If (Me.imgFoto.width / Me.imgFoto.height) > (slika.XsizePicture / slika.YsizePicture) Then
        Me.imgFoto.width = (slika.XsizePicture / slika.YsizePicture) * Me.imgFoto.height
        Me.imgFoto.Left = (Me.ScaleWidth / 2) - (Me.imgFoto.width / 2)
    ElseIf (Me.imgFoto.height / Me.imgFoto.width) > (slika.YsizePicture / slika.XsizePicture) Then
        Me.imgFoto.height = (slika.YsizePicture / slika.XsizePicture) * Me.imgFoto.width
        Me.imgFoto.Top = ((Me.ScaleHeight - Me.ctlStatusBar.height) / 2) - (Me.imgFoto.height / 2)
    End If
End If
Me.imgFoto.Visible = True
End Sub

Private Sub mnuClose_Click()
Me.Hide
End Sub

Public Sub Prikazi(sif As String, md As String, dob As String, dat As String, img As String, ID As Long)
SlikaFile = img
Me.ctlStatusBar.Panels(1).Text = "Šifra: " & sif & " "
Me.ctlStatusBar.Panels(2).Text = "Model: " & md & " "
Me.ctlStatusBar.Panels(3).Text = "Dobavljaè: " & dob & " "
Me.ctlStatusBar.Panels(4).Text = "Datum nabavke: " & dat & " "
Me.ctlStatusBar.Panels(5).Text = "Naziv fajla: p" & ID & ".jpg "
If FileExists(img) Then
    Me.imgFoto.Picture = LoadPicture(img)
End If
Me.width = Screen.width * 3 / 4
Me.height = Screen.height * 3 / 4
Me.Show
End Sub


