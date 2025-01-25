VERSION 5.00
Begin VB.Form frmPicPreview 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   2730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image ctlImage 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "frmPicPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.ctlImage.width = Me.ScaleWidth
Me.ctlImage.height = Me.ScaleHeight
Me.Top = frmGlavni.Top + frmGlavni.height - Me.height - 2 * frmGlavni.ctlListView.Top - frmGlavni.ctlStatusBar.height
Me.Left = frmGlavni.Left + frmGlavni.width - Me.width - frmGlavni.ctlListView.Left - 300
End Sub

Public Sub PrikaziMe()
Me.Show
KeepOnTop Me
frmGlavni.SetFocus
Me.Display
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
frmGlavni.mnuModeliPreview.Checked = False
End Sub

Public Sub Display()
    Dim fn As String, modelkey As String
    If frmGlavni.fnMultiple = 1 Then
        modelkey = frmGlavni.ctlListView.SelectedItem.Key
        fn = MyDirectory & "\pic\p" & Mid(modelkey, 3) & ".jpg"
        Me.ctlImage.Visible = False
        Me.ctlImage.Left = 0
        Me.ctlImage.Top = 0
        Me.ctlImage.width = Me.ScaleWidth
        Me.ctlImage.height = Me.ScaleHeight
        Set Me.ctlImage.Picture = Nothing
        If FileExists(fn) Then
            Dim c As New clsJPEGparser
            c.ParseJpegFile fn
            If Me.ScaleWidth / Me.ScaleHeight > c.XsizePicture / c.YsizePicture Then
                Me.ctlImage.width = Me.ScaleHeight * c.XsizePicture / c.YsizePicture
                Me.ctlImage.Left = (Me.ScaleWidth - Me.ctlImage.width) / 2
            ElseIf Me.ScaleWidth / Me.ScaleHeight < c.XsizePicture / c.YsizePicture Then
                Me.ctlImage.height = Me.ScaleWidth * c.YsizePicture / c.XsizePicture
                Me.ctlImage.Top = (Me.ScaleHeight - Me.ctlImage.height) / 2
            End If
            Set c = Nothing
            Me.ctlImage.Picture = LoadPicture(fn)
        End If
        Me.ctlImage.Visible = True
        Me.Caption = frmGlavni.ctlListView.ListItems(modelkey).SubItems(1)
    Else
        Me.ctlImage.Visible = False
        Me.ctlImage.Left = 0
        Me.ctlImage.Top = 0
        Me.ctlImage.width = Me.ScaleWidth
        Me.ctlImage.height = Me.ScaleHeight
        Set Me.ctlImage.Picture = Nothing
        Me.ctlImage.Visible = True
        Me.Caption = ""
    End If
End Sub


