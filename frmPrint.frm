VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Štampanje"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkOddEven 
      Caption         =   "Prvo sve neparne strane"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4560
      TabIndex        =   20
      Top             =   3360
      Width           =   2055
   End
   Begin VB.PictureBox ctlPicture 
      Height          =   495
      Left            =   2760
      ScaleHeight     =   435
      ScaleWidth      =   1395
      TabIndex        =   31
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame fraMargine 
      Caption         =   "Margine (mm)"
      Height          =   2775
      Left            =   4560
      TabIndex        =   7
      Top             =   120
      Width           =   2055
      Begin VB.CheckBox chkMirrorLR 
         Caption         =   "Suprotno na parnim stranama"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   13
         ToolTipText     =   "Ukoliko štampate obostrano"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CheckBox chkMirrorUD 
         Caption         =   "Suprotno na parnim stranama"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   10
         ToolTipText     =   "Ukoliko štampate obostrano"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox MarginRight 
         Height          =   285
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   12
         Text            =   "10"
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox MarginLeft 
         Height          =   285
         Left            =   240
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "15"
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox MarginDown 
         Height          =   285
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "10"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox MarginUp 
         Height          =   285
         Left            =   240
         MaxLength       =   2
         TabIndex        =   8
         Text            =   "10"
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Desna:"
         Height          =   255
         Left            =   1200
         TabIndex        =   30
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Lijeva:"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Donja:"
         Height          =   255
         Left            =   1200
         TabIndex        =   28
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Gornja:"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CheckBox chkPageNums 
      Caption         =   "Štampaj brojeve stranica"
      Height          =   255
      Left            =   4560
      TabIndex        =   19
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Odustani"
      Height          =   375
      Left            =   4560
      TabIndex        =   22
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Štampaj"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Frame fraRange 
      Caption         =   "Štampaj"
      Height          =   1575
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   4335
      Begin VB.OptionButton optModels 
         Caption         =   "Modeli:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         ToolTipText     =   "Štampa odreðene modele iz arhive"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtModels 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   18
         ToolTipText     =   "npr. 7,8-15,20 (brojevi se odnose na lokacije modela)"
         Top             =   1060
         Width           =   3015
      End
      Begin VB.OptionButton optSelected 
         Caption         =   "Oznaèeno"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         ToolTipText     =   "Štampa odabrane modele iz arhive"
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton optAll 
         Caption         =   "Sve modele"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         ToolTipText     =   "Štampa sve modele iz arhive"
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame fraOpcije 
      Caption         =   "Opcije"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.ComboBox cboPageSize 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   690
         Width           =   2250
      End
      Begin VB.OptionButton optLandscape 
         Caption         =   "Horizontalno"
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optPortrait 
         Caption         =   "Uspravno"
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   1080
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox txtCols 
         Height          =   285
         Left            =   3315
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "3"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtRows 
         Height          =   285
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "5"
         Top             =   1440
         Width           =   375
      End
      Begin VB.ComboBox cboPrinterList 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   310
         Width           =   2655
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Velièina stranice:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   1230
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Broj kolona:"
         Height          =   255
         Left            =   2235
         TabIndex        =   25
         Top             =   1470
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Broj redova:"
         Height          =   255
         Left            =   255
         TabIndex        =   24
         Top             =   1470
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Štampaè:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub IspisPrikaz()
    If frmGlavni.fnMultiple = 0 Then
        Me.optSelected.Enabled = False
    Else
        Me.optSelected.Value = True
    End If
    
    Me.Show
    KeepOnTop Me
    ValidatePrintButton
End Sub

Private Sub cboPrinterList_Click()
ValidatePrintButton
End Sub

Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdPrint_Click()
Me.cmdPrint.Enabled = False
Me.MousePointer = 11

Dim KolekcijaZaStampu As New Collection
Dim iPrintOption As Integer

iPrintOption = 1
If Me.optSelected Then
    iPrintOption = 2
ElseIf Me.optModels Then
    iPrintOption = 3
End If


Dim c As CPrintItem

Select Case iPrintOption
Dim t As MSComctlLib.ListItem
    Case 1
        For Each t In frmGlavni.ctlListView.ListItems
            Set c = New CPrintItem
            c.ID = Mid(t.Key, 3)
            c.Sifra = Val(t.Text)
            c.Model = t.SubItems(1)
            KolekcijaZaStampu.Add c
            Set c = Nothing
        Next
    Case 2
        For Each t In frmGlavni.ctlListView.ListItems
            If t.Selected Then
                Set c = New CPrintItem
                c.ID = Mid(t.Key, 3)
                c.Sifra = Val(t.Text)
                c.Model = t.SubItems(1)
                KolekcijaZaStampu.Add c
                Set c = Nothing
            End If
        Next
    Case 3
        Dim str1 As String, str2 As String
        Dim i As Integer
        For i = 1 To Len(Me.txtModels)
            If Mid(Me.txtModels, i, 1) = "," Then
                If Val(str1) <= rsModeli.RecordCount Then
                    Set c = New CPrintItem
                    c.ID = Mid(frmGlavni.ctlListView.ListItems(Val(str1)).Key, 3)
                    c.Sifra = Val(frmGlavni.ctlListView.ListItems(Val(str1)).Text)
                    c.Model = frmGlavni.ctlListView.ListItems(Val(str1)).SubItems(1)
                    KolekcijaZaStampu.Add c
                    Set c = Nothing
                    str1 = ""
                    GoTo again
                End If

            ElseIf Mid(Me.txtModels, i, 1) = "-" Then
                Dim j As Integer
                For j = i + 1 To Len(Me.txtModels)
                    If Mid(Me.txtModels, j, 1) <> "," Then
                        str2 = str2 & Mid(Me.txtModels, j, 1)
                    Else
                        Exit For
                    End If
                Next
                If Val(str1) < Val(str2) Then
                    Dim k As Integer
                    For k = Val(str1) To Val(str2)
                        If k <= rsModeli.RecordCount Then
                            Set c = New CPrintItem
                            c.ID = Mid(frmGlavni.ctlListView.ListItems(k).Key, 3)
                            c.Sifra = Val(frmGlavni.ctlListView.ListItems(k).Text)
                            c.Model = frmGlavni.ctlListView.ListItems(k).SubItems(1)
                            KolekcijaZaStampu.Add c
                            Set c = Nothing
                        End If
                    Next
                Else
                    For k = Val(str1) To Val(str2) Step -1
                        If k <= rsModeli.RecordCount Then
                            Set c = New CPrintItem
                            c.ID = Mid(frmGlavni.ctlListView.ListItems(k).Key, 3)
                            c.Sifra = Val(frmGlavni.ctlListView.ListItems(k).Text)
                            c.Model = frmGlavni.ctlListView.ListItems(k).SubItems(1)
                            KolekcijaZaStampu.Add c
                            Set c = Nothing
                        End If
                    Next
                End If
                i = i + Len(str2) + 1
                str1 = ""
                str2 = ""

            Else
                str1 = str1 & Mid(Me.txtModels, i, 1)
            End If
again:
        Next
    If str1 <> "" Then
        If Val(str1) <= rsModeli.RecordCount Then
            Set c = New CPrintItem
            c.ID = Mid(frmGlavni.ctlListView.ListItems(Val(str1)).Key, 3)
            c.Sifra = Val(frmGlavni.ctlListView.ListItems(Val(str1)).Text)
            c.Model = frmGlavni.ctlListView.ListItems(Val(str1)).SubItems(1)
            KolekcijaZaStampu.Add c
            Set c = Nothing
            str1 = ""

        End If
    End If

End Select

' PODEŠAVANJE ŠTAMPAÈA

Printer.FontName = "Arial CE"
Printer.ForeColor = RGB(0, 0, 0)
Printer.FontSize = 8
Printer.FillStyle = 1
Printer.FontTransparent = False
Printer.PrintQuality = vbPRPQHigh
Printer.ScaleMode = vbMillimeters
Printer.PaperSize = 9


Dim iOrientation As Integer
Dim iColNum As Integer, iRowNum As Integer
Dim iMarginU As Integer, iMarginL As Integer
Dim bMirrorLR As Boolean, bMirrorUD As Boolean
Dim bPageNums As Boolean, bBooklet As Boolean
Dim sPaperWidth As Single, sPaperHeight As Single
Dim sPageW As Single, sPageH As Single

iOrientation = 1
If Me.optLandscape Then iOrientation = 2

Printer.Orientation = iOrientation

If iOrientation = 1 Then
    sPaperWidth = 210
    sPaperHeight = 297
Else
    sPaperWidth = 297
    sPaperHeight = 210
End If

iColNum = Me.txtCols
iRowNum = Me.txtRows

iMarginU = Me.MarginUp - (sPaperHeight - Int(Printer.ScaleHeight)) / 2
iMarginL = Me.MarginLeft - (sPaperWidth - Int(Printer.ScaleWidth)) / 2

sPageW = sPaperWidth - Me.MarginLeft - Me.MarginRight
sPageH = sPaperHeight - Me.MarginDown - Me.MarginUp


bMirrorLR = CBool(Me.chkMirrorLR)
bMirrorUD = CBool(Me.chkMirrorUD)

bPageNums = CBool(Me.chkPageNums)
bBooklet = CBool(Me.chkOddEven)

If bPageNums Then
    Printer.FontSize = 12
    sPageH = sPageH - Printer.TextHeight("00")
End If
If iOrientation = 1 Then
  Printer.FontSize = 10 * 3 / iColNum
Else
  Printer.FontSize = 10 * 4 / iColNum
End If


' ISPIS

For i = 1 To KolekcijaZaStampu.Count
    
    Dim sSifraTMP As String, sModelTMP As String
    sSifraTMP = KolekcijaZaStampu(i).Sifra
    sModelTMP = KolekcijaZaStampu(i).Model
    Dim sPicW As Single, sPicH As Single
    sPicW = sPageW / iColNum
    sPicH = sPageH / iRowNum
    
       

' ŠTAMPANJE FOTOGRAFIJA

    If FileExists(MyDirectory & "\pic\p" & KolekcijaZaStampu(i).ID & ".jpg") Then
    
    Dim cPic As New clsJPEGparser
    Dim TMPW As Single, TMPH As Single
    TMPW = sPicW
    TMPH = sPicH - 0.5 - Printer.TextHeight("0000")
        Me.ctlPicture.Picture = LoadPicture(MyDirectory & "\pic\p" & KolekcijaZaStampu(i).ID & ".jpg")
        cPic.ParseJpegFile MyDirectory & "\pic\p" & KolekcijaZaStampu(i).ID & ".jpg"
        If cPic.XsizePicture / cPic.YsizePicture > TMPW / TMPH Then
            Printer.PaintPicture Me.ctlPicture.Picture, iMarginL + ((i - 1) Mod iColNum) * sPicW, iMarginU + (Int((i - 1) / iColNum) Mod iRowNum) * sPicH + (TMPH - cPic.YsizePicture / cPic.XsizePicture * TMPW) / 2, TMPW, cPic.YsizePicture / cPic.XsizePicture * TMPW
        ElseIf cPic.XsizePicture / cPic.YsizePicture < TMPW / TMPH Then
            Printer.PaintPicture Me.ctlPicture.Picture, iMarginL + ((i - 1) Mod iColNum) * sPicW + (TMPW - TMPH * cPic.XsizePicture / cPic.YsizePicture) / 2, iMarginU + (Int((i - 1) / iColNum) Mod iRowNum) * sPicH, TMPH * cPic.XsizePicture / cPic.YsizePicture, TMPH
        Else
            Printer.PaintPicture Me.ctlPicture.Picture, iMarginL + ((i - 1) Mod iColNum) * sPicW, iMarginU + (Int((i - 1) / iColNum) Mod iRowNum) * sPicH, TMPW, TMPH
        End If
        Set cPic = Nothing
    End If

  ' ŠTAMPANJE PODATAKA O SLICI
    
    Printer.CurrentX = iMarginL + ((i - 1) Mod iColNum) * sPicW + sPicW - 1 - Printer.TextWidth(sModelTMP)
    Printer.CurrentY = iMarginU + (Int((i - 1) / iColNum) Mod iRowNum) * sPicH + sPicH - 0.5 - Printer.TextHeight(sSifraTMP)
    Printer.Print sModelTMP
    Printer.CurrentX = iMarginL + ((i - 1) Mod iColNum) * sPicW + 1
    Printer.CurrentY = iMarginU + (Int((i - 1) / iColNum) Mod iRowNum) * sPicH + sPicH - 0.5 - Printer.TextHeight(sSifraTMP)
    Printer.Print sSifraTMP

' ŠTAMPANJE OKVIRA SLIKE
 
'x1=iMarginL + ((i - 1) Mod iColNum) * sPicW
'x2=iMarginL + ((i - 1) Mod iColNum) * sPicW + sPicW
'Y1=iMarginU + (Int((i - 1) / iColNum) Mod iRowNum) * sPicH
'y2=iMarginU + (Int((i - 1) / iColNum) Mod iRowNum) * sPicH + sPicH

 
    Printer.Line (iMarginL + ((i - 1) Mod iColNum) * sPicW, iMarginU + (Int((i - 1) / iColNum) Mod iRowNum) * sPicH)-(iMarginL + ((i - 1) Mod iColNum) * sPicW + sPicW, iMarginU + (Int((i - 1) / iColNum) Mod iRowNum) * sPicH)
    Printer.Line (iMarginL + ((i - 1) Mod iColNum) * sPicW + sPicW, iMarginU + (Int((i - 1) / iColNum) Mod iRowNum) * sPicH)-(iMarginL + ((i - 1) Mod iColNum) * sPicW + sPicW, iMarginU + (Int((i - 1) / iColNum) Mod iRowNum) * sPicH + sPicH)
    Printer.Line (iMarginL + ((i - 1) Mod iColNum) * sPicW + sPicW, iMarginU + (Int((i - 1) / iColNum) Mod iRowNum) * sPicH + sPicH)-(iMarginL + ((i - 1) Mod iColNum) * sPicW, iMarginU + (Int((i - 1) / iColNum) Mod iRowNum) * sPicH + sPicH)
    Printer.Line (iMarginL + ((i - 1) Mod iColNum) * sPicW, iMarginU + (Int((i - 1) / iColNum) Mod iRowNum) * sPicH + sPicH)-(iMarginL + ((i - 1) Mod iColNum) * sPicW, iMarginU + (Int((i - 1) / iColNum) Mod iRowNum) * sPicH)
    



' ŠTAMPANJE BROJEVA STRANICA

    If bPageNums And (i Mod (iRowNum * iColNum) = 1) Then
        Printer.FontSize = 12
        Dim pn As String
        pn = CStr((i - 1) / (iRowNum * iColNum) + 1)
        Printer.CurrentX = iMarginL + sPageW / 2 - Printer.TextWidth(pn) / 2
        Printer.CurrentY = iMarginU + sPageH
        Printer.Print pn
        Printer.FontSize = 8
    End If

    
' PRELEZAK NA SLJEDEÆU STRANU
    
    If (i Mod (iRowNum * iColNum)) = 0 Then
        Printer.NewPage
    End If
    
Next

Printer.EndDoc

Me.Hide
Me.cmdPrint.Enabled = True
Me.MousePointer = 0


End Sub
'
Private Sub Form_Load()
    Dim i As Integer
    For i = 0 To Printers.Count - 1
        Me.cboPrinterList.AddItem Printers(i).DeviceName, i
    Next
    Me.cboPrinterList = Printer.DeviceName
    Me.cboPageSize.AddItem "A4 (210×297 mm)", 0
    Me.cboPageSize.ListIndex = 0
End Sub


Private Sub MarginDown_Change()
ValidatePrintButton
End Sub

Private Sub MarginDown_GotFocus()
Me.MarginDown.SelStart = 0
Me.MarginDown.SelLength = Len(Me.MarginDown)
End Sub

Private Sub MarginDown_KeyPress(KeyAscii As Integer)
If (KeyAscii > 57 Or KeyAscii < 48) And KeyAscii <> 8 Then KeyAscii = 0
End Sub


Private Sub MarginLeft_Change()
ValidatePrintButton
End Sub

Private Sub MarginLeft_GotFocus()
Me.MarginLeft.SelStart = 0
Me.MarginLeft.SelLength = Len(Me.MarginLeft)
End Sub

Private Sub MarginLeft_KeyPress(KeyAscii As Integer)
If (KeyAscii > 57 Or KeyAscii < 48) And KeyAscii <> 8 Then KeyAscii = 0
End Sub


Private Sub MarginRight_Change()
ValidatePrintButton
End Sub

Private Sub MarginRight_GotFocus()
Me.MarginRight.SelStart = 0
Me.MarginRight.SelLength = Len(Me.MarginRight)
End Sub

Private Sub MarginRight_KeyPress(KeyAscii As Integer)
If (KeyAscii > 57 Or KeyAscii < 48) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub MarginUp_Change()
ValidatePrintButton
End Sub

Private Sub MarginUp_GotFocus()
Me.MarginUp.SelStart = 0
Me.MarginUp.SelLength = Len(Me.MarginUp)
End Sub

Private Sub MarginUp_KeyPress(KeyAscii As Integer)
If (KeyAscii > 57 Or KeyAscii < 48) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub optAll_Click()
optModels_Click
End Sub


Private Sub optModels_Click()
If Me.optModels = True Then
    Me.txtModels.BackColor = vbWindowBackground
    Me.txtModels.Enabled = True
Else
    Me.txtModels.BackColor = vbButtonFace
    Me.txtModels.Enabled = False
End If
ValidatePrintButton
End Sub

Private Sub optSelected_Click()
optModels_Click
End Sub

Private Sub txtCols_Change()
ValidatePrintButton
End Sub

Private Sub txtCols_GotFocus()
Me.txtCols.SelStart = 0
Me.txtCols.SelLength = Len(Me.txtCols)
End Sub

Private Sub txtCols_KeyPress(KeyAscii As Integer)
If (KeyAscii > 57 Or KeyAscii < 48) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtModels_Change()
ValidatePrintButton
End Sub

Private Sub txtModels_GotFocus()
Me.txtModels.SelStart = 0
Me.txtModels.SelLength = Len(Me.txtModels)
End Sub

Private Sub txtModels_KeyPress(KeyAscii As Integer)
If (KeyAscii > 57 Or KeyAscii < 48) And (KeyAscii < 44 Or KeyAscii > 45) And KeyAscii <> 8 Then KeyAscii = 0
If KeyAscii = 44 Or KeyAscii = 45 Then
    If Me.txtModels.SelStart <> 0 Then
        If Mid(Me.txtModels, Me.txtModels.SelStart + 1, 1) = Chr(44) Or _
            Mid(Me.txtModels, Me.txtModels.SelStart + 1, 1) = Chr(45) Or _
            Mid(Me.txtModels, Me.txtModels.SelStart, 1) = Chr(44) Or _
            Mid(Me.txtModels, Me.txtModels.SelStart, 1) = Chr(45) Then KeyAscii = 0
        
    Else
        If Mid(Me.txtModels, Me.txtModels.SelStart + 1, 1) = Chr(44) Or _
            Mid(Me.txtModels, Me.txtModels.SelStart + 1, 1) = Chr(45) Then KeyAscii = 0
    End If
End If
If InStrRev(Me.txtModels, "-") <> 0 Then
    If KeyAscii = 45 Then
        If InStrRev(Me.txtModels, "-") > InStrRev(Me.txtModels, ",") Then KeyAscii = 0
    End If
End If
End Sub

Private Sub txtRows_Change()
ValidatePrintButton
End Sub

Private Sub txtRows_GotFocus()
Me.txtRows.SelStart = 0
Me.txtRows.SelLength = Len(Me.txtRows)
End Sub

Private Sub txtRows_KeyPress(KeyAscii As Integer)
If (KeyAscii > 57 Or KeyAscii < 48) And KeyAscii <> 8 Then KeyAscii = 0
End Sub


Private Sub ValidatePrintButton()
Dim LMargin As Single, UMargin As Single
Dim bGo As Boolean
bGo = True
Set Printer = Printers(Me.cboPrinterList.ListIndex)
Printer.ScaleMode = vbMillimeters
Printer.PaperSize = 9
LMargin = Int((210 - Printer.ScaleWidth) / 2 + 1)
UMargin = Int((297 - Printer.ScaleHeight) / 2 + 1)
bGo = bGo And Me.txtRows <> ""
bGo = bGo And Me.txtCols <> ""
bGo = bGo And (Val(Me.txtRows) <= 10 And Val(Me.txtRows) > 0)
bGo = bGo And (Val(Me.txtCols) <= 10 And Val(Me.txtCols) > 0)
bGo = bGo And Val(Me.MarginLeft) >= LMargin
bGo = bGo And Val(Me.MarginRight) >= LMargin
bGo = bGo And Val(Me.MarginUp) >= UMargin
bGo = bGo And Val(Me.MarginDown) >= UMargin
If Me.optModels Then
    bGo = bGo And Me.txtModels <> ""
End If
Me.cmdPrint.Enabled = bGo
End Sub


