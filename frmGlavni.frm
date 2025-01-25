VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGlavni 
   Caption         =   " Arhiva modela"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7800
   Icon            =   "frmGlavni.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView ctlListView 
      Height          =   3495
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Šifra"
         Object.Width           =   1535
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Model"
         Object.Width           =   5874
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Dobavljaè"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Datum nabavke"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Postoji u odab. obj."
         Object.Width           =   2831
      EndProperty
   End
   Begin MSComctlLib.StatusBar ctlStatusBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   4080
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3201
            MinWidth        =   2
            Text            =   "Ukupno modela u arhivi:"
            TextSave        =   "Ukupno modela u arhivi:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   4577
            MinWidth        =   2
            Text            =   "Pozicija trenutno izabranog modela:"
            TextSave        =   "Pozicija trenutno izabranog modela:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1535
            MinWidth        =   2
            Text            =   "Naziv fajla:"
            TextSave        =   "Naziv fajla:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2619
            Text            =   "Izabrano: 0 modela "
            TextSave        =   "Izabrano: 0 modela "
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuModeli 
      Caption         =   "&Modeli"
      Begin VB.Menu mnuModeliFoto 
         Caption         =   "Potpuni prikaz &modela"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuModeliPreview 
         Caption         =   "Pokazuj &fotografiju"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuStanje 
         Caption         =   "Pokazuj &stanje"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuModeliSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModeliDodaj 
         Caption         =   "&Dodaj"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuModeliIzmijeni 
         Caption         =   "&Izmijeni"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuModeliBriši 
         Caption         =   "&Briši"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuModeliSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModeliTrazi 
         Caption         =   "&Pretraživanje"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuModeliSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModeliStampanje 
         Caption         =   "Š&tampanje"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuModeliKopiraj 
         Caption         =   "&Kopiraj"
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu mnuSortiranje 
      Caption         =   "&Sortiranje"
      Begin VB.Menu mnuSortSortirano 
         Caption         =   "&Sortirano"
      End
      Begin VB.Menu mnuSortSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSortAscending 
         Caption         =   "&Rastuæi redosljed"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSortDescending 
         Caption         =   "&Opadajuæi redosljed"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSortSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSortSifra 
         Caption         =   "Po šifri"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSortModeli 
         Caption         =   "Po modelu"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSortDobavljac 
         Caption         =   "Po dobavljaèu"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSortDatum 
         Caption         =   "Po datumu nabavke"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSortPrisutnost 
         Caption         =   "Po prisutnosti"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuAlati 
      Caption         =   "&Alati"
      Begin VB.Menu mnuAlatiDobavljaci 
         Caption         =   "Popravi listu dobavljaèa"
      End
      Begin VB.Menu mnuAlatiIme 
         Caption         =   "Popravi ime dobavljaèa"
      End
      Begin VB.Menu mnuAlatiReplaceString 
         Caption         =   "Popravi polje 'Model'"
      End
      Begin VB.Menu mnuPrisutnostUObjektu 
         Caption         =   "Prikaži prisutnost u objektu"
      End
   End
End
Attribute VB_Name = "frmGlavni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub DisplayMe()
On Error GoTo greska

frmSplash.Show
frmSplash.Refresh

Me.ctlListView.ListItems.Clear

Dim i As Integer
If rsModeli.RecordCount = 0 Then GoTo over

rsModeli.MoveLast
rsModeli.MoveFirst

For i = 1 To rsModeli.RecordCount
    Me.ctlListView.ListItems.Add , "ID" & rsModeli("ID"), rsModeli("Sifra")
    Me.ctlListView.ListItems("ID" & rsModeli("ID")).SubItems(1) = rsModeli("Model")
    Me.ctlListView.ListItems("ID" & rsModeli("ID")).SubItems(2) = rsModeli("Dobavljac")
    Me.ctlListView.ListItems("ID" & rsModeli("ID")).SubItems(3) = Format(rsModeli("Datum nabavke"), "d. M. yyyy.")
    Me.ctlListView.ListItems("ID" & rsModeli("ID")).Bold = True
    Me.ctlListView.ListItems("ID" & rsModeli("ID")).SubItems(4) = "DA"
    If Not rsModeli("Aktuelan") Then
        Me.ctlListView.ListItems("ID" & rsModeli("ID")).ForeColor = RGB(180, 180, 180)
        Me.ctlListView.ListItems("ID" & rsModeli("ID")).SubItems(4) = ""
    End If
    
    rsModeli.MoveNext
Next

If Me.ctlListView.SelectedItem.Selected Then Me.ctlListView.SelectedItem.Selected = False
Me.ctlListView.ListItems.Item(rsModeli.RecordCount).EnsureVisible
Set Me.ctlListView.SelectedItem = Me.ctlListView.ListItems(rsModeli.RecordCount)
ctlListView_ItemClick Me.ctlListView.ListItems.Item(rsModeli.RecordCount)

over:
Me.ValidateStatusBar
Me.Show
Unload frmSplash
GoTo kraj

greska:
If Err.Number = 94 Then
    Resume Next
Else
    Err.Raise Err.Number, Err.Source, Err.Description
End If

kraj:
End Sub


Private Sub ctlListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If Not Me.mnuSortSortirano.Checked Then

    Me.mnuSortSortirano.Checked = True
    
    Me.mnuSortAscending.Enabled = True
    Me.mnuSortDescending.Enabled = True
    Me.mnuSortModeli.Enabled = True
    Me.mnuSortSifra.Enabled = True
    Me.mnuSortDatum.Enabled = True
    Me.mnuSortDobavljac.Enabled = True
    Me.mnuSortPrisutnost.Enabled = True
    
    Me.mnuSortAscending.Checked = True
    Me.mnuSortDescending.Checked = False
    Me.mnuSortModeli.Checked = False
    Me.mnuSortSifra.Checked = False
    Me.mnuSortDatum.Checked = False
    Me.mnuSortDobavljac.Checked = False
    Me.mnuSortPrisutnost.Checked = False
    Select Case ColumnHeader.Index
        Case 1
            Me.mnuSortSifra.Checked = True
        Case 2
            Me.mnuSortModeli.Checked = True
        Case 3
            Me.mnuSortDobavljac.Checked = True
        Case 4
            Me.mnuSortDatum.Checked = True
        Case 5
            Me.mnuSortPrisutnost.Checked = True
    End Select
Else
    If Me.ctlListView.SortKey + 1 <> ColumnHeader.Index Then
        Me.mnuSortAscending.Checked = True
        Me.mnuSortDescending.Checked = False
        Me.mnuSortModeli.Checked = False
        Me.mnuSortSifra.Checked = False
        Me.mnuSortDatum.Checked = False
        Me.mnuSortDobavljac.Checked = False
        Me.mnuSortPrisutnost.Checked = False
        Select Case ColumnHeader.Index
            Case 1
                Me.mnuSortSifra.Checked = True
            Case 2
                Me.mnuSortModeli.Checked = True
            Case 3
                Me.mnuSortDobavljac.Checked = True
            Case 4
                Me.mnuSortDatum.Checked = True
            Case 5
                Me.mnuSortPrisutnost.Checked = True
        End Select
    Else
        Select Case Me.ctlListView.SortOrder
            Case lvwAscending
                Me.mnuSortAscending.Checked = False
                Me.mnuSortDescending.Checked = True
            Case lvwDescending
                Me.mnuSortAscending.Checked = True
                Me.mnuSortDescending.Checked = False
        End Select
    End If
End If
Me.Sortiraj
End Sub

Public Sub ctlListView_ItemClick(ByVal Item As MSComctlLib.ListItem)
Me.ValidateStatusBar
Me.MousePointer = 11
If Me.mnuModeliPreview.Checked Then frmPicPreview.Display
If Me.mnuStanje.Checked Then frmStanje.Display
Me.MousePointer = 0
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.Width < 9000 Then Me.Width = 9000
If Me.Height < 6000 Then Me.Height = 6000
Me.ctlListView.Width = Me.ScaleWidth - (2 * Me.ctlListView.Left)
Me.ctlListView.Height = Me.ScaleHeight - (1 * Me.ctlListView.Top) - Me.ctlStatusBar.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
rsModeli.Close
rsDobavljaci.Close
dbModeli.Close
If Me.mnuModeliPreview.Checked Then Unload frmPicPreview
End
End Sub


Private Sub mnuAlatiDobavljaci_Click()
Dim i As Long

If rsModeli.RecordCount = 0 Then
    If rsDobavljaci.RecordCount = 0 Then
        Exit Sub
    Else
        rsDobavljaci.MoveLast
        Do Until rsDobavljaci.BOF
            rsDobavljaci.Delete
            rsDobavljaci.MovePrevious
        Loop
        rsModeli.Close
        rsDobavljaci.Close
        dbModeli.Close
        DBEngine.CompactDatabase MyDirectory & "\ARHIVA.MDB", MyDirectory & "\DB_tmp.tmp"
        ' , , , ";pwd=n4123ekZ"
        Kill MyDirectory & "\ARHIVA.MDB"
        FileCopy MyDirectory & "\DB_tmp.tmp", MyDirectory & "\ARHIVA.MDB"
        Kill MyDirectory & "\DB_tmp.tmp"
        Set dbModeli = OpenDatabase(MyDirectory & "\ARHIVA.MDB", False, False) ', ";pwd=n4123ekZ"
        Set rsModeli = dbModeli.OpenRecordset("Modeli", dbOpenDynaset, dbSeeChanges)
        Set rsDobavljaci = dbModeli.OpenRecordset("Dobavljaci", dbOpenDynaset, dbSeeChanges)
    End If
Else
    If rsDobavljaci.RecordCount > 0 Then
        rsDobavljaci.MoveLast
        Do Until rsDobavljaci.BOF
            rsDobavljaci.Delete
            rsDobavljaci.MovePrevious
        Loop
    End If
    
        rsModeli.Close
        rsDobavljaci.Close
        dbModeli.Close
        DBEngine.CompactDatabase MyDirectory & "\ARHIVA.MDB", MyDirectory & "\DB_tmp.tmp"
        ', , , ";pwd=n4123ekZ"
        Kill MyDirectory & "\ARHIVA.MDB"
        FileCopy MyDirectory & "\DB_tmp.tmp", MyDirectory & "\ARHIVA.MDB"
        Kill MyDirectory & "\DB_tmp.tmp"
        Set dbModeli = OpenDatabase(MyDirectory & "\ARHIVA.MDB", False, False)
        ', ";pwd=n4123ekZ"
        Set rsModeli = dbModeli.OpenRecordset("Modeli", dbOpenDynaset, dbSeeChanges)
        Set rsDobavljaci = dbModeli.OpenRecordset("Dobavljaci", dbOpenDynaset, dbSeeChanges)
    
    rsModeli.MoveLast
    rsModeli.MoveFirst
    Dim bNoviDobavljac As Boolean
    For i = 1 To rsModeli.RecordCount
        If Trim(rsModeli("Dobavljac")) <> "" Then
            bNoviDobavljac = True
            If rsDobavljaci.RecordCount > 0 Then
                rsDobavljaci.MoveFirst
                Do Until rsDobavljaci.EOF
                    If rsDobavljaci("Ime") = rsModeli("Dobavljac") Then
                        bNoviDobavljac = False
                        Exit Do
                    End If
                    rsDobavljaci.MoveNext
                Loop
                If bNoviDobavljac Then
                    rsDobavljaci.AddNew
                    rsDobavljaci("Ime") = rsModeli("Dobavljac")
                    rsDobavljaci.Update
                End If
            Else
                rsDobavljaci.AddNew
                rsDobavljaci("Ime") = rsModeli("Dobavljac")
                rsDobavljaci.Update
            End If
        End If
    rsModeli.MoveNext
    Next
End If
Dim k As Integer
k = MsgBox("Lista dobavljaèa popravljena!")

End Sub

Private Sub mnuAlatiIme_Click()
frmPopravakImenaDobavljaca.Start
End Sub

Private Sub mnuAlatiReplaceString_Click()
If rsModeli.RecordCount = 0 Then Exit Sub
frmReplaceString.Zamijeni
End Sub

Private Sub mnuModeliBriši_Click()
    
Dim i As Integer, odgovor As Integer

If rsModeli.RecordCount = 0 Then Exit Sub

    odgovor = MsgBox("Da li ste sigurni da želite da obrišete trenutno odabrane modele?", vbYesNo + vbQuestion, "Brisanje modela")
    
    Select Case odgovor
        Case vbYes
            For i = Me.ctlListView.ListItems.Count To 1 Step -1
                If Me.ctlListView.ListItems(i).Selected Then
                    If FileExists(MyDirectory & "\pic\p" & Mid(Me.ctlListView.ListItems(i).Key, 3) & ".jpg") Then Kill MyDirectory & "\pic\p" & Mid(Me.ctlListView.ListItems(i).Key, 3) & ".jpg"
                    rsModeli.FindFirst "ID=" & Mid(Me.ctlListView.ListItems(i).Key, 3)
                    rsModeli.Delete
                    Me.ctlListView.ListItems.Remove Me.ctlListView.ListItems(i).Key
                End If
            Next
        Case vbNo
            Exit Sub
    End Select
Me.ValidateStatusBar
End Sub

Private Sub mnuModeliDodaj_Click()
    Me.Enabled = False
    frmEditModel.DodajModel
End Sub

Private Sub mnuModeliFoto_Click()
If Me.fnMultiple = 1 Then
    frmPrikaz.Prikazi Me.ctlListView.SelectedItem.Text, Me.ctlListView.SelectedItem.SubItems(1), Me.ctlListView.SelectedItem.SubItems(2), Me.ctlListView.SelectedItem.SubItems(3), MyDirectory & "\pic\p" & Mid(Me.ctlListView.SelectedItem.Key, 3) & ".jpg", Mid(Me.ctlListView.SelectedItem.Key, 3)
    frmPrikaz.Form_Resize
End If
End Sub

Private Sub mnuModeliIzmijeni_Click()
If Me.fnMultiple = 1 Then
Me.Enabled = False
frmEditModel.IzmijeniModel
End If
End Sub

Private Sub mnuModeliKopiraj_Click()
Dim i As MSComctlLib.ListItem
BrowseFolder Me.hWnd, "Izaberite folder"
If BrowseFolder_Successful Then
    If Right(BrowseFolder_FolderName, 1) <> "\" Then BrowseFolder_FolderName = BrowseFolder_FolderName & "\"
    For Each i In Me.ctlListView.ListItems
        If i.Selected Then FileCopy MyDirectory & "\pic\p" & Mid(i.Key, 3) & ".jpg", BrowseFolder_FolderName & "p" & Mid(i.Key, 3) & "-" & i.Text & ".jpg"
    Next
End If
End Sub

Private Sub mnuModeliPreview_Click()
If Me.mnuModeliPreview.Checked Then
    frmPicPreview.Hide
    Me.mnuModeliPreview.Checked = False
Else
    frmPicPreview.PrikaziMe
    Me.mnuModeliPreview.Checked = True
End If
End Sub

Public Function fnMultiple() As Integer
Dim X As MSComctlLib.ListItem, i As Long
i = 0
For Each X In Me.ctlListView.ListItems
    If X.Selected Then i = i + 1
    If i > 1 Then Exit For
Next
fnMultiple = i
If fnMultiple > 1 Then fnMultiple = 2
End Function

Public Sub ValidateStatusBar()
On Error Resume Next
Me.ctlStatusBar.Panels(1) = "Ukupno modela u arhivi: " & rsModeli.RecordCount
If Me.fnMultiple = 1 Then
    Me.ctlStatusBar.Panels(2) = "Pozicija oznaèenog modela: " & Me.ctlListView.SelectedItem.Index & " "
    Select Case FileExists(MyDirectory & "\pic\p" & Mid(Me.ctlListView.SelectedItem.Key, 3) & ".jpg")
        Case True
            Me.ctlStatusBar.Panels(3) = "Naziv fajla: p" & Mid(Me.ctlListView.SelectedItem.Key, 3) & ".jpg "
        Case False
            Me.ctlStatusBar.Panels(3) = "Naziv fajla: ? "
    End Select
Else
    Me.ctlStatusBar.Panels(2) = "Pozicija oznaèenog modela: "
    Me.ctlStatusBar.Panels(3) = "Naziv fajla: "
End If
Me.ctlStatusBar.Panels(4) = "Izabrano modela: " & fnSelectedCount & " "
End Sub

Private Function fnSelectedCount() As Long
Dim sum As Long
sum = 0
Dim i As MSComctlLib.ListItem
For Each i In Me.ctlListView.ListItems
    If i.Selected Then sum = sum + 1
Next
fnSelectedCount = sum
End Function

Private Sub mnuModeliStampanje_Click()
If rsModeli.RecordCount = 0 Then Exit Sub
frmPrint.IspisPrikaz
End Sub

Private Sub mnuModeliTrazi_Click()
If Me.fnMultiple = 1 Then
    Me.Enabled = False
    frmTrazenje.Show
    frmTrazenje.txtSearchString.SetFocus
End If
End Sub

Private Sub mnuPrisutnostUObjektu_Click()

Dim i As Integer

Dim strAnswer As String
strAnswer = frmObjekat.Prikazi(Me)

If strAnswer = "" Then Exit Sub

Me.MousePointer = vbHourglass

If strAnswer = "*" Then
    For i = 1 To Me.ctlListView.ListItems.Count
        If Me.ctlListView.ListItems(i).ForeColor = Me.ctlListView.ForeColor Then
            Me.ctlListView.ListItems.Item(i).SubItems(4) = "DA"
        Else
            Me.ctlListView.ListItems.Item(i).SubItems(4) = ""
        End If
    Next
Else
    
    Dim rs As Recordset
    Set rs = dbModeli.OpenRecordset("SELECT DISTINCT ARTIKL FROM ARHIVA WHERE PROD='" & strAnswer & "'")
    For i = 1 To Me.ctlListView.ListItems.Count
        rs.MoveFirst
        rs.FindFirst "ARTIKL=" & Me.ctlListView.ListItems.Item(i).Text
        If Not rs.NoMatch Then
            Me.ctlListView.ListItems.Item(i).SubItems(4) = "DA"
        Else
            Me.ctlListView.ListItems.Item(i).SubItems(4) = ""
        End If
    Next
    
End If

Me.MousePointer = vbDefault
End Sub

Private Sub mnuSortAscending_Click()
Me.mnuSortAscending.Checked = True
Me.mnuSortDescending.Checked = False
Me.Sortiraj
End Sub

Private Sub mnuSortDatum_Click()
Me.mnuSortSifra.Checked = False
Me.mnuSortModeli.Checked = False
Me.mnuSortDobavljac.Checked = False
Me.mnuSortPrisutnost.Checked = False
Me.mnuSortDatum.Checked = True
Me.Sortiraj

End Sub

Private Sub mnuSortDescending_Click()
Me.mnuSortAscending.Checked = False
Me.mnuSortDescending.Checked = True
Me.Sortiraj
End Sub

Private Sub mnuSortDobavljac_Click()
Me.mnuSortSifra.Checked = False
Me.mnuSortModeli.Checked = False
Me.mnuSortDobavljac.Checked = True
Me.mnuSortDatum.Checked = False
Me.mnuSortPrisutnost.Checked = False

Me.Sortiraj

End Sub

Private Sub mnuSortModeli_Click()
Me.mnuSortSifra.Checked = False
Me.mnuSortModeli.Checked = True
Me.mnuSortDobavljac.Checked = False
Me.mnuSortDatum.Checked = False
Me.mnuSortPrisutnost.Checked = False
Me.Sortiraj

End Sub

Private Sub mnuSortPrisutnost_Click()
Me.mnuSortSifra.Checked = False
Me.mnuSortModeli.Checked = False
Me.mnuSortDobavljac.Checked = False
Me.mnuSortDatum.Checked = False
Me.mnuSortPrisutnost.Checked = True
Me.Sortiraj
End Sub

Private Sub mnuSortSifra_Click()
Me.mnuSortSifra.Checked = True
Me.mnuSortModeli.Checked = False
Me.mnuSortDobavljac.Checked = False
Me.mnuSortDatum.Checked = False
Me.mnuSortPrisutnost.Checked = False
Me.Sortiraj
End Sub



Private Sub mnuSortSortirano_Click()
'If Me.mnuSortSortirano.Checked Then
'    Me.mnuSortSortirano.Checked = False
'    Me.mnuSortAscending.Enabled = False
'    Me.mnuSortDescending.Enabled = False
'    Me.mnuSortSifra.Enabled = False
'    Me.mnuSortModeli.Enabled = False
'    Me.mnuSortDobavljac.Enabled = False
'    Me.mnuSortDatum.Enabled = False
'    Me.mnuSortPrisutnost.Enabled = False
'Else
'    Me.mnuSortSortirano.Checked = True
'    Me.mnuSortAscending.Enabled = True
'    Me.mnuSortDescending.Enabled = True
'    Me.mnuSortSifra.Enabled = True
'    Me.mnuSortModeli.Enabled = True
'    Me.mnuSortDobavljac.Enabled = True
'    Me.mnuSortDatum.Enabled = True
'    Me.mnuSortPrisutnost.Enabled = True
'End If
If Not Me.mnuSortSortirano.Checked Then
    Me.mnuSortSortirano.Checked = True
    Me.mnuSortAscending.Enabled = True
    Me.mnuSortDescending.Enabled = True
    Me.mnuSortSifra.Enabled = True
    Me.mnuSortModeli.Enabled = True
    Me.mnuSortDobavljac.Enabled = True
    Me.mnuSortDatum.Enabled = True
    Me.mnuSortPrisutnost.Enabled = True
End If

Me.Sortiraj
End Sub

Public Sub Sortiraj()
If Me.mnuSortSortirano.Checked Then
    Dim TMP As MSComctlLib.ListItem
        
    If Me.mnuSortSifra.Checked Then
        For Each TMP In Me.ctlListView.ListItems
            TMP.Text = String(10 - Len(TMP.Text), "0") & TMP.Text
        Next
    End If
    
    If Me.mnuSortDatum.Checked Then
        For Each TMP In Me.ctlListView.ListItems
            Select Case Right(TMP.SubItems(3), 1) = "."
                Case True
                    TMP.SubItems(3) = Format(Left(TMP.SubItems(3), Len(TMP.SubItems(3)) - 1), "yyyy-MM-dd")
                Case False
                    TMP.SubItems(3) = Format(TMP.SubItems(3), "yyyy-MM-dd")
            End Select
        Next
    End If
                
    
    Me.ctlListView.Sorted = True
    Select Case Me.mnuSortAscending.Checked
        Case True
            Me.ctlListView.SortOrder = lvwAscending
        Case False
            Me.ctlListView.SortOrder = lvwDescending
    End Select
    If Me.mnuSortSifra.Checked Then
        Me.ctlListView.SortKey = 0
    ElseIf Me.mnuSortModeli.Checked Then
        Me.ctlListView.SortKey = 1
    ElseIf Me.mnuSortDobavljac.Checked Then
        Me.ctlListView.SortKey = 2
    ElseIf Me.mnuSortDatum.Checked Then
        Me.ctlListView.SortKey = 3
    ElseIf Me.mnuSortPrisutnost.Checked Then
        Me.ctlListView.SortKey = 4
    End If
    
    If Me.mnuSortSifra.Checked Then
        For Each TMP In Me.ctlListView.ListItems
            TMP.Text = Val(TMP.Text)
        Next
    End If
    
    If Me.mnuSortDatum.Checked Then
        For Each TMP In Me.ctlListView.ListItems
            TMP.SubItems(3) = Format(TMP.SubItems(3), "d. M. yyyy.")
        Next
    End If

End If
'    Me.ctlListView.Sorted = False
'
'    Me.ctlListView.ListItems.Clear
'
'    Dim i As Integer
'    If rsModeli.RecordCount > 0 Then
'
'        rsModeli.MoveLast
'        rsModeli.MoveFirst
'
'        For i = 1 To rsModeli.RecordCount
'            Me.ctlListView.ListItems.Add , "ID" & rsModeli("ID"), rsModeli("Sifra")
'            Me.ctlListView.ListItems("ID" & rsModeli("ID")).SubItems(1) = rsModeli("Model")
'            Me.ctlListView.ListItems("ID" & rsModeli("ID")).SubItems(2) = rsModeli("Dobavljac")
'            Me.ctlListView.ListItems("ID" & rsModeli("ID")).SubItems(3) = Format(rsModeli("Datum nabavke"), "d. M. yyyy.")
'            Me.ctlListView.ListItems("ID" & rsModeli("ID")).Bold = True
'            Me.ctlListView.ListItems("ID" & rsModeli("ID")).SubItems(4) = "DA"
'            If Not rsModeli("Aktuelan") Then
'                Me.ctlListView.ListItems("ID" & rsModeli("ID")).ForeColor = RGB(180, 180, 180)
'                Me.ctlListView.ListItems("ID" & rsModeli("ID")).SubItems(4) = ""
'            End If
'
'            rsModeli.MoveNext
'        Next
'
'    End If
'
'    If rsModeli.RecordCount > 0 Then
'        If Me.ctlListView.SelectedItem.Selected Then Me.ctlListView.SelectedItem.Selected = False
'            Me.ctlListView.SelectedItem = Me.ctlListView.ListItems(X)
'            ctlListView_ItemClick Me.ctlListView.ListItems.Item(X)
'        End If
'        Me.ctlListView.SelectedItem.EnsureVisible
'    End If
    
    Me.ValidateStatusBar
    
End Sub

Private Sub mnuStanje_Click()
    If Me.mnuStanje.Checked Then
        frmStanje.Hide
        Me.mnuStanje.Checked = False
    Else
        frmStanje.PrikaziMe
        Me.mnuStanje.Checked = True
    End If
End Sub
