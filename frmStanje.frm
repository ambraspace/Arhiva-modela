VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStanje 
   Caption         =   "Stanje"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lstStanje 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   12303
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "id"
         Text            =   "ID"
         Object.Width           =   900
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "prodavnica"
         Text            =   "Prodavnica"
         Object.Width           =   6006
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "pocStanje"
         Text            =   "Poè. stanje"
         Object.Width           =   1164
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   "prijem"
         Text            =   "Prijem"
         Object.Width           =   1164
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   "otprema"
         Text            =   "Otprema"
         Object.Width           =   1164
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Key             =   "prodaja"
         Text            =   "Prodaja"
         Object.Width           =   1164
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Key             =   "stanej"
         Text            =   "Stanje"
         Object.Width           =   1164
      EndProperty
   End
End
Attribute VB_Name = "frmStanje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub PrikaziMe()
    Me.Show
    KeepOnTop Me
    frmGlavni.SetFocus
    Me.Display
End Sub

Public Sub Display()

    Dim rsStanje As Recordset
    
    Me.lstStanje.ListItems.Clear
        
    Me.Caption = frmGlavni.ctlListView.SelectedItem.ListSubItems(1).Text
    
    Set rsStanje = dbModeli.OpenRecordset("SELECT * FROM ARHIVA WHERE ARTIKL=" & frmGlavni.ctlListView.SelectedItem.Text & _
            " ORDER BY PROD")
    
    
    If rsStanje.RecordCount > 0 Then
        
        Dim ID As Long, ukPocStanje As Long, ukPrijem As Long, ukOtprema As Long, ukProdaja As Long, ukStanje As Long
        ukPocStanje = 0
        ukPrijem = 0
        ukOtprema = 0
        ukProdaja = 0
        ukStanje = 0
        Do While Not rsStanje.EOF
            ID = rsStanje.AbsolutePosition + 1
            If rsStanje("POCSTANJE") <> 0 Or rsStanje("PRIJEM") <> 0 Or rsStanje("OTPREMA") <> 0 Or _
                rsStanje("PRODAJA") <> 0 Or rsStanje("ZSTANJE") <> 0 Then
                ukPocStanje = ukPocStanje + rsStanje("POCSTANJE")
                ukPrijem = ukPrijem + rsStanje("PRIJEM")
                ukOtprema = ukOtprema + rsStanje("OTPREMA")
                ukProdaja = ukProdaja + rsStanje("PRODAJA")
                ukStanje = ukStanje + rsStanje("ZSTANJE")
                Me.lstStanje.ListItems.Add , "ID" & ID, rsStanje("PROD")
                If Not IsNull(rsStanje("PRODAVNICA")) Then Me.lstStanje.ListItems("ID" & ID).SubItems(1) = rsStanje("PRODAVNICA")
                Me.lstStanje.ListItems("ID" & ID).SubItems(2) = rsStanje("POCSTANJE")
                Me.lstStanje.ListItems("ID" & ID).SubItems(3) = rsStanje("PRIJEM")
                Me.lstStanje.ListItems("ID" & ID).SubItems(4) = rsStanje("OTPREMA")
                Me.lstStanje.ListItems("ID" & ID).SubItems(5) = rsStanje("PRODAJA")
                Me.lstStanje.ListItems("ID" & ID).SubItems(6) = rsStanje("ZSTANJE")
                
            End If
            rsStanje.MoveNext
        Loop
        Me.lstStanje.ListItems.Add , "UK", " "
        Me.lstStanje.ListItems("UK").SubItems(1) = "UKUPNO"
        Me.lstStanje.ListItems("UK").SubItems(2) = ukPocStanje
        Me.lstStanje.ListItems("UK").SubItems(3) = ukPrijem
        Me.lstStanje.ListItems("UK").SubItems(4) = ukOtprema
        Me.lstStanje.ListItems("UK").SubItems(5) = ukProdaja
        Me.lstStanje.ListItems("UK").SubItems(6) = ukStanje
                 
    End If
    
    rsStanje.Close
    Set rsStanje = Nothing

End Sub


Private Sub Form_Load()
    Me.Top = frmGlavni.Top
    Me.Left = frmGlavni.Left + frmGlavni.Width - Me.Width
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    frmGlavni.mnuStanje.Checked = False
End Sub

Private Sub Form_Resize()
    Me.lstStanje.Width = Me.ScaleWidth - 2 * Me.lstStanje.Left
    Me.lstStanje.Height = Me.ScaleHeight - 2 * Me.lstStanje.Top
End Sub
