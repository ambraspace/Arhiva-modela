Attribute VB_Name = "modGlavni"
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Const MAX_PATH = 260

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Public dbModeli As Database
Public rsModeli As Recordset
Public rsDobavljaci As Recordset


Public BrisatiIzvornuFotografiju As Boolean
Public PosljednjiDatum As String
Public CurrentFolder As String
Public MyDirectory As String

Sub Main()
If App.PrevInstance Then End
PrikupiPodatke
frmGlavni.DisplayMe
End Sub


Private Sub PrikupiPodatke()
Dim sTMP As String * 255, strlen As Integer, NewMyDir As String
Dim READONLY As Boolean
MyDirectory = CurDir
'MyDirectory = "C:\Documents and Settings\ambra\Desktop\Arhiva modela"
READONLY = False
'READONLY = True
If Right(MyDirectory, 1) = "\" Then MyDirectory = Left(MyDirectory, Len(MyDirectory) - 1)
strlen = GetPrivateProfileString("Main", "RemotePath", "", sTMP, 255, MyDirectory & "\settings.ini")
NewMyDir = Left(sTMP, strlen)
If NewMyDir <> "" Then MyDirectory = NewMyDir

Dim answer As Integer
answer = MsgBox("Želite li uvesti nove podatke?", vbQuestion + vbYesNo + vbDefaultButton2, "Uvoz novih podataka")
If answer = vbYes Then
    If FileExists(MyDirectory & "\ARHIVA.ldb") Then
        MsgBox "Neki korisnici trenutno koriste Arhivu modela." & vbCrLf & _
            "Uvoz je trenutno nemoguæ." & vbCrLf & _
            "Rad æe biti nastavljen bez uvoza novih podataka!", vbCritical + vbOKOnly
    Else
        CompactArhivaModela
    End If
End If

Set dbModeli = OpenDatabase(MyDirectory & "\ARHIVA.MDB", False, READONLY)

Set rsModeli = dbModeli.OpenRecordset("Modeli", dbOpenDynaset, dbSeeChanges)
Set rsDobavljaci = dbModeli.OpenRecordset("Dobavljaci", dbOpenDynaset, dbSeeChanges)

BrisatiIzvornuFotografiju = True
PosljednjiDatum = Format(Date, "dd-MM-yy")
CurrentFolder = MyDirectory
End Sub

Private Sub CompactArhivaModela()
Dim db As DAO.Database
Dim rsSource As Recordset, rsDestination As Recordset, rsModeli As Recordset

frmImport.Show

frmImport.lblOperation.Caption = "Brišem stare podatke..."
frmImport.ctrlProgress.Value = frmImport.ctrlProgress.Min
frmImport.Refresh
DoEvents
Set db = OpenDatabase(MyDirectory & "\ARHIVA.MDB")
db.Execute "DELETE FROM ARHIVA"
db.Close
frmImport.ctrlProgress.Value = frmImport.ctrlProgress.Max
frmImport.Refresh
DoEvents

frmImport.lblOperation.Caption = "Optimizujem bazu podataka..."
frmImport.ctrlProgress.Value = frmImport.ctrlProgress.Min
frmImport.Refresh
DoEvents
If FileExists(MyDirectory & "\compacted.MDB") Then Kill MyDirectory & "\compacted.MDB"
Access.CompactRepair MyDirectory & "\ARHIVA.MDB", MyDirectory & "\compacted.MDB"
If FileExists(MyDirectory & "\ARHIVA.MDB") Then Kill MyDirectory & "\ARHIVA.MDB"
FileCopy MyDirectory & "\compacted.MDB", MyDirectory & "\ARHIVA.MDB"
If FileExists(MyDirectory & "\compacted.MDB") Then Kill MyDirectory & "\compacted.MDB"
frmImport.ctrlProgress.Value = frmImport.ctrlProgress.Max
frmImport.Refresh
DoEvents

Set db = OpenDatabase(MyDirectory & "\ARHIVA.MDB")

frmImport.lblOperation.Caption = "Podešavam status modela..."
frmImport.ctrlProgress.Value = frmImport.ctrlProgress.Min
Set rsModeli = db.OpenRecordset("Modeli", dbOpenDynaset)
rsModeli.MoveLast
rsModeli.MoveFirst
Do Until rsModeli.EOF
    
    rsModeli.Edit
    rsModeli("Aktuelan") = False
    rsModeli.Update
    frmImport.Refresh
    DoEvents
    rsModeli.MoveNext
Loop

frmImport.lblOperation.Caption = "Uvozim nove podatke..."
frmImport.ctrlProgress.Value = frmImport.ctrlProgress.Min
frmImport.Refresh
DoEvents
Set rsSource = db.OpenRecordset("ARHIVA1")
Set rsDestination = db.OpenRecordset("ARHIVA")
rsSource.MoveLast
rsSource.MoveFirst

Do Until rsSource.EOF
    rsDestination.AddNew
    rsDestination("PROD") = rsSource("PROD")
    rsDestination("PRODAVNICA") = rsSource("PRODAVNICA")
    rsDestination("ARTIKL") = rsSource("ARTIKL")
    rsDestination("NAZIV") = rsSource("NAZIV")
    rsDestination("NAZIV11") = rsSource("NAZIV11")
    rsDestination("NAZIVA") = rsSource("NAZIVA")
    rsDestination("NAZIV22") = rsSource("NAZIV22")
    rsDestination("NAZIVB") = rsSource("NAZIVB")
    rsDestination("NAZIV33") = rsSource("NAZIV33")
    rsDestination("NAZIVC") = rsSource("NAZIVC")
    rsDestination("NAZIV44") = rsSource("NAZIV44")
    rsDestination("NAZIVD") = rsSource("NAZIVD")
    rsDestination("NAZIV55") = rsSource("NAZIV55")
    rsDestination("NAZIVE") = rsSource("NAZIVE")
    rsDestination("NAZIV66") = rsSource("NAZIV66")
    rsDestination("NAZIVF") = rsSource("NAZIVF")
    rsDestination("NAZIV77") = rsSource("NAZIV77")
    rsDestination("NAZIVG") = rsSource("NAZIVG")
    rsDestination("NAZIV88") = rsSource("NAZIV88")
    rsDestination("NAZIVH") = rsSource("NAZIVH")
    If IsNull(rsSource("POCSTANJE")) Then
        rsDestination("POCSTANJE") = 0
    Else
        rsDestination("POCSTANJE") = rsSource("POCSTANJE")
    End If
    If IsNull(rsSource("PRIJEM")) Then
        rsDestination("PRIJEM") = 0
    Else
        rsDestination("PRIJEM") = rsSource("PRIJEM")
    End If
    If IsNull(rsSource("OTPREMA")) Then
        rsDestination("OTPREMA") = 0
    Else
        rsDestination("OTPREMA") = rsSource("OTPREMA")
    End If
    If IsNull(rsSource("PRODAJA")) Then
        rsDestination("PRODAJA") = 0
    Else
        rsDestination("PRODAJA") = rsSource("PRODAJA")
    End If
    If IsNull(rsSource("ZSTANJE")) Then
        rsDestination("ZSTANJE") = 0
    Else
        rsDestination("ZSTANJE") = rsSource("ZSTANJE")
    End If
    'rsDestination("PRIJEM") = rsSource("PRIJEM")
    'rsDestination("OTPREMA") = rsSource("OTPREMA")
    'rsDestination("PRODAJA") = rsSource("PRODAJA")
    'rsDestination("ZSTANJE") = rsSource("ZSTANJE")
    rsDestination.Update
    rsModeli.FindFirst "Sifra=" & rsSource("ARTIKL")
    If Not rsModeli.NoMatch Then
        rsModeli.Edit
        rsModeli("Aktuelan") = True
        rsModeli.Update
    End If
    frmImport.ctrlProgress.Value = rsSource.AbsolutePosition / rsSource.RecordCount * (frmImport.ctrlProgress.Max - frmImport.ctrlProgress.Min)
    DoEvents
    rsSource.MoveNext
Loop

frmImport.Hide

db.Close

End Sub

Public Function FileExists(fn As String) As Boolean
    Dim retVal As Long, FindData As WIN32_FIND_DATA, retval2 As Long
    retVal = FindFirstFile(fn, FindData)
    If retVal <> -1 Then FileExists = True
    retval2 = FindClose(retVal)
End Function


Public Sub KeepOnTop(f As Form)
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

    SetWindowPos f.hWnd, -1, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub

