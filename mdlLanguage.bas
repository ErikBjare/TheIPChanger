Attribute VB_Name = "mdlLanguage"
Option Explicit

'Language array indexes
Public Const LI_NOTFOUND = 0
Public Const LI_HOSTLONG = 1
Public Const LI_UNKNOWNCLIENT = 2
Public Const LI_EXPORTLIST = 3
Public Const LI_ALREADYRUNNING = 4
Public Const LI_ONLINE = 5
Public Const LI_HOST = 6
Public Const LI_PORT = 7
Public Const LI_CHANGE = 8
Public Const LI_CHECK = 9
Public Const LI_ADD = 10
Public Const LI_DELETE = 11
Public Const LI_EXPORT = 12
Public Const LI_MINIMIZE = 13
Public Const LI_ERROR = 14
Public Const LI_NOFILE = 15
Public Const LI_CHOOSE = 16
Public Const LI_CANCEL = 17
Public Const LI_TIMEOUT = 18
Public Const LI_ABOUT = 19

Public sLanguagePath 'Current language
Public sLang() As String 'Language array

Public Sub LoadLanguage(ByVal sPath As String)
    On Error GoTo ErrOut:
    Dim iFile As Integer
    Dim sTemp As String
    Dim sTest() As String
    sLanguagePath = sPath
    sPath = App.Path & "\Languages\" & sPath
    iFile = FreeFile
    Open sPath For Binary As #iFile
    sTemp = Input(FileLen(sPath), iFile) 'Load the whole file content to sTemp
    Close #iFile
    
    sTest = Split(sTemp, vbCrLf)
    If UBound(sTest) = 19 Then 'If the file contained 18 lines then set the real language array
        sLang = Split(sTemp, vbCrLf)
    End If
    UpdateControls
    Exit Sub 'If no errors, the sub ends here
    
ErrOut:
    On Error Resume Next 'Just to not cause an endless loop incase Close #iFile generates an error
    Close #iFile
    DefaultLanguage 'Load the default language
End Sub

Public Sub DefaultLanguage()
    sLanguagePath = "" 'No path
    ReDim sLang(18) 'Redim the array

    sLang(LI_NOTFOUND) = "Tibia not found"
    sLang(LI_HOSTLONG) = "Host too long"
    sLang(LI_UNKNOWNCLIENT) = "Unknown client"
    sLang(LI_EXPORTLIST) = "Export list"
    sLang(LI_ALREADYRUNNING) = "TheIPChanger is already running"
    sLang(LI_ONLINE) = "Server %ip is currently online at port %port"
    sLang(LI_HOST) = "Host"
    sLang(LI_PORT) = "Port"
    sLang(LI_CHANGE) = "Change"
    sLang(LI_CHECK) = "Check"
    sLang(LI_ADD) = "Add"
    sLang(LI_DELETE) = "Delete"
    sLang(LI_EXPORT) = "Export"
    sLang(LI_MINIMIZE) = "Minimize to tray"
    sLang(LI_ERROR) = "Error"
    sLang(LI_NOFILE) = "File not found"
    sLang(LI_CHOOSE) = "Choose language"
    sLang(LI_CANCEL) = "Cancel"
    sLang(LI_TIMEOUT) = "Server %ip is currently not online at port %port"
    sLang(LI_ABOUT) = "About"
    UpdateControls
End Sub

Public Sub UpdateControls()
    'Update controls with the right language
    frmMain.cmdAdd.Caption = sLang(LI_ADD)
    frmMain.cmdChange.Caption = sLang(LI_CHANGE)
    frmMain.cmdCheck.Caption = sLang(LI_CHECK)
    frmMain.cmdDelete.Caption = sLang(LI_DELETE)
    frmMain.cmdExport.Caption = sLang(LI_EXPORT)
    frmMain.lblHost.Caption = sLang(LI_HOST)
    frmMain.lblPort.Caption = sLang(LI_PORT)
    frmMain.chkTray.Caption = sLang(LI_MINIMIZE)
    frmMain.cmdLanguage.Caption = sLang(LI_CHOOSE)
    frmLanguage.Caption = sLang(LI_CHOOSE)
    frmMain.chkTray.Caption = sLang(LI_MINIMIZE)
    frmMain.About.Caption = sLang(LI_ABOUT)
End Sub
