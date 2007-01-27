VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "TheIPChanger"
   ClientHeight    =   6435
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   4770
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   429
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   318
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optVersion 
      Caption         =   "Tibia 7.91 "
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   21
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton About 
      Caption         =   "About"
      Height          =   255
      Left            =   2760
      TabIndex        =   19
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdLanguage 
      Caption         =   "Choose language"
      Height          =   375
      Left            =   2760
      TabIndex        =   18
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CheckBox chkTray 
      Caption         =   "Minimize to tray"
      Height          =   495
      Left            =   2760
      TabIndex        =   17
      Top             =   3480
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   2760
      TabIndex        =   16
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtPort 
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   2760
      TabIndex        =   14
      Text            =   "7171"
      Top             =   360
      Width           =   1935
   End
   Begin VB.ListBox lstPort 
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   2520
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.ListBox lstVersion 
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   2160
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.OptionButton optVersion 
      Caption         =   "Tibia 7.1"
      Height          =   255
      Index           =   4
      Left            =   2760
      TabIndex        =   10
      Top             =   5760
      Width           =   1695
   End
   Begin VB.OptionButton optVersion 
      Caption         =   "Tibia 7.24"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   9
      Top             =   5400
      Width           =   1695
   End
   Begin VB.OptionButton optVersion 
      Caption         =   "Tibia 7.3"
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   8
      Top             =   5040
      Width           =   1695
   End
   Begin VB.OptionButton optVersion 
      Caption         =   "Tibia 7.5"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   7
      Top             =   4680
      Width           =   1695
   End
   Begin VB.OptionButton optVersion 
      Caption         =   "Tibia 7.6"
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   6
      Top             =   4320
      Value           =   -1  'True
      Width           =   1695
   End
   Begin MSWinsockLib.Winsock sckCheck 
      Left            =   360
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txtIP 
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
   Begin VB.ListBox lstList 
      ForeColor       =   &H80000007&
      Height          =   5325
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change"
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version 0.1.0 (BETA)"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Label lblPort 
      Caption         =   "Port:"
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblHost 
      Caption         =   "Host:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   2295
   End
   Begin VB.Menu mnuTray 
      Caption         =   "mnuTray"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Outside of the form borders there are 2 list boxes used for holding port
'and the client protocol used by the server


'Some API declares
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

'I regret naming this constant to FILE
Private Const FILE = "iplist.txt"

'These constants both serve as memory addresses
'and for internal use to handle client versions
Private Const T791 As Long = &H5EFBC0 'Len 96 (+ 1, Zero terminated)
Private Const T76 As Long = &H5EFBC0 'Len 96 (+ 1, Zero terminated)
Private Const T75 As Long = &H5EB998 'Len 96 (+ 1, Zero terminated)
Private Const T73 As Long = &H51CE10 'Len 96 (+ 1, Zero terminated)
Private Const T724 As Long = &H51B7C0 'Len 96 (+ 1, Zero terminated)
Private Const T71 As Long = &H51BF64 'Len 45 (+ 1, Zero terminated)

'The others only need one memadress per version but 7.1 needs 2
Private Const T71Port As Long = &H51BF60


Private lLast As Long 'Last selected listitem
Private lVersion As Long 'Version of the current server

Private Sub ChangeIPs(sIP As String, lPort As Long)
    Dim lHwnd As Long
    Dim lID As Long
    Dim lHandle As Long
    Dim sTemp As String
    Dim lTemp As Long
    Dim lDesktop As Long
    
    'We need this because we are using FindWindowEx to
    'get all tibiaclient childs of the desktop
    lDesktop = GetDesktopWindow
    
    'Just check if there is any tibiaclient at all
    'We could use FindWindow(vbNullString, "Tibia   ")
    'But some tools change the title of the tibia window so that is not recomended
    If FindWindow("tibiaclient", vbNullString) = 0 Then
        MsgBox sLang(LI_NOTFOUND), vbCritical, sLang(LI_ERROR)
        GoTo ErrOut
    End If
    
    'This loop is what makes it a multi ip changer
    Do
        'What does this do then?
        'It gets all tibiaclients open in orded
        'The second parameter tells the function which client we had before
        'if we don't give that parameter it would always return the first client
        lHwnd = FindWindowEx(lDesktop, lHwnd, "tibiaclient", vbNullString)
        
        
        If lHwnd = 0 Then Exit Do
        
        'Get process id from the window
        GetWindowThreadProcessId lHwnd, lID
        
        'Get process handle from the process id
        lHandle = OpenProcess(&H1F0FFF, False, lID)
        
        'If the handle is NULL, display error and exit sub
        If (lHandle = 0) Then
            MsgBox sLang(LI_NOTFOUND), vbCritical, sLang(LI_ERROR)
            GoTo ErrOut
        End If
    
        sTemp = sIP 'String to write in the memory
        lTemp = lPort 'Long integer to write in the memory
        
        'This client only supports hostnames of 45 characters and has only 1 copy of the hostname
        If lVersion = T71 Then
            'Lenght check
            If Len(sTemp) > 45 Then
                MsgBox sLang(LI_HOSTLONG), vbCritical, sLang(LI_ERROR)
                GoTo ErrOut
            End If
            'Change
            WriteProcessMemory lHandle, T71, ByVal sTemp, Len(sTemp) + 1, 0&
            WriteProcessMemory lHandle, T71Port, lTemp, 4, 0&
        'These clients supports hostnames of 96 characters and has 4 copies of the hostname
        ElseIf (lVersion = T791) Or (lVersion = T76) Or (lVersion = T75) Or (lVersion = T73) Or (lVersion = T724) Then
            'Lenght check
            If Len(sTemp) > 96 Then
                MsgBox sLang(LI_HOSTLONG), vbCritical, sLang(LI_ERROR)
                GoTo ErrOut
            End If

            'Change all ips, they are at 4 places with a distance of 112 bytes
            'between
            WriteProcessMemory lHandle, lVersion, ByVal sTemp, Len(sTemp) + 1, 0&
            WriteProcessMemory lHandle, lVersion + 112, ByVal sTemp, Len(sTemp) + 1, 0&
            WriteProcessMemory lHandle, lVersion + 224, ByVal sTemp, Len(sTemp) + 1, 0&
            WriteProcessMemory lHandle, lVersion + 336, ByVal sTemp, Len(sTemp) + 1, 0&
            
            'Change all ports, they are at 4 places with a distance of 112 bytes
            'between, they also are 100 bytes from the corrensponding ip
            WriteProcessMemory lHandle, lVersion + 100, lTemp, 4, 0&
            WriteProcessMemory lHandle, lVersion + 212, lTemp, 4, 0&
            WriteProcessMemory lHandle, lVersion + 324, lTemp, 4, 0&
            WriteProcessMemory lHandle, lVersion + 436, lTemp, 4, 0&
        Else 'This ain't a valid client
            MsgBox sLang(LI_UNKNOWNCLIENT), vbCritical, sLang(LI_ERROR)
            GoTo ErrOut
        End If
        FlashWindow lHwnd, 1
    Loop
    CloseHandle lHandle
    Exit Sub
ErrOut:
    On Error Resume Next
    CloseHandle lHandle
End Sub

'This isn't used
Private Sub ClearIPs()
    'Clear all listboxes
    lstList.Clear
    lstPort.Clear
    lstVersion.Clear
    'Set last selected listitem to none
    lLast = -1
End Sub

Private Sub LoadIPs()
    On Error GoTo NoFile:
    Dim sTemp As String
    Dim iFile As Integer
    Dim lC As Long
    
    iFile = FreeFile 'Get a free filehandle
    Open FILE For Input As #iFile
        Input #iFile, sTemp 'Get first line to check if its a simple or advanced file
        If sTemp = "Advanced mode" Then 'Advanced
            While Not EOF(iFile)
                Input #iFile, sTemp
                lstList.AddItem sTemp
                Input #iFile, sTemp
                lstPort.AddItem sTemp
                Input #iFile, sTemp
                lstVersion.AddItem sTemp
            Wend
        Else 'Simple
            'Don't add blank lines
            If Trim$(sTemp) <> "" Then
                lstList.AddItem sTemp
                lstPort.AddItem "7171" 'Default port 7171
                lstVersion.AddItem Trim$(Str(T76)) 'Default version 7.6
            End If
            While Not EOF(iFile)
                Input #iFile, sTemp
                'Don't add blank lines
                If Trim$(sTemp) <> "" Then
                    lstList.AddItem sTemp
                    lstPort.AddItem "7171" 'Default port 7171
                    lstVersion.AddItem Trim$(Str(T76)) 'Default version 7.6
                End If
            Wend
        End If
NoFile:
    On Error Resume Next 'To not cause endless loop, incase Close #iFile raises error
    Close #iFile
End Sub

Private Sub SaveIPs()
    Dim iC As Integer
    Dim iFile As Integer
    'Save the lists to the iplist file
    iFile = FreeFile
    Open App.Path & "\" & FILE For Output As #iFile
    Print #iFile, "Advanced mode"
    For iC = 0 To lstList.ListCount - 1
        Print #iFile, lstList.List(iC)
        Print #iFile, lstPort.List(iC)
        Print #iFile, Trim$(lstVersion.List(iC))
    Next iC
    Close #iFile
End Sub

Private Sub About_Click()
    frmAbout.Show 1
End Sub

Private Sub cmdAdd_Click()
    'If txtIP is empty or only contains spaces, exit sub
    If Trim$(txtIP.Text) = "" Then Exit Sub
        'Add all fields to their listboxes
        lstList.AddItem Trim$(txtIP.Text)
        lstPort.AddItem txtPort.Text
        lstVersion.AddItem Str$(lVersion)
    'Last selected item = NONE
    lLast = -1
    txtIP.Text = "" 'Clear host files
    txtPort.Text = "7171" 'Reset port field
End Sub

Private Sub cmdChange_Click()
    'Call the function to change the ips and ports
    ChangeIPs Trim$(txtIP.Text), CLng(Val(txtPort.Text))
End Sub

Private Sub cmdCheck_Click()
    'Close the socket, incase it is not closed already
    sckCheck.Close
    'Attempt to connect to the host specified in txtIP
    sckCheck.Connect Trim$(txtIP.Text), CLng(Val(txtPort.Text))
End Sub

Private Sub cmdDelete_Click()

    Dim lTemp As Long
    'If no listitem is selected, exit sub
    If lstList.ListIndex <= -1 Then Exit Sub
    lTemp = lstList.ListIndex 'Set lTemp to the index of the item that is going to be deleted
     'Remove currentitem from all listboxes
    lstPort.RemoveItem lstList.ListIndex
    lstVersion.RemoveItem lstList.ListIndex
    lstList.RemoveItem lstList.ListIndex
    
    'If the deleted item was the last item, lower lTemp by one
    If lstList.ListCount = lTemp Then lTemp = lTemp - 1
    
    'Last selected item = lTemp
    lLast = lTemp
    
    'If lTemp is a valid item index, load fields from listboxes
    If lTemp > -1 Then
        txtIP.Text = lstList.List(lTemp)
        txtPort.Text = lstPort.List(lTemp)
        lVersion = CLng(lstVersion.List(lTemp))
    Else 'lTemp is an invalid index
        'Reset ip field and port field
        txtIP.Text = ""
        txtPort.Text = "7171"
    End If
    lstList.ListIndex = lTemp 'Select newly loaded item
    
    'Set version optionbuttons to the newly loaded version
    Select Case lVersion
    Case T76: optVersion(0).Value = True
    Case T75: optVersion(1).Value = True
    Case T73: optVersion(2).Value = True
    Case T724: optVersion(3).Value = True
    Case T71: optVersion(4).Value = True
    End Select
End Sub



Private Sub cmdExport_Click()
    Dim iC As Integer
    Dim iFile As Integer
    Dim sPath As String
    
    'Prompt for a filename to export the list to
    sPath = InputBox(sLang(LI_EXPORTLIST), sLang(LI_EXPORTLIST), "export.txt")
    
    'If cancel was selected, exit sub
    If sPath = "" Then Exit Sub
    
    iFile = FreeFile
    Open App.Path & "\" & sPath For Output As #iFile
    For iC = 0 To lstList.ListCount - 1
        Print #iFile, lstList.List(iC) 'Save all ips, nothing else
    Next iC
    Close #iFile
End Sub


Private Sub cmdLanguage_Click()
    'Show frmLanguage, the 1 is there to disable the main form while the
    'language form is open
    frmLanguage.Show 1
End Sub

Private Sub Form_Load()
    'Get the language setting, and load language
    'This needs to be first so the errormessages can be displayed correctly
    LoadLanguage GetSetting("TheIPChanger", "Settings", "Language", "")
    
    'Check if a previous instance of the app is running
    'This is kind of lame since you only have to move the exe to another location
    'But it's not that important really
    If App.PrevInstance Then
        MsgBox sLang(LI_ALREADYRUNNING), vbExclamation
        End
    End If
    
    'Get the tray setting, and load it to the chkTray.Value property
    chkTray.Value = GetSetting("TheIPChanger", "Settings", "Tray", 0)
    
    
    lLast = -1 'Set last selected ip from list to none
    lVersion = T76 'Default version = 7.6
    LoadIPs 'Load iplist.txt
End Sub




Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Mouse events with the tray icon
Select Case X
    'Case meMouseMove
    'Case meLeftUp
    'Case meLeftDown
    Case meLeftDbClick
        TrayDelete 'Delete icon from tray
        Me.WindowState = vbNormal 'Set window state to normal
        Me.Show 'Show the form
    'Case meMiddleUp
    'Case meMiddleDown
    'Case meMiddleDbClick
    'Case meRightUp
        'Show a menu if the tray icon is rightclicked
        'You got to make the menu first, user the menu editor
        'PopupMenu mnuTrayMenu
    'Case meRightDown
    'Case meRightDbClick
End Select
End Sub

Private Sub Form_Resize()
    'This event is fired when the form is resized
    'this includes maximizing and minimizing
    
    'Check if the minimize to tray is checked
    If chkTray.Value = vbChecked Then
        'Check if the window was minimized
        If Me.WindowState = vbMinimized Then
            TrayAdd Me.hwnd, Me.Icon, "TheIPChanger", meMouseMove 'Add icon to tray
            Me.Hide 'Hide form from taskbar
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrOut: 'If any errors, instantly go to ErrOut, and the app will end
    'Save the settings, if you want to experiment with settings  in your own app,
    'try out the following functions: GetSetting, SaveSetting, DeleteSetting
    SaveSetting "TheIPChanger", "Settings", "Tray", chkTray.Value
    SaveSetting "TheIPChanger", "Settings", "Language", sLanguagePath
    
    TrayDelete 'Delete icon from tray
    'Check if there is a last selected list item
    If lLast > -1 Then
        'Save the current fields to the last selected list item
        lstList.List(lLast) = txtIP.Text
        lstPort.List(lLast) = txtPort.Text
        lstVersion.List(lLast) = Str$(lVersion)
    End If
    'Save the ips to iplist.txt
    SaveIPs
ErrOut:
    End 'End the app
End Sub

Private Sub lstList_Click()
    'Check if there is a last selected list item
    If lLast > -1 Then
        'Save the current fields to the last selected list item
        lstList.List(lLast) = txtIP.Text
        lstPort.List(lLast) = txtPort.Text
        lstVersion.List(lLast) = Str$(lVersion)
    End If
    'Set last selected list item to current list item
    lLast = lstList.ListIndex
    
    'Load ip and port from the listboxes
    txtIP.Text = lstList.List(lstList.ListIndex)
    txtPort.Text = lstPort.List(lstList.ListIndex)
    
    'Load the version from the version listbox
    lVersion = CLng(Val(lstVersion.List(lstList.ListIndex)))
    
    'Set the optionbuttons to the newly loaded version
    Select Case lVersion
    Case T791: optVersion(5).Value = True
    Case T76: optVersion(0).Value = True
    Case T75: optVersion(1).Value = True
    Case T73: optVersion(2).Value = True
    Case T724: optVersion(3).Value = True
    Case T71: optVersion(4).Value = True
    End Select
End Sub

Private Sub optVersion_Click(Index As Integer)
    'As you can see I use the optionbutton caption
    'to detarmain which version was choosen
    'If you are going to change the captions, you will have to change here too
    Select Case optVersion(Index).Caption
    Case "Tibia 7.91": lVersion = T791
    Case "Tibia 7.6": lVersion = T76
    Case "Tibia 7.5": lVersion = T75
    Case "Tibia 7.3": lVersion = T73
    Case "Tibia 7.24": lVersion = T724
    Case "Tibia 7.1": lVersion = T71
    End Select
End Sub

Private Sub sckCheck_Connect()
    'This fires if the sckCheck socket is connected to a server

    'Replace the variables in the language file with their real values, and display a msgbox
    MsgBox Replace(Replace(sLang(LI_ONLINE), "%ip", txtIP.Text), "%port", CLng(Val(txtPort.Text))), vbInformation, "TheIPChanger"
    
    'Close the socket, we are done checking the online status
    sckCheck.Close
End Sub

Private Sub sckCheck_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Check if connection timed out
    If Number = 10060 Then
        'Replace the variables in the language file with their real values, and display a msgbox
        MsgBox Replace(Replace(sLang(LI_TIMEOUT), "%ip", txtIP.Text), "%port", CLng(Val(txtPort.Text))), vbInformation, "TheIPChanger"
    
        'Close the socket, we are done checking the online status
        sckCheck.Close
    End If
End Sub
