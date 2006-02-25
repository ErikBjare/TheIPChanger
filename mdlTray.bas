Attribute VB_Name = "mdlTray"
Option Explicit
'This is the only part of the program that I haven't done myself
'Who remembers all IP constants anyways


'[Tray Constants]
Public Const NIF_MESSAGE    As Long = &H1     'Message
Public Const NIF_ICON       As Long = &H2     'Icon
Public Const NIF_TIP        As Long = &H4     'TooTipText
Public Const NIM_ADD        As Long = &H0     'Add to tray
Public Const NIM_MODIFY     As Long = &H1     'Modify
Public Const NIM_DELETE     As Long = &H2     'Delete From Tray

Public Const meMouseMove    As Long = &H200     'On Mousemove
Public Const meLeftUp       As Long = &H202     'Left Button Mouse Up
Public Const meLeftDown     As Long = &H201     'Left Button MouseDown
Public Const meLeftDbClick  As Long = &H203     'Left Button Double Click
Public Const meRightUp      As Long = &H205     'Right Button Up
Public Const meRightDown    As Long = &H204     'Right Button Down
Public Const meRightDbClick As Long = &H206     'Right Button Double Click
Public Const meMiddleUp     As Long = &H208     'Middle Button Up
Public Const meMiddleDown   As Long = &H207     'Middle Button Down
Public Const meMiddleDbClick As Long = &H209    'Middle Button Double Click
    
'[Type NotifyIconData For Tray Icon]
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type


    

'[Modify Items]
Public Enum ModifyItemEnum
    ToolTip = 1             'Modify ToolTip
    Icon = 2                'Modify Icon
End Enum

'[API]
Private TrayIcon As NOTIFYICONDATA
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'[Add to Tray]
Public Sub TrayAdd(hwnd As Long, Icon As Picture, ToolTip As String, ReturnCallEvent As Long)
    With TrayIcon
        .cbSize = Len(TrayIcon)
        .hwnd = hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = ReturnCallEvent
        .hIcon = Icon
        .szTip = ToolTip & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, TrayIcon
End Sub

'[Remove From tray]
Public Sub TrayDelete()
    Shell_NotifyIcon NIM_DELETE, TrayIcon
End Sub

'[Modify the tray]
Public Sub TrayModify(Item As ModifyItemEnum, vNewValue As Variant)
    Select Case Item
        Case ToolTip
            TrayIcon.szTip = vNewValue & vbNullChar
        Case Icon
            TrayIcon.hIcon = vNewValue
    End Select
    Shell_NotifyIcon NIM_MODIFY, TrayIcon
End Sub


