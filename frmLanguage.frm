VERSION 5.00
Begin VB.Form frmLanguage 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Choose language"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   1815
   End
   Begin VB.FileListBox lstLanguage 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me 'Close language form
End Sub

Private Sub Form_Load()
    'If the directory Langauges does not exist an error will be raised
    'this line prevents that from happening
    On Error Resume Next
    
    'Set the filelistbox to show all files in the language directory
    lstLanguage.Path = App.Path & "\Languages\"
End Sub

Private Sub lstLanguage_Click()
    'Load language from selected item
    LoadLanguage lstLanguage.FileName
End Sub
