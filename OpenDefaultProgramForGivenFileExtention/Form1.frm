VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Open Default Program for Given Extention"
   ClientHeight    =   1035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1035
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOpen 
      Caption         =   "Open Deafult Program"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox txtFileExt 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the File Extention E.g: doc, rtf, exe, bas, vbp, mp3"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'##############'
'### API'S ####'
'##############'
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub CmdOpen_Click()
On Error GoTo FileAccessErr 'Create the Temp File
Open "C:\Temp." & txtFileExt.Text For Append As #1
Close #1 'Close the Temp File
'Open the Temp File
ShellExecute hwnd, "Open", "C:\Temp." & txtFileExt.Text, _
vbNull, "C:\", 1
Sleep 3000 'Sleep for 3 Seconds to Allow time for Loading
Kill "C:\Temp." & txtFileExt.Text 'Delete the Temp File
Exit Sub
'Error Handling
FileAccessErr:
    MsgBox "The File Extention ." & txtFileExt.Text & " Was not found in the File Assosiation List", vbInformation, "File Extention not Found"
    Err.Clear
End Sub

'#########################################
'The Function Can be Done Easier by the
'Following function which is basically
'the same but but without all the CRAP
'#########################################

'Private Sub CmdOpen_Click()
'Open "C:\Temp." txtfileexe.text for append as #1
'Close #1
'ShellExecute hWnd, "Open", "C:\Temp." & txtFileExt.Text, _
'vbNull, "C:\", 1
'Sleep 3000
'Kill "C:\Temp." txtFileExt.Text
'End sub
