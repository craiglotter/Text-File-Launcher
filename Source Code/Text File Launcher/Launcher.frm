VERSION 5.00
Begin VB.Form Launcher 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text File Launcher"
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6210
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "Launcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
Dim path As String
path = App.path & "\readme.txt"
path = Replace(path, "\\", "\")
If Dir$(path) <> "" Then
Label1.Caption = "Success: """ & path & """ has been passed to the Shell Executor."
Dim pid As Double
        pid = ShellExecute(0&, vbNullString, "notepad", path, vbNullString, vbNormalFocus)
       
Unload Me
Else
Label1.Caption = "Error Encountered: """ & path & """ cannot be located in the current file system."
End If


End Sub
