VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Tray Icon Zoom"
   ClientHeight    =   1440
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2970
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   2970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Click on the Exit Button to Zoom to Tray"
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   1170
      Width           =   2925
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "&Show Main"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'UnloadMode possibilities:
'0   The user has chosen the Close command from the Control-menu box on the form.
'1   The Unload method has been invoked from code.
'2   The current Windows-environment session is ending.
'3   The Microsoft Windows Task Manager is closing the application.
'4   An MDI child form is closing because the MDI form is closing.

    If UnloadMode = 0 Then
        SendToTray
    Cancel = True
    End If
    
    Me.Hide
End Sub

Private Sub mnuExit_Click()
    Unload Me
    End
End Sub

Private Sub mnuShow_Click()
    ClearTray
    RestoreWindow
End Sub
