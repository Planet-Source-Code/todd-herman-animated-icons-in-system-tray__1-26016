VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "System Tray test"
   ClientHeight    =   2265
   ClientLeft      =   4395
   ClientTop       =   1545
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   6585
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   150
      Top             =   1065
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuStartAnimation 
         Caption         =   "Start &Animation"
      End
      Begin VB.Menu mnuStopAnimation 
         Caption         =   "&Stop Animation"
      End
      Begin VB.Menu SEP01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore Window"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents SysTray As CSysTray
Attribute SysTray.VB_VarHelpID = -1
Private Sub Form_Load()

    Set SysTray = New CSysTray
    Set SysTray.SourceWindow = Me
    
    SysTray.ChangeIcon App.Path & "\globe.ani"
    SysTray.ToolTip = Me.Caption
    
    SysTray.MinToSysTray
    Timer1.Enabled = True
    
End Sub
Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then
        SysTray.MinToSysTray
        Timer1.Enabled = True
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    Timer1.Enabled = False
    SysTray.RemoveFromSysTray

End Sub
Private Sub SysTray_RButtonUP()

    'Display popup menu when user presses the right mouse button on the System Tray icon
    PopupMenu Me.mnuPopup
    
End Sub
Private Sub mnuRestore_Click()

    'This restores the BIT Manager application
    Timer1.Enabled = False
    
    Me.WindowState = vbNormal
    Me.Show
    App.TaskVisible = True
    SysTray.RemoveFromSysTray
   
End Sub
Private Sub mnuExit_Click()

    SysTray.RemoveFromSysTray
    End
    
End Sub
Private Sub mnuStartAnimation_Click()

    Timer1.Enabled = True
    
End Sub
Private Sub mnuStopAnimation_Click()

    Timer1.Enabled = False
    SysTray.ChangeIcon App.Path & "\globe.ico"
    
End Sub
Private Sub Timer1_Timer()

    SysTray.ChangeIcon App.Path & "\globe.ani"
    
End Sub
