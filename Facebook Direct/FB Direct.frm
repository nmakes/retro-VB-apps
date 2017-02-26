VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Facebook Direct"
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   15240
   Icon            =   "FB Direct.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleWidth      =   15240
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   2040
      Top             =   2640
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8400
      Top             =   3480
   End
   Begin SHDocVwCtl.WebBrowser wb1 
      Height          =   10455
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   15255
      ExtentX         =   26908
      ExtentY         =   18441
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Menu mnweb 
      Caption         =   "&Facebook"
      Begin VB.Menu mnhome 
         Caption         =   "&Home"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mncf 
         Caption         =   "&Creator's Profile"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnmessages 
         Caption         =   "&Messages"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnstop 
      Caption         =   "&Stop"
   End
   Begin VB.Menu mnfip 
      Caption         =   "&Facebook Inventor's Page"
   End
   Begin VB.Menu mnabout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnquit 
      Caption         =   "&Quit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
wb1.Navigate2 ("http://www.facebook.com")
End Sub

Private Sub Form_Resize()
wb1.Height = Form1.Height - 725
wb1.Width = Form1.Width - 75
End Sub

Private Sub mnabout_Click()
MsgBox "EVERYTHING MADE BY NAVEEN VENKAT. Copyright (c) 2010 Naveen | Spotlite Productions | www.funwithspotlite.webs.com"
End Sub

Private Sub mncf_Click()
wb1.Navigate2 ("http://www.facebook.com/profile.php?id=1638176623")
End Sub

Private Sub mnfip_Click()
wb1.Navigate2 ("http://www.facebook.com/markzuckerberg")
End Sub

Private Sub mnhome_Click()
wb1.Navigate2 ("www.facebook.com")
End Sub

Private Sub mnmessages_Click()
wb1.Navigate2 ("http://www.facebook.com/?sk=messages")
End Sub

Private Sub mnquit_Click()
MsgBox "Thank you for using Facebook Direct"
End
End Sub

Private Sub mnstop_Click()
wb1.Stop
End Sub

Private Sub Timer1_Timer()
If wb1.Busy = True Then
Me.Caption = "Working... - Facebook Direct"
Else
Me.Caption = wb1.LocationName & " - Facebook Direct"
End If
End Sub

Private Sub Timer2_Timer()
On Error GoTo errbug
Exit Sub
errbug:
MsgBox "Sorry, but some error has occured and we need to close. Try again later or please report this error in the Spotlite Forums"
End Sub

