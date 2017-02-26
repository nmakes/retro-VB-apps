VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   BackColor       =   &H80000003&
   Caption         =   "Surf"
   ClientHeight    =   8655
   ClientLeft      =   1590
   ClientTop       =   2010
   ClientWidth     =   14655
   Icon            =   "surfos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   14655
   Begin SHDocVwCtl.WebBrowser b 
      Height          =   7935
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   14655
      ExtentX         =   25850
      ExtentY         =   13996
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
      Location        =   "http:///"
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   4920
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   5400
      Top             =   0
   End
   Begin VB.ComboBox addr 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1320
      TabIndex        =   0
      Text            =   "about:blank"
      Top             =   120
      Width           =   12015
   End
   Begin VB.Image bstop 
      Height          =   480
      Left            =   13440
      Picture         =   "surfos.frx":0CCA
      ToolTipText     =   "Stop"
      Top             =   120
      Width           =   480
   End
   Begin VB.Image reload 
      Height          =   480
      Left            =   14040
      Picture         =   "surfos.frx":1994
      Top             =   120
      Width           =   480
   End
   Begin VB.Image d7 
      Height          =   480
      Left            =   14040
      Picture         =   "surfos.frx":1C9E
      Top             =   120
      Width           =   480
   End
   Begin VB.Image d6 
      Height          =   480
      Left            =   14040
      Picture         =   "surfos.frx":2968
      Top             =   120
      Width           =   480
   End
   Begin VB.Image d5 
      Height          =   480
      Left            =   14040
      Picture         =   "surfos.frx":3632
      Top             =   120
      Width           =   480
   End
   Begin VB.Image d4 
      Height          =   480
      Left            =   14040
      Picture         =   "surfos.frx":42FC
      Top             =   120
      Width           =   480
   End
   Begin VB.Image d3 
      Height          =   480
      Left            =   14040
      Picture         =   "surfos.frx":4FC6
      Top             =   120
      Width           =   480
   End
   Begin VB.Image d2 
      Height          =   480
      Left            =   14040
      Picture         =   "surfos.frx":5C90
      Top             =   120
      Width           =   480
   End
   Begin VB.Image d1 
      Height          =   480
      Left            =   14040
      Picture         =   "surfos.frx":695A
      Top             =   120
      Width           =   480
   End
   Begin VB.Image frwd 
      Height          =   480
      Left            =   720
      Picture         =   "surfos.frx":7624
      Top             =   120
      Width           =   480
   End
   Begin VB.Image back 
      Height          =   480
      Left            =   120
      Picture         =   "surfos.frx":792E
      Top             =   120
      Width           =   480
   End
   Begin VB.Menu mnsoftsol 
      Caption         =   "&Softsol"
      Begin VB.Menu mnhome 
         Caption         =   "&Home"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function surfcrashhandler()
b.Stop
End Function

Private Function elementsplacing()
On Error Resume Next
If Me.Width > 5535 Then
addr.Width = Me.Width - 2760
reload.Left = Me.Width - 735
bstop.Left = Me.Width - 1335
d1.Left = Me.Width - 735
d2.Left = Me.Width - 735
d3.Left = Me.Width - 735
d4.Left = Me.Width - 735
d5.Left = Me.Width - 735
d6.Left = Me.Width - 735
d7.Left = Me.Width - 735
End If
b.Width = Me.Width - 125
b.Height = Me.Height - 1520
End Function

Private Function addrnotfound()
Dim msg As Double
msg = MsgBox("The address was not found! Try something else.", vbMsgBoxRight, "oops.. not found!")
b.Stop
End Function

Private Function animate()
Timer1.Enabled = True
reload.Visible = False
End Function

Private Function stopanimate()
Timer1.Enabled = False
Call d1vis
reload.Visible = True
End Function

Private Function d1vis()
d1.Visible = True
d2.Visible = False
d3.Visible = False
d4.Visible = False
d5.Visible = False
d6.Visible = False
d7.Visible = False

End Function

Private Function navi()
b.Navigate2 (addr.Text)
End Function

Private Function surfanim()
If d1.Visible = True Then
d1.Visible = False
d2.Visible = True
d3.Visible = False
d4.Visible = False
d5.Visible = False
d6.Visible = False
d7.Visible = False

ElseIf d2.Visible = True Then
d1.Visible = False
d2.Visible = False
d3.Visible = True
d4.Visible = False
d5.Visible = False
d6.Visible = False
d7.Visible = False

ElseIf d3.Visible = True Then
d1.Visible = False
d2.Visible = False
d3.Visible = False
d4.Visible = True
d5.Visible = False
d6.Visible = False
d7.Visible = False

ElseIf d4.Visible = True Then
d1.Visible = False
d2.Visible = False
d3.Visible = False
d4.Visible = False
d5.Visible = True
d6.Visible = False
d7.Visible = False

ElseIf d5.Visible = True Then
d1.Visible = False
d2.Visible = False
d3.Visible = False
d4.Visible = False
d5.Visible = False
d6.Visible = True
d7.Visible = False

ElseIf d6.Visible = True Then
d1.Visible = False
d2.Visible = False
d3.Visible = False
d4.Visible = False
d5.Visible = False
d6.Visible = False
d7.Visible = True

ElseIf d7.Visible = True Then
d1.Visible = True
d2.Visible = False
d3.Visible = False
d4.Visible = False
d5.Visible = False
d6.Visible = False
d7.Visible = False
End If

End Function


Private Sub addr_KeyPress(KeyAscii As Integer)
On Error GoTo addressnotfound

If KeyAscii = vbKeyReturn Then
Call navi
Exit Sub

addressnotfound:
Call addrnotfound
End If

End Sub

Private Sub b_DownloadComplete()
addr.AddItem (addr.Text)
End Sub



Private Sub b_StatusTextChange(ByVal Text As String)
addr.Text = b.LocationURL
Me.Caption = b.LocationName & " - SurfOS " & App.Major & "." & App.Minor & "." & App.Revision
End Sub



Private Sub back_Click()
On Error Resume Next
b.GoBack
End Sub

Private Sub bstop_Click()
b.Stop
End Sub



Private Sub Form_Load()
Call stopanimate
b.Navigate2 ("about:blank")
Call elementsplacing
End Sub

Private Sub Form_Resize()
Call elementsplacing
End Sub

Private Sub frwd_Click()
On Error Resume Next
b.GoForward
End Sub


Private Sub mnhome_Click()
b.Navigate2 ("http://www.softsolmcs.weebly.com")
End Sub

Private Sub reload_Click()
On Error Resume Next
Call navi
End Sub

Private Sub stop_Click()
b.Stop
End Sub

Private Sub Timer1_Timer()
Call surfanim
End Sub

Private Sub Timer2_Timer()
If b.Busy = True Then
Call animate
Me.Caption = "Busy..." & b.LocationName & " - SurfOS " & App.Major & "." & App.Minor & "." & App.Revision
Else
Call stopanimate
Me.Caption = b.LocationName & " - SurfOS " & App.Major & "." & App.Minor & "." & App.Revision
End If
End Sub

