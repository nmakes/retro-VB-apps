VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form form1 
   BackColor       =   &H00808080&
   Caption         =   "NMakes Surf 3.4"
   ClientHeight    =   9900
   ClientLeft      =   885
   ClientTop       =   1935
   ClientWidth     =   14550
   BeginProperty Font 
      Name            =   "Kartika"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Surfer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9900
   ScaleWidth      =   14550
   Begin VB.Frame Frame2 
      Caption         =   " [ History ] "
      Height          =   6735
      Left            =   3240
      TabIndex        =   6
      Top             =   1200
      Width           =   8775
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   6960
         Width           =   8535
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Clear History"
         Height          =   375
         Left            =   3120
         TabIndex        =   10
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton add 
         Caption         =   "Add to History"
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   120
         Width           =   1335
      End
      Begin VB.ListBox List1 
         Height          =   6000
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   8
         Top             =   600
         Width           =   8535
      End
      Begin VB.CommandButton Command19 
         Caption         =   "x"
         Height          =   375
         Left            =   7800
         TabIndex        =   7
         Top             =   120
         Width           =   855
      End
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   5
      Top             =   9600
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   529
      Style           =   1
      SimpleText      =   "Done"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox address 
      Height          =   360
      Left            =   1920
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   120
      Width           =   10575
   End
   Begin VB.Timer Timer2 
      Interval        =   120
      Left            =   6000
      Top             =   0
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5520
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "STOP"
      Height          =   495
      Left            =   13440
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go!"
      Height          =   495
      Left            =   12600
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin SHDocVwCtl.WebBrowser wb1 
      Height          =   8895
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   14535
      ExtentX         =   25638
      ExtentY         =   15690
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
   Begin VB.Image sico 
      Height          =   480
      Left            =   30
      Picture         =   "Surfer.frx":030A
      Top             =   90
      Width           =   480
   End
   Begin VB.Image l4 
      Height          =   480
      Left            =   60
      Picture         =   "Surfer.frx":0614
      Top             =   60
      Width           =   480
   End
   Begin VB.Image l3 
      Height          =   480
      Left            =   60
      Picture         =   "Surfer.frx":091E
      Top             =   60
      Width           =   480
   End
   Begin VB.Image l2 
      Height          =   480
      Left            =   60
      Picture         =   "Surfer.frx":0C28
      Top             =   60
      Width           =   480
   End
   Begin VB.Image l1 
      Height          =   480
      Left            =   60
      Picture         =   "Surfer.frx":0F32
      Top             =   60
      Width           =   480
   End
   Begin VB.Menu mnweb 
      Caption         =   "&Web"
      Begin VB.Menu mngo 
         Caption         =   "&Go"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnstop 
         Caption         =   "&Stop"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnrefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnpreloads 
         Caption         =   "&Preloads"
         Begin VB.Menu mngoogle 
            Caption         =   "&Google"
            Shortcut        =   {F6}
         End
         Begin VB.Menu mngmail 
            Caption         =   "&Gmail"
            Shortcut        =   {F7}
         End
         Begin VB.Menu mnfacebook 
            Caption         =   "&Facebook"
            Shortcut        =   {F8}
         End
         Begin VB.Menu mntwitter 
            Caption         =   "&Twitter"
         End
         Begin VB.Menu mnyoutube 
            Caption         =   "&Youtube"
         End
         Begin VB.Menu mnorkut 
            Caption         =   "&Orkut"
         End
      End
      Begin VB.Menu mnquit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnthemes 
      Caption         =   "&Themes"
      Begin VB.Menu mnwhite 
         Caption         =   "&White"
      End
      Begin VB.Menu mnblack 
         Caption         =   "&Black"
      End
      Begin VB.Menu mnblue 
         Caption         =   "&Blue"
      End
      Begin VB.Menu mngreen 
         Caption         =   "&Green"
      End
      Begin VB.Menu mnred 
         Caption         =   "&Red"
      End
      Begin VB.Menu mnpink 
         Caption         =   "&Pink"
      End
      Begin VB.Menu mngrey 
         Caption         =   "&Grey (Default)"
      End
   End
   Begin VB.Menu mntools 
      Caption         =   "&Tools"
      Begin VB.Menu mnhistory 
         Caption         =   "&History"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mncalender 
         Caption         =   "&Calender"
      End
      Begin VB.Menu mnlocal 
         Caption         =   "&Local Device"
         Begin VB.Menu mnA 
            Caption         =   "&A:\"
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu mnc 
            Caption         =   "&C:\"
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu mnd 
            Caption         =   "&D:\"
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu mne 
            Caption         =   "&E:\"
            Shortcut        =   ^{F4}
         End
         Begin VB.Menu mnf 
            Caption         =   "&F:\"
            Shortcut        =   ^{F5}
         End
         Begin VB.Menu mng 
            Caption         =   "&G:\"
            Shortcut        =   ^{F6}
         End
         Begin VB.Menu mnh 
            Caption         =   "&H:\"
            Shortcut        =   ^{F7}
         End
         Begin VB.Menu mni 
            Caption         =   "&I:\"
            Shortcut        =   ^{F8}
         End
      End
   End
   Begin VB.Menu mnhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnabout 
         Caption         =   "&About"
      End
      Begin VB.Menu mncredits 
         Caption         =   "&Credits"
      End
      Begin VB.Menu mnlicence 
         Caption         =   "&LICENCE"
      End
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub add_Click()
If address.Text <> Text1.Text Then
List1.AddItem (address.Text)
End If
End Sub



Private Sub address_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Then
Command1_Click
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
wb1.Navigate2 (address.Text)
Me.Caption = wb1.LocationName & " - NMakes Surf " & App.Major & "." & App.Minor & "." & App.Revision
address.AddItem (address.Text)
End Sub


Private Sub Command19_Click()
Frame2.Visible = False
End Sub

Private Sub Command2_Click()
wb1.Stop
Me.Caption = wb1.LocationName & " - NMakes Surf " & App.Major & "." & App.Minor & "." & App.Revision
End Sub



Private Sub Command21_Click()
List1.Clear
End Sub

Private Sub Command3_Click()
On Error GoTo err
wb1.GoBack
Exit Sub
err:
End Sub

Private Sub Command4_Click()
On Error GoTo errf
wb1.GoForward
Exit Sub
errf:
End Sub



Private Sub Form_Load()
address.Text = "about:blank"
wb1.Navigate2 (address.Text)
wb1.Offline = False
Frame2.Visible = False
End Sub

Private Sub Form_Resize()
On Error Resume Next
wb1.Height = form1.Height - 1695
wb1.Width = form1.Width - 100
If form1.Width > 7860 Then
address.Width = form1.Width - 4080
Command1.Left = form1.Width - 2055
Command2.Left = form1.Width - 1215
End If
End Sub

Private Sub mnA_Click()
wb1.Navigate2 ("file:///A:/")
End Sub

Private Sub mnabout_Click()
frmAbout.Show
End Sub

Private Sub mnblack_Click()
Me.BackColor = &H0&

End Sub

Private Sub mnblue_Click()
Me.BackColor = &HFF0000

End Sub

Private Sub mnc_Click()
wb1.Navigate2 ("file:///C:/")

End Sub

Private Sub mncafw_Click()
address.Text = "http://www.webs.com"
Command1_Click
End Sub

Private Sub mncalender_Click()
Form3.Show
End Sub

Private Sub mncredits_Click()
MsgBox "Everything Made in VB6 By NAVEEN VENKAT. For more information on running this program, please read the manual. Updates will be available on the website www.nmakes.webs.com"
End Sub

Private Sub mnd_Click()
wb1.Navigate2 ("file:///D:/")

End Sub

Private Sub mne_Click()
wb1.Navigate2 ("file:///E:/")

End Sub

Private Sub mnf_Click()

wb1.Navigate2 ("file:///F:/")
End Sub

Private Sub mnfacebook_Click()
address.Text = "http://www.facebook.com"
Command1_Click
End Sub

Private Sub mng_Click()
wb1.Navigate2 ("file:///G:/")

End Sub

Private Sub mngmail_Click()
address.Text = "http://www.gmail.com"
Command1_Click
End Sub

Private Sub mngo_Click()
Command1_Click
End Sub



Private Sub mngoogle_Click()
address.Text = "http://www.google.com"
Command1_Click
End Sub

Private Sub mngreen_Click()
Me.BackColor = &HFF00&

End Sub

Private Sub mngrey_Click()
Me.BackColor = &H808080

End Sub

Private Sub mnhistory_Click()
Frame2.Visible = True
End Sub

Private Sub mni_Click()
wb1.Navigate2 ("file:///I:/")
End Sub

Private Sub mnlicence_Click()
Form2.Show
End Sub

Private Sub mnorkut_Click()
address.Text = "http://www.orkut.com"
Command1_Click
End Sub

Private Sub mnpink_Click()
Me.BackColor = &HFF00FF

End Sub


Private Sub mnquit_Click()
End
End Sub

Private Sub mnred_Click()
Me.BackColor = &HFF&
End Sub

Private Sub mnrefresh_Click()
wb1.Refresh2
End Sub

Private Sub mnstop_Click()
wb1.Stop
End Sub

Private Sub mntwitter_Click()
address.Text = "http://www.twitter.com"
Command1_Click
End Sub

Private Sub mnwhite_Click()
Me.BackColor = &HFFFFFF

End Sub

Private Sub mnyoutube_Click()
address.Text = "http://www.youtube.com"
Command1_Click
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If wb1.Busy = True Then
Me.Caption = "Working..."
Timer2.Enabled = True
sb.SimpleText = "[" & Time & "]" & " [" & Date & "] Busy... " & address.Text
sico.Visible = False
Else

Me.Caption = wb1.LocationName & " - NMakes Surf " & App.Major & "." & App.Minor & "." & App.Revision
Timer2.Enabled = False
l1.Visible = False
l2.Visible = False
l3.Visible = False
l4.Visible = True
sb.SimpleText = "[" & Time & "]" & " [" & Date & "] Done"
sico.Visible = True
End If
End Sub

Private Sub Timer2_Timer()
sico.Visible = False
If l1.Visible = True Then
l1.Visible = False
l2.Visible = True
l3.Visible = False
l4.Visible = False
Else
If l2.Visible = True Then
l1.Visible = False
l2.Visible = False
l3.Visible = True
l4.Visible = False
Else
If l3.Visible = True Then
l1.Visible = False
l2.Visible = False
l3.Visible = False
l4.Visible = True
Else
If l4.Visible = True Then
l1.Visible = True
l2.Visible = False
l3.Visible = False
l4.Visible = False
End If
End If
End If
End If
End Sub


Private Sub wb1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
Text1.Text = wb1.LocationURL
End Sub




Private Sub wb1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
address.Text = wb1.LocationURL
Me.Caption = wb1.LocationName
address.AddItem (wb1.LocationURL)
add_Click
End Sub
