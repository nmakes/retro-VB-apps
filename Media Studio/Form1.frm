VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   Caption         =   "Media Studio"
   ClientHeight    =   5310
   ClientLeft      =   3720
   ClientTop       =   2790
   ClientWidth     =   8130
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   8130
   Begin VB.Frame Frame1 
      Caption         =   "Open"
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton Command2 
         Caption         =   "Open"
         Height          =   315
         Left            =   5880
         TabIndex        =   8
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   5775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "x"
         Height          =   255
         Left            =   6360
         TabIndex        =   6
         Top             =   0
         Width           =   255
      End
      Begin VB.FileListBox File1 
         Height          =   1260
         Left            =   4320
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
      Begin VB.DirListBox Dir1 
         Height          =   1215
         Left            =   2040
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Select the drive above. Then select the folder from the right. Then select the file."
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer mp1 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   14420
      _cy             =   9340
   End
   Begin VB.Menu mnfile 
      Caption         =   "&File"
      Begin VB.Menu mnopen 
         Caption         =   "&Open"
      End
   End
   Begin VB.Menu mnabout 
      Caption         =   "&About"
      Begin VB.Menu mncredits 
         Caption         =   "&Credits"
      End
      Begin VB.Menu mnontheweb 
         Caption         =   "&On The Web"
      End
   End
   Begin VB.Menu mnquit 
      Caption         =   "&Quit"
      Begin VB.Menu mnays 
         Caption         =   "&Are you sure?"
         Begin VB.Menu mnyes 
            Caption         =   "&Yes"
         End
         Begin VB.Menu mnno 
            Caption         =   "&No"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Frame1.Visible = False
End Sub

Private Sub Command2_Click()
mp1.URL = Text1.Text
Frame1.Visible = False
End Sub

Private Sub Dir1_Change()
Text1.Text = ""
Text1.Text = Dir1
File1.Path = Dir1
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Text1.Text = ""
Dir1.Path = Drive1
End Sub

Private Sub File1_Click()
Text1.Text = ""
Text1.Text = Dir1 & "\" & File1
End Sub

Private Sub Form_Resize()
On Error Resume Next
mp1.Width = Form1.Width - 120
mp1.Height = Form1.Height - 830
End Sub

Private Sub mncredits_Click()
frmAbout.Show
End Sub

Private Sub mnontheweb_Click()
MsgBox "Website: http://www.nmakes.webs.com"
End Sub

Private Sub mnopen_Click()
Frame1.Visible = True
End Sub

Private Sub mnyes_Click()
End
End Sub

