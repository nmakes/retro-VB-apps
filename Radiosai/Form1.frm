VERSION 5.00
Object = "{1339B53E-3453-11D2-93B9-000000000000}#1.0#0"; "mozctl.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Radiosai Live Stream"
   ClientHeight    =   6660
   ClientLeft      =   8190
   ClientTop       =   2130
   ClientWidth     =   5400
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1440
      Top             =   3000
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6405
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MOZILLACONTROLLibCtl.MozillaBrowser MozillaBrowser1 
      Height          =   7095
      Left            =   0
      OleObjectBlob   =   "Form1.frx":0CCA
      TabIndex        =   0
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Dim k As String
k = MsgBox("Are you sure?", vbYesNo, "Quitting?...")
If k = vbYes Then
End
End If
End Sub

Private Sub Form_Load()
MozillaBrowser1.Navigate2 ("http://media.radiosai.org/www/Bhajan.html")
StatusBar1.SimpleText = "[" & Time & "] [" & Date & "] Created By Naveen Venkat"
End Sub

Private Sub Timer1_Timer()
StatusBar1.SimpleText = "[" & Time & "] [" & Date & "] Created By Naveen Venkat"
End Sub
