VERSION 5.00
Object = "{1B773E42-2509-11CF-942F-008029004347}#3.6#0"; "sysmon.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form AdvancedSystemMonitor 
   BackColor       =   &H8000000D&
   Caption         =   "Advanced System Monitor"
   ClientHeight    =   8505
   ClientLeft      =   4845
   ClientTop       =   1905
   ClientWidth     =   12165
   Icon            =   "AdvancedSystemMonitor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   12165
   Begin SystemMonitorCtl.SystemMonitor sm 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12135
      _Version        =   393219
      _ExtentX        =   21405
      _ExtentY        =   14420
      DisplayType     =   1
      ReportValueType =   0
      MaximumScale    =   100
      MinimumScale    =   0
      ShowLegend      =   -1  'True
      ShowToolbar     =   -1  'True
      ShowScaleLabels =   -1  'True
      ShowHorizontalGrid=   0   'False
      ShowVerticalGrid=   0   'False
      ShowValueBar    =   -1  'True
      ManualUpdate    =   0   'False
      Highlight       =   0   'False
      ReadOnly        =   0   'False
      MonitorDuplicateInstances=   -1  'True
      UpdateInterval  =   1
      DisplayFilter   =   1
      BackColorCtl    =   -2147483633
      ForeColor       =   -1
      BackColor       =   -1
      GridColor       =   -2147483633
      TimeBarColor    =   255
      Appearance      =   1
      BorderStyle     =   1
      NextCounterColor=   0
      NextCounterWidth=   0
      NextCounterLineStyle=   0
      GraphTitle      =   ""
      YAxisLabel      =   ""
      DataSourceType  =   1
      SqlDsnName      =   ""
      SqlLogSetName   =   ""
      LogFileCount    =   0
      AmbientFont     =   -1  'True
      LegendColumnWidths=   $"AdvancedSystemMonitor.frx":030A
      LegendSortDirection=   1364
      LegendSortColumn=   0
      CounterCount    =   0
      MaximumSamples  =   100
      SampleCount     =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2760
      Top             =   4200
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8250
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnfile 
      Caption         =   "&File"
      Begin VB.Menu mnexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mntheme 
      Caption         =   "&Theme"
      Begin VB.Menu mnblue 
         Caption         =   "&Blue (Default)"
      End
      Begin VB.Menu mngreen 
         Caption         =   "&Green"
      End
      Begin VB.Menu mnwhite 
         Caption         =   "&White"
      End
      Begin VB.Menu mnblack 
         Caption         =   "&Black"
      End
      Begin VB.Menu mnyellow 
         Caption         =   "&Yellow"
      End
      Begin VB.Menu mngrey 
         Caption         =   "&Grey"
      End
   End
   Begin VB.Menu mnabout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "AdvancedSystemMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
On Error Resume Next
sm.Width = Me.Width - 120
sm.Height = Me.Height - 1072
End Sub

Private Sub mnabout_Click()
frmAbout.Show
End Sub

Private Sub mnblack_Click()
Me.BackColor = &H0&
End Sub

Private Sub mnblue_Click()
Me.BackColor = &H8000000D
End Sub

Private Sub mnexit_Click()
kkr = MsgBox("Are you sure?", vbYesNo, "Quitting?")
If kkr = vbYes Then
End
End If
End Sub

Private Sub mngreen_Click()
Me.BackColor = &HC000&
End Sub

Private Sub mngrey_Click()
Me.BackColor = &HC0C0C0
End Sub

Private Sub mnwhite_Click()
Me.BackColor = &HFFFFFF
End Sub

Private Sub mnyellow_Click()
Me.BackColor = &HFFFF&
End Sub



Private Sub Timer1_Timer()
sb.SimpleText = "[ " & Time & " ] " & "Copyright (c) 2011 Naveen Venkat | All Rights Reserved"
End Sub
