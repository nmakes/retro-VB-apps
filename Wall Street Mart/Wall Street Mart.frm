VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wall Street Mart"
   ClientHeight    =   4605
   ClientLeft      =   8115
   ClientTop       =   3165
   ClientWidth     =   6735
   Icon            =   "Wall Street Mart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Wall Street Mart.frx":030A
   ScaleHeight     =   4605
   ScaleWidth      =   6735
   Begin VB.Frame Expandstore 
      Caption         =   "[ Expand Store ]"
      Height          =   3735
      Left            =   2040
      TabIndex        =   23
      Top             =   600
      Width           =   4575
      Begin VB.CommandButton Command2 
         Caption         =   "Expand Store and Increase Capacity"
         Height          =   975
         Left            =   360
         TabIndex        =   29
         Top             =   2040
         Width           =   3855
      End
      Begin VB.TextBox capcst 
         Height          =   285
         Left            =   1680
         TabIndex        =   28
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox capinc 
         Height          =   285
         Left            =   1680
         TabIndex        =   26
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "x"
         Height          =   375
         Left            =   4080
         TabIndex        =   24
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Capcost 
         Alignment       =   1  'Right Justify
         Caption         =   "Cost:"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Capacity Increase:"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   1335
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   22
      Top             =   4335
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   476
      Style           =   1
      SimpleText      =   " .:|:. Wall Street Mart .:|:.       Copyright (c) 2011 Naveen Venkat        All Rights Reserved"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Stockframe 
      Caption         =   "[ Stock ]"
      Height          =   3735
      Left            =   2040
      TabIndex        =   11
      Top             =   600
      Width           =   4575
      Begin VB.CommandButton Buystockbutton 
         Caption         =   "Buy Stock"
         Height          =   1095
         Left            =   720
         TabIndex        =   19
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox finalgaintext 
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox costtext 
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox STOCKTEXT 
         Height          =   285
         Left            =   2160
         TabIndex        =   14
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton Stock_Close 
         Caption         =   "x"
         Height          =   375
         Left            =   4080
         TabIndex        =   12
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Final Gain:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Cost:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Enter the amount of stock:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Timer stockstimer 
      Interval        =   1
      Left            =   120
      Top             =   720
   End
   Begin VB.TextBox cap 
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "100"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox money 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "1000000"
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      Caption         =   "Game"
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
      Begin VB.CommandButton Quit 
         Caption         =   "Quit"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton Help 
         Caption         =   "Help"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton Load 
         Caption         =   "Load"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton Save 
         Appearance      =   0  'Flat
         Caption         =   "Save"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Stock"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1815
      Begin VB.CommandButton Buy 
         Caption         =   "Buy"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Store"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
      Begin VB.CommandButton Expand 
         Caption         =   "Expand"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.TextBox expirytime 
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Expiry time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3120
      TabIndex        =   20
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   3360
      X2              =   3360
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Capacity:"
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Money:"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   120
      Width           =   615
   End
   Begin VB.Image stock 
      Height          =   2550
      Left            =   3120
      Picture         =   "Wall Street Mart.frx":121B4
      Top             =   1080
      Width           =   2340
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      Height          =   3615
      Left            =   2040
      Top             =   720
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Buy_Click()
Stockframe.Visible = True
Expandstore.Visible = False
End Sub

Private Sub Buystockbutton_Click()
l = Val(money.Text)
o = Val(costtext.Text)
m = Val(cap.Text)
n = Val(STOCKTEXT.Text)


If l >= o And m >= n Then
lk = l - o
money.Text = Str$(lk)
stockstimer.Enabled = True
expirytime.Text = STOCKTEXT.Text
Else
MsgBox "Not enough funds or capacity"
Exit Sub
End If
End Sub



Private Sub capinc_Change()
o = Val(capinc.Text)
capcsts = o * 80
capcst.Text = Str$(capcsts)
End Sub

Private Sub Command1_Click()
Expandstore.Visible = False
End Sub

Private Sub Command2_Click()
p = Val(capcst.Text)
s = Val(money.Text)
a = Val(cap.Text)
g = Val(capinc.Text)
If p <= s Then
capp = a + g
cap.Text = Str$(capp)
hhjj = s - p
money.Text = Str$(hhjj)
End If
End Sub

Private Sub Expand_Click()
Expandstore.Visible = True
Stockframe.Visible = False
End Sub

Private Sub expirytime_Change()
j = Val(expirytime.Text)
If j <= 0 Then
stockstimer.Enabled = False
stock.Visible = False
End If
End Sub

Private Sub Form_Load()
Me.Caption = "Wall Street Mart v" & App.Major & "." & App.Minor & "." & App.Revision
stock.Visible = False
Stockframe.Visible = False
stockstimer.Enabled = False
Expandstore.Visible = False
End Sub

Private Sub Stock_Close_Click()
Stockframe.Visible = False
End Sub


Private Sub stockstimer_Timer()
If stock.Visible = False Then
stock.Visible = True
End If

'--------

a = Val(expirytime.Text)
tl = a - 1
expirytime.Text = Str$(tl)

'---------

k = Val(money.Text)
If expirytime.Text > -1 Then
kk = k + 8
money.Text = Str$(kk)
End If
End Sub

Private Sub STOCKTEXT_Change()
a = Val(STOCKTEXT.Text)
costtext.Text = a * 4
finalgaintext.Text = a * 8
End Sub


