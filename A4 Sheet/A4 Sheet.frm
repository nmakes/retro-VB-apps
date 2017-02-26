VERSION 5.00
Begin VB.Form a4 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A4 Sheet"
   ClientHeight    =   9120
   ClientLeft      =   4125
   ClientTop       =   1395
   ClientWidth     =   8295
   Icon            =   "A4 Sheet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "A4 Sheet.frx":030A
   MousePointer    =   99  'Custom
   ScaleHeight     =   9120
   ScaleWidth      =   8295
   Begin VB.CommandButton Command5 
      Caption         =   "ERASE SHEET"
      BeginProperty Font 
         Name            =   "Kartika"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      MouseIcon       =   "A4 Sheet.frx":0BD4
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Brush"
      BeginProperty Font 
         Name            =   "Kartika"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      MouseIcon       =   "A4 Sheet.frx":149E
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   8160
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pen"
      BeginProperty Font 
         Name            =   "Kartika"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      MouseIcon       =   "A4 Sheet.frx":1D68
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   8160
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pencil"
      BeginProperty Font 
         Name            =   "Kartika"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      MouseIcon       =   "A4 Sheet.frx":2632
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   8160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Resize the brush"
      BeginProperty Font 
         Name            =   "Kartika"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      MouseIcon       =   "A4 Sheet.frx":2EFC
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   8160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Kartika"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      MouseIcon       =   "A4 Sheet.frx":37C6
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Text            =   "30"
      Top             =   8280
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000E&
      Height          =   8055
      Left            =   0
      MouseIcon       =   "A4 Sheet.frx":4090
      MousePointer    =   99  'Custom
      ScaleHeight     =   7995
      ScaleWidth      =   8235
      TabIndex        =   0
      Top             =   0
      Width           =   8295
   End
   Begin VB.Label Label2 
      Caption         =   "Copyright (c) 2010 Naveen Venkat | Spotlite Productions | All Rights Reserved"
      BeginProperty Font 
         Name            =   "Kartika"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      MouseIcon       =   "A4 Sheet.frx":495A
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   8760
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "Size:"
      BeginProperty Font 
         Name            =   "Kartika"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      MouseIcon       =   "A4 Sheet.frx":5224
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   8280
      Width           =   375
   End
End
Attribute VB_Name = "a4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim drawing As Boolean

Private Sub Command1_Click()
On Error Resume Next
If Text1.Text = "" Or Text1.Text <= 0 Then
MsgBox "Please Enter a correct brush size (Minimum 1)"
Text1.Text = "1"
Picture1.DrawWidth = Text1.Text
Else
Picture1.DrawWidth = Text1.Text
End If
End Sub



Private Sub Command2_Click()
Text1.Text = "1.8"
Picture1.DrawWidth = Text1.Text
End Sub

Private Sub Command3_Click()
Text1.Text = "3"
Picture1.DrawWidth = Text1.Text
End Sub

Private Sub Command4_Click()
Text1.Text = "8"
Picture1.DrawWidth = Text1.Text
End Sub

Private Sub Command5_Click()
Picture1.Picture = LoadPicture("")
End Sub

Private Sub Form_Load()
drawing = False
End Sub



Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
drawing = True
If Text1.Text = "" Or Text1.Text <= 0 Then
MsgBox "Enter a correct size of the brush"
Text1.Text = "1"
Picture1.DrawWidth = Text1.Text
Else
drawing = True
Picture1.DrawWidth = Text1.Text
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ha
If drawing = True Then
    Picture1.PSet (X, Y)
End If
Picture1.DrawWidth = Text1.Text
Exit Sub
ha:
MsgBox "Enter a correct size of the brush (minimum 1)"
Text1.Text = "1"
Picture1.DrawWidth = Text1.Text
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
drawing = False
End Sub

