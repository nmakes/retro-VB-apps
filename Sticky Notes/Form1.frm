VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sticky Notes"
   ClientHeight    =   2055
   ClientLeft      =   14970
   ClientTop       =   10395
   ClientWidth     =   2295
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   2295
   Begin VB.TextBox Text1 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Kartika"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   2055
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":030A
      Top             =   0
      Width           =   2295
   End
   Begin VB.Menu mnclear 
      Caption         =   "&Clear"
   End
   Begin VB.Menu mnabout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnabout_Click()
MsgBox "Coded by Naveen Venkat | Copyright (c) 2011 Naveen Venkat | All Rights Reserved | www.nmakes.webs.com"
End Sub

Private Sub mnclear_Click()
Text1.Text = ""
End Sub
