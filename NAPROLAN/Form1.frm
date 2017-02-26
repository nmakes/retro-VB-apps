VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NaProLan"
   ClientHeight    =   5085
   ClientLeft      =   6690
   ClientTop       =   2850
   ClientWidth     =   7470
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7470
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   8160
      TabIndex        =   2
      Top             =   360
      Width           =   150
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   4800
      Width           =   7455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   4815
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Text1.Text = Text1.Text & vbCrLf & "-----" & vbCrLf & "NAPROLAN - NAveen's PROgramming LANguage" & vbCrLf & "Copyright (c) 2011 Naveen Venkat" & vbCrLf & "All Rights Reserved" & vbCrLf & "www.nmakes.webs.com"
End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn And Text2.Text <> "" Then
Text1.Text = Text1.Text & vbCrLf & Text2.Text
GoTo CodeHandler
Else
Exit Sub
End If

'[][][][][][][][][][][][][][][][][]'
'----------------------------------'
'[C][O][D][E] [H][A][N][D][L][E][R]'
'----------------------------------'
'[][][][][][][][][][][][][][][][][]'

CodeHandler:

Select Case Text2.Text

'=================='
'== Clear Screen =='
'=================='

    Case "cls"
    Text1.Text = ""

'============='
'== Library =='
'============='
    
    Case "library"
    Form2.Show
    Text1.Text = Text1.Text & vbCrLf & "Library Loaded Successfully!"
    
'==========='
'== Input =='
'==========='
    
    Case "input"
    Text1.Text = Text1.Text & vbCrLf & "No input device loaded!"
    
    
    
    
    
'==========='
'== About =='
'==========='

    Case "about"
    Text1.Text = Text1.Text & vbCrLf & "-----" & vbCrLf & "NAPROLAN - NAveen's PROgramming LANguage" & vbCrLf & "Copyright (c) 2011 Naveen Venkat" & vbCrLf & "All Rights Reserved" & vbCrLf & "www.nmakes.webs.com"


'=========='
'== Exit =='
'=========='

    Case "exit"
    End





'=============================================='
'== If none of the above case is there then: =='
'=============================================='
    
    Case Else
    Text1.Text = Text1.Text & vbCrLf & "Undefined Function!"
    
End Select
End Sub
