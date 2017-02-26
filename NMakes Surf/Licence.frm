VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Licence - NMakes Surf"
   ClientHeight    =   5850
   ClientLeft      =   5475
   ClientTop       =   2190
   ClientWidth     =   8745
   Icon            =   "Licence.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5850
   ScaleWidth      =   8745
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   5895
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Licence.frx":030A
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Caption = "Licence - NMakes Surf " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Resize()
Text1.Height = Form2.Height - 250
Text1.Width = Form2.Width - 70
End Sub

