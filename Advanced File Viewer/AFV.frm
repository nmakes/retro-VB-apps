VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advanced File Viewer"
   ClientHeight    =   8415
   ClientLeft      =   5070
   ClientTop       =   1785
   ClientWidth     =   6735
   Icon            =   "AFV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "AFV.frx":030A
   MousePointer    =   99  'Custom
   ScaleHeight     =   8415
   ScaleWidth      =   6735
   Begin VB.CommandButton Command3 
      Caption         =   "CONTACT NAVEEN"
      Height          =   615
      Left            =   120
      MouseIcon       =   "AFV.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   7680
      Width           =   6375
   End
   Begin VB.Frame Frame1 
      Caption         =   "File Tasks"
      Height          =   1215
      Left            =   1380
      MouseIcon       =   "AFV.frx":091E
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   3600
      Width           =   4095
      Begin VB.CommandButton Command2 
         Caption         =   "DELETE"
         Height          =   495
         Left            =   2160
         MouseIcon       =   "AFV.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Open"
         Height          =   495
         Left            =   240
         MouseIcon       =   "AFV.frx":0F32
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.Line Line1 
         X1              =   2040
         X2              =   2040
         Y1              =   240
         Y2              =   1080
      End
      Begin VB.Label Label3 
         Caption         =   "Delete the selected file"
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Open the selected file"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.TextBox textBox 
      Height          =   285
      Left            =   960
      MouseIcon       =   "AFV.frx":123C
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "The current selected file"
      Top             =   3240
      Width           =   5655
   End
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   3360
      MouseIcon       =   "AFV.frx":1546
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   "Select the ""FILE"" to perform action"
      Top             =   360
      Width           =   3255
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   0
      MouseIcon       =   "AFV.frx":1850
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Select the ""Folder"""
      Top             =   360
      Width           =   3255
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      MouseIcon       =   "AFV.frx":1B5A
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Select the ""Drive"""
      Top             =   0
      Width           =   6615
   End
   Begin VB.Label Label7 
      Caption         =   "All Rights Reserved"
      Height          =   255
      Left            =   2640
      TabIndex        =   13
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "(Founder, CEO) Spotlite Production"
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   $"AFV.frx":1E64
      Height          =   1935
      Left            =   120
      MouseIcon       =   "AFV.frx":2110
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   5640
      Width           =   6375
   End
   Begin VB.Label Label4 
      Caption         =   "Copyright (c) 2010 Naveen Venkat"
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Current File:"
      Height          =   255
      Left            =   0
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3240
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo ErrorHandler
If textBox.Text = "" Then
MsgBox "Please select a file to open!"
Else
Shell textBox.Text, vbMaximizedFocus
Drive1.Refresh
Dir1.Refresh
File1.Refresh
End If
Exit Sub
ErrorHandler:
MsgBox "Cannot Open File.CHECK THE PATH OF THE FILE IN THE ABOVE TEXTBOX. IF IT IS OK THEN I CANNOT/WILL NOT OPEN THIS KIND OF FILE AND/OR NEED ANOTHER SOFTWARE."
End Sub

Private Sub Command2_Click()
On Error GoTo ERHL
If textBox.Text = "" Then
MsgBox "Please select a file to delete!"
Else
Kill textBox.Text
MsgBox "Deleted!"
Drive1.Refresh
Dir1.Refresh
File1.Refresh
End If
Exit Sub
ERHL:
MsgBox "Cannot Delete File. MAKE SURE THAT THE PATH OF THE FILE IN THE ABOVE TEXT BOX IS CORRECT. IF IT IS CORRECT THEN WINDOWS CANNOT/WILL NOT DELETE THIS KIND OF FILE"
End Sub

Private Sub Command3_Click()
contact.Show
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1
End Sub

Private Sub File1_Click()
textBox.Text = Dir1 & "\" & File1
End Sub

