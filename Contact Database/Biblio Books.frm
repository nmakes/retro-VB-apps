VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Biblio Books"
   ClientHeight    =   4440
   ClientLeft      =   4725
   ClientTop       =   3255
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   5760
   Begin VB.TextBox Text8 
      DataField       =   "Comments"
      DataSource      =   "adobooks"
      Height          =   855
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      DataField       =   "Notes"
      DataSource      =   "adobooks"
      Height          =   1215
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      DataField       =   "Description"
      DataSource      =   "adobooks"
      Height          =   285
      Left            =   1320
      TabIndex        =   17
      Top             =   1680
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc adobooks 
      Height          =   330
      Left            =   6840
      Top             =   3840
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Program Files\Microsoft Visual Studio\VB98\BIBLIO.MDB;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Program Files\Microsoft Visual Studio\VB98\BIBLIO.MDB;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Titles"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton ext 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   3960
      Width           =   5775
   End
   Begin VB.CommandButton nex 
      Caption         =   "&>"
      Height          =   495
      Left            =   2880
      TabIndex        =   10
      Top             =   3000
      Width           =   2895
   End
   Begin VB.CommandButton pre 
      Caption         =   "&<"
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   3000
      Width           =   2895
   End
   Begin VB.CommandButton can 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4320
      TabIndex        =   8
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton del 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton add 
      Caption         =   "&Add"
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton sav 
      Caption         =   "&Save"
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      DataField       =   "Subject"
      DataSource      =   "adobooks"
      Height          =   285
      Left            =   3720
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      DataField       =   "PubID"
      DataSource      =   "adobooks"
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      DataField       =   "ISBN"
      DataSource      =   "adobooks"
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      DataField       =   "Year Published"
      DataSource      =   "adobooks"
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      DataField       =   "Title"
      DataSource      =   "adobooks"
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Biblio Books"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1500
      TabIndex        =   23
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label8 
      Caption         =   "Comments:"
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Notes:"
      Height          =   255
      Left            =   3000
      TabIndex        =   20
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Description:"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Subject:"
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Publisher's ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "ISBN:"
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Year Published:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add_Click()
On Error Resume Next
adobooks.Recordset.AddNew
End Sub


Private Sub can_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
End Sub



Private Sub del_Click()
On Error Resume Next
confirm = MsgBox("Are you sure?", vbYesNo, "Delete Confirmation")
If confirm = vbYes Then
adobooks.Recordset.Delete
Else
MsgBox "Record not deleted!", , "Message"
End If
End Sub

Private Sub ext_Click()
End
End Sub

Private Sub nex_Click()
On Error Resume Next
If Not adobooks.Recordset.EOF Then
adobooks.Recordset.MoveNext
If adobooks.Recordset.EOF Then
adobooks.Recordset.MovePrevious
End If
End If

End Sub

Private Sub pre_Click()
On Error Resume Next
If Not adobooks.Recordset.BOF Then
adobooks.Recordset.MovePrevious
If adobooks.Recordset.BOF Then
adobooks.Recordset.MoveNext
End If
End If
End Sub

Private Sub sav_Click()
On Error GoTo erh
adobooks.Recordset.Fields("Title") = Text1.Text
adobooks.Recordset.Fields("Year Published") = Text2.Text
adobooks.Recordset.Fields("ISBN") = Text3.Text
adobooks.Recordset.Fields("PubID") = Text4.Text
adobooks.Recordset.Fields("Subject") = Text5.Text
adobooks.Recordset.Fields("Description") = Text6.Text
adobooks.Recordset.Fields("Notes") = Text7.Text
adobooks.Recordset.Fields("Comments") = Text8.Text
adobooks.Recordset.Update
'-------
erh:
MsgBox "A field is empty"

End Sub
