VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contact Database"
   ClientHeight    =   3855
   ClientLeft      =   4935
   ClientTop       =   3555
   ClientWidth     =   6015
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6015
   Begin VB.Frame Frame1 
      Caption         =   "Quick Search"
      Height          =   3495
      Left            =   0
      TabIndex        =   27
      Top             =   360
      Width           =   6015
      Begin VB.ListBox List4 
         Height          =   2985
         Left            =   4200
         TabIndex        =   32
         Top             =   360
         Width           =   1695
      End
      Begin VB.ListBox List3 
         Height          =   2985
         Left            =   2760
         TabIndex        =   31
         Top             =   360
         Width           =   1455
      End
      Begin VB.ListBox List2 
         Height          =   2985
         Left            =   1440
         TabIndex        =   30
         Top             =   360
         Width           =   1335
      End
      Begin VB.ListBox List1 
         DataField       =   "First Name"
         DataSource      =   "adodc1"
         Height          =   2985
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "x"
         Height          =   255
         Left            =   5520
         TabIndex        =   28
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3960
      TabIndex        =   25
      Top             =   3240
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   330
      Left            =   0
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
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
      Connect         =   $"Form1.frx":030A
      OLEDBString     =   $"Form1.frx":0391
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Contacts"
      Caption         =   "<> Browse Contacts <>"
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
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3960
      TabIndex        =   24
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   375
      Left            =   3960
      TabIndex        =   23
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   3960
      TabIndex        =   22
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text11 
      DataField       =   "Relation"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   840
      TabIndex        =   21
      Top             =   3120
      Width           =   2895
   End
   Begin VB.TextBox Text10 
      DataField       =   "City"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   3600
      TabIndex        =   19
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox Text9 
      DataField       =   "State"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   3600
      TabIndex        =   17
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      DataField       =   "Country"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   3600
      TabIndex        =   15
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      DataField       =   "Company/Job"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   4080
      TabIndex        =   13
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      DataField       =   "Email ID"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   840
      TabIndex        =   11
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      DataField       =   "Address"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   840
      TabIndex        =   9
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      DataField       =   "Mobile"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   720
      TabIndex        =   7
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      DataField       =   "Phone"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   720
      TabIndex        =   5
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      DataField       =   "Last Name"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   "First Name"
      DataSource      =   "adodc1"
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 2011 Spotlite Productions"
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   3600
      Width           =   6015
   End
   Begin VB.Line Line8 
      X1              =   3840
      X2              =   0
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label11 
      Caption         =   "Relation:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "City:"
      Height          =   255
      Left            =   3000
      TabIndex        =   18
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "State:"
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Country:"
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   960
      Width           =   615
   End
   Begin VB.Line Line4 
      X1              =   2880
      X2              =   6000
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label7 
      Caption         =   "Company/Job:"
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   480
      Width           =   1095
   End
   Begin VB.Line Line7 
      X1              =   0
      X2              =   3840
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label6 
      Caption         =   "Email ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   735
   End
   Begin VB.Line Line6 
      X1              =   3840
      X2              =   3840
      Y1              =   3480
      Y2              =   2040
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   3840
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line3 
      X1              =   2880
      X2              =   2880
      Y1              =   2040
      Y2              =   240
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   2880
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6000
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label5 
      Caption         =   "Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Mobile:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Phone:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Last Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "First Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin VB.Menu mndata 
      Caption         =   "&Data"
      Begin VB.Menu mnadd 
         Caption         =   "&Add"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mndelete 
         Caption         =   "&Delete"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnsave 
         Caption         =   "&Save"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnview 
      Caption         =   "&View"
      Begin VB.Menu mnsearch 
         Caption         =   "&Search"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnabout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnexit 
      Caption         =   "&Exit"
      Begin VB.Menu mnays 
         Caption         =   "Are you Sure?"
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
Private Sub adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
End Sub



Private Sub Command1_Click()
mnadd_Click
End Sub

Private Sub Command2_Click()
mnsave_Click
End Sub

Private Sub Command3_Click()
mndelete_Click
End Sub



Private Sub Command5_Click()
mnyes_Click
End Sub

Private Sub Command6_Click()
Frame1.Visible = False
End Sub

Private Sub mnadd_Click()
On Error Resume Next
adodc1.Recordset.AddNew
End Sub

Private Sub mndelete_Click()
On Error Resume Next
Confirm = MsgBox("Are you sure to delete this record?", vbYesNo, "Delete Confirmation")
If Confirm = vbYes Then
adodc1.Recordset.Delete
MsgBox "Record Deleted", , "Message"
Else
MsgBox "Record Not Deleted!", , "Message"
End If
End Sub

Private Sub mnsave_Click()
On Error Resume Next
adodc1.Recordset.Fields("First Name") = Text1.Text
adodc1.Recordset.Fields("Last Name") = Text2.Text
adodc1.Recordset.Fields("Phone") = Text3.Text
adodc1.Recordset.Fields("Mobile") = Text4.Text
adodc1.Recordset.Fields("Address") = Text5.Text
adodc1.Recordset.Fields("Email ID") = Text6.Text
adodc1.Recordset.Fields("Company/Job") = Text7.Text
adodc1.Recordset.Fields("Country") = Text8.Text
adodc1.Recordset.Fields("State") = Text9.Text
adodc1.Recordset.Fields("City") = Text10.Text
adodc1.Recordset.Fields("Relation") = Text11.Text
adodc1.Recordset.Update
End Sub



Private Sub mnsearch_Click()
Frame1.Visible = True
End Sub

Private Sub mnyes_Click()
On Error Resume Next
End
End Sub
