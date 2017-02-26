VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7455
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      Height          =   4590
      Left            =   37
      TabIndex        =   0
      Top             =   7
      Width           =   7380
      Begin VB.CommandButton Command1 
         Caption         =   "Validate"
         Height          =   315
         Left            =   5280
         TabIndex        =   11
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         Top             =   4200
         Width           =   4095
      End
      Begin VB.PictureBox picLogo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   1320
         Picture         =   "frmSplash.frx":030A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   2
         Top             =   1800
         Width           =   510
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "Licenced By Naveen Venkat"
         Height          =   255
         Left            =   270
         TabIndex        =   1
         Tag             =   "LicenseTo"
         Top             =   300
         Width           =   6855
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Writer"
         BeginProperty Font 
            Name            =   "Edwardian Script ITC"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   1680
         TabIndex        =   9
         Tag             =   "Product"
         Top             =   1200
         Width           =   3315
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Naveen Makes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2505
         TabIndex        =   8
         Tag             =   "CompanyProduct"
         Top             =   765
         Width           =   2595
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6720
         TabIndex        =   7
         Tag             =   "Platform"
         Top             =   2400
         Width           =   420
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version 4.2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5775
         TabIndex        =   6
         Tag             =   "Version"
         Top             =   2760
         Width           =   1380
      End
      Begin VB.Label lblWarning 
         Caption         =   $"frmSplash.frx":0614
         Height          =   795
         Left            =   360
         TabIndex        =   3
         Tag             =   "Warning"
         Top             =   3600
         Width           =   6855
      End
      Begin VB.Label lblCompany 
         Alignment       =   1  'Right Justify
         Caption         =   "Naveen Makes"
         Height          =   255
         Left            =   4710
         TabIndex        =   5
         Tag             =   "Company"
         Top             =   3330
         Width           =   2415
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         Caption         =   "Copyright (c) 2011 Naveen Venkat"
         Height          =   255
         Left            =   4710
         TabIndex        =   4
         Tag             =   "Copyright"
         Top             =   3120
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "8a1b6c3d" Then
MsgBox "[C][O][R][R][E][C][T] Thankyou for buying an ORIGINAL SOFTWARE!"
Else
MsgBox "[W][R][O][N][G] The Code is INVALID!"
End If
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

