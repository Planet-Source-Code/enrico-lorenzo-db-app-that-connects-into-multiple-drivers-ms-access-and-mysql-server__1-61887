VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAccess 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Microsoft Access"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   735
      Left            =   4200
      Picture         =   "frmAccess.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   3120
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   735
      Left            =   4200
      Picture         =   "frmAccess.frx":04B2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox txtFileName 
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Filename"
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "frmAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
frmConnection.Show
Unload Me
End Sub

Private Sub cmdBrowse_Click()
cdl.Filter = "Microsoft Access Database | *.mdb"
cdl.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
cdl.ShowOpen

txtFileName = cdl.FileName
End Sub

Private Sub cmdConnect_Click()
ConnectAccess txtFileName, txtPassword
End Sub
