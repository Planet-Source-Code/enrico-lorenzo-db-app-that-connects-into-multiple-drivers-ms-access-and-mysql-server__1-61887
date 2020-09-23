VERSION 5.00
Begin VB.Form frmMySQL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MySQL ODBC 3.51"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
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
   ScaleHeight     =   3195
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   735
      Left            =   3480
      Picture         =   "frmMySQL.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   735
      Left            =   3480
      Picture         =   "frmMySQL.frx":04B2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox txtUser 
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox txtServer 
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox txtDatabase 
      Height          =   360
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "User"
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Server"
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Database"
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   825
   End
End
Attribute VB_Name = "frmMySQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
frmConnection.Show
Unload Me
End Sub

Private Sub cmdConnect_Click()
ConnectMySQL txtServer, txtUser, txtPassword, txtDatabase
End Sub

