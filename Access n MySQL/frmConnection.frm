VERSION 5.00
Begin VB.Form frmConnection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Driver Connection"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
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
   ScaleHeight     =   2625
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   735
      Left            =   120
      Picture         =   "frmConnection.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ListBox lstDriver 
      Height          =   1020
      ItemData        =   "frmConnection.frx":04B1
      Left            =   120
      List            =   "frmConnection.frx":04BB
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Driver"
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNext_Click()
If lstDriver.ListIndex = 0 Then
    frmAccess.Show
Else
    frmMySQL.Show
End If
Unload Me
End Sub

Private Sub Form_Activate()
lstDriver.Selected(0) = True
End Sub

