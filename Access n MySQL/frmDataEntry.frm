VERSION 5.00
Begin VB.Form frmDataEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Entry"
   ClientHeight    =   3720
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
   ScaleHeight     =   3720
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMiddleName 
      Height          =   360
      Left            =   120
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox txtAddress 
      Height          =   1020
      Left            =   120
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox txtFirstName 
      Height          =   360
      Left            =   120
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox txtLastName 
      Height          =   360
      Left            =   120
      MaxLength       =   20
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   735
      Left            =   3480
      Picture         =   "frmDataEntry.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   735
      Left            =   3480
      Picture         =   "frmDataEntry.frx":0513
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Middle Name"
      Height          =   240
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1140
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Address"
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "First Name"
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Last Name"
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "frmDataEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
frmPerson.Person.Save txtLastName, txtFirstName, txtMiddleName, txtAddress
End Sub

