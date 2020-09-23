VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPerson 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Person Records"
   ClientHeight    =   6165
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8595
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
   ScaleHeight     =   6165
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   2990
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Students"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "PersonNo"
         Caption         =   "Person No."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "LastName"
         Caption         =   "Last Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "FirstName"
         Caption         =   "First Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "MiddleName"
         Caption         =   "Middle Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Address"
         Caption         =   "Address"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1995.024
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   735
      Left            =   2520
      Picture         =   "frmPerson.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   735
      Left            =   1320
      Picture         =   "frmPerson.frx":0555
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Height          =   735
      Left            =   120
      Picture         =   "frmPerson.frx":0A6F
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox txtPersonNo 
      BackColor       =   &H00C0FFFF&
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   735
      Left            =   2880
      Picture         =   "frmPerson.frx":0F1C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   7815
      Begin VB.Label lblPersonNo 
         AutoSize        =   -1  'True
         Caption         =   "Person No:"
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblLastName 
         AutoSize        =   -1  'True
         Caption         =   "Last Name:"
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label lblFirstName 
         AutoSize        =   -1  'True
         Caption         =   "First Name:"
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1020
      End
      Begin VB.Label lblMiddleName 
         AutoSize        =   -1  'True
         Caption         =   "Middle Name:"
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         Caption         =   "Address:"
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   780
      End
      Begin VB.Image Image1 
         Height          =   1440
         Left            =   6120
         Picture         =   "frmPerson.frx":1405
         Top             =   120
         Width           =   1440
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Person No."
      Height          =   240
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Person As New clsPerson
Public Insrt As Boolean

Private Sub cmdDelete_Click()
If rs.EOF Then
    MsgBox "No record to be deleted.", vbInformation
Else
    If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo) = vbYes Then
        Person.Delete
    End If
End If
End Sub

Private Sub cmdInsert_Click()
Insrt = True
frmDataEntry.Show vbModal
End Sub

Private Sub cmdSearch_Click()
On Error GoTo esearch
Person.Search txtPersonNo

If rs.EOF Then
    MsgBox "No record found.", vbInformation
End If
Exit Sub

esearch:
MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdUpdate_Click()
Insrt = False
If rs.EOF Then
    MsgBox "No record to be updated.", vbInformation
Else
    With frmDataEntry
        .txtLastName = rs(1)
        .txtFirstName = rs(2)
        .txtMiddleName = rs(3)
        .txtAddress = rs(4)
        .Show vbModal
    End With
End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
getRecord
End Sub

Private Sub Form_Load()
Set DataGrid1.DataSource = rs.DataSource
getRecord
End Sub
