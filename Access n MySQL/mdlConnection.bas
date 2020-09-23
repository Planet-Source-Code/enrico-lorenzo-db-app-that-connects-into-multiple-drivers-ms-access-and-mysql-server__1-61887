Attribute VB_Name = "mdlConnection"
Option Explicit

Public cn As New ADODB.Connection
Public rs As New ADODB.Recordset

Public Function ConnectMySQL(Server As String, User As String, Password As String, Database As String)
On Error GoTo econnectmysql

cn.CursorLocation = adUseClient
cn.Open "DRIVER={MySQL ODBC 3.51 Driver};Server=" & Server & ";UID=" & User & ";PWD=" & Password & ";Database=" & Database

Record
Unload frmMySQL
Exit Function

econnectmysql:
MsgBox Err.Description, vbCritical
End Function

Public Function ConnectAccess(FileName As String, Password As String)
On Error GoTo econnectaccess

cn.CursorLocation = adUseClient
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileName & ";Jet OLEDB:Database Password=" & Password

Record
Unload frmAccess
Exit Function

econnectaccess:
MsgBox Err.Description, vbCritical
End Function

Public Function Record()
rs.Open "SELECT * FROM Person", cn, adOpenDynamic, adLockOptimistic
rs.Sort = "PersonNo"

frmPerson.Show
End Function

Public Function getRecord()
With frmPerson
    On Error GoTo egetrecord
    .lblPersonNo = "Person No.: " & rs(0)
    .lblLastName = "Last Name: " & rs(1)
    .lblFirstName = "First Name: " & rs(2)
    .lblMiddleName = "Middle Name: " & rs(3)
    .lblAddress = "Address: " & rs(4)
    Exit Function

egetrecord:
    .lblPersonNo = "Person No.: "
    .lblLastName = "Last Name: "
    .lblFirstName = "First Name: "
    .lblMiddleName = "Middle Name: "
    .lblAddress = "Address: "
End With
End Function
