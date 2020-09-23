VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Save(LastName As String, FirstName As String, MiddleName As String, Address As String)
On Error GoTo iError
If frmPerson.Insrt = True Then
    cn.Execute "INSERT INTO Person(LastName, FirstName, MiddleName, Address) VALUES('" & LastName & "', '" & FirstName & "', '" & MiddleName & "', '" & Address & "')"
    MsgBox "Record sucessfully added.", vbInformation
Else
    cn.Execute "UPDATE Person SET LastName='" & LastName & "', FirstName='" & FirstName & "', MiddleName='" & MiddleName & "', Address='" & Address & "' WHERE PersonNo=" & rs(0)
    MsgBox "Record sucessfully updated.", vbInformation
End If

Unload frmDataEntry
rs.Requery
getRecord
Exit Function

iError:
MsgBox Err.Description, vbCritical
End Function

Public Function Delete()
cn.Execute "DELETE FROM Person WHERE PersonNo=" & rs(0)
rs.Requery
getRecord
End Function

Public Function Search(PersonNo As Integer)
rs.Requery
rs.Find "PersonNo=" & PersonNo
getRecord
End Function

