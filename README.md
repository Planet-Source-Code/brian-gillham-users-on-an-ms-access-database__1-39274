<div align="center">

## Users on an MS Access Database


</div>

### Description

These Functions will allow you to determine WHO is logged on to the Acces Database as well as HOW MANY users are logged-in. As always the code is FREE, if you want support can consult me in my professional capacity. ENJOY
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brian Gillham](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brian-gillham.md)
**Level**          |Intermediate
**User Rating**    |3.5 (14 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brian-gillham-users-on-an-ms-access-database__1-39274/archive/master.zip)





### Source Code

```
Public Function EmptyRS(ByVal oRS) As Boolean
  On Error Resume Next
  'Checks for an EMPTY RecordSet
  EmptyRS = True
  If Not oRS Is Nothing Then
    EmptyRS = ((oRS.BOF = True) And (oRS.EOF = True))
  End If
End Function
Public Function GetDBUsers() As ADODB.Recordset
  ' NOTES: Fields as follows
  ' 0 - COMPUTER_NAME:  Workstation
  ' 1 - LOGIN_NAME:    Name used to Login to DB
  ' 2 - CONNECTED:    True if Lock in LDB File
  ' 3 - SUSPECTED_STATE: True if user has left database in a suspect state(else Null)
  On Error GoTo LocalError
  Const JET_SCHEMA_USERROSTER = "{947bb102-5d43-11d1-bdbf-00c04fb92675}"
  ' Return a Disconnected RecordSet
  If cnADO.State = adStateOpen Then
    Set GetDBUsers = cnADO.OpenSchema(adSchemaProviderSpecific, , JET_SCHEMA_USERROSTER)
    Set GetDBUsers.ActiveConnection = Nothing
  End If
LocalError:
End Function
Public Function GetDBUserCount() As Long
  On Error GoTo LocalError
  Dim lRS As ADODB.Recordset
  Set lRS = GetDBUsers
  If Not EmptyRS(lRS) Then
    GetDBUserCount = lRS.RecordCount
    lRS.Close
  End If
LocalError:
  Set lRS = Nothing
End Function
```

