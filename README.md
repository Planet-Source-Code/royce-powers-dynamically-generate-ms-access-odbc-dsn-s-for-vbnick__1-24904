<div align="center">

## Dynamically generate MS Access ODBC DSN's \(for VbNick\)


</div>

### Description

Class object that can be compiled or copied and pasted into your application that will dynamically create MS Access ODBC DSN's for you.
 
### More Info
 
ODBC Type, ODBC Name, MDB Path, Optional User ID, Optional Password

Copy and Paste into your app, or compile and call the object.

Success code-Integer; Creates ODBC DSN


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Royce Powers](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/royce-powers.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 6\.0, ASP \(Active Server Pages\) , VBA MS Access
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/royce-powers-dynamically-generate-ms-access-odbc-dsn-s-for-vbnick__1-24904/archive/master.zip)

### API Declarations

```
Declare Function SQLAllocConnect Lib "odbc32.dll" (ByVal henv _
 As Long, phdbc As Long) As Integer
Declare Function SQLDisconnect Lib "odbc32.dll" (ByVal hdbc As _
 Long) As Integer
Declare Function SQLConnect Lib "odbc32.dll" (ByVal hdbc As _
 Long, ByVal szDSN As String, ByVal cbDSN As Integer, ByVal szUID As _
 String, ByVal cbUID As Integer, ByVal szAuthStr As String, ByVal _
 cbAuthStr As Integer) As Integer
Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal henv As _
 Long) As Integer
Declare Function SQLFreeConnect Lib "odbc32.dll" (ByVal hdbc _
 As Long) As Integer
Declare Function SQLError Lib "odbc32.dll" (ByVal henv As _
 Long, ByVal hdbc As Long, ByVal hstmt As Long, ByVal szSqlState As _
 String, pfNativeError As Long, ByVal szErrorMsg As String, ByVal _
 cbErrorMsgMax As Integer, pcbErrorMsg As Integer) As Integer
Declare Function SQLConfigDataSource Lib "ODBCCP32" _
 (ByVal hwndParent As Long, ByVal fRequest As Long, _
 ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
```


### Source Code

```
' in Module (.bas)
Option Explicit
Public Const vbAPINull As Long = 0&
Private Const SQL_SUCCESS As Long = 0
Private Const SQL_SUCCESS_WITH_INFO As Long = 1
Declare Function SQLAllocConnect Lib "odbc32.dll" (ByVal henv _
 As Long, phdbc As Long) As Integer
Declare Function SQLDisconnect Lib "odbc32.dll" (ByVal hdbc As _
 Long) As Integer
Declare Function SQLConnect Lib "odbc32.dll" (ByVal hdbc As _
 Long, ByVal szDSN As String, ByVal cbDSN As Integer, ByVal szUID As _
 String, ByVal cbUID As Integer, ByVal szAuthStr As String, ByVal _
 cbAuthStr As Integer) As Integer
Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal henv As _
 Long) As Integer
Declare Function SQLFreeConnect Lib "odbc32.dll" (ByVal hdbc _
 As Long) As Integer
Declare Function SQLError Lib "odbc32.dll" (ByVal henv As _
 Long, ByVal hdbc As Long, ByVal hstmt As Long, ByVal szSqlState As _
 String, pfNativeError As Long, ByVal szErrorMsg As String, ByVal _
 cbErrorMsgMax As Integer, pcbErrorMsg As Integer) As Integer
Declare Function SQLConfigDataSource Lib "ODBCCP32" _
 (ByVal hwndParent As Long, ByVal fRequest As Long, _
 ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
' In Class (.cls)
Option Explicit
Public Enum peDSN_OPTIONS
 ODBC_ADD_DSN = 1
 ODBC_CONFIG_DSN = 2
 ODBC_ADD_SYS_DSN = 4
 ODBC_CONFIG_SYS_DSN = 5
End Enum
Public Function RegisterDataSource(iFunction As peDSN_OPTIONS, sDSNName As String, sMDBPath As String, _
         Optional sUserID As String, Optional sPassword As String) As Integer
 Dim sAttributes As String
 Dim iRetVal As Integer
 If sUserID = "" Then sUserID = "Admin"
 sAttributes = "DSN=" & sDSNName _
 & Chr$(0) & "Description=Microsoft Access Database (" & sMDBPath & ")" _
 & Chr$(0) & "UID = " & sUserID _
 & Chr$(0) & "DefaultDir=" & sMDBPath _
 & Chr$(0) & "DBQ=" & sMDBPath _
 & Chr$(0)
 iRetVal = SQLConfigDataSource(vbAPINull, iFunction, "Microsoft Access Driver (*.mdb)", sAttributes)
End Function
```

