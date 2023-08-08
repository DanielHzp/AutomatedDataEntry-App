Attribute VB_Name = "Module1"
Option Explicit

Sub UpdateDataBaseADO()

Dim conn As New ADODB.Connection
Dim cmd As ADODB.Command
Dim StrCon As String
Dim SqlString
Dim UserId
Dim Password As String
Dim NewRecordAdded As Range

'Search new records added by the macro executed in the userform
'The search should be done per attribute value to add in the DB
NewRecordAdded = Range("C99999").End(xlUp).Row + 1


'Open new ADO SQL connection to insert new data obtained in the internal application
Set dbADOConnection = New ADODB.Connection


StrCon = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Initial Catalog=Reporting;Data Source=dpsql01"


dbADOConnection.ConnectionString = StrCon

'Set connection credentials

UserId = "Insert id"
Password = "Insert Password"


'Activate connection setup
dbADOConnection.Open ConnectionString, UserId, Password, OpenOptions
'conn.Open

'Set data insertion into a new DB record
'The target data base is a sample replica of the solution

SqlString = "INSERT INTO ZoqueDataBaseName Values(Getdate(), ActiveSheet.Range("C" & NewRecordAdded).value, Application.UserName)"

Open SqlString, dbADOConnection


dbADOConnection.Close

End Sub
