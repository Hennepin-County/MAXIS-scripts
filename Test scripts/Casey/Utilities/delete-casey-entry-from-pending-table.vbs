Const adOpenStatic = 3
Const adLockOptimistic = 3

'declare the SQL statement that will query the database

'Creating objects for Access
Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

'This is the file path for the statistics Access database.
' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""

' objSQL = "SELECT * FROM ES.ES_CasesPending"
' objRecordSet.Open objSQL, objConnection
' Do While NOT objRecordSet.Eof
'     MsgBox objRecordSet("AuditLoadBy")
'     ' If objRecordSet("CaseNumber") = eight_digit_case_number Then
'     '     current_case_record_found = True
'     ' End If
'     objRecordSet.MoveNext
' Loop
my_name = "HC_ACCT\CALO001"
objRecordSet.Open "DELETE FROM ES.ES_CasesPending WHERE AuditLoadBy = '" & my_name & "'", objConnection

' objRecordSet.Close
objConnection.Close
Set objRecordSet=nothing
Set objConnection=nothing
