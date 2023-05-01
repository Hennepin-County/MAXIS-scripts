'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "UTILITIES - TEST STAFF INFO.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 1                      'manual run time in seconds
STATS_denomination = "C"       			   'C is for each CASE
'END OF stats block==============================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
        IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
            FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
        Else											'Everyone else should use the release branch.
            FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
        End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                           "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                           "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                           vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
           StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

current_user_name = ""
current_user_email = ""
Current_user_x_number = ""

'Creating objects for Access
Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

' Connection_String = "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
SQL_table = "SELECT * from ES.ES_StaffHierarchyDim"

'This is the file path for the statistics Access database.
' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
objRecordSet.Open SQL_table, objConnection

' windows_user_ID

Do While NOT objRecordSet.Eof
	table_user_id = objRecordSet("EmpLogOnID")

	If table_user_id = windows_user_ID Then

		current_user_name = objRecordSet("EmpFullName")
		current_user_email = objRecordSet("EmployeeEmail")
		Current_user_x_number = objRecordSet("EmpStateLogOnID")

		Exit Do
	End If

	objRecordSet.MoveNext
Loop

objRecordSet.Close
objConnection.Close
Set objRecordSet=nothing
Set objConnection=nothing

If current_user_name <> "" Then
	Name_array = split(current_user_name, ",")
	current_user_name = trim(Name_array(1)) & " " & trim(Name_array(0))

	show_my_info = "Your Information has been found in script data."
	show_my_info = show_my_info & vbCr & vbCR & "Name: " & current_user_name
	show_my_info = show_my_info & vbCr & "E-Mail: " & current_user_email
	show_my_info = show_my_info & vbCr & "X Number: " & Current_user_x_number
	end_msg = "Success! The script has been able to identify you!" & vbCr & vbCr & show_my_info & vbcr & vbCr & "The BlueZone Script Team has been notified that it worked. Thank you!"

Else
	show_my_info = "Your information could not be found."
	end_msg = "The test did not fail, but it didn't exactly work.!" & vbCr & vbCr & show_my_info & vbcr & vbCr & "The BlueZone Script Team has been notified. Thank you!"

End If

email_body = show_my_info & vbCr & vbCr & "Windows ID: " & windows_user_ID
Call create_outlook_email("hsph.ews.bluezonescripts@hennepin.us", "", "StaffHierarchy User Table Test", email_body, "", True)

script_end_procedure(end_msg)


' "HsStaffHierarchyDimKey"
' "EmpFullName"
' "EmployeeNumber"
' "OrgInfoServiceArea"
' "ServiceArea"
' "L1ManagerEmployeeNumber"
' "L1Manager"
' "L2ManagerEmployeeNumber"
' "L2Manager"
' "L3ManagerEmployeeNumber"
' "L3Manager"
' "L4ManagerEmployeeNumber"
' "L4Manager"
' "AuditLoadDate"
' "AuditLoadBy"
' "AuditChangeDate"
' "AuditChangeBy"
' "EmployeeEmail"
' "EmpStateLogOnID"
' objRecordSet("EmpLogOnID")


function find_user_name(the_person_running_the_script)
'--- This function finds the outlook name of the person using the script
'~~~~~ the_person_running_the_script:the variable for the person's name to output
'===== Keywords: MAXIS, worker name, email signature
	On Error Resume Next
	Set objOutlook = CreateObject("Outlook.Application")
	' MsgBox IsEmpty(objOutlook)
	' MsgBox "error number - " & Err.Number
	' If Err.Number = 0 Then
	Set the_person_running_the_script = objOutlook.GetNamespace("MAPI").CurrentUser
	' Set the_person_running_the_script = objOutlook.GetNamespace("MAPI").CurrentProfileName
	the_person_running_the_script = the_person_running_the_script & ""
	If Err.Number <> 0 Then MsgBox "NO NAME FOUND"
	MsgBox "the_person_running_the_script - " & the_person_running_the_script
	Set objOutlook = Nothing
	On Error Goto 0
end function