'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - DATA ACCESS TEST.vbs"
start_time = timer
STATS_counter = 1			     'sets the stats counter at one
STATS_manualtime = 	100			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
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


'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone


Call MAXIS_case_number_finder(MAXIS_case_number)
Call find_user_name(assigned_worker)

' MAXIS_case_number = "124312"
'Initial Dialog - Case number
Dialog1 = ""                                        'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 190, 85, "Data Test"
  EditBox 60, 25, 45, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 80, 65, 50, 15
    CancelButton 135, 65, 50, 15
  Text 5, 10, 185, 10, "Enter a CASE Number for the test."
  Text 5, 30, 50, 10, "Case Number:"
  Text 5, 45, 185, 20, "This script is just to test access to some of the data structures we have. "
EndDialog

Do
	err_msg = ""
	Dialog Dialog1
	cancel_without_confirmation
	Call validate_MAXIS_case_number(err_msg, "*")

	If err_msg <> "" Then MsgBox "Resolve to continue:" & vbCr & err_msg
Loop until err_msg = ""

'============== DBO USEAGE LOG TEST ==================================

stop_time = timer				'TODO - delete when the new data recording function is in place
script_run_end_time = time		'TODO - delete when the new data recording function is in place
script_run_end_date = date		'TODO - delete when the new data recording function is in place
script_run_time = stop_time - start_time
trial_name_of_script = "TEST - Database Access Test.vbs"
closing_message = "Test complete"


'Setting constants
Const adOpenStatic = 3
Const adLockOptimistic = 3

'Defaulting script success to successful
SCRIPT_success = -1

'Creating objects for Access
Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

'Fixing a bug when the script_end_procedure has an apostrophe (this interferes with Access)
closing_message = replace(closing_message, "'", "")

'Opening DB
objConnection.Open db_full_string

objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX, STATS_COUNTER, STATS_MANUALTIME, STATS_DENOMINATION, WORKER_COUNTY_CODE, SCRIPT_SUCCESS, CASE_NUMBER)" &  _
"VALUES ('" & user_ID & "', '" & script_run_end_date & "', '" & script_run_end_time & "', '" & trial_name_of_script & "', " & abs(script_run_time) & ", '" & closing_message & "', " & abs(STATS_counter) & ", " & abs(STATS_manualtime) & ", '" & STATS_denomination & "', '" & worker_county_code & "', " & SCRIPT_success & ", '" & MAXIS_CASE_NUMBER & "')", objConnection, adOpenStatic, adLockOptimistic

'Closing the connection
objConnection.Close

email_subject = "Usage Log Database Access Test Completed for " & assigned_worker
email_body = "The Usage Log test was successful."

Call create_outlook_email("", "hsph.ews.bluezonescripts@hennepin.us", "", "", email_subject, 1, False, "", "", False, "", email_body, False, "", True)
MsgBox "Test 1 Successful"

'==== PENDING CASES TABLE =========================================

' MAXIS_case_number = trim(MAXIS_case_number)
' eight_digit_case_number = right("00000000"&MAXIS_case_number, 8)            'The SQL table functionality needs the leading 0s added to the Case Number

' cash_stat_code = "P"                    'determining the program codes for the table entry
' hc_stat_code = "I"
' ga_stat_code = "P"
' grh_stat_code = "P"
' emer_stat_code = "I"
' mfip_stat_code = "I"
' snap_stat_code = "P"

' If no_transfer_checkbox = checked Then worker_id_for_data_table = initial_pw_for_data_table     'determining the X-Number for table entry
' If no_transfer_checkbox = unchecked Then worker_id_for_data_table = transfer_to_worker
' If len(worker_id_for_data_table) = 3 Then worker_id_for_data_table = "X127" & worker_id_for_data_table

' 'Creating objects for Access
' Set objPENDConnection = CreateObject("ADODB.Connection")
' Set objPENDRecordSet = CreateObject("ADODB.Recordset")

' 'This is the BZST connection to SQL Database'
' objPENDConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlsw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""

' 'delete a record if the case number matches
' ' objRecordSet.Open "DELETE FROM ES.ES_CasesPending WHERE CaseNumber = '" & eight_digit_case_number & "'", objPENDConnection
' 'Add a new record with this case information'
' objPENDRecordSet.Open 	"INSERT INTO ES.ES_CasesPending (WorkerID, CaseNumber, CaseName, ApplDate, FSStatusCode, CashStatusCode, HCStatusCode, GAStatusCode, GRStatusCode, EAStatusCode, MFStatusCode, IsExpSnap, UpdateDate)" &  _
'                   		"VALUES ('" & worker_id_for_data_table & "', '" & eight_digit_case_number & "', '" & case_name_for_data_table & "', '" & application_date & "', '" & snap_stat_code & "', '" & cash_stat_code & "', '" & hc_stat_code & "', '" & ga_stat_code & "', '" & grh_stat_code & "', '" & emer_stat_code & "', '" & mfip_stat_code & "', '" & 1 & "', '" & date & "')", objPENDConnection, adOpenStatic, adLockOptimistic

' 'close the connection and recordset objects to free up resources
' objPENDConnection.Close
' Set objPENDRecordSet=nothing
' Set objPENDConnection=nothing

' email_subject = "Cases Pending Database Access Test Completed for " & assigned_worker
' email_body = "The Cases Pending test was successful."

' Call create_outlook_email("", "hsph.ews.bluezonescripts@hennepin.us", "", "", email_subject, 1, False, "", "", False, "", email_body, False, "", True)
' MsgBox "Test 2 Successful"


Call script_end_procedure("Test Complete.")