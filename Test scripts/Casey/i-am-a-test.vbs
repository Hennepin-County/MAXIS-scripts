'Required for statistical purposes==========================================================================================
name_of_script = "I am a test.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================
run_locally = TRUE
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


' function script_end_procedure_with_error_report(closing_message)
' '--- This function is how all user stats are collected when a script ends.
' '~~~~~ closing_message: message to user in a MsgBox that appears once the script is complete. Example: "Success! Your actions are complete."
' '===== Keywords: MAXIS, MMIS, PRISM, end, script, statistics, stopscript
' 	stop_time = timer
'     send_error_message = ""
' 	If closing_message <> "" AND left(closing_message, 3) <> "~PT" then        '"~PT" forces the message to "pass through", i.e. not create a pop-up, but to continue without further diversion to the database, where it will write a record with the message
'         send_error_message = MsgBox(closing_message & vbNewLine & vbNewLine & "Do you need to send an error report about this script run?", vbSystemModal + vbDefaultButton2 + vbYesNo, "Script Run Completed")
'     End If
'     script_run_time = stop_time - start_time
' 	If is_county_collecting_stats  = True then
' 		'Getting user name
' 		Set objNet = CreateObject("WScript.NetWork")
' 		user_ID = objNet.UserName
'
' 		'Setting constants
' 		Const adOpenStatic = 3
' 		Const adLockOptimistic = 3
'
'         'Determining if the script was successful
'         If closing_message = "" or left(ucase(closing_message), 7) = "SUCCESS" THEN
'             SCRIPT_success = -1
'         else
'             SCRIPT_success = 0
'         end if
'
' 		'Determines if the value of the MAXIS case number - BULK and UTILITIES scripts will not have case number informaiton input into the database
' 		IF left(name_of_script, 4) = "BULK" or left(name_of_script, 4) = "UTIL" then
' 			MAXIS_CASE_NUMBER = ""
' 		End if
'
' 		'Creating objects for Access
' 		Set objConnection = CreateObject("ADODB.Connection")
' 		Set objRecordSet = CreateObject("ADODB.Recordset")
'
' 		'Fixing a bug when the script_end_procedure has an apostrophe (this interferes with Access)
' 		closing_message = replace(closing_message, "'", "")
'
' 		'Opening DB
' 		IF using_SQL_database = TRUE then
'     		objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" & stats_database_path & ""
' 		ELSE
' 			objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & "" & stats_database_path & ""
' 		END IF
'
'         'Adds some data for users of the old database, but adds lots more data for users of the new.
'         If STATS_enhanced_db = false or STATS_enhanced_db = "" then     'For users of the old db
'     		'Opening usage_log and adding a record
'     		objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX)" &  _
'     		"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & script_run_time & ", '" & closing_message & "')", objConnection, adOpenStatic, adLockOptimistic
' 		'collecting case numbers counties
' 		Elseif collect_MAXIS_case_number = true then
' 			objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX, STATS_COUNTER, STATS_MANUALTIME, STATS_DENOMINATION, WORKER_COUNTY_CODE, SCRIPT_SUCCESS, CASE_NUMBER)" &  _
' 			"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & abs(script_run_time) & ", '" & closing_message & "', " & abs(STATS_counter) & ", " & abs(STATS_manualtime) & ", '" & STATS_denomination & "', '" & worker_county_code & "', " & SCRIPT_success & ", '" & MAXIS_CASE_NUMBER & "')", objConnection, adOpenStatic, adLockOptimistic
' 		 'for users of the new db
' 		Else
'             objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX, STATS_COUNTER, STATS_MANUALTIME, STATS_DENOMINATION, WORKER_COUNTY_CODE, SCRIPT_SUCCESS)" &  _
'             "VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & abs(script_run_time) & ", '" & closing_message & "', " & abs(STATS_counter) & ", " & abs(STATS_manualtime) & ", '" & STATS_denomination & "', '" & worker_county_code & "', " & SCRIPT_success & ")", objConnection, adOpenStatic, adLockOptimistic
'         End if
'
' 		'Closing the connection
' 		objConnection.Close
' 	End if
'
'     If send_error_message = vbYes Then
'         'dialog here to gather more detail
'         Dialog1 = ""
'         BeginDialog Dialog1, 0, 0, 401, 175, "Report Error Detail"
'           Text 60, 35, 55, 10, MAXIS_case_number
'           ComboBox 220, 30, 175, 45, ""+chr(9)+"BUG - somethng happened that was wrong"+chr(9)+"ENHANCEMENT - somthing could be done better"+chr(9)+"TYPO - gramatical/spelling type errors", error_type
'           EditBox 65, 50, 330, 15, error_detail
'           CheckBox 20, 100, 65, 10, "CASE/NOTE", case_note_checkbox
'           CheckBox 95, 100, 65, 10, "Update in STAT", stat_update_checkbox
'           CheckBox 170, 100, 75, 10, "Problems with Dates", date_checkbox
'           CheckBox 265, 100, 65, 10, "Math is incorrect", math_checkbox
'           CheckBox 20, 115, 65, 10, "TIKL is incorrect", tikl_checkbox
'           CheckBox 95, 115, 65, 10, "MEMO or WCOM", memo_wcom_checkbox
'           CheckBox 170, 115, 75, 10, "Created Document", document_checkbox
'           CheckBox 265, 115, 115, 10, "Missing a place for Information", missing_spot_checkbox
'           EditBox 60, 140, 165, 15, worker_signature
'           ButtonGroup ButtonPressed
'             OkButton 290, 140, 50, 15
'             CancelButton 345, 140, 50, 15
'           Text 10, 10, 300, 10, "Information is needed about the error for our scriptwriters to review and resolve the issue. "
'           Text 5, 35, 50, 10, "Case Number:"
'           Text 125, 35, 95, 10, "What type of error occured?"
'           Text 5, 55, 60, 10, "Explain in detail:"
'           GroupBox 10, 75, 380, 60, "Common areas of issue"
'           Text 20, 85, 200, 10, "Check any that were impacted by the error you are reporting."
'           Text 10, 145, 50, 10, "Worker Name:"
'           Text 25, 160, 335, 10, "*** Remember to leave the case as is if possible. We can resolve error better when in a live case. ***"
'         EndDialog
'
'         Dialog Dialog1
'
'         'sent email here
'         If ButtonPressed = -1 Then
'             bzt_email = "HSPH.EWS.BlueZoneScripts@hennepin.us"
'             subject_of_email = "Script Error -- " & name_of_script & " (Automated Report)"
'
'             full_text = "Error occured on " & date & " at " & time
'             full_text = full_text & vbCr & "Error type - " & error_type
'             full_text = full_text & vbCr & "Script name - " & name_of_script & " was run on Case #" & MAXIS_case_number & " with a runtime of " & script_run_time & " seconds."
'             full_text = full_text & vbCr & "Information: " & error_detail
'             If case_note_checkbox = checked OR stat_update_checkbox = checked OR date_checkbox = checked OR math_checkbox = checked OR tikl_checkbox = checked OR memo_wcom_checkbox = checked OR document_checkbox = checked OR missing_spot_checkbox = checked Then full_text = full_text & vbCr & vbCr & "Script has issues/concerns in the following areas:"
'
'             If case_note_checkbox = checked Then full_text = full_text & vbCr & " - CASE/NOTE"
'             If stat_update_checkbox = checked Then full_text = full_text & vbCr & " - Update in STAT"
'             If date_checkbox = checked Then full_text = full_text & vbCr & " - Dates are incorrect"
'             If math_checkbox = checked Then full_text = full_text & vbCr & " - Math is incorrect"
'             If tikl_checkbox = checked Then full_text = full_text & vbCr & " - TIKL"
'             If memo_wcom_checkbox = checked Then full_text = full_text & vbCr & " - NOTICES (WCOM/MEMO)"
'             If document_checkbox = checked Then full_text = full_text & vbCr & " - The Excel or Word Document"
'             If missing_spot_checkbox = checked Then full_text = full_text & vbCr & " - There is no space to enter particular information"
'
'             full_text = full_text & vbCr & vbCr & "Sent by: " & worker_signature
'
'             If script_run_lowdown <> "" Then full_text = full_text & vbCr & vbCr & "All Script Run Details:" & vbCr & script_run_lowdown
'
'             Call create_outlook_email(bzt_email, "", subject_of_email, full_text, "", true)
'
'             MsgBox "Error Report completed!" & vbNewLine & vbNewLine & "Thank you for working with us for Continuous Improvement."
'         Else
'             MsgBox "Your error report has been cancelled and has NOT been sent to the BlueZone Script Team"
'         End If
'     End If
' 	If disable_StopScript = FALSE or disable_StopScript = "" then stopscript
' end function


' Call MAXIS_case_number_finder(MAXIS_case_number)
'
' Dialog1 = ""
' BeginDialog Dialog1, 0, 0, 126, 55, "Dialog"
'   ButtonGroup ButtonPressed
'     OkButton 70, 30, 50, 15
'   Text 10, 15, 50, 10, "Case Number"
'   EditBox 60, 10, 60, 15, MAXIS_case_number
' EndDialog
'
' Do
'     err_msg = ""
'
'     dialog Dialog1
'     Call validate_MAXIS_case_number(err_msg, "-")
'     If err_msg <> "" Then MsgBox("Please review the foloowing in order for the script to continue:" & vbNewLine & err_msg)
'
' Loop until err_msg = ""
'
' Call start_a_blank_CASE_NOTE
' notes_variable = "03/19 for 01 is BANKED MONTH - Banked Month: 3.; 04/19 for 01 is BANKED MONTH - Banked Month: 4.;"
' bullet_variable = "This is where the bullet would be all the things."
' time_variable = "Now is the time and this is the place."
' order_variable = "Everything in it's place."
'
' Call write_variable_in_CASE_NOTE("*** SNAP approved starting in 03/19 ***")
' Call write_variable_in_CASE_NOTE("* SNAP approved for 03/19")
' Call write_variable_in_CASE_NOTE("    Eligible Household Members: 01,")
' Call write_variable_in_CASE_NOTE("    Income: Earned: $522.00 Unearned: $0.00")
' Call write_variable_in_CASE_NOTE("    Shelter Costs: $0.00")
' Call write_variable_in_CASE_NOTE("    SNAP BENEFTIT: $115.00 Reporting Status: NON-HRF")
' Call write_variable_in_CASE_NOTE("* SNAP approved for 04/19")
' Call write_variable_in_CASE_NOTE("    Eligible Household Members: 01,")
' Call write_variable_in_CASE_NOTE("    Income: Earned: $522.00 Unearned: $0.00")
' Call write_variable_in_CASE_NOTE("    Shelter Costs: $0.00")
' Call write_variable_in_CASE_NOTE("    SNAP BENEFTIT: $115.00 Reporting Status: NON-HRF")
' Call write_bullet_and_variable_in_CASE_NOTE("Notes", notes_variable)
' Call write_variable_in_CASE_NOTE("This is a thing")
' Call write_variable_in_CASE_NOTE("   this is another thing")
' Call write_variable_in_CASE_NOTE("How now brown cow")
' Call write_variable_in_CASE_NOTE("the thing and thing and stuff")
' Call write_variable_in_CASE_NOTE("all the writing")
' Call write_variable_in_CASE_NOTE("blah blah blah")
' Call write_bullet_and_variable_in_CASE_NOTE("BULLET", bullet_variable)
' Call write_bullet_and_variable_in_CASE_NOTE("Time", time_variable)
' Call write_bullet_and_variable_in_CASE_NOTE("Order", order_variable)
'
' Call write_variable_in_CASE_NOTE("H.Lamb/QI")

' script_list_URL = "C:\MAXIS-scripts\Test scripts\Casey\User Group\COMPLETE LIST OF TESTERS.vbs"
' Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
' Set fso_command = run_another_script_fso.OpenTextFile(script_list_URL)
' text_from_the_other_script = fso_command.ReadAll
' fso_command.Close
' Execute text_from_the_other_script

' Call confirm_tester_information

'Initial Dialog which requests a file path for the excel file
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 361, 105, "On Demand Recertifications"
  EditBox 130, 60, 175, 15, recertification_cases_excel_file_path
  ButtonGroup ButtonPressed
    PushButton 310, 60, 45, 15, "Browse...", select_a_file_button
  EditBox 75, 85, 140, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 250, 85, 50, 15
    CancelButton 305, 85, 50, 15
  Text 10, 10, 170, 10, "Welcome to the On Demand Recertification Notifier."
  Text 10, 25, 340, 30, "This script will send an Appointment Notice or NOMI for recertification for a list of cases in a county that currently has an On Demand Waiver in effect for interviews. If your county does not have this waiver, this script should not be used."
  Text 10, 65, 120, 10, "Select an Excel file for recert cases:"
  Text 10, 90, 60, 10, "Worker Signature"
EndDialog


'Confirmation Diaglog will require worker to afirm the appointment notices/NOMIs should actually be sent

'END DIALOGS ===============================================================================================================

'SCRIPT ====================================================================================================================
'Connects to BlueZone
EMConnect ""

'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
'Show initial dialog
Do
	Dialog Dialog1
	If ButtonPressed = cancel then stopscript
	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(recertification_cases_excel_file_path, ".xlsx")
Loop until ButtonPressed = OK and recertification_cases_excel_file_path <> "" and worker_signature <> ""

'Opens Excel file here, as it needs to populate the dialog with the details from the spreadsheet.
call excel_open(recertification_cases_excel_file_path, True, True, ObjExcel, objWorkbook)

'Set objWorkSheet = objWorkbook.Worksheet
For Each objWorkSheet In objWorkbook.Worksheets
	If instr(objWorkSheet.Name, "Sheet") = 0 and objWorkSheet.Name <> "controls" then scenario_list = scenario_list & chr(9) & objWorkSheet.Name
Next

'Dialog to select worksheet
'DIALOG is defined here so that the dropdown can be populated with the above code
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 151, 75, "On Demand Recertifications"
  DropListBox 5, 35, 140, 15, "Select One..." & scenario_list, scenario_dropdown
  ButtonGroup ButtonPressed
    OkButton 40, 55, 50, 15
    CancelButton 95, 55, 50, 15
  Text 5, 10, 130, 20, "Select the correct worksheet to run for recertification interview notifications:"
EndDialog

'Shows the dialog to select the correct worksheet
Do
    Dialog Dialog1
    If ButtonPressed = cancel then stopscript
Loop until scenario_dropdown <> "Select One..."

objExcel.worksheets(scenario_dropdown).Activate

excel_row = 2
leave_loop = FALSE
Do
    Call back_to_SELF
    MAXIS_case_number = trim(ObjExcel.Cells(excel_row, 2).value)

    If MAXIS_case_number <> "" Then
        Call navigate_to_MAXIS_screen("STAT", "REVW")

        EMReadScreen cash_revw_status, 1, 7, 40
        EMReadScreen snap_revw_status, 1, 7, 60
        EMReadScreen hc_revw_status, 1, 7, 73

        If cash_revw_status = "U" Then leave_loop = TRUE
        If snap_revw_status = "U" Then leave_loop = TRUE
        If hc_revw_status = "U" Then leave_loop = TRUE
        If cash_revw_status = "A" Then leave_loop = TRUE
        If snap_revw_status = "A" Then leave_loop = TRUE
        If hc_revw_status = "A" Then leave_loop = TRUE

    Else
        leave_loop = TRUE
    End If
    MAXIS_case_number = ""
    excel_row = excel_row + 1

Loop until leave_loop = TRUE

Call script_end_procedure_with_error_report("The End")
