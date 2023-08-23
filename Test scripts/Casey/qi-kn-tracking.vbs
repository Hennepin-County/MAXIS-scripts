'STATS GATHERING=============================================================================================================
name_of_script = "TYPE - NEW SCRIPT TEMPLATE.vbs"       'Replace TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer

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


'Required for statistical purposes===========================================================================================
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 1            'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block==========================================================================================================

'DIALOGS FOR THE SCRIPT======================================================================================================
    '------Paste any dialogs needed in from the dialog editor here. Dialogs typically include MAXIS_case_number and worker_signature fields


'END DIALOGS=================================================================================================================

'TAGS TO FOLLOW
'FOLLOW UP QUESTION

'THE SCRIPT==================================================================================================================

'Connects to BlueZone
EMConnect ""

'Grabs the MAXIS case number automatically
CALL MAXIS_case_number_finder(MAXIS_case_number)			'here we find a case number
contact_date = date & ""
contact_end_time = time & ""
Call find_user_name(qi_worker)


'First dialog - capture the worker, date, case, time, what we are doing.
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 336, 205, "Knowledge Now Tracking"
  EditBox 90, 20, 130, 15, contact_with
  EditBox 90, 40, 50, 15, contact_date
  EditBox 90, 60, 50, 15, contact_start_time
  EditBox 90, 80, 50, 15, contact_end_time
  EditBox 75, 120, 50, 15, MAXIS_case_number
  CheckBox 75, 140, 180, 10, "Check here if no Case number was provided", no_case_number_checkbox
  DropListBox 75, 155, 190, 45, "Case Support"+chr(9)+"PRIV Case Transfer", kn_proess
  EditBox 100, 185, 110, 15, qi_worker
  ButtonGroup ButtonPressed
    OkButton 225, 185, 50, 15
    CancelButton 280, 185, 50, 15
  GroupBox 5, 10, 325, 95, "Contact Detail"
  Text 35, 25, 55, 10, "KN Contact with"
  Text 225, 25, 90, 10, "(Copy from Teams or Email)"
  Text 30, 45, 60, 10, "Date of Request:"
  Text 20, 65, 70, 10, "Intital Request Time:"
  Text 145, 65, 110, 10, "(email or Teams chat time stamp)"
  Text 30, 85, 60, 10, " Completed Time:"
  GroupBox 5, 110, 325, 65, "Issue Detail"
  Text 25, 125, 50, 10, "Case Number:"
  Text 10, 160, 60, 10, "Type of Contact:"
  Text 10, 190, 90, 10, "QI Knowledge Now Worker"
EndDialog

DO
	err_msg = ""                                       'Blanks this out every time the loop runs. If mandatory fields aren't entered, this variable is updated below with messages, which then display for the worker.
	Dialog Dialog1                               'The Dialog command shows the dialog. Replace sample_dialog with your actual dialog pasted above.
	cancel_without_confirmation

	contact_with = trim(contact_with)
    'Handling for error messaging (in the case of mandatory fields or fields requiring a specific format)-----------------------------------
	Call validate_MAXIS_case_number(err_msg, "*")
	If no_case_number_checkbox = checked and kn_proess = "Case Support" Then err_msg = ""

	If no_case_number_checkbox = checked  AND trim(MAXIS_case_number) <> "" Then err_ms = err_msg & vbCr & "* You have indicated that no Case Number was provided but entered detail in the Case Number field. Either remove the Case Number information or uncheck the box."
	If InStr(contact_start_time, ":") = 0 Then
		err_msg = err_msg & vbCr & "* Enter the start time in hours and minutes using a colon between the hours and minutes."
	Else
		If UCase(right(contact_start_time, 1)) <> "M" Then
			time_array = split(contact_start_time, ":")
			start_hour = time_array(0)
			start_hour = start_hour * 1
			If start_hour < 6 or start_hour = 12 Then contact_start_time = contact_start_time & " PM"
			If start_hour > 5 and start_hour <> 12 Then contact_start_time = contact_start_time & " AM"
		End If
	End If
	If InStr(contact_end_time, ":") = 0 Then
		err_ms = err_msg & vbCr & "* Enter the end time in hours and minutes using a colon between the hours and minutes."
	Else
		If UCase(right(contact_end_time, 1)) <> "M" Then
			time_array = split(contact_end_time, ":")
			start_hour = time_array(0)
			start_hour = start_hour * 1
			If start_hour < 6 or start_hour = 12 Then contact_end_time = contact_end_time & " PM"
			If start_hour > 5 and start_hour <> 12 Then contact_end_time = contact_scontact_end_timetart_time & " AM"
		End If
	End If
    'If the error message isn't blank, it'll pop up a message telling you what to do!
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue."     '
	err_msg = "LOOP"
LOOP UNTIL err_msg = ""     'It only exits the loop when all mandatory fields are resolved!



	'If NO Case Number - how do we handle for this
	'TIME reports - can we add handling for AM/PM if it is not entered Core hours 8 - 4 should have repeats - DO NOT mandate AM/PM
	'Check if start time is aftter end time and don't allow that
	'can check the box for no case number for case support option
	'We only have 2 options for what we are doing:
		'PRIV Case transfer - MUST HAVE CASE NUMBER
		'Case Supports - questions/information

'IF PRIV Case transfer -log detail AND stopr script
If kn_proess = "PRIV Case Transfer" Then

	Call script_end_procedure("PRIV Case Transfer has been logged and script run is completed.")
End If

'Between the dialogs if there is a case number - capture MAXIS detail
	'Current active programs
	'Current pending programs.
	'REVW data?
	'Interview data?
'If NO Case number - add some checkboxes to check off program specifics
'Second dialog for 'Case Support' option
	'display found program information OR program check boxes
	'million check boxes
	'Oped edit box?
	'forward to script team - checkbox


















'Shows dialog (replace "sample_dialog" with the actual dialog you entered above)----------------------------------

'End dialog section-----------------------------------------------------------------------------------------------

' MsgBox "Information saved to the Knowledge Now Database"
' 'Checks Maxis for password prompt
' CALL check_for_MAXIS(True)
'
' 'Now it navigates to a blank case note
' start_a_blank_case_note
'
' '...and enters a title (replace variables with your own content)...
' CALL write_variable_in_case_note("*** CASE NOTE HEADER ***")
'
' '...some editboxes or droplistboxes (replace variables with your own content)...
' CALL write_bullet_and_variable_in_case_note( "Here's the first bullet",  a_variable_from_your_dialog        )
' CALL write_bullet_and_variable_in_case_note( "Here's another bullet",    another_variable_from_your_dialog  )
'
' '...checkbox responses (replace variables with your own content)...
' If some_checkbox_from_your_dialog = checked     then CALL write_variable_in_case_note( "* The checkbox was checked."     )
' If some_checkbox_from_your_dialog = unchecked   then CALL write_variable_in_case_note( "* The checkbox was not checked." )
'
' '...and a worker signature.
' CALL write_variable_in_case_note("---")
' CALL write_variable_in_case_note(worker_signature)

'End the script. Put any success messages in between the quotes, *always* starting with the word "Success!"
script_end_procedure("Information saved to the Knowledge Now Database")
