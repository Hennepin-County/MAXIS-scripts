'Required for statistical purposes===============================================================================
name_of_script = "BULK - CEI PREMIUM NOTER.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 38                      'manual run time in seconds
STATS_denomination = "C"       						 'C is for each CASE
'END OF stats block==============================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'CUSTOM FUNCTIONS---------------------------------------------------------------------------------------------
'This one creates a quasi-two-dimensional array of all cases, using "|" to split cases and "~" to split case info within cases.
Function combine_CEI_data_to_array(info_array)
	If case_number_01 <> "" then info_array = info_array & case_number_01 & "~" & CEI_amount_01 & "~" & Mo_Yr_01 & "~" & date_01 & "|"
	If case_number_02 <> "" then info_array = info_array & case_number_02 & "~" & CEI_amount_02 & "~" & Mo_Yr_02 & "~" & date_02 & "|"
	If case_number_03 <> "" then info_array = info_array & case_number_03 & "~" & CEI_amount_03 & "~" & Mo_Yr_03 & "~" & date_03 & "|"
	If case_number_04 <> "" then info_array = info_array & case_number_04 & "~" & CEI_amount_04 & "~" & Mo_Yr_04 & "~" & date_04 & "|"
	If case_number_05 <> "" then info_array = info_array & case_number_05 & "~" & CEI_amount_05 & "~" & Mo_Yr_05 & "~" & date_05 & "|"
	If case_number_06 <> "" then info_array = info_array & case_number_06 & "~" & CEI_amount_06 & "~" & Mo_Yr_06 & "~" & date_06 & "|"
End function

'DIALOGS---------------------------------------------------------------------------------------------------------
BeginDialog CEI_premium_dialog, 0, 0, 391, 165, "CEI premium dialog"
  EditBox 55, 5, 70, 15, case_number_01
  EditBox 165, 5, 45, 15, CEI_amount_01
  EditBox 245, 5, 45, 15, Mo_Yr_01
  EditBox 340, 5, 45, 15, date_01
  EditBox 55, 25, 70, 15, case_number_02
  EditBox 165, 25, 45, 15, CEI_amount_02
  EditBox 245, 25, 45, 15, Mo_Yr_02
  EditBox 340, 25, 45, 15, date_02
  EditBox 55, 45, 70, 15, case_number_03
  EditBox 165, 45, 45, 15, CEI_amount_03
  EditBox 245, 45, 45, 15, Mo_Yr_03
  EditBox 340, 45, 45, 15, date_03
  EditBox 55, 65, 70, 15, case_number_04
  EditBox 165, 65, 45, 15, CEI_amount_04
  EditBox 245, 65, 45, 15, Mo_Yr_04
  EditBox 340, 65, 45, 15, date_04
  EditBox 55, 85, 70, 15, case_number_05
  EditBox 165, 85, 45, 15, CEI_amount_05
  EditBox 245, 85, 45, 15, Mo_Yr_05
  EditBox 340, 85, 45, 15, date_05
  EditBox 55, 105, 70, 15, case_number_06
  EditBox 165, 105, 45, 15, CEI_amount_06
  EditBox 245, 105, 45, 15, Mo_Yr_06
  EditBox 340, 105, 45, 15, date_06
  EditBox 80, 145, 50, 15, check_will_be_mailed_date
  EditBox 215, 145, 50, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 280, 145, 50, 15
    CancelButton 335, 145, 50, 15
    PushButton 5, 125, 85, 15, "Need more lines?", need_more_lines_button
    PushButton 220, 125, 70, 15, "Prefill months?", prefill_months_button
  Text 5, 10, 50, 10, "Case number: "
  Text 135, 10, 30, 10, "CEI amt:"
  Text 220, 10, 25, 10, "Month:"
  Text 300, 10, 35, 10, "Submitted:"
  Text 5, 30, 50, 10, "Case number: "
  Text 135, 30, 30, 10, "CEI amt:"
  Text 220, 30, 25, 10, "Month:"
  Text 300, 30, 35, 10, "Submitted:"
  Text 5, 50, 50, 10, "Case number: "
  Text 135, 50, 30, 10, "CEI amt:"
  Text 220, 50, 25, 10, "Month:"
  Text 300, 50, 35, 10, "Submitted:"
  Text 5, 70, 50, 10, "Case number: "
  Text 135, 70, 30, 10, "CEI amt:"
  Text 220, 70, 25, 10, "Month:"
  Text 300, 70, 35, 10, "Submitted:"
  Text 5, 90, 50, 10, "Case number: "
  Text 135, 90, 30, 10, "CEI amt:"
  Text 220, 90, 25, 10, "Month:"
  Text 300, 90, 35, 10, "Submitted:"
  Text 5, 110, 50, 10, "Case number: "
  Text 135, 110, 30, 10, "CEI amt:"
  Text 220, 110, 25, 10, "Month:"
  Text 300, 110, 35, 10, "Submitted:"
  Text 5, 150, 70, 10, "Check will be mailed:"
  Text 140, 150, 75, 10, "Sign your case notes:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""

'Shows dialog, allows for cancel, and checks for MAXIS
Do
	Do
		Do
			Dialog CEI_premium_dialog
			cancel_confirmation
			'If the user selects the prefill months option, it'll add the current month to the dialog
			If buttonpressed = prefill_months_button then
				prefill_date = datepart("m", dateadd("m", -1, date)) & "/" & datepart("yyyy", dateadd("m", -1, date))		'Determines the date
				If instr(prefill_date, "/") = 2 then prefill_date = "0" & prefill_date		'Adding the zero if the month is a single digit
				'Inserts the above date in when there's already a case number in each field
				If case_number_01 <> "" then mo_yr_01 = prefill_date
				If case_number_02 <> "" then mo_yr_02 = prefill_date
				If case_number_03 <> "" then mo_yr_03 = prefill_date
				If case_number_04 <> "" then mo_yr_04 = prefill_date
				If case_number_05 <> "" then mo_yr_05 = prefill_date
				If case_number_06 <> "" then mo_yr_06 = prefill_date
			End if
		Loop until buttonpressed <> prefill_months_button
		'Now, it checks to make sure each case number has info, and that no info exists without a case number.
		'It uses a true/false system to make the do...loop simpler with less code.
		If (case_number_01 = "" and (CEI_amount_01 <> "" or Mo_Yr_01 <> "" or date_01 <> "")) or _
		(case_number_02 = "" and (CEI_amount_02 <> "" or Mo_Yr_02 <> "" or date_02 <> "")) or _
		(case_number_03 = "" and (CEI_amount_03 <> "" or Mo_Yr_03 <> "" or date_03 <> "")) or _
		(case_number_04 = "" and (CEI_amount_04 <> "" or Mo_Yr_04 <> "" or date_04 <> "")) or _
		(case_number_05 = "" and (CEI_amount_05 <> "" or Mo_Yr_05 <> "" or date_05 <> "")) or _
		(case_number_06 = "" and (CEI_amount_06 <> "" or Mo_Yr_06 <> "" or date_06 <> "")) then
			MsgBox "You either have a case number without CEI, mo/yr, or date info, OR you have CEI info, mo/yr info, or date info without a case number." & chr(10) & chr(10) & "Please make sure you include required info for each case number, and do not enter info on this dialog without a case number."
			dialog_complete = False
		ElseIf (case_number_01 <> "" and (CEI_amount_01 = "" or Mo_Yr_01 = "" or date_01 = "")) or _
		(case_number_02 <> "" and (CEI_amount_02 = "" or Mo_Yr_02 = "" or date_02 = "")) or _
		(case_number_03 <> "" and (CEI_amount_03 = "" or Mo_Yr_03 = "" or date_03 = "")) or _
		(case_number_04 <> "" and (CEI_amount_04 = "" or Mo_Yr_04 = "" or date_04 = "")) or _
		(case_number_05 <> "" and (CEI_amount_05 = "" or Mo_Yr_05 = "" or date_05 = "")) or _
		(case_number_06 <> "" and (CEI_amount_06 = "" or Mo_Yr_06 = "" or date_06 = "")) then
			MsgBox "You either have a case number without CEI, mo/yr, or date info, OR you have CEI info, mo/yr info, or date info without a case number." & chr(10) & chr(10) & "Please make sure you include required info for each case number, and do not enter info on this dialog without a case number."
			dialog_complete = False
		Else
			dialog_complete = True
		End if
	Loop until dialog_complete = True
	'If the user selects the "need more lines" button, it'll add the existing data to an array and clear the dialog.
	If buttonpressed = need_more_lines_button then
		add_info_to_array_msgbox = MsgBox("This will clear your existing info, moving it into the computer memory, and clearing the lines on this dialog. Is this OK?", vbYesNo)
		If add_info_to_array_msgbox = vbYes then
			'Combine the info to the array
			call combine_CEI_data_to_array(info_array)
			'Clear the existing dialog info
			case_number_01 = ""
			case_number_02 = ""
			case_number_03 = ""
			case_number_04 = ""
			case_number_05 = ""
			case_number_06 = ""
			CEI_amount_01 = ""
			CEI_amount_02 = ""
			CEI_amount_03 = ""
			CEI_amount_04 = ""
			CEI_amount_05 = ""
			CEI_amount_06 = ""
			Mo_Yr_01 = ""
			Mo_Yr_02 = ""
			Mo_Yr_03 = ""
			Mo_Yr_04 = ""
			Mo_Yr_05 = ""
			Mo_Yr_06 = ""
			date_01 = ""
			date_02 = ""
			date_03 = ""
			date_04 = ""
			date_05 = ""
			date_06 = ""
		End if
	End if

Loop until buttonpressed = OK

'checking for an active MAXIS session
Call check_for_MAXIS(False)

'Heading back to self
back_to_self

'Creates a quasi-two-dimensional array of all cases, using "|" to split cases and "~" to split case info within cases. Function declared above.
call combine_CEI_data_to_array(info_array)

'Splits the array
info_array = split(info_array, "|")

'Now the script will go to case note the contents of each case listed.
For each case_info in info_array

	'Goes into each line of the array, skipping blank cases
	If case_info <> "" then

		STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter

		'Splits the case_info variable into an array containing (0) MAXIS_case_number, (1) CEI_amount, (2) mo_yr, and (3) date_sent
		case_specific_info_array = split(case_info, "~")	'That's the character we used above to designate objects for the array

		'Assigns value to each variable needed for the next part
		MAXIS_case_number = case_specific_info_array(0)
		CEI_amount = case_specific_info_array(1)
		mo_yr = case_specific_info_array(2)
		date_sent = case_specific_info_array(3)

		'navigates to case/curr
		call navigate_to_MAXIS_screen("case", "curr")

		'Checking to make sure case is active. This will skip cases without MA or IMD active.
		row = 1											'Declaring prior to the EMSearch feature
		col = 1
		EMSearch "MA: ACTIVE", row, col					'Searching for MA: ACTIVE
		If row = 0 then								'If not found... search again for IMD: ACTIVE
			row = 1										'Declaring again
			col = 1
			EMSearch "IMD: ACTIVE", row, col			'Searching for IMD: ACTIVE. If still not found, lets worker know on next line.
			If row = 0 then MsgBox "This case is not open on MA or IMD. Check the case status for this client after the script runs. The script will skip this client and move on to the next case."
		End if

		'If it was found the entire time, then make that case note.
		If row <> 0 then
			'Navigates to case note and creates a new case note.
			call start_a_blank_CASE_NOTE
			'Now it is case noting the contents.
			Call write_variable_in_CASE_NOTE("CEI reimbursement for " & Mo_Yr & " sent to fiscal" & " on " & date_sent)
			call write_bullet_and_variable_in_case_note("CEI amount", CEI_amount)
			call write_bullet_and_variable_in_case_note("Check will be mailed", check_will_be_mailed_date)
			call write_variable_in_CASE_NOTE("---")
			call write_variable_in_CASE_NOTE(worker_signature)
		End if
	End if
Next

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Your cases have been case noted! Don't forget to send the authorization for payment forms. See a supervisor for more information.")
