'Required for statistical purposes===============================================================================
name_of_script = "BULK - RETURNED MAIL.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 64                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
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

'DIALOGS-------------------------------------------------------------------------------------------------------------------
'The bulk-loading-case numbers dialog
BeginDialog many_case_numbers_dialog, 0, 0, 366, 250, "Enter Many Case Numbers Dialog"
  Text 5, 5, 220, 10, "Enter each MAXIS case number, then press ''Next...'' when finished."
  EditBox 5, 20, 55, 15, case_number_01
  DropListBox 65, 20, 55, 20, "No Forwarding "+chr(9)+"Forwarding ", MailType_01
  EditBox 125, 20, 55, 15, case_number_02
  DropListBox 185, 20, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_02
  EditBox 245, 20, 55, 15, case_number_03
  DropListBox 305, 20, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_03
  EditBox 5, 40, 55, 15, case_number_04
  DropListBox 65, 40, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_04
  EditBox 125, 40, 55, 15, case_number_05
  DropListBox 185, 40, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_05
  EditBox 245, 40, 55, 15, case_number_06
  DropListBox 305, 40, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_06
  EditBox 5, 60, 55, 15, case_number_07
  DropListBox 65, 60, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_07
  EditBox 125, 60, 55, 15, case_number_08
  DropListBox 185, 60, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_08
  EditBox 245, 60, 55, 15, case_number_09
  DropListBox 305, 60, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_09
  EditBox 5, 80, 55, 15, case_number_10
  DropListBox 65, 80, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_10
  EditBox 125, 80, 55, 15, case_number_11
  DropListBox 185, 80, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_11
  EditBox 245, 80, 55, 15, case_number_12
  DropListBox 305, 80, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_12
  EditBox 5, 100, 55, 15, case_number_13
  DropListBox 65, 100, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_13
  EditBox 125, 100, 55, 15, case_number_14
  DropListBox 185, 100, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_14
  EditBox 245, 100, 55, 15, case_number_15
  DropListBox 305, 100, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_15
  EditBox 5, 120, 55, 15, case_number_16
  DropListBox 65, 120, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_16
  EditBox 125, 120, 55, 15, case_number_17
  DropListBox 185, 120, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_17
  EditBox 245, 120, 55, 15, case_number_18
  DropListBox 305, 120, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_18
  EditBox 5, 140, 55, 15, case_number_19
  DropListBox 65, 140, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_19
  EditBox 125, 140, 55, 15, case_number_20
  DropListBox 185, 140, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_20
  EditBox 245, 140, 55, 15, case_number_21
  DropListBox 305, 140, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_21
  EditBox 5, 160, 55, 15, case_number_22
  DropListBox 65, 160, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_22
  EditBox 125, 160, 55, 15, case_number_23
  DropListBox 185, 160, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_23
  EditBox 245, 160, 55, 15, case_number_24
  DropListBox 305, 160, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_24
  EditBox 5, 180, 55, 15, case_number_25
  DropListBox 65, 180, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_25
  EditBox 125, 180, 55, 15, case_number_26
  DropListBox 185, 180, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_26
  EditBox 245, 180, 55, 15, case_number_27
  DropListBox 305, 180, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_27
  EditBox 5, 200, 55, 15, case_number_28
  DropListBox 65, 200, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_28
  EditBox 125, 200, 55, 15, case_number_29
  DropListBox 185, 200, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_29
  EditBox 245, 200, 55, 15, case_number_30
  DropListBox 305, 200, 55, 15, "No Forwarding "+chr(9)+"Forwarding ", MailType_30
  ButtonGroup ButtonPressed
    PushButton 305, 230, 50, 15, "Next...", next_button
EndDialog

'THE SCRIPT----------------------------------------------------------------
'Opening message box with info about script and warning to make sure worker is in production in Maxis, not inquiry
warning_box = MsgBox("PLEASE READ!!" & chr(10) & chr(10) & "NEW!!! Script can now handle casees with an allowed forwarding address.  It will case note and TIKL for up-to 30 cases worth of returned mail.  Consult a supervisor if you have questions about returned mail policy." & _
		vbNewline & vbNewline & "NOTE: Make sure you are in production before continuing as Maxis cannot case note or TIKL in inquiry.", vbOKCancel)
If warning_box = vbCancel then stopscript

'Connect to BlueZone
EMConnect ""

'Showing the dialog
DO
	Dialog many_case_numbers_dialog
	Cancel_confirmation
	If (isnumeric(case_number_01) = FALSE and case_number_01 <> "") or (isnumeric(case_number_02) = FALSE and case_number_02 <> "") or _
	  (isnumeric(case_number_03) = FALSE and case_number_03 <> "") or (isnumeric(case_number_04) = FALSE and case_number_04 <> "") or _
	  (isnumeric(case_number_05) = FALSE and case_number_05 <> "") or (isnumeric(case_number_06) = FALSE and case_number_06 <> "") or _
	  (isnumeric(case_number_07) = FALSE and case_number_07 <> "") or (isnumeric(case_number_08) = FALSE and case_number_08 <> "") or _
	  (isnumeric(case_number_09) = FALSE and case_number_09 <> "") or (isnumeric(case_number_10) = FALSE and case_number_10 <> "") or _
	  (isnumeric(case_number_11) = FALSE and case_number_11 <> "") or (isnumeric(case_number_12) = FALSE and case_number_12 <> "") or _
	  (isnumeric(case_number_13) = FALSE and case_number_13 <> "") or (isnumeric(case_number_14) = FALSE and case_number_14 <> "") or _
	  (isnumeric(case_number_15) = FALSE and case_number_15 <> "") or (isnumeric(case_number_16) = FALSE and case_number_16 <> "") or _
	  (isnumeric(case_number_17) = FALSE and case_number_17 <> "") or (isnumeric(case_number_18) = FALSE and case_number_18 <> "") or _
	  (isnumeric(case_number_19) = FALSE and case_number_19 <> "") or (isnumeric(case_number_20) = FALSE and case_number_20 <> "") or _
	  (isnumeric(case_number_21) = FALSE and case_number_21 <> "") or (isnumeric(case_number_22) = FALSE and case_number_22 <> "") or _
	  (isnumeric(case_number_23) = FALSE and case_number_23 <> "") or (isnumeric(case_number_24) = FALSE and case_number_24 <> "") or _
	  (isnumeric(case_number_25) = FALSE and case_number_25 <> "") or (isnumeric(case_number_26) = FALSE and case_number_26 <> "") or _
	  (isnumeric(case_number_27) = FALSE and case_number_27 <> "") or (isnumeric(case_number_28) = FALSE and case_number_28 <> "") or _
	  (isnumeric(case_number_29) = FALSE and case_number_29 <> "") or (isnumeric(case_number_30) = FALSE and case_number_30 <> "") then
		MsgBox "You must enter a numeric case number for each item, or leave it blank."
	End if
Loop until (isnumeric(case_number_01) = True or case_number_01 = "") and (isnumeric(case_number_02) = True or case_number_02 = "") and _
  (isnumeric(case_number_03) = True or case_number_03 = "") and (isnumeric(case_number_04) = True or case_number_04 = "") and _
  (isnumeric(case_number_05) = True or case_number_05 = "") and (isnumeric(case_number_06) = True or case_number_06 = "") and _
  (isnumeric(case_number_07) = True or case_number_07 = "") and (isnumeric(case_number_08) = True or case_number_08 = "") and _
  (isnumeric(case_number_09) = True or case_number_09 = "") and (isnumeric(case_number_10) = True or case_number_10 = "") and _
  (isnumeric(case_number_11) = True or case_number_11 = "") and (isnumeric(case_number_12) = True or case_number_12 = "") and _
  (isnumeric(case_number_13) = True or case_number_13 = "") and (isnumeric(case_number_14) = True or case_number_14 = "") and _
  (isnumeric(case_number_15) = True or case_number_15 = "") and (isnumeric(case_number_16) = True or case_number_16 = "") and _
  (isnumeric(case_number_17) = True or case_number_17 = "") and (isnumeric(case_number_18) = True or case_number_18 = "") and _
  (isnumeric(case_number_19) = True or case_number_19 = "") and (isnumeric(case_number_20) = True or case_number_20 = "") and _
  (isnumeric(case_number_21) = True or case_number_21 = "") and (isnumeric(case_number_22) = True or case_number_22 = "") and _
  (isnumeric(case_number_23) = True or case_number_23 = "") and (isnumeric(case_number_24) = True or case_number_24 = "") and _
  (isnumeric(case_number_25) = True or case_number_25 = "") and (isnumeric(case_number_26) = True or case_number_26 = "") and _
  (isnumeric(case_number_27) = True or case_number_27 = "") and (isnumeric(case_number_28) = True or case_number_28 = "") and _
  (isnumeric(case_number_29) = True or case_number_29 = "") and (isnumeric(case_number_30) = True or case_number_30 = "")

'Worker signature
worker_signature = InputBox("Sign your case note:", vbOKCancel)
If worker_signature = vbCancel then stopscript

'Splits the MAXIS_case_number(s) into a case_number_array
case_number_array = array(case_number_01, case_number_02, case_number_03, case_number_04, case_number_05, _
  case_number_06, case_number_07, case_number_08, case_number_09, case_number_10, _
  case_number_11, case_number_12, case_number_13, case_number_14, case_number_15, _
  case_number_16, case_number_17, case_number_18, case_number_19, case_number_20, _
  case_number_21, case_number_22, case_number_23, case_number_24, case_number_25, _
  case_number_26, case_number_27, case_number_28, case_number_29, case_number_30)
  'End crazy array splitting

'Splits the MailType(s) into a MailType_array
MailType_array = array(MailType_01, MailType_02, MailType_03, MailType_04, MailType_05, _
  MailType_06, MailType_07, MailType_08, MailType_09, MailType_10, _
  MailType_11, MailType_12, MailType_13, MailType_14, MailType_15, _
  MailType_16, MailType_17, MailType_18, MailType_19, MailType_20, _
  MailType_21, MailType_22, MailType_23, MailType_24, MailType_25, _
  MailType_26, MailType_27, MailType_28, MailType_29, MailType_30)
  'End crazy array splitting

'Checking for MAXIS
call check_for_MAXIS(false)

'Setting variables so we can compare between two arrays
array_count=0

For each MAXIS_case_number in case_number_array

	If MAXIS_case_number <> "" then	'skip blanks

		'Getting to case note
		Call navigate_to_MAXIS_screen("case", "note")

		'If there was an error after trying to go to CASE/NOTE
		EMReadScreen SELF_error_check, 27, 2, 28
		If SELF_error_check = "Select Function Menu (SELF)" then
			MsgBox "Script stopped on case " & MAXIS_case_number & "."
			error_message = error_message & MAXIS_case_number & ", " 'Building error message to contain every failed case number, this will allow script to continue if it fails except for out of county cases.
		Else
			'Opening a new case/note
			start_a_blank_CASE_NOTE

			'Writing the case note depending on MailType
			If MailType_array(array_count) = "No Forwarding " then
				EMSendKey "<home>" & "-->Returned mail received<--" & "<newline>"
				call write_variable_in_case_note("* No forwarding address was indicated.")
				call write_variable_in_case_note("* Sending verification request to last known address. TIKLed for 10-day return.")
				call write_variable_in_case_note("(NOTE generated via BULK script)")
				call write_variable_in_case_note("---")
				call write_variable_in_case_note(worker_signature)
			Else
				EMSendKey "<home>" & "-->Returned Mail Received<--" & "<newline>"
				call write_variable_in_case_note("* Forwarding address indicated.")
				call write_variable_in_case_note("* Updated ADDR to match forwarding address.  Forwarded returned mail to current address. Sent appropriate returned mail paperwork for current programs.")
				call write_variable_in_case_note("(NOTE generated via BULK script)")
				call write_variable_in_case_note("---")
				call write_variable_in_case_note(worker_signature)
			End If

			'Exiting the case note
			PF3
			'Getting to DAIL/WRIT
			call navigate_to_MAXIS_screen("dail", "writ")
			'Inserting the date
			call create_MAXIS_friendly_date(date, 10, 5, 18)

			'Writes TIKL depending on MailType
			If MailType_array(array_count) = "No Forwarding " then
				write_variable_in_TIKL("Request for address sent 10 days ago. If not responded to, take appropriate action. (TIKL generated via BULK script)")
			Else
				write_variable_in_TIKL("Returned mail processed 10 days ago.  If verifs were requested and have not been received back, take appropriate action. (TIKL generated via BULK script)")
			End If

			'Exits case note
			PF3

		End if


	End if


	'Increasing the array count for each case number processed from MAXIS_case_number array.
	array_count=array_count+1
	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
Next

'Error message box that lists case numbers that the script failed on so workers can process manually.
If error_message <> "" then msgbox "These cases were not able to be processed, they may be privileged or invalid case numbers. Please review and process manually if needed. " & vbNewline & vbNewline & error_message

'Script ends
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Using " & EDMS_choice & ", send the appropriate returned mail paperwork. Send the completed forms to the most recent address(es). The script has case noted that returned mail was received and TIKLed out for 10-day return for each case indicated.")
