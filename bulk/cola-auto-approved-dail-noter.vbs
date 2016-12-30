'Required for statistical purposes===============================================================================
name_of_script = "BULK - COLA AUTO APPROVED DAIL NOTER.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 95                               'manual run time in seconds
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

'DIALOGS----------------------------------------------------------------------------------------------
BeginDialog Auto_Approved_COLA_DAIL_Message_Dialog, 0, 0, 251, 185, "Auto Approved COLA DAIL Message"
  Text 5, 15, 240, 20, "Which of the following AUTO APPROVED COLA DAIL messages do you want to delete and case note?"
  CheckBox 35, 45, 35, 15, "SNAP", SNAP_COLA_Message_Checkbox
  CheckBox 35, 60, 35, 15, "GRH", GRH_COLA_Message_Checkbox
  CheckBox 35, 75, 35, 15, "MSA", MSA_COLA_Message_Checkbox
  Text 5, 115, 70, 10, "Sign your case note"
  EditBox 90, 115, 65, 15, Worker_Signature
  ButtonGroup ButtonPressed
	OkButton 135, 150, 50, 15
	CancelButton 190, 150, 50, 15
EndDialog

BeginDialog worker_sig_dlg, 0, 0, 211, 135, "COLA Scrubber"
  EditBox 110, 10, 65, 15, worker_signature
  EditBox 110, 30, 65, 15, x_number
  CheckBox 10, 55, 165, 10, "Check here to have the script delete the DAIL", delete_dail_check
  CheckBox 10, 70, 195, 10, "Check here to have the script run on ALL COLA messages", bulk_check
  ButtonGroup ButtonPressed
    OkButton 105, 110, 50, 15
    CancelButton 155, 110, 50, 15
  Text 20, 85, 160, 10, "NOTE: This option also creates a report in Excel."
  Text 10, 30, 95, 20, "Please enter an X number (7 digit format):"
  Text 10, 15, 95, 10, "Please sign your case note:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------------
'Connecting to BlueZone
EMConnect ""

'Grabbing worker signature
DO
	DIALOG worker_sig_dlg
		IF ButtonPressed = cancel THEN stopscript
		IF worker_signature = "" THEN MsgBox "Sign your case note."
		IF UCASE(left(x_number, 1)) <> "X" or len(x_number) <> 7 THEN MsgBox "Please enter a valid X number."
LOOP UNTIL worker_signature <> "" AND UCASE(left(x_number, 1)) = "X" AND len(x_number) = 7

IF bulk_check = checked THEN
	CALL check_for_MAXIS(FALSE)
	'Creating the Excel file for tracking
	SET objExcel = CreateObject("Excel.Application")
	objExcel.Visible = FALSE
	SET objWorkbook = objExcel.Workbooks.Add()
	objExcel.DisplayAlerts = TRUE

	objExcel.Cells(1, 1).Value = "MAXIS CASE NUMBER"
	objExcel.Cells(1, 2).Value = "COLA MESSAGE"

	CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
	EMWriteScreen x_number, 21, 6
	transmit

	'Having the script display ONLY COLA messages
	EMWriteScreen "X", 4, 12
	transmit

	EMWriteScreen "_", 7, 39	'"ALL"
	EMWriteScreen "X", 8, 39	'"COLA"
	EMWriteScreen "X", 13, 39	'"INFO" because GRH auto-approved is an INFO message rather than COLA
	pick_row = 9
	DO
		EMWriteScreen "_", pick_row, 39		'Deselects all the other options
		pick_row = pick_row + 1
		IF pick_row = 13 then pick_row = pick_row + 1 'Skips info row when clearing all other picked rows.
	LOOP UNTIL pick_row = 21
	transmit

	dail_row = 6
	excel_row = 2
	DO

		'Checking for a case number.
		check_for_case_number_row = dail_row
		DO
			EMReadScreen look_for_case_number, 8, check_for_case_number_row, 63
			IF look_for_case_number = "CASE NBR" THEN
				EMReadScreen MAXIS_case_number, 8, check_for_case_number_row, 73
				MAXIS_case_number = trim(MAXIS_case_number)
			ELSE
				check_for_case_number_row = check_for_case_number_row - 1
			END IF
		LOOP UNTIL look_for_case_number = "CASE NBR"

		'Making sure this is a COLA message and not a client's name.
		DO
			EMReadScreen is_this_a_cola_message, 5, dail_row, 6
			IF is_this_a_cola_message <> "COLA " THEN
				IF is_this_a_cola_message <> "INFO " THEN dail_row = dail_row + 1
				IF dail_row = 19 THEN
					PF8
					EMReadScreen last_page, 21, 24, 2
					dail_row = 6
				END IF
			END IF
		LOOP UNTIL is_this_a_cola_message = "COLA " or is_this_a_cola_message = "INFO " or last_page = "THIS IS THE LAST PAGE"
		IF last_page = "THIS IS THE LAST PAGE" THEN EXIT DO

		EMReadScreen cola_message, 60, dail_row, 20

		IF trim(cola_message) = "SNAP: NEW VERSION AUTO-APPROVED" OR trim(cola_message) = "GRH: NEW VERSION AUTO-APPROVED" OR trim(cola_message) = "NEW MSA ELIG AUTO-APPROVED" THEN
			EMWriteScreen "N", dail_row, 3
			transmit

			'Reading the top case note header to see if it matches. If it does it will end the script as the script has derailed and needs to be stopped.
			EMReadScreen double_message_catch, 55, 5, 25
			IF trim(double_message_catch) = trim(cola_message) Then
				catch_msgbox = msgbox ("Looks like this already was case noted, are you sure about this?", vbYesNo)
				IF catch_msgbox = vbNo then script_end_procedure("An error has occurred and the script is re-reading a dail causing it to duplicate a case note. Please check with your Alpha Script User.")
			END IF

			PF9
			'This is error_msg to determine if you do not have write access to the case.
			EMReadScreen error_msg, 3, 24, 2
			IF error_msg <> "YOU" THEN
				case_note_auto_approval = trim(cola_message)
				CALL write_variable_in_case_note(case_note_auto_approval)
				CALL write_variable_in_case_note("Case was auto approved due to COLA changes")
				CALL write_variable_in_case_note("---")
				CALL write_variable_in_case_note(worker_signature)

				'Navigating back to DAIL/DAIL
				PF3
			END IF
			PF3

			'Resetting dail_row because when the script backs out to DAIL/DAIL, the case will now be the top case on DAIL.
			dail_row = 6
			STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter

			'The case number is now at the top of the DAIL
			IF delete_dail_check = checked THEN
				DO
					EMReadScreen original_message, 60, dail_row, 20
					original_message = trim(original_message)
					IF original_message = case_note_auto_approval THEN
						EMWriteScreen "D", dail_row, 3
						transmit
						EMReadScreen current_user, 7, 21, 73
						IF UCASE(current_user) <> UCASE(x_number) THEN transmit
					ELSEIF original_message = "-------------------------------" THEN script_end_procedure("The original DAIL could not be found.")

					ELSE
						dail_row = dail_row + 1
					END IF
				LOOP UNTIL original_message = case_note_auto_approval
			ELSE
				dail_row = dail_row + 1
			END IF
		ELSE
			dail_row = dail_row + 1  'accounting for when dail message isn't deleted and remains on top. without this it will re-read same message and case note infinitely.
		END IF

		IF dail_row = 19 THEN
			PF8
			EMReadScreen last_page, 21, 24, 2
			dail_row = 6
		END IF

		objExcel.Cells(excel_row, 1).Value = MAXIS_case_number
		objExcel.Cells(excel_row, 2).Value = cola_message
		excel_row = excel_row + 1
	LOOP UNTIL last_page = "THIS IS THE LAST PAGE"
IF objExcel.visible = False THEN objExcel.visible = TRUE
script_end_procedure("Success!")

ELSE
	EMGetCursor read_row, read_col   'must be read first otherwise read for DAIL changes cursor and the maxis check changes cursor

	CALL check_for_MAXIS(FALSE)

	'The code below is a safeguard to make sure the worker is on DAIL and the cursor is set to a COLA DAIL.
	EMReadScreen on_dail, 4, 2, 48
	IF on_dail <> "DAIL" THEN script_end_procedure("You are not in DAIL. Please navigate to DAIL and run the script again.")

	EMReadScreen is_right_line, 4, read_row, 6
	IF is_right_line <> "COLA" THEN			'must be nested otherwise does OR does not work.
		IF is_right_line <> "INFO" THEN script_end_procedure("You are not on the correct line. Please select a COLA/INFO message on your DAIL.")
	END IF

	'Now the script needs to read the specific COLA message to determine what action to take next.
	EMReadScreen cola_message, 60, read_row, 20

	IF trim(cola_message) = "SNAP: NEW VERSION AUTO-APPROVED" OR trim(cola_message) = "GRH: NEW VERSION AUTO-APPROVED" OR trim(cola_message) = "NEW MSA ELIG AUTO-APPROVED" THEN

		'IF the COLA message is for an auto-approved SNAP case, the script will case note that the SNAP was auto-approved and give the worker the option to delete the DAIL.
		EMWriteScreen "N", read_row, 3
		'replacing TRANSMIT with CALL check_for_MAXIS(True) because there is already a TRANSMIT at the start of that function
		transmit

		'Reading the top case note header to see if it matches. If it does it will end the script as the script has derailed and needs to be stopped.
		EMReadScreen double_message_catch, 55, 5, 25
		IF trim(double_message_catch) = trim(cola_message) Then
			catch_msgbox = msgbox ("Looks like this already was case noted, are you sure about this?", vbYesNo)
			IF catch_msgbox = vbNo then script_end_procedure("An error has occurred and the script is re-reading a dail causing it to duplicate a case note. Please check with your Alpha Script User.")
		END IF

		PF9
		case_note_auto_approval = trim(cola_message)
		CALL write_variable_in_case_note(case_note_auto_approval)
		CALL write_variable_in_case_note("Case was auto approved due to COLA changes")
		CALL write_variable_in_case_note("---")
		CALL write_variable_in_case_note(worker_signature)
		'Navigating back to DAIL/DAIL
		PF3
		PF3

		'The case number is now at the top of the DAIL
		IF delete_dail_check = checked THEN
			dail_row = 6
			DO
				EMReadScreen original_message, 31, dail_row, 20
				IF original_message = case_note_auto_approval THEN
					EMWriteScreen "D", dail_row, 3
					transmit
					EMReadScreen current_user, 7, 21, 73
					IF UCASE(current_user) <> UCASE(x_number) THEN transmit
				ELSEIF original_message = "-------------------------------" THEN
					script_end_procedure("The original DAIL could not be found.")
				ELSE
					dail_row = dail_row + 1
				END IF
			LOOP UNTIL original_message = case_note_auto_approval
		END IF

	ELSE
		script_end_procedure("This COLA/INFO message is not yet supported by the script.")
	END IF
	STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
	script_end_procedure("Success!")
END IF
