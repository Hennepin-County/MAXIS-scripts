'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - DUPLICATE ASSISTANCE WCOM.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 90                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
'END OF stats block=========================================================================================================

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
call changelog_update("04/04/2017", "Added handling for multiple recipient changes to SPEC/WCOM", "David Courtright, St Louis County")
call changelog_update("01/17/2017", "Added new option to write out of state duplicate assistance alternate text in SNAP WCOMS.", "Charles Potter, DHS")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Dialog--------------------------------------------
BeginDialog dup_dlg, 0, 0, 156, 115, "Duplicate Assistance WCOM"
  EditBox 65, 5, 75, 15, MAXIS_case_number
  EditBox 75, 25, 65, 15, worker_signature
  EditBox 60, 45, 20, 15, MAXIS_footer_month
  EditBox 130, 45, 20, 15, MAXIS_footer_year
  CheckBox 15, 70, 125, 10, "Out of State Duplicate Assistance?", out_of_state_checkbox
  ButtonGroup ButtonPressed
    OkButton 25, 90, 50, 15
    CancelButton 80, 90, 50, 15
  Text 10, 10, 55, 10, "Case Number: "
  Text 10, 30, 60, 10, "Worker Signature: "
  Text 10, 45, 45, 20, "Footer Month (MM):"
  Text 85, 45, 40, 20, "Footer Year (YY):"
EndDialog


'The script-------------------------------------
EMConnect ""
call MAXIS_case_number_finder(MAXIS_case_number)

'warning box
Msgbox "Warning: If you have multiple waiting SNAP or MFIP results this script may be unable to find the most recent one. Please process manually in those instances." & vbNewLine & vbNewLine &_
		"- If this case includes members who are residing in a battered women's shelter please review approval." & vbNewLine &_
		"- If this was an expedited case where client reported they did not receive benefits in another state please review approval" & vbNewLine &_
		"- See CM 001.21 for more details on these two situations and how they qualify for duplicate assistance."

'the dialog
Do
	Do
		Do
			dialog dup_dlg
			cancel_confirmation
			If MAXIS_footer_month = "" or MAXIS_footer_year = "" THEN Msgbox "Please fill in footer month and year (MM YY format)."
			If MAXIS_case_number = "" THEN MsgBox "Please enter a case number."
			If worker_signature = "" THEN MsgBox "Please sign your note."
		Loop until MAXIS_footer_month <> "" & MAXIS_footer_year <> ""
	Loop until MAXIS_case_number <> ""
Loop until worker_signature <> ""

'Converting dates into useable forms
If len(MAXIS_footer_month) < 2 THEN MAXIS_footer_month = "0" & MAXIS_footer_month
If len(MAXIS_footer_year) > 2 THEN MAXIS_footer_year = right(MAXIS_footer_year, 2)

'Navigating to the spec wcom screen
CALL Check_for_MAXIS(false)

back_to_self

Emwritescreen MAXIS_case_number, 18, 43
Emwritescreen MAXIS_footer_month, 20, 43
Emwritescreen MAXIS_footer_year, 20, 46

'This section will check for whether forms go to AREP and SWKR
call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
EMReadscreen forms_to_arep, 1, 10, 45
call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
EMReadscreen forms_to_swkr, 1, 15, 63


CALL navigate_to_MAXIS_screen("SPEC", "WCOM")

'Searching for waiting SNAP notice
wcom_row = 6
Do
	wcom_row = wcom_row + 1
	Emreadscreen program_type, 2, wcom_row, 26
	Emreadscreen print_status, 7, wcom_row, 71
	If program_type = "FS" then
		If print_status = "Waiting" then
			Emwritescreen "x", wcom_row, 13
			Transmit
			PF9
			'The script is now on the recipient selection screen.  Mark all recipients that need NOTICES
			row = 4                             'Defining row and col for the search feature.
			col = 1
			EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
			IF row > 4 THEN  arep_row = row  'locating ALTREP location if it exists'
			row = 4                             'reset row and col for the next search
			col = 1
			EMSearch "SOCWKR", row, col
			IF row > 4 THEN  swkr_row = row     'Logs the row it found the SOCWKR string as swkr_row
			EMWriteScreen "x", 5, 10                                        'We always send notice to client
			IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
			IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
			transmit                                                        'Transmits to start the memo writing process'
			Emreadscreen fs_wcom_exists, 3, 3, 15
			If fs_wcom_exists <> "   " then script_end_procedure ("It appears you already have a WCOM added to this notice. The script will now end.")
			If program_type = "FS" AND print_status = "Waiting" then
				fs_wcom_writen = true
				'This will write if the notice is for SNAP only
				IF out_of_state_checkbox = checked THEN
					state_of_assistance = inputbox("Please enter the other state where client received SNAP")
					CALL write_variable_in_SPEC_MEMO("******************************************************")
					CALL write_variable_in_SPEC_MEMO("Dear Client,")
					CALL write_variable_in_SPEC_MEMO("")
					CALL write_variable_in_SPEC_MEMO("You received SNAP benefits from the state of " & state_of_assistance & " during the month of " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". You cannot receive SNAP benefits from two states at the same time.")
					CALL write_variable_in_SPEC_MEMO("")
					CALL write_variable_in_SPEC_MEMO("If you have any questions or concerns please feel free to contact your worker.")
					CALL write_variable_in_SPEC_MEMO("---")
					CALL write_variable_in_SPEC_MEMO(worker_signature)
					CALL write_variable_in_SPEC_MEMO("")
					CALL write_variable_in_SPEC_MEMO("******************************************************")
				ELSE
					CALL write_variable_in_SPEC_MEMO("******************************************************")
					CALL write_variable_in_SPEC_MEMO("Dear Client,")
					CALL write_variable_in_SPEC_MEMO("")
					CALL write_variable_in_SPEC_MEMO("You will not be eligible for SNAP benefits this month since you have received SNAP benefits on another case for the same month.")
					CALL write_variable_in_SPEC_MEMO("Per program rules SNAP participants are not eligible for duplicate benefits in the same benefit month.")
					CALL write_variable_in_SPEC_MEMO("")
					CALL write_variable_in_SPEC_MEMO("If you have any questions or concerns please feel free to contact your worker.")
					CALL write_variable_in_SPEC_MEMO("---")
					CALL write_variable_in_SPEC_MEMO(worker_signature)
					CALL write_variable_in_SPEC_MEMO("")
					CALL write_variable_in_SPEC_MEMO("******************************************************")
				END IF
				PF4
				PF3
			End if
		End If
	End If
	If fs_wcom_writen = true then Exit Do
	If wcom_row = 17 then
		PF8
		Emreadscreen spec_edit_check, 6, 24, 2
		wcom_row = 6
	end if
	If spec_edit_check = "NOTICE" THEN no_fs_waiting = true
Loop until spec_edit_check = "NOTICE"

program_type = " "
print_status = " "
spec_edit_check = " "

wcom_row = 6
Do
	wcom_row = wcom_row + 1
	Emreadscreen program_type, 2, wcom_row, 26
	Emreadscreen print_status, 7, wcom_row, 71
	If program_type = "MF" then
		If print_status = "Waiting" then
			Emwritescreen "x", wcom_row, 13
			Transmit
			PF9
			'The script is now on the recipient selection screen.  Mark all recipients that need NOTICES
			row = 4                             'Defining row and col for the search feature.
			col = 1
			EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
			IF row > 4 THEN  arep_row = row  'locating ALTREP location if it exists'
			row = 4                             'reset row and col for the next search
			col = 1
			EMSearch "SOCWKR", row, col
			IF row > 4 THEN  swkr_row = row     'Logs the row it found the SOCWKR string as swkr_row
			EMWriteScreen "x", 5, 10                                        'We always send notice to client
			IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
			IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
			transmit                                                        'Transmits to start the memo writing process'
			Emreadscreen mf_wcom_exists, 3, 3, 15
			If mf_wcom_exists <> "   " then script_end_procedure ("It appears you already have a WCOM added to this notice. The script will now end.")
			If program_type = "MF" AND print_status = "Waiting" then
				mf_wcom_writen = true
				'This will write if it is for an MFIP notice
				CALL write_variable_in_SPEC_MEMO("******************************************************")
				CALL write_variable_in_SPEC_MEMO("Dear Client,")
				CALL write_variable_in_SPEC_MEMO("")
				CALL write_variable_in_SPEC_MEMO("One or more of your household members will not be eligible for SNAP benefits in this month since you have received SNAP benefits on another case for the same month.")
				CALL write_variable_in_SPEC_MEMO("Per program rules SNAP participants are not eligible for duplicate benefits in the same benefit month.")
				CALL write_variable_in_SPEC_MEMO("")
				CALL write_variable_in_SPEC_MEMO("If you have any questions or concerns please feel free to contact your worker.")
				CALL write_variable_in_SPEC_MEMO("---")
				CALL write_variable_in_SPEC_MEMO(worker_signature)
				CALL write_variable_in_SPEC_MEMO("")
				CALL write_variable_in_SPEC_MEMO("******************************************************")
				PF4
				PF3
			End If
		End If
	End If
	If mf_wcom_writen = true then Exit Do
	If wcom_row = 17 then
		PF8
		Emreadscreen spec_edit_check, 6, 24, 2
		wcom_row = 6
	end if
	If spec_edit_check = "NOTICE" THEN no_mf_waiting = true
Loop until spec_edit_check = "NOTICE"

If no_fs_waiting = true AND no_mf_waiting = true then script_end_procedure("No waiting FS or MFIP notices were found for the requested month")

script_end_procedure("WCOM has been added to the first found waiting SNAP and/or MFIP notice for the month and case selected. Please review the notice.")
