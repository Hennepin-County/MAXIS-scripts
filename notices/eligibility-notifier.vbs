'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - ELIGIBILITY NOTIFIER.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 195                               'manual run time in seconds
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

'DIALOGS----------------------------------------------------------------------------------------------------

BeginDialog Potential_Eligibility_MEMO_dialog, 0, 0, 181, 120, "Potential Eligibility MEMO"
  EditBox 75, 5, 50, 15, MAXIS_case_number
  CheckBox 10, 30, 30, 10, "SNAP", SNAP_checkbox
  CheckBox 55, 30, 30, 10, "CASH", CASH_checkbox
  CheckBox 100, 30, 25, 10, "MA", MA_checkbox
  CheckBox 140, 30, 30, 10, "MSP", MSP_checkbox
  DropListBox 100, 55, 75, 10, ""+chr(9)+"Apply in MAXIS"+chr(9)+"Apply in MNSure", HC_apply_method
  EditBox 90, 75, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 30, 95, 50, 15
    CancelButton 90, 95, 50, 15
  Text 25, 10, 50, 10, "Case number:"
  Text 20, 80, 65, 10, "Worker signature:"
  Text 10, 50, 85, 20, "If HC was checked please pick system to apply in:"
EndDialog
'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

'Searches for a case number
call MAXIS_case_number_finder(MAXIS_case_number)

'This Do...loop shows the appointment letter dialog, and contains logic to require most fields.
DO
	Do
		err_msg = ""
		Dialog Potential_Eligibility_MEMO_dialog
		If ButtonPressed = cancel then stopscript
		If SNAP_checkbox <> checked AND CASH_checkbox <> checked AND MA_checkbox <> checked AND MSP_checkbox <> checked THEN err_msg = err_msg & "Please select a program." & vbNewLine
		If MSP_checkbox = checked AND HC_apply_method <> "Apply in MAXIS" THEN err_msg = err_msg & "You selected MSP, at this time you cannot apply in Mnsure if you have Medicare. Please review selections" & vbNewLine
		If (MSP_checkbox = checked or MA_checkbox = checked) AND HC_apply_method = "" THEN err_msg = err_msg & "You selected a HC program, please select a system to apply in." & vbNewLine
		If isnumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & "You must fill in a valid case number." & vbNewLine
		If worker_signature = "" then err_msg = err_msg & "You must sign your case note." & vbNewLine
		IF err_msg <> "" THEN msgbox err_msg
	Loop until err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'Checks for MAXIS
call check_for_MAXIS(False)

'Navigating to SPEC/MEMO
call navigate_to_MAXIS_screen("SPEC", "MEMO")

'Creates a new MEMO. If it's unable the script will stop.
PF5
EMReadScreen memo_display_check, 12, 2, 33
If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")

'Checking for an AREP. If there's an AREP it'll navigate to STAT/AREP, check to see if the forms go to the AREP. If they do, it'll write X's in those fields below.
row = 4                             'Defining row and col for the search feature.
col = 1
EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
IF row > 4 THEN                     'If it isn't 4, that means it was found.
    arep_row = row                                          'Logs the row it found the ALTREP string as arep_row
    call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
    EMReadscreen forms_to_arep, 1, 10, 45                   'Reads for the "Forms to AREP?" Y/N response on the panel.
    call navigate_to_MAXIS_screen("SPEC", "MEMO")           'Navigates back to SPEC/MEMO
    PF5                                                     'PF5s again to initiate the new memo process
END IF
'Checking for SWKR
row = 4                             'Defining row and col for the search feature.
col = 1
EMSearch "SOCWKR", row, col         'Row and col are variables which change from their above declarations if "SOCWKR" string is found.
IF row > 4 THEN                     'If it isn't 4, that means it was found.
    swkr_row = row                                          'Logs the row it found the SOCWKR string as swkr_row
    call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
    EMReadscreen forms_to_swkr, 1, 15, 63                'Reads for the "Forms to SWKR?" Y/N response on the panel.
    call navigate_to_MAXIS_screen("SPEC", "MEMO")         'Navigates back to SPEC/MEMO
    PF5                                           'PF5s again to initiate the new memo process
END IF
EMWriteScreen "x", 5, 10                                        'Initiates new memo to client
IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
transmit                                                        'Transmits to start the memo writing process


'formatting variable
IF SNAP_checkbox = checked THEN progs_to_apply_in_maxis = "SNAP or "
IF CASH_checkbox = checked THEN progs_to_apply_in_maxis = progs_to_apply_in_maxis & "CASH or "
IF MA_checkbox = checked AND HC_apply_method = "Apply in MAXIS" THEN progs_to_apply_in_maxis = progs_to_apply_in_maxis & "MA(in MAXIS) or "
IF MA_checkbox = checked AND HC_apply_method = "Apply in MNSure" THEN progs_to_apply_in_maxis = progs_to_apply_in_maxis & "MA(in MNSure) or "
IF MSP_checkbox = checked THEN progs_to_apply_in_maxis = progs_to_apply_in_maxis & "MSP or "
progs_to_apply_in_maxis = left(progs_to_apply_in_maxis,(len(progs_to_apply_in_maxis) - "3"))


'Writes the MEMO.
call write_variable_in_SPEC_MEMO("***********************************************************")
IF SNAP_checkbox = checked THEN call write_variable_in_SPEC_MEMO("You appear to be eligible for the Supplemental Nutritional Assistance Program (SNAP).")
IF CASH_checkbox = checked THEN call write_variable_in_SPEC_MEMO("You appear to be eligible for CASH assistance.")
IF MA_checkbox = checked THEN call write_variable_in_SPEC_MEMO("You appear to be eligible for Medical assistance(MA).")
IF MSP_checkbox = checked THEN call write_variable_in_SPEC_MEMO("You appear to be eligible for the Medicare Savings Program (MSP).")
call write_variable_in_SPEC_MEMO("")
IF SNAP_checkbox = checked or CASH_checkbox = checked or HC_apply_method = "Apply in MAXIS" THEN call write_variable_in_SPEC_MEMO("To apply for " & progs_to_apply_in_maxis & "apply online at applymn.org, contact your worker to request an application, or complete an application at your local County or Tribal agency.")
call write_variable_in_SPEC_MEMO("")
IF SNAP_checkbox = checked or CASH_checkbox = checked THEN
	call write_variable_in_SPEC_MEMO("When applying for SNAP and/or CASH you can submit the first page of the paper application to set your date of application. Your first month's benefit will be determined based on your application date.")
	call write_variable_in_SPEC_MEMO("")
END IF
IF HC_apply_method = "Apply in MNSure" THEN
	call write_variable_in_SPEC_MEMO("To apply for MA you can apply online at MNsure.org, you can contact your worker to request an application, or complete an application at your local County or Tribal Agency.")
	call write_variable_in_SPEC_MEMO("")
END IF
call write_variable_in_SPEC_MEMO("This is a notice to inform you of programs you might have eligibility for. Elibility will be determined during the application process.")
call write_variable_in_SPEC_MEMO("***********************************************************")
'Exits the MEMO
PF4

'Navigates to CASE/NOTE and starts a blank one
start_a_blank_CASE_NOTE

'Writes the case note--------------------------------------------
call write_variable_in_CASE_NOTE("**Potential Eligibility Notice Sent**")
call write_bullet_and_variable_in_CASE_NOTE("Programs client may be eligible for", progs_to_apply_in_maxis)
If forms_to_arep = "Y" then call write_variable_in_CASE_NOTE("* Copy of notice sent to AREP.")              'Defined above
If forms_to_swkr = "Y" then call write_variable_in_CASE_NOTE("* Copy of notice sent to Social Worker.")     'Defined above
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")
