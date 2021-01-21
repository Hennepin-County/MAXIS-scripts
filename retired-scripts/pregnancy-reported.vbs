'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - PREGNANCY REPORTED.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 180          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT------------------------------------------------------------------------------------------------------
'Connects to BLUEZONE
EMConnect ""
'Grabs the MAXIS case number
CALL MAXIS_case_number_finder(MAXIS_case_number)
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 296, 140, "Pregnancy Reported"
  EditBox 60, 5, 55, 15, maxis_case_number
  EditBox 165, 5, 20, 15, MEMB_number
  CheckBox 200, 20, 20, 10, "CA", cash_checkbox
  CheckBox 230, 20, 25, 10, "FS", FS_CHECKBOX
  CheckBox 260, 20, 25, 10, "MA", ma_checkbox
  CheckBox 200, 40, 70, 10, "Updated in MMIS", mmis_checkbox
  EditBox 60, 25, 55, 15, due_date
  DropListBox 60, 45, 95, 15, "Select One:"+chr(9)+"Self Attestation"+chr(9)+"Change Report Form"+chr(9)+"Pregnancy Verification Form"+chr(9)+"Renewal Form"+chr(9)+"Other", report_method
  EditBox 60, 65, 230, 15, verif_requested
  CheckBox 60, 85, 125, 10, "Verification Request sent for CASH", verification_checkbox
  EditBox 60, 100, 230, 15, other_notes
  EditBox 60, 120, 125, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 195, 120, 45, 15
    CancelButton 245, 120, 45, 15
  Text 120, 10, 45, 10, "HH MEMB#:"
  Text 5, 10, 50, 10, "Case Number:"
  Text 5, 50, 45, 10, "Reported Via:"
  GroupBox 195, 5, 95, 30, "Active Programs:"
  Text 5, 125, 40, 10, "Worker Sig:"
  Text 5, 30, 35, 10, "Due Date:"
  Text 5, 105, 45, 10, "Other Notes:"
  Text 5, 70, 55, 10, "Verif Requested:"
EndDialog

'Shows dialog
DO
    DO
    	err_msg = ""
    	Dialog Dialog1				'Calling a dialog without a assigned variable will call the most recently defined dialog
    	cancel_confirmation
    	IF report_method = "Select One:" THEN err_msg = err_msg & vbCr & "* You must select how the pregnancy was reported!"
    	IF IsNumeric(MAXIS_case_number) = FALSE THEN err_msg = err_msg & vbCr & "* You must type a valid numeric case number."
    	IF due_date = "" OR (due_date <> "" AND IsDate(due_date) = False) THEN err_msg = err_msg & vbCr & "* You must enter a due date in a MM/DD/YY format."
    	IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* You must sign your case note!"
    	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
    LOOP UNTIL err_msg = ""
LOOP UNTIL are_we_passworded_out = false

'Checks Maxis for password prompt
CALL check_for_MAXIS(True)

'Script calculates the Conception date based off the due date entered in the dialog box
conception_date = DateAdd("d", -280, due_date)

'The script reads what member number was manually entered, and navigates to that member's stat/preg panel
CALL navigate_to_MAXIS_screen("STAT", "PREG")
EMWriteScreen MEMB_number, 20, 76
EMWriteScreen "nn", 20, 79
transmit

'Writes the auto-calucated conception date in the Conception Date field and the Due date in that field
CALL create_MAXIS_friendly_date(conception_date, 0, 6, 53)
CALL create_MAXIS_friendly_date(due_date, 0, 10, 53)

EMWriteScreen "n", 8, 75

'If under Program Pregnancy applied for, FW has check MA or MA/CASH then script will write Y in the Verified field on stat/preg
IF ma_checkbox = checked and cash_checkbox = checked THEN EMWritescreen "Y", 6, 75

'If under Program Pregnancy applied for, FW has checked CASH then script will write N in the Verified field on stat/preg
IF cash_checkbox = checked THEN EMWritescreen "N", 6, 75
TRANSMIT 'to save the updates'
'formatting for casenote'
IF cash_checkbox = CHECKED THEN programs = "CA,"
IF FS_CHECKBOX = CHECKED THEN programs = "FS,"
IF ma_checkbox = CHECKED THEN programs = "MA,"
'trims excess spaces of programs
programs = trim(programs)
'takes the last comma off of programs when autofilled into dialog
IF right(programs, 1) = "," THEN programs = left(programs, len(programs) - 1)
'Opens new case note
start_a_blank_case_note
CALL write_variable_in_case_note ("---Pregnancy Reported---")
CALL write_bullet_and_variable_in_case_note("Household Member", MEMB_number)
CALL write_bullet_and_variable_in_case_note("Conception Date", conception_date)
CALL write_bullet_and_variable_in_case_note("Pregnancy Due Date", due_date)
CALL write_bullet_and_variable_in_case_note("Pregnancy Reported Via", report_method)
CALL write_variable_in_CASE_NOTE("Active Programs", programs)
IF mmis_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Updated in MMIS")
IF verification_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Sent verification request for CASH")
CALL write_bullet_and_variable_in_CASE_NOTE("Verifications Requested", verif_requested)
CALL write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

script_end_procedure("")
