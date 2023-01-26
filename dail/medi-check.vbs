'Required for statistical purposes===============================================================================
name_of_script = "DAIL - MEDI CHECK.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 127         'manual run time in seconds
STATS_denomination = "C"       'C is for case
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

'===========================================================================================================CHANGELOG BLOCK
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("01/26/2023", "Removed term 'ECF' from the case note per DHS guidance, and referencing the case file instead.", "Ilse Ferris, Hennepin County")
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note.", "Ilse Ferris")
call changelog_update("05/01/2019", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'=======================================================================================================END CHANGELOG BLOCK
'THE MAIN DIALOG--------------------------------------------------------------------------------------------------
EMConnect ""

EMWriteScreen "N", 6, 3         'Goes to Case Note - maintains tie with DAIL
TRANSMIT
Call MAXIS_case_number_finder(MAXIS_case_number)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 130, DAIL_type & " MESSAGE PROCESSED"
  EditBox 60, 35, 15, 15, memb_number
  CheckBox 185, 35, 75, 10, "Referral sent in ECF", ECF_sent_checkbox
  CheckBox 5, 55, 140, 10, "Client is eligible for the Medicare buy-in", medi_checkbox
  EditBox 210, 50, 50, 15, ELIG_date
  EditBox 210, 70, 50, 15, ELIG_year
  EditBox 50, 90, 210, 15, other_notes
  EditBox 65, 110, 105, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 175, 110, 40, 15
    CancelButton 220, 110, 40, 15
  GroupBox 5, 5, 255, 25, "DAIL for case #"  &  MAXIS_case_number
  Text 10, 15, 250, 10, full_message
  Text 5, 95, 40, 10, "Other notes:"
  Text 5, 115, 60, 10, "Worker signature:"
  Text 170, 55, 35, 10, "ELIG date"
  Text 5, 75, 195, 10, "If INELIG year that client will be eligible for Medicare Buy-In"
  Text 5, 40, 50, 10, "Memb number:"
EndDialog

Do
    Do
        err_msg = ""
		Dialog Dialog1
		cancel_confirmation

		ELIG_year = trim(ELIG_year)
		IF medi_checkbox = CHECKED THEN
			IF isdate(ELIG_date) = False then err_msg = err_msg & vbnewline & "* Since you indicated the client is eligible for the Medicare Buy-In, enter a valid date of eligibility."
		ELSE
			IF ELIG_year = "" THEN
				err_msg = err_msg & vbnewline & "* Since you did not check the box to indicate the client is eligible for the Medicare Buy-in, you must indicate the year the client will be expected to be eligible."
			ELSE
				If len(ELIG_year) <> 2 Then err_msg = err_msg & vbnewline & "* Enter just the last 2 digits of the year - the script will enter the '20' at the begninning."
			END IF
		END IF

        If (isnumeric(memb_number) = False OR len(memb_number) > 2) then err_msg = err_msg & vbcr & "* Enter a valid member number."
		If trim(worker_signature) = "" then err_msg = err_msg & vbcr & "* Please ensure your case note is signed."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'checking for an active MAXIS session
Call check_for_MAXIS(False)
EmReadScreen panel_check, 04, 02, 45
DO
	'before we write checking for casenote'
	IF panel_check =  "NOTE" THEN EXIT DO
	IF panel_check <> "NOTE" THEN
		case_note_confirmation = MsgBox("Press YES to confirm that you are back to case notes and ready write." & vbNewLine & "To navigate to CASE/NOTE, press NO." & vbNewLine & vbNewLine & _
    	"Panel Check -" & panel_check, vbYesNoCancel, "Case note confirmation")
	 	IF case_note_confirmation = vbNo THEN
			Call navigate_to_MAXIS_screen("CASE", "NOTE")
			PF9
			EmReadScreen write_mode_casenote, 06, 03, 03
			If write_mode_casenote <> "Please" THEN msgbox script_end_procedure_with_error_report("The script has ended. Unable to access case note.")
			IF write_mode_casenote = "Please" THEN EXIT DO
		END IF
		IF case_note_confirmation = vbYes THEN EXIT DO
		IF case_note_confirmation = vbCancel THEN script_end_procedure_with_error_report("The script has ended. The DAIL has not been acted on.")
	END IF
LOOP UNTIL case_note_confirmation = vbYes


'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
due_date = dateadd("d", 30, date)
IF medi_checkbox = CHECKED and ELIG_date <> "" THEN Call create_TIKL("Referral made for medicare, please check on proof of application filed. Due " & due_date & ".", 30, date, True, TIKL_note_text)

IF IsNumeric(ELIG_year) = TRUE THEN
	reminder_year = ELIG_year - 1
    nov_date = "11/01/" & reminder_year
    'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
    Call create_TIKL("Reminder to mail the Medicare Referral for January 20" & ELIG_year & ".", 0, nov_date, False, TIKL_note_text)
END IF

'----------------------------------------------------------------------------the casenote
Call navigate_to_MAXIS_screen("CASE", "NOTE")
PF9
EMReadScreen case_note_mode_check, 7, 20, 3
If case_note_mode_check <> "Mode: A" then script_end_procedure("You are not in a case note on edit mode. You might be in inquiry. Try the script again in production.")

IF medi_checkbox = CHECKED and ELIG_date <> "" THEN
	Call write_variable_in_case_note("** Medicare Buy-in Referral mailed for M" & memb_number & " **")
	Call write_variable_in_case_note("* Client is eligible for the Medicare buy-in as of " & ELIG_date & ".")
	Call write_variable_in_case_note("* Proof due by " & due_date & " to apply.")
	Call write_variable_in_case_note("* Mailed DHS-3439-ENG MHCP Medicare Buy-In Referral Letter")
	Call write_variable_in_case_note("* TIKL set to follow up.")
ELSEIF ELIG_year <> "" THEN
	Call write_variable_in_case_note("** Medicare Referral for M" & memb_number & " **")
	Call write_variable_in_case_note("* Client is not eligible for the Medicare buy-in. Enrollment is not until January 20" & ELIG_year & ", unable to apply until the enrollment time.")
	Call write_variable_in_case_note("* TIKL set to mail the Medicare Referral for November 20" & reminder_year & ".")
END IF
IF ECF_sent_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Case file reviewed.")
CALL write_bullet_and_variable_in_case_note("Other notes", other_notes)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure_with_error_report("DAIL has been case noted. Please remember to send forms out of ECF and delete the PEPR.")
