'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - CLAIM REFERRAL TRACKING.vbs" 'BULK script that creates a list of cases that require an interview, and the contact phone numbers'
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 200            'manual run time in seconds
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
call changelog_update("08/05/2019", "Updated the term claim referral to use the action taken on MISC as well as to read for active programs.", "MiKayla Handley")
call changelog_update("07/30/2019", "Reverted the term claim referral to use the action taken on MISC as well as to read for active programs.", "MiKayla Handley")
call changelog_update("10/15/2018", "Updated claim referral dialog to read for active programs.", "MiKayla Handley")
call changelog_update("09/20/2018", "Updated claim referral dialog to match MAXIS panel.", "MiKayla Handley")
call changelog_update("06/26/2017", "Initial version.", "MiKayla Handley")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'----------------------------------------------------------------------------------------------------Thescript
EMCONNECT ""
Call MAXIS_case_number_finder(MAXIS_case_number)
'MEMB_number = "01"
action_date = date & ""
'verif_requested = "TEST"
'other_notes = "TEST"

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 221, 155, "Claim Referral Tracking"
  EditBox 65, 30, 45, 15, MAXIS_case_number
  EditBox 165, 30, 45, 15, action_date
  DropListBox 65, 50, 145, 15, "Select One:"+chr(9)+"Sent Request for Additional Info"+chr(9)+"Overpayment Exists"+chr(9)+"No Overpayment Exists", action_taken
  EditBox 65, 70, 145, 15, verif_requested
  EditBox 65, 90, 145, 15, other_notes
  EditBox 110, 110, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 115, 135, 45, 15
    CancelButton 165, 135, 45, 15
    PushButton 5, 135, 85, 15, "Claims Procedures", claims_procedures
  Text 5, 5, 210, 20, "This script will only enter a STAT/MISC panel for a SNAP or MFIP federal food claim. "
  Text 15, 35, 50, 10, "Case Number: "
  Text 120, 35, 40, 10, "Action Date: "
  Text 15, 55, 45, 10, "Action Taken:"
  Text 5, 75, 55, 10, "Verif Requested:"
  Text 20, 95, 45, 10, "Other Notes:"
  Text 65, 115, 40, 10, "Worker Sig:"
EndDialog

Do
	Do
		err_msg = ""
		Do
            dialog Dialog1
            cancel_confirmation
            If ButtonPressed = claims_procedures then CreateObject("WScript.Shell").Run("https://dept.hennepin.us/hsphd/manuals/hsrm/Pages/Claims_Maxis_Procedures.aspx")
        Loop until ButtonPressed = -1
		IF buttonpressed = 0 then stopscript
		IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
		IF isdate(action_date) = False then err_msg = err_msg & vbnewline & "* Enter a valid action date."
		IF action_taken = "Select One:" then err_msg = err_msg & vbnewline & "* Select the action taken for next step in overpayment."
        IF action_taken = "Sent Request for Additional Info" and verif_requested = "" then err_msg = err_msg & vbnewline & "* You selected that a request for additional information was sent, please advise what verifications were requested."
		IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "* Enter your worker signature."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Determines which programs are currently status_checking in the month of application
CALL navigate_to_MAXIS_screen("STAT", "PROG")		'Goes to STAT/PROG

DW_STATUS = FALSE 'Diversionary Work Program'
GA_STATUS = FALSE 'General Assistance'
MS_STATUS = FALSE 'Mn Suppl Aid '
MF_STATUS = FALSE 'Mn Family Invest Program '
RC_STATUS = FALSE 'Refugee Cash Assistance'
SNAP_STATUS = FALSE
CASH_STATUS = FALSE
'Reading the status and program
EMReadScreen cash1_status_check, 4, 6, 74
EMReadScreen cash2_status_check, 4, 7, 74
EMReadScreen emer_status_check, 4, 8, 74
EMReadScreen grh_status_check, 4, 9, 74
EMReadScreen fs_status_check, 4, 10, 74
EMReadScreen ive_status_check, 4, 11, 74
EMReadScreen hc_status_check, 4, 12, 74
EMReadScreen cca_status_check, 4, 14, 74

EMReadScreen cash1_prog_check, 2, 6, 67
EMReadScreen cash2_prog_check, 2, 7, 67
EMReadScreen emer_prog_check, 2, 8, 67
EMReadScreen grh_prog_check, 2, 9, 67
EMReadScreen fs_prog_check, 2, 10, 67
EMReadScreen ive_prog_check, 2, 11, 67
EMReadScreen hc_prog_check, 2, 12, 67
EMReadScreen cca_prog_check, 2, 14, 67

IF FS_status_check = "ACTV" or FS_status_check = "PEND"  Then SNAP_STATUS = TRUE

'Logic to determine if MFIP is active
If cash1_prog_check = "MF" THEN
	If cash1_status_check = "ACTV" Then CASH_STATUS = TRUE
	If cash1_status_check = "PEND" Then CASH_STATUS = TRUE
	If cash1_status_check = "INAC" Then CASH_STATUS = FALSE
	If cash1_status_check = "SUSP" Then CASH_STATUS = FALSE
	If cash1_status_check = "DENY" Then CASH_STATUS = FALSE
	If cash1_status_check = ""     Then CASH_STATUS = FALSE
END IF
If cash2_prog_check = "MF" THEN
	If cash2_status_check = "ACTV" THEN CASH_STATUS = TRUE
	If cash2_status_check = "PEND" THEN CASH_STATUS = TRUE
	If cash2_status_check = "INAC" THEN CASH_STATUS = FALSE
	If cash2_status_check = "SUSP" THEN CASH_STATUS = FALSE
	If cash2_status_check = "DENY" THEN CASH_STATUS = FALSE
	If cash2_status_check = ""     THEN CASH_STATUS = FALSE
END IF
programs = ""
IF SNAP_STATUS = TRUE THEN programs = programs & "Food Support, "
IF CASH_STATUS = TRUE THEN programs = programs & "MFIP, "

programs = trim(programs)
'takes the last comma off of programs when autofilled into dialog
If right(programs, 1) = "," THEN programs = left(programs, len(programs) - 1)


IF SNAP_STATUS = TRUE or CASH_STATUS = TRUE THEN
    'MsgBox programs
    Call navigate_to_MAXIS_screen ("STAT", "MISC")
	Row = 6
	EmReadScreen panel_number, 1, 02, 73
	If panel_number = "0" then
		EMWriteScreen "NN", 20,79
		TRANSMIT
		'CHECKING FOR MAXIS PROGRAMS ARE INACTIVE'
		EmReadScreen MISC_error_check,  74, 24, 02
		IF trim(MISC_error_check) = "" THEN
			case_note_only = FALSE
		else
			maxis_error_check = MsgBox("*** NOTICE!!!***" & vbNewLine & "Continue to case note only?" & vbNewLine & MISC_error_check & vbNewLine, vbYesNo + vbQuestion, "Message handling")
			IF maxis_error_check = vbYes THEN
				case_note_only = TRUE 'this will case note only'
			END IF
			IF maxis_error_check= vbNo THEN
				case_note_only = FALSE 'this will update the panels and case note'
			END IF
		END IF
	END IF

	Do
		'Checking to see if the MISC panel is empty, if not it will find a new line'
		EmReadScreen MISC_description, 25, row, 30
		MISC_description = replace(MISC_description, "_", "")
		If trim(MISC_description) = "" then
			'PF9
			EXIT DO
		Else
			row = row + 1
		End if
	Loop Until row = 17
	If row = 17 then MsgBox("There is not a blank field in the MISC panel. Please delete a line(s), and run script again or update manually.")

	'writing in the action taken and date to the MISC panel
	PF9
    'writing in the action taken and date to the MISC panel
    IF action_taken = "Sent Request for Additional Info" THEN MISC_action_taken = "Initial Claim Referral"
    IF action_taken = "Overpayment Exists" THEN MISC_action_taken =  "Determination-OP Entered" '"Claim Determination 25 character available
    IF action_taken = "No Overpayment Exists" THEN MISC_action_taken = "Determination-No Savings"
    EMWriteScreen MISC_action_taken, Row, 30
    EMWriteScreen date, Row, 66
    TRANSMIT
    'PF3

    due_date = dateadd("d", 10, date)
    'The case note-------------------------------------------------------------------------------------------------
    start_a_blank_CASE_NOTE
    Call write_variable_in_case_note("***Claim Referral Tracking-" & action_taken & "***")
    Call write_bullet_and_variable_in_case_note("Action Date", action_date)
    Call write_bullet_and_variable_in_case_note("Active Program(s)", programs)
    IF action_taken = "Sent Request for Additional Info" THEN Call write_bullet_and_variable_in_case_note("Action taken", MISC_action_taken)
    IF action_taken = "Sent Request for Additional Info" THEN CALL write_variable_in_case_note("* Additional verifications requested, TIKL set for 10 day return.")
    If action_taken = "Sent Request for Additional Info" THEN  Call write_bullet_and_variable_in_case_note("Verification requested", verif_requested)
    If action_taken = "Overpayment Exists" THEN Call write_variable_in_case_note("* Overpayment exists, claims procedure to follow.")
    IF action_taken = "No Overpayment Exists" THEN Call write_variable_in_case_note("* No overpayment exists, maxis has been updated with reported changes.")
    Call write_bullet_and_variable_in_case_note("Other Notes", other_notes)
    Call write_variable_in_case_note("* Entries for these potential claims must be retained until further notice.")
    Call write_variable_in_case_note("---")
    Call write_variable_in_case_note(worker_signature)
    PF3

    IF action_taken = "Sent Request for Additional Info" THEN
    'set TIKL------------------------------------------------------------------------------------------------------
        Call navigate_to_MAXIS_screen("DAIL", "WRIT")
        call create_MAXIS_friendly_date(due_date, 10, 5, 18)
        Call write_variable_in_TIKL("Potential overpayment exists on case. Please review case for receipt of additional requested information.")
        PF3
    	script_end_procedure("You have indicated that you sent a request for additional information. Please follow the agency's procedure(s) for claim entry once received.")
    Else
    	script_end_procedure("You have indicated that an overpayment exists. Please follow the agency's procedure(s) for claim entry.")
    End if
ELSE
    IF SNAP_STATUS = FALSE and CASH_STATUS = FALSE THEN
        Do
            case_note_only_confirmation = MsgBox(" " & vbNewLine & "*** Only enter a STAT/MISC panel for a SNAP or MFIP federal food claim. Please press no if SNAP or MFIP were not open.***", vbYesNo, "Do you wish to case note only?")
            IF case_note_only_confirmation = vbNo THEN
                case_note_only = FALSE
                EXIT DO
            END IF
            IF case_note_only_confirmation = vbYes THEN
                case_note_only = TRUE
                EXIT DO
            END IF
        LOOP
    END IF
    IF case_note_only = TRUE THEN
         start_a_blank_CASE_NOTE
         Call write_variable_in_case_note("***Claim Referral Tracking-" & action_taken & "***")
         Call write_bullet_and_variable_in_case_note("Action Date", action_date)
         Call write_bullet_and_variable_in_case_note("Active Program(s)", programs)
         IF action_taken = "Sent Request for Additional Info" THEN Call write_bullet_and_variable_in_case_note("Action taken", MISC_action_taken)
         IF action_taken = "Sent Request for Additional Info" THEN CALL write_variable_in_case_note("* Additional verifications requested, TIKL set for 10 day return.")
         If action_taken = "Sent Request for Additional Info" THEN  Call write_bullet_and_variable_in_case_note("Verification requested", verif_requested)
         If action_taken = "Overpayment Exists" THEN Call write_variable_in_case_note("* Overpayment exists, claims procedure to follow.")
         IF action_taken = "No Overpayment Exists" THEN Call write_variable_in_case_note("* No overpayment exists, maxis has been updated with reported changes.")
         Call write_bullet_and_variable_in_case_note("Other Notes", other_notes)
         Call write_variable_in_case_note("* Entries for these potential claims must be retained until further notice. Case is currently inactive and the MISC panel was not entered.")
         Call write_variable_in_case_note("---")
         Call write_variable_in_case_note(worker_signature)
         PF3
    ELSE
        script_end_procedure_with_error_report("Claim Referral Tracking is for MFIP and SNAP cases only. Please let us know if there are further considerations needed.")
    END IF
END IF
