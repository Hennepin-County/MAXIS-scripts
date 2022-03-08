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
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note.", "Ilse Ferris")
call changelog_update("08/05/2019", "Updated the term claim referral to use the action taken on MISC as well as to read for active programs.", "MiKayla Handley")
call changelog_update("07/30/2019", "Reverted the term claim referral to use the action taken on MISC as well as to read for active programs.", "MiKayla Handley")
call changelog_update("10/15/2018", "Updated claim referral dialog to read for active programs.", "MiKayla Handley")
call changelog_update("09/20/2018", "Updated claim referral dialog to match MAXIS panel.", "MiKayla Handley")
call changelog_update("06/26/2017", "Initial version.", "MiKayla Handley")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'---------------------------------------------------------------------------------------The script
EMConnect ""                                        'Connecting to BlueZone
CALL MAXIS_case_number_finder(MAXIS_case_number)    'Grabbing the CASE Number
CALL Check_for_MAXIS(false)                         'Ensuring we are not passworded out
action_date = date & ""

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
    PushButton 5, 135, 85, 15, "Claims Procedures", claims_procedures_btn
  Text 5, 5, 210, 20, "This script will only enter a STAT/MISC panel for a SNAP or MFIP federal food claim. "
  Text 15, 35, 50, 10, "Case Number: "
  Text 120, 35, 40, 10, "Action Date: "
  Text 15, 55, 45, 10, "Action Taken:"
  Text 5, 75, 55, 10, "Verif Requested:"
  Text 20, 95, 45, 10, "Other Notes:"
  Text 65, 115, 40, 10, "Worker Sig:"
EndDialog

DO
    DO
	    err_msg = ""
	    DO
            dialog Dialog1
            cancel_without_confirmation
            If ButtonPressed = claims_procedures_btn then CreateObject("WScript.Shell").Run("https://hennepin.sharepoint.com/teams/hs-es-manual/sitepages/Claims_and_Overpayments.aspx")
        Loop until ButtonPressed = -1
	    IF buttonpressed = 0 then stopscript
	    IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
	    IF isdate(action_date) = False then err_msg = err_msg & vbnewline & "* Please enter a valid action date."
	    IF action_taken = "Select One:" then err_msg = err_msg & vbnewline & "* Please select the action taken for next step in overpayment."
        IF action_taken = "Sent Request for Additional Info" and verif_requested = "" then err_msg = err_msg & vbnewline & "* You selected that a request for additional information was sent, please advise what verifications were requested."
	    IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "* Please enter your worker signature."
	    IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
    LOOP UNTIL err_msg = ""									'loops until all errors are resolved
CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

'Checking for PRIV cases.
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "SUMM", is_this_priv)
IF is_this_priv = True THEN script_end_procedure_with_error_report("This case is privileged. Please request access before running the script again. ")
MAXIS_background_check      'Making sure we are out of background.

'Grabbing case and program status information from MAXIS.
'For tis script to work correctly, these must be correct BEFORE running the script.
Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status)
EMReadScreen case_status, 15, 8, 9                  'Now we are reading the CASE STATUS string from the panel - we want to make sure this does NOT read CAF1 PENDING
EMReadScreen appl_date, 8, 8, 29               'Grabbing the PND2 date from CASE CURR in case the information cannot be pulled from REPT/PND2
ega_status = "INACTIVE"
ea_status = "INACTIVE"

'\This functionality is how the above function reads for program information - just pulled out for these specific programs

row = 1                                             'Looking for EGA information
col = 1
EMSearch "EGA", row, col
If row <> 0 Then
    EMReadScreen ega_status, 9, row, col + 6
    ega_status = trim(ega_status)
    If ega_status = "ACTIVE" or ega_status = "APP CLOSE" or ega_status = "APP OPEN" Then ega_status = "ACTIVE"
    If ega_status = "PENDING" Then case_pending = True      'Updating the case_pending variable from the function
End If

row = 1                                             'Looking for EA information
col = 1
EMSearch "EA: ", row, col
If row <> 0 Then
    EMReadScreen ea_status, 9, row, col + 5
    ea_status = trim(ea_status)
    If ea_status = "ACTIVE" or ea_status = "APP CLOSE" or ea_status = "APP OPEN" Then ea_status = "ACTIVE"
    If ea_status = "PENDING" Then case_pending = True      'Updating the case_pending variable from the function
End If

case_status = trim(case_status)     'cutting off any excess space from the case_status read from CASE/CURR above
active_programs = ""        'Creates a variable that lists all the active programs on the case.
If ga_status = "ACTIVE" Then active_programs = active_programs & "GA, "
If msa_status = "ACTIVE" Then active_programs = active_programs & "MSA, "
If mfip_status = "ACTIVE" Then active_programs = active_programs & "MFIP, "
If dwp_status = "ACTIVE" Then active_programs = active_programs & "DWP, "
If ive_status = "ACTIVE" Then active_programs = active_programs & "IV-E, "
If grh_status = "ACTIVE" Then active_programs = active_programs & "GRH, "
If snap_status = "ACTIVE" Then active_programs = active_programs & "SNAP, "
If ega_status = "ACTIVE" Then active_programs = active_programs & "EGA, "
If ea_status = "ACTIVE" Then active_programs = active_programs & "EA, "
If cca_status = "ACTIVE" Then active_programs = active_programs & "CCA, "
If ma_status = "ACTIVE" OR msp_status = "ACTIVE" Then active_programs = active_programs & "HC, "

active_programs = trim(active_programs)  'trims excess spaces of active_programs
If right(active_programs, 1) = "," THEN active_programs = left(active_programs, len(active_programs) - 1)

pending_programs = ""        'Creates a variable that lists all the pending programs on the case.
If unknown_cash_pending = True Then pending_programs = pending_programs & "Cash, "
If ga_status = "PENDING" Then pending_programs = pending_programs & "GA, "
If msa_status = "PENDING" Then pending_programs = pending_programs & "MSA, "
If mfip_status = "PENDING" Then pending_programs = pending_programs & "MFIP, "
If dwp_status = "PENDING" Then pending_programs = pending_programs & "DWP, "
If ive_status = "PENDING" Then pending_programs = pending_programs & "IV-E, "
If grh_status = "PENDING" Then pending_programs = pending_programs & "GRH, "
If snap_status = "PENDING" Then pending_programs = pending_programs & "SNAP, "
If ega_status = "PENDING" Then pending_programs = pending_programs & "EGA, "
If ea_status = "PENDING" Then pending_programs = pending_programs & "EA, "
If cca_status = "PENDING" Then pending_programs = pending_programs & "CCA, "
If ma_status = "PENDING" OR msp_status = "PENDING" OR unknown_hc_pending = True Then pending_programs = pending_programs & "HC, "

pending_programs = trim(pending_programs)  'trims excess spaces of pending_programs
If right(pending_programs, 1) = "," THEN pending_programs = left(pending_programs, len(pending_programs) - 1)

Call back_to_SELF

CALL navigate_to_MAXIS_screen("STAT", "PROG")		'Goes to STAT/PROG
CASH_STATUS = FALSE 'overall variable'
'CCA_STATUS = FALSE
DW_STATUS = FALSE 'Diversionary Work Program'
EMER_STATUS = FALSE 'Emergency Assistance'
FS_STATUS = FALSE
GA_STATUS = FALSE 'General Assistance'
GRH_STATUS = FALSE
'HC_STATUS = FALSE
MS_STATUS = FALSE 'Mn Suppl Aid '
MF_STATUS = FALSE 'Mn Family Invest Program '
RC_STATUS = FALSE 'Refugee Cash Assistance'

'Reading the status and program
EMReadScreen cash1_status_check, 4, 6, 74
EMReadScreen cash2_status_check, 4, 7, 74
EMReadScreen emer_status_check, 4, 8, 74
EMReadScreen grh_status_check, 4, 9, 74
EMReadScreen fs_status_check, 4, 10, 74
'EMReadScreen ive_status_check, 4, 11, 74
'EMReadScreen hc_status_check, 4, 12, 74
'EMReadScreen cca_status_check, 4, 14, 74
EMReadScreen cash1_prog_check, 2, 6, 67
EMReadScreen cash2_prog_check, 2, 7, 67
EMReadScreen emer_prog_check, 2, 8, 67
EMReadScreen grh_prog_check, 2, 9, 67
EMReadScreen fs_prog_check, 2, 10, 67
'EMReadScreen ive_prog_check, 2, 11, 67
'EMReadScreen hc_prog_check, 2, 12, 67
'EMReadScreen cca_prog_check, 2, 14, 67

IF FS_status_check = "ACTV" or FS_status_check = "PEND"  THEN FS_STATUS = TRUE
IF emer_status_check = "ACTV" or emer_status_check = "PEND"  THEN ER_STATUS = TRUE
IF grh_status_check = "ACTV" or grh_status_check = "PEND"  THEN GRH_STATUS = TRUE
'IF hc_status_check = "ACTV" or hc_status_check = "PEND"  THEN HC_STATUS = TRUE
IF cca_status_check = "ACTV" or cca_status_check = "PEND"  THEN CCA_STATUS = TRUE
'Logic to determine if MFIP is active
If cash1_prog_check = "MF" or cash1_prog_check = "GA" or cash1_prog_check = "DW" or cash1_prog_check = "RC" or cash1_prog_check = "MS" THEN
	If cash1_status_check = "ACTV" Then CASH_STATUS = TRUE
	If cash1_status_check = "PEND" Then CASH_STATUS = TRUE
	If cash1_status_check = "INAC" Then CASH_STATUS = FALSE
	If cash1_status_check = "SUSP" Then CASH_STATUS = FALSE
	If cash1_status_check = "DENY" Then CASH_STATUS = FALSE
	If cash1_status_check = ""     Then CASH_STATUS = FALSE
END IF
If cash2_prog_check = "MF" or cash2_prog_check = "GA" or cash2_prog_check = "DW" or cash2_prog_check = "RC" or cash2_prog_check = "MS" THEN
	If cash2_status_check = "ACTV" Then CASH_STATUS = TRUE
	If cash2_status_check = "PEND" Then CASH_STATUS = TRUE
	If cash2_status_check = "INAC" Then CASH_STATUS = FALSE
	If cash2_status_check = "SUSP" Then CASH_STATUS = FALSE
	If cash2_status_check = "DENY" Then CASH_STATUS = FALSE
	If cash2_status_check = ""     Then CASH_STATUS = FALSE
END IF

IF SNAP_STATUS = FALSE or CASH_STATUS = FALSE THEN
    PROG_check = MsgBox("*** NOTICE!!!***" & vbNewLine & "This case does not appear to have snap or cash active."  & vbNewLine & "Continue to case note only?" & vbNewLine, vbYesNo + vbQuestion, "No cash or snap programs")
    IF PROG_check = vbYes THEN  case_note_only = TRUE
    IF PROG_check = vbNo THEN script_end_procedure("Please review the case if cash or snap were active previously select yes and case note only.")
END IF

IF SNAP_STATUS = TRUE or CASH_STATUS = TRUE THEN
    case_note_only = FALSE
    'Going to the MISC panel to add claim referral tracking information
    CALL navigate_to_MAXIS_screen ("STAT", "MISC")
    Row = 6
    EMReadScreen panel_number, 1, 02, 73
    IF panel_number = "0" THEN
        EMWriteScreen "NN", 20,79
        TRANSMIT
    ELSE
        DO
            'Checking to see if the MISC panel is empty, if not it will find a new line'
            EMReadScreen MISC_description, 25, row, 30
            MISC_description = replace(MISC_description, "_", "")
            IF trim(MISC_description) = "" THEN
                PF9
                EXIT DO
            ELSE
              row = row + 1
            END IF
        LOOP UNTIL row = 17
        IF row = 17 THEN MsgBox("There is not a blank field in the MISC panel. Please delete a line(s), and run script again or update manually.")
    END IF
    'writing in the action taken and date to the MISC panel
    PF9
    '_________________________ 25 characters to write on MISC
    IF claim_referral_tracking_dropdown = "Initial" THEN MISC_action_taken = "Claim Referral Initial"
    IF claim_referral_tracking_dropdown = "OP Non-Collectible (please specify)" THEN MISC_action_taken = "Determination-Non-Collect"
    IF claim_referral_tracking_dropdown = "No Savings/Overpayment" THEN MISC_action_taken = "Determination-No Savings"
    IF claim_referral_tracking_dropdown = "Overpayment Exists" THEN MISC_action_taken =  "Determination-OP Entered" '"Claim Determination 25 character available
    EMWriteScreen MISC_action_taken, Row, 30
    EMWriteScreen date, Row, 66
    TRANSMIT

    'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
    IF action_taken = "Sent Request for Additional Info" THEN Call create_TIKL("Potential overpayment exists on case. Please review case for receipt of additional requested information.", 10, date, False, TIKL_note_text)

    'The case note-------------------------------------------------------------------------------------------------
    start_a_blank_CASE_NOTE
    Call write_variable_in_case_note("***Claim Referral Tracking-" & action_taken & "***")
    Call write_bullet_and_variable_in_case_note("Action Date", action_date)
    Call write_bullet_and_variable_in_case_note("Pending Program(s)", pending_programs)
    CALL write_bullet_and_variable_in_CASE_NOTE ("Active Programs", active_programs)
    IF action_taken = "Sent Request for Additional Info" THEN Call write_bullet_and_variable_in_case_note("Action taken", MISC_action_taken)
    IF action_taken = "Sent Request for Additional Info" THEN CALL write_variable_in_case_note(TIKL_note_text)
    If action_taken = "Sent Request for Additional Info" THEN Call write_bullet_and_variable_in_case_note("Verification requested", verif_requested)
    Call write_variable_in_case_note("---")
    Call write_variable_in_case_note(worker_signature)

    IF action_taken = "Sent Request for Additional Info" THEN
        end_msg = end_msg & vbCr & ("Claim Referral Tracking - you have indicated that you sent a request for additional information. Please follow the agency's procedure(s) for claim entry once received.")
    ELSE
    	end_msg = end_msg & vbCr & ("Claim Referral Tracking - you have indicated that an overpayment exists. Please follow the agency's procedure(s) for claim entry.")
    END IF
ELSE
    case_note_only = TRUE
    IF case_note_only = TRUE THEN
        end_msg = end_msg & vbCr & "Claim Referral Tracking " & programs & " action " & action_taken 'we create some messaging to explain what happened in the script run.
        start_a_blank_CASE_NOTE
        Call write_variable_in_case_note("***Claim Referral Tracking " & action_taken & "***")
        Call write_bullet_and_variable_in_case_note("Pending Program(s)", pending_programs)
        CALL write_bullet_and_variable_in_CASE_NOTE ("Active Programs", active_programs)
        IF action_taken = "Sent Request for Additional Info" THEN Call write_bullet_and_variable_in_case_note("Action taken", MISC_action_taken)
        IF action_taken = "Sent Request for Additional Info" THEN CALL write_variable_in_case_note(TIKL_note_text)
        If action_taken = "Sent Request for Additional Info" THEN  Call write_bullet_and_variable_in_case_note("Verification requested", verif_requested)
        Call write_variable_in_case_note(worker_signature)
        PF3
    ELSE
        end_msg = "Claim Referral Tracking is for MFIP and SNAP cases only. Please let us know if there are further considerations needed."
    END IF
END IF

Call script_end_procedure_with_error_report(end_msg)

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------12/10/2021
'--Tab orders reviewed & confirmed----------------------------------------------12/29/2021
'--Mandatory fields all present & Reviewed--------------------------------------12/29/2021
'--All variables in dialog match mandatory fields-------------------------------12/29/2021
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------12/29/2021
'--CASE:NOTE Header doesn't look funky------------------------------------------12/29/2021
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------12/29/2021
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------12/29/2021
'--MAXIS_background_check reviewed (if applicable)------------------------------12/29/2021
'--PRIV Case handling reviewed -------------------------------------------------12/29/2021
'--Out-of-County handling reviewed----------------------------------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------12/29/2021
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------12/29/2021
'--Incrementors reviewed (if necessary)-----------------------------------------N/A
'--Denomination reviewed -------------------------------------------------------N/A
'--Script name reviewed---------------------------------------------------------12/29/2021
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub taks are complete-----------------------------------------12/29/2021
'--comment Code-----------------------------------------------------------------03/08/2022
'--Update Changelog for release/update------------------------------------------03/08/2022
'--Remove testing message boxes-------------------------------------------------03/08/2022
'--Remove testing code/unnecessary code-----------------------------------------03/08/2022
'--Review/update SharePoint instructions----------------------------------------03/08/2022
'--Review Best Practices using BZS page ----------------------------------------N/A
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------03/08/2022
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------03/08/2022
