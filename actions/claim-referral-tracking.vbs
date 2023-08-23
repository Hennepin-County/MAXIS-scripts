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
function claim_referral_tracking(action_taken, action_date)
'--- This function tracks the date a worker first suspects there may be a SNAP or MFIP claim. It also helps to track the discovery date and the established date of a claim. This will create or update the MISC panel and case note the referral.
'~~~~~ action_taken: 4 options exist for clearing claim referral "Sent Verification Request", "Determination-OP Entered","Determination-Non-Collect", "No Savings/Overpayment" each has different handling
'===== Keywords: MAXIS, Claim, MISC, CCOL, overpayment
    CALL MAXIS_case_number_finder(MAXIS_case_number)    'Grabbing the CASE Number
    CALL MAXIS_footer_month_confirmation()
    CALL Check_for_MAXIS(false)                         'Ensuring we are not passworded out
    action_date = date & ""

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 221, 155, "Claim Referral Tracking"
      EditBox 65, 30, 45, 15, MAXIS_case_number
      EditBox 165, 30, 45, 15, action_date
      DropListBox 65, 50, 145, 15, "Select One:"+chr(9)+"Sent Verification Request"+chr(9)+"Determination-OP Entered"+chr(9)+"Determination-Non-Collect"+chr(9)+"No Savings/Overpayment", action_taken
      EditBox 65, 70, 145, 15, verif_requested
      EditBox 65, 90, 145, 15, other_notes
      EditBox 110, 110, 100, 15, worker_signature
      ButtonGroup ButtonPressed
        OkButton 115, 135, 45, 15
        CancelButton 165, 135, 45, 15
        PushButton 5, 135, 85, 15, "Claims Procedures", claims_procedures_btn
      Text 5, 5, 210, 20, "This script will only enter a STAT/MISC panel for a SNAP or MFIP federal food claim."
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
    	    cancel_without_confirmation
    	    IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
    	    IF isdate(action_date) = False then err_msg = err_msg & vbnewline & "* Please enter a valid action date."
    	    IF action_taken = "Select One:" then err_msg = err_msg & vbnewline & "* Please select the action taken for next step in overpayment."
            IF action_taken = "Sent Verification Request" and verif_requested = "" then err_msg = err_msg & vbnewline & "* You selected that a request for additional information was sent, please advise what verifications were requested."
            IF action_taken = "Determination-Non-Collect" and other_notes = "" THEN err_msg = err_msg & vbNewLine & "* Please the reason the claim is non-collectible."
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
    Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, active_programs, pending_programs)
    EMReadScreen appl_date, 8, 8, 29               'Grabbing the PND2 date from CASE CURR in case the information cannot be pulled from REPT/PND2

    Call back_to_SELF

    claim_referral = False
    If snap_case = True Then claim_referral = True
    If mfip_case = True Then claim_referral = True

    IF claim_referral = False then
        PROG_check = MsgBox("*** NOTICE!!!***" & vbNewLine & "This case does not appear to have snap or cash active."  & vbNewLine & "Continue to case note only?" & vbNewLine, vbYesNo + vbQuestion, "No cash or snap programs")
        IF PROG_check = vbYes THEN case_note = True
        IF PROG_check = vbNo THEN
            case_note = False
            end_msg = end_msg & "Please review the case if cash or snap were active previously select yes and case note only."
        End if
    END IF

    IF claim_referral = True then
        case_note = True
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
        EMWriteScreen action_taken, Row, 30
        EMWriteScreen date, Row, 66
        TRANSMIT 'to save the work'

        'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
        IF action_taken = "Sent Verification Request" THEN Call create_TIKL("Potential overpayment exists on case. Please review case for receipt of additional requested information.", 10, date, False, TIKL_note_text)
    End if

    '-------------------------------------------------------------------------------------------------case note
    If case_note = True then
        start_a_blank_CASE_NOTE
        Call write_variable_in_case_note("***Claim Referral Tracking-" & action_taken & "***")
        Call write_bullet_and_variable_in_case_note("Action Date", action_date)
        Call write_bullet_and_variable_in_case_note("Pending Program(s)", pending_programs)
        CALL write_bullet_and_variable_in_CASE_NOTE("Active Programs", active_programs)
        If action_taken = "Sent Verification Request" THEN Call write_bullet_and_variable_in_case_note("Verification requested", verif_requested)
        IF action_taken = "Sent Verification Request" THEN CALL write_variable_in_case_note(TIKL_note_text)
        If action_taken = "Determination-OP Entered" THEN CALL write_variable_in_case_note("* Claim entered") 'this should call OP?'
        IF action_taken = "Determination-Non-Collect" THEN Call write_variable_in_case_note("* Claim is non-collectible.")
        IF action_taken = "No Savings/Overpayment" THEN Call write_variable_in_case_note("* No overpayment was found after review.")
        CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
        If action_taken <> "Determination-OP Entered" THEN CALL write_variable_in_case_note("* Entries for these potential claims must be retained until further notice.")
        Call write_variable_in_case_note("---")
        Call write_variable_in_case_note(worker_signature)
        IF action_taken = "Sent Verification Request" THEN end_msg = end_msg & vbCr & ("Claim Referral Tracking - you have indicated that you sent a request for additional information. Please follow the agency's procedure(s) for claim entry once received.")
        IF action_taken = "Determination-OP Entered" THEN end_msg = end_msg & vbCr & ("Claim Referral Tracking - you have indicated that an overpayment exists. Please follow the agency's procedure(s) for claim entry.")
        IF action_taken = "Determination-Non-Collect" THEN end_msg = end_msg & vbCr & ("Claim Referral Tracking - you have indicated that an overpayment exists, but is non-collectible. Please follow the agency's procedure(s) for claim entry.")
        PF3
    ELSE
        IF case_active = FALSE or case_rein = FALSE THEN
            CALL write_variable_in_case_note("Maxis case is inactive unable to add or update MISC panel")
            end_msg = end_msg & vbCr & ("Claim Referral Tracking is for MFIP and SNAP cases only. Please let us know if there are further considerations needed.")
        END IF
    END IF
    Call script_end_procedure_with_error_report(end_msg)
END FUNCTION

'----------------------------------------------------------------------------------------------------The script :)
EMConnect ""                                        'Connecting to BlueZone
Call claim_referral_tracking(action_taken, action_date)

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
