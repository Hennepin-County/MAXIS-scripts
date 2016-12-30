'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - MTAF.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 300          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS----------------------------------------------------------------------------------------------------
'CASE NUMBER DIALOG
BeginDialog case_number_dialog, 0, 0, 126, 45, "Case number dialog"
  EditBox 55, 5, 65, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 10, 25, 50, 15
    CancelButton 65, 25, 50, 15
  Text 5, 10, 50, 10, "Case number:"
EndDialog

'MTAF DIALOG
BeginDialog MTAF_dialog, 0, 0, 506, 345, "MTAF dialog"
  EditBox 45, 5, 60, 15, MTAF_date
  EditBox 160, 5, 60, 15, MFIP_elig_date
  EditBox 275, 5, 60, 15, interview_date
  DropListBox 275, 25, 60, 15, "Select one..."+chr(9)+"complete"+chr(9)+"incomplete", MTAF_status_dropdown
  EditBox 75, 45, 260, 15, ADDR_change
  EditBox 75, 65, 260, 15, HHcomp_change
  EditBox 75, 85, 260, 15, asset_change
  EditBox 105, 105, 230, 15, earned_income_change
  EditBox 105, 125, 230, 15, unearned_income_change
  EditBox 105, 145, 230, 15, shelter_costs_change
  EditBox 175, 165, 160, 15, subsidized_housing
  DropListBox 175, 185, 160, 15, "Select one..."+chr(9)+"Not subsidized"+chr(9)+"Verification provided"+chr(9)+"Verification pending", sub_housing_droplist
  EditBox 110, 200, 225, 15, child_adult_care_costs
  EditBox 110, 220, 225, 15, relationship_proof
  EditBox 175, 240, 160, 15, referred_to_OMB_PBEN
  EditBox 125, 260, 210, 15, elig_results_fiated
  EditBox 75, 280, 260, 15, other_notes
  EditBox 75, 300, 260, 15, verifications_needed
  If worker_county_code = "x127" or worker_county_code = "x162" then CheckBox 5, 325, 180, 10, "Sent MFIP financial orientation DVD to participant(s).", MFIP_DVD_checkbox
  CheckBox 350, 45, 135, 10, "Rights and responsibilities explained.", RR_explained_checkbox
  CheckBox 350, 60, 55, 10, "MTAF signed.", mtaf_signed_checkbox
  CheckBox 350, 75, 140, 10, "MFIP/financial orientation completed.", mfip_financial_orientation_checkbox
  CheckBox 350, 90, 150, 10, "Client exempt from cooperation with ES.", ES_exemption_checkbox
  EditBox 295, 320, 95, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 395, 320, 50, 15
    CancelButton 450, 320, 50, 15
  Text 110, 10, 50, 10, "MFIP elig date:"
  Text 225, 10, 50, 10, "Interview date:"
  Text 5, 30, 210, 10, "**Changes reported on MTAF**  (Complete boxes as applicable.)"
  Text 225, 30, 45, 10, "MTAF status:"
  Text 5, 50, 70, 10, "Change in address:"
  Text 5, 70, 70, 10, "Change in HH comp:"
  Text 5, 90, 70, 10, "Change in assets:"
  Text 5, 110, 90, 10, "*Change in earned income:"
  Text 5, 130, 95, 10, "Change in unearned income:"
  Text 5, 150, 95, 10, "Change in shelter costs:"
  Text 5, 170, 170, 10, "Is housing subsidized? If so, what is the amount?"
  Text 75, 185, 100, 10, "**Subsidized housing status:"
  Text 5, 205, 85, 10, "Child or adult care costs:"
  Text 5, 225, 95, 10, "Proof of relationship on file:"
  Text 5, 245, 160, 10, "Client has been referred to apply for OMB/PBEN:"
  Text 5, 265, 115, 10, "Eligibility results fiated? If so, why:"
  Text 5, 285, 45, 10, "Other notes:"
  Text 5, 305, 70, 10, "Verifications needed:"
  GroupBox 350, 105, 150, 100, ""
  Text 360, 115, 135, 35, "*STOP WORK - Verification only necessary to verify income in the month of application/eligibility. (CM 0010.18.01)"
  Text 360, 155, 135, 45, "**SUBSIDY - Verification of housing subsidy is a mandatory verification for MFIP. STAT must be appropriately updated to ensure accurate approval of housing grant. (CM 0010.18.01)"
  Text 230, 325, 60, 10, "Worker signature:"
  Text 5, 10, 40, 10, "MTAF date:"
EndDialog

'SCRIPT STUFF=================================================================================================================
EMConnect ""
'Grabs case number from the screen.
Call MAXIS_case_number_finder(MAXIS_case_number)

'This is the running of the case number dialog.
Do
	Dialog case_number_dialog 'Runs the "case number" dialog
	If buttonpressed = 0 then stopscript 'If someone hits "cancel" the script stops.
	If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then MsgBox "You need to type a valid case number." 'If the case number is blank, is not numeric, or is longer than 8 characters, then message box.
Loop until MAXIS_case_number <> "" and IsNumeric(MAXIS_case_number) = True and len(MAXIS_case_number) <= 8 'Loop until case number is not blank, is numeric, and is smaller than or equal to 8 characters in length.

'Time to make a case note!
DO
	Do
		Do
			err_msg = ""		'Resetting the error message variable to be blank.
			Dialog MTAF_dialog	'Displays the MTAF dialog box.
			cancel_confirmation		'Asks if you're sure you want to cancel, and cancels if you select that.
			If MTAF_status_dropdown = "Select one..." then err_msg = err_msg & vbNewLine & "- Please indicate MTAF status." 'MTAF status must be selected, or error message is displayed.
			If sub_housing_droplist = "Select one..." then err_msg = err_msg & vbNewLine & "- Please indicate whether or not housing is subsidized, and status of verification if applicable" 'Subsidized housing status/verifications must be selected, or error message is displayed.
			If relationship_proof = "" then err_msg = err_msg & vbNewLine & "- You must indicate what form of proof is being used as verification of relationship for all household members." 'Proof of relationship must be indicated, or error message is displayed.
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		Loop until err_msg = ""
		CALL proceed_confirmation(case_note_confirm)	'Checks to make sure that we are ready to case note.
	Loop until case_note_confirm = TRUE
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'Takes script to a blank case note.
Call start_a_blank_case_note

'THE CASE NOTE===========================================================================
CALL write_variable_in_CASE_NOTE("***MTAF Interview Completed***")
CALL write_bullet_and_variable_in_CASE_NOTE ("Date received", MTAF_date)
CALL write_bullet_and_variable_in_CASE_NOTE ("Date of eligibility", MFIP_elig_date)
CALL write_bullet_and_variable_in_CASE_NOTE ("Date of interview", interview_date)
CALL write_bullet_and_variable_in_CASE_NOTE ("Address change", ADDR_change)
CALL write_bullet_and_variable_in_CASE_NOTE ("Household composition change", HHcomp_change)
CALL write_bullet_and_variable_in_CASE_NOTE ("Change in assets", asset_change)
CALL write_bullet_and_variable_in_CASE_NOTE ("Change in earned income", earned_income_change)
CALL write_bullet_and_variable_in_CASE_NOTE ("Change in unearned income", unearned_income_change)
CALL write_bullet_and_variable_in_CASE_NOTE ("Change in shelter costs", shelter_costs_change)
CALL write_bullet_and_variable_in_CASE_NOTE ("Is housing subsidized? If so, what is the amount?", subsidized_housing)
CALL write_bullet_and_variable_in_CASE_NOTE ("Subsidized housing status", sub_housing_droplist)
CALL write_bullet_and_variable_in_CASE_NOTE ("Child or adult care costs", child_adult_care_costs)
CALL write_bullet_and_variable_in_CASE_NOTE ("Proof of relationship on file", relationship_proof)
CALL write_bullet_and_variable_in_CASE_NOTE ("Referred to apply for OMB/PBEN", referred_to_OMB_PBEN)
CALL write_bullet_and_variable_in_CASE_NOTE ("ELIG results fiated", elig_results_fiated)
CALL write_bullet_and_variable_in_CASE_NOTE ("Other notes", other_notes)
CALL write_bullet_and_variable_in_CASE_NOTE ("Verifications Needed", verifications_needed)
IF MFIP_DVD_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Sent MFIP orientation DVD to participant(s).")
If RR_explained_checkbox = checked THEN CALL write_variable_in_CASE_NOTE ("* Rights & responsibilities explained.")
If mtaf_signed_checkbox = checked THEN CALL write_variable_in_CASE_NOTE ("* MTAF was signed.")
If mfip_financial_orientation_checkbox = checked THEN CALL write_variable_in_CASE_NOTE ("* MFIP orientation information reviewed/completed.")
If ES_exemption_checkbox = checked THEN CALL write_variable_in_CASE_NOTE ("* Client is exempt from cooperation with ES at this time.")
CALL write_bullet_and_variable_in_CASE_NOTE ("MTAF Status", MTAF_status_dropdown)
CALL write_variable_in_CASE_NOTE ("---")
CALL write_variable_in_CASE_NOTE (worker_signature)

script_end_procedure ("")
