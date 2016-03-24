'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - MTAF.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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

'Required for statistical purposes==========================================================================================
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 300          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'DIALOGS----------------------------------------------------------------------------------------------------
'CASE NUMBER DIALOG
BeginDialog case_number_dialog, 0, 0, 126, 45, "Case number dialog"
  EditBox 55, 5, 65, 15, case_number
  ButtonGroup ButtonPressed
    OkButton 10, 25, 50, 15
    CancelButton 65, 25, 50, 15
  Text 5, 10, 50, 10, "Case number:"
EndDialog


'MTAF DIALOG
BeginDialog MTAF_dialog, 0, 0, 526, 340, "MTAF dialog"
 EditBox 45, 5, 60, 15, MTAF_date
  EditBox 160, 5, 60, 15, MFIP_elig_date
  EditBox 275, 5, 60, 15, interview_date
  DropListBox 275, 30, 60, 15, "Select one..."+chr(9)+"complete"+chr(9)+"incomplete", MTAF_status_dropdown
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
  CheckBox 350, 45, 135, 10, "Rights and responsibilities explained.", RR_explained_checkbox
  CheckBox 350, 60, 55, 10, "MTAF signed.", mtaf_signed_checkbox
  CheckBox 350, 75, 150, 10, "MFIP/financial orientation completed.", mfip_financial_orientation_checkbox
  CheckBox 350, 90, 200, 10, "Client exempt from cooperation with ES.", ES_exemption_checkbox
  'CheckBox 5, 325, 115, 10, "Open approved programs script", open_approved_programs_checkbox
  'CheckBox 130, 325, 110, 10, "Open denied programs script", open_denied_programs_checkbox
  EditBox 340, 320, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 425, 320, 50, 15
    CancelButton 475, 320, 50, 15
  Text 5, 10, 40, 10, "MTAF date:"
  Text 110, 10, 50, 10, "MFIP elig date:"
  Text 225, 10, 55, 15, "Interview date:"
  Text 5, 30, 225, 15, "**Changes reported on MTAF**  (Complete boxes as applicable.)"
  Text 225, 30, 45, 15, "MTAF status:"
  Text 5, 50, 70, 15, "Change in address:"
  Text 5, 70, 70, 15, "Change in HH comp:"
  Text 5, 90, 70, 15, "Change in assets:"
  Text 5, 110, 90, 15, "*Change in earned income:"
  Text 5, 130, 95, 15, "Change in unearned income:"
  Text 5, 150, 95, 15, "Change in shelter costs:"
  Text 5, 170, 170, 15, "Is housing subsidized? If so, what is the amount?"
  Text 75, 185, 90, 10, "**Subsidized housing status:"
  Text 5, 200, 85, 15, "Child or adult care costs:"
  Text 5, 220, 95, 15, "Proof of relationship on file:"
  Text 5, 240, 160, 15, "Client has been referred to apply for OMB/PBEN:"
  Text 5, 265, 115, 15, "Eligibility results fiated? If so, why:"
  Text 5, 285, 45, 10, "Other notes:"
  Text 5, 305, 70, 10, "Verifications needed:"
  GroupBox 350, 105, 150, 100, ""
  Text 360, 115, 135, 35, "*STOP WORK - Verification only necessary to verify income in the month of application/eligibility. (CM 0010.18.01)"
  Text 360, 155, 135, 45, "**SUBSIDY - Verification of housing subsidy is a mandatory verification for MFIP. STAT must be appropriately updated to ensure accurate approval of housing grant. (CM 0010.18.01)"
  Text 275, 325, 60, 10, "Worker signature:"
EndDialog

'SCRIPT STUFF=================================================================================================================
EMConnect ""


'Makes sure you are in MAXIS, to avoid password-out scenario.
Call check_for_MAXIS(True)

'Grabs case number from the screen.
Call MAXIS_case_number_finder(case_number)

'This is the running of the case number dialog.
Do 
	Dialog case_number_dialog 'Runs the "case number" dialog
	If buttonpressed = 0 then stopscript 'If someone hits "cancel" the script stops.
	If case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 then MsgBox "You need to type a valid case number." 'If the case number is blank, is not numeric, or is longer than 8 characters, then message box.
Loop until case_number <> "" and IsNumeric(case_number) = True and len(case_number) <= 8 'Loop until case number is not blank, is numeric, and is smaller than or equal to 8 characters in length.

'Makes sure you are in MAXIS, to avoid password-out scenario.
Call check_for_MAXIS(True)

'Time to make a case note!
DO
	Do
		Do
			err_msg = ""		'Resetting the error message variable to be blank.
			Dialog MTAF_dialog	'Displays the MTAF dialog box.
			cancel_confirmation		'Asks if you're sure you want to cancel, and cancels if you select that.
			If MTAF_status_dropdown = "Select one..." then err_msg = "- Please indicate MTAF status." 'MTAF status must be selected, or error message is displayed.
			If sub_housing_droplist = "Select one..." then err_msg = err_msg & vbNewLine & "- Please indicate whether or not housing is subsidized, and status of verification if applicable" 'Subsidized housing status/verifications must be selected, or error message is displayed.
			If relationship_proof = "" then err_msg = err_msg & vbNewLine & "- You must indicate what form of proof is being used as verification of relationship for all household members." 'Proof of relationship must be indicated, or error message is displayed.
			If err_msg <> "" Then Msgbox err_msg
		Loop until err_msg = ""
		CALL proceed_confirmation(case_note_confirm)	'Checks to make sure that we are ready to case note.
	Loop until case_note_confirm = TRUE
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'Makes sure you are in MAXIS, to avoid password-out scenario.
Call check_for_MAXIS(True)

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
If RR_explained_checkbox = checked THEN CALL write_variable_in_CASE_NOTE ("* Rights & responsibilities explained.")
If mtaf_signed_checkbox = checked THEN CALL write_variable_in_CASE_NOTE ("* MTAF was signed.")
If mfip_financial_orientation_checkbox = checked THEN CALL write_variable_in_CASE_NOTE ("* MFIP orientation information reviewed/completed.")
If ES_exemption_checkbox = checked THEN CALL write_variable_in_CASE_NOTE ("* Client is exempt from cooperation with ES at this time.")
CALL write_bullet_and_variable_in_CASE_NOTE ("MTAF Status", MTAF_status_dropdown)
CALL write_variable_in_CASE_NOTE ("---")
CALL write_variable_in_CASE_NOTE (worker_signature)

script_end_procedure ("")


