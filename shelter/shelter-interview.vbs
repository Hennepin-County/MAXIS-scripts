'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - SHELTER-SHELTER INTERVIEW.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 600        	'manual run time in seconds was told 1800
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
CALL changelog_update("04/12/2022", "Elimination of Self-Pay: removal of mention from scripts.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/11/2016", "Initial version.", "Ilse Ferris, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Connecting to MAXIS, and grabbing the case number and footer month'
EMConnect ""
CALL check_for_maxis(FALSE) 'checking for passord out, brings up dialog'
CALL MAXIS_case_number_finder(MAXIS_case_number)
when_contact_was_made = now & ""
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 316, 325, "Shelter Interview"
  EditBox 70, 5, 50, 15, MAXIS_case_number
  EditBox 290, 5, 15, 15, MEMB_number
  EditBox 70, 25, 90, 15, when_contact_was_made
  EditBox 100, 50, 205, 15, barriers_housing
  EditBox 100, 70, 205, 15, shelter_history
  EditBox 105, 100, 200, 15, reason_homeless
  EditBox 105, 120, 200, 15, name_verified
  EditBox 105, 140, 70, 15, address_verf
  EditBox 255, 140, 50, 15, phone_number
  EditBox 105, 160, 200, 15, comments_notes
  EditBox 105, 185, 200, 15, referrals_made
  EditBox 105, 205, 200, 15, social_worker
  EditBox 105, 225, 200, 15, other_income
  EditBox 105, 245, 200, 15, other_notes
  DropListBox 240, 265, 65, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO"+chr(9)+"N/A", screening_questions_dropdown
  EditBox 175, 285, 130, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 200, 305, 50, 15
    CancelButton 255, 305, 50, 15
	PushButton 5, 305, 65, 15, "SHELTER ", shelter_button
  Text 20, 10, 50, 10, "Case Number:"
  Text 175, 10, 115, 10, "Interview Completed with Member "
  Text 225, 20, 80, 10, "(Enter member Number)"
  Text 5, 30, 65, 10, "Contact Date/Time:"
  Text 15, 55, 70, 10, "Barrier(s) to housing:"
  Text 35, 75, 55, 10, "Shelter history:"
  GroupBox 5, 90, 305, 90, "Homelessness Verified"
  Text 10, 105, 90, 10, "Reason for homelessness:"
  Text 75, 125, 25, 10, "Name:"
  Text 70, 145, 30, 10, "Address:"
  Text 65, 165, 40, 10, "Comments: "
  Text 35, 190, 60, 10, "Referrals made to:"
  Text 50, 210, 50, 10, "Social worker:"
  Text 50, 230, 45, 10, "Other income:"
  Text 55, 250, 40, 10, "Other notes:"
  Text 105, 270, 130, 10, "Health screening questions answered:"
  Text 105, 290, 65, 10, "Worker Signature:"
  Text 200, 145, 50, 10, "Phone number: "
EndDialog

DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
        Call validate_MAXIS_case_number(err_msg, "*")
		If reason_homeless = "" then err_msg = err_msg & vbNewLine & "* Please enter the reason for family's homelessness."
		If barriers_housing = "" then err_msg = err_msg & vbNewLine & "* Please enter the family's barrier(s) to housing."
		If referrals_made = "" then err_msg = err_msg & vbNewLine & "* Please enter referrals made for the family."
		IF screening_questions_dropdown =  "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please enter if health screening questions was answered."
		If worker_signature = "" Then err_msg = err_msg & vbNewLine & "* Please enter your worker signature."
		If ButtonPressed = shelter_button Then               'Pulling up the hsr page if the button was pressed.
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Shelter_Team.aspx"
			err_msg = "LOOP"
		Else                                                'If the instructions button was NOT pressed, we want to display the error message if it exists.
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
		End If
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

CALL determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
CALL back_to_self 'to ensure no funky business happens before the case note'
'The case note'
start_a_blank_CASE_NOTE
CALL write_variable_in_CASE_NOTE("### Household in Shelter - Interview ####")
Call write_bullet_and_variable_in_CASE_NOTE("Interview Completed With", MEMB_number)
Call write_bullet_and_variable_in_CASE_NOTE("Active programs", list_active_programs)
Call write_bullet_and_variable_in_CASE_NOTE("Pending programs", list_pending_programs)
CALL write_bullet_and_variable_in_CASE_NOTE("Date:", when_contact_was_made)
Call write_bullet_and_variable_in_CASE_NOTE("Barrier(s) to housing", barriers_housing)
Call write_bullet_and_variable_in_CASE_NOTE("Shelter history", shelter_history)
Call write_bullet_and_variable_in_CASE_NOTE("Reason for homelessness", reason_homeless)
IF name_verified <> "" THEN
	CALL write_variable_in_CASE_NOTE("* Homelessness verified")
    CALL write_variable_in_CASE_NOTE("   Name: " & name_verified)
    CALL write_variable_in_CASE_NOTE("   Phone number: " & phone_number)
    CALL write_variable_in_CASE_NOTE("   Address: " & address_verf)
ELSE
 	CALL write_variable_in_CASE_NOTE("Homelessness not verified")
END IF
CALL write_variable_in_CASE_NOTE("* Comments: " & comments_notes)
Call write_bullet_and_variable_in_CASE_NOTE("Social worker", social_worker)
Call write_bullet_and_variable_in_CASE_NOTE("Referrals made to", referrals_made)
Call write_bullet_and_variable_in_CASE_NOTE("Other income", other_income)
Call write_bullet_and_variable_in_CASE_NOTE("Health screening questions answered", screening_questions_dropdown) 'not provided these questions unclear if this is a barrier '
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE("* Explained shelter policies and client options to shelter such as bus tickets, temporary housing, private shelters, etc.")
Call write_variable_in_CASE_NOTE("* Client given family social services number (348-4111) to discuss any family issues/barriers.")
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

Call script_end_procedure_with_error_report("Shelter Interview Note Entered.")


'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------04/12/2022
'--Tab orders reviewed & confirmed----------------------------------------------04/12/2022
'--Mandatory fields all present & Reviewed--------------------------------------04/12/2022				It is very unclear about what the script is case noting here
'--All variables in dialog match mandatory fields-------------------------------04/12/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------04/12/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------04/12/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------05/02/2022
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------04/12/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------04/12/2022
'--PRIV Case handling reviewed -------------------------------------------------04/12/2022
'--Out-of-County handling reviewed----------------------------------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------04/12/2022
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------04/12/2022 					I was told an interview takes 30-60 minutes but this script does not reflect that
'--Incrementors reviewed (if necessary)-----------------------------------------N/A
'--Denomination reviewed -------------------------------------------------------N/A
'--Script name reviewed---------------------------------------------------------N/A
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------N/A
'--comment Code-----------------------------------------------------------------N/A
'--Update Changelog for release/update------------------------------------------N/A
'--Remove testing message boxes-------------------------------------------------N/A
'--Remove testing code/unnecessary code-----------------------------------------N/A
'--Review/update SharePoint instructions----------------------------------------N/A
'--Review Best Practices using BZS page ----------------------------------------N/A
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------N/A
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------N/A
