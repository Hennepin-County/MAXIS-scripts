'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - IMIG - EMA.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 270          'manual run time in seconds
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

'retirement redirect of NOTES - IMIG - EMA
redirect_verbiage = "This script (NOTES - IMIG - EMA) was retired on 4/25/23."
redirect_verbiage = redirect_verbiage & vbCr & vbCr & "NOTES - Health Care Evaulation is the script to support documenting details about new requests for Health Care."
redirect_verbiage = redirect_verbiage & vbCr & "- - You can check out the Hot Topic and "
redirect_verbiage = redirect_verbiage & vbCr & "    Video Demo about using this new script."
redirect_verbiage = redirect_verbiage & vbCr & "----------------------------------------------"
redirect_verbiage = redirect_verbiage & vbCr & "This script will now run NOTES - Health Care Evaluation"
redirect_to_hc_eval = MsgBox(redirect_verbiage, 48, "Redirecting to new NOTES - Health Care Evaluation Script")
Call run_from_GitHub(script_repository & "notes/health-care-evaluation.vbs")

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("04/11/2023", "This script will be replaced with a new script 'NOTES - Health Care Evaluation' in the coming weeks. Look for additional information next week.", "Casey Love, Hennepin County")
call changelog_update("09/12/2022", "Updated TIKL functionality and TIKL text in the case note.", "MiKayla Handley, Hennepin County")   '#944
call changelog_update("05/24/2022", "CASE/NOTE format updated to exclude the 'How App Received' detail. This information is important for the script operation, but is not necessary to be included in the CASE/NOTE", "Casey Love, Hennepin County")   '#799
call changelog_update("07/29/2021", "GitHub Issue #543 add verifications requested to dialog.", "MiKayla Handley, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'---------------------------------------------------------------------------------------The script
'Grabs the case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)
CALL check_for_MAXIS(False)
back_to_SELF 'to ensure we are not inside  case note'
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 291, 160, "Application Received for EMERGENCY MEDICAL ASSISTANCE"
  EditBox 80, 5, 55, 15, MAXIS_case_number
  EditBox 80, 25, 55, 15, application_date
  EditBox 80, 45, 55, 15, begin_date
  EditBox 80, 65, 55, 15, end_date
  DropListBox 210, 5, 75, 15, "Select One:"+chr(9)+"Health Jeopardy"+chr(9)+"Serious Impairment"+chr(9)+"Serious Dysfunction", consequence_type
  DropListBox 210, 25, 75, 15, "Select One:"+chr(9)+"Approved"+chr(9)+"Denied"+chr(9)+"Incomplete", application_status
  EditBox 210, 45, 30, 15, HH_comp
  EditBox 210, 65, 75, 15, memb_id
  EditBox 90, 85, 195, 15, notes_income
  EditBox 90, 105, 195, 15, verification_requested
  EditBox 90, 125, 100, 15, worker_signature
  CheckBox 5, 145, 245, 10, "Check here to have the script create a TIKL to deny at the 45 day mark.", TIKL_check
  ButtonGroup ButtonPressed
    OkButton 195, 125, 45, 15
    CancelButton 240, 125, 45, 15
  Text 30, 10, 50, 10, "Case Number:"
  Text 20, 30, 55, 10, "Application Date:"
  Text 5, 50, 75, 10, "Coverage Begin Date:"
  Text 40, 70, 35, 10, "End Date:"
  Text 160, 10, 50, 10, "Consequence:"
  Text 145, 30, 65, 10, "Application Status:"
  Text 170, 50, 35, 10, "HH Comp:"
  Text 160, 70, 45, 10, "Identification:"
  Text 25, 90, 60, 10, "Notes on Income:"
  Text 5, 110, 85, 10, "Verifications Requested:"
  Text 25, 130, 60, 10, "Worker Signature:"
EndDialog

'Runs the dialog - which confirms the case number
Do
	DO
	   	err_msg = ""
	    Dialog Dialog1
	    cancel_without_confirmation
	    CALL validate_MAXIS_case_number(err_msg, "*")
		IF isdate(begin_date) = FALSE THEN err_msg = err_msg & vbnewline & "Please enter the begin date."
		IF end_date <> "" THEN
			IF isdate(end_date) = FALSE THEN err_msg = err_msg & vbnewline & "Please enter the end date."
		END IF
	    IF IsDate(application_date) = False Then err_msg = err_msg & vbNewLine & "* Please enter a valid date for the date the application was received."
		IF consequence_type = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select type of consequence."
		IF application_status = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select the application status."
		IF application_status = "Approved" and TIKL_check = CHECKED THEN err_msg = err_msg & vbNewLine & "* Please check the application status and TIKL request, you indicated the case will be approved this TIKL is for denial."
		IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "* Please enter your worker signature, for help see utilities."
	    IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

'Navigating to STAT/HCRE
CALL navigate_to_MAXIS_screen_review_PRIV("STAT", "HCRE", is_this_priv)
IF is_this_priv = TRUE THEN script_end_procedure("This case is privileged, the script will now end.")

'TIKL portion-----------------------------------------------------------------------------------------------------------------------------
'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
If TIKL_check = 1 then Call create_TIKL("EMA pending 45 days. Evaluate for possible denial.", 45, application_date, False, TIKL_note_text)
'-----------------------------------------------------------------------------------------CASENOTE
start_a_blank_CASE_NOTE
CALL write_variable_in_CASE_NOTE ("~ EMA Application received for " & application_date & " ~")										'writes title in case note
CALL write_bullet_and_variable_in_case_note("EMMA begin date", begin_date)							' writes the date the EMA began
CALL write_bullet_and_variable_in_case_note("EMMA end date", end_date)
CALL write_bullet_and_variable_in_case_note("Consequencee", consequence_type)		' writes how EMA affects clients health
CALL write_bullet_and_variable_in_case_note("HH Comp", HH_comp)										' writes the number of people in HH
CALL write_bullet_and_variable_in_case_note("Identification provided", memb_id)
CALL write_bullet_and_variable_in_case_note("Notes on income", notes_income)							' writes what type of income client has
CALL write_bullet_and_variable_in_case_note("Application status", application_status)		' writes outcome of application
CALL write_bullet_and_variable_in_case_note("Verifications requested", verification_requested)		' writes outcome of application
IF TIKL_check = 1 THEN CALL write_variable_in_case_note(TIKL_note_text)
CALL write_variable_in_case_note ("---")
CALL write_variable_in_case_note (worker_signature)

CALL script_end_procedure_with_error_report("")

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------05/24/2022
'--Tab orders reviewed & confirmed----------------------------------------------05/24/2022
'--Mandatory fields all present & reviewed--------------------------------------05/24/2022
'--All variables in dialog match mandatory fields-------------------------------05/24/2022
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------05/24/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------05/24/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------05/24/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------N/A
'--Out-of-County handling reviewed----------------------------------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------05/24/2022
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--write_variable_in_CASE_NOTE confirm that proper punctuation is used----------09/12/2022 not used except in header
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------05/24/2022
'--Incrementors reviewed (if necessary)-----------------------------------------N/A
'--Denomination reviewed -------------------------------------------------------05/24/2022
'--Script name reviewed---------------------------------------------------------05/24/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A
'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub taks are complete-----------------------------------------05/24/2022
'--Comment Code-----------------------------------------------------------------05/24/2022
'--Update Changelog for release/update------------------------------------------09/12/2022
'--Remove testing message boxes-------------------------------------------------05/24/2022
'--Remove testing code/unnecessary code-----------------------------------------05/24/2022
'--Review/update SharePoint instructions----------------------------------------N/A
'--Review Best Practices using BZS page ----------------------------------------N/A
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------N/A
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------N/A
