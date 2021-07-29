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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("07/29/2021", "GitHub Issue #543 add verifications requested to dialog.", "MiKayla Handley, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'---------------------------------------------------------------------------------------The script
'Grabs the case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 296, 170, "Application Received for EMA"
  EditBox 90, 5, 55, 15, MAXIS_case_number
  DropListBox 90, 25, 55, 15, "Select One:"+chr(9)+"Fax"+chr(9)+"Mail"+chr(9)+"Office"+chr(9)+"Online", how_application_rcvd
  EditBox 90, 45, 55, 15, application_date
  DropListBox 210, 5, 75, 15, "Select One:"+chr(9)+"Health Jeopardy"+chr(9)+"Serious Impairment"+chr(9)+"Serious Dysfunction", consequence_type
  DropListBox 210, 25, 75, 15, "Select One:"+chr(9)+"Approved"+chr(9)+"Denied"+chr(9)+"Incomplete", action_taken
  EditBox 90, 65, 55, 15, start_date
  EditBox 90, 85, 55, 15, end_date
  EditBox 210, 45, 30, 15, HH_comp
  EditBox 210, 65, 60, 15, memb_id
  EditBox 90, 105, 195, 15, notes_income
  EditBox 90, 125, 195, 15, verification_requested
  EditBox 90, 150, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 195, 150, 45, 15
    CancelButton 240, 150, 45, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 160, 10, 50, 10, "Consequence:"
  Text 5, 30, 70, 10, "Application Received:"
  Text 5, 75, 70, 10, "Coverage Start Date:"
  Text 5, 130, 85, 10, "Verifications Requested:"
  Text 160, 50, 35, 10, "HH Comp:"
  Text 40, 95, 35, 10, "End Date:"
  Text 5, 55, 55, 10, "Application Date:"
  Text 5, 155, 60, 10, "Worker Signature:"
  Text 160, 70, 45, 10, "Identification:"
  Text 5, 110, 60, 10, "Notes on Income:"
  Text 160, 30, 50, 10, "Action Taken:"
EndDialog

'Runs the first dialog - which confirms the case number
Do
	Dialog Dialog1
	cancel_without_confirmation
	IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
	IF how_application_rcvd = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select how the application was received."
	IF consequence_type = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select type of Consequence."
	IF action_taken = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select type of action taken."
	IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "* Please enter your worker signature, for help see utilities."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

'-----------------------------------------------------------------------------------------CASENOTE
start_a_blank_CASE_NOTE
CALL write_variable_in_CASE_NOTE ("~ EMA Application received via " & how_application_rcvd & " for " & application_date & " ~")										'writes title in case note
CALL write_bullet_and_variable_in_case_note("EMMA begin date", start_date)							' writes the date the EMA began
CALL write_bullet_and_variable_in_case_note("EMMA end date", end_date)
CALL write_bullet_and_variable_in_case_note("Consequencee", consequence_type)		' writes how EMA affects clients health
CALL write_bullet_and_variable_in_case_note("HH Comp", HH_comp)										' writes the number of people in HH
CALL write_bullet_and_variable_in_case_note("Identification provided", memb_id)
CALL write_bullet_and_variable_in_case_note("Notes on income", notes_income)							' writes what type of income client has
CALL write_bullet_and_variable_in_case_note("Action taken", action_taken)		' writes outcome of application
CALL write_bullet_and_variable_in_case_note("Verifications requested:", verification_requested)		' writes outcome of application
CALL write_variable_in_case_note ("---")
CALL write_variable_in_case_note (worker_signature)

CALL script_end_procedure_with_error_report("")
