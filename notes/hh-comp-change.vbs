'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - HH COMP CHANGE.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 150                	'manual run time in seconds - INCLUDES A POLICY LOOKUP
STATS_denomination = "C"       		'C is for each CASE
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

'THIS SCRIPT IS BEING USED IN A WORKFLOW SO DIALOGS ARE NOT NAMED
'DIALOGS MAY NOT BE DEFINED AT THE BEGINNING OF THE SCRIPT BUT WITHIN THE SCRIPT FILE
'This script currently only has one dialog and so it can be defined in the beginning but is unnamedBeginDialog , 0, 0, 291, 175, "Household Comp Change"
  Text 5, 15, 50, 10, "Case Number"
  EditBox 60, 10, 100, 15, MAXIS_case_number
  Text 5, 35, 80, 10, "Unit Member HH Change"
  EditBox 90, 30, 45, 15, HH_member
  Text 5, 55, 85, 10, "Date Reported/Addendum"
  EditBox 95, 50, 60, 15, date_reported
  Text 160, 55, 55, 10, "Effective Date"
  EditBox 215, 50, 70, 15, effective_date
  CheckBox 110, 70, 120, 10, "Check if the change is temporary.", temporary_change_checkbox
  Text 10, 90, 45, 10, "Action Taken"
  EditBox 60, 85, 225, 15, actions_taken
  Text 5, 110, 60, 10, "Additional Notes"
  EditBox 60, 105, 225, 15, additional_notes
  Text 10, 130, 45, 15, "Worker Name"
  EditBox 60, 125, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 15, 150, 50, 15
    CancelButton 230, 150, 50, 15
EndDialog


'---SCRIPTS--------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone
EMConnect ""

'Finds the case number
Call MAXIS_case_number_finder(MAXIS_case_number)

'Finds the benefit month
EMReadScreen on_SELF, 4, 2, 50
IF on_SELF = "SELF" THEN
	CALL find_variable("Benefit Period (MM YY): ", MAXIS_footer_month, 2)
	IF MAXIS_footer_month <> "" THEN CALL find_variable("Benefit Period (MM YY): " & MAXIS_footer_month & " ", MAXIS_footer_year, 2)
ELSE
	CALL find_variable("Month: ", MAXIS_footer_month, 2)
	IF MAXIS_footer_month <> "" THEN CALL find_variable("Month: " & MAXIS_footer_month & " ", MAXIS_footer_year, 2)
END IF

check_for_maxis(False)

'Do loop for HHLD Comp Change Dialogbox
DO
	DO
		err_msg = ""
		DIALOG  					'Calling a dialog without a assigned variable will call the most recently defined dialog
		cancel_confirmation
		IF MAXIS_case_number = "" THEN err_msg = "You must enter case number!"
		IF HH_Member = "" THEN err_msg = err_msg & vbNewLine & "You must enter a HH Member"
		IF date_reported = "" THEN err_msg = err_msg & vbNewLine & "You must enter date reported"
		IF effective_date = "" THEN err_msg = err_msg & vbNewLine & "You must enter effective date"
		IF actions_taken = "" THEN err_msg = err_msg & vbNewLine & "You must enter the actions taken"
		IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "Please sign your note"
		IF err_msg <> "" THEN msgbox "*** Notice!!! ***" & vbNewLine & err_msg
	LOOP UNTIL err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'Checks MAXIS for password prompt
CALL check_for_MAXIS(false)

'Navigates to case note
CALL navigate_to_MAXIS_screen("CASE", "NOTE")

'Send PF9 to case note
PF9

CALL write_variable_in_case_note("HH Comp Change Reported")
CALL write_bullet_and_variable_in_Case_Note("Unit member HH Member", HH_Member)
CALL write_bullet_and_variable_in_Case_Note("Date Reported/Addendum", date_reported)
CALL write_bullet_and_variable_in_Case_Note("Date Effective", effective_date)
CALL write_bullet_and_variable_in_Case_Note("Actions Taken", actions_taken)
CALL write_bullet_and_variable_in_Case_Note("Additional Notes", additional_notes)

'case notes if the change is temporary
IF Temporary_Change_Checkbox = 1 THEN CALL write_variable_in_Case_Note("***Change is temporary***")
IF Temporary_Change_Checkbox = 0 THEN CALL write_variable_in_Case_Note("***Change is NOT temporary***")

'signs case note
CALL write_variable_in_Case_Note("----")
CALL write_variable_in_Case_Note(worker_signature)

script_end_procedure ("")
