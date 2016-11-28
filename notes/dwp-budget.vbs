'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - DWP BUDGET.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 90           'manual run time in seconds
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

'This is the dialog box information/code
BeginDialog DWP_budget_dialog, 0, 0, 426, 165, "DWP Budget Dialog"
  EditBox 60, 5, 45, 15, MAXIS_case_number
  EditBox 195, 5, 45, 15, ES_appointment_date
  EditBox 370, 5, 45, 15, ES_deadline_date
  EditBox 55, 25, 365, 15, income_info
  EditBox 55, 45, 365, 15, shelter_info
  EditBox 170, 65, 15, 15, personal_needs
  EditBox 75, 85, 160, 15, vendor_information
  EditBox 50, 105, 230, 15, other_notes
  EditBox 65, 125, 120, 15, months_eligible
  EditBox 260, 125, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 315, 145, 50, 15
    CancelButton 370, 145, 50, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 120, 10, 75, 10, "ES Appointment Date:  "
  Text 255, 10, 115, 10, "ES Deadline (10 Business Days):"
  Text 5, 30, 45, 10, "Income Info:"
  Text 5, 50, 45, 10, "Shelter Info: "
  Text 5, 70, 165, 10, "Personal Needs (Number of DWP HH Members):"
  Text 190, 65, 230, 20, "(This will multiply the number of eligible DWP household members by $70.00/person.)"
  Text 5, 90, 70, 10, "Vendor Information: "
  Text 5, 110, 45, 10, "Other Notes: "
  Text 5, 130, 60, 10, "Months Eligible: "
  Text 195, 130, 65, 10, "Worker Signature:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------
'Connects to BlueZone & grabbing the case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Displays the dialog
DO
	err_msg = ""
	Dialog DWP_budget_dialog
	cancel_confirmation
	IF MAXIS_case_number = "" OR (MAXIS_case_number <> "" AND IsNumeric(MAXIS_case_number) = False) THEN err_msg = err_msg & vbNewLine & "*Please enter a valid case number"
	If personal_needs = "" THEN err_msg = err_msg & vbNewLine & "*You must enter the number of DWP household members"
	IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "*You must sign your case note"
	IF err_msg <> "" THEN Msgbox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue"
LOOP UNTIL err_msg = ""

'Calculates personal needs info
personal_needs = "$" & personal_needs * 70

'Checks to make sure worker is not passworded out
CALL check_for_MAXIS(False)

'Writing to CASE/NOTE----------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
CALL write_variable_in_case_note("***DWP ES Referral and Budget Info***")
IF ES_appointment_date <> "" THEN CALL write_bullet_and_variable_in_case_note("ES Appointment Date", ES_appointment_date)
IF ES_deadline_date <> "" THEN CALL write_bullet_and_variable_in_case_note("ES Deadline Date", ES_deadline_date)
IF income_info <> "" THEN CALL write_bullet_and_variable_in_case_note("Income Info", income_info)
IF shelter_info <> "" THEN CALL write_bullet_and_variable_in_case_note("Shelter Info", shelter_info)
IF personal_needs <> "" THEN CALL write_bullet_and_variable_in_case_note("Personal Needs", personal_needs)
IF vendor_information <> "" THEN CALL write_bullet_and_variable_in_case_note("Vendor Information", vendor_information)
IF other_notes <> "" THEN CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
IF months_eligible <> "" THEN CALL write_bullet_and_variable_in_case_note("Months Eligible", months_eligible)
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

script_end_procedure("")
