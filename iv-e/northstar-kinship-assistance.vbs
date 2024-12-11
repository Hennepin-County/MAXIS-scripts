'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - NORTHSTAR KINSHIP ASSISTANCE.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 0         	'manual run time in seconds
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
call changelog_update("09/09/2024", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
Call check_for_MAXIS(False)							'ensure we are in MAXIS
CALL MAXIS_case_number_finder(MAXIS_case_number)

TIKL_checkbox = checked								'default the checkbox to checked - TIKLing is needed in most cases.

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog dialog1, 0, 0, 321, 165, "Details about Northstar Kinship Assistance"
  EditBox 70, 10, 60, 15, MAXIS_case_number
  EditBox 250, 10, 60, 15, effective_date
  EditBox 70, 35, 150, 15, arep_entered
  EditBox 70, 55, 240, 15, memi_information
  EditBox 70, 75, 240, 15, unea_income
  CheckBox 70, 100, 170, 10, "Check here to TIKL to close when the child is 18.", TIKL_checkbox
  EditBox 70, 120, 240, 15, other_notes
  EditBox 70, 140, 125, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 205, 140, 50, 15
    CancelButton 260, 140, 50, 15
  Text 20, 15, 45, 10, "Case number:"
  Text 140, 15, 110, 10, "Northstar Kinship Effective Date:"
  Text 45, 40, 25, 10, "AREP:"
  Text 5, 60, 65, 10, "MEMI Information:"
  Text 15, 80, 55, 10, "UNEA Income:"
  Text 25, 125, 45, 10, "Other Notes:"
  Text 10, 145, 60, 10, "Worker signature:"
EndDialog

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog dialog1
        cancel_confirmation

		memi_information = trim(memi_information)
		worker_signature = trim(worker_signature)

		call validate_MAXIS_case_number(err_msg, "*")																'Mandatory fields: Case number, effective date, MEMI information, Worker Signature
		If IsDate(effective_date) = False Then err_msg = err_msg & vbNewLine & "* Enter a valid effective date."
		If memi_information = "" Then err_msg = err_msg & vbNewLine & "* Enter information from the MEMI panel."
		If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."

		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in
Call check_for_MAXIS(False)
end_msg = "Information about Northstar Kinship Assistance has been entered into CASE/NOTE."							'Start of the closing message information

TIKL_note_text = ""					'defaulting this to blank - it will indicate if the TIKL is set or not.
If TIKL_checkbox = checked Then
	Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_PRIV)											'Going to get the member DOB to determine the date of the TIKL
	If is_this_PRIV = False Then
		EMReadScreen memb_01_DOB, 10, 8, 42								'reading the DOB of MEMB 01 and making it a date
		memb_01_DOB = replace(memb_01_DOB, " ", "/")
		memb_01_DOB = DateAdd("d", 0, memb_01_DOB)

		dob_for_18_today = DateAdd("yyyy", -18, date)					'setting the date for the birthday of someone turing 18 today.

		If DateDiff("d", memb_01_DOB, dob_for_18_today) >=0 Then		'If the DOB is before or the same as someone who turns 18 today, the child is already 18
			end_msg = end_msg & vbCr & vbCr & "TIKL for 18 years could not be entered because it appears MEMB 01 is already 18."
		Else
			eighteenth_birthday = DateAdd("yyyy", 18, memb_01_DOB)		'creating a variable for the child's 18th birthday
			month_turns_18 = DatePart("m", eighteenth_birthday)
			year_turns_18 = DatePart("yyyy", eighteenth_birthday)
			TIKL_date = month_turns_18 & "/1/" & year_turns_18			'Setting a variable for the date the TIKL should be set
			TIKL_date = DateAdd("d", 0, TIKL_date)

			If DateDiff("d", date, TIKL_date) > 0 Then					'If the TIKL is after today, it will be set.
				Call create_TIKL("Review MA-25X eligibility as Member 01 is turning 18 this month.", 0, TIKL_date, False, TIKL_note_text)
				end_msg = end_msg & vbCr & vbCr & "TIKL created for " & TIKL_date & " to remind to close MA-25x when the resident turns 18."
			Else														'If the TIKL date is for today or before today, it cannot be set.
				end_msg = end_msg & vbCr & vbCr & "TIKL for 18 years could not be entered because the TIKL date (" & TIKL_date & ") is not in the future."
			End If
		End If
	Else
		end_msg = end_msg & vbCr & vbCr & "TIKL for 18 years could not be entered because the case is Privileged."
	End If
End If
effective_month = DatePart("m", effective_date)			'creating an effective month and year for the CASE/NOTE
effective_year = DatePart("yyyy", effective_date)

'The case note
start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
Call write_variable_in_CASE_NOTE("APPROVAL - Northstar Kinship ELIGIBLE eff " & effective_month & "/" & effective_year & " - Ongoing")
Call write_variable_in_CASE_NOTE("Effective Date: " & effective_date)
Call write_variable_in_CASE_NOTE("Child is ELIGIBLE for MA.")
Call write_variable_in_CASE_NOTE("  Elig Type: 25")
Call write_variable_in_CASE_NOTE("- No Health Care Reviews for Northstar Kinship Assistance Cases")
If TIKL_note_text <> "" Then Call write_variable_in_CASE_NOTE("  -TIKL set to close MA-25x when the child turns 18.")
Call write_variable_in_CASE_NOTE ("---")
Call write_bullet_and_variable_in_CASE_NOTE("AREP", arep_entered)
Call write_bullet_and_variable_in_CASE_NOTE("MEMI", memi_information)
Call write_bullet_and_variable_in_CASE_NOTE("UNEA", unea_income)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE ("---")
Call write_variable_in_CASE_NOTE (worker_signature)

Call script_end_procedure(end_msg)

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 05/23/2024
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------09/10/2024
'--Tab orders reviewed & confirmed----------------------------------------------09/10/2024
'--Mandatory fields all present & Reviewed--------------------------------------09/10/2024
'--All variables in dialog match mandatory fields-------------------------------09/10/2024
'Review dialog names for content and content fit in dialog----------------------09/10/2024
'--FIRST DIALOG--NEW EFF 5/23/2024----------------------------------------------N/A - Aligning with the other scripts in this category
'--Include script category and name somewhere on first dialog-------------------
'--Create a button to reference instructions------------------------------------
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------09/10/2024
'--CASE:NOTE Header doesn't look funky------------------------------------------09/10/2024
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------09/10/2024
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------09/10/2024
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------09/10/2024
'--MAXIS_background_check reviewed (if applicable)------------------------------09/10/2024
'--PRIV Case handling reviewed -------------------------------------------------09/10/2024					PRIV and Out of County are limited for IV-E cases
'--Out-of-County handling reviewed----------------------------------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------09/10/2024
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------09/10/2024
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------N/A							No stats in any of the IV-E cases.
'--Incrementors reviewed (if necessary)-----------------------------------------N/A
'--Denomination reviewed -------------------------------------------------------N/A
'--Script name reviewed---------------------------------------------------------09/10/2024
'--BULK - remove 1 incrementor at end of script reviewed------------------------

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------09/10/2024
'--comment Code-----------------------------------------------------------------
'--Update Changelog for release/update------------------------------------------09/10/2024
'--Remove testing message boxes-------------------------------------------------09/10/2024
'--Remove testing code/unnecessary code-----------------------------------------09/10/2024
'--Review/update SharePoint instructions----------------------------------------09/10/2024
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------09/10/2024
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------09/11/2024
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------09/10/2024
