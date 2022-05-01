'GATHERING STATS===========================================================================================
name_of_script = "NOTES - SHELTER-MONEY MISMANAGEMENT.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 180
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

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
CALL changelog_update("04/12/2022", "Elimination of Self-Pay: Removal of mention from scripts.", "MiKayla Handley, Hennepin County")
call changelog_update("06/26/2017", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'--------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

when_contact_was_made = date & ""
date_requested = date & ""
income_checkbox = checked
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 246, 235, "Money Mismanagement"
  EditBox 60, 5, 45, 15, maxis_case_number
  EditBox 150, 5, 70, 15, phone_number
  ComboBox 60, 25, 55, 15, "Phone call"+chr(9)+"Voicemail"+chr(9)+"Email"+chr(9)+"Office visit"+chr(9)+"Letter", contact_type
  DropListBox 120, 25, 25, 15, "to"+chr(9)+"from", contact_direction
  ComboBox 155, 25, 65, 15, "client"+chr(9)+"Other HH Memb"+chr(9)+"AREP", who_contacted
  EditBox 80, 45, 45, 15, date_requested
  EditBox 80, 65, 45, 15, when_contact_was_made
  DropListBox 60, 100, 80, 15, "Select One:"+chr(9)+"1st Instance"+chr(9)+"2nd Instance"+chr(9)+"Grant Management", occurrence_droplist
  EditBox 180, 115, 25, 15, first_occurrence_mm
  EditBox 180, 135, 25, 15, second_occurrence_mm
  DropListBox 180, 155, 55, 45, "Select One:"+chr(9)+"January"+chr(9)+"February"+chr(9)+"March"+chr(9)+"April"+chr(9)+"May"+chr(9)+"June"+chr(9)+"July"+chr(9)+"August"+chr(9)+"September"+chr(9)+"October"+chr(9)+"November"+chr(9)+"December", first_month_of_grant_reduction
  EditBox 180, 170, 30, 15, grant_reduction_amount
  EditBox 50, 195, 190, 15, comments_notes
  EditBox 50, 215, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 135, 215, 50, 15
    CancelButton 190, 215, 50, 15
  Text 120, 10, 25, 10, "Phone:"
  Text 5, 30, 50, 10, "Contact Type:"
  Text 5, 70, 50, 10, "Contact date:"
  Text 5, 50, 75, 10, "Shelter requested on: "
  GroupBox 5, 85, 235, 105, "Money Mismanagement"
  Text 15, 105, 40, 10, "Occurrence:"
  Text 15, 120, 145, 10, "First occurrence of money mismanagement: "
  Text 210, 120, 25, 10, "MM/YY"
  Text 15, 140, 155, 10, "Second occurrence of money mismanagement:"
  Text 210, 140, 25, 10, "MM/YY"
  Text 15, 160, 55, 10, "Grant Reduction:"
  Text 125, 160, 45, 10, "Initial Month:"
  Text 125, 175, 50, 10, "Reduced to: $"
  Text 5, 200, 40, 10, "Comments:"
  Text 5, 10, 50, 10, "Case Number:"
  Text 5, 220, 40, 10, "Worker Sig:"
EndDialog

DO
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation

		phone_number = trim(phone_number)
		first_occurrence_mm = trim(first_occurrence_mm)
		second_occurrence_mm = trim(second_occurrence_mm)
		grant_reduction_amount = trim(grant_reduction_amount)
		comments_notes = trim(comments_notes)

		Call validate_MAXIS_case_number(err_msg, "*")
		IF phone_number = "" then err_msg = err_msg & vbNewLine & "* Please enter the phone number."
		IF isdate(when_contact_was_made) = FALSE then err_msg = err_msg & vbNewLine & "* Please enter the date contact was made."
		IF isdate(date_requested) = FALSE then err_msg = err_msg & vbNewLine & "* Please enter the date shelter was requested."
		IF occurrence_droplist = "Select One:" THEN
			err_msg = err_msg & vbNewLine & "* Please select the occurrence of money mismanagement."
		ElseIf occurrence_droplist = "1st Instance" Then
			IF first_occurrence_mm = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the month and year of the first instance of money mismanagement."
			If second_occurrence_mm <> "" Then err_msg = err_msg & vbNewLine & "* Since this is only the first month of grant mismanagment, no month should be listed in the second instance of grant mismanagement."
			If first_month_of_grant_reduction <> "Select One:" Then err_msg = err_msg & vbNewLine & "* Since this is not a 'Grant Management' action, a month for grant reduction should not be entered."
			If grant_reduction_amount <> "" Then err_msg = err_msg & vbNewLine & "* Since this is not a 'Grant Management' action, a grant reduction amount should not be entered."
		ElseIf occurrence_droplist = "2nd Instance" Then
			IF first_occurrence_mm = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the month and year of the first instance of money mismanagement."
			If second_occurrence_mm = "" Then err_msg = err_msg & vbNewLine & "* Please enter the month and year of the second instance of money mismanagement."
			If first_month_of_grant_reduction <> "Select One:" Then err_msg = err_msg & vbNewLine & "* Since this is not a 'Grant Management' action, a month for grant reduction should not be entered."
			If grant_reduction_amount <> "" Then err_msg = err_msg & vbNewLine & "* Since this is not a 'Grant Management' action, a grant reduction amount should not be entered."
		ElseIf occurrence_droplist = "Grant Management" Then
			IF first_occurrence_mm = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the month and year of the first instance of money mismanagement."
			If second_occurrence_mm = "" Then err_msg = err_msg & vbNewLine & "* Please enter the month and year of the second instance of money mismanagement."
			If first_month_of_grant_reduction = "Select One:" Then err_msg = err_msg & vbNewLine & "* Indicate which month will first have a grant reduction to follow the 'Grant Management' process."
			If grant_reduction_amount = "" Then err_msg = err_msg & vbNewLine & "* Indicate the amount the grant will be reduced to for the 'Grant Management' process."
		End If
		If worker_signature = "" Then err_msg = err_msg & vbNewLine & "* Please enter your worker signature."
		If err_msg <> "" Then MsgBox "******  NOTICE  ******" & vbNewLine & "Resolve to continue:" & vbNewLine & err_msg
	LOOP UNTIL err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'1st Instance of Money Mismanagement = occurrence_MM1--------SAVE for ENHANCEMNT
'2nd Instance of Money Mismanagement = occurrence_MM2
'Grant Management = occurrence_MM3

IF first_month_of_grant_reduction = "January" THEN months_variable = "January, February, March"
IF first_month_of_grant_reduction = "February" THEN months_variable = "February, March, April"
IF first_month_of_grant_reduction = "March" THEN months_variable = "March, April, May"
IF first_month_of_grant_reduction = "April" THEN months_variable = "April, May, June"
IF first_month_of_grant_reduction = "May" THEN months_variable = "May, June, July"
IF first_month_of_grant_reduction = "June" THEN months_variable = "June, July, August"
IF first_month_of_grant_reduction = "July" THEN months_variable = "July, August, September"
IF first_month_of_grant_reduction = "August" THEN months_variable = "August, September, October"
IF first_month_of_grant_reduction = "September" THEN months_variable = "September, October, November"
IF first_month_of_grant_reduction = "October" THEN months_variable = "October, November, December"
IF first_month_of_grant_reduction = "November" THEN months_variable = "November, December, January"
IF first_month_of_grant_reduction = "December" THEN months_variable = "December, January, February"

'----------------------------------------------------------------------------------------------------CASENOTE
start_a_blank_case_note
CALL write_variable_in_CASE_NOTE("### Money Mismanagement: " & occurrence_droplist & " ###")
CALL write_variable_in_CASE_NOTE("* Contacted " & who_contacted & " on " & when_contact_was_made & " by " & contact_type & " " & contact_direction & " "& phone_number & " ")
CALL write_variable_in_CASE_NOTE("* Client requested shelter on " & date_requested & " and all GA/SSI is gone." )
IF occurrence_droplist = "Grant Management" THEN
    CALL write_variable_in_CASE_NOTE("*** Grant reduction 3 MONTHS/Grant Management/even if client is no longer in shelter ***")
    CALL write_variable_in_CASE_NOTE("* 1st money mismanagement was: " & first_occurrence_mm )
    CALL write_variable_in_CASE_NOTE("* Second money mismanagement was: " & second_occurrence_mm)
    CALL write_variable_in_CASE_NOTE("* No matter where client lives the grant will be $" & grant_reduction_amount & " for three months.")
    CALL write_variable_in_CASE_NOTE("* Grant reduced to $" & grant_reduction_amount & " effective: " & months_variable)
ELSE
	If first_occurrence_mm <> "" Then CALL write_variable_in_CASE_NOTE("* 1st money mismanagement was: " & first_occurrence_mm )
	If second_occurrence_mm <> "" Then CALL write_variable_in_CASE_NOTE("* Second money mismanagement was: " & second_occurrence_mm)
END IF
Call write_bullet_and_variable_in_CASE_NOTE("Comments", Comments_notes)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

Call script_end_procedure_with_error_report("Money mismanagement case note entered please follow all next steps to assist the resident.")
'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------04/29/2022
'--Tab orders reviewed & confirmed----------------------------------------------04/29/2022
'--Mandatory fields all present & Reviewed--------------------------------------04/29/2022
'--All variables in dialog match mandatory fields-------------------------------04/29/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------04/29/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------04/29/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------04/29/2022
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------n/a
'--MAXIS_background_check reviewed (if applicable)------------------------------n/a
'--PRIV Case handling reviewed -------------------------------------------------n/a
'--Out-of-County handling reviewed----------------------------------------------n/a
'--script_end_procedures (w/ or w/o error messaging)----------------------------04/29/2022
'--BULK - review output of statistics and run time/count (if applicable)--------n/a
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------n/a
'--Incrementors reviewed (if necessary)-----------------------------------------04/29/2022
'--Denomination reviewed -------------------------------------------------------04/29/2022
'--Script name reviewed---------------------------------------------------------04/29/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------n/a

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub taks are complete-----------------------------------------04/29/2022
'--comment Code-----------------------------------------------------------------n/a
'--Update Changelog for release/update------------------------------------------04/29/2022
'--Remove testing message boxes-------------------------------------------------04/29/2022
'--Remove testing code/unnecessary code-----------------------------------------04/29/2022
'--Review/update SharePoint instructions----------------------------------------04/29/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------04/29/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------n/a
'--Complete misc. documentation (if applicable)---------------------------------04/29/2022
'--Update project team/issue contact (if applicable)----------------------------04/29/2022
