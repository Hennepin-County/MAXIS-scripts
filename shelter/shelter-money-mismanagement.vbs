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
closing_message = "Money mismanagement case note entered please follow all next steps to assist the resident."
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

when_contact_was_made = date
date_requested = date & ""
income_checkbox = checked
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 241, 205, " Money Mismanagement "
  EditBox 60, 5, 45, 15, maxis_case_number
  EditBox 145, 5, 70, 15, phone_number
  ComboBox 60, 25, 45, 15, "Phone call"+chr(9)+"Voicemail"+chr(9)+"Email"+chr(9)+"Office visit"+chr(9)+"Letter", contact_type
  DropListBox 115, 25, 20, 12, "to"+chr(9)+"from", contact_direction
  ComboBox 145, 25, 65, 15, "client"+chr(9)+"Other HH Memb"+chr(9)+"AREP", who_contacted
  EditBox 80, 45, 45, 15, date_requested
  DropListBox 60, 85, 70, 15, "Select One:"+chr(9)+"1st Instance"+chr(9)+"2nd Instance"+chr(9)+"Grant Management", occurrence_droplist
  EditBox 175, 100, 20, 15, first_month_grant_reduction
  EditBox 175, 120, 20, 15, first_occurrence_mm
  EditBox 175, 140, 20, 15, second_occurrence_mm
  EditBox 50, 165, 185, 15, comments_notes
  ButtonGroup ButtonPressed
    OkButton 130, 185, 50, 15
    CancelButton 185, 185, 50, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 115, 10, 25, 10, "Phone:"
  Text 5, 30, 50, 10, "Contact Type:"
  Text 5, 50, 75, 10, "Shelter requested on: "
  GroupBox 5, 70, 230, 90, "Money Mismanagement"
  Text 15, 90, 40, 10, "Occurrence:"
  Text 15, 105, 100, 10, "First month of grant reduction: "
  Text 15, 125, 145, 10, "First occurrence of money mismanagement: "
  Text 15, 145, 155, 10, "Second occurrence of money mismanagement:"
  Text 5, 170, 40, 10, "Comments:"
  Text 200, 105, 25, 10, "MM"
  Text 200, 125, 25, 10, "MM/YY"
  Text 200, 145, 25, 10, "MM/YY"
EndDialog

DO
	Do
		Dialog Dialog1
		cancel_confirmation
		Call validate_MAXIS_case_number(err_msg, "*")
		IF phone_number = "" then err_msg = err_msg & vbNewLine & "* Please enter the phone number."
		IF isdate(date_requested) = FALSE then err_msg = err_msg & vbNewLine & "* Please enter the date shelter was requested."
		IF occurrence_droplist = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select the occurrence of money mismanagement."
		IF first_month_grant_reduction = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the first month of grant reduction(MM)."
		IF first_occurrence_mm = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the first occurrence of money mismanagement (MM/YY)."
	LOOP UNTIL err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'1st Instance of Money Mismanagement = occurrence_MM1--------SAVE for ENHANCEMNT
'2nd Instance of Money Mismanagement = occurrence_MM2
'Grant Management = occurrence_MM3

back_to_self
EMWriteScreen "________", 18, 43
EMWriteScreen MAXIS_case_number, 18, 43
EMWriteScreen CM_mo, 20, 43	'entering current footer month/year
EMWriteScreen CM_yr, 20, 46

months_variable = CM_mo 'this is only use when they are on their last instance'
IF CM_MO = "01" THEN months_variable = "January, February, March"
IF CM_MO = "02" THEN months_variable = "February, March, April"
IF CM_MO = "03" THEN months_variable = "March, April, May"
IF CM_MO = "04" THEN months_variable = "April, May, June"
IF CM_MO = "05" THEN months_variable = "May, June, July"
IF CM_MO = "06" THEN months_variable = "June, July, August"
IF CM_MO = "07" THEN months_variable = "July, August, September"
IF CM_MO = "08" THEN months_variable = "August, September, October"
IF CM_MO = "09" THEN months_variable = "September, October, November"
IF CM_MO = "10" THEN months_variable = "October, November, December"
IF CM_MO = "11" THEN months_variable = "November, December, January"
IF CM_MO = "12" THEN months_variable = "December, January, February"

'----------------------------------------------------------------------------------------------------CASENOTE
start_a_blank_case_note
CALL write_variable_in_CASE_NOTE("### Grant reduction - " & Occurrence_droplist & " ###")
CALL write_variable_in_CASE_NOTE("* Contacted " & who_contacted & "on " & when_contact_was_made & " by " & contact_type & " " & contact_direction & " "& phone_number & " ")
CALL write_variable_in_CASE_NOTE("* Client requested shelter on " & date_requested & " and all GA/SSI is gone." )
CALL write_variable_in_CASE_NOTE("* 1st month of grant reduction: " & first_month_grant_reduction)
Call write_bullet_and_variable_in_CASE_NOTE("Comments", Comments_notes)
IF Occurrence_droplist = "2nd Instance" THEN CALL write_variable_in_CASE_NOTE("* 1st money mismanagement was " & first_occurrence_mm)
IF Occurrence_droplist = "Grant Management" THEN
    CALL write_variable_in_CASE_NOTE("*** Grant reduction 3 MONTHS/Grant Management/even if client is no longer in shelter ***")
    CALL write_variable_in_CASE_NOTE("* 1st money mismanagement was: " & first_occurrence_mm )
    CALL write_variable_in_CASE_NOTE("* Second money mismanagement was: " & second_occurrence_mm)
    CALL write_variable_in_CASE_NOTE("* No matter where client lives the grant will be $97.00 for three months. ")
    CALL write_variable_in_CASE_NOTE("* Grant reduced to $97.00 effective: " & months_variable)
END IF
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

Call script_end_procedure_with_error_report(closing_message)
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
