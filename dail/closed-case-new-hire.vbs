'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.As such, it does NOT include protections to be ran independently.

'Required for statistical purposes===============================================================================
name_of_script = "DAIL - CLOSED CASE NEW HIRE.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 345         'manual run time in seconds
STATS_denomination = "C"       'C is for each MEMBER
'END OF stats block==============================================================================================

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
call changelog_update("01/12/2023", "Initial version.", "MiKayla Handley, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------------Script
EMConnect ""
EMSendKey "T"   'TYPES "T" TO BRING THE SELECTED MESSAGE TO THE TOP
transmit

EmReadScreen MAXIS_case_number, 8, 5, 73
MAXIS_case_number = trim(MAXIS_case_number)

'determining if the old message with the SSN functionality will be needed or not.
EMReadScreen HIRE_check, 11, 6, 37
If HIRE_check = "JOB DETAILS" then
    SSN_present = True
Else
    EmReadscreen fed_match, 4, 6, 20
    If left(fed_match, 4) = "NDNH" then SSN_present = False
    SSN_present = False
    EMReadScreen full_message, 60, 6, 20
    full_message = trim(full_message)
End if

If right(full_message, 2) = "00" then
    clear_DAIL = msgbox ("MEMB 00 message are system errors, and only need to be cleared in INFC. Would you like the script to clear this match in INFC?", vbQuestion, vbYesNo, "Non-Actionable DAIL found:")
    If clear_DAIL = vbNo then script_end_procedure("The script will not end. Process the MEMB 00 HIRE message in INFC manually.")
End if

'SELECTS THE DAIL MESSAGE AND READS THE RESPONSE
EMSendKey "X"
transmit

'Reading information fom the HIRE pop-up
'Date Hired and Employer Name
EMReadScreen new_hire_first_line, 61, 9, 5
new_hire_first_line = trim(new_hire_first_line)

EMReadScreen new_hire_second_line, 61, 10, 5
new_hire_second_line = trim(new_hire_second_line)
'MAXIS Name
EMReadScreen new_hire_third_line, 61, 11, 5
new_hire_third_line = trim(new_hire_third_line)
new_hire_third_line = replace(new_hire_third_line, ",", ", ")
'New Hire Name
EMReadScreen new_hire_fourth_line, 61, 12, 5
new_hire_fourth_line = trim(new_hire_fourth_line)
new_hire_fourth_line = replace(new_hire_fourth_line, ",", ", ")

row = 1 						'Now it's searching for info on the hire date as well as employer
col = 1
'Now it's searching for info on the hire date as well as employer
EMSearch "DATE HIRED", row, col
EMReadScreen date_hired, 10, row, col + 15
date_hired = trim(date_hired)
If date_hired = "  -  -  EM" OR date_hired = "UNKNOWN  E" then
    date_hired = CM_mo & "-" & current_day & "-" & CM_yr '??? Why is this code necessary?
Else
    Call ONLY_create_MAXIS_friendly_date(date_hired)
    month_hired = left(date_hired, 2)       'will be used to determine what dates to use on the JOBS panel
End if

EMSearch "EMPLOYER:", row, col
EMReadScreen employer, 25, row, col + 10
employer = TRIM(employer)
If SSN_present = True then
    EMReadScreen new_HIRE_SSN, 9, 9, 5
Else
    EmReadScreen HH_memb, 2, 9, 15
End if
PF3 ' to exit pop-up

call bring_correct_message_to_top
EMSendKey "H"
transmit

EMReadScreen case_status, 8, 8, 9
'TODO - Future Enhancement to determine if a case has been closed for at least 4 months.'
' Call write_value_and_transmit ("X", 4, 9)

PF3 		'Back to DAIL.

If case_status <> "INACTIVE" Then script_end_procedure("This case does does not appear to be INACTIVE. This script will now end.")


'----------------------------------------------------------------------------------------------------STAT Information
If clear_DAIL <> vbYes then

	Call write_value_and_transmit("S", 6, 3)
	'PRIV Handling
	EMReadScreen priv_check, 6, 24, 14              'If it can't get into the case then it's a priv case
	If priv_check = "PRIVIL" THEN script_end_procedure("This case is priviledged. The script will now end.")
	EMReadScreen stat_check, 4, 20, 21
	If stat_check <> "STAT" then script_end_procedure_with_error_report("Unable to get to stat due to an error screen. Clear the error screen and return to the DAIL. Then try the script again.")
	'GOING TO MEMB, NEED TO CHECK THE HH MEMBER
	Call write_value_and_transmit("MEMB", 20, 71)

	If SSN_present = True then
	    Do
	    	EMReadScreen MEMB_current, 1, 2, 73
	    	EMReadScreen MEMB_total, 1, 2, 78
	    	EMReadScreen MEMB_SSN, 11, 7, 42
	    	If new_HIRE_SSN = replace(MEMB_SSN, " ", "") then
	            exit do
	        Else
	    		transmit
	        End if
	    LOOP UNTIL (MEMB_current = MEMB_total) or (new_HIRE_SSN = replace(MEMB_SSN, " ", ""))
	    EMReadScreen HH_memb, 2, 4, 33
	Else
	    Call write_value_and_transmit(HH_memb, 20, 76) 'SSN_present = False information here
	    EmReadscreen MEMB_SSN, 11, 7, 42    'gathering the SSN
	    MEMB_SSN = replace(MEMB_SSN, " ", "")
	    If match_answer_droplist = "YES-INFC clear match" then PF3 'back to DAIL
	End if

	EMWriteScreen "JOBS", 20, 71
	Call write_value_and_transmit(HH_memb, 20, 76)

	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 276, 185, "NDNH Match Resolution Information"
	  DropListBox 170, 60, 95, 15, "Select One:"+chr(9)+"YES-No Further Action"+chr(9)+"NO-See Next Question", Emp_known_droplist
	  EditBox 10, 90, 260, 15, case_closure_detail
	  EditBox 10, 120, 260, 15, other_notes
	  EditBox 75, 140, 195, 15, worker_signature
	  ButtonGroup ButtonPressed
	    OkButton 165, 165, 50, 15
	    CancelButton 220, 165, 50, 15
		Text 10, 10, 260, 10, new_hire_first_line
  	  Text 10, 20, 260, 10, new_hire_second_line
  	  Text 10, 30, 260, 10, new_hire_third_line
  	  Text 10, 40, 260, 10, new_hire_fourth_line
	  Text 10, 65, 145, 10, "Was this employment known to the agency?"
	  Text 10, 80, 90, 10, "Case Closure Information:"
	  Text 10, 110, 40, 10, "Other notes:"
	  Text 10, 145, 60, 10, "Worker Signature:"
	  Text 10, 170, 115, 10, "Match Cleared Info in next Dialog."
	EndDialog

	'Show dialog
	DO
		DO
			err_msg = ""
			Dialog Dialog1
			cancel_confirmation
			IF Emp_known_droplist = "Select One:" THEN err_msg = err_msg & vbCr & "* You must select yes or no for was this employment known to the agency?"
			IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
		LOOP UNTIL err_msg = ""									'loops until all errors are resolved
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

	PF3
End If

'navigating to the INFC screens
EMSendKey "I"
transmit
'PRIV Handling
EMReadScreen priv_check, 6, 24, 14              'If it can't get into the case then it's a priv case
If priv_check = "PRIVIL" THEN script_end_procedure("This case is priviledged. The script will now end.")
If SSN_present = False then EmWriteScreen MEMB_SSN, 3, 63
Call write_value_and_transmit("HIRE", 20, 71)

'checking for IRS non-disclosure agreement.
EMReadScreen agreement_check, 9, 2, 24
IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")

row = 9
DO
	EMReadScreen case_number, 8, row, 5
	case_number = trim(case_number)
	IF case_number = MAXIS_case_number THEN
		EMReadScreen employer_match, 20, row, 36
		employer_match = trim(employer_match)
		IF trim(employer_match) = "" THEN script_end_procedure("An employer match could not be found. The script will now end.")
		IF employer_match = employer THEN
			EMReadScreen cleared_value, 1, row, 61
			IF cleared_value = " " THEN
				EmReadscreen date_of_hire, 8, row, 20
				EmReadscreen match_month, 5, row, 14
				info_confirmation = MsgBox("Press YES to confirm this is the match you wish to act on." & vbNewLine & "For the next match, press NO." & vbNewLine & vbNewLine & _
				"HIRE INFO from DAIL:" & vbNewLine & "   Employer: " & employer & vbNewLine & "   Date Hired: " & date_hired & vbNewLine & vbNewLine & _
				"INFC Information:" & vbNewLine & "   " & employer_match & vbNewLine & "     Case: " & case_number & vbNewLine & "     Hire Date: " & date_of_hire & vbNewLine & "     Month: " & match_month, vbYesNoCancel, "Please confirm this match")
				IF info_confirmation = vbCancel THEN script_end_procedure ("The script has ended. The match has not been acted on.")
				IF info_confirmation = vbYes THEN
					hire_match = TRUE
					match_row = row
					EXIT DO
				END IF
			END IF
		END IF
	END IF
	row = row + 1
	IF row = 19 THEN
		PF8
		EmReadscreen end_of_list, 9, 24, 14
		If end_of_list = "LAST PAGE" Then Exit Do
		row = 9
	END IF
LOOP UNTIL case_number = ""
IF hire_match <> TRUE THEN script_end_procedure("No pending HIRE match found for: " & employer & "." & vbcr & "Please review case for potential manual updates.")

If clear_DAIL = vbYes then
	'entering the INFC/HIRE match '
	Call write_value_and_transmit("U", match_row, 3)
	EMReadscreen panel_check, 4, 2, 49
	IF panel_check <> "NHMD" THEN msgbox "We did not enter to clear the match"
	EMWriteScreen "N", 16, 54
	EMWriteScreen "NA", 17, 54
	TRANSMIT 'enters the information then a warning message comes up WARNING: ARE YOU SURE YOU WANT TO UPDATE? PF3 TO CANCEL OR TRANSMIT TO UPDATE '
	TRANSMIT 'this confirms the cleared status'
	PF3
	EMReadscreen cleared_confirmation, 1, match_row, 61
	IF cleared_confirmation = "" THEN MsgBox "the match did not appear to clear"
	PF3' this takes us back to DAIL/DAIL
Else

	If Emp_known_droplist = "NO-See Next Question" Then Action_taken_droplist = "NA-No Action Taken"
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 276, 205, "NDNH Match Resolution Information"
	  Text 170, 65, 95, 10, Emp_known_droplist
	  If Emp_known_droplist = "NO-See Next Question" Then DropListBox 170, 80, 95, 15, "Select One:"+chr(9)+"NA-No Action Taken"+chr(9)+"BR-Benefits Reduced"+chr(9)+"CC-Case Closed", Action_taken_droplist
	  If Emp_known_droplist = "YES-No Further Action" Then Text 170, 85, 95, 20, "Employment Known, no additional action detail"
	  EditBox 10, 110, 260, 15, case_closure_detail
	  EditBox 10, 140, 260, 15, other_notes
	  EditBox 75, 160, 195, 15, worker_signature
	  ButtonGroup ButtonPressed
	    OkButton 165, 185, 50, 15
	    CancelButton 220, 185, 50, 15
	  Text 10, 10, 260, 10, new_hire_first_line
	  Text 10, 20, 260, 10, new_hire_second_line
	  Text 10, 30, 260, 10, new_hire_third_line
	  Text 10, 40, 260, 10, new_hire_fourth_line
	  Text 10, 65, 145, 10, "Was this employment known to the agency?"
	  Text 10, 85, 155, 10, "If unknown: what action was taken by agency?"
	  Text 10, 100, 90, 10, "Case Closure Information:"
	  Text 10, 130, 40, 10, "Other notes:"
	  Text 10, 165, 60, 10, "Worker Signature:"
	EndDialog

	'Show dialog
	DO
		DO
			err_msg = ""
			Dialog Dialog1
			cancel_confirmation
			IF Emp_known_droplist = "Select One:" THEN err_msg = err_msg & vbCr & "* You must select yes or no for was this employment known to the agency?"
			IF Emp_known_droplist = "NO-See Next Question" AND Action_taken_droplist = "Select One:" THEN err_msg = err_msg & vbCr & "* You must select an action taken."
	   		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
		LOOP UNTIL err_msg = ""									'loops until all errors are resolved
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

	'entering the INFC/HIRE match '
	EMWriteScreen "U", match_row, 3
	transmit
	EMReadscreen panel_check, 4, 2, 49
	IF panel_check <> "NHMD" THEN msgbox "We did not enter to clear the match"
	IF Emp_known_droplist = "NO-See Next Question" THEN EMWriteScreen "N", 16, 54
	IF Emp_known_droplist = "YES-No Further Action" THEN EMWriteScreen "Y", 16, 54
	IF Action_taken_droplist = "NA-No Action Taken" THEN EMWriteScreen "NA", 17, 54
	IF Action_taken_droplist = "BR-Benefits Reduced" THEN EMWriteScreen "BR", 17, 54
	IF Action_taken_droplist = "CC-Case Closed" THEN EMWriteScreen "CC", 17, 54
	TRANSMIT 'enters the information then a warning message comes up WARNING: ARE YOU SURE YOU WANT TO UPDATE? PF3 TO CANCEL OR TRANSMIT TO UPDATE '
	TRANSMIT 'this confirms the cleared status'
	PF3
	EMReadscreen cleared_confirmation, 1, match_row, 61
	IF cleared_confirmation = "" THEN MsgBox "the match did not appear to clear"
	PF3' this takes us back to DAIL/DAIL

End If

Call start_a_blank_CASE_NOTE

CALL write_variable_in_case_note("-NDNH Match for (M" & HH_memb & ") INFC cleared: Reported-")
CALL write_variable_in_case_note("DATE HIRED: " & date_hired)
CALL write_variable_in_case_note("EMPLOYER: " & employer)
CALL write_variable_in_case_note(new_hire_third_line)
CALL write_variable_in_case_note(new_hire_fourth_line)
IF Emp_known_droplist = "NO-See Next Question" THEN CALL write_variable_in_case_note("This employment was NOT known to the agency.")
IF Emp_known_droplist = "YES-No Further Action" THEN CALL write_variable_in_case_note("This employment was known to the agency.")
CALL write_variable_in_case_note("---")
IF Action_taken_droplist = "NA-No Action Taken" THEN CALL write_variable_in_case_note("* No futher action taken on this match at this time")
IF Action_taken_droplist = "BR-Benefits Reduced" THEN CALL write_variable_in_case_note("* Action taken: Benefits Reduced")
IF Action_taken_droplist = "CC-Case Closed" THEN CALL write_variable_in_case_note("* Action taken: Case Closed (allowing for 10 day cutoff if applicable)")

CALL write_bullet_and_variable_in_case_note("Case Closure Info", case_closure_detail)
CALL write_bullet_and_variable_in_case_note("Other notes", other_notes)
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)
closing_message = "Success! The NDNH HIRE message has been cleared. Please start overpayment process if necessary."
script_end_procedure_with_error_report(closing_message)

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------01/12/2023
'--Tab orders reviewed & confirmed----------------------------------------------01/12/2023
'--Mandatory fields all present & Reviewed--------------------------------------01/12/2023
'--All variables in dialog match mandatory fields-------------------------------01/12/2023
'Review dialog names for content and content fit in dialog----------------------01/12/2023
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------01/12/2023
'--CASE:NOTE Header doesn't look funky------------------------------------------01/12/2023
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------01/12/2023
'--write_variable_in_CASE_NOTE function:
'    confirm that proper punctuation is used -----------------------------------01/12/2023
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------N/A - Dail Scrubber
'--MAXIS_background_check reviewed (if applicable)------------------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------01/12/2023
'--Out-of-County handling reviewed----------------------------------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------01/12/2023
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------01/12/2023
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------01/12/2023
'--Incrementors reviewed (if necessary)-----------------------------------------N/A
'--Denomination reviewed -------------------------------------------------------01/12/2023
'--Script name reviewed---------------------------------------------------------01/12/2023
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------01/12/2023
'--comment Code-----------------------------------------------------------------01/12/2023
'--Update Changelog for release/update------------------------------------------N/A
'--Remove testing message boxes-------------------------------------------------01/12/2023
'--Remove testing code/unnecessary code-----------------------------------------01/12/2023
'--Review/update SharePoint instructions----------------------------------------N/A
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------N/A
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------N/A
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------							Worker Specific functionality for KW
