'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MEMO - SIGNIFICANT CHANGE.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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

BeginDialog SigChange_Dialog, 0, 0, 291, 245, "Significant Change"
  EditBox 75, 0, 60, 15, case_number
  DropListBox 75, 20, 65, 15, "Select one..."+chr(9)+"Requested"+chr(9)+"Pending"+chr(9)+"Approved"+chr(9)+"Denied", Sig_change_status
  DropListBox 75, 35, 215, 15, "Select one..."+chr(9)+"Income didn't decline at least 50% in the benefit month. "+chr(9)+"Income change was due to an extra paycheck in the budget month."+chr(9)+"The decrease in income is due to a unit member on strike."+chr(9)+"Self Employment Income does not apply to Significant Change. "+chr(9)+"Significant Change was used twice in last 12 months.", Denial_reason
  EditBox 75, 60, 35, 15, Month_requested
  EditBox 190, 60, 35, 15, Last_month_used
  EditBox 75, 85, 210, 15, What_income
  EditBox 75, 105, 210, 15, Income_verified
  EditBox 75, 135, 210, 15, Verifs_needed
  EditBox 75, 155, 210, 15, Action_taken
  EditBox 75, 175, 110, 15, Worker_signature
  CheckBox 70, 200, 10, 10, "", Tikl_future_month
  CheckBox 190, 200, 10, 10, "", Spec_memo_denial
  ButtonGroup ButtonPressed
    OkButton 5, 220, 50, 15
    CancelButton 65, 220, 50, 15
  Text 5, 5, 45, 10, "Case Number"
  Text 5, 20, 65, 10, "Significant Change"
  Text 15, 35, 55, 10, "Reason if denied"
  Text 5, 60, 65, 10, "Month Requested"
  Text 125, 60, 55, 10, "Last Month Used"
  Text 5, 80, 65, 20, "What Income has decreased by 50%"
  Text 5, 105, 55, 20, "Income Change Verified?"
  Text 5, 140, 70, 10, "Verifications Needed"
  Text 5, 155, 50, 10, "Action Taken"
  Text 5, 170, 55, 20, "Worker Signature"
  Text 5, 200, 60, 10, "Tikl Future Month?"
  Text 105, 200, 85, 10, "Spec/Memo Denial Letter"
  Text 125, 220, 160, 20, "* See Combined Manual 0008.06.15 and TEMP  Manual TE02.13.11 for determining eligibility."
EndDialog

'THE SCRIPT------------------------------------------------------------------------------------------------------------------
'Connects to Bluezone
EMConnect "" 
'Brings Bluezone to foreground
EMFocus
'Password Check
call check_for_MAXIS(True)
'Searches for case number
call MAXIS_case_number_finder(case_number)
'This Do Loop makes sure a variable is filled for Sig Change status, month requested, and worker signature

DO
	DO
		DO
			DO
				DO 
					Dialog Sigchange_dialog
					cancel_confirmation
					call check_for_MAXIS (False)
					IF IsNumeric(case_number) = FALSE THEN MsgBox "You must type a valid numeric case number" 
				LOOP UNTIL IsNumeric(case_number) = TRUE 
				IF Sig_change_status = "Select one..." THEN MsgBox "You must select a Significant Change status type" 
			LOOP UNTIL Sig_change_status <> "Select one..." 
			IF Month_requested = "" THEN Msgbox "You must enter a month requested" 
		LOOP UNTIL Month_requested <> ""
		IF worker_signature = "" THEN MsgBox "You must sign your case note"
	LOOP UNTIL worker_signature <> ""
	If Sig_change_status = "Denied" AND Denial_reason = "Select one..." THEN
		Msgbox "You have selected Denied, you must select a denial reason"
	ELSE IF Sig_change_status <> "Denied" THEN EXIT DO
	END IF 
LOOP UNTIL Denial_reason <> "Select one..."
	
'TIKL to review/process sig change request for future month (check box selected)
If TIKL_future_month = checked THEN 
	'navigates to DAIL/WRIT 
	Call navigate_to_MAXIS_screen ("DAIL", "WRIT")	
	
	TIKL_date = dateadd("m", 1, date)		'Creates a TIKL_date variable with the current date + 1 month (to determine what the month will be next month)
	TIKL_date = datepart("m", TIKL_date) & "/01/" & datepart("yyyy", TIKL_date)		'Modifies the TIKL_date variable to reflect the month, the string "/01/", and the year from TIKL_date, which creates a TIKL date on the first of next month.
	
	'The following will generate a TIKL formatted date for 10 days from now.
	Call create_MAXIS_friendly_date(TIKL_date, 0, 5, 18) 'updates to first day of the next available month dateadd(m, 1)
	'Writes TIKL to worker
	Call write_variable_in_TIKL("A Significant Change was requested for this month. Please review and process")
	'Saves TIKL and enters out of TIKL function
	transmit
	PF3
END If

If Spec_memo_denial = checked THEN
	'Navigating to SPEC/MEMO
	call navigate_to_screen("SPEC", "MEMO")

	'This checks to make sure we've moved passed SELF.
	EMReadScreen SELF_check, 27, 2, 28
	If SELF_check = "Select Function Menu (SELF)" then script_end_procedure("An error has occurred preventing the script from moving past the SELF menu. Your case might be in background. Check for errors and try again.")

	'Creates a new MEMO. If it's unable the script will stop.
	PF5
	EMReadScreen memo_display_check, 12, 2, 33
	If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
	EMWriteScreen "x", 5, 10
	transmit

	'Writes the MEMO.
	EMSetCursor 3, 15
	call write_variable_in_SPEC_MEMO("************************************************************")
	call write_variable_in_SPEC_MEMO("Your request for Significant Change for the month of " & Month_requested & " has been received")
	call write_variable_in_SPEC_MEMO("Your household is not eligible to receive a significant change supplement for the month requested.")
	call write_variable_in_SPEC_MEMO("This is because " & Denial_reason)
	call write_variable_in_SPEC_MEMO("Please contact your worker if you have any questions")
	call write_variable_in_SPEC_MEMO("************************************************************")

	'Exits the MEMO
	PF4
END IF
	
'Starts the case note
call start_a_blank_case_note

'Writes the case note
call write_bullet_and_variable_in_CASE_NOTE ("Significant Change", Sig_change_status)
call write_bullet_and_variable_in_CASE_NOTE ("Month Requested", Month_requested)
call write_bullet_and_variable_in_CASE_NOTE ("Last Month Used", Last_month_used)
call write_bullet_and_variable_in_CASE_NOTE ("What Income has decreased by 50%", What_income)
call write_bullet_and_variable_in_CASE_NOTE ("Income Change Verified?", Income_verified)
call write_bullet_and_variable_in_CASE_NOTE ("Verifications Needed", Verifs_needed)
call write_bullet_and_variable_in_CASE_NOTE ("Action Taken", Action_taken)
call write_variable_in_CASE_NOTE (Worker_signature)

script_end_procedure("Success!")



