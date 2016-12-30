'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - SIGNIFICANT CHANGE.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 90                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
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

'DIALOGS---------------------------------
BeginDialog SigChange_Dialog, 0, 0, 291, 260, "Significant Change"
  EditBox 75, 5, 60, 15, MAXIS_case_number
  DropListBox 75, 25, 65, 15, "Select one..."+chr(9)+"Requested"+chr(9)+"Pending"+chr(9)+"Approved"+chr(9)+"Denied", Sig_change_status_dropdown
  DropListBox 75, 45, 215, 15, "Select one..."+chr(9)+"Income did not decrease enough to qualify."+chr(9)+"Income change was due to an extra paycheck in the budget month."+chr(9)+"The decrease in income is due to a unit member on strike."+chr(9)+"Self Employment Income does not apply to Significant Change."+chr(9)+"Significant Change was used twice in last 12 months.", Denial_reason_dropdown
  DropListBox 75, 75, 55, 15, "Select one..."+chr(9)+"January"+chr(9)+"February"+chr(9)+"March"+chr(9)+"April"+chr(9)+"May"+chr(9)+"June"+chr(9)+"July"+chr(9)+"August"+chr(9)+"September"+chr(9)+"October"+chr(9)+"November"+chr(9)+"December", Month_requested_dropdown
  EditBox 160, 75, 25, 15, Month_Requested_Year
  EditBox 75, 95, 35, 15, Last_month_used
  EditBox 75, 120, 210, 15, Income_decreased
  EditBox 75, 140, 210, 15, Income_verified
  EditBox 75, 165, 210, 15, Verifs_needed
  EditBox 75, 185, 210, 15, Action_taken
  CheckBox 5, 210, 110, 10, "Tikl Future Month Requested", Tikl_future_month_checkbox
  EditBox 215, 210, 70, 15, Worker_signature
  ButtonGroup ButtonPressed
    OkButton 175, 235, 50, 15
    CancelButton 235, 235, 50, 15
  Text 5, 10, 45, 10, "Case Number"
  Text 5, 30, 65, 10, "Significant Change"
  Text 5, 45, 60, 10, "Reason if denied"
  Text 15, 60, 245, 10, "*If Significant Change is denied a Denial Letter will be sent automatically"
  Text 5, 80, 65, 10, "Month Requested"
  Text 5, 100, 55, 10, "Last Month Used"
  Text 5, 120, 65, 15, "What Income has decreased?"
  Text 5, 140, 55, 15, "Income Change Verified?"
  Text 5, 170, 70, 10, "Verifications Needed"
  Text 5, 190, 50, 10, "Action Taken"
  Text 150, 215, 60, 10, "Worker Signature"
  Text 5, 230, 160, 20, "* See Combined Manual 0008.06.15 and TEMP  Manual TE02.13.11 for determining eligibility."
  Text 140, 80, 20, 10, "Year"
  Text 190, 80, 70, 10, "*Enter 4 digit year"
EndDialog


'THE SCRIPT------------------------------------------------------------------------------------------------------------------
EMConnect "" 'Connects to Bluezone
EMFocus 'Brings Bluezone to foreground

call check_for_MAXIS(True) 'Password Check- Script will shut down if passworded out

call MAXIS_case_number_finder(MAXIS_case_number) 'Searches for case number

'This is the new Do Loop process that makes mandatory fields in the dialog box
Do
	err_msg = ""
	Dialog SigChange_Dialog
	cancel_confirmation 'Are you sure you want to quit? message
	call check_for_MAXIS (False) 'Password check- If passworded out, dialog box wont close
	IF MAXIS_case_number = "" OR (MAXIS_case_number <> "" AND IsNumeric(MAXIS_case_number) = FALSE) THEN err_msg = err_msg & vbNewLine & "*Please enter a valid case number" 'Makes sure there is a numeric case number
	IF Sig_change_status_dropdown = "Select one..." THEN err_msg = err_msg & vbNewLine & "*You must select a Significant Change status type" 'Selecting the status of the sig change request is a mandatory field
	IF Sig_change_status_dropdown = "Denied" AND Denial_reason_dropdown = "Select one..." THEN err_msg = err_msg & vbNewLine & "*You have selected Denied, you must select a denial reason" 'If your status is Denied then you have to select a denial reason (this will pull into the spec/Memo denial letter)
	IF Sig_change_status_dropdown = "Denied" AND Denial_reason_dropdown <> "Select one..." AND Month_requested_dropdown = "Select one..." THEN err_msg = err_msg & vbNewLine & "*You must enter a month requested" 'I made the month requested a mandatory field only if it is denied because it pulls into the Spec/Memo, also clients do not always state the month they are requesting
	IF Month_requested_dropdown <> "Select one..." AND (Month_requested_year = "" OR IsNumeric(Month_Requested_Year) = FALSE) THEN err_msg = err_msg & vbNewLine & "*You must enter a valid year" 'This just makes you put in a numeric year if you select a month requested. Basicallly if you know the month then you should know the year
	IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "*You must sign your case note" 'Mandatory field
	IF err_msg <> "" THEN Msgbox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue" 'Msgbox starts out with Notice!!! then makes new line (this should give it a space it before the error messages because each message starts out with new line) and then adds a couple lines to space after the error messages before the saying that "Please resolve for script to continue" "
LOOP UNTIL err_msg = ""

'TIKL to review/process sig change request for future month (check box selected)
If TIKL_future_month_checkbox = checked THEN
	'navigates to DAIL/WRIT
	Call navigate_to_MAXIS_screen ("DAIL", "WRIT")

	TIKL_date = dateadd("m", 1, date)		'Creates a TIKL_date variable with the current date + 1 month (to determine what the month will be next month)
	TIKL_date = datepart("m", TIKL_date) & "/01/" & datepart("yyyy", TIKL_date)		'Modifies the TIKL_date variable to reflect the month, the string "/01/", and the year from TIKL_date, which creates a TIKL date on the first of next month.

	Call create_MAXIS_friendly_date(TIKL_date, 0, 5, 18) 'updates to first day of the next available month dateadd(m, 1)
	'Writes TIKL to worker
	Call write_variable_in_TIKL("A Significant Change was requested for this month. Please review and process")
	'Saves TIKL and enters out of TIKL function
	transmit
	PF3
END If

If Sig_change_status_dropdown = "Denied" THEN
	'Navigating to SPEC/MEMO
	call navigate_to_MAXIS_screen("SPEC", "MEMO")

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
	call write_variable_in_SPEC_MEMO("Your request for Significant Change for the month of " & Month_requested_dropdown & " " & Month_requested_year & " has been received.")
	call write_variable_in_SPEC_MEMO("Your household is not eligible to receive a significant change supplement for the month requested.")
	call write_variable_in_SPEC_MEMO("This is because " & Denial_reason_dropdown)
	call write_variable_in_SPEC_MEMO("Please contact your worker if you have any questions")
	call write_variable_in_SPEC_MEMO("************************************************************")

	'Exits the MEMO
	PF4
END If

'Starts the case note
call start_a_blank_case_note

'Writes the case note
call write_bullet_and_variable_in_CASE_NOTE ("Significant Change", Sig_change_status_dropdown)
IF Sig_change_status_dropdown = "Denied" THEN call write_bullet_and_variable_in_CASE_NOTE ("Denial Reason", Denial_reason_dropdown)
IF Month_requested_dropdown <> "Select one..." THEN call write_bullet_and_variable_in_CASE_NOTE ("Month Requested", Month_requested_dropdown & " " & Month_requested_year)
call write_bullet_and_variable_in_CASE_NOTE ("Last Month Used", Last_month_used)
call write_bullet_and_variable_in_CASE_NOTE ("What Income has decreased?", Income_decreased)
call write_bullet_and_variable_in_CASE_NOTE ("Income Change Verified?", Income_verified)
call write_bullet_and_variable_in_CASE_NOTE ("Verifications Needed", Verifs_needed)
call write_bullet_and_variable_in_CASE_NOTE ("Action Taken", Action_taken)
IF Tikl_future_month_checkbox = "1" THEN write_variable_in_case_note ("* Tikl set to review Significant Change for future month")
IF Sig_change_status_dropdown = "Denied" THEN write_variable_in_case_note ("* Denial letter sent via Spec/Memo")
call write_variable_in_CASE_NOTE (Worker_signature)

script_end_procedure("Success!")
