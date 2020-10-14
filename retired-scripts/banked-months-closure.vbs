'GATHERING STATS===========================================================================================
name_of_script = "NOTES - BANKED MONTHS CLOSURE.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 300
STATS_denominatinon = "M"
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
call changelog_update("12/29/2017", "Added Other closure reason to close banked months case. This will send the 'banked months notifier' WCOM to the client.", "Ilse Ferris, Hennepin County")
call changelog_update("11/17/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to MAXIS & case number
EMConnect ""
call maxis_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
member_number = "01"

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 141, 95, "Enter the case number & footer month/year"
  EditBox 75, 5, 45, 15, MAXIS_case_number
  EditBox 75, 25, 20, 15, MAXIS_footer_month
  EditBox 100, 25, 20, 15, MAXIS_footer_year
  EditBox 75, 45, 20, 15, member_number
  ButtonGroup ButtonPressed
    OkButton 15, 65, 50, 15
    CancelButton 70, 65, 50, 15
  Text 5, 30, 70, 10, "Closure month/year:"
  Text 20, 10, 55, 10, "Case Number:"
  Text 35, 50, 40, 10, "Member #:"
EndDialog

'the dialog
Do
	Do
  		err_msg = ""
  		Dialog Dialog1
  		If ButtonPressed = 0 then stopscript
  		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
  		If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) > 2 or len(MAXIS_footer_month) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer month."
  		If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) > 2 or len(MAXIS_footer_year) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid two-digit footer year."
		If IsNumeric(member_number) = False or len(member_number) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid two-digit footer year."
  		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

add_WCOM_check = 1		'defaulting the send WCOM to be checked since all cases should have a WCOM sent.

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 241, 205, "Banked Months Closure Dialog"
  DropListBox 105, 5, 70, 15, "Select one..."+chr(9)+"Non-Cooperation"+chr(9)+"All months used"+chr(9)+"Other closure", closure_reason
  EditBox 105, 25, 130, 15, closure_info
  EditBox 105, 45, 70, 15, counted_banked_months
  CheckBox 15, 80, 165, 10, "SNAP approved to continue for the next month?", next_month_checkbox
  CheckBox 15, 95, 155, 10, "Client does not meet an ABAWD exemption.", exemption_check
  CheckBox 15, 110, 170, 10, "Client does not meet ABAWD 2nd set criteria.", second_set_check
  CheckBox 15, 125, 120, 10, "Add applicable WCOM to notice?", add_WCOM_check
  EditBox 130, 145, 45, 15, next_ABAWD_month
  EditBox 65, 165, 170, 15, other_notes
  EditBox 65, 185, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 150, 185, 40, 15
    CancelButton 195, 185, 40, 15
  Text 5, 190, 60, 10, "Worker signature:"
  GroupBox 5, 65, 230, 75, "Check all that apply to be added to the case note:"
  Text 5, 10, 100, 10, "Banked months closure type:"
  Text 20, 170, 40, 10, "Other notes:"
  Text 10, 150, 120, 10, "When SNAP will be available again:"
  Text 5, 50, 100, 10, "Counted Banked Months here:"
  Text 5, 30, 100, 10, "If 'Other closure', explain here:"
EndDialog
'the dialog
Do
	Do
  		err_msg = ""
  		Dialog Dialog1
  		cancel_confirmation
		If closure_reason = "Select one..." then err_msg = err_msg & vbNewLine & "* Select a banked months closure type."
		If closure_reason = "Other closure" and trim(closure_info) = "" then err_msg = err_msg & vbNewLine & "* Enter the other closure reason(s)."
		If trim(counted_banked_months) = "" then err_msg = err_msg & vbNewLine & "* Enter all the banked months used."
		If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
  		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

check_for_MAXIS(False) 'in case users are in the ABAWD tracking record.
MAXIS_background_check	'Ensuring case is out of background

'Grabbing the client's 1st name
Call navigate_to_MAXIS_screen("STAT", "MEMB")
Call write_value_and_transmit(member_number, 20, 76)
EMReadScreen first_name, 12, 6, 63
first_name = replace(first_name, "_", "")
first_name = Trim(first_name)
Call fix_case_for_name(first_name)

'----------------------------------------------------------------------------------------------------NOTICE Coding
'This section will check for whether forms go to AREP and SWKR
call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
EMReadscreen forms_to_arep, 1, 10, 45
call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
EMReadscreen forms_to_swkr, 1, 15, 63

call navigate_to_MAXIS_screen("spec", "wcom")

EMWriteScreen MAXIS_footer_month, 3, 46
EMWriteScreen MAXIS_footer_year, 3, 51
transmit

DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
	EMReadScreen more_pages, 8, 18, 72
	IF more_pages = "MORE:  -" THEN PF7
LOOP until more_pages <> "MORE:  -"

read_row = 7
DO
	waiting_check = ""
	EMReadscreen prog_type, 2, read_row, 26
	EMReadscreen waiting_check, 7, read_row, 71 'finds if notice has been printed
	If waiting_check = "Waiting" and prog_type = "FS" THEN 'checking program type and if it's been printed
		EMSetcursor read_row, 13
		EMSendKey "x"
		Transmit
		pf9
		'The script is now on the recipient selection screen.  Mark all recipients that need NOTICES
		row = 4                             'Defining row and col for the search feature.
		col = 1
		EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
		IF row > 4 THEN  arep_row = row  'locating ALTREP location if it exists'
		row = 4                             'reset row and col for the next search
		col = 1
		EMSearch "SOCWKR", row, col
		IF row > 4 THEN  swkr_row = row     'Logs the row it found the SOCWKR string as swkr_row
		EMWriteScreen "x", 5, 12                                        'We always send notice to client
		IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 12     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
		IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 12     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
		transmit                                                        'Transmits to start the memo writing process'
		EMSetCursor 03, 15
		If closure_reason = "All months used" then CALL write_variable_in_SPEC_MEMO("You have been receiving SNAP banked months. Your SNAP is closing for using all available banked months. If you meet one of the exemptions listed above AND all other eligibility factors you may still be eligible for SNAP. Please call the EZ info line at 612-596-1300 if you have questions.")
		If closure_reason = "Non-Cooperation" then CALL write_variable_in_SPEC_MEMO("You have been receiving SNAP banked months. Your SNAP case is closing because" & first_name & " did not meet the requirements of working with Employment and Training. If you feel you have Good Cause for not cooperating with this requirement please contact your employment and training contact before your SNAP closes. If your SNAP closes for not cooperating with Employment and Training you will not be eligible for future banked months. If you meet an exemption listed above, AND all other eligibility factors you may be eligible for SNAP. If you have questions please call the EZ info line at 612-596-1300.")
		If closure_reason = "Other closure" then CALL write_variable_in_SPEC_MEMO("You have used all of your available ABAWD months. You may be eligible for SNAP banked months if you are cooperating with Employment Services. Please call the EZ info line at 612-596-1300 if you have questions.")
		PF4
		PF4
		PF3
		WCOM_count = WCOM_count + 1
		exit do
	ELSE
		read_row = read_row + 1
	END IF
	IF read_row = 18 THEN
		PF8          'Navigates to the next page of notices.  DO/LOOP until read_row = 18
		read_row = 7
	End if
LOOP until prog_type = "  "

'----------------------------------------------------------------------------------------------------All for the case note
'Informatoin to be added to the case note based on the option selected
If closure_reason = "All months used" then
	header = "Closed"
	closure_reason = "All available banked months used."
	wcom_info = "all eligible banked months have been used."
End if
If closure_reason = "Non-Cooperation" then
	header = "FORFEITED!"
	closure_reason = "Non-cooperation with E & T. Client cannot use banked months any longer even if all months have not been used."
	wcom_info = "banked months ending for SNAP E & T non-coop."
End if

If closure_reason = "Other closure" then
	header = "Closed"
	closure_reason = closure_info
	wcom_info = "banked months may be available."
End if

If next_month_checkbox = 1 then
	end_header = ", new SNAP elig app'd"
Else
	end_header = ""
End if
'----------------------------------------------------------------------------------------------------The case note
start_a_blank_CASE_NOTE
call write_variable_in_CASE_NOTE(header & " Banked Months eff. " & MAXIS_footer_month & "/" & MAXIS_footer_year & " MEMB " & member_number & end_header)
call write_bullet_and_variable_in_CASE_NOTE("Client closed due to", closure_reason)
call write_variable_in_CASE_NOTE("---Counted banked months: " & counted_banked_months & "---")
call write_variable_in_CASE_NOTE("---")
If exemption_check = 1 then call write_variable_in_CASE_NOTE("* Client does not meet an ABAWD exemption.")
If second_set_check = 1 then call write_variable_in_CASE_NOTE("* Client does not meet ABAWD 2nd set criteria.")
IF next_ABAWD_month <> "" then
	If second_set_check = 1 then
		call write_variable_in_CASE_NOTE("* Client will not be eligble for SNAP again until " & ABAWD_months & "unless the meets an exemption.")
	Else
		call write_variable_in_CASE_NOTE("* Client will not be eligble for SNAP again until " & ABAWD_months & "unless the meets an exemption, or qualifies for ABAWD 2nd set.")
	End if
End if
If next_month_checkbox = 1 then
	call write_variable_in_CASE_NOTE("---")
	call write_variable_in_CASE_NOTE("* SNAP has been approved to continue for the next month.")
End if
call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
If add_WCOM_check = 1 then call write_variable_in_CASE_NOTE("* Worker comments have been added to the notice re: " & wcom_info)
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(Worker_Signature)

script_end_procedure("")
