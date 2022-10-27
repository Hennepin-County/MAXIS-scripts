'GATHERING STATS===========================================================================================
name_of_script = "NOTES - COUNTY PAID MEDI B.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 90         'manual run time in seconds
STATS_denomination = "C"       'C is for case
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

'===========================================================================================================CHANGELOG BLOCK
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("03/20/2020", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'=======================================================================================================END CHANGELOG BLOCK
'THE SCRIPT
'GRABBING THE CASE NUMBER, THE MEMB NUMBERS, AND THE FOOTER MONTH------------------------------------------------------------------------------------------------------------------------------------------------
EMConnect ""
Call maxis_case_number_finder(MAXIS_case_number)

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 211, 105, "COUNTY-PAID MEDI B"
  EditBox 65, 5, 45, 15, maxis_case_number
  EditBox 160, 5, 45, 15, start_date
  EditBox 65, 25, 45, 15, case_amount
  EditBox 160, 25, 45, 15, end_date
  EditBox 65, 45, 140, 15, other_notes
  EditBox 65, 65, 140, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 120, 85, 40, 15
    CancelButton 165, 85, 40, 15
  Text 120, 10, 40, 10, "Start Date:"
  Text 5, 10, 50, 10, "Case Number:"
  Text 5, 70, 60, 10, "Worker Signature:"
  Text 5, 50, 45, 10, "Other Notes:"
  Text 120, 30, 35, 10, "End Date:"
  Text 5, 30, 30, 10, "Amount:"
EndDialog

Do
    Do
        err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
		IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
		If trim(worker_signature) = "" then err_msg = err_msg & vbcr & "* Please ensure your case note is signed."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in
EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case THEN it gets identified, and will not be updated in MMIS

'----------------------------------------------------------------------------------------Gathering the member information

CALL navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv) 'navigating to stat memb to gather the ref number and name.
IF is_this_priv = TRUE THEN script_end_procedure("This case is privileged, the script will now end.")

client_array = "Select One:" & "|"

DO								'reads the reference number, last name, first name, and THEN puts it into a single string THEN into the array
EMReadscreen ref_nbr, 3, 4, 33
EMReadScreen access_denied_check, 13, 24, 2
'MsgBox access_denied_check
If access_denied_check = "ACCESS DENIED" Then
	PF10
	last_name = "UNABLE TO FIND"
	first_name = " - Access Denied"
	mid_initial = ""
Else
	EMReadscreen last_name, 25, 6, 30
	EMReadscreen first_name, 12, 6, 63
	EMReadscreen mid_initial, 1, 6, 79
	last_name = trim(replace(last_name, "_", "")) & " "
	first_name = trim(replace(first_name, "_", "")) & " "
	mid_initial = replace(mid_initial, "_", "")
End If
	EMReadscreen MEMB_number, 3, 4, 33
	EMReadscreen last_name, 25, 6, 30
	EMReadscreen first_name, 12, 6, 63
	EMReadscreen mid_initial, 1, 6, 79
    EMReadScreen client_DOB, 10, 8, 42
	EMReadscreen client_SSN, 11, 7, 42
	client_SSN = replace(client_SSN, " ", "")
	last_name = trim(replace(last_name, "_", "")) & " "
	first_name = trim(replace(first_name, "_", "")) & " "
	mid_initial = replace(mid_initial, "_", "")
	client_string = MEMB_number & last_name & first_name & client_SSN
	client_array = client_array & trim(client_string) & "|"

	transmit
	Emreadscreen edit_check, 7, 24, 2
LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

client_array = TRIM(client_array)
client_selection = split(client_array, "|")
CALL convert_array_to_droplist_items(client_selection, hh_member_dropdown)
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 171, 60, "HH Composition"
DropListBox 5, 20, 160, 15, hh_member_dropdown, ievs_member
  ButtonGroup ButtonPressed
    OkButton 70, 40, 45, 15
    CancelButton 120, 40, 45, 15
  Text 5, 5, 165, 10, "Please select the HH Member for the IEVS match:"
EndDialog

DO
    DO
       	err_msg = ""
       	Dialog Dialog1
       	cancel_without_confirmation
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
       LOOP UNTIL err_msg = ""
	CALL check_for_password_without_transmit(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = false

ievs_member = trim(ievs_member)
IEVS_ssn = right(ievs_member, 9)
IEVS_MEMB_number = left(ievs_member, 2)
'MsgBox IEVS_MEMB_number


'create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
Call create_TIKL("AVS 10-day follow up is required. If you do not have access to the AVS system, QI can assist you with the results. Contact your supervisor about your AVS access.", 10, date, False, TIKL_note_text)

start_a_blank_case_note
Call write_variable_in_case_note("***AVS Request Submitted - 10 day follow-up needed***")
Call write_variable_in_case_note(TIKL_note_text)
CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)


***MED B REIMB CHECK ISSUED*** Accounting has issued a COUNTY-PAID reimb
Check made payable to the client for the period 8/1/18 - 11/30/18
In the amount of $534.40.
.
This is per the Use Form email Request from the HSR.  Automatic monthly
COUNTY-PAID reimbursement checks will continue to be paid to the client
(or AREP/ALTP) on an ongoing basis as long as they are eligible for them.
..
Teams, please let Accounts. Payable know if this client ever becomes
Ineligible for COUNTY-PAID Med B reimbursements.  Thank you.
------.Barry Rausch, Accounts Payable 13-A MC 134


'NEED MORE SPECIFIC STOP AND START INFORMATION PULLED FROM DIALOG'
*MED B REIMB CHECK ISSUED* Accounting has issued a COUNTY-PAID
Reimb check made payable to the client.
.
This is per the Use Form email Request from the HSR.  Automatic monthly
COUNTY-PAID reimbursement checks will continue to be paid to the client
on an ongoing basis as long as they are eligible for them.
..
Teams, please let Accounts. Payable know if this client ever becomes
Ineligible for COUNTY-PAID Med B reimbursements.  Thank you.
******************************************************************************



script_end_procedure_with_error_report("Case note entered and copied to PERS please review case note to ensure accuracy.")
