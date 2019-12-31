'Required for statistical purposes===============================================================================
name_of_script = "NOTES - OTHER BENEFITS REFERRAL.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 127         'manual run time in seconds
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
call changelog_update("05/01/2019", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'=======================================================================================================END CHANGELOG BLOCK
'THE SCRIPT
'GRABBING THE CASE NUMBER, THE MEMB NUMBERS, AND THE FOOTER MONTH------------------------------------------------------------------------------------------------------------------------------------------------
EMConnect ""
call maxis_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
memb_number = "01"

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
'DHS-2116-ENG Notice to Apply for Other Maintenance Benefits
BeginDialog Dialog1, 0, 0, 311, 205, "Other Maintenance Benefits"
  EditBox 60, 5, 45, 15, maxis_case_number
  EditBox 175, 5, 15, 15, memb_number
  EditBox 260, 5, 45, 15, other_elig_date
  CheckBox 15, 35, 80, 10, "Railroad Retirement", railroad_retirement_checkbox
  CheckBox 15, 45, 95, 10, "Unemployment Insurance", unemployment_insurance_checkbox
  CheckBox 15, 55, 125, 10, "Supplemental Security Income (SSI)", SSI_checkbox
  CheckBox 15, 65, 180, 10, "Retirement, Survivors, and Disability Income (RSDI)", RSDI_checkbox
  CheckBox 175, 35, 90, 10, "Worker's Compensation", worker_compensation_checkbox
  CheckBox 175, 45, 120, 10, "Veterans' Disability Benefits (VA)", VA_checkbox
  CheckBox 175, 55, 130, 10, "Other (please specify in other notes)", other_checkbox
  CheckBox 15, 90, 65, 10, "Medicare Buy-In", medi_checkbox
  EditBox 70, 100, 50, 15, ELIG_date
  EditBox 285, 100, 15, 15, ELIG_year
  CheckBox 10, 120, 235, 10, "Sent Notice to Apply for Other Maintenance Benefits DHS-2116-ENG", ECF_sent_checkbox
  CheckBox 10, 130, 285, 10, "Client will need to fill out an Interim Assistance Agreement DHS-1795 or DHS-1795A", IAA_needed
  EditBox 65, 145, 240, 15, action_taken
  EditBox 65, 165, 240, 15, other_notes
  EditBox 65, 185, 150, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 220, 185, 40, 15
    CancelButton 265, 185, 40, 15
  GroupBox 5, 80, 300, 40, "Medicare Buy-In only **   "
  Text 15, 105, 50, 10, "Eligibility Date:"
  Text 135, 105, 150, 10, "If ineligible the year that the client will be elig:"
  Text 15, 150, 50, 10, "Actions Taken:"
  Text 20, 170, 45, 10, "Other Notes:"
  Text 5, 190, 60, 10, "Worker signature:"
  Text 285, 90, 15, 10, "(YY)"
  Text 205, 10, 50, 10, "Eligibility Date:"
  Text 120, 10, 50, 10, "Memb Number:"
  GroupBox 5, 25, 300, 55, "Client may be able to get  benefits from the programs checked below:"
  Text 10, 10, 50, 10, "Case Number:"
EndDialog

Do
    Do
        err_msg = ""
		Dialog Dialog1
		cancel_confirmation
		IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
		If (isnumeric(memb_number) = False and len(MAXIS_case_number) > 2) then err_msg = err_msg & vbcr & "* Please enter a valid member number."
		If (ELIG_year <> "" and isnumeric(ELIG_year) = False and len(ELIG_year) > 2) then err_msg = err_msg & vbcr & "* Enter a valid 2 digit year for eligibility."
		IF (medi_checkbox = CHECKED and ELIG_date = "" and ELIG_year = "") THEN err_msg = err_msg & vbcr & "* Please advise of Medicare Buy-In eligibility date or the year the client will be eligible."
		IF (other_checkbox = CHECKED and other_notes = "") THEN err_msg = err_msg & vbcr & "* Please describe what benefits the client may be eligible for."
		If trim(worker_signature) = "" then err_msg = err_msg & vbcr & "* Please ensure your case note is signed."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 406, 130, "SSI Benefits Reminder"
  ButtonGroup ButtonPressed
    OkButton 295, 110, 50, 15
    CancelButton 350, 110, 50, 15
  Text 10, 5, 390, 15, "Applicants/participants who do not appear eligible to apply for Social Security benefits should not be referred to apply. For example:"
  Text 55, 25, 345, 20, "If an applicant/participant has a medical condition that will only last for three months and no other medical conditions, do not require them to apply for Social Security benefits."
  Text 55, 50, 345, 25, "If an applicant’s/participants disability makes cooperation with the Social Security application process impossible. In this case, refer applicant/participant to a skilled Social Security advocate or to your county or tribal social worker for assistance."
  Text 55, 80, 345, 25, "Do not require applicants/participants to reapply for Social Security benefits, which were previously denied unless there has been a change in their circumstances or the eligibility requirements of the benefit program."
EndDialog

IF SSI_checkbox = CHECKED THEN
    Do
        Do
            err_msg = ""
    		Dialog Dialog1
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
    	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
END IF

start_a_blank_case_note
IF medi_checkbox <> CHECKED and ELIG_year = "" THEN
	CALL write_variable_in_CASE_NOTE("Sent Notice to M" & memb_number & " to apply for Other Maintenance Benefits")
	IF ECF_sent_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Mailed DHS-2116-ENG Notice to Apply for Other Maintenance Benefits.")
	IF other_elig_date <> "" THEN CALL write_variable_in_case_note("* Client may be eligible for the following benefits as of " & other_elig_date & ":")
    IF RSDI_checkbox = CHECKED THEN CALL write_variable_in_case_note("- Retirement, Survivors, and Disability Income-RSDI")
    IF railroad_retirement_checkbox  = CHECKED THEN CALL write_variable_in_case_note("- Railroad Retirement")
    IF SSI_checkbox = CHECKED THEN CALL write_variable_in_case_note("- Supplemental Security Income-SSI")
    IF worker_compensation_checkbox = CHECKED THEN CALL write_variable_in_case_note("- Worker's Compensation")
    IF VA_checkbox = CHECKED THEN CALL write_variable_in_case_note("- Veterans' Disability Benefits (VA)")
    IF unemployment_insurance_checkbox = CHECKED THEN CALL write_variable_in_case_note("- Unemployment Insurance")
    IF other_checkbox = CHECKED THEN CALL write_variable_in_case_note("- Other")
    IF IAA_needed = CHECKED THEN CALL write_variable_in_case_note("* The client will need to fill out an Interim Assistance Agreement DHS-1795 or DHS-1795A.")
    IF IAA_needed = UNCHECKED THEN CALL write_variable_in_case_note("* The client will not need to fill out an Interim Assistance Agreement DHS-1795 or DHS-1795A.")
END IF

IF medi_checkbox = CHECKED and ELIG_date <> "" THEN
	due_date = dateadd("d", 30, date)
	Call write_variable_in_case_note("** Medicare Buy-in Referral mailed for M" & memb_number & " **")
	Call write_variable_in_case_note("* Client is eligible for the Medicare buy-in as of " & ELIG_date & ".")
	Call write_variable_in_case_note("* Proof due by " & due_date & " to apply.")
	Call write_variable_in_case_note("* Mailed DHS-3439-ENG MHCP Medicare Buy-In Referral Letter.")
	Call write_variable_in_case_note("* TIKL set to follow up.")
ELSEIF ELIG_year <> "" THEN
	Call write_variable_in_case_note("** Medicare Referral for M" & memb_number & " **")
	Call write_variable_in_case_note("* Client is not eligible for the Medicare buy-in. Enrollment is not until January 20" & ELIG_year & ", unable to apply until the enrollment time.")
	Call write_variable_in_case_note("* TIKL set to mail the Medicare Referral for November " & ELIG_year & ".")
END IF

CALL write_bullet_and_variable_in_case_note("Action Taken", action_taken)
CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)
PF3

'TIKLING
IF medi_checkbox = CHECKED and ELIG_date <> "" THEN
	CALL navigate_to_MAXIS_screen("DAIL","WRIT")
	call create_MAXIS_friendly_date(Due_date, 10, 5, 18)
	CALL write_variable_in_TIKL("Referral made for medicare, please check on proof of application filed. Due " & due_date & ".")
	PF3
END IF
IF ELIG_year <> "" THEN
	CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
	EMWriteScreen "11", 5, 18
	EMWriteScreen "01", 5, 21
	EMWriteScreen ELIG_year, 5, 24
	CALL write_variable_in_TIKL("Reminder to mail the Medicare Referral for November 20" & ELIG_year & ".")
	PF3
END IF


script_end_procedure_with_error_report("Case has been noted. Please remember to send forms out of ECF.")
