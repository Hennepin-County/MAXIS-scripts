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

'Connects to BlueZone
EMConnect ""
'Calls a MAXIS case number
call MAXIS_case_number_finder(MAXIS_case_number)


'--------------------------------------------------------------------------------------------------THE MAIN DIALOG
'DHS-2116-ENG Notice to Apply for Other Maintenance Benefits

B. When not to refer
Applicants/participants who do not appear eligible to apply for Social Security benefits should not be referred to apply. For example:
 If an applicant/participant has a medical condition that will only last for three months and no other medical conditions, do not require them to apply for Social Security benefits.
 If an applicant’s/participants disability makes cooperation with the Social Security application process impossible. In this case, refer applicant/participant to a skilled Social Security advocate or to your county or tribal social worker for assistance.
 Do not require applicants/participants to reapply for Social Security benefits, which were previously denied unless there has
'Notice to clients informing them of the requirement to apply for other maintenance benefits for which they may be eligible.
The agency believes that you may be able to get cash benefits from the programs checked below:
Railroad Retirement
Supplemental Security Income (SSI)
Worker's Compensation
Unemployment Insurance
Retirement, Survivors, and Disability Income (RSDI)
Veterans' Disability Benefits (VA)
Other (describe)
You
will
will not
need to fill out an Interim Assistance Agreement (DHS-1795 or DHS-1795A).
BeginDialog other_bene_dialog, 0, 0, 311, 190, "Other Maintenance Benefits" & maxis_case_number
  CheckBox 15, 20, 65, 10, "Medicare Buy-in", medi_checkbox
  CheckBox 125, 20, 180, 10, "Retirement, Survivors, and Disability Income (RSDI)", RSDI_checkbox
  CheckBox 15, 30, 80, 10, "Railroad Retirement", railroad_retirement_checkbox
  CheckBox 125, 30, 125, 10, "Supplemental Security Income (SSI)", SSI_checkbox
  CheckBox 15, 40, 90, 10, "Worker's Compensation", worker_compensation_checkbox
  CheckBox 125, 40, 120, 10, "Veterans' Disability Benefits (VA)", VA_checkbox
  CheckBox 15, 50, 85, 10, "Other(please specify)", other_checkbox
  CheckBox 125, 50, 95, 10, "Unemployment Insurance", unemployment_insurance_checkbox
  EditBox 65, 70, 50, 15, ELIG_date
  EditBox 250, 70, 50, 15, ELIG_year
  CheckBox 10, 95, 235, 10, "Sent Notice to Apply for Other Maintenance Benefits - DHS-2116-ENG", ECF_sent_checkbox
  CheckBox 10, 110, 285, 10, "Client will need to fill out an Interim Assistance Agreement DHS-1795 or DHS-1795A", IAA_needed
  EditBox 70, 125, 230, 15, action_taken
  EditBox 70, 145, 230, 15, other_notes
  EditBox 70, 165, 140, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 215, 165, 40, 15
    CancelButton 260, 165, 40, 15
  GroupBox 5, 5, 300, 60, "Client may be able to get  benefits from the programs checked below:"
  Text 10, 75, 50, 10, "Eligibility Date:"
  Text 140, 70, 105, 20, "Medicare Buy-In only - If INELIG the year that client will be ELIG"
  Text 10, 130, 50, 10, "Actions Taken:"
  Text 10, 150, 45, 10, "Other Notes:"
  Text 5, 170, 60, 10, "Worker signature:"
EndDialog

EMConnect ""
EMWriteScreen "N", 6, 3         'Goes to Case Note - maintains tie with DAIL
TRANSMIT
'Starts a blank case note
PF9
EMReadScreen case_note_mode_check, 7, 20, 3
If case_note_mode_check <> "Mode: A" then script_end_procedure("You are not in a case note on edit mode. You might be in inquiry. Try the script again in production.")

Do
    Do
        err_msg = ""
		Dialog medi_dialog
		cancel_confirmation
        If (isnumeric(MAXIS_case_number) = False and len(MAXIS_case_number) <> 8) then err_msg = err_msg & vbcr & "* Enter a valid case number."
		If trim(worker_signature) = "" then err_msg = err_msg & vbcr & "* Please ensure your case note is signed."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

due_date = dateadd("d", 30, ELIG_date)

start_a_blank_case_note
CALL write_variable_in_CASE_NOTE("Sent Notice to Apply for Other Maintenance Benefits - DHS-2116-ENG ")
CALL write_variable_in_case_note("* Mailed DHS-2116-ENG Notice to Apply for Other Maintenance Benefits.")
CALL write_variable_in_case_note("* Client may be eligible for the following benefits:")
IF RSDI_checkbox = CHECKED THEN CALL write_variable_in_case_note("Retirement, Survivors, and Disability Income-RSDI")
IF railroad_retirement_checkbox  = CHECKED THEN CALL write_variable_in_case_note("Supplemental Security Income-SSI")
IF SSI_checkbox = CHECKED THEN CALL write_variable_in_case_note("Supplemental Security Income-SSI")
IF worker_compensation_checkbox = CHECKED THEN CALL write_variable_in_case_note("Worker's Compensation")
IF VA_checkbox = CHECKED THEN CALL write_variable_in_case_note("Veterans' Disability Benefits (VA)")
IF other_checkbox = CHECKED THEN CALL write_variable_in_case_note("Other")
IF unemployment_insurance_checkbox = CHECKED THEN CALL write_variable_in_case_note("Unemployment Insurance")
IF IAA_needed = CHECKED THEN CALL write_variable_in_case_note("The client will need to fill out an Interim Assistance Agreement DHS-1795 or DHS-1795A.")
IF IAA_needed = UNCHECKED THEN CALL write_variable_in_case_note("The client will not need to fill out an Interim Assistance Agreement DHS-1795 or DHS-1795A.")
CALL write_variable_in_case_note("---")
IF medi_checkbox = CHECKED THEN
	Call write_variable_in_case_note("** Medicare Buy-in Referral mailed **")
	Call write_variable_in_case_note("Client is eligible for the Medicare buy-in as of " & ELIG_date & ". Proof due by " & due_date & "to apply.")
	Call write_variable_in_case_note("Mailed DHS-3439-ENG MHCP Medicare Buy-In Referral Letter - TIKL set to follow up.")
ELSE
	Call write_variable_in_case_note("** Medicare Referral **")
	Call write_variable_in_case_note("Client is not eligible for the Medicare buy-in. Enrollment is not until January " & ELIG_year & ", unable	to apply until the enrollment time.")
	Call write_variable_in_case_note("TIKL set to mail the Medicare Referral for November " & ELIG_year & ".")
END IF
IF ECF_sent_checkbox = CHECKED THEN CALL write_variable_in_case_note("* ECF reviewed and appropriate action taken")
CALL write_bullet_and_variable_in_case_note("Action Taken", action_taken)
CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)
PF3

'TIKLING
	IF TIKL_checkbox = checked THEN CALL navigate_to_MAXIS_screen("dail", "writ")
	'If worker checked to TIKL out, it goes to DAIL WRIT
	IF TIKL_checkbox = checked THEN
		CALL navigate_to_MAXIS_screen("DAIL","WRIT")
		CALL create_MAXIS_friendly_date(date, 10, 5, 18)
		EMSetCursor 9, 3
		IF medi_checkbox = CHECKED THEN
			EMSendKey "Medicare Referral made, please check on proof of application filed."
		ELSE
			EMSendKey "TIKL set to mail the Medicare Referral for November " & ELIG_year & "."
		END IF
	END IF

script_end_procedure_with_error_report(DAIL_type & vbcr &  first_line & vbcr & " DAIL has been case noted. Please remember to send forms out of ECF.")
