'LOADING GLOBAL VARIABLES
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("\\hcgg.fr.co.hennepin.mn.us\lobroot\HSPH\Team\Eligibility Support\Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - ACF REQUEST PENDING.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 0         	'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog acf_dialog, 0, 0, 281, 215, "ACF Request Pending"
  EditBox 50, 5, 55, 15, MAXIS_case_number
  EditBox 225, 5, 50, 15, date_request_sent
  EditBox 50, 30, 55, 15, monthly_rent
  EditBox 225, 30, 50, 15, damage_deposit
  EditBox 225, 55, 50, 15, amount_vendored
  EditBox 225, 80, 50, 15, account_balance
  EditBox 70, 105, 50, 15, earned_income
  EditBox 225, 105, 50, 15, unearned_income
  EditBox 70, 125, 50, 15, mfip
  EditBox 225, 125, 50, 15, dwp
  EditBox 115, 150, 160, 15, income_used_for
  EditBox 80, 170, 195, 15, reason_for_issuance
  EditBox 50, 195, 110, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 170, 195, 50, 15
    CancelButton 225, 195, 50, 15
  Text 5, 35, 45, 10, "Monthly rent: "
  Text 25, 85, 195, 10, "Balance in HCEA shelter account available towards rent/DD:"
  Text 80, 60, 140, 10, "Amount vendored to HCEA shelter account:"
  Text 15, 110, 55, 10, "Earned Income:"
  Text 155, 110, 60, 10, "Unearned Income:"
  Text 5, 155, 105, 10, "Income this month was used for:"
  Text 115, 10, 110, 10, "Request sent to HSS JW/GLA on:"
  Text 45, 125, 20, 10, "MFIP:"
  Text 195, 130, 20, 10, "DWP:"
  Text 20, 10, 25, 10, "Case #"
  Text 5, 175, 70, 10, "Reason for Issuance:"
  Text 160, 35, 55, 10, "Damage deposit:"
  Text 5, 200, 40, 10, "Other notes:"
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'autofilling the review_date variable with the current date
date_request_sent = date & ""

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog acf_dialog
		cancel_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		If Isdate(date_request_sent) = False then err_msg = err_msg & vbNewLine & "* Enter the date the request was sent."
		If monthly_rent = "" then err_msg = err_msg & vbNewLine & "* Enter the Monthly Rent amount."
		If damage_deposit = "" then err_msg = err_msg & vbNewLine & "* Enter Damage Deposit amount"
		If amount_vendored = "" then err_msg = err_msg & vbNewLine & "* Enter the Vendored amount"
		If account_balance = "" then err_msg = err_msg & vbNewLine & "* Enter the Account Balance amount"
		If earned_income = "" then err_msg = err_msg & vbNewLine & "* Enter the Earned Income amount."
		If unearned_income = "" then err_msg = err_msg & vbNewLine & "* Enter the Unearned Income amount."	
		If mfip = "" then err_msg = err_msg & vbNewLine & "* Enter the MFIP amount."
		If dwp = "" then err_msg = err_msg & vbNewLine & "* Enter the DWP amount."
		If income_used_for = "" then err_msg = err_msg & vbNewLine & "* Enter what the applicant income was used for."
		If reason_for_issuance = "" then err_msg = err_msg & vbNewLine & "* Enter the reason for Issuance."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & "(enter NA in all fields that do not apply)" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in					
		
'adding the case number 
back_to_self
EMWriteScreen "________", 18, 43
EMWriteScreen MAXIS_case_number, 18, 43
EMWriteScreen CM_mo, 20, 43	'entering current footer month/year
EMWriteScreen CM_yr, 20, 46

'The case note'
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("### All County Funds Request Pending ###")	
Call write_bullet_and_variable_in_CASE_NOTE("Request was sent to HSS JW and GLA on", date_request_sent)
Call write_bullet_and_variable_in_CASE_NOTE("Monthly Rent Amount", monthly_rent)
Call write_bullet_and_variable_in_CASE_NOTE("Damage Deposit", damage_deposit)
Call write_bullet_and_variable_in_CASE_NOTE("Amount Vendored", amount_vendored)
Call write_bullet_and_variable_in_CASE_NOTE("Account Balance", account_balance)
Call write_bullet_and_variable_in_CASE_NOTE("Earned Income", earned_income)
Call write_bullet_and_variable_in_CASE_NOTE("Unearned Income", unearned_income)
Call write_bullet_and_variable_in_CASE_NOTE("MFIP", mfip)
Call write_bullet_and_variable_in_CASE_NOTE("DWP", dwp)
Call write_bullet_and_variable_in_CASE_NOTE("Applicant income used for", income_used_for)
Call write_bullet_and_variable_in_CASE_NOTE("Reason for issuance", reason_for_issuance)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE ("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

script_end_procedure("")