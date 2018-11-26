'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - SHELTER-SHELTER INTERVIEW.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 0         	'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

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

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog Shelter_interview, 0, 0, 311, 325, "Shelter Interview: Do no release funds, family in shelter."
  EditBox 70, 10, 45, 15, MAXIS_case_number
  DropListBox 210, 10, 90, 15, "Select one..."+chr(9)+"DWP"+chr(9)+"MFIP", cash_type
  EditBox 240, 40, 60, 15, one_time_issuance
  EditBox 95, 65, 205, 15, other_income
  EditBox 95, 85, 205, 15, money_mismanagement
  EditBox 95, 105, 205, 15, reason_homeless
  EditBox 95, 125, 205, 15, barriers_housing
  EditBox 95, 145, 205, 15, shelter_history
  EditBox 95, 165, 205, 15, social_worker
  EditBox 95, 185, 205, 15, referrals_made
  EditBox 95, 205, 100, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 200, 205, 50, 15
    CancelButton 250, 205, 50, 15
  Text 45, 70, 45, 10, "Other income:"
  Text 10, 90, 80, 10, "Money mismanagement:"
  Text 40, 170, 50, 10, "Social worker:"
  Text 20, 130, 70, 10, "Barrier(s) to housing:"
  Text 50, 210, 40, 10, "Other notes: "
  Text 20, 260, 215, 15, "* Explained shelter policies and client options to shelter such as:   bus tickets, temporary housing, private shelters, etc."
  Text 25, 285, 265, 15, "* Client given family social services number (348-4111) to discuss any family issues/barriers."
  GroupBox 5, 230, 295, 75, "Additional text added to case note:"
  Text 155, 15, 50, 10, "Cash program:"
  Text 20, 15, 45, 10, "Case number:"
  Text 30, 190, 60, 10, "Referrals made to:"
  Text 35, 150, 55, 10, "Shelter history:"
  Text 20, 245, 225, 10, "* 100% of cash benefit to be issued to HCEA shelter account #52871."
  Text 5, 110, 90, 10, "Reason for homelessness:"
  Text 15, 45, 225, 10, "Amt issued to EBT as one-time only (10%) for PN ($20 med co-pays):"
  GroupBox 10, 30, 295, 30, "If MFIP recipient:"
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog Shelter_interview
		cancel_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		If cash_type = "Select one..." then err_msg = err_msg & vbNewLine & "* Select the family's cash program."      
        If cash_type = "MFIP" and one_time_issuance = "" then err_msg = err_msg & vbNewLine & "* Enter the amount to issue as a one-time only payment."      
        If reason_homeless = "" then err_msg = err_msg & vbNewLine & "* Enter the reason for family's homelessness."
		If barriers_housing = "" then err_msg = err_msg & vbNewLine & "* Enter the family's barrier(s) to housing."
		If referrals_made = "" then err_msg = err_msg & vbNewLine & "* Enter referals made for the family."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
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
Call write_variable_in_CASE_NOTE(">>>DO NOT release " & cash_type & " funds, family in shelter<<<")
Call write_variable_in_CASE_NOTE("* 100% of cash benefit to be issued to HCEA shelter account #52871.") 
If cash_type = "DWP" then 
    Call write_variable_in_CASE_NOTE("* DWP families are not eligible for the one-time only personal needs and medical co-pays")
ELSE 
    Call write_variable_in_CASE_NOTE(" Except $" & one_time_issuance & " to EBT for one-time only (10%) for personal needs. $20 for medical co-pays.")
END IF 
Call write_variable_in_CASE_NOTE("---")
Call write_bullet_and_variable_in_CASE_NOTE("Other income", other_income)
Call write_bullet_and_variable_in_CASE_NOTE("Money mismanagement", money_mismanagement)
Call write_bullet_and_variable_in_CASE_NOTE("Reason for homelessness", reason_homeless)
Call write_bullet_and_variable_in_CASE_NOTE("Barrier(s) to housing", barriers_housing)
Call write_bullet_and_variable_in_CASE_NOTE("Shelter history", shelter_history)
Call write_bullet_and_variable_in_CASE_NOTE("Social worker", social_worker)
Call write_bullet_and_variable_in_CASE_NOTE("Referrals made to", referrals_made)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE("* Explained shelter policies and client options to shelter such as bus tickets, temporary housing, private shelters, etc.") 
Call write_variable_in_CASE_NOTE("* Client given family social services number (348-4111) to discuss any family issues/barriers.")
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team") 

script_end_procedure("")