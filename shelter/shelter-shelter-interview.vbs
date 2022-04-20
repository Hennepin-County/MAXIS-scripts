'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - SHELTER-SHELTER INTERVIEW.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 300         	'manual run time in seconds
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
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine & "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine & "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
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

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to MAXIS, and grabbing the case number and footer month'
EMConnect ""
CALL check_for_maxis(FALSE) 'checking for passord out, brings up dialog'
CALL MAXIS_case_number_finder(MAXIS_case_number)

If MAXIS_case_number <> "" Then 		'If a case number is found the script will get the list of
	Call Generate_Client_List(HH_Memb_DropDown, "Select One:")
End If
'Running the dialog for case number and client
Do
	err_msg = ""
    Dialog1 = "" 'Blanking out previous dialog detail
	BeginDialog Dialog1, 0, 0, 201, 90, "Shelter Interview"
	  EditBox 55, 5, 45, 15, MAXIS_case_number
	  DropListBox 80, 25, 115, 15, HH_Memb_DropDown, clt_to_update
	  EditBox 80, 45, 115, 15, worker_signature
	  ButtonGroup ButtonPressed
	    OkButton 100, 70, 45, 15
	    CancelButton 150, 70, 45, 15
	  Text 5, 10, 45, 10, "Case number:"
	  Text 5, 30, 70, 10, "Household member:"
	  Text 5, 50, 60, 10, "Worker signature:"
	  ButtonGroup ButtonPressed
	    PushButton 110, 5, 85, 15, "HH MEMB SEARCH", search_button
	EndDialog

	Dialog Dialog1
	IF ButtonPressed = cancel Then StopScript
	IF ButtonPressed = search_button Then
		If MAXIS_case_number = "" Then
			MsgBox "Cannot search without a case number, please try again."
		Else
			HH_Memb_DropDown = ""
			Call Generate_Client_List(HH_Memb_DropDown, "Select One:")
			err_msg = err_msg & "Start Over"
		End If
	End If
	Call validate_MAXIS_case_number(err_msg, "*")
	IF clt_to_update = "Select One:" Then err_msg = err_msg & vbNewLine & "* Please select a client to update."
    IF trim(worker_signature) = "" THEN err_msg = err_msg & vbNewLine & "* Please enter your worker signature."
Loop until err_msg = ""
CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in


CALL navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv)
IF is_this_priv = TRUE THEN script_end_procedure("Please send a request for access to the case to Knowledge Now.")

'redefine ref_numb'
MEMB_number = left(clt_to_update, 2)	'Setting the reference number to use
EMWriteScreen MEMB_number, 20, 76
TRANSMIT
EMReadScreen client_first_name, 12, 6, 63
client_first_name = replace(client_first_name, "_", "")
client_first_name = trim(client_first_name)
EMReadScreen client_last_name, 25, 6, 30
client_last_name = replace(client_last_name, "_", "")
client_last_name = trim(client_last_name)
EMReadscreen client_mid_initial, 1, 6, 79
EMReadScreen client_DOB, 10, 8, 42
EMReadscreen client_SSN, 11, 7, 42
client_SSN = replace(client_SSN, " ", "")

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 316, 255, "Shelter Interview"
  EditBox 100, 5, 205, 15, barriers_housing
  EditBox 100, 25, 205, 15, shelter_history
  EditBox 105, 50, 200, 15, reason_homeless
  EditBox 105, 70, 200, 15, other_income
  EditBox 105, 90, 70, 15, address_verf
  EditBox 255, 90, 50, 15, phone_number
  EditBox 105, 110, 200, 15, comments_notes
  EditBox 105, 135, 200, 15, referrals_made
  EditBox 105, 155, 200, 15, social_worker
  EditBox 105, 175, 200, 15, other_income
  EditBox 105, 195, 200, 15, other_notes
  DropListBox 240, 215, 65, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO"+chr(9)+"N/A", screening_questions_dropdown
  ButtonGroup ButtonPressed
    OkButton 200, 235, 50, 15
    CancelButton 255, 235, 50, 15
	PushButton 10, 235, 65, 15, "SHELTER ", shelter_button
  Text 35, 25, 55, 10, "Shelter history:"
  Text 10, 55, 90, 10, "Reason for homelessness:"
  Text 75, 75, 25, 10, "Name:"
  Text 70, 95, 30, 10, "Address:"
  Text 195, 95, 50, 10, "Phone number: "
  Text 60, 115, 40, 10, "Comments: "
  Text 35, 140, 60, 10, "Referrals made to:"
  Text 50, 180, 45, 10, "Other income:"
  Text 55, 200, 40, 10, "Other notes:"
  Text 15, 10, 70, 10, "Barrier(s) to housing:"
  Text 105, 220, 130, 10, "Health screening questions answered:"
  GroupBox 5, 40, 305, 90, "Homelessness verified:"
  Text 50, 160, 50, 10, "Social worker:"
EndDialog

DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
        If reason_homeless = "" then err_msg = err_msg & vbNewLine & "* Please enter the reason for family's homelessness."
		If barriers_housing = "" then err_msg = err_msg & vbNewLine & "* Please enter the family's barrier(s) to housing."
		If referrals_made = "" then err_msg = err_msg & vbNewLine & "* Please enter referrals made for the family."

		If ButtonPressed = shelter_button Then               'Pulling up the hsr page if the button was pressed.
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Shelter_Team.aspx?xsdata=MDV8MDF8fGM0ZjRlYmM5ZWY4NzQyNjBiNTZhMDhkYTIzMWFmNTMwfDhhZWZkZjlmODc4MDQ2YmY4ZmI3NGM5MjQ2NTNhOGJlfDB8MHw2Mzc4NjA4OTU5MTYwNzg0NDJ8R29vZHxWR1ZoYlhOVFpXTjFjbWwwZVZObGNuWnBZMlY4ZXlKV0lqb2lNQzR3TGpBd01EQWlMQ0pRSWpvaVYybHVNeklpTENKQlRpSTZJazkwYUdWeUlpd2lWMVFpT2pFeGZRPT18MXxNVGs2WkdVeE1qUTVPREV0TURaaFlpMDBOMkprTFRneE16a3RPV0V4TXpnME1tWmlNREU0WDJZeE1EZGpOemxpTFRFeE5qTXROREF3TWkxaU56WTNMV1V4WmpabFptSXlabVZrTVVCMWJuRXVaMkpzTG5Od1lXTmxjdz09fHw%3D&sdata=ckFaYmY4T0dXMmlUVGZpcmhHRDNUNUY0ZmhtblBidmROcHlHVTBOWnJBTT0%3D&ovuser=8aefdf9f-8780-46bf-8fb7-4c924653a8be%2CMikayla.Handley%40hennepin.us&OR=Teams-HL&CT=1650492794967&params=eyJBcHBOYW1lIjoiVGVhbXMtRGVza3RvcCIsIkFwcFZlcnNpb24iOiIyNy8yMjAzMDcwMTYxMCJ9"
			err_msg = "LOOP"
		Else                                                'If the instructions button was NOT pressed, we want to display the error message if it exists.
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
		End If
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'adding the case number
back_to_self
EMWriteScreen "________", 18, 43
EMWriteScreen MAXIS_case_number, 18, 43
EMWriteScreen CM_mo, 20, 43	'entering current footer month/year
EMWriteScreen CM_yr, 20, 46

CALL determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)

when_contact_was_made = date & ", " & time
IF name_verf <> "" THEN homeless_verified = "YES"

'The case note'
start_a_blank_CASE_NOTE
CALL write_variable_in_CASE_NOTE("~ Family in Shelter Interview with M" &  MEMB_number & " ~")
Call write_bullet_and_variable_in_CASE_NOTE("Active programs", list_active_programs)
Call write_bullet_and_variable_in_CASE_NOTE("Pending programs", list_pending_programs)
CALL write_bullet_and_variable_in_CASE_NOTE("Date:", when_contact_was_made)
Call write_bullet_and_variable_in_CASE_NOTE("Barrier(s) to housing", barriers_housing)
Call write_bullet_and_variable_in_CASE_NOTE("Reason for homelessness", reason_homeless)
Call write_bullet_and_variable_in_CASE_NOTE("Homelessness verified", homeless_verified)
CALL write_variable_in_CASE_NOTE("Name: " & name_verf)
CALL write_variable_in_CASE_NOTE("Phone number: " & phone_number)
CALL write_variable_in_CASE_NOTE("Address: " & address_verf)
CALL write_variable_in_CASE_NOTE("Comments: " & comments_notes)
Call write_bullet_and_variable_in_CASE_NOTE("Shelter history", shelter_history)
Call write_bullet_and_variable_in_CASE_NOTE("Social worker", social_worker)
Call write_bullet_and_variable_in_CASE_NOTE("Referrals made to", referrals_made)
Call write_bullet_and_variable_in_CASE_NOTE("Other income", other_income)
Call write_bullet_and_variable_in_CASE_NOTE("Health screening questions answered", screening_questions_dropdown)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE("* Explained shelter policies and client options to shelter such as bus tickets, temporary housing, private shelters, etc.")
Call write_variable_in_CASE_NOTE("* Client given family social services number (348-4111) to discuss any family issues/barriers.")
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

IF end_msg <> "" Then
    closing_message = closing_message & vbCr & vbCr & "" & vbCr & end_msg
ELSE
    closing_message = closing_message & vbCr & vbCr & " " &
END IF
Call script_end_procedure_with_error_report(closing_message)


'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------04/12/2022
'--Tab orders reviewed & confirmed----------------------------------------------04/12/2022
'--Mandatory fields all present & Reviewed--------------------------------------04/12/2022
'--All variables in dialog match mandatory fields-------------------------------04/12/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------04/12/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------04/12/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------04/12/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------04/12/2022
'--PRIV Case handling reviewed -------------------------------------------------04/12/2022
'--Out-of-County handling reviewed----------------------------------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------04/12/2022
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------04/12/2022
'--Incrementors reviewed (if necessary)-----------------------------------------N/A
'--Denomination reviewed -------------------------------------------------------N/A
'--Script name reviewed---------------------------------------------------------N/A
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------N/A
'--comment Code-----------------------------------------------------------------N/A
'--Update Changelog for release/update------------------------------------------N/A
'--Remove testing message boxes-------------------------------------------------N/A
'--Remove testing code/unnecessary code-----------------------------------------N/A
'--Review/update SharePoint instructions----------------------------------------N/A
'--Review Best Practices using BZS page ----------------------------------------N/A
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------N/A
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------N/A
