'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - COUNTY BURIAL APPLICATION.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 120           'manual run time in seconds
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/30/2016", "Corrected Typo: Creamation to Cremation.", "Charles Potter, DHS")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Dialog---------------------------------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 126, 45, "Case number dialog"
  EditBox 60, 5, 60, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 10, 25, 50, 15
    CancelButton 70, 25, 50, 15
  Text 5, 10, 50, 10, "Case number:"
EndDialog

BeginDialog County_Burial_Application_Received, 0, 0, 266, 325, "County Burial Application Received"
  EditBox 55, 5, 60, 15, MAXIS_case_number
  EditBox 200, 5, 60, 15, date_received
  EditBox 55, 25, 60, 15, date_of_death
  EditBox 200, 25, 30, 15, CFR_resp
  EditBox 45, 50, 215, 15, CASH_ACCTs
  EditBox 55, 70, 205, 15, other_assets
  EditBox 80, 90, 180, 15, Total_Counted_Assets
  ComboBox 80, 115, 180, 45, ""+chr(9)+"Cremation"+chr(9)+"Burial", services_requested
  EditBox 60, 135, 200, 15, funeral_home
  EditBox 60, 155, 75, 15, funeral_home_phone
  EditBox 225, 155, 35, 15, funeral_home_amount
  EditBox 60, 175, 200, 15, cemetary
  EditBox 60, 195, 75, 15, cemetary_phone
  EditBox 225, 195, 35, 15, cemetary_amount
  EditBox 60, 220, 200, 15, contact_name
  EditBox 60, 240, 75, 15, contact_phone
  EditBox 185, 240, 75, 15, contact_fax
  EditBox 55, 265, 205, 15, other_notes
  EditBox 55, 285, 205, 15, action_taken
  EditBox 65, 305, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 160, 305, 50, 15
    CancelButton 210, 305, 50, 15
  Text 5, 10, 50, 10, "Case Number: "
  Text 145, 10, 50, 10, "Date received: "
  Text 5, 30, 50, 10, "Date of death:"
  Text 175, 30, 20, 10, "CFR:"
  Text 5, 55, 35, 10, "Accounts:"
  Text 5, 75, 45, 10, "Other Assets:"
  Text 5, 95, 75, 10, "Total Counted Assets:"
  Text 5, 120, 75, 10, "Services Requested:"
  Text 5, 140, 50, 10, "Funeral Home:"
  Text 5, 160, 50, 10, "Phone Number:"
  Text 155, 160, 65, 10, "Amount Requested:"
  Text 5, 180, 35, 10, "Cemetary:"
  Text 5, 200, 50, 10, "Phone Number:"
  Text 155, 200, 65, 10, "Amount Requested:"
  Text 5, 225, 50, 10, "Contact/AREP:"
  Text 5, 245, 50, 10, "Phone Number:"
  Text 160, 245, 20, 10, "Fax:"
  Text 5, 270, 40, 10, "Other notes: "
  Text 5, 290, 50, 10, "Actions taken: "
  Text 5, 310, 60, 10, "Worker Signature: "
  GroupBox 0, 40, 265, 70, ""
  GroupBox 0, 210, 265, 50, ""
  GroupBox 0, 105, 265, 110, ""
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------
'connecting to BlueZone, and grabbing the case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)	'Finds the case number
get_county_code										'Sets the county of financial resp. as current county by default
CFR_resp = right(worker_county_code, 2)

If MAXIS_case_number = "" Then
	Dialog case_number_dialog
	cancel_confirmation
End If

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Reads the date of death if it has been entered in MAXIS
Call navigate_to_MAXIS_screen ("STAT", "MEMB")
EMReadScreen date_of_death, 10, 19, 42
If date_of_death = "__ __ ____" Then
	date_of_death = ""
Else
	date_of_death = replace(date_of_death, " ", "/")
End If
'Pulls some asset information
call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", CASH_ACCTs)
call autofill_editbox_from_MAXIS(HH_member_array, "CASH", CASH_ACCTs)
call autofill_editbox_from_MAXIS(HH_member_array, "SECU", other_assets)

'calling the dialog---------------------------------------------------------------------------------------------------------------
Do
	DO
		err_msg = ""
		Dialog County_Burial_Application_Received
		cancel_confirmation
		IF buttonpressed = 0 THEN stopscript
		If MAXIS_case_number = "" Then err_msg = err_msg & vbNewLine & "You must enter a case number."
		If date_received = "" Then err_msg = err_msg & vbNewLine & "Enter the date the application was recieved."
		If date_of_death = "" Then err_msg = err_msg & vbNewLine & "Enter the date of death."
		If CFR_resp = "" Then err_msg = err_msg & vbNewLine & "List the County of Financial Responsibility."
		If worker_signature = "" Then err_msg = err_msg & vbNewLine & "Sign your case note."
		If err_msg <> "" Then MsgBox ("The following must be resolved before continuing." & vbNewLine & err_msg)
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'checking for an active MAXIS session
CALL check_for_MAXIS(FALSE)

If contact_name <> "" Then contact_info = contact_name
If contact_phone <> "" Then contact_info = contact_info & " - Phone: " & contact_phone
If contact_fax <> "" THen contact_info = contact_info & " - Fax: " & contact_fax

If funeral_home_amount <> "" Then funeral_home_amount = FormatCurrency(funeral_home_amount)
If cemetary_amount <> "" Then cemetary_amount = FormatCurrency(cemetary_amount)

'Format funeral home and cemetary like above'
If funeral_home <> "" Then funeral_info = "Services at " & funeral_home & ". "
If funeral_home_amount <> "" THen funeral_info = funeral_info & "Ammount requested: " & funeral_home_amount
If funeral_home_phone <> "" Then funeral_info = funeral_info & " Phone Number: " & funeral_home_phone

If cemetary <> "" Then cemetary_info = "Services at " & cemetary & ". "
If cemetary_amount <> "" THen cemetary_info = cemetary_info & "Ammount requested: " & cemetary_amount
If cemetary_phone <> "" Then cemetary_info = cemetary_info & " Phone Number: " & cemetary_phone

'The case note---------------------------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
CALL write_variable_in_CASE_NOTE("***County Burial Application Received")
CALL write_bullet_and_variable_in_CASE_NOTE("Date Received", Date_Received)
CALL write_bullet_and_variable_in_CASE_NOTE("Date of death", Date_of_death)
CALL write_bullet_and_variable_in_CASE_NOTE("CFR", CFR_resp)
CALL write_bullet_and_variable_in_CASE_NOTE("Accounts", CASH_ACCTs)
CALL write_bullet_and_variable_in_CASE_NOTE("Assets", other_assets)
CALL write_bullet_and_variable_in_CASE_NOTE("Total Counted Assets", Total_Counted_Assets)
Call write_bullet_and_variable_in_CASE_NOTE("Services Requested", services_requested)
If funeral_home <> "" Then Call write_variable_in_CASE_NOTE("* Requesting " & funeral_home_amount & " for services at " & funeral_home & ". Contact: " & funeral_home_phone)
If cemetary <> "" Then Call write_variable_in_CASE_NOTE("* Requesting " & cemetary_amount & " for services at " & cemetary & ". Contact: " & cemetary_phone)
Call write_bullet_and_variable_in_CASE_NOTE("Contact", contact_info)
CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
CALL write_bullet_and_variable_in_CASE_NOTE("Action taken", action_taken)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

Script_end_procedure("")
