'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - SHELTER-PERMANENT HOUSING FOUND.vbs"
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
'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()
'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("06/11/2023", "Updated Dialog: Removed self-pay fields.", "Megan Geissler, Hennepin County")
CALL changelog_update("09/21/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("06/19/2017", "Initial version.", "MiKayla Handley")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)
Call check_for_MAXIS(False)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 176, 65, "Case Number Dialog"
  ButtonGroup ButtonPressed
    OkButton 75, 45, 45, 15
    CancelButton 125, 45, 45, 15
  EditBox 75, 5, 45, 15, MAXIS_case_number
  EditBox 75, 25, 95, 15, worker_signature
  Text 20, 10, 50, 10, "Case Number:"
  Text 10, 30, 60, 10, "Worker Signature:"
EndDialog

Do
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
      	Call validate_MAXIS_case_number(err_msg, "*")
        If trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Sign your case note."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call check_for_MAXIS(False)
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "ADDR", is_this_priv)
If is_this_priv = True then script_end_procedure("This case is privileged and cannot be accessed. The script will now stop.")


'-------------------------------------------------------------------------------------------------DIALOG
Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_number_one, phone_number_two, phone_number_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)
phone_number_list = "Select or Type|"
If phone_number_one <> "" Then phone_number_list = phone_number_list & phone_number_one & "|"
If phone_number_two <> "" Then phone_number_list = phone_number_list & phone_number_two & "|"
If phone_number_three <> "" Then phone_number_list = phone_number_list & phone_number_three & "|"

phone_number_array = split(phone_number_list, "|")  'creating an array of phone numbers to choose from that are active on the case, splitting by the delimiter "|"
Call convert_array_to_droplist_items(phone_number_array, phone_numbers) 'function to add phone_number array to a droplist - variable called phone_numbers


Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 456, 245, "Permanent Housing Found - Case#" & MAXIS_case_number
  EditBox 105, 5, 40, 15, move_date
  EditBox 210, 5, 40, 15, monthly_rent
  EditBox 335, 5, 20, 15, num_days
  EditBox 110, 40, 60, 15, rent_needed
  EditBox 110, 60, 60, 15, DD_needed
  EditBox 110, 80, 60, 15, rent_subsidy_paid
  EditBox 270, 40, 40, 15, earned_income
  EditBox 270, 60, 40, 15, rent_subsidy
  EditBox 400, 40, 40, 15, MFIP_DWP_income
  EditBox 400, 60, 40, 15, UNEA_income
  EditBox 400, 80, 40, 15, additional_funds
  EditBox 110, 115, 145, 15, LL_name
  EditBox 110, 135, 145, 15, LL_ADDR
  EditBox 110, 155, 145, 15, REW
  EditBox 110, 175, 40, 15, closed_servicepoint
  EditBox 390, 115, 50, 15, mand_vend_date
  ComboBox 365, 135, 75, 15, phone_numbers+chr(9)+client_phone, client_phone
  EditBox 365, 155, 75, 15, client_work_phone
  EditBox 65, 195, 375, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 345, 225, 50, 15
    CancelButton 400, 225, 50, 15
  Text 270, 10, 65, 10, "# of days in shelter:"
  Text 320, 160, 45, 10, "Client work #:"
  Text 20, 45, 80, 10, "Rent needed to move in:"
  Text 55, 10, 45, 10, "Move in date:"
  Text 20, 65, 90, 10, "DD amt needed to move in:"
  Text 20, 200, 45, 10, "Other notes: "
  Text 20, 85, 75, 10, "Rent subsidy paid by:"
  Text 165, 10, 45, 10, "Monthly rent:"
  Text 265, 120, 120, 10, "Mandatory vendor changed effective:"
  Text 20, 180, 90, 10, "ServicePoint closed out on:"
  Text 20, 120, 85, 10, "Landlord name/Vendor #:"
  Text 20, 160, 85, 10, "Rapid Exit Worker (REW):"
  Text 315, 140, 50, 10, "Client phone #:"
  Text 220, 65, 45, 10, "Rent subsidy:"
  Text 20, 140, 55, 10, "Landlord ADDR:"
  Text 215, 45, 55, 10, "Earned income:"
  Text 375, 65, 25, 10, "UNEA:"
  Text 230, 85, 170, 10, "Balance of funds owed to the LL will be issued from:"
  GroupBox 5, 30, 445, 80, "Resources for shelter funds:"
  Text 345, 45, 55, 10, "MFIP/DWP/HG:"
  GroupBox 5, 105, 445, 115, "Additional Info"
EndDialog

'Running the initial dialog
DO
	  DO
      err_msg = ""
      Dialog Dialog1
      cancel_confirmation
      IF IsDate(move_date) = False then err_msg = err_msg & vbNewLine & "* Enter a numeric move in date."
      If isNumeric(monthly_rent) = False then err_msg = err_msg & vbNewLine & "* Enter the monthly rent."
      If isNumeric(num_days) = False then err_msg = err_msg & vbNewLine & "* Enter a numeric number of days in shelter."
      If isNumeric(rent_needed) = False then err_msg = err_msg & vbNewLine & "* Enter a numeric rent amount needed to move in."
      If isNumeric(DD_needed) = False then err_msg = err_msg & vbNewLine & "* Enter a numeric DD amount needed to move in."
      If rent_subsidy_paid = "" then err_msg = err_msg & vbNewLine & "* Enter who will be paying for the rental subsidy."
      If isNumeric(rent_subsidy) = False then err_msg = err_msg & vbNewLine & "* Enter a numeric subsidy amount."
      If additional_funds = "" then err_msg = err_msg & vbNewLine & "* Enter who will be issuing the additional funds."
      If LL_name = "" then err_msg = err_msg & vbNewLine & "* Enter the Landlord's name."
      If LL_ADDR = "" then err_msg = err_msg & vbNewLine & "* Enter the Landlord's address."
      If REW = "" then err_msg = err_msg & vbNewLine & "* Enter the name of the Rapid Exit Worker."
      If isdate(closed_servicepoint) = False then err_msg = err_msg & vbNewLine & "* Enter the date that ServicePoint was closed out."
      If isdate(mand_vend_date) = False then err_msg = err_msg & vbNewLine & "* Enter the date the mandatory vendor status changed."
      If trim(client_phone) = "" or trim(client_phone) = "Select or Type" then err_msg = err_msg & vbcr & "* Enter the client's phone number."
      IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	  LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in


'The case note---------------------------------------------------------------------------------------
back_to_SELF
Call MAXIS_background_check
start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
Call write_variable_in_CASE_NOTE("### PERMANENT HOUSING FOUND/SHELTER ACCT VENDOR TO LL ###")
Call write_bullet_and_variable_in_CASE_NOTE("Move in date", move_date)
Call write_bullet_and_variable_in_CASE_NOTE("Monthly rent", monthly_rent)
Call write_bullet_and_variable_in_CASE_NOTE("Number of days in shelter", num_days)
Call write_bullet_and_variable_in_CASE_NOTE("Amt of rent needed to move in", rent_needed)
Call write_bullet_and_variable_in_CASE_NOTE("Amt of DD needed to move in", DD_needed)
Call write_bullet_and_variable_in_CASE_NOTE("Earned income", earned_income)
Call write_bullet_and_variable_in_CASE_NOTE("MFIP/DWP/HG", MFIP_DWP_income)
Call write_bullet_and_variable_in_CASE_NOTE("UNEA", UNEA_income)
Call write_bullet_and_variable_in_CASE_NOTE("Rent subsidy amt", rent_subsidy)
Call write_bullet_and_variable_in_CASE_NOTE("Rent subsidy paid by", rent_subsidy_paid)
Call write_bullet_and_variable_in_CASE_NOTE("Balance of funds owed to the LL will be issued from", additional_funds)
Call write_variable_in_CASE_NOTE ("---")
Call write_bullet_and_variable_in_CASE_NOTE("Landlord name/vendor #", LL_name)
Call write_bullet_and_variable_in_CASE_NOTE("Landlord address", LL_ADDR)
Call write_bullet_and_variable_in_CASE_NOTE("Rapid Exit Worker (REW)", REW)
Call write_bullet_and_variable_in_CASE_NOTE("Client phone #", client_phone)
Call write_bullet_and_variable_in_CASE_NOTE("Client work phone #", client_work_phone)
Call write_variable_in_CASE_NOTE ("---")
Call write_bullet_and_variable_in_CASE_NOTE("Mandatory vendor changed effective", mand_vend_date)
Call write_bullet_and_variable_in_CASE_NOTE("Service Point closed out on ", closed_servicepoint)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE ("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

end_msg = "Success!"
end_msg = end_msg & vbCr & vbCr & "Case noted Permanent Housing Found"
script_end_procedure(end_msg)


'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------06/11/2024  
'--Tab orders reviewed & confirmed----------------------------------------------06/11/2024
'--Mandatory fields all present & Reviewed--------------------------------------06/11/2024
'--All variables in dialog match mandatory fields-------------------------------06/11/2024
'Review dialog names for content and content fit in dialog----------------------06/11/2024
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------06/11/2024
'--CASE:NOTE Header doesn't look funky------------------------------------------06/11/2024
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------06/11/2024
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------06/11/2024
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------06/11/2024
'--MAXIS_background_check reviewed (if applicable)------------------------------06/11/2024
'--PRIV Case handling reviewed -------------------------------------------------NA
'--Out-of-County handling reviewed----------------------------------------------NA
'--script_end_procedures (w/ or w/o error messaging)----------------------------06/11/2024
'--BULK - review output of statistics and run time/count (if applicable)--------NA
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------06/11/2024
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------06/11/2024
'--Incrementors reviewed (if necessary)-----------------------------------------
'--Denomination reviewed -------------------------------------------------------06/11/2024
'--Script name reviewed---------------------------------------------------------06/11/2024
'--BULK - remove 1 incrementor at end of script reviewed------------------------NA

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------06/11/2024
'--comment Code-----------------------------------------------------------------06/11/2024
'--Update Changelog for release/update------------------------------------------06/11/2024
'--Remove testing message boxes-------------------------------------------------06/11/2024
'--Remove testing code/unnecessary code-----------------------------------------06/11/2024
'--Review/update SharePoint instructions----------------------------------------06/11/2024
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------06/11/2024
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------06/11/2024
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------06/11/2024
'--Complete misc. documentation (if applicable)---------------------------------06/11/2024
'--Update project team/issue contact (if applicable)----------------------------06/11/2024