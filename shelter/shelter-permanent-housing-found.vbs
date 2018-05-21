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
BeginDialog perm_housing_found_dialog, 0, 0, 416, 330, "Permanent Housing Found"
  EditBox 120, 10, 55, 15, MAXIS_case_number
  EditBox 225, 10, 55, 15, move_date
  EditBox 350, 10, 55, 15, monthly_rent
  EditBox 120, 35, 55, 15, vendored_to_HCEA
  EditBox 225, 35, 55, 15, shelter_cost
  EditBox 360, 35, 45, 15, num_nights
  EditBox 100, 95, 55, 15, rent_needed
  EditBox 100, 115, 55, 15, DD_needed
  EditBox 235, 70, 55, 15, balance_HCEA_acct
  EditBox 235, 90, 55, 15, client_funds
  EditBox 235, 110, 55, 15, rent_subsidy_paid
  EditBox 235, 130, 55, 15, vendor_to_LL
  EditBox 180, 150, 135, 15, additional_funds
  EditBox 350, 70, 55, 15, other_partners
  EditBox 350, 90, 55, 15, earned_income
  EditBox 350, 110, 55, 15, rent_subsidy
  EditBox 350, 130, 55, 15, MFIP_DWP_income
  EditBox 350, 150, 55, 15, UNEA_income
  EditBox 75, 185, 145, 15, LL_name
  EditBox 60, 205, 200, 15, LL_ADDR
  EditBox 130, 225, 130, 15, ESP
  EditBox 130, 245, 130, 15, REW
  EditBox 130, 265, 80, 15, closed_servicepoint
  EditBox 350, 185, 55, 15, mand_vend_date
  EditBox 350, 205, 55, 15, client_phone
  EditBox 350, 225, 55, 15, client_work_phone
  EditBox 350, 245, 55, 15, TANF_months
  EditBox 350, 265, 55, 15, num_days_shelter
  EditBox 130, 285, 275, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 300, 310, 50, 15
    CancelButton 355, 310, 50, 15
  Text 285, 40, 70, 10, "for number of nights:"
  Text 285, 250, 65, 10, "TANF months used:"
  Text 300, 230, 45, 10, "Client work #:"
  Text 15, 95, 80, 10, "Rent needed to move in:"
  Text 180, 15, 45, 10, "Move in date:"
  Text 10, 115, 90, 10, "DD amt needed to move in:"
  Text 160, 95, 75, 10, "Client funds available:"
  Text 80, 290, 45, 10, "Other notes: "
  Text 160, 115, 75, 10, "Rent subsidy paid by:"
  Text 70, 15, 45, 10, "Case number:"
  Text 65, 135, 170, 10, "Balance in HCEA Shelter acct to be vendored to LL:"
  Text 305, 15, 45, 10, "Monthly rent:"
  Text 225, 190, 120, 10, "Mandatory vendor changed effective:"
  Text 35, 270, 90, 10, "ServicePoint closed out on:"
  Text 15, 230, 105, 10, "Employ. Service Provider (ESP):"
  Text 30, 75, 205, 10, "Balance in HCEA Shelter acct available toward rent and/or DD:"
  Text 295, 75, 55, 10, "Other partners:"
  Text 5, 190, 65, 10, "LL name/Vendor #: "
  Text 40, 250, 85, 10, "Rapid Exit Worker (REW):"
  Text 300, 210, 50, 10, "Client phone #:"
  Text 300, 115, 45, 10, "Rent subsidy:"
  Text 5, 210, 55, 10, "Landlord ADDR:"
  Text 635, 40, 10, 0, "-10"
  Text 5, 40, 115, 10, "Amt vendored to HCEA sheltet acct:"
  Text 295, 95, 55, 10, "Earned income:"
  Text 180, 40, 45, 10, "Shelter cost:"
  Text 325, 155, 25, 10, "UNEA:"
  Text 10, 155, 170, 10, "Balance of funds owed to the LL will be issued from:"
  Text 280, 270, 65, 10, "# of days in shelter:"
  GroupBox 5, 60, 405, 115, "Resources for shelter funds:"
  Text 295, 135, 55, 10, "MFIP/DWP/HG:"
EndDialog


'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog perm_housing_found_dialog
        cancel_confirmation
		IF len(MAXIS_case_number) > 8 or IsNumeric(MAXIS_case_number) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		IF IsDate(move_date) = False then err_msg = err_msg & vbNewLine & "* Enter a numeric move in date."
		If monthly_rent = "" then err_msg = err_msg & vbNewLine & "* Enter the move in date." 
		If vendored_to_HCEA = "" then err_msg = err_msg & vbNewLine & "* Enter the amt vendored to HCEA shelter acct."
		If shelter_cost = "" then err_msg = err_msg & vbNewLine & "* Enter the shelter costs."
		If isNumeric(num_nights) = False then err_msg = err_msg & vbNewLine & "* Enter a numeric number of nights at the shelter."
		If isNumeric(balance_HCEA_acct) = False then err_msg = err_msg & vbNewLine & "* Enter the HCEA Shelter account balance available for rent/DD."
		If isNumeric(rent_needed) = False then err_msg = err_msg & vbNewLine & "* Enter a numeric rent amount needed to move in."
		If isNumeric(DD_needed) = False then err_msg = err_msg & vbNewLine & "* Enter a numeric DD amount needed to move in."
		If isNumeric(client_funds) = False then err_msg = err_msg & vbNewLine & "* Enter a numeric amount of client funds available."
		If rent_subsidy_paid = "" then err_msg = err_msg & vbNewLine & "* Enter who will be paying for the rental subsidy."
		If isNumeric(rent_subsidy) = False then err_msg = err_msg & vbNewLine & "* Enter a numeric subsidy amount."
		If isNumeric(vendor_to_LL) = False then err_msg = err_msg & vbNewLine & "* Enter a numeric amount for the balance in HCEA shelter acct that will be vendored to LL."
		If additional_funds = "" then err_msg = err_msg & vbNewLine & "* Enter who will be issuing the additional funds."
		If LL_name = "" then err_msg = err_msg & vbNewLine & "* Enter the Landlord's name."
		If LL_ADDR = "" then err_msg = err_msg & vbNewLine & "* Enter the Landlord's address."
		If ESP = "" then err_msg = err_msg & vbNewLine & "* Enter the name of the Employment Services Provider."
		If REW = "" then err_msg = err_msg & vbNewLine & "* Enter the name of the Rapid Exit Worker."
		If isdate(closed_servicepoint) = False then err_msg = err_msg & vbNewLine & "* Enter the date that ServicePoint was closed out."
		If isdate(mand_vend_date) = False then err_msg = err_msg & vbNewLine & "* Enter the date the mandatory vendor status changed."
		If client_phone = "" then err_msg = err_msg & vbNewLine & "* Enter the client's phone number."
		If isNumeric(TANF_months) = False then err_msg = err_msg & vbNewLine & "* Enter numeric number of TANF months used."
		If isNumeric(num_days_shelter) = False then err_msg = err_msg & vbNewLine & "* Enter the numeric number of days in shelter."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
 Call check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False
 
back_to_SELF
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen CM_mo, 20, 43	'entering current footer month/year
EMWriteScreen CM_yr, 20, 46

'The case note---------------------------------------------------------------------------------------
start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
Call write_variable_in_CASE_NOTE("### PERMANENT HOUSING FOUND/SHELTER ACCT VENDOR TO LL ###")
Call write_bullet_and_variable_in_CASE_NOTE("Move in date", move_date)
Call write_bullet_and_variable_in_CASE_NOTE("Monthly rent", monthly_rent)
Call write_bullet_and_variable_in_CASE_NOTE("Amount vendored to HCEA shelter account", vendored_to_HCEA)
Call write_variable_in_CASE_NOTE("* Shelter cost: " & shelter_cost & " for " & num_nights &  " nights.")
Call write_bullet_and_variable_in_CASE_NOTE("Balance in HCEA shelter acct available towards rent and/or DD", balance_HCEA_acct)
Call write_bullet_and_variable_in_CASE_NOTE("Amt of rent needed to move in", rent_needed)
Call write_bullet_and_variable_in_CASE_NOTE("Amt of DD needed to move in", DD_needed)
Call write_bullet_and_variable_in_CASE_NOTE("Client funds available", client_funds)
Call write_bullet_and_variable_in_CASE_NOTE("Other partners", other_partners)
Call write_bullet_and_variable_in_CASE_NOTE("Earned income", earned_income)
Call write_bullet_and_variable_in_CASE_NOTE("MFIP/DWP/HG", MFIP_DWP_income)
Call write_bullet_and_variable_in_CASE_NOTE("UNEA", UNEA_income)
Call write_bullet_and_variable_in_CASE_NOTE("Rent subsidy amt", rent_subsidy)
Call write_bullet_and_variable_in_CASE_NOTE("Rent subsidy paid by", rent_subsidy_paid)
Call write_bullet_and_variable_in_CASE_NOTE("Balance in HCEA shelter acct to be vendored to LL", vendor_to_LL )
Call write_bullet_and_variable_in_CASE_NOTE("Balance of funds owed to the LL will be issued from", additional_funds)
Call write_variable_in_CASE_NOTE ("---")
Call write_bullet_and_variable_in_CASE_NOTE("Landlord name/vendor #", LL_name)
Call write_bullet_and_variable_in_CASE_NOTE("Landlord address", LL_ADDR)
Call write_bullet_and_variable_in_CASE_NOTE("Employment Services Provider (ESP)", ESP)
Call write_bullet_and_variable_in_CASE_NOTE("Rapid Exit Worker (REW)", REW)
Call write_bullet_and_variable_in_CASE_NOTE("Client phone #", client_phone)
Call write_bullet_and_variable_in_CASE_NOTE("Client work phone #", client_work_phone)
Call write_variable_in_CASE_NOTE ("---")
Call write_bullet_and_variable_in_CASE_NOTE("Mandatory vendor changed effective", mand_vend_date)
Call write_bullet_and_variable_in_CASE_NOTE("Service Point closed out on ", closed_servicepoint)
Call write_bullet_and_variable_in_CASE_NOTE("TANF months used", TANF_months)
Call write_bullet_and_variable_in_CASE_NOTE("Number of days in shelter", num_days_shelter)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE ("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

script_end_procedure("")	