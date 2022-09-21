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
CALL changelog_update("09/21/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("06/19/2017", "Initial version.", "MiKayla Handley")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 416, 310, "Permanent Housing Found"
  EditBox 55, 5, 45, 15, MAXIS_case_number
  EditBox 160, 5, 40, 15, move_date
  EditBox 255, 5, 40, 15, monthly_rent
  EditBox 160, 25, 40, 15, vendored_to_HCEA
  EditBox 255, 25, 40, 15, shelter_cost
  EditBox 370, 25, 20, 15, num_nights
  EditBox 100, 75, 40, 15, rent_needed
  EditBox 100, 95, 40, 15, DD_needed
  EditBox 215, 55, 40, 15, balance_HCEA_acct
  EditBox 225, 75, 40, 15, client_funds
  EditBox 225, 95, 40, 15, rent_subsidy_paid
  EditBox 185, 115, 40, 15, vendor_to_LL
  EditBox 185, 135, 40, 15, additional_funds
  EditBox 350, 55, 40, 15, other_partners
  EditBox 350, 75, 40, 15, earned_income
  EditBox 350, 95, 40, 15, rent_subsidy
  EditBox 350, 115, 40, 15, MFIP_DWP_income
  EditBox 350, 135, 40, 15, UNEA_income
  EditBox 75, 165, 145, 15, LL_name
  EditBox 75, 185, 145, 15, LL_ADDR
  EditBox 120, 205, 100, 15, ESP
  EditBox 120, 225, 130, 15, REW
  EditBox 120, 245, 80, 15, closed_servicepoint
  EditBox 350, 165, 50, 15, mand_vend_date
  EditBox 350, 185, 50, 15, client_phone
  EditBox 350, 205, 50, 15, client_work_phone
  EditBox 350, 225, 50, 15, TANF_months
  EditBox 350, 245, 50, 15, num_days_shelter
  EditBox 120, 265, 280, 15, other_notes
  EditBox 120, 285, 165, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 295, 285, 50, 15
    CancelButton 350, 285, 50, 15
  Text 310, 30, 60, 10, "number of nights:"
  Text 285, 230, 65, 10, "TANF months used:"
  Text 300, 210, 45, 10, "Client work #:"
  Text 10, 80, 80, 10, "Rent needed to move in:"
  Text 110, 10, 45, 10, "Move in date:"
  Text 10, 100, 90, 10, "DD amt needed to move in:"
  Text 150, 80, 75, 10, "Client funds available:"
  Text 10, 270, 45, 10, "Other notes: "
  Text 150, 100, 75, 10, "Rent subsidy paid by:"
  Text 5, 10, 45, 10, "Case number:"
  Text 10, 120, 170, 10, "Balance in HCEA Shelter acct to be vendored to LL:"
  Text 205, 10, 45, 10, "Monthly rent:"
  Text 225, 170, 120, 10, "Mandatory vendor changed effective:"
  Text 10, 250, 90, 10, "ServicePoint closed out on:"
  Text 10, 210, 105, 10, "Employ. Service Provider (ESP):"
  Text 10, 60, 205, 10, "Balance in HCEA Shelter acct available toward rent and/or DD:"
  Text 295, 60, 55, 10, "Other partners:"
  Text 10, 170, 65, 10, "LL name/Vendor #: "
  Text 10, 230, 85, 10, "Rapid Exit Worker (REW):"
  Text 295, 190, 50, 10, "Client phone #:"
  Text 300, 100, 45, 10, "Rent subsidy:"
  Text 10, 190, 55, 10, "Landlord ADDR:"
  Text 40, 30, 115, 10, "Amt vendored to HCEA sheltet acct:"
  Text 295, 80, 55, 10, "Earned income:"
  Text 210, 30, 45, 10, "Shelter cost:"
  Text 325, 140, 25, 10, "UNEA:"
  Text 10, 140, 170, 10, "Balance of funds owed to the LL will be issued from:"
  Text 280, 250, 65, 10, "# of days in shelter:"
  GroupBox 5, 45, 395, 115, "Resources for shelter funds:"
  Text 295, 120, 55, 10, "MFIP/DWP/HG:"
  Text 10, 290, 60, 10, "Worker Signature:"
EndDialog

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog Dialog1
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
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Sign your case note."
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
