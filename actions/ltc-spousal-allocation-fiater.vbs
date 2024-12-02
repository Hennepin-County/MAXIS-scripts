'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - LTC - SPOUSAL ALLOCATION FIATER.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 336                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
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
call changelog_update("10/02/2019", "Bug fix in ELIG/HC when finding the starting HC month. Also removed outdated JOBS panel coding.", "Ilse Ferris, Hennepin County")
call changelog_update("09/12/2019", "Updated main dialog to only one date selection to enter the allocation start month. Added mandatory fields in case note dialog.", "Ilse Ferris, Hennepin County")
call changelog_update("09/11/2019", "Updated HEST deductions when electric and phones deductions are present for community spouse.", "Ilse Ferris, Hennepin County")
call changelog_update("01/05/2018", "Updated coordinates in STAT/JOBS for income type and verification codes.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to MAXIS, grabs case number and footer month/year
EMConnect ""
call MAXIS_case_number_finder(MAXIS_case_number)
call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 201, 115, "Case number Dialog"
  EditBox 110, 5, 45, 15, MAXIS_case_number
  EditBox 110, 25, 20, 15, MAXIS_footer_month
  EditBox 135, 25, 20, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 60, 45, 45, 15
    CancelButton 110, 45, 45, 15
  Text 60, 10, 45, 10, "Case number:"
  Text 10, 30, 100, 10, "Allocation budget month/year:"
  GroupBox 5, 65, 195, 45, "If LTC spouse is open on CASH programs"
  Text 10, 80, 190, 30, "You will need to process the spousal allocation manually. The script currently does not support budgeting public assistance CASH programs into the spousal allocation."
EndDialog

'Shows case number dialog
Do
    Do
        err_msg = ""
        dialog Dialog1
        cancel_without_confirmation
        IF len(MAXIS_case_number) > 8 or isnumeric(MAXIS_case_number) = false THEN err_msg = err_msg & vbCr & "Enter a valid case number."		'mandatory field
        If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid 2-digit footer month."
        If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid 2-digit footer year."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
    LOOP UNTIL err_msg = ""									'loops until all errors are resolved
CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

spousal_allocation_footer_month = MAXIS_footer_month
spousal_allocation_footer_year = MAXIS_footer_year

'Enters into STAT for the client
Call navigate_to_MAXIS_screen("STAT", "MEMB")

'Checks for which HH member is the spouse. The spouse is coded as "02" on STAT/MEMB.
Do
  EMReadScreen spouse_check, 2, 10, 42
  If spouse_check = "02" then EMReadScreen spousal_reference_number, 2, 4, 33
  EMReadScreen current_memb, 1, 2, 73
  EMReadScreen total_membs, 1, 2, 78
  transmit
Loop until cint(current_memb) = cint(total_membs)

'Jumps to STAT/SHEL.
Call navigate_to_MAXIS_screen("STAT", "SHEL")

'Reads the info off of STAT/SHEL into variables for each type of shelter expense. This is used to autofill the allocation dialog.
EMReadScreen rent_verif, 2, 11, 67
If rent_verif <> "__" and rent_verif <> "NO" and rent_verif <> "?_" then EMReadScreen rent, 8, 11, 56
If rent_verif = "__" or rent_verif = "NO" or rent_verif = "?_" then rent = "0"
EMReadScreen lot_rent_verif, 2, 12, 67
If lot_rent_verif <> "__" and lot_rent_verif <> "NO" and lot_rent_verif <> "?_" then EMReadScreen lot_rent, 8, 12, 56
If lot_rent_verif = "__" or lot_rent_verif = "NO" or lot_rent_verif = "?_" then lot_rent = "0"
EMReadScreen mortgage_verif, 2, 13, 67
If mortgage_verif <> "__" and mortgage_verif <> "NO" and mortgage_verif <> "?_" then EMReadScreen mortgage, 8, 13, 56
If mortgage_verif = "__" or mortgage_verif = "NO" or mortgage_verif = "?_" then mortgage = "0"
EMReadScreen insurance_verif, 2, 14, 67
If insurance_verif <> "__" and insurance_verif <> "NO" and insurance_verif <> "?_" then EMReadScreen insurance, 8, 14, 56
If insurance_verif = "__" or insurance_verif = "NO" or insurance_verif = "?_" then insurance = "0"
EMReadScreen taxes_verif, 2, 15, 67
If taxes_verif <> "__" and taxes_verif <> "NO" and taxes_verif <> "?_" then EMReadScreen taxes, 8, 15, 56
If taxes_verif = "__" or taxes_verif = "NO" or taxes_verif = "?_" then taxes = "0"
EMReadScreen room_verif, 2, 16, 67
If room_verif <> "__" and room_verif <> "NO" and room_verif <> "?_" then EMReadScreen room, 8, 16, 56
If room_verif = "__" or room_verif = "NO" or room_verif = "?_" then room = "0"
EMReadScreen garage_verif, 2, 17, 67
If garage_verif <> "__" and garage_verif <> "NO" and garage_verif <> "?_" then EMReadScreen garage, 8, 17, 56
If garage_verif = "__" or garage_verif = "NO" or garage_verif = "?_" then garage = "0"
EMReadScreen subsidy_verif, 2, 18, 67
If subsidy_verif <> "__" and subsidy_verif <> "NO" and subsidy_verif <> "?_" then EMReadScreen subsidy, 8, 18, 56
If subsidy_verif = "__" or subsidy_verif = "NO" or subsidy_verif = "?_" then subsidy = "0"
mort_rent_payment = cint(rent) + cint(mortgage)
mort_rent_payment = "" & mort_rent_payment
If mort_rent_payment = "0" then mort_rent_payment = ""
taxes_and_insuance = cint(taxes) + cint(insurance)
taxes_and_insuance = "" & taxes_and_insuance
If taxes_and_insuance = "0" then taxes_and_insuance = ""
coop_condo_maint_fees = cint(lot_rent)
coop_condo_maint_fees = "" & coop_condo_maint_fees
If coop_condo_maint_fees = "0" then coop_condo_maint_fees = ""

'Jumps to STAT/HEST
Call navigate_to_MAXIS_screen("STAT", "HEST")

'Reads the info off of STAT/HEST into variables for the utility expenses. This is used to autofill the allocation dialog.
EMReadScreen heat_air, 6, 13, 75
IF trim(heat_air) <> "" then
    utility_allowance =  heat_air
Else
    'electric
    EMReadScreen elec_allow, 6, 14, 75
    elec_allow = trim(elec_allow)
    If elec_allow = "" then elec_allow = "0"
    'phone
    EMReadScreen phone_allow, 6, 15, 75
    phone_allow = trim(phone_allow)
    If phone_allow = "" then phone_allow = "0"
    utility_allowance = cint(elec_allow) + cint(phone_allow) & ""
End if

'Navigates to UNEA and grabs info on the spouse's income (if a spouse is found in MEMB).
If spousal_reference_number <> "" then
  EMWriteScreen "unea", 20, 71
  EMWriteScreen spousal_reference_number, 20, 76
  transmit
  EMReadScreen current_panel_number, 1, 2, 73
  If current_panel_number = "1" then
    EMReadScreen gross_spousal_unearned_income_type_01, 2, 5, 37
    EMReadScreen gross_spousal_unearned_income_01, 8, 18, 68
    transmit
  End if
  EMReadScreen current_panel_number, 1, 2, 73
  If current_panel_number = "2" then
    EMReadScreen gross_spousal_unearned_income_type_02, 2, 5, 37
    EMReadScreen gross_spousal_unearned_income_02, 8, 18, 68
    transmit
  End if
  EMReadScreen current_panel_number, 1, 2, 73
  If current_panel_number = "3" then
    EMReadScreen gross_spousal_unearned_income_type_03, 2, 5, 37
    EMReadScreen gross_spousal_unearned_income_03, 8, 18, 68
    transmit
  End if
  EMReadScreen current_panel_number, 1, 2, 73
  If current_panel_number = "4" then
    EMReadScreen gross_spousal_unearned_income_type_04, 2, 5, 37
    EMReadScreen gross_spousal_unearned_income_04, 8, 18, 68
    transmit
  End if
  EMWriteScreen "jobs", 20, 71
  EMWriteScreen spousal_reference_number, 20, 76
  transmit
  EMReadScreen current_panel_number, 1, 2, 73
'changes unearned income coding types as coding from JOBS panel and spousal allocation screen are not the same
  If current_panel_number = "1" then
	earned_income_number = 1
	EMReadScreen gross_spousal_earned_income_type_01, 1, 5, 34

	If gross_spousal_earned_income_type_01 = "J" THEN gross_spousal_earned_income_type_01 = "01"
	If gross_spousal_earned_income_type_01 = "W" then gross_spousal_earned_income_type_01 = "02"
	If gross_spousal_earned_income_type_01 = "E" THEN gross_spousal_earned_income_type_01 = "03"
	If gross_spousal_earned_income_type_01 = "G" then gross_spousal_earned_income_type_01 = "04"
	If gross_spousal_earned_income_type_01 = "F" THEN gross_spousal_earned_income_type_01 = "05"
	If gross_spousal_earned_income_type_01 = "S" then gross_spousal_earned_income_type_01 = "06"
	If gross_spousal_earned_income_type_01 = "O" THEN gross_spousal_earned_income_type_01 = "07"
	If gross_spousal_earned_income_type_01 = "I" then gross_spousal_earned_income_type_01 = "08"
	If gross_spousal_earned_income_type_01 = "M" THEN gross_spousal_earned_income_type_01 = "09"
	If gross_spousal_earned_income_type_01 = "C" then gross_spousal_earned_income_type_01 = "10"
	If gross_spousal_earned_income_type_01 = "T" then gross_spousal_earned_income_type_01 = "07"
	If gross_spousal_earned_income_type_01 = "P" then gross_spousal_earned_income_type_01 = "07"
	If gross_spousal_earned_income_type_01 = "R" then gross_spousal_earned_income_type_01 = "07"
	EMReadScreen gross_spousal_earned_income_01, 8, 17, 67
 	transmit
  End if
  EMReadScreen current_panel_number, 1, 2, 73
  If current_panel_number = "2" then
    earned_income_number = earned_income_number + 1
	EMReadScreen gross_spousal_earned_income_type_02, 1, 5, 34

	If gross_spousal_earned_income_type_02 = "J" THEN gross_spousal_earned_income_type_02 = "01"
	If gross_spousal_earned_income_type_02 = "W" then gross_spousal_earned_income_type_02 = "02"
	If gross_spousal_earned_income_type_02 = "E" THEN gross_spousal_earned_income_type_02 = "03"
	If gross_spousal_earned_income_type_02 = "G" then gross_spousal_earned_income_type_02 = "04"
	If gross_spousal_earned_income_type_02 = "F" THEN gross_spousal_earned_income_type_02 = "05"
	If gross_spousal_earned_income_type_02 = "S" then gross_spousal_earned_income_type_02 = "06"
	If gross_spousal_earned_income_type_02 = "O" THEN gross_spousal_earned_income_type_02 = "07"
	If gross_spousal_earned_income_type_02 = "I" then gross_spousal_earned_income_type_02 = "08"
	If gross_spousal_earned_income_type_02 = "M" THEN gross_spousal_earned_income_type_02 = "09"
	If gross_spousal_earned_income_type_02 = "C" then gross_spousal_earned_income_type_02 = "10"
	If gross_spousal_earned_income_type_02 = "T" then gross_spousal_earned_income_type_02 = "07"
	If gross_spousal_earned_income_type_02 = "P" then gross_spousal_earned_income_type_02 = "07"
	If gross_spousal_earned_income_type_02 = "R" then gross_spousal_earned_income_type_02 = "07"
    EMReadScreen gross_spousal_earned_income_02, 8, 17, 67
    transmit
  End if
  EMReadScreen current_panel_number, 1, 2, 73
  If current_panel_number = "3" then
    earned_income_number = earned_income_number + 1
    EMReadScreen gross_spousal_earned_income_type_03, 1, 5, 34

	If gross_spousal_earned_income_type_03 = "J" THEN gross_spousal_earned_income_type_03 = "01"
	If gross_spousal_earned_income_type_03 = "W" then gross_spousal_earned_income_type_03 = "02"
	If gross_spousal_earned_income_type_03 = "E" THEN gross_spousal_earned_income_type_03 = "03"
	If gross_spousal_earned_income_type_03 = "G" then gross_spousal_earned_income_type_03 = "04"
	If gross_spousal_earned_income_type_03 = "F" THEN gross_spousal_earned_income_type_03 = "05"
	If gross_spousal_earned_income_type_03 = "S" then gross_spousal_earned_income_type_03 = "06"
	If gross_spousal_earned_income_type_03 = "O" THEN gross_spousal_earned_income_type_03 = "07"
	If gross_spousal_earned_income_type_03 = "I" then gross_spousal_earned_income_type_03 = "08"
	If gross_spousal_earned_income_type_03 = "M" THEN gross_spousal_earned_income_type_03 = "09"
	If gross_spousal_earned_income_type_03 = "C" then gross_spousal_earned_income_type_03 = "10"
	If gross_spousal_earned_income_type_03 = "T" then gross_spousal_earned_income_type_03 = "07"
	If gross_spousal_earned_income_type_03 = "P" then gross_spousal_earned_income_type_03 = "07"
	If gross_spousal_earned_income_type_03 = "R" then gross_spousal_earned_income_type_03 = "07"
    EMReadScreen gross_spousal_earned_income_03, 8, 17, 67
    transmit
  End if
  EMReadScreen current_panel_number, 1, 2, 73
  If current_panel_number = "4" then
    earned_income_number = earned_income_number + 1
	EMReadScreen gross_spousal_earned_income_type_04, 1, 5, 34
	If gross_spousal_earned_income_type_04 = "J" THEN gross_spousal_earned_income_type_04 = "01"
	If gross_spousal_earned_income_type_04 = "W" then gross_spousal_earned_income_type_04 = "02"
	If gross_spousal_earned_income_type_04 = "E" THEN gross_spousal_earned_income_type_04 = "03"
	If gross_spousal_earned_income_type_04 = "G" then gross_spousal_earned_income_type_04 = "04"
	If gross_spousal_earned_income_type_04 = "F" THEN gross_spousal_earned_income_type_04 = "05"
	If gross_spousal_earned_income_type_04 = "S" then gross_spousal_earned_income_type_04 = "06"
	If gross_spousal_earned_income_type_04 = "O" THEN gross_spousal_earned_income_type_04 = "07"
	If gross_spousal_earned_income_type_04 = "I" then gross_spousal_earned_income_type_04 = "08"
	If gross_spousal_earned_income_type_04 = "M" THEN gross_spousal_earned_income_type_04 = "09"
	If gross_spousal_earned_income_type_04 = "C" then gross_spousal_earned_income_type_04 = "10"
	If gross_spousal_earned_income_type_04 = "T" then gross_spousal_earned_income_type_04 = "07"
	If gross_spousal_earned_income_type_04 = "P" then gross_spousal_earned_income_type_04 = "07"
	If gross_spousal_earned_income_type_04 = "R" then gross_spousal_earned_income_type_04 = "07"
    EMReadScreen gross_spousal_earned_income_04, 8, 17, 67
    transmit
  End if
  EMWriteScreen "busi", 20, 71
  EMWriteScreen spousal_reference_number, 20, 76
  transmit
  EMReadScreen current_panel_number, 1, 2, 73
  If current_panel_number <> "0" then
    MsgBox "There is self-employment for the spouse. Add this manually to the calculation. The dialog will pop-up after it checks for RBIC."
    has_self_employment = True
  End if
  EMWriteScreen "rbic", 20, 71
  EMWriteScreen spousal_reference_number, 20, 76
  transmit
  EMReadScreen current_panel_number, 1, 2, 73
  If current_panel_number <> "0" then MsgBox "There is room/board income for the spouse. Add this manually to the calculation. The dialog will pop-up next."
  If current_panel_number = "0" and has_self_employment = True then
    EMWriteScreen "busi", 20, 71
    EMWriteScreen spousal_reference_number, 20, 76
    transmit
  End if
End if

'It should trim up the spaces from the screens.
gross_spousal_unearned_income_01 = trim(gross_spousal_unearned_income_01)
gross_spousal_unearned_income_02 = trim(gross_spousal_unearned_income_02)
gross_spousal_unearned_income_03 = trim(gross_spousal_unearned_income_03)
gross_spousal_unearned_income_04 = trim(gross_spousal_unearned_income_04)
gross_spousal_earned_income_01 = trim(gross_spousal_earned_income_01)
gross_spousal_earned_income_02 = trim(gross_spousal_earned_income_02)
gross_spousal_earned_income_03 = trim(gross_spousal_earned_income_03)
gross_spousal_earned_income_04 = trim(gross_spousal_earned_income_04)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 256, 185, "Spousal Maintenance Dialog"
  EditBox 5, 15, 35, 15, gross_spousal_unearned_income_type_01
  EditBox 60, 15, 50, 15, gross_spousal_unearned_income_01
  EditBox 5, 35, 35, 15, gross_spousal_unearned_income_type_02
  EditBox 60, 35, 50, 15, gross_spousal_unearned_income_02
  EditBox 5, 55, 35, 15, gross_spousal_unearned_income_type_03
  EditBox 60, 55, 50, 15, gross_spousal_unearned_income_03
  EditBox 5, 75, 35, 15, gross_spousal_unearned_income_type_04
  EditBox 60, 75, 50, 15, gross_spousal_unearned_income_04
  EditBox 5, 105, 35, 15, gross_spousal_earned_income_type_01
  EditBox 60, 105, 50, 15, gross_spousal_earned_income_01
  EditBox 5, 125, 35, 15, gross_spousal_earned_income_type_02
  EditBox 60, 125, 50, 15, gross_spousal_earned_income_02
  EditBox 5, 145, 35, 15, gross_spousal_earned_income_type_03
  EditBox 60, 145, 50, 15, gross_spousal_earned_income_03
  EditBox 5, 165, 35, 15, gross_spousal_earned_income_type_04
  EditBox 60, 165, 50, 15, gross_spousal_earned_income_04
  EditBox 205, 85, 35, 15, mort_rent_payment
  EditBox 205, 105, 35, 15, taxes_and_insuance
  EditBox 210, 125, 35, 15, coop_condo_maint_fees
  EditBox 190, 145, 45, 15, utility_allowance
  ButtonGroup ButtonPressed
    OkButton 135, 165, 50, 15
    CancelButton 190, 165, 50, 15
    PushButton 125, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 125, 25, 45, 10, "next panel", next_panel_button
    PushButton 185, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 185, 25, 45, 10, "next memb", next_memb_button
    PushButton 125, 60, 25, 10, "HEST", HEST_button
    PushButton 150, 60, 25, 10, "JOBS", JOBS_button
    PushButton 175, 60, 25, 10, "SHEL", SHEL_button
    PushButton 200, 60, 25, 10, "UNEA", UNEA_button
  Text 10, 5, 25, 10, "UI type"
  Text 75, 5, 30, 10, "Amount"
  Text 10, 95, 25, 10, "EI type"
  Text 75, 95, 30, 10, "Amount"
  GroupBox 120, 5, 115, 35, "STAT-based navigation"
  GroupBox 120, 50, 110, 25, "MAXIS panels"
  Text 130, 90, 65, 10, "Mort-rent payment:"
  Text 130, 110, 75, 10, "Taxes and insurance:"
  Text 130, 130, 80, 10, "Coop-condo maint fees:"
  Text 130, 150, 60, 10, "Utility allowance:"
EndDialog

'Defining the HH_memb_row variable for the navigation buttons
HH_memb_row = 6

Do
    Do
        'Shows the spousal maintenance dialog
        dialog Dialog1
        cancel_confirmation
        MAXIS_dialog_navigation
    Loop until ButtonPressed = -1
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'changes unearned income coding types as coding from UNEA panel and spousal allocation screen are not the same
'unearned income 01
IF gross_spousal_unearned_income_type_01 = "11" THEN gross_spousal_unearned_income_type_01 = "09"
IF gross_spousal_unearned_income_type_01 = "12" THEN gross_spousal_unearned_income_type_01 = "10"
IF gross_spousal_unearned_income_type_01 = "13" THEN gross_spousal_unearned_income_type_01 = "11"
IF gross_spousal_unearned_income_type_01 = "14" THEN gross_spousal_unearned_income_type_01 = "12"
IF gross_spousal_unearned_income_type_01 = "15" THEN gross_spousal_unearned_income_type_01 = "13"
IF gross_spousal_unearned_income_type_01 = "16" THEN gross_spousal_unearned_income_type_01 = "14"
IF gross_spousal_unearned_income_type_01 = "17" THEN gross_spousal_unearned_income_type_01 = "15"
IF gross_spousal_unearned_income_type_01 = "18" THEN gross_spousal_unearned_income_type_01 = "16"
IF gross_spousal_unearned_income_type_01 = "19" THEN gross_spousal_unearned_income_type_01 = "17"
IF gross_spousal_unearned_income_type_01 = "20" THEN gross_spousal_unearned_income_type_01 = "18"
IF gross_spousal_unearned_income_type_01 = "21" THEN gross_spousal_unearned_income_type_01 = "19"
IF gross_spousal_unearned_income_type_01 = "22" THEN gross_spousal_unearned_income_type_01 = "20"
IF gross_spousal_unearned_income_type_01 = "23" THEN gross_spousal_unearned_income_type_01 = "21"
IF gross_spousal_unearned_income_type_01 = "24" THEN gross_spousal_unearned_income_type_01 = "22"
IF gross_spousal_unearned_income_type_01 = "25" THEN gross_spousal_unearned_income_type_01 = "23"
IF gross_spousal_unearned_income_type_01 = "26" THEN gross_spousal_unearned_income_type_01 = "24"
IF gross_spousal_unearned_income_type_01 = "27" THEN gross_spousal_unearned_income_type_01 = "25"
IF gross_spousal_unearned_income_type_01 = "28" THEN gross_spousal_unearned_income_type_01 = "26"
IF gross_spousal_unearned_income_type_01 = "29" THEN gross_spousal_unearned_income_type_01 = "27"
IF gross_spousal_unearned_income_type_01 = "31" THEN gross_spousal_unearned_income_type_01 = "29"
IF gross_spousal_unearned_income_type_01 = "32" THEN gross_spousal_unearned_income_type_01 = "30"
IF gross_spousal_unearned_income_type_01 = "35" THEN gross_spousal_unearned_income_type_01 = "04"
IF gross_spousal_unearned_income_type_01 = "36" THEN gross_spousal_unearned_income_type_01 = "05"
IF gross_spousal_unearned_income_type_01 = "37" THEN gross_spousal_unearned_income_type_01 = "07"
IF gross_spousal_unearned_income_type_01 = "38" THEN gross_spousal_unearned_income_type_01 = "34"
IF gross_spousal_unearned_income_type_01 = "39" THEN gross_spousal_unearned_income_type_01 = "35"
IF gross_spousal_unearned_income_type_01 = "40" THEN gross_spousal_unearned_income_type_01 = "36"
IF gross_spousal_unearned_income_type_01 = "43" THEN gross_spousal_unearned_income_type_01 = "43"
IF gross_spousal_unearned_income_type_01 = "44" THEN gross_spousal_unearned_income_type_01 = "27"
IF gross_spousal_unearned_income_type_01 = "47" THEN gross_spousal_unearned_income_type_01 = "27"
IF gross_spousal_unearned_income_type_01 = "48" THEN gross_spousal_unearned_income_type_01 = "27"
IF gross_spousal_unearned_income_type_01 = "49" THEN gross_spousal_unearned_income_type_01 = "27"

'unearned income 02
IF gross_spousal_unearned_income_type_02 = "11" THEN gross_spousal_unearned_income_type_02 = "09"
IF gross_spousal_unearned_income_type_02 = "12" THEN gross_spousal_unearned_income_type_02 = "10"
IF gross_spousal_unearned_income_type_02 = "13" THEN gross_spousal_unearned_income_type_02 = "11"
IF gross_spousal_unearned_income_type_02 = "14" THEN gross_spousal_unearned_income_type_02 = "12"
IF gross_spousal_unearned_income_type_02 = "15" THEN gross_spousal_unearned_income_type_02 = "13"
IF gross_spousal_unearned_income_type_02 = "16" THEN gross_spousal_unearned_income_type_02 = "14"
IF gross_spousal_unearned_income_type_02 = "17" THEN gross_spousal_unearned_income_type_02 = "15"
IF gross_spousal_unearned_income_type_02 = "18" THEN gross_spousal_unearned_income_type_02 = "16"
IF gross_spousal_unearned_income_type_02 = "19" THEN gross_spousal_unearned_income_type_02 = "17"
IF gross_spousal_unearned_income_type_02 = "20" THEN gross_spousal_unearned_income_type_02 = "18"
IF gross_spousal_unearned_income_type_02 = "21" THEN gross_spousal_unearned_income_type_02 = "19"
IF gross_spousal_unearned_income_type_02 = "22" THEN gross_spousal_unearned_income_type_02 = "20"
IF gross_spousal_unearned_income_type_02 = "23" THEN gross_spousal_unearned_income_type_02 = "21"
IF gross_spousal_unearned_income_type_02 = "24" THEN gross_spousal_unearned_income_type_02 = "22"
IF gross_spousal_unearned_income_type_02 = "25" THEN gross_spousal_unearned_income_type_02 = "23"
IF gross_spousal_unearned_income_type_02 = "26" THEN gross_spousal_unearned_income_type_02 = "24"
IF gross_spousal_unearned_income_type_02 = "27" THEN gross_spousal_unearned_income_type_02 = "25"
IF gross_spousal_unearned_income_type_02 = "28" THEN gross_spousal_unearned_income_type_02 = "26"
IF gross_spousal_unearned_income_type_02 = "29" THEN gross_spousal_unearned_income_type_02 = "27"
IF gross_spousal_unearned_income_type_02 = "31" THEN gross_spousal_unearned_income_type_02 = "29"
IF gross_spousal_unearned_income_type_02 = "32" THEN gross_spousal_unearned_income_type_02 = "30"
IF gross_spousal_unearned_income_type_02 = "35" THEN gross_spousal_unearned_income_type_02 = "04"
IF gross_spousal_unearned_income_type_02 = "36" THEN gross_spousal_unearned_income_type_02 = "05"
IF gross_spousal_unearned_income_type_02 = "37" THEN gross_spousal_unearned_income_type_02 = "07"
IF gross_spousal_unearned_income_type_02 = "38" THEN gross_spousal_unearned_income_type_02 = "34"
IF gross_spousal_unearned_income_type_02 = "39" THEN gross_spousal_unearned_income_type_02 = "35"
IF gross_spousal_unearned_income_type_02 = "40" THEN gross_spousal_unearned_income_type_02 = "36"
IF gross_spousal_unearned_income_type_02 = "43" THEN gross_spousal_unearned_income_type_02 = "43"
IF gross_spousal_unearned_income_type_02 = "44" THEN gross_spousal_unearned_income_type_02 = "27"
IF gross_spousal_unearned_income_type_02 = "47" THEN gross_spousal_unearned_income_type_02 = "27"
IF gross_spousal_unearned_income_type_02 = "48" THEN gross_spousal_unearned_income_type_02 = "27"
IF gross_spousal_unearned_income_type_02 = "49" THEN gross_spousal_unearned_income_type_02 = "27"

'unearned income 03
IF gross_spousal_unearned_income_type_03 = "11" THEN gross_spousal_unearned_income_type_03 = "09"
IF gross_spousal_unearned_income_type_03 = "12" THEN gross_spousal_unearned_income_type_03 = "10"
IF gross_spousal_unearned_income_type_03 = "13" THEN gross_spousal_unearned_income_type_03 = "11"
IF gross_spousal_unearned_income_type_03 = "14" THEN gross_spousal_unearned_income_type_03 = "12"
IF gross_spousal_unearned_income_type_03 = "15" THEN gross_spousal_unearned_income_type_03 = "13"
IF gross_spousal_unearned_income_type_03 = "16" THEN gross_spousal_unearned_income_type_03 = "14"
IF gross_spousal_unearned_income_type_03 = "17" THEN gross_spousal_unearned_income_type_03 = "15"
IF gross_spousal_unearned_income_type_03 = "18" THEN gross_spousal_unearned_income_type_03 = "16"
IF gross_spousal_unearned_income_type_03 = "19" THEN gross_spousal_unearned_income_type_03 = "17"
IF gross_spousal_unearned_income_type_03 = "20" THEN gross_spousal_unearned_income_type_03 = "18"
IF gross_spousal_unearned_income_type_03 = "21" THEN gross_spousal_unearned_income_type_03 = "19"
IF gross_spousal_unearned_income_type_03 = "22" THEN gross_spousal_unearned_income_type_03 = "20"
IF gross_spousal_unearned_income_type_03 = "23" THEN gross_spousal_unearned_income_type_03 = "21"
IF gross_spousal_unearned_income_type_03 = "24" THEN gross_spousal_unearned_income_type_03 = "22"
IF gross_spousal_unearned_income_type_03 = "25" THEN gross_spousal_unearned_income_type_03 = "23"
IF gross_spousal_unearned_income_type_03 = "26" THEN gross_spousal_unearned_income_type_03 = "24"
IF gross_spousal_unearned_income_type_03 = "27" THEN gross_spousal_unearned_income_type_03 = "25"
IF gross_spousal_unearned_income_type_03 = "28" THEN gross_spousal_unearned_income_type_03 = "26"
IF gross_spousal_unearned_income_type_03 = "29" THEN gross_spousal_unearned_income_type_03 = "27"
IF gross_spousal_unearned_income_type_03 = "31" THEN gross_spousal_unearned_income_type_03 = "29"
IF gross_spousal_unearned_income_type_03 = "32" THEN gross_spousal_unearned_income_type_03 = "30"
IF gross_spousal_unearned_income_type_03 = "35" THEN gross_spousal_unearned_income_type_03 = "04"
IF gross_spousal_unearned_income_type_03 = "36" THEN gross_spousal_unearned_income_type_03 = "05"
IF gross_spousal_unearned_income_type_03 = "37" THEN gross_spousal_unearned_income_type_03 = "07"
IF gross_spousal_unearned_income_type_03 = "38" THEN gross_spousal_unearned_income_type_03 = "34"
IF gross_spousal_unearned_income_type_03 = "39" THEN gross_spousal_unearned_income_type_03 = "35"
IF gross_spousal_unearned_income_type_03 = "40" THEN gross_spousal_unearned_income_type_03 = "36"
IF gross_spousal_unearned_income_type_03 = "43" THEN gross_spousal_unearned_income_type_03 = "43"
IF gross_spousal_unearned_income_type_03 = "44" THEN gross_spousal_unearned_income_type_03 = "27"
IF gross_spousal_unearned_income_type_03 = "47" THEN gross_spousal_unearned_income_type_03 = "27"
IF gross_spousal_unearned_income_type_03 = "48" THEN gross_spousal_unearned_income_type_03 = "27"
IF gross_spousal_unearned_income_type_03 = "49" THEN gross_spousal_unearned_income_type_03 = "27"

'unearned income 04
IF gross_spousal_unearned_income_type_04 = "11" THEN gross_spousal_unearned_income_type_04 = "09"
IF gross_spousal_unearned_income_type_04 = "12" THEN gross_spousal_unearned_income_type_04 = "10"
IF gross_spousal_unearned_income_type_04 = "13" THEN gross_spousal_unearned_income_type_04 = "11"
IF gross_spousal_unearned_income_type_04 = "14" THEN gross_spousal_unearned_income_type_04 = "12"
IF gross_spousal_unearned_income_type_04 = "15" THEN gross_spousal_unearned_income_type_04 = "13"
IF gross_spousal_unearned_income_type_04 = "16" THEN gross_spousal_unearned_income_type_04 = "14"
IF gross_spousal_unearned_income_type_04 = "17" THEN gross_spousal_unearned_income_type_04 = "15"
IF gross_spousal_unearned_income_type_04 = "18" THEN gross_spousal_unearned_income_type_04 = "16"
IF gross_spousal_unearned_income_type_04 = "19" THEN gross_spousal_unearned_income_type_04 = "17"
IF gross_spousal_unearned_income_type_04 = "20" THEN gross_spousal_unearned_income_type_04 = "18"
IF gross_spousal_unearned_income_type_04 = "21" THEN gross_spousal_unearned_income_type_04 = "19"
IF gross_spousal_unearned_income_type_04 = "22" THEN gross_spousal_unearned_income_type_04 = "20"
IF gross_spousal_unearned_income_type_04 = "23" THEN gross_spousal_unearned_income_type_04 = "21"
IF gross_spousal_unearned_income_type_04 = "24" THEN gross_spousal_unearned_income_type_04 = "22"
IF gross_spousal_unearned_income_type_04 = "25" THEN gross_spousal_unearned_income_type_04 = "23"
IF gross_spousal_unearned_income_type_04 = "26" THEN gross_spousal_unearned_income_type_04 = "24"
IF gross_spousal_unearned_income_type_04 = "27" THEN gross_spousal_unearned_income_type_04 = "25"
IF gross_spousal_unearned_income_type_04 = "28" THEN gross_spousal_unearned_income_type_04 = "26"
IF gross_spousal_unearned_income_type_04 = "29" THEN gross_spousal_unearned_income_type_04 = "27"
IF gross_spousal_unearned_income_type_04 = "31" THEN gross_spousal_unearned_income_type_04 = "29"
IF gross_spousal_unearned_income_type_04 = "32" THEN gross_spousal_unearned_income_type_04 = "30"
IF gross_spousal_unearned_income_type_04 = "35" THEN gross_spousal_unearned_income_type_04 = "04"
IF gross_spousal_unearned_income_type_04 = "36" THEN gross_spousal_unearned_income_type_04 = "05"
IF gross_spousal_unearned_income_type_04 = "37" THEN gross_spousal_unearned_income_type_04 = "07"
IF gross_spousal_unearned_income_type_04 = "38" THEN gross_spousal_unearned_income_type_04 = "34"
IF gross_spousal_unearned_income_type_04 = "39" THEN gross_spousal_unearned_income_type_04 = "35"
IF gross_spousal_unearned_income_type_04 = "40" THEN gross_spousal_unearned_income_type_04 = "36"
IF gross_spousal_unearned_income_type_04 = "43" THEN gross_spousal_unearned_income_type_04 = "43"
IF gross_spousal_unearned_income_type_04 = "44" THEN gross_spousal_unearned_income_type_04 = "27"
IF gross_spousal_unearned_income_type_04 = "47" THEN gross_spousal_unearned_income_type_04 = "27"
IF gross_spousal_unearned_income_type_04 = "48" THEN gross_spousal_unearned_income_type_04 = "27"
IF gross_spousal_unearned_income_type_04 = "49" THEN gross_spousal_unearned_income_type_04 = "27"

'Navigates to ELIG/HC.
Call navigate_to_MAXIS_screen("ELIG", "HC")

'Checks to see if MEMB 01 has HC, and puts an "x" there. If not it'll try MEMB 02.
EMReadScreen person_check, 1, 8, 26
If person_check <> "_" then
  MsgBox "Person 01 does not have HC on this case. The script will attempt to execute this on person 02. Please check this for errors before approving any results."
  EMWriteScreen "x", 9, 26
End if
If person_check = "_" then EMWriteScreen "x", 8, 26
'Gets into ELIG/HC for that particular member.
transmit

'Turns on FIAT mode, checks to make sure it worked and sets the reason as "06".
PF9
EMReadScreen FIAT_check, 4, 24, 45
If FIAT_check <> "FIAT" then
  EMSendKey "06"
  transmit
End if

footer_month_and_year = MAXIS_footer_month & "/01/" & MAXIS_footer_year 'defining footer month/year as date 
new_elig_hc_panel_date = "01/01/25"                   'defining date ELIG/HC panel format/positions changed

If DateDiff("D", new_elig_hc_panel_date, footer_month_and_year) < 0 Then  'Panel prior to 1/1/25
  'Defining the variables for the following search
  ELIG_HC_row = 6
  ELIG_HC_col = 1
End If

If DateDiff("D", new_elig_hc_panel_date, footer_month_and_year) >= 0 THEN 'Panel on/after 1/1/25
  'Defining the variables for the following search
  ELIG_HC_row = 5
  ELIG_HC_col = 1
End If


'Determining the col variable based on the indicated footer month/year
EMSearch spousal_allocation_footer_month & "/" & spousal_allocation_footer_year, ELIG_HC_row, ELIG_HC_col
If ELIG_HC_col = 0 then script_end_procedure("Requested footer month not found. You may have entered ELIG/HC in an invalid footer month, or results haven't been generated for that month. Check these out and try again.")
col = ELIG_HC_col + 1

'setting the variable for the next do...loop
budget_months = 0

'Fills all budget months with "x's", so that the script will go into each one in succession.
Do
  If DateDiff("D", new_elig_hc_panel_date, footer_month_and_year) < 0 Then  'Panel prior to 1/1/25
    EMReadScreen budget_check, 1, 12, col
  End If 

  If DateDiff("D", new_elig_hc_panel_date, footer_month_and_year) >= 0 THEN 'Panel on/after 1/1/25
    EMReadScreen budget_check, 1, 11, col
  End If

  If budget_check = "/" then
    budget_months = budget_months + 1
    If DateDiff("D", new_elig_hc_panel_date, footer_month_and_year) < 0 Then  'Panel prior to 1/1/25
      EMWriteScreen "x", 9, col + 1
    End If

    If DateDiff("D", new_elig_hc_panel_date, footer_month_and_year) >= 0 THEN 'Panel on/after 1/1/25
      EMWriteScreen "x", 8, col + 1
    End If
  End if
  col = col + 11
Loop until col > 75

'Jumps into the budget screen (LBUD or SBUD)
transmit

For amt_of_months_to_do = 1 to budget_months

  'For an unknown (as of 06/24/2013) reason, some cases seem to stay in the first budget month and not move on. This gathers the current month to see if we've moved past it later.
  EMReadScreen starting_bdgt_month, 5, 6, 14

  'Checks to see if this is an LBUD or SBUD. It'll stop if it's neither.
  EMReadScreen LBUD_check, 4, 3, 45
  If LBUD_check <> "LBUD" then
    EMReadScreen SBUD_check, 4, 3, 44
    If SBUD_check <> "SBUD" then script_end_procedure("This is not a method L or method S case.")
  End if

  'Transmits to the next screen after putting an "x" on the "spousal allocation" line. This is different on LBUD and SBUD, so the script includes logic.
  If LBUD_check = "LBUD" then EMWriteScreen "x", 8, 44
  If SBUD_check = "SBUD" then EMWriteScreen "x", 9, 44
  transmit

  'Puts an "x" on the Gross Unearned Income and transmits
  EMWriteScreen "x", 5, 5
  transmit

  'Blanks out what's already there, then writes the unearned income type, and income amount, and whether or not it's excluded. Repeats for up to four income types.
  EMWriteScreen "__", 8, 8
  EMWriteScreen gross_spousal_unearned_income_type_01, 8, 8
  EMWriteScreen "__________", 8, 43
  EMWriteScreen gross_spousal_unearned_income_01, 8, 43
  EMWriteScreen "_", 8, 58
  If gross_spousal_unearned_income_type_01 <> "" then
    If gross_spousal_unearned_income_excluded_check_01 = 1 then EMWriteScreen "Y", 8, 58
    If gross_spousal_unearned_income_excluded_check_01 <> 1 then EMWriteScreen "N", 8, 58
  End if
  EMWriteScreen "__", 9, 8
  EMWriteScreen gross_spousal_unearned_income_type_02, 9, 8
  EMWriteScreen "__________", 9, 43
  EMWriteScreen gross_spousal_unearned_income_02, 9, 43
  EMWriteScreen "_", 9, 58
  If gross_spousal_unearned_income_type_02 <> "" then
    If gross_spousal_unearned_income_excluded_check_02 = 1 then EMWriteScreen "Y", 9, 58
    If gross_spousal_unearned_income_excluded_check_02 <> 1 then EMWriteScreen "N", 9, 58
  End if
  EMWriteScreen "__", 10, 8
  EMWriteScreen gross_spousal_unearned_income_type_03, 10, 8
  EMWriteScreen "__________", 10, 43
  EMWriteScreen gross_spousal_unearned_income_03, 10, 43
  EMWriteScreen "_", 10, 58
  If gross_spousal_unearned_income_type_03 <> "" then
    If gross_spousal_unearned_income_excluded_check_03 = 1 then EMWriteScreen "Y", 10, 58
    If gross_spousal_unearned_income_excluded_check_03 <> 1 then EMWriteScreen "N", 10, 58
  End if
  EMWriteScreen "__", 11, 8
  EMWriteScreen gross_spousal_unearned_income_type_04, 11, 8
  EMWriteScreen "__________", 11, 43
  EMWriteScreen gross_spousal_unearned_income_04, 11, 43
  EMWriteScreen "_", 11, 58
  If gross_spousal_unearned_income_type_04 <> "" then
    If gross_spousal_unearned_income_excluded_check_04 = 1 then EMWriteScreen "Y", 11, 58
    If gross_spousal_unearned_income_excluded_check_04 <> 1 then EMWriteScreen "N", 11, 58
  End if

  'Gets out of unearned and heads into earned income
  PF3
  EMWriteScreen "x", 6, 5
  transmit


  'Blanks out what's already there, then writes the earned income type, and income amount, and whether or not it's excluded. Repeats for up to four income types.
  EMWriteScreen "__", 8, 8
  EMWriteScreen gross_spousal_earned_income_type_01, 8, 8
  EMWriteScreen "___________", 8, 43
  EMWriteScreen gross_spousal_earned_income_01, 8, 43
  EMWriteScreen "_", 8, 59
  If gross_spousal_earned_income_type_01 <> "" then
    If gross_spousal_earned_income_excluded_check_01 = 1 then EMWriteScreen "Y", 8, 59
    If gross_spousal_earned_income_excluded_check_01 <> 1 then EMWriteScreen "N", 8, 59
  End if
  EMWriteScreen "__", 9, 8
  EMWriteScreen gross_spousal_earned_income_type_02, 9, 8
  EMWriteScreen "___________", 9, 43
  EMWriteScreen gross_spousal_earned_income_02, 9, 43
  EMWriteScreen "_", 9, 59
  If gross_spousal_earned_income_type_02 <> "" then
    If gross_spousal_earned_income_excluded_check_02 = 1 then EMWriteScreen "Y", 9, 59
    If gross_spousal_earned_income_excluded_check_02 <> 1 then EMWriteScreen "N", 9, 59
  End if
  EMWriteScreen "__", 10, 8
  EMWriteScreen gross_spousal_earned_income_type_03, 10, 8
  EMWriteScreen "___________", 10, 43
  EMWriteScreen gross_spousal_earned_income_03, 10, 43
  EMWriteScreen "_", 10, 59
  If gross_spousal_earned_income_type_03 <> "" then
    If gross_spousal_earned_income_excluded_check_03 = 1 then EMWriteScreen "Y", 10, 59
    If gross_spousal_earned_income_excluded_check_03 <> 1 then EMWriteScreen "N", 10, 59
  End if
  EMWriteScreen "__", 11, 8
  EMWriteScreen gross_spousal_earned_income_type_04, 11, 8
  EMWriteScreen "___________", 11, 43
  EMWriteScreen gross_spousal_earned_income_04, 11, 43
  EMWriteScreen "_", 11, 59
  If gross_spousal_earned_income_type_04 <> "" then
    If gross_spousal_earned_income_excluded_check_04 = 1 then EMWriteScreen "Y", 11, 59
    If gross_spousal_earned_income_excluded_check_04 <> 1 then EMWriteScreen "N", 11, 59
  End if
  'Gets out of earned income
  PF3

  'Blanks out Mort/Rent Payment, Taxes & Insurance, Coop/Condo Maint Fees and Utility Allowance, writes in the amounts from the dialog
  EMWriteScreen "__________", 9, 33
  EMWriteScreen mort_rent_payment, 9, 33
  EMWriteScreen "__________", 10, 33
  EMWriteScreen taxes_and_insuance, 10, 33
  EMWriteScreen "__________", 11, 33
  EMWriteScreen coop_condo_maint_fees, 11, 33
  EMWriteScreen "__________", 12, 33
  EMWriteScreen utility_allowance, 12, 33

  'Transmits to get past the spousal allocation screen
  transmit

  'Checks to see if the ACT maint needs check box popped up. If it did, it'll transmit to get past that
  EMReadScreen ACT_maint_needs_check, 3, 17, 4
  If ACT_maint_needs_check = "ACT" then transmit
  'Transmits to jump to the next month.
  transmit
  'For an unknown (as of 06/24/2013) reason, some cases seem to stay in the first budget month and not move on. This is a fix for that.
  EMReadScreen ending_bdgt_month, 5, 6, 14
  If starting_bdgt_month = ending_bdgt_month then transmit

  'Resets the variables to check on the next month.
  LBUD_check = ""
  SBUD_check = ""
next

script_end_procedure("")
