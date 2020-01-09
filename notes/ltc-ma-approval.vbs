'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - LTC - MA APPROVAL.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 300          'manual run time in seconds
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
call changelog_update("01/09/2020", "Added additional mandatory fields in dialogs, and updated back end handling.", "Ilse Ferris, Hennepin County")
call changelog_update("12/11/2019", "Added COLA as an option in the special header droplist", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""    'Connects to BlueZone
Call MAXIS_case_number_finder(MAXIS_case_number) 'Grabbing case number & footer month/year
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 161, 61, "Case number"
  Text 5, 5, 85, 10, "Enter your case number:"
  EditBox 95, 0, 60, 15, MAXIS_case_number
  Text 15, 25, 50, 10, "Footer month:"
  EditBox 65, 20, 25, 15, MAXIS_footer_month
  Text 95, 25, 20, 10, "Year:"
  EditBox 120, 20, 25, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 25, 40, 50, 15
    CancelButton 85, 40, 50, 15
EndDialog

DO
	DO
		err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
		dialog Dialog1				'main dialog
		cancel_without_confirmation
		IF len(MAXIS_case_number) > 8 or isnumeric(MAXIS_case_number) = false THEN err_msg = err_msg & vbCr & "* Enter a valid case number."		'mandatory fields
        If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid 2-digit MAXIS footer month."
        If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid 2-digit MAXIS footer year."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Sends transmit to check for MAXIS
Call check_for_MAXIS(FALSE)
Call MAXIS_footer_month_confirmation
Call navigate_to_MAXIS_screen("ELIG", "HC  ")

'Checks to make sure it's in HCMI, due to error prone cases
call navigate_to_MAXIS_screen("STAT", "HCMI")
EMReadScreen HCMI_check, 4, 2, 55
If HCMI_check <> "HCMI" then transmit

'Checks the spenddown option. If one is indicated it will navigate to FACI and pull the current FACI into the designated provider box. If no FACI is given it will generate a warning message to the worker to check MMIS.
EMReadScreen spenddown_option, 2, 10, 57
If spenddown_option <> "__" then
  call autofill_editbox_from_MAXIS(array("01"), "FACI", FACI)
  If FACI = "" then
    MsgBox "A current facility could not be found. Check MMIS for a designated provider."
  Else
    designated_provider = FACI
  End if
End if

'Jumps into ELIG/HC for the footer month listed earlier
Call navigate_to_MAXIS_screen("ELIG", "HC__")

'Checks to see if MEMB 01 has a case. If not it'll try MEMB 02. If that doesn't work the script will error out on its own due to MAXIS intervention.
EMReadScreen person_check, 2, 8, 31
If person_check = "NO" then
  MsgBox "Person 01 does not have HC on this case. The script will attempt to execute this on person 02. Please check this for errors before approving any results."
  EMWriteScreen "x", 9, 26
End if
If person_check <> "NO" then EMWriteScreen "x", 8, 26
transmit

'Searching for the footer month span after defining row/col variables. If a span can't be found the script will shut down.
row = 1
col = 1
EMSearch " " & MAXIS_footer_month & "/" & MAXIS_footer_year & " ", row, col
If row = 0 then script_end_procedure("A " & MAXIS_footer_month & "/" & MAXIS_footer_year & " span could not be found. Try this again. You may need to run the case through background.")

'Grabbing the elig type and budget type
EMReadScreen elig_type, 2, 12, col - 1
EMReadScreen budget_type, 1, 13, col + 3
If (budget_type <> "L" AND budget_type <> "S" AND budget_type <> "B") THEN
	script_end_procedure ("This case is not a L, S or B budget case.  Use the ""Approved Programs"" script instead.")
END if

'Transmitting into the budget breakdown screen
EMWriteScreen "x", 9, col + 3
transmit

'Checking to see if this is an LBUD. If so, it'll grab the info from the appropriate places.
EMReadScreen LBUD_check, 4, 3, 45
If LBUD_check = "LBUD" then
  EMReadScreen recipient_amt, 10, 15, 70
  recipient_amt = "$" & trim(recipient_amt)
  EMReadScreen income, 10, 12, 32
  income = "$" & trim(income)
  EMReadScreen LTC_exclusions, 10, 14, 32
  If LTC_exclusions <> "__________" then deductions = deductions & "LTC exclusions ($" & replace(LTC_exclusions, "_", "") & "). "
  EMReadScreen medicare_premium, 10, 15, 32
  If medicare_premium <> "__________" then deductions = deductions & "Medicare ($" & replace(medicare_premium, "_", "") & "). "
  EMReadScreen pers_cloth_needs, 10, 16, 32
  If pers_cloth_needs <> "__________" then deductions = deductions & "Personal needs ($" & replace(pers_cloth_needs, "_", "") & "). "
  EMReadScreen home_maintenance_allowance, 10, 17, 32
  If home_maintenance_allowance <> "__________" then deductions = deductions & "Home maintenance allowance ($" & replace(home_maintenance_allowance, "_", "") & "). "
  EMReadScreen guard_rep_payee_fee, 10, 18, 32
  If guard_rep_payee_fee <> "__________" then deductions = deductions & "Payee fee ($" & replace(guard_rep_payee_fee, "_", "") & "). "
  EMReadScreen spousal_allocation, 10, 8, 70
  If spousal_allocation <> "          " then deductions = deductions & "Spousal allocation ($" & replace(spousal_allocation, " ", "") & "). "
  EMReadScreen family_allocation, 10, 9, 70
  If family_allocation <> "__________" then deductions = deductions & "Family allocation ($" & replace(family_allocation, "_", "") & "). "
  EMReadScreen health_ins_premium, 10, 10, 70
  If health_ins_premium <> "__________" then deductions = deductions & "Health insurance premium ($" & replace(health_ins_premium, "_", "") & "). "
  EMReadScreen other_med_expense, 10, 11, 70
  If other_med_expense <> "__________" then deductions = deductions & "Other medical expense ($" & replace(other_med_expense, "_", "") & "). "
  EMReadScreen SSI_1611_benefits, 10, 12, 70
  If SSI_1611_benefits <> "__________" then deductions = deductions & "SSI 1611 benefits ($" & replace(SSI_1611_benefits, "_", "") & "). "
  EMReadScreen other_deductions, 10, 13, 70
  If other_deductions <> "__________" then deductions = deductions & "Other deductions ($" & replace(other_deductions, "_", "") & "). "
End if

'Now it checks to see if this is an SBUD. If so, it'll grab the info from the appropriate places.
EMReadScreen SBUD_check, 4, 3, 44
If SBUD_check = "SBUD" then
  EMReadScreen recipient_amt, 10, 16, 71
  recipient_amt = "$" & trim(recipient_amt)
  EMReadScreen income, 10, 13, 32
  income = "$" & trim(income)
  EMReadScreen LTC_exclusions, 10, 15, 32
  If LTC_exclusions <> "__________" then deductions = deductions & "LTC exclusions ($" & replace(LTC_exclusions, "_", "") & "). "
  EMReadScreen medicare_premium, 10, 16, 32
  If medicare_premium <> "__________" then deductions = deductions & "Medicare ($" & replace(medicare_premium, "_", "") & "). "
  EMReadScreen pers_cloth_needs, 10, 17, 32
  If pers_cloth_needs <> "__________" then deductions = deductions & "Maintenance needs allowance ($" & replace(pers_cloth_needs, "_", "") & "). "
  EMReadScreen guard_rep_payee_fee, 10, 18, 32
  If guard_rep_payee_fee <> "__________" then deductions = deductions & "Payee fee ($" & replace(guard_rep_payee_fee, "_", "") & "). "
  EMReadScreen spousal_allocation, 10, 9, 71
  If spousal_allocation <> "          " then deductions = deductions & "Spousal allocation ($" & replace(spousal_allocation, " ", "") & "). "
  EMReadScreen family_allocation, 10, 10, 71
  If family_allocation <> "__________" then deductions = deductions & "Family allocation ($" & replace(family_allocation, "_", "") & "). "
  EMReadScreen health_ins_premium, 10, 11, 71
  If health_ins_premium <> "__________" then deductions = deductions & "Health insurance premium ($" & replace(health_ins_premium, "_", "") & "). "
  EMReadScreen other_med_expense, 10, 12, 71
  If other_med_expense <> "__________" then deductions = deductions & "Other medical expense ($" & replace(other_med_expense, "_", "") & "). "
  EMReadScreen SSI_1611_benefits, 10, 13, 71
  If SSI_1611_benefits <> "__________" then deductions = deductions & "SSI 1611 benefits ($" & replace(SSI_1611_benefits, "_", "") & "). "
  EMReadScreen other_deductions, 10, 14, 71
  If other_deductions <> "__________" then deductions = deductions & "Other deductions ($" & replace(other_deductions, "_", "") & "). "
End if

'Now it checks to see if this is an EBUD. If so, it'll grab the info from the appropriate places.
EMReadScreen EBUD_check, 4, 3, 60
If EBUD_check = "EBUD" then
  EMReadScreen income, 10, 9, 69
  income = "$" & trim(income)
  EMReadScreen MA_EPD_premium, 10, 13, 69
  other = "MA-EPD premium is $" & trim(MA_EPD_premium) & "/mo."
End if

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 191, 76, "BBUD"
  Text 5, 10, 180, 10, "This is a method B budget. What would you like to do?"
  ButtonGroup ButtonPressed
    PushButton 20, 25, 70, 15, "Jump to STAT/BILS", BILS_button
    PushButton 100, 25, 70, 15, "Stay in ELIG/HC", ELIG_button
    CancelButton 135, 55, 50, 15
EndDialog

'Now it checks to see if this is a BBUD. If so, it'll read the info, then offer the worker the chance to navigate to BILS
EMReadScreen BBUD_check, 4, 3, 47
If BBUD_check = "BBUD" then
    EMReadScreen income, 10, 12, 32
    income = "$" & trim(income)
    Do
        Dialog Dialog1
        If ButtonPressed = ELIG_button then exit do
        If ButtonPressed = BILS_Button then 
            PF3
            Call navigate_to_MAXIS_screen("STAT", "BILS")
            EMReadScreen BILS_check, 4, 2, 54
            If BILS_check <> "BILS" then transmit
            exit do
        End if 
  	    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back inn
    
    'If ButtonPressed = 4 then
    '    'BILS_button navigation 
    '    PF3
    ''    Do
    ''        Dialog Dialog1
    ''        cancel_confirmation
    ''        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    ''    Loop until are_we_passworded_out = false					'loops until user passwords back inn
    ''End if
    'Call navigate_to_MAXIS_screen("STAT", "BILS")
    'EMReadScreen BILS_check, 4, 2, 54
    'If BILS_check <> "BILS" then transmit
End if

'auto-fills and cleans up information to entered into the approval_dialog
If recipient_amt = "$" or recipient_amt = "" then recipient_amt = "$0"
If income = "$" or income = "" then income = "$0"
If deductions = "$" or deductions = "" then deductions = "$0"

BeginDialog Dialog1, 0, 0, 376, 165, "Approval dialog"
  DropListBox 45, 5, 30, 15, "AX"+chr(9)+"EX"+chr(9)+"DX"+chr(9)+"DP", elig_type
  DropListBox 135, 5, 30, 15, "L"+chr(9)+"S"+chr(9)+"B", budget_type
  EditBox 285, 5, 85, 15, recipient_amt
  EditBox 90, 25, 280, 15, income
  EditBox 50, 45, 320, 15, deductions
  CheckBox 5, 65, 70, 10, "Updated RSPD?", updated_RSPD_check
  CheckBox 75, 65, 110, 10, "Approved new MAXIS results?", approved_check
  CheckBox 190, 65, 70, 10, "Sent DHS-3050?", DHS_3050_check
  CheckBox 5, 80, 125, 15, "Sent DHS-5181 to Case Manager", sent_5181_check
  EditBox 75, 100, 140, 15, designated_provider
  EditBox 75, 120, 295, 15, other
  DropListBox 60, 145, 60, 15, "None"+chr(9)+"COLA"+chr(9)+"Paperless IR"+chr(9)+"HRF", special_header_droplist
  EditBox 190, 145, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 265, 145, 50, 15
    CancelButton 320, 145, 50, 15
    PushButton 220, 100, 35, 10, "ELIG/HC", ELIG_HC_button
    PushButton 265, 100, 25, 10, "BILS", BILS_button
    PushButton 290, 100, 25, 10, "FACI", FACI_button
    PushButton 315, 100, 25, 10, "HCMI", HCMI_button
    PushButton 340, 100, 25, 10, "UNEA", UNEA_button
  Text 5, 10, 35, 10, "Elig type:"
  Text 85, 10, 45, 10, "Budget type:"
  Text 175, 10, 110, 10, "Waiver obilgation/recipient amt:"
  Text 5, 30, 80, 10, "Total countable income:"
  Text 5, 50, 45, 10, "Deductions:"
  GroupBox 260, 90, 110, 25, "STAT based navigation"
  Text 5, 105, 70, 10, "Designated provider:"
  Text 5, 125, 65, 10, "Other (if applicable):"
  Text 5, 150, 55, 10, "Special header:"
  Text 130, 150, 60, 10, "Worker signature:"
EndDialog

'Shows the MA approval dialog, checks for MAXIS and allows navigation buttons
Do
    Do
    	err_msg = ""
    	Dialog Dialog1
    	cancel_confirmation
    	MAXIS_dialog_navigation
        If trim(recipient_amt) = "" then err_msg = err_msg & "Enter recipient amount, even if amount is 0."
        If trim(income) = "" then err_msg = err_msg & "Enter income information, even if amount is 0."
        If trim(deductions) = "" then err_msg = err_msg & "Enter deductions, even if amount is 0."
    	If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "Sign your case note."
    	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    LOOP until err_msg = ""
CALL check_for_password_without_transmit(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = false

'Variables for the case note
If special_header_droplist = "None" then note_header = " for "
If special_header_droplist = "COLA" then note_header = " for COLA for "
If special_header_droplist = "HRF" then note_header = " for HRF "
If special_header_droplist = "Paperless IR" then note_header = " for paperless IR for "

If budget_type = "L" then SD_type = " LTC SD**"
If budget_type = "S" then SD_type = " SISEW waiver obl**"
If budget_type = "B" then SD_type = " recip amt**"

'THE CASE NOTE----------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("**Approved " & elig_type & "-" & budget_type & note_header & MAXIS_footer_month & "/" & MAXIS_footer_year & ", " & recipient_amt & SD_type)
call write_bullet_and_variable_in_case_note ("Income", income)
call write_bullet_and_variable_in_case_note ("Deductions", deductions)
call write_variable_in_case_note ("")
If updated_RSPD_check = 1 then call write_variable_in_case_note("* Updated RSPD in MMIS.")
call write_bullet_and_variable_in_case_note ("Designated provider", designated_provider)
If approved_check = 1 then call write_variable_in_case_note ("* Approved new MAXIS results.")
If DHS_3050_check = 1 then call write_variable_in_case_note ("* Sent DHS-3050 LTC communication form to facility.")
IF sent_5181_check = 1 then call write_variable_in_case_note ("* Sent DHS-5181 LTC communication to Case Manager")
call write_bullet_and_variable_in_case_note ("Other", other)
call write_variable_in_case_note ("---")
call write_variable_in_case_note (worker_signature)

script_end_procedure("")