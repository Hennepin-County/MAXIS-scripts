'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - EMERGENCY.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 480          'manual run time in seconds
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
call changelog_update("12/28/2019", "Updated EGA screening determination when emer has been used before, but the elig period has expired.", "Ilse Ferris, Hennepin County")
call changelog_update("04/11/2019", "Updated backend processing.", "Ilse Ferris, Hennepin County")
call changelog_update("04/08/2019", "Updated 200% FPG for 2019.", "Ilse Ferris, Hennepin County")
call changelog_update("10/22/2018", "Updated EGA eligibilty period to a year and a day after the start of the eligibilty period, per EGA group.", "Ilse Ferris, Hennepin County")
call changelog_update("09/01/2018", "FPG standards updated.", "Ilse Ferris, Hennepin County")
call changelog_update("03/06/2018", "FPG standards updated.", "Ilse Ferris, Hennepin County")
call changelog_update("09/25/2017", "Fixed header for case notes that have an EGA screening, header information was duplicating.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
'creating month variable 13 months prior to current footer month/year to search for EMER programs issued (for EMER SCREENING portion of the script)
begin_search_month = dateadd("m", -13, date)
begin_search_year = datepart("yyyy", begin_search_month)
begin_search_year = right(begin_search_year, 2)
begin_search_month = datepart("m", begin_search_month)
If len(begin_search_month) = 1 then begin_search_month = "0" & begin_search_month
'End of date calculations----------------------------------------------------------------------------------------------

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 141, 115, "Case number dialog"
  EditBox 75, 5, 55, 15, MAXIS_case_number
  EditBox 75, 25, 25, 15, MAXIS_footer_month
  EditBox 105, 25, 25, 15, MAXIS_footer_year
  CheckBox 10, 60, 30, 10, "cash", cash_check
  CheckBox 55, 60, 30, 10, "HC", HC_check
  CheckBox 95, 60, 35, 10, "SNAP", SNAP_check
  CheckBox 10, 80, 120, 10, "Check here if program is EGA?", EGA_screening_check
  ButtonGroup ButtonPressed
    OkButton 15, 95, 50, 15
    CancelButton 75, 95, 50, 15
  Text 10, 30, 65, 10, "Footer month/year:"
  GroupBox 5, 45, 130, 30, "Other programs open or applied for:"
  Text 25, 10, 45, 10, "Case number:"
EndDialog

'This dialog contains a customized "percent rule" variable, as well as a customized "income days" variable. As such, it can't directly be edited in the dialog editor.
BeginDialog emergency_dialog, 0, 0, 326, 395, "Emergency Dialog"
  EditBox 60, 45, 65, 15, interview_date
  EditBox 170, 45, 150, 15, HH_comp
  CheckBox 25, 75, 40, 10, "Eviction", eviction_check
  CheckBox 75, 75, 70, 10, "Utility disconnect", utility_disconnect_check
  CheckBox 155, 75, 60, 10, "Homelessness", homelessness_check
  CheckBox 230, 75, 65, 10, "Security deposit", security_deposit_check
  EditBox 65, 100, 255, 15, cause_of_crisis
  EditBox 85, 160, 235, 15, income
  EditBox 105, 180, 215, 15, income_under_200_FPG
  EditBox 60, 200, 260, 15, percent_rule_notes
  EditBox 70, 220, 250, 15, monthly_expense
  EditBox 55, 240, 265, 15, assets
  EditBox 55, 260, 265, 15, verifs_needed
  EditBox 80, 280, 240, 15, crisis_resolvable
  EditBox 80, 300, 240, 15, discussion_of_crisis
  EditBox 55, 320, 265, 15, actions_taken
  EditBox 55, 340, 265, 15, referrals
  CheckBox 5, 360, 90, 10, "Sent forms to AREP?", sent_arep_checkbox
  EditBox 70, 375, 140, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 215, 375, 50, 15
    CancelButton 270, 375, 50, 15
    PushButton 10, 15, 25, 10, "ADDR", ADDR_button
    PushButton 35, 15, 25, 10, "MEMB", MEMB_button
    PushButton 60, 15, 25, 10, "MEMI", MEMI_button
    PushButton 10, 25, 25, 10, "PROG", PROG_button
    PushButton 35, 25, 25, 10, "TYPE", TYPE_button
    PushButton 125, 20, 50, 10, "ELIG/EMER", ELIG_EMER_button
    PushButton 210, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 210, 25, 45, 10, "next panel", next_panel_button
    PushButton 270, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 270, 25, 45, 10, "next memb", next_memb_button
    PushButton 75, 130, 25, 10, "BUSI", BUSI_button
    PushButton 100, 130, 25, 10, "JOBS", JOBS_button
    PushButton 75, 140, 25, 10, "RBIC", RBIC_button
    PushButton 100, 140, 25, 10, "UNEA", UNEA_button
    PushButton 150, 130, 25, 10, "ACCT", ACCT_button
    PushButton 175, 130, 25, 10, "CARS", CARS_button
    PushButton 200, 130, 25, 10, "CASH", CASH_button
    PushButton 225, 130, 25, 10, "OTHR", OTHR_button
    PushButton 150, 140, 25, 10, "REST", REST_button
    PushButton 175, 140, 25, 10, "SECU", SECU_button
    PushButton 200, 140, 25, 10, "TRAN", TRAN_button
  GroupBox 5, 5, 85, 35, "other STAT panels:"
  GroupBox 205, 5, 115, 35, "STAT-based navigation"
  Text 5, 50, 50, 10, "Interview date:"
  Text 130, 50, 35, 10, "HH Comp:"
  GroupBox 20, 65, 280, 25, "Crisis (check all that apply):"
  Text 5, 105, 55, 10, "Cause of crisis:"
  GroupBox 70, 120, 60, 35, "Income panels"
  GroupBox 145, 120, 110, 35, "Asset panels"
  Text 5, 165, 75, 10, "Income (past " & emer_number_of_income_days & " days):"
  Text 5, 185, 100, 10, "Is income under 200% FPG?:"
  Text 5, 205, 55, 10, emer_percent_rule_amt & "% rule notes:"
  Text 5, 225, 60, 10, "Monthly expense:"
  Text 25, 245, 25, 10, "Assets:"
  Text 5, 265, 50, 10, "Verifs needed:"
  Text 5, 285, 65, 10, "Crisis resolvable?:"
  Text 5, 305, 70, 10, "Discussion of Crisis:"
  Text 5, 325, 50, 10, "Actions taken:"
  Text 15, 345, 35, 10, "Referrals:"
  Text 5, 380, 65, 10, "Worker signature:"
EndDialog

BeginDialog case_note_dialog, 0, 0, 136, 51, "Case note dialog"
  ButtonGroup ButtonPressed
    PushButton 15, 20, 105, 10, "Yes, take me to case note.", yes_case_note_button
    PushButton 5, 35, 125, 10, "No, take me back to the script dialog.", no_case_note_button
  Text 10, 5, 125, 10, "Are you sure you want to case note?"
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number & footer month/year
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)
CALL MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'Showing the case number dialog
DO
	Do
        err_msg = ""
		Dialog case_number_dialog
		cancel_without_confirmation
        Call validate_MAXIS_case_number(err_msg, "*")
        If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid 2-digit MAXIS footer month."
        If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid 2-digit MAXIS footer year."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    LOOP UNTIL err_msg = ""
CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'EMER screnning code----------------------------------------------------------------------------------------------------
If EGA_screening_check = 1 then
    'EGA screening dialog
    BeginDialog emergency_screening_dialog, 0, 0, 286, 170, "Emergency Screening dialog"
      ComboBox 255, 5, 25, 15, "1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"5"+chr(9)+"6"+chr(9)+"7"+chr(9)+"8"+chr(9)+"9"+chr(9)+"10"+chr(9)+"11"+chr(9)+"12"+chr(9)+"13"+chr(9)+"14"+chr(9)+"15"+chr(9)+"16"+chr(9)+"17"+chr(9)+"18"+chr(9)+"19"+chr(9)+"20", HH_members
      CheckBox 15, 45, 40, 10, "Eviction", eviction_check
      CheckBox 65, 45, 70, 10, "Utility disconnect", utility_disconnect_check
      CheckBox 140, 45, 60, 10, "Homelessness", homelessness_check
      CheckBox 210, 45, 65, 10, "Security deposit", security_deposit_check
      ComboBox 230, 65, 50, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No", meets_residency
      EditBox 230, 85, 50, 15, shelter_costs
      EditBox 230, 105, 50, 15, net_income
      EditBox 155, 125, 125, 15, worker_signature
      ButtonGroup ButtonPressed
        PushButton 85, 145, 90, 15, "HSR Manual EMER page ", EMER_HSR_manual_button
        OkButton 180, 145, 50, 15
        CancelButton 230, 145, 50, 15
        PushButton 10, 95, 30, 10, "ADDR", ADDR_button
        PushButton 40, 95, 30, 10, "BUSI", BUSI_button
        PushButton 10, 105, 30, 10, "JOBS", JOBS_button
        PushButton 40, 105, 30, 10, "MEMB", MEMB_button
        PushButton 10, 115, 30, 10, "PROG", PROG_button
        PushButton 40, 115, 30, 10, "SHEL", SHEL_button
        PushButton 10, 125, 30, 10, "TYPE", TYPE_button
        PushButton 40, 125, 30, 10, "UNEA", UNEA_button
        PushButton 15, 135, 50, 10, "CASE/CURR", CURR_button
        PushButton 15, 145, 50, 10, "MONY/INQX", MONY_button
      Text 145, 10, 105, 10, "Number of EMER HH members:"
      Text 100, 110, 125, 10, "What is the household's NET income?"
      Text 10, 10, 120, 10, "Case number: " & MAXIS_case_number
      GroupBox 5, 30, 275, 30, "Crisis (Check all that apply. If none, do not check any):"
      Text 5, 70, 220, 10, "Has anyone in the HH been residing in MN for more than 30 days?"
      Text 100, 90, 125, 10, "What is the household's shelter cost?"
      Text 90, 130, 60, 10, "Worker signature:"
      GroupBox 0, 85, 80, 75, "STAT navigation"
    EndDialog

    'Running the initial dialog
    DO
    	DO
            err_msg = ""
    		DO
    			Dialog emergency_screening_dialog
    			cancel_without_confirmation
				MAXIS_dialog_navigation
                IF buttonpressed = EMER_HSR_manual_button then CreateObject("WScript.Shell").Run("https://dept.hennepin.us/hsphd/manuals/hsrm/Pages/Emergency_Assistance_Policy.aspx") 'HSR manual policy page
            LOOP until ButtonPressed = -1
    		If HH_members = "" or IsNumeric(HH_members) = False then err_msg = err_msg & vbNewLine & "* Enter the number of household members."
    		If meets_residency = "Select one..." then err_msg = err_msg & vbNewLine & "* Answer the MN residency question."
    		If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
			If IsNumeric(shelter_costs) = false then err_msg = err_msg & vbNewLine & "* Enter a numeric shelter cost amount."
    		If net_income = "" or IsNumeric(net_income) = False then err_msg = err_msg & vbNewLine & "* Enter the household's net income."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    	LOOP until err_msg = ""
    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

    Call navigate_to_MAXIS_screen("MONY", "INQX")
    EMWriteScreen begin_search_month, 6, 38		'entering footer month/year 13 months prior to current footer month/year'
    EMWriteScreen begin_search_year, 6, 41
    EMWriteScreen CM_mo, 6, 53		'entering current footer month/year
    EMWriteScreen CM_yr, 6, 56
    EMWriteScreen "x", 9, 50		'selecting EA
    EMWriteScreen "x", 11, 50		'selecting EGA
    transmit

    'searching for EA/EG issued on the INQD screen
    DO
    	row = 6
    	DO
    		EMReadScreen emer_issued, 1, row, 16		'searching for EMER programs as they start with E
    		IF emer_issued = "E" then
    			'reading the EMER information for EMER issuance
    			EMReadScreen EMER_type, 2, row, 16
    			EMReadScreen EMER_amt_issued, 7, row, 39
    			EMReadScreen EMER_elig_start_date, 8, row, 62
    			EMReadScreen EMER_elig_end_date, 8, row, 73
    			exit do
    		ELSE
    			row = row + 1
    		END IF
    	Loop until row = 18				'repeats until the end of the page
    	PF8
    	EMReadScreen last_page_check, 21, 24, 2
    	If last_page_check <> "THIS IS THE LAST PAGE" then row = 6		're-establishes row for the new page
    LOOP UNTIL last_page_check = "THIS IS THE LAST PAGE"

    'creating variables and conditions for EMER screening
    New_EMER_year = dateadd("YYYY", 1, EMER_elig_start_date)
    EMER_available_date = dateadd("d", 1, New_EMER_year)	'creating emer available date that is 1 day & 1 year past the EMER_elig_end_date
    EMER_last_used_dates = EMER_elig_start_date & " - " & EMER_elig_end_date	'combining dates into new variable

    If emer_issued <> "E" then	'creating variables for cases that have not had EMER issued in current 13 months
     	EMER_last_used_dates = "n/a"
    	EMER_available_date = "Currently available"
        emer_availble = True 
    Elseif datediff("D", EMER_elig_end_date, date) > 0 then
        emer_availble = True        'If emer was used but the elig end date has passed 
    Else 
        emer_availble = False       'not eligible 
    End if 

	crisis = ""
    'Logic to enter what the "crisis" variable is from the check boxes indicated
    If eviction_check = 1 then crisis = crisis & "eviction, "
    If utility_disconnect_check = 1 then crisis = crisis & "utility disconnect, "
    If homelessness_check = 1 then crisis = crisis & "homelessness, "
    If security_deposit_check = 1 then crisis = crisis & "security deposit, "
    If eviction_check = 0 and utility_disconnect_check = 0 and homelessness_check = 0 and security_deposit_check = 0 then
      crisis = "no crisis given"
    Else
      crisis = trim(crisis)
      crisis = left(crisis, len(crisis) - 1)
    End if

    'determining  200% FPG (using last year's amounts) per HH member---handles up to 20 members. Updated with 2019's FPG. Changes April 1 every year.
	If HH_members = "1"  then monthly_standard = "2082"
    If HH_members = "2"  then monthly_standard = "2818"
    If HH_members = "3"  then monthly_standard = "3555"
    If HH_members = "4"  then monthly_standard = "4292"
    If HH_members = "5"  then monthly_standard = "5028"
    If HH_members = "6"  then monthly_standard = "5765"
    If HH_members = "7"  then monthly_standard = "6502"
    If HH_members = "8"  then monthly_standard = "7238"
    If HH_members = "9"  then monthly_standard = "7975"
    If HH_members = "10" then monthly_standard = "8712"
    If HH_members = "11" then monthly_standard = "9449"
    If HH_members = "12" then monthly_standard = "10186"
    If HH_members = "13" then monthly_standard = "10923"
    If HH_members = "14" then monthly_standard = "11660"
    If HH_members = "15" then monthly_standard = "12397"
    If HH_members = "16" then monthly_standard = "13134"
    If HH_members = "17" then monthly_standard = "13871"
    If HH_members = "18" then monthly_standard = "14608"
    If HH_members = "19" then monthly_standard = "15345"
    If HH_members = "20" then monthly_standard = "16082"

    seventy_percent_income = net_income * .70   'This is to determine if shel costs exceed 70% of the HH's income

    'determining if client is potentially elig for EMER or not'
	If crisis <> "no crisis given" AND meets_residency = "Yes" AND abs(net_income) < abs(monthly_standard) AND net_income <> "0" AND emer_availble = True AND abs(seventy_percent_income) > abs(shelter_costs) then
        screening_determination = "potentially eligible for EGA."
	Else
        screening_determination = "NOT eligible for EGA for the following reasons:" & vbcr
        'if client is not elig, reason(s) for not being elig will be listed in the msgbox
        If crisis = "no crisis given" then screening_determination = screening_determination & vbNewLine & "* No crisis meeting program requirements."
        If abs(seventy_percent_income) < abs(shelter_costs) then screening_determination = screening_determination & vbNewLine & "* The HH's shelter costs are more than 70% of the HH's net income."
	    IF meets_residency = "No" then screening_determination = screening_determination & vbNewLine & "* No one in the household has met 30 day residency requirements."
        If abs(net_income) > abs(monthly_standard) then screening_determination = screening_determination & vbNewLine & "* Net income exceeds program guidelines."
        IF net_income = "0" then screening_determination = screening_determination & vbNewLine & "* Household does not have current/ongoing income."
        If EMER_last_used_dates <> "n/a" then screening_determination = screening_determination & vbNewLine & "* Emergency funds were used within the last year from the eligibility period."
		'If EMER_available_date = > Cdate then screening_determination = screening_determination & vbNewLine & "* Emergency funds were used within the last year from the eligibility period."
    End if

    'Msgbox with screening results. Will give the user the option to cancel the script, case note the results, or use the EMER notes script
    Screening_options = MsgBox ("Based on the information provided, this HH appears to " & screening_determination & vbNewLine & vbNewLine &"The last date emergency funds were used was: " & EMER_last_used_dates & "." & _
    vbNewLine & "Emergency programs are available to the HH effective: " & EMER_available_date & "." & vbNewLine & vbNewLine & "Would you like to start the NOTES - EMERGENCY script?" , vbYesNoCancel, "Screening results dialog")

    IF Screening_options = vbCancel then script_end_procedure("")	'ends the script
    IF Screening_options = vbNO then
    	'The case note
    	Call start_a_blank_CASE_NOTE
    	Call write_variable_in_CASE_NOTE("--//--Emergency Programs Screening--//--")
    	Call write_bullet_and_variable_in_CASE_NOTE("Number of HH members", HH_members)
    	Call  write_bullet_and_variable_in_CASE_NOTE("Crisis/Type of Emergency", crisis)
    	Call write_bullet_and_variable_in_CASE_NOTE("Living situation is", affordbable_housing)
    	Call write_bullet_and_variable_in_CASE_NOTE("Does any member of the HH meet 30 day residency requirements", meets_residency)
    	Call write_bullet_and_variable_in_CASE_NOTE("Shelter cost for HH", shelter_costs)
		Call write_bullet_and_variable_in_CASE_NOTE("Net income for HH", net_income)
    	IF screening_determination = "potentially eligible for emergency programs." then
    		Call write_variable_in_CASE_NOTE("* HH is potentially eligible for EMER programs.")
    	Else
    		Call write_variable_in_CASE_NOTE("* HH does not appear eligible for EMER programs.")
    	END IF
    	Call write_variable_in_CASE_NOTE("---")
    	Call write_bullet_and_variable_in_CASE_NOTE("Last date EMER programs were used", EMER_last_used_dates)
    	Call write_variable_in_CASE_NOTE("* Date EMER programs will be available to HH: " & EMER_available_date)
    	Call write_variable_in_CASE_NOTE("---")
    	Call write_variable_in_CASE_NOTE(worker_signature)
		script_end_procedure("")
	END IF
END IF
'End of EMER screening code----------------------------------------------------------------------------------------------------

'Jumping into STAT
call navigate_to_MAXIS_screen("stat", "hcre")
'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Autofilling
call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", income)
call autofill_editbox_from_MAXIS(HH_member_array, "CARS", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "CASH", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "COEX", monthly_expense)
call autofill_editbox_from_MAXIS(HH_member_array, "DCEX", monthly_expense)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", income)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)
call autofill_editbox_from_MAXIS(HH_member_array, "OTHR", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", income)
call autofill_editbox_from_MAXIS(HH_member_array, "REST", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "SECU", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "SHEL", monthly_expense)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", income)
call autofill_editbox_from_MAXIS(HH_member_array, "HEST", monthly_expense) 'Does this last because people like it tacked on to the end, not before. The rest are alphabetical.

'Showing the case note
DO
	Do
        err_msg = ""
		Do
			Dialog emergency_dialog
			cancel_confirmation
			MAXIS_dialog_navigation
		Loop until ButtonPressed = -1
		If ButtonPressed = -1 then dialog case_note_dialog
	    If trim(income) = "" then err_msg = err_msg & vbcr & "* Enter income information."
        If trim(actions_taken) = "" then err_msg = err_msg & vbcr & "* Enter your actions taken."
        If trim(worker_signature) = "" then err_msg = err_msg & vbcr & "* Sign your case note."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    LOOP until err_msg = ""
    Call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

crisis = ""
'Logic to enter what the "crisis" variable is from the checkboxes indicated
If eviction_check = 1 then crisis = crisis & "eviction, "
If utility_disconnect_check = 1 then crisis = crisis & "utility disconnect, "
If homelessness_check = 1 then crisis = crisis & "homelessness, "
If security_deposit_check = 1 then crisis = crisis & "security deposit, "
If eviction_check = 0 and utility_disconnect_check = 0 and homelessness_check = 0 and security_deposit_check = 0 then
  crisis = "no crisis given."
Else
  crisis = trim(crisis)
  crisis = left(crisis, len(crisis) - 1) & "."
End if

'Writing the case note
call start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("***Emergency app: "& replace(crisis, ".", "") & "***")
call write_bullet_and_variable_in_CASE_NOTE("Interview date", interview_date)
call write_bullet_and_variable_in_CASE_NOTE("HH comp", HH_comp)
call write_bullet_and_variable_in_CASE_NOTE("Crisis", crisis)
call write_bullet_and_variable_in_CASE_NOTE("Cause of crisis", cause_of_crisis)
call write_bullet_and_variable_in_CASE_NOTE("Income, past " & emer_number_of_income_days & " days", income)
call write_bullet_and_variable_in_CASE_NOTE("Income under 200% FPG", income_under_200_FPG)
call write_bullet_and_variable_in_CASE_NOTE(emer_percent_rule_amt & "% rule notes", percent_rule_notes)
call write_bullet_and_variable_in_CASE_NOTE("Monthly expense", monthly_expense)
call write_bullet_and_variable_in_CASE_NOTE("Assets", assets)
call write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)
call write_bullet_and_variable_in_CASE_NOTE("Crisis resolvable?", crisis_resolvable)
call write_bullet_and_variable_in_CASE_NOTE("Discussion of crisis", discussion_of_crisis)
call write_bullet_and_variable_in_CASE_NOTE("Actions taken", actions_taken)
call write_bullet_and_variable_in_CASE_NOTE("Referrals", referrals)
IF Sent_arep_checkbox = checked THEN CALL write_variable_in_case_note("* Sent form(s) to AREP.")
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure_with_error_report("")
