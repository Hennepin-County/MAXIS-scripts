'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - EMERGENCY.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 480          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
    IF on_the_desert_island = TRUE Then
        FuncLib_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\desert-island\MASTER FUNCTIONS LIBRARY.vbs"
        Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
        Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
        text_from_the_other_script = fso_command.ReadAll
        fso_command.Close
        Execute text_from_the_other_script
    ELSEIF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
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
call changelog_update("04/06/2023", "Updated 200% FPG for 2023.", "Ilse Ferris, Hennepin County")
call changelog_update("11/15/2022", "The display of the EGA Screening result has been updated to repeat the information provided and have buttons to indicate what next action the script should take.", "Casey Love, Hennepin County")
call changelog_update("09/30/2022", "BUG Fix: EGA screening will now indicate that a case screens as potentially eligible for EGA if the net inocme is equal to the standard or if 70% of the net income is exactly equal to the shelter costs.##~####~##Previously if they were equal the script would screen as not eligible but did not give details around why the screening appears ineligble (there was no handling for situations where the amounts were exactly equal.)##~####~##EGA screening was not correctly identifying if EGA was available/had already been used in the past 12 months.##~##", "Casey Love, Hennepin County")
call changelog_update("04/02/2022", "Updated 200% FPG for 2022.", "Ilse Ferris, Hennepin County")
call changelog_update("04/01/2021", "Updated 200% FPG for 2021.", "Ilse Ferris, Hennepin County")
call changelog_update("11/12/2020", "Updated HSR Manual link for EGA policy/procedure due to SharePoint Online Migration.", "Ilse Ferris, Hennepin County")
call changelog_update("07/30/2020", "BUG Fix: The script was getting stuck and would not continue to the note if a required field was not completed.##~##Updated so that the dialog with the missing field would appear.##~##", "Casey Love, Hennepin County")
call changelog_update("04/01/2020", "Updated 200% FPG for 2020.", "Ilse Ferris, Hennepin County")
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

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number & footer month/year
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)
CALL MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 141, 115, "Case number dialog"
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
'Showing the case number dialog
DO
	Do
	    err_msg = ""
		Dialog Dialog1
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
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 286, 170, "Emergency Screening dialog"
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
    DO
    	DO
		    err_msg = ""
    		DO
    			Dialog Dialog1
    			cancel_without_confirmation
				MAXIS_dialog_navigation
                IF buttonpressed = EMER_HSR_manual_button then CreateObject("WScript.Shell").Run("https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/EGA_Policy.aspx") 'HSR manual policy page
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
    Elseif datediff("D", EMER_available_date, date) > 0 then
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

    'determining  200% FPG per HH member---handles up to 20 members. Changes April 1 every year. CM0016.18.01 - 200 Percent of FPG
	If HH_members = "1"  then monthly_standard = "2430"
    If HH_members = "2"  then monthly_standard = "3287"
    If HH_members = "3"  then monthly_standard = "4143"
    If HH_members = "4"  then monthly_standard = "5000"
    If HH_members = "5"  then monthly_standard = "5857"
    If HH_members = "6"  then monthly_standard = "6713"
    If HH_members = "7"  then monthly_standard = "7570"
    If HH_members = "8"  then monthly_standard = "8427"
    If HH_members = "9"  then monthly_standard = "9283"
    If HH_members = "10" then monthly_standard = "10140"
    If HH_members = "11" then monthly_standard = "10997"
    If HH_members = "12" then monthly_standard = "11854"
    If HH_members = "13" then monthly_standard = "12711"
    If HH_members = "14" then monthly_standard = "13568"
    If HH_members = "15" then monthly_standard = "14425"
    If HH_members = "16" then monthly_standard = "15282"
    If HH_members = "17" then monthly_standard = "16139"
    If HH_members = "18" then monthly_standard = "16996"
    If HH_members = "19" then monthly_standard = "17853"
    If HH_members = "20" then monthly_standard = "18710"

    seventy_percent_income = net_income * .70   'This is to determine if shel costs exceed 70% of the HH's income

    'determining if client is potentially elig for EMER or not'
	If crisis <> "no crisis given" AND meets_residency = "Yes" AND abs(net_income) =< abs(monthly_standard) AND net_income <> "0" AND emer_availble = True AND abs(seventy_percent_income) >= abs(shelter_costs) then
        ega_results_dlg_len = 220
	Else
        screening_determination = "NOT eligible for EGA for the following reasons:" & vbcr
        ega_results_dlg_len = 220
        'if client is not elig, reason(s) for not being elig will be listed in the msgbox
        If crisis = "no crisis given" then ega_results_dlg_len = ega_results_dlg_len + 10
        If abs(seventy_percent_income) < abs(shelter_costs) then ega_results_dlg_len = ega_results_dlg_len + 10'"* The HH's shelter costs are more than 70% of the HH's net income."
	    IF meets_residency = "No" then ega_results_dlg_len = ega_results_dlg_len + 10'"* No one in the household has met 30 day residency requirements."
        If abs(net_income) > abs(monthly_standard) then ega_results_dlg_len = ega_results_dlg_len + 10'"* Net income exceeds program guidelines."
        IF net_income = "0" then ega_results_dlg_len = ega_results_dlg_len + 10'"* Household does not have current/ongoing income."
        If EMER_last_used_dates <> "n/a" then ega_results_dlg_len = ega_results_dlg_len + 10'"* Emergency funds were used within the last year from the eligibility period."
		'If EMER_available_date = > Cdate then screening_determination = screening_determination & vbNewLine & "* Emergency funds were used within the last year from the eligibility period."
    End if

    ega_screening_note_made = False
    Do
        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 301, ega_results_dlg_len, "EGA Screening Results"
          ButtonGroup ButtonPressed
            If ega_screening_note_made = False Then
                PushButton 185, ega_results_dlg_len-60, 110, 15, "Enter Screening in CASE/NOTE", enter_screening_note_btn
                PushButton 80, ega_results_dlg_len-40, 215, 15, "Do NOT CASE/NOTE Screening - Continue to Emergency Script", continue_to_emer_script_btn
                PushButton 140, ega_results_dlg_len-20, 155, 15, "Do NOT CASE/NOTE Screening - End Script", end_script_btn
            End If
            If ega_screening_note_made = True Then
                ' PushButton 185, ega_results_dlg_len-60, 110, 15, "Enter Screening in CASE/NOTE", enter_screening_note_btn
                PushButton 80, ega_results_dlg_len-40, 215, 15, "CASE/NOTE Created - Continue to Emergency Script", continue_to_emer_script_btn
                PushButton 140, ega_results_dlg_len-20, 155, 15, "CASE/NOTE Created - End Script", end_script_btn
            End If
          GroupBox 10, 10, 285, 110, "EGA Screening Details"
          Text 20, 25, 270, 10, "Crisis: " & crisis
          Text 30, 40, 125, 10, "Shelter Expense: $ " & shelter_costs
          Text 25, 50, 125, 10, "Household Income: $ " & net_income & "  (net)"
          Text 140, 50, 115, 10, " 70% of HH Income: $ " & seventy_percent_income
          Text 25, 60, 75, 10, "  # of HH Members: " & HH_members
          Text 130, 60, 115, 10, "EGA Monthly Standard: $ " & monthly_standard
          Text 30, 70, 120, 10, " State Residency: " & meets_residency
          Text 35, 90, 130, 10, "  EMER Last Used: " & EMER_last_used_dates
          Text 25, 100, 130, 10, "EMER Available Date: " & EMER_available_date
          Text 10, 130, 75, 10, "Case " & MAXIS_case_number & " is:"
          If crisis <> "no crisis given" AND meets_residency = "Yes" AND abs(net_income) =< abs(monthly_standard) AND net_income <> "0" AND emer_availble = True AND abs(seventy_percent_income) >= abs(shelter_costs) then

            Text 20, 140, 125, 10, "POTENTIALLY ELIGIBLE FOR EGA"
          Else
            y_pos = 140
            Text 20, y_pos, 250, 10, "NOT eligible for EGA for the following reasons:"
            y_pos = y_pos + 10
            If crisis = "no crisis given" then
                Text 25, y_pos, 250, 10, "* No crisis meeting program requirements."
                y_pos = y_pos + 10
            End If
            If abs(seventy_percent_income) < abs(shelter_costs) then
                Text 25, y_pos, 250, 10, "* The HH's shelter costs are more than 70% of the HH's net income."
                y_pos = y_pos + 10
            End If
            IF meets_residency = "No" then
                Text 25, y_pos, 250, 10, "* No one in the household has met 30 day residency requirements."
                y_pos = y_pos + 10
            End If
            If abs(net_income) > abs(monthly_standard) then
                Text 25, y_pos, 250, 10, "* Net income exceeds program guidelines."
                y_pos = y_pos + 10
            End If
            IF net_income = "0" then
                Text 25, y_pos, 250, 10, "* Household does not have current/ongoing income."
                y_pos = y_pos + 10
            End If
            If EMER_last_used_dates <> "n/a" then
                Text 25, y_pos, 250, 10, "* Emergency funds were used within the last year from the eligibility period."
                y_pos = y_pos + 10
            End If

          End If
        EndDialog

        dialog Dialog1
        cancel_without_confirmation

        If ButtonPressed = end_script_btn Then Call script_end_procedure("EGA Screening completed, script ended per your request.")

        If ButtonPressed = enter_screening_note_btn Then
            'The case note
            ega_screening_note_made = True
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
        End If
    Loop until ButtonPressed = continue_to_emer_script_btn
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

'-------------------------------------------------------------------------------------------------DIALOG
DO
	Do
        Dialog1 = "" 'Blanking out previous dialog detail
        'This dialog contains a customized "percent rule" variable, as well as a customized "income days" variable. As such, it can't directly be edited in the dialog editor.
        BeginDialog Dialog1, 0, 0, 326, 395, "Emergency Dialog"
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

	    err_msg = ""
		Do
			Dialog Dialog1
			cancel_confirmation
			MAXIS_dialog_navigation
		Loop until ButtonPressed = -1
	    If trim(income) = "" then err_msg = err_msg & vbcr & "* Enter income information."
        If trim(actions_taken) = "" then err_msg = err_msg & vbcr & "* Enter your actions taken."
        If trim(worker_signature) = "" then err_msg = err_msg & vbcr & "* Sign your case note."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
        If err_msg = "" then
            '-------------------------------------------------------------------------------------------------DIALOG
            Dialog1 = "" 'Blanking out previous dialog detail
            BeginDialog Dialog1, 0, 0, 136, 51, "Case note dialog"
              ButtonGroup ButtonPressed
                PushButton 15, 20, 105, 10, "Yes, take me to case note.", yes_case_note_button
                PushButton 5, 35, 125, 10, "No, take me back to the script dialog.", no_case_note_button
              Text 10, 5, 125, 10, "Are you sure you want to case note?"
            EndDialog
            dialog Dialog1

            If ButtonPressed = no_case_note_button Then err_msg = "LOOP"
        END IF
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
