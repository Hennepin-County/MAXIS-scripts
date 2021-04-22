'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - ABAWD Waived Approval.vbs"
start_time = timer
STATS_counter = 1			 'sets the stats counter at one
STATS_manualtime = 270			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
'END OF stats block==============================================================================================
run_locally = TRUE
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
call changelog_update("04/17/2020", "Added verbiage to the CASE:NOTE header for Current Month plus 1 to indicate the benefit information in the note is approved for ongoing.", "Casey Love, Hennepin County")
call changelog_update("04/08/2020", "Updated functionality to add a separate note for each month of approval completed. Enter the initial month of approval and the script will loop through the entire information gathering, WCOMs, and CASE:NOTE for each month from the initial to current month plus one.##~##", "Casey Love, Hennepin County")
call changelog_update("04/07/2020", "BUG FIX - sometimes the script couldn't read the dates from MAXIS correctly. Added handling to adjust date formats to match exactly.##~##", "Casey Love, Hennepin County")
call changelog_update("03/03/2020", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

EMConnect ""

'Checks to make sure we're in MAXIS
call check_for_MAXIS(True)

Call MAXIS_case_number_finder(MAXIS_case_number)
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 126, 100, "Case Number Dialog"
  EditBox 60, 35, 60, 15, MAXIS_case_number
  EditBox 85, 55, 15, 15, MAXIS_footer_month
  EditBox 105, 55, 15, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 15, 80, 50, 15
    CancelButton 70, 80, 50, 15
  Text 10, 40, 50, 10, "Case Number:"
  Text 10, 60, 65, 10, "Month of Approval"
  Text 10, 5, 100, 25, "SNAP Approved with ABAWD coded in a waivered area."
EndDialog

Do
    err_msg = ""

    Dialog Dialog1

    cancel_without_confirmation

    Call validate_MAXIS_case_number(err_msg, "*")
    If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg

Loop until err_msg = ""

Do
    Call Back_to_SELF
    Call navigate_to_MAXIS_screen("ELIG", "SUMM")
    EMWriteScreen MAXIS_footer_month, 19, 56
    EMWriteScreen MAXIS_footer_year, 19, 59
    transmit

    EMReadScreen versions_exist, 1, 17, 40
    EMReadScreen version_date, 8, 17, 48
    today_month = DatePart("m", date)
    today_month = right("0" & today_month, 2)
    today_day = DatePart("d", date)
    today_day = right("0" & today_day, 2)
    today_year = DatePart("yyyy", date)
    today_year = right(today_year, 2)
    todays_date = today_month & "/" & today_day & "/" & today_year

    If versions_exist = " " Then
        end_msg = "It does not appear there are any Approved versions of SNAP for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". This scrpit requires a version of SNAP approved today to run accurately. "
        call script_end_procedure_with_error_report(end_msg)
    End if
    If version_date = todays_date Then
        version_to_note = "0" & versions_exist
    Else
        end_msg = "The most recent version of SNAP Eligibility was created on " & version_date & " and in order to correctly note the approval, it should be noted on the day of the approval. Review ELIG/SNAP and complete a new approval and rerun the script."
        call script_end_procedure_with_error_report(end_msg)
    End If

    Call navigate_to_MAXIS_screen("ELIG", "FS  ")
    EMWriteScreen version_to_note, 19, 78
    EMReadScreen version_status, 10, 3, 3
    version_status = trim(version_status)
    If version_status <> "APPROVED" Then
        EMWriteScreen "99", 19, 78
        transmit
        list_row = 7
        Do
            EMReadScreen version_number, 1, list_row, 23
            EMReadScreen process_date, 8, list_row, 26
            EMReadScreen approval_status, 10, list_row, 50
            approval_status = trim(approval_status)
            If approval_status = "APPROVED" Then
                If process_date = todays_date Then
                    version_to_note = "0" & version_number
                    exit do
                End If
            End If
            list_row = list_row + 1
        Loop until version_number = " "
        transmit

        If version_number = " " Then
            end_msg = "The most recent version of SNAP Eligibility that was created today has not yet been approved. Review SNAP ELIG and either approve the current version or delete incorrect elig versions."
            call script_end_procedure_with_error_report(end_msg)
        End If
    End If

    start_month = MAXIS_footer_month
    start_year = MAXIS_footer_year

    Call Back_to_SELF
    Call navigate_to_MAXIS_screen("STAT", "EATS")
    EMReadScreen eats_all_together, 1, 4, 72
    all_the_eats_ref = ""
    ' MsgBox eats_all_together
    If eats_all_together = "N" Then
        eats_col = 39
        Do
            EMReadScreen person_ref, 2, 13, eats_col
            If person_ref <> "__" Then all_the_eats_ref = all_the_eats_ref & "~" & person_ref
            eats_col = eats_col + 4
        Loop until person_ref = "__"

    Else
        stat_row = 5
        Do
            EMReadScreen person_ref, 2, stat_row, 3
            ' MsgBox person_ref
            If person_ref <> "  " Then all_the_eats_ref = all_the_eats_ref & "~" & person_ref
            stat_row = stat_row + 1
        Loop until person_ref = "  "
    End If
    ' MsgBox all_the_eats_ref
    If left(all_the_eats_ref, 1) = "~" Then all_the_eats_ref = right(all_the_eats_ref, len(all_the_eats_ref) - 1)
    If InStr(all_the_eats_ref, "~") <> 0 Then
        member_array = split(all_the_eats_ref, "~")
    Else
        member_array = array(all_the_eats_ref)
    End If

    Call navigate_to_MAXIS_screen("STAT", "ADDR")
    EMReadScreen homeless_indicator, 1, 10, 43
    If homeless_indicator = "Y" Then homeless_wcom_checkbox = checked

    all_members_wreg_info = ""
    waiver_ABAWD_on_case = FALSE
    Call navigate_to_MAXIS_screen("STAT", "WREG")
    For each member in member_array
        ' MsgBox member
        Call write_value_and_transmit(member, 20, 76)

        EMReadScreen wreg_exists, 14, 24, 13
        If wreg_exists <> "DOES NOT EXIST" Then
            EMReadScreen member_fs_pwe, 1, 6, 68
            EMReadScreen member_fset_wreg_status, 2, 8, 50
            EMReadScreen member_abawd_status, 2, 13, 50
            EMReadScreen member_bm_indicator, 1, 14, 50

            EmWriteScreen "x", 13, 57
            transmit
            bene_mo_col = (15 + (4*cint(MAXIS_footer_month)))
            bene_yr_row = 10
            abawd_counted_months = 0
            abawd_info_list = ""
            second_abawd_period = 0
            second_set_info_list = ""
            exemption_months = 0
            exemption_months_list = ""
            banked_months_count = 0
            banked_months_list = ""
            meets_work_req_count = 0
            meets_work_req_list = ""
            If member_abawd_status = "09" Then waiver_ABAWD_on_case = TRUE

            month_count = 0
            DO
                'establishing variables for specific ABAWD counted month dates
                If bene_mo_col = "19" then counted_date_month = "01"
                If bene_mo_col = "23" then counted_date_month = "02"
                If bene_mo_col = "27" then counted_date_month = "03"
                If bene_mo_col = "31" then counted_date_month = "04"
                If bene_mo_col = "35" then counted_date_month = "05"
                If bene_mo_col = "39" then counted_date_month = "06"
                If bene_mo_col = "43" then counted_date_month = "07"
                If bene_mo_col = "47" then counted_date_month = "08"
                If bene_mo_col = "51" then counted_date_month = "09"
                If bene_mo_col = "55" then counted_date_month = "10"
                If bene_mo_col = "59" then counted_date_month = "11"
                If bene_mo_col = "63" then counted_date_month = "12"
                'reading to see if a month is counted month or not
                EMReadScreen is_counted_month, 1, bene_yr_row, bene_mo_col

                If is_counted_month = "P" Then

                End If

                'counting and checking for counted ABAWD months
                IF is_counted_month = "X" or is_counted_month = "M" THEN
                    EMReadScreen counted_date_year, 2, bene_yr_row, 15			'reading counted year date
                    abawd_counted_months_string = counted_date_month & "/" & counted_date_year
                    abawd_info_list = abawd_info_list & ", " & abawd_counted_months_string			'adding variable to list to add to array
                    abawd_counted_months = abawd_counted_months + 1				'adding counted months
                END IF

                'counting and checking for second set of ABAWD months
                IF is_counted_month = "Y" or is_counted_month = "N" THEN
                    EMReadScreen counted_date_year, 2, bene_yr_row, 15			'reading counted year date
                    second_abawd_period = second_abawd_period + 1				'adding counted months
                    second_counted_months_string = counted_date_month & "/" & counted_date_year			'creating new variable for array
                    second_set_info_list = second_set_info_list & ", " & second_counted_months_string	'adding variable to list to add to array
                END IF

                If is_counted_month = "E" or is_counted_month = "F" Then
                    EMReadScreen counted_date_year, 2, bene_yr_row, 15			'reading counted year date
                    exemption_months = exemption_months + 1				'adding counted months
                    exemption_months_string = counted_date_month & "/" & counted_date_year			'creating new variable for array
                    exemption_months_list = exemption_months_list & ", " & exemption_months_string	'adding variable to list to add to array
                End If

                If is_counted_month = "B" or is_counted_month = "C" Then
                    EMReadScreen counted_date_year, 2, bene_yr_row, 15			'reading counted year date
                    banked_months_count = banked_months_count + 1				'adding counted months
                    banked_month_string = counted_date_month & "/" & counted_date_year			'creating new variable for array
                    banked_months_list = banked_months_list & ", " & banked_month_string	'adding variable to list to add to array
                End If

                If is_counted_month = "W" or is_counted_month = "V" Then
                    EMReadScreen counted_date_year, 2, bene_yr_row, 15			'reading counted year date
                    meets_work_req_count = meets_work_req_count + 1				'adding counted months
                    meets_work_req_string = counted_date_month & "/" & counted_date_year			'creating new variable for array
                    meets_work_req_list = meets_work_req_list & ", " & meets_work_req_string	'adding variable to list to add to array
                End If

                bene_mo_col = bene_mo_col - 4
                IF bene_mo_col = 15 THEN
                    bene_yr_row = bene_yr_row - 1
                    bene_mo_col = 63
                END IF
                month_count = month_count + 1
            LOOP until month_count = 36
            'declaring & splitting the abawd months array
            If left(abawd_info_list, 1) = "," then abawd_info_list = right(abawd_info_list, len(abawd_info_list) - 1)
            'declaring & splitting the second set of abawd months array
            If left(second_set_info_list, 1) = "," then second_set_info_list = right(second_set_info_list, len(second_set_info_list) - 1)
            If left(exemption_months_list, 1) = "," then exemption_months_list = right(exemption_months_list, len(exemption_months_list) - 1)
            If left(banked_months_list, 1) = "," then banked_months_list = right(banked_months_list, len(banked_months_list) - 1)
            If left(meets_work_req_list, 1) = "," then meets_work_req_list = right(meets_work_req_list, len(meets_work_req_list) - 1)
            PF3

            member_fs_pwe = replace(member_fs_pwe, "_", "")
            If member_fset_wreg_status = "03" Then member_fset_wreg_status = "03 - Unfit for Employment"
            If member_fset_wreg_status = "04" Then member_fset_wreg_status = "04 - Resp for Care of Incapacitated Person"
            If member_fset_wreg_status = "05" Then member_fset_wreg_status = "05 - Age 60 or Older"
            If member_fset_wreg_status = "06" Then member_fset_wreg_status = "06 - Under Age 16"
            If member_fset_wreg_status = "07" Then member_fset_wreg_status = "07 - Age 16-17, Living w/ Caregiver"
            If member_fset_wreg_status = "08" Then member_fset_wreg_status = "08 - Resp for Care of Child under 6"
            If member_fset_wreg_status = "09" Then member_fset_wreg_status = "09 - Empl 30 hrs/wk or Earnings of 30 hrs/wk"
            If member_fset_wreg_status = "10" Then member_fset_wreg_status = "10 - Matching Grant Participant"
            If member_fset_wreg_status = "11" Then member_fset_wreg_status = "11 - Receiving or Applied for UI"
            If member_fset_wreg_status = "12" Then member_fset_wreg_status = "12 - Enrolled in School, Training, or Higher Ed"
            If member_fset_wreg_status = "13" Then member_fset_wreg_status = "13 - Participating in CD Program"
            If member_fset_wreg_status = "14" Then member_fset_wreg_status = "14 - Receiving MFIP"
            If member_fset_wreg_status = "20" Then member_fset_wreg_status = "20 - Pending/Receiving DWP"
            If member_fset_wreg_status = "15" Then member_fset_wreg_status = "15 - Age 16-17, NOT Living w/ Caregiver"
            If member_fset_wreg_status = "16" Then member_fset_wreg_status = "16 - 50-59 Years Old"
            If member_fset_wreg_status = "17" Then member_fset_wreg_status = "17 - Receiving RCA or GA"
            If member_fset_wreg_status = "21" Then member_fset_wreg_status = "21 - Resp for Care of Child under 18"
            If member_fset_wreg_status = "30" Then member_fset_wreg_status = "30 - Mandatory FSET Participant"
            If member_fset_wreg_status = "02" Then member_fset_wreg_status = "02 - Fail to Cooperate with FSET"
            If member_fset_wreg_status = "33" Then member_fset_wreg_status = "33 - Non-Coop being Referred"

            If member_abawd_status = "01" Then member_abawd_status = "01 - Work Reg Exempt"
            If member_abawd_status = "02" Then member_abawd_status = "02 - Under Age 18"
            If member_abawd_status = "03" Then member_abawd_status = "03 - Age 50 or Over"
            If member_abawd_status = "04" Then member_abawd_status = "04 - Caregiver of Minor Child"
            If member_abawd_status = "05" Then member_abawd_status = "05 - Pregnant"
            If member_abawd_status = "06" Then member_abawd_status = "06 - Employed Avg of 20 hrs/wk"
            If member_abawd_status = "07" Then member_abawd_status = "07 - Work Experience Participant"
            If member_abawd_status = "08" Then member_abawd_status = "08 - Other E&T Services"
            If member_abawd_status = "09" Then member_abawd_status = "09 - Resides in a Waivered Area"
            If member_abawd_status = "10" Then member_abawd_status = "10 - ABAWD Counted Month"
            If member_abawd_status = "11" Then member_abawd_status = "11 - 2nd-3rd Month Period of Elig"
            If member_abawd_status = "12" Then member_abawd_status = "12 - RCA or GA Recipient"
            If member_abawd_status = "13" Then member_abawd_status = "13 - ABAWD Banked Months"
            member_bm_indicator = replace(member_bm_indicator, "_", "")

            all_members_wreg_info = all_members_wreg_info & " - Memb " & member & ": FSET Status: " & member_fset_wreg_status & "; "
            all_members_wreg_info = all_members_wreg_info & "           ABAWD Status: " & member_abawd_status & "; "
            If abawd_info_list = "" Then all_members_wreg_info = all_members_wreg_info & "           Counted ABAWD Months Used: " & abawd_counted_months & "; "
            If abawd_info_list <> "" Then all_members_wreg_info = all_members_wreg_info & "           Counted ABAWD Months Used: " & abawd_counted_months & " months: " & abawd_info_list & "; "

            If second_set_info_list = "" Then all_members_wreg_info = all_members_wreg_info & "           Second Set Months Used: " & second_abawd_period & "; "
            If second_set_info_list <> "" Then all_members_wreg_info = all_members_wreg_info & "           Second Set Months Used: " & second_abawd_period & " months: " & second_set_info_list & "; "

            ' If banked_months_list = "" Then all_members_wreg_info = all_members_wreg_info & "           Banked Months Used: " & banked_months_count & "; "
            ' If banked_months_list <> "" Then all_members_wreg_info = all_members_wreg_info & "           Banked Months Used: " & banked_months_count & " months: " & banked_months_list & "; "

        End If
    Next
    If waiver_ABAWD_on_case = FALSE Then script_end_procedure_with_error_report("There does not appear to be a member whose ABAWD is waived on this case. This script is made to accomodate processing for the ABAWD emergency order waivered status. Please use 'Approved Programs' for any other kind of case.")

    Call navigate_to_MAXIS_screen("ELIG", "FS  ")
    EMWriteScreen version_to_note, 19, 78
    transmit
    'Reading the HH Members on SNAP
    FSPR_row = 7
    membs_elig_for_SNAP = ""
    membs_inelig = ""
    Do
        EMReadScreen memb_ref_nbr, 2, FSPR_row, 10
        EMReadScreen memb_name, 18, FSPR_row, 13
        EMReadScreen state_food, 1, FSPR_row, 50
        EMReadScreen memb_elig, 10, FSPR_row, 57
        memb_name = trim(memb_name)
        memb_elig = trim(memb_elig)

        If memb_elig = "ELIGIBLE" Then
            If state_food = "N" Then membs_elig_for_SNAP = membs_elig_for_SNAP & " - Memb " &  memb_ref_nbr & " - " & memb_name & "; "
            If state_food = "Y" Then membs_elig_for_SNAP = membs_elig_for_SNAP & " - Memb " &  memb_ref_nbr & " - " & memb_name & " - STATE FOOD BENEFIT; "
        ElseIf memb_elig = "INELIGIBLE" Then
            EMReadScreen memb_request, 1, FSPR_row, 32
            EMreadScreen member_counted, 7, FSPR_row, 39
            If member_counted = "N" then
                membs_inelig = membs_inelig & " - Memb " & memb_ref_nbr & " - " & memb_name & " - did not request SNAP.; "
            ElseIf member_counted = "COUNTED" Then
                membs_inelig = membs_inelig & " - Memb " & memb_ref_nbr & " - " & memb_name & " - Ineligible but income is counted.; "
            Else
                membs_inelig = membs_inelig & " - Memb " & memb_ref_nbr & " - " & memb_name & "; "
            End If
        End If
        FSPR_row = FSPR_row + 1
    Loop until memb_ref_nbr = "  "

    EMWriteScreen "FSCR", 19, 70
    transmit

    EMWriteScreen "FSB1", 19, 70
    transmit

    EMReadScreen all_wages, 9, 5, 32
    EMReadScreen all_self_emp, 9, 6, 32
    EMReadScreen total_earned_income, 9, 8, 32

    EMReadScreen all_pa_grants, 9, 10, 32
    EMReadScreen all_RSDI, 9, 11, 32
    EMReadScreen all_SSI, 9, 12, 32
    EMReadScreen all_VA, 9, 13, 32
    EMReadScreen all_UC_WC, 9, 14, 32
    EMReadScreen all_cses, 9, 15, 32
    EMReadScreen all_other_unea, 9, 16, 32
    EMReadScreen total_UNEA, 9, 18, 32

    EMReadScreen total_gross_income, 9, 7, 72
    EMReadScreen max_gross_income, 9, 8, 72

    EMReadScreen standard_deduction, 8, 10, 73
    EMReadScreen earned_income_deduction, 8, 11, 73
    EMReadScreen medical_deduction, 8, 12, 73
    EMReadScreen dcex_deduction, 8, 13, 73
    EMReadScreen cses_deduction, 8, 14, 73
    EMReadScreen total_deduction, 8, 16, 73

    EMReadScreen total_net_income, 9, 18, 72

    EMWriteScreen "FSB2", 19, 70
    transmit

    EMReadScreen shel_rent, 9, 5, 28
    EMReadScreen shel_tax, 9, 6, 28
    EMReadScreen shel_ins, 9, 7, 28
    EMReadScreen hest_elec, 9, 8, 28
    EMReadScreen hest_heat, 9, 9, 28
    EMReadScreen hest_water, 9, 10, 28
    EMReadScreen hest_phone, 9, 11, 28
    EMReadScreen shel_other, 9, 12, 28

    EMReadScreen total_shel_cost, 9, 14, 28
    EMReadScreen half_net_income, 9, 15, 28
    EMReadScreen adj_shel_cost, 9, 17, 28

    EMReadScreen max_allow_shel, 9, 5, 72
    EMReadScreen shel_expense, 9, 6, 72
    EMReadScreen net_adj_income, 9, 7, 72
    EMReadScreen max_net_adj_income, 9, 8, 72

    If total_shel_cost = "         " Then total_shel_cost = "     0.00"
    If adj_shel_cost = "         " Then adj_shel_cost = "     0.00"
    If shel_expense = "         " Then shel_expense = "     0.00"

    EMReadScreen monthly_fs_allotment, 9, 10, 72
    EMReadScreen recoup_amt, 9, 14, 72
    EMReadScreen benefit_amount, 9, 16, 72
    EMReadScreen state_food_amt, 9, 17, 72
    EMReadScreen fed_food_amt, 9, 18, 72

    EMWriteScreen "FSSM", 20, 70
    transmit

    If benefit_amount = "     0.00" Then EMReadScreen benefit_amount, 9, 10, 72
    EMReadScreen current_prog_status, 12, 6, 31
    EMReadScreen elig_result, 12, 7, 31
    EMReadScreen reporting_status, 12, 8, 31
    EMReadScreen info_source, 4, 9, 31
    EMReadScreen benefit_impact, 12, 10, 31
    EMReadScreen elig_revw_date, 8, 11, 31
    EMReadScreen budget_cycle, 12, 12, 31
    EMReadScreen numb_in_assistance_unit, 2, 13, 31

    abawd_waived_wcom_checkbox = checked

    If right(all_members_wreg_info, 2) = "; " Then all_members_wreg_info = left(all_members_wreg_info, len(all_members_wreg_info)- 2)
    If right(membs_elig_for_SNAP, 2) = "; " Then membs_elig_for_SNAP = left(membs_elig_for_SNAP, len(membs_elig_for_SNAP)- 2)
    If right(membs_inelig, 2) = "; " Then membs_inelig = left(membs_inelig, len(membs_inelig)- 2)
    all_membs_wreg_array = Split(all_members_wreg_info, "; ")
    elig_membs_array = Split(membs_elig_for_SNAP, "; ")
    inelig_membs_array = Split(membs_inelig, "; ")

    wreg_groupbox_hght = 25 + UBound(all_membs_wreg_array) * 10
    elig_memb_hght_one = 35 + UBound(elig_membs_array) * 10
    elig_memb_hght_two = 35 + UBound(inelig_membs_array) * 10
    If elig_memb_hght_one > elig_memb_hght_two Then elig_memb_hght = elig_memb_hght_one
    If elig_memb_hght_two > elig_memb_hght_one Then elig_memb_hght = elig_memb_hght_two
    dlg_hgt = 165 + wreg_groupbox_hght + elig_memb_hght
    Do
        Do
            err_msg = ""

            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 496, dlg_hgt, "CASE NOTE Details"
              GroupBox 5, 5, 485, 75, "Food Support Eligibility Information"
              Text 15, 15, 90, 25, "Benefit Amount: $" & trim(benefit_amount) & "    -Fed Amount: $" & trim(fed_food_amt) & "        -State Amount: $" & trim(state_food_amt)
              Text 135, 10, 105, 65, "            ===Income===              Total Earned Income: $" & trim(total_earned_income) & " Total Unearned Income: $" & trim(total_UNEA) & " ----------------------------------------------------- Total Gross Income: $" & trim(total_gross_income) & "    Max Gross Income: $" & trim(max_gross_income) & "    Total Deductions: $" & trim(total_deduction) & "        Net Income: $" & trim(total_net_income)
              Text 260, 10, 100, 65, "        ===Calculations===            Shelter Expense: $" & shel_expense & "    Net Adj Income: $" & net_adj_income & " ----------------------------------------------------- Monthly FS Allotment: $" & trim(monthly_fs_allotment) & "    Recoup Amt: $" & trim(recoup_amt) & "            Benefit Amt: $" & trim(benefit_amount)
              Text 380, 10, 90, 65, "        ===Summary===         SNAP is " & trim(current_prog_status) & "            Case is " & trim(elig_result) & "            Reporting: " & trim(reporting_status) & "        Number in Unit: " & trim(numb_in_assistance_unit)
              GroupBox 5, 85, 485, wreg_groupbox_hght, "WREG Information"

              y_pos = 95
              For each member in all_membs_wreg_array
                  Text 10, y_pos, 400, 10, member
                  y_pos = y_pos + 10
              Next
              y_pos = y_pos + 10
              ' Text 10, 95, 115, 10, " - Memb 01 NAME - WREG"
              ' Text 10, 105, 110, 10, " - Memb 02 NAME - WREG"
              GroupBox 5, y_pos, 485, elig_memb_hght, "Members "
              y_pos = y_pos + 10
              elig_btn_y_pos = y_pos
              Text 10, y_pos, 405, 10, "===ELIGIBILE===                                                                                     ===INELIGIBLE==="
              y_pos = y_pos + 10
              y_pos_over = y_pos
              For each member in elig_membs_array
                  Text 10, y_pos, 200, 10, member
                  y_pos = y_pos + 10
              Next
              For each member in inelig_membs_array
                  Text 235, y_pos_over, 200, 10, member
                  y_pos_over = y_pos_over + 10
              Next
              If y_pos_over > y_pos Then y_pos = y_pos_over
              y_pos = y_pos + 10
              ' Text 10, 145, 115, 10, " - Memb 01 NAME - ELIG"
              ' Text 10, 155, 115, 10, " - Memb 02 NAME - ELIG"
              ' Text 235, 145, 115, 10, " - Memb 20 NAME - INELIG"
              Text 10, y_pos + 5, 25, 10, "NOTES"
              EditBox 40, y_pos, 450, 15, other_notes
              y_pos = y_pos + 25
              CheckBox 270, y_pos, 210, 10, "Check here to add WCOM about current ABAWD Waiver", abawd_waived_wcom_checkbox
              y_pos = y_pos + 10
              CheckBox 270, y_pos, 220, 10, "Check here to add WCOM about a possible homeless exemption.", homeless_wcom_checkbox
              y_pos = y_pos + 10
              Text 10, y_pos + 5, 65, 10, "Worker Signature:"
              EditBox 75, y_pos, 155, 15, worker_signature
              ButtonGroup ButtonPressed
                PushButton 450, 95, 35, 10, "NOTE", edit_wreg_detail_btn
                PushButton 450, elig_btn_y_pos, 35, 10, "NOTE", edit_elig_membs_btn
                OkButton 385, y_pos + 5, 50, 15
                CancelButton 440, y_pos + 5, 50, 15
            EndDialog

            Dialog Dialog1

            cancel_confirmation

            ' Dim new_wreg_array ()
            ' ReDim new_wreg_array(0)
            ' array_counter = 0
            ' new_wreg_notes = ""
            ' the_top = UBound(all_membs_wreg_array)
            If ButtonPressed = edit_wreg_detail_btn Then
                wreg_dlg_hgt = 90 + UBound(all_membs_wreg_array) * 20
                Dialog1 = ""
                BeginDialog Dialog1, 0, 0, 371, wreg_dlg_hgt, "WREG Detail"
                  Text 10, 10, 260, 10, "Add detail to WREG information as needed."
                  y_pos = 30
                  For the_thing = 0 to UBound(all_membs_wreg_array)
                      EditBox 10, y_pos, 355, 15, all_membs_wreg_array(the_thing)
                      y_pos = y_pos + 20

                  Next
                  Text 10, y_pos + 5, 50, 10, "Remember:"
                  Text 10, y_pos + 15, 290, 10, "Changing the spaces at the beginning of text fields will change case case note format."
                  Text 10, y_pos + 25, 290, 10, "Use '; ' for a new line in a case note"
                  ButtonGroup ButtonPressed
                    OkButton 315, y_pos + 20, 50, 15
                EndDialog

                dialog Dialog1

                err_msg = "LOOP" & err_msg
            ElseIf ButtonPressed = edit_elig_membs_btn Then
                elig_dlg_hgt = 120 + UBound(elig_membs_array) * 20 + UBOUND(inelig_membs_array) * 20
                Dialog1 = ""
                BeginDialog Dialog1, 0, 0, 371, elig_dlg_hgt, "ELIG MEMB Detail"
                  Text 10, 10, 260, 10, "Add detial about member eligbility as needed"
                  Text 10, 25, 50, 10, "ELIGIBLE"
                  y_pos = 35
                  For the_thing = 0 to UBound(elig_membs_array)
                      EditBox 10, y_pos, 355, 15, elig_membs_array(the_thing)
                      y_pos = y_pos + 20
                  Next
                  Text 10, y_pos, 50, 10, "INELIGIBLE"
                  y_pos = y_pos + 10
                  For the_thing = 0 to UBound(inelig_membs_array)
                      EditBox 10, y_pos, 355, 15, inelig_membs_array(the_thing)
                      y_pos = y_pos + 20
                  Next
                  Text 10, y_pos, 50, 10, "Remember:"
                  Text 10, y_pos + 10, 290, 10, "Changing the spaces at the beginning of text fields will change case case note format."
                  Text 10, y_pos + 20, 290, 10, "Use '; ' for a new line in a case note"
                  ButtonGroup ButtonPressed
                    OkButton 315, y_pos + 15, 50, 15
                EndDialog

                dialog Dialog1

                err_msg = "LOOP" & err_msg
            Else
                If worker_signature = "" Then err_msg = err_msg & vbNewLine & "* Enter your name in the worker signature field for the case note."
            End If
            If err_msg <> "" and left(err_msg, 4) <> "LOOP" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg

        Loop until err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE


    ' CALL add_words_to_message("You receive time-limited SNAP as you are an ABAWD (Able-bodied adult without dependents).")
    ' CALL add_words_to_message("You previously reported that you are homeless, specifically defined for this purpose as lacking both:; *Fixed/regular nighttime residence (inc. temporary housing); *Access to work-related necessities (shower/laundry/etc.); Based on this information, you may qualify for SNAP benefits that are not time-limited. If you believe you meet the homeless and unfit for employment exemption (or any other exemption), please contact your team.")
    ' CALL add_words_to_message("Minnesota has changed the rules for time-limited SNAP recipients. " & abawd_memb_name & " is not required to participate in SNAP Employment and Training (SNAP E&T), but may choose to. Participation in SNAP E&T may extend your SNAP benefits and offer you support as you seek employment. Ask your worker about SNAP E&T.")
    wcom_message = ""
    If abawd_waived_wcom_checkbox = checked Then
        ' wcom_message = "You receive time-limited SNAP as you are an ABAWD (Able-bodied adult without dependents). Currently in Minnesota all ABAWDs are receiving SNAP as Banked Months, a time-limited funding source. Eligibility for this type of SNAP benefit will end 08/31/20, a redetermination will happen at that time."
        wcom_message = "You are considered an ABAWD (Able-bodied adult without dependents) for SNAP (Food Support) purposes. Currently all SNAP participants have no time-limitation on SNAP benefits due to the current National State of Emergency. When this order is no longer in effect, your SNAP eligibility may change."
    End If
    If homeless_wcom_checkbox = checked Then
        If wcom_message = "" Then wcom_message = "You receive time-limited SNAP as you are an ABAWD (Able-bodied adult without dependents). "
        wcom_msg_two = "You previously reported that you are homeless, specifically defined for this purpose as lacking both:; *Fixed/regular nighttime residence (inc. temporary housing); *Access to work-related necessities (shower/laundry/etc.); Based on this information, you may qualify for SNAP benefits that are not time-limited. If you believe you meet the homeless and unfit for employment exemption (or any other exemption), please contact your team."
    End If
    If abawd_waived_wcom_checkbox = checked OR homeless_wcom_checkbox = checked Then
        'Navigate to the correct SPEC screen to select the notice
        Call navigate_to_MAXIS_screen ("SPEC", "WCOM")

        EMWriteScreen MAXIS_footer_month, 3, 46
        EMWriteScreen MAXIS_footer_year, 3, 51

        transmit

        spec_row = 6
        Do
            spec_row = spec_row + 1
            EMReadScreen approval_date, 8, spec_row, 16
            EMReadScreen prog_type, 2, spec_row, 26
            EMReadScreen notice_doc, 20, spec_row, 30
            EMReadScreen print_status, 7, spec_row, 71
            ' MsgBox approval_date
            ' approval_date = DateAdd("d", 0, approval_date)
            If approval_date = todays_date AND prog_type = "FS" AND notice_doc = "ELIG Approval Notice" AND print_status = "Waiting" Then
            'Open the Notice
                EmWriteScreen "X", spec_row, 13
                transmit

                PF9     'Put in to edit mode - the worker comment input screen
                EMSetCursor 03, 13
                Exit Do
            End if
        Loop until approval_date = "        "
        If approval_date = "        " Then
            script_run_lowdown = script_run_lowdown & vbCr & "WCOM(s) selected but failed due to not finding a notice with the approval date: "
            If homeless_wcom_checkbox = checked Then script_run_lowdown = script_run_lowdown & " - Homeless WCOM checkbox - "
            If abawd_waived_wcom_checkbox = checked Then script_run_lowdown = script_run_lowdown & " - ABAWD Waived WCOM checkbox - "
            homeless_wcom_checkbox = unchecked
            abawd_waived_wcom_checkbox = unchecked
            wcom_failed_msg = MsgBox("**** WCOM(s) have not been entered *****" & vbNewLine & vbNewLine & "The script could not find a notice from today for FS ELIG Approval that is still in waiting status." & vbNewLine & "The selected WCOM(s) have not been added.", vbCritical, "WCOM Failure")
        Else
            EMReadScreen wcom_line_one, 60, 3, 15
            wcom_line_one = trim(wcom_line_one)
            If wcom_line_one <> "" Then
                delete_existing_wcom_msg = MsgBox("It appears there is already verbiage in the WCOM." & vbNewLine & vbNewLine & "Would you like to have the script remove the current WCOM and place the requested WCOM?", vbQuestion + vbYesNo, "Replace WCOM")
                If delete_existing_wcom_msg = vbYes Then
                    wcom_row = 3
                    Do
                        wcom_col = 15
                        Do
                            EMWriteScreen " ", wcom_row, wcom_col
                            wcom_col = wcom_col + 1
                        Loop until wcom_col = 75
                        wcom_row = wcom_row + 1
                    Loop until wcom_row = 18
                Else
                    PF4
                    script_run_lowdown = script_run_lowdown & vbCr & "WCOM(s) selected but failed due to a WCOM already existing: "
                    If homeless_wcom_checkbox = checked Then script_run_lowdown = script_run_lowdown & " - Homeless WCOM checkbox - "
                    If abawd_waived_wcom_checkbox = checked Then script_run_lowdown = script_run_lowdown & " - ABAWD Waived WCOM checkbox - "
                    homeless_wcom_checkbox = unchecked
                    abawd_waived_wcom_checkbox = unchecked
                End If
            End If
        End If

        If abawd_waived_wcom_checkbox = checked OR homeless_wcom_checkbox = checked Then
            Call write_variable_in_SPEC_MEMO(wcom_message)
            Call write_variable_in_SPEC_MEMO(wcom_msg_two)
            PF4
            PF3
            PF3
        End If
    End If

    Call start_a_blank_CASE_NOTE

    If MAXIS_footer_month = CM_plus_1_mo AND MAXIS_footer_year = CM_plus_1_yr Then
        Call write_variable_in_CASE_NOTE("SNAP eff " & MAXIS_footer_month & "/" & MAXIS_footer_year & " and Ongoing APP with WAIVERED ABAWD")
    Else
        Call write_variable_in_CASE_NOTE("SNAP eff " & MAXIS_footer_month & "/" & MAXIS_footer_year & " APP with WAIVERED ABAWD")
    End If
    Call write_variable_in_CASE_NOTE("* * * SNAP Benefit : $" & trim(benefit_amount) & " * * *")
    Call write_variable_in_CASE_NOTE("Approval completed on " & date)
    Call write_variable_in_CASE_NOTE("=== Benefit Calculation ===")

    Call write_variable_in_CASE_NOTE("  -Total Earned Income:     $" & total_earned_income)
    Call write_variable_in_CASE_NOTE("  -Total Unearned Income:   $" & total_UNEA)
    Call write_variable_in_CASE_NOTE("* Total Gross Income:       $" & total_gross_income)
    Call write_variable_in_CASE_NOTE("  -Total Deductions:        $ " & total_deduction)
    Call write_variable_in_CASE_NOTE("* Total Net Income:         $" & total_net_income)
    Call write_variable_in_CASE_NOTE("  -Total Shelter Costs:     $" & total_shel_cost)
    Call write_variable_in_CASE_NOTE("  -Half of Net income:      $" & total_shel_cost & " - subtract from costs.")
    Call write_variable_in_CASE_NOTE("  -Adjusted Shelter Costs:  $" & adj_shel_cost)
    Call write_variable_in_CCOL_NOTE("  -MAXIMUM Shelter Cost:    $" & max_allow_shel)
    Call write_variable_in_CASE_NOTE("  -Allowed Shelter Expense: $" & shel_expense)
    Call write_variable_in_CASE_NOTE("* Net Adjusted Income:      $" & net_adj_income)
    Call write_variable_in_CASE_NOTE("  -Monthly FS Allotment:    $" & monthly_fs_allotment)
    Call write_variable_in_CASE_NOTE("  -Recoupment:              $" & recoup_amt)
    Call write_variable_in_CASE_NOTE("* Benefit Amount:           $" & benefit_amount)
    Call write_variable_in_CASE_NOTE("=== Eligibility Summary ===")
    Call write_variable_in_CASE_NOTE("* SNAP is " & trim(current_prog_status))
    Call write_variable_in_CASE_NOTE("* Case is " & trim(elig_result))
    Call write_variable_in_CASE_NOTE("* Reporting: " & trim(reporting_status))
    Call write_variable_in_CASE_NOTE("* Number in Unit: " & trim(numb_in_assistance_unit))
    Call write_variable_in_CASE_NOTE("=== Member Eligibility ===")
    Call write_variable_in_CASE_NOTE("* Household Members Eligibile for SNAP")
    For each member in elig_membs_array
        Call write_variable_in_CASE_NOTE(member)
    Next
    Call write_variable_in_CASE_NOTE("* Household Members NOT Eligibile for SNAP")
    For each member in inelig_membs_array
        Call write_variable_in_CASE_NOTE(member)
    Next
    Call write_variable_in_CASE_NOTE("=== WREG Information ===")
    For each member in all_membs_wreg_array
        Call write_variable_in_CASE_NOTE(member)
    Next
    Call write_variable_in_CASE_NOTE("---")
    Call write_bullet_and_variable_in_CASE_NOTE("Notes", other_notes)
    If homeless_wcom_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Added WCOM regarding potention ABAWD exemption due to homelessness.")
    If abawd_waived_wcom_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Added WCOM with information ABAWD Waiver.")
    Call write_variable_in_CASE_NOTE("Due to the 'Families First Coronavirus Response Act' ABAWD status will be considered waived and no longer time-limited. Cases with ABAWD participants will be coded '09' - Resides in a Waivered Area for the duration of the public health emergency.")
    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)

    next_month = MAXIS_footer_month & "/1/" & MAXIS_footer_year
    next_month = DateAdd("m", 1, next_month)
    Call convert_date_into_MAXIS_footer_month(next_month, MAXIS_footer_month, MAXIS_footer_year)
Loop until MAXIS_footer_month = CM_plus_2_mo AND MAXIS_footer_year = CM_plus_2_yr

script_end_procedure_with_error_report("Success! The script has noted information about a case subject to the Federal ABAWD Waiver. As this is a new script, please report any issue, concern, bug, or enhancement and we will review and update as we can.")
