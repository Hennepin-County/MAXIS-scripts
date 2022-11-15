'Required for statistical purposes==========================================================================================
name_of_script = "UTILITIES - Display Benefits.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 500                     'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
'END OF stats block================================================================================

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
' call run_from_GitHub(script_repository & "application-received.vbs")

'END FUNCTIONS LIBRARYBLOCK================================================================================================

'CHANGELOG BLOCK===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County
CALL changelog_update("11/15/2022", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

function find_last_approved_ELIG_version(cmd_row, cmd_col, version_number, version_date, version_result, approval_found)
	Call write_value_and_transmit("99", cmd_row, cmd_col)
	approval_found = True

	row = 7
	Do
		EMReadScreen elig_version, 2, row, 22
		EMReadScreen elig_date, 8, row, 26
		EMReadScreen elig_result, 10, row, 37
		EMReadScreen approval_status, 10, row, 50

		elig_version = trim(elig_version)
		elig_result = trim(elig_result)
		approval_status = trim(approval_status)

		If approval_status = "APPROVED" Then Exit Do

		row = row + 1
	Loop until approval_status = ""

	Call clear_line_of_text(18, 54)
	If approval_status = "" Then
		approval_found = false
		PF3
	Else
		Call write_value_and_transmit(elig_version, 18, 54)
		version_number = "0" & elig_version
		version_date = elig_date
		version_result = elig_result
	End If
end function


Dim months_to_go_back
Dim run_from_client_contact

const footer_month_const    = 0
const footer_year_const     = 1
const snap_issued_const     = 2
const snap_recoup_const     = 3
const ga_issued_const       = 4
const ga_recoup_const       = 5
const msa_issued_const      = 6
const msa_recoup_const      = 7
const mf_mf_issued_const    = 8
const mf_mf_recoup_const    = 9
const mf_fs_issued_const    = 10
const dwp_issued_const      = 11
const dwp_recoup_const      = 12
const emer_issued_const     = 13
const emer_prog_const       = 14
const grh_issued_const      = 15
const grh_recoup_const      = 16
const no_issuance_const     = 17

const last_const            = 25

Dim ISSUED_BENEFITS_ARRAY()
ReDim ISSUED_BENEFITS_ARRAY(last_const, 0)

Dim ongoing_snap_amount
Dim ongoing_ga_amount
Dim ongoing_msa_amount
Dim ongoing_mfip_cash_amount
Dim ongoing_mfip_food_amount
Dim ongoing_mfip_hg_amount
Dim ongoing_dwp_amount
Dim ongoing_grh_amount_one
Dim ongoing_grh_amount_two


'THE SCRIPT ================================================================================================================
EMConnect ""        'connect to BZ'
CALL MAXIS_case_number_finder(MAXIS_case_number)        'Find CASe Number
MAXIS_footer_month = CM_mo                              'setting the footermonth to the current month
MAXIS_footer_year = CM_yr

'case number dialog
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 196, 105, "Display Benefits"
  EditBox 65, 40, 50, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 85, 85, 50, 15
    CancelButton 140, 85, 50, 15
  Text 10, 15, 150, 20, "This script will display information about the benefits that have been issued or approved."
  Text 15, 45, 50, 10, "Case Number:"
  Text 15, 65, 180, 10, "This script will not CASE/NOTE or create any Notices."
EndDialog

Do
    Do
        err_msg = ""
        dialog Dialog1

        cancel_without_confirmation
        Call validate_MAXIS_case_number("*", MAXIS_case_number)
    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = False

months_to_go_back = 6
run_from_client_contact = False

function gather_case_benefits_details()
    now_month = CM_mo & "/1/" & CM_yr
    now_month = DateAdd("d", 0, now_month)

    subtract_months = 0-months_to_go_back
    start_month = DateAdd("m", subtract_months, now_month)
    start_month_mo = right("00"&DatePart("m", start_month), 2)
    start_month_yr = right(DatePart("yyyy", start_month), 2)
    beginning_footer_month = start_month_mo & "/" & start_month_yr
    month_to_review = start_month

    snap_found = False
    ga_found = False
    msa_found = False
    mfip_found = False
    dwp_found = False
    grh_found = False

    count_months = 0

    Do
        ReDim Preserve ISSUED_BENEFITS_ARRAY(last_const, count_months)
        Call convert_date_into_MAXIS_footer_month(month_to_review, MAXIS_footer_month, MAXIS_footer_year)
        year_to_search = DatePart("yyyy", month_to_review)
        year_to_search = year_to_search & ""
        month_to_search = MonthName(DatePart("m", month_to_review))
        ISSUED_BENEFITS_ARRAY(footer_month_const, count_months) = MAXIS_footer_month
        ISSUED_BENEFITS_ARRAY(footer_year_const, count_months) = MAXIS_footer_year
        ISSUED_BENEFITS_ARRAY(no_issuance_const, count_months) = True

        Call back_to_SELF
        Call navigate_to_MAXIS_screen("MONY", "INQB")
        inqb_row = 6
        Do
            EMReadScreen inqb_month, 12, inqb_row, 3
            EMReadScreen inqb_year, 4, inqb_row, 16
            inqb_month = trim(inqb_month)
            ' MsgBox "inqb_month - " & inqb_month & "-" & vbCr & "month_to_search -" & month_to_search & "-" & vbCr & vbCr & "inqb_year - " & inqb_year & "-" & vbCr & "year_to_search -" & year_to_search & "-" & vbCr & vbCr & "inqb_row - " & inqb_row
            If inqb_month = month_to_search and inqb_year = year_to_search Then
                EMReadScreen inqb_prog, 2, inqb_row, 23
                EMReadScreen inqb_amt, 10, inqb_row, 38
                EMReadScreen inqb_recoup, 10, inqb_row, 53
                EMReadScreen inqb_food, 10, inqb_row, 69
                EMReadScreen inqb_full, 77, inqb_row, 3

                ' MsgBox "inqb_full - " & inqb_full
                If InStr(inqb_full, "FS") <> 0 Then
                ' If inqb_prog = "FS" Then
                    ISSUED_BENEFITS_ARRAY(snap_issued_const, count_months) = trim(inqb_amt)
                    ISSUED_BENEFITS_ARRAY(snap_recoup_const, count_months) = trim(inqb_recoup)
                    ISSUED_BENEFITS_ARRAY(no_issuance_const, count_months) = False
                    snap_found = True
                End If
                If InStr(inqb_full, "GA") <> 0 Then
                ' If inqb_prog = "GA" Then
                    ISSUED_BENEFITS_ARRAY(ga_issued_const, count_months) = trim(inqb_amt)
                    ISSUED_BENEFITS_ARRAY(ga_recoup_const, count_months) = trim(inqb_recoup)
                    ISSUED_BENEFITS_ARRAY(no_issuance_const, count_months) = False
                    ga_found = True
                End If
                If InStr(inqb_full, "MS") <> 0 Then
                ' If inqb_prog = "MS" Then
                    ISSUED_BENEFITS_ARRAY(msa_issued_const, count_months) = trim(inqb_amt)
                    ISSUED_BENEFITS_ARRAY(msa_recoup_const, count_months) = trim(inqb_recoup)
                    ISSUED_BENEFITS_ARRAY(no_issuance_const, count_months) = False
                    msa_found = True
                End If
                If InStr(inqb_full, "MF") <> 0 Then
                ' If inqb_prog = "MF" Then
                    ISSUED_BENEFITS_ARRAY(mf_mf_issued_const, count_months) = trim(inqb_amt)
                    ISSUED_BENEFITS_ARRAY(mf_fs_issued_const, count_months) = trim(inqb_food)
                    ISSUED_BENEFITS_ARRAY(mf_mf_recoup_const, count_months) = trim(inqb_recoup)
                    ISSUED_BENEFITS_ARRAY(no_issuance_const, count_months) = False
                    mfip_found = True
                End If
                If InStr(inqb_full, "DW") <> 0 Then
                ' If inqb_prog = "DW" Then
                    ISSUED_BENEFITS_ARRAY(dwp_issued_const, count_months) = trim(inqb_amt)
                    ISSUED_BENEFITS_ARRAY(dwp_recoup_const, count_months) = trim(inqb_recoup)
                    ISSUED_BENEFITS_ARRAY(no_issuance_const, count_months) = False
                    dwp_found = True
                End If
                If InStr(inqb_full, "GR") <> 0 Then
                ' If inqb_prog = "GR" Then
                    ISSUED_BENEFITS_ARRAY(grh_issued_const, count_months) = trim(inqb_amt)
                    ISSUED_BENEFITS_ARRAY(grh_recoup_const, count_months) = trim(inqb_recoup)
                    ISSUED_BENEFITS_ARRAY(no_issuance_const, count_months) = False
                    grh_found = True
                End If
                If InStr(inqb_full, "AE") <> 0 Then
                ' If inqb_prog = "EA" Then
                    ISSUED_BENEFITS_ARRAY(emer_issued_const, count_months) = trim(inqb_amt)
                    ISSUED_BENEFITS_ARRAY(no_issuance_const, count_months) = False
                End If
            End If

            inqb_row = inqb_row + 1
            EMReadScreen next_prog, 2, inqb_row, 23
        Loop until next_prog = "  "
        Call Back_to_SELF

        month_to_review = DateAdd("m", 1, month_to_review)
        count_months = count_months + 1
    Loop Until DateDiff("d", now_month, month_to_review) > 0

    MAXIS_footer_month = CM_plus_1_mo                              'setting the footermonth to the current month
    MAXIS_footer_year = CM_plus_1_yr


    Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
    Call Back_to_SELF

    If snap_status = "ACTIVE" or snap_status = "APP OPEN" Then
        call navigate_to_MAXIS_screen("ELIG", "FS  ")
        Call find_last_approved_ELIG_version(19, 78, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
        Call write_value_and_transmit("FSSM", 19, 70)

        EMReadScreen snap_benefit_monthly_fs_allotment, 10, 8, 71
        EMReadScreen snap_benefit_prorated_amt, 		10, 9, 71
        EMReadScreen snap_benefit_prorated_date,		8, 9, 58
        EMReadScreen snap_benefit_amt, 					10, 13, 71

        snap_benefit_monthly_fs_allotment = trim(snap_benefit_monthly_fs_allotment)
        snap_benefit_prorated_amt = trim(snap_benefit_prorated_amt)
        snap_benefit_prorated_date = trim(snap_benefit_prorated_date)
        ongoing_snap_amount = trim(snap_benefit_amt)

        Call Back_to_SELF
    End If
    If ga_status = "ACTIVE" or ga_status = "APP OPEN" Then
        call navigate_to_MAXIS_screen("ELIG", "GA  ")
        Call find_last_approved_ELIG_version(20, 78, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
        Call write_value_and_transmit("GASM", 20, 70)

        EMReadScreen ga_elig_summ_monthly_grant, 10, 9, 71
        EMReadScreen ga_elig_summ_amount_to_be_paid, 10, 14, 71

        ga_elig_summ_monthly_grant = trim(ga_elig_summ_monthly_grant)
        ongoing_ga_amount = trim(ga_elig_summ_amount_to_be_paid)

        Call Back_to_SELF
    End If
    If msa_status = "ACTIVE" or msa_status = "APP OPEN" Then
        call navigate_to_MAXIS_screen("ELIG", "MSA ")
        Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
        Call write_value_and_transmit("MSSM", 20, 71)

        EMReadScreen msa_elig_summ_grant, 9, 11, 72
        EMReadScreen msa_elig_summ_current_payment, 9, 17, 72

        msa_elig_summ_grant = trim(msa_elig_summ_grant)
        ongoing_msa_amount = trim(msa_elig_summ_current_payment)

        Call Back_to_SELF
    End If
    If mfip_status = "ACTIVE" or mfip_status = "APP OPEN" Then
        call navigate_to_MAXIS_screen("ELIG", "MFIP")
        Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
        Call write_value_and_transmit("MFSM", 20, 71)

        EMReadScreen mfip_case_summary_grant_amount, 10, 11, 71
        EMReadScreen mfip_case_summary_net_grant_amount, 10, 13, 71
        EMReadScreen mfip_case_summary_cash_portion, 10, 14, 71
        EMReadScreen mfip_case_summary_food_portion, 10, 15, 71
        EMReadScreen mfip_case_summary_housing_grant, 10, 16, 71

        mfip_case_summary_grant_amount = trim(mfip_case_summary_grant_amount)
        mfip_case_summary_net_grant_amount = trim(mfip_case_summary_net_grant_amount)
        ongoing_mfip_cash_amount = trim(mfip_case_summary_cash_portion)
        ongoing_mfip_food_amount = trim(mfip_case_summary_food_portion)
        ongoing_mfip_hg_amount = trim(mfip_case_summary_housing_grant)

        Call Back_to_SELF
    End If
    If dwp_status = "ACTIVE" or dwp_status = "APP OPEN" Then
        call navigate_to_MAXIS_screen("ELIG", "DWP ")
        Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
        Call write_value_and_transmit("DWSM", 20, 71)


        EMReadScreen dwp_case_summary_grant_amount, 10, 10, 71
        EMReadScreen dwp_case_summary_net_grant_amount, 10, 12, 71
        EMReadScreen dwp_case_summary_shelter_benefit_portion, 10, 13, 71
        EMReadScreen dwp_case_summary_personal_needs_portion, 10, 14, 71

        dwp_case_summary_grant_amount = trim(dwp_case_summary_grant_amount)
        ongoing_dwp_amount = trim(dwp_case_summary_net_grant_amount)
        dwp_case_summary_shelter_benefit_portion = trim(dwp_case_summary_shelter_benefit_portion)
        dwp_case_summary_personal_needs_portion = trim(dwp_case_summary_personal_needs_portion)

        Call Back_to_SELF
    End If
    If grh_status = "ACTIVE" or ga_stagrh_statustus = "APP OPEN" Then
        call navigate_to_MAXIS_screen("ELIG", "GRH ")
        Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
        Call write_value_and_transmit("GRSM", 20, 71)

        EMReadScreen ongoing_grh_amount_one, 		9, 12, 31
        EMReadScreen ongoing_grh_amount_two, 		9, 12, 50

        ongoing_grh_amount_one = trim(ongoing_grh_amount_one)
        ongoing_grh_amount_two = trim(ongoing_grh_amount_two)

        Call Back_to_SELF
    End If




    Do
        Do

            programs_with_no_cm_plus_one_issuance = ""
            If ongoing_snap_amount = "" Then programs_with_no_cm_plus_one_issuance = programs_with_no_cm_plus_one_issuance & ", SNAP"
            If ongoing_ga_amount = "" Then programs_with_no_cm_plus_one_issuance = programs_with_no_cm_plus_one_issuance & ", GA"
            If ongoing_msa_amount = "" Then programs_with_no_cm_plus_one_issuance = programs_with_no_cm_plus_one_issuance & ", MSA"
            If ongoing_mfip_cash_amount = "" and ongoing_mfip_food_amount = "" and ongoing_mfip_hg_amount = "" Then programs_with_no_cm_plus_one_issuance = programs_with_no_cm_plus_one_issuance & ", MFIP"
            If ongoing_dwp_amount = "" Then programs_with_no_cm_plus_one_issuance = programs_with_no_cm_plus_one_issuance & ", DWP"
            If ongoing_grh_amount_one = "" Then programs_with_no_cm_plus_one_issuance = programs_with_no_cm_plus_one_issuance & ", GRH"
            if left(programs_with_no_cm_plus_one_issuance, 1) = "," Then programs_with_no_cm_plus_one_issuance = right(programs_with_no_cm_plus_one_issuance, len(programs_with_no_cm_plus_one_issuance)-1)
            programs_with_no_cm_plus_one_issuance = trim(programs_with_no_cm_plus_one_issuance)

            programs_with_no_past_issuance = ""
            If snap_found = False Then programs_with_no_past_issuance = programs_with_no_past_issuance & ", SNAP"
            If ga_found = False Then programs_with_no_past_issuance = programs_with_no_past_issuance & ", GA"
            If msa_found = False Then programs_with_no_past_issuance = programs_with_no_past_issuance & ", MSA"
            If mfip_found = False Then programs_with_no_past_issuance = programs_with_no_past_issuance & ", MFIP"
            If dwp_found = False Then programs_with_no_past_issuance = programs_with_no_past_issuance & ", DWP"
            If grh_found = False Then programs_with_no_past_issuance = programs_with_no_past_issuance & ", GRH"
            if left(programs_with_no_past_issuance, 1) = "," Then programs_with_no_past_issuance = right(programs_with_no_past_issuance, len(programs_with_no_past_issuance)-1)
            programs_with_no_past_issuance = trim(programs_with_no_past_issuance)
            ' MsgBox "programs_with_no_past_issuance - " & programs_with_no_past_issuance

            prog_count = 1
            If mfip_found = True Then prog_count = prog_count + 1
            If snap_found = True Then prog_count = prog_count + 1
            If ga_found = True Then prog_count = prog_count + 1
            If msa_found = True Then prog_count = prog_count + 1
            If dwp_found = True Then prog_count = prog_count + 1
            If grh_found = True Then prog_count = prog_count + 1
            prog_len_multiplier = prog_count/2
            prog_len_multiplier = INT(prog_len_multiplier)
            grp_bx_len = 45
            grp_bx_len = grp_bx_len + prog_len_multiplier * 25
            grp_bx_len = grp_bx_len + months_to_go_back * 10 * prog_len_multiplier

            dlg_len = 160 + grp_bx_len

            err_msg = ""

            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 441, dlg_len, "Case " & MAXIS_case_number & " Issuance Details"
             ButtonGroup ButtonPressed
                EditBox 500, 600, 50, 15, fake_edit_box
                GroupBox 10, 10, 420, 105, "Current Approval Amounts"
                Text 20, 25, 180, 10, "Based on ELIG for current month plus 1  (" & CM_plus_1_mo & "/" & CM_plus_1_yr &")"

                x_pos = 30
                If ongoing_snap_amount <> "" Then
                    Text x_pos, 40, 25, 10, "SNAP"
                    Text x_pos+5, 50, 30, 10, "$ " & ongoing_snap_amount
                    PushButton x_pos, 80, 35, 10, "ELIG/FS", elig_fs_btn
                    x_pos = x_pos + 60
                End If
                If ongoing_ga_amount <> "" Then
                    Text x_pos, 40, 25, 10, "GA"
                    Text x_pos+5, 50, 30, 10, "$ " & ongoing_ga_amount
                    PushButton x_pos, 80, 35, 10, "ELIG/GA", elig_ga_btn
                    x_pos = x_pos + 60
                End If
                If ongoing_msa_amount <> "" Then
                    Text x_pos, 40, 25, 10, "MSA"
                    Text x_pos+5, 50, 30, 10, "$ " & ongoing_msa_amount
                    PushButton x_pos, 80, 40, 10, "ELIG/MSA", elig_msa_btn
                    x_pos = x_pos + 65
                End If
                If ongoing_mfip_cash_amount <> "" or ongoing_mfip_food_amount <> "" or ongoing_mfip_hg_amount <> "" Then
                    Text x_pos, 40, 25, 10, "MFIP"
                    Text x_pos+5, 50, 40, 10, "MF - $ " & ongoing_mfip_cash_amount
                    Text x_pos+5, 60, 40, 10, "FS - $ " & ongoing_mfip_food_amount
                    Text x_pos+5, 70, 40, 10, "HG - $ " & ongoing_mfip_hg_amount
                    PushButton x_pos, 80, 45, 10, "ELIG/MFIP", elig_mfip_btn
                    x_pos = x_pos + 70
                End If
                If ongoing_dwp_amount <> "" Then
                    Text x_pos, 40, 25, 10, "DWP"
                    Text 300, 50, 30, 10, "$ " & ongoing_dwp_amount
                    PushButton x_pos, 80, 40, 10, "ELIG/DWP", elig_dwp_btn
                    x_pos = x_pos + 65
                End If
                If ongoing_grh_amount_one <> "" Then
                    Text x_pos, 40, 25, 10, "GRH"
                    Text x_pos+5, 50, 45, 10, "One - $ " & ongoing_grh_amount_one
                    if ongoing_grh_amount_two <> "" Then Text x_pos+5, 60, 45, 10, "Two - $ " & ongoing_grh_amount_two
                    PushButton x_pos, 80, 40, 10, "ELIG/GRH", elig_grh_btn
                End If
                Text 140, 100, 280, 10, "No Eligibility for: " & programs_with_no_cm_plus_one_issuance



                ' GroupBox 10, 125, 420, 150, "Past Issuance Amounts"
                ' Text 25, 140, 155, 10, "Information going back 6 months to MM/YY"
                ' ButtonGroup ButtonPressed
                '   PushButton 265, 135, 160, 15, "Change the Number of Months to Go Back", change_lookbackk_month_count_btn
                ' Text 30, 155, 35, 10, "SNAP"
                ' Text 40, 165, 150, 10, "MM/YY  .  .  .  $ 1234        Recoup: $ 52"
                ' Text 40, 175, 150, 10, "MM/YY  .  .  .  $ 1234        Recoup: $ 52"
                ' Text 40, 185, 150, 10, "MM/YY  .  .  .  $ 1234        Recoup: $ 52"
                ' Text 260, 155, 35, 10, "GA"
                ' Text 270, 165, 150, 10, "MM/YY  .  .  .  $ 1234        Recoup: $ 52"
                ' Text 270, 175, 150, 10, "MM/YY  .  .  .  $ 1234        Recoup: $ 52"
                ' Text 270, 185, 150, 10, "MM/YY  .  .  .  $ 1234        Recoup: $ 52"
                ' Text 30, 205, 35, 10, "MFIP"
                ' Text 40, 215, 205, 10, "MM/YY  .  .  .  Cash $ 1234  -  Food $  4321      Recoup: $ 52"
                ' Text 40, 225, 205, 10, "MM/YY  .  .  .  Cash $ 1234  -  Food $  4321      Recoup: $ 52"
                ' Text 40, 235, 205, 10, "MM/YY  .  .  .  Cash $ 1234  -  Food $  4321      Recoup: $ 52"
                ' Text 260, 205, 35, 10, "GA"
                ' Text 270, 215, 150, 10, "MM/YY  .  .  .  $ 1234        Recoup: $ 52"
                ' Text 270, 225, 150, 10, "MM/YY  .  .  .  $ 1234        Recoup: $ 52"
                ' Text 270, 235, 150, 10, "MM/YY  .  .  .  $ 1234        Recoup: $ 52"
                ' Text 135, 260, 295, 10, "No Issuances for: "
                '

                GroupBox 10, 125, 420, grp_bx_len, "Past Issuance Amounts"
                Text 25, 140, 155, 10, "Information going back " & months_to_go_back & " months to " & beginning_footer_month
                PushButton 265, 135, 160, 15, "Change the Number of Months to Go Back", change_lookbackk_month_count_btn

                x_pos = 30
                y_pos = 155
                y_pos_reset = 155

                If mfip_found = True Then
                    Text x_pos, y_pos, 35, 10, "MFIP"
                    y_pos = y_pos + 10

                    For each_mf_issue = 0 to UBound(ISSUED_BENEFITS_ARRAY, 2)
                        If ISSUED_BENEFITS_ARRAY(mf_mf_recoup_const, each_mf_issue) <> "" Then Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(footer_month_const, each_mf_issue) & "/" & ISSUED_BENEFITS_ARRAY(footer_year_const, each_mf_issue) & "  .  .  .  Cash $ " & ISSUED_BENEFITS_ARRAY(mf_mf_issued_const, each_mf_issue) & "  -  Food $  " & ISSUED_BENEFITS_ARRAY(mf_fs_issued_const, each_mf_issue) & "        Recoup: $ " & ISSUED_BENEFITS_ARRAY(mf_mf_recoup_const, each_mf_issue)
                        If ISSUED_BENEFITS_ARRAY(mf_mf_recoup_const, each_mf_issue) = "" Then Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(footer_month_const, each_mf_issue) & "/" & ISSUED_BENEFITS_ARRAY(footer_year_const, each_mf_issue) & "  .  .  .  Cash $ " & ISSUED_BENEFITS_ARRAY(mf_mf_issued_const, each_mf_issue) & "  -  Food $  " & ISSUED_BENEFITS_ARRAY(mf_fs_issued_const, each_mf_issue)
                        y_pos = y_pos + 10
                    Next
                    y_pos_end = y_pos
                    If x_pos = 30 Then
                        x_pos = 260
                        y_pos = y_pos_reset
                    ElseIf x_pos = 260 Then
                        x_pos = 30
                        y_pos_reset = y_pos + 10
                        y_pos = y_pos_reset
                    End If
                End If

                If snap_found = True Then
                    Text x_pos, y_pos, 35, 10, "SNAP"
                    y_pos = y_pos + 10

                    For each_fs_issue = 0 to UBound(ISSUED_BENEFITS_ARRAY, 2)
                        If ISSUED_BENEFITS_ARRAY(snap_recoup_const, each_fs_issue) <> "" Then Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(footer_month_const, each_fs_issue) & "/" & ISSUED_BENEFITS_ARRAY(footer_year_const, each_fs_issue) & "  .  .  .  $ " & ISSUED_BENEFITS_ARRAY(snap_issued_const, each_fs_issue) & "        Recoup: $ " & ISSUED_BENEFITS_ARRAY(snap_recoup_const, each_fs_issue)
                        If ISSUED_BENEFITS_ARRAY(snap_recoup_const, each_fs_issue) = "" Then Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(footer_month_const, each_fs_issue) & "/" & ISSUED_BENEFITS_ARRAY(footer_year_const, each_fs_issue) & "  .  .  .  $ " & ISSUED_BENEFITS_ARRAY(snap_issued_const, each_fs_issue)
                        y_pos = y_pos + 10
                    Next
                    y_pos_end = y_pos
                    If x_pos = 30 Then
                        x_pos = 260
                        y_pos = y_pos_reset
                    ElseIf x_pos = 260 Then
                        x_pos = 30
                        y_pos_reset = y_pos + 10
                        y_pos = y_pos_reset
                    End If
                End If

                If ga_found = True Then
                    Text x_pos, y_pos, 35, 10, "GA"
                    y_pos = y_pos + 10

                    For each_ga_issue = 0 to UBound(ISSUED_BENEFITS_ARRAY, 2)
                        If ISSUED_BENEFITS_ARRAY(ga_recoup_const, each_ga_issue) <> "" Then Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(footer_month_const, each_ga_issue) & "/" & ISSUED_BENEFITS_ARRAY(footer_year_const, each_ga_issue) & "  .  .  .  $ " & ISSUED_BENEFITS_ARRAY(ga_issued_const, each_ga_issue) & "        Recoup: $ " & ISSUED_BENEFITS_ARRAY(ga_recoup_const, each_ga_issue)
                        If ISSUED_BENEFITS_ARRAY(ga_recoup_const, each_ga_issue) = "" Then Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(footer_month_const, each_ga_issue) & "/" & ISSUED_BENEFITS_ARRAY(footer_year_const, each_ga_issue) & "  .  .  .  $ " & ISSUED_BENEFITS_ARRAY(ga_issued_const, each_ga_issue)
                        y_pos = y_pos + 10
                    Next
                    y_pos_end = y_pos
                    If x_pos = 30 Then
                        x_pos = 260
                        y_pos = y_pos_reset
                    ElseIf x_pos = 260 Then
                        x_pos = 30
                        y_pos_reset = y_pos + 10
                        y_pos = y_pos_reset
                    End If
                End If

                If msa_found = True Then
                    Text x_pos, y_pos, 35, 10, "MSA"
                    y_pos = y_pos + 10

                    For each_msa_issue = 0 to UBound(ISSUED_BENEFITS_ARRAY, 2)
                        If ISSUED_BENEFITS_ARRAY(msa_recoup_const, each_msa_issue) <> "" Then Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(footer_month_const, each_msa_issue) & "/" & ISSUED_BENEFITS_ARRAY(footer_year_const, each_msa_issue) & "  .  .  .  $ " & ISSUED_BENEFITS_ARRAY(msa_issued_const, each_msa_issue) & "        Recoup: $ " & ISSUED_BENEFITS_ARRAY(msa_recoup_const, each_msa_issue)
                        If ISSUED_BENEFITS_ARRAY(msa_recoup_const, each_msa_issue) = "" Then Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(footer_month_const, each_msa_issue) & "/" & ISSUED_BENEFITS_ARRAY(footer_year_const, each_msa_issue) & "  .  .  .  $ " & ISSUED_BENEFITS_ARRAY(msa_issued_const, each_msa_issue)
                        y_pos = y_pos + 10
                    Next
                    y_pos_end = y_pos
                    If x_pos = 30 Then
                        x_pos = 260
                        y_pos = y_pos_reset
                    ElseIf x_pos = 260 Then
                        x_pos = 30
                        y_pos_reset = y_pos + 10
                        y_pos = y_pos_reset
                    End If
                End If

                If dwp_found = True Then
                    Text x_pos, y_pos, 35, 10, "DWP"
                    y_pos = y_pos + 10

                    For each_dwp_issue = 0 to UBound(ISSUED_BENEFITS_ARRAY, 2)
                        If ISSUED_BENEFITS_ARRAY(dwp_recoup_const, each_dwp_issue) <> "" Then Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(footer_month_const, each_dwp_issue) & "/" & ISSUED_BENEFITS_ARRAY(footer_year_const, each_dwp_issue) & "  .  .  .  $ " & ISSUED_BENEFITS_ARRAY(dwp_issued_const, each_dwp_issue) & "        Recoup: $ " & ISSUED_BENEFITS_ARRAY(dwp_recoup_const, each_dwp_issue)
                        If ISSUED_BENEFITS_ARRAY(dwp_recoup_const, each_dwp_issue) = "" Then Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(footer_month_const, each_dwp_issue) & "/" & ISSUED_BENEFITS_ARRAY(footer_year_const, each_dwp_issue) & "  .  .  .  $ " & ISSUED_BENEFITS_ARRAY(dwp_issued_const, each_dwp_issue)
                        y_pos = y_pos + 10
                    Next
                    y_pos_end = y_pos
                    If x_pos = 30 Then
                        x_pos = 260
                        y_pos = y_pos_reset
                    ElseIf x_pos = 260 Then
                        x_pos = 30
                        y_pos_reset = y_pos + 10
                        y_pos = y_pos_reset
                    End If
                End If

                If grh_found = True Then
                    Text x_pos, y_pos, 35, 10, "GRH"
                    y_pos = y_pos + 10

                    For each_grh_issue = 0 to UBound(ISSUED_BENEFITS_ARRAY, 2)
                        If ISSUED_BENEFITS_ARRAY(grh_recoup_const, each_grh_issue) <> "" Then Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(footer_month_const, each_grh_issue) & "/" & ISSUED_BENEFITS_ARRAY(footer_year_const, each_grh_issue) & "  .  .  .  $ " & ISSUED_BENEFITS_ARRAY(grh_issued_const, each_grh_issue) & "        Recoup: $ " & ISSUED_BENEFITS_ARRAY(grh_recoup_const, each_grh_issue)
                        If ISSUED_BENEFITS_ARRAY(grh_recoup_const, each_grh_issue) = "" Then Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(footer_month_const, each_grh_issue) & "/" & ISSUED_BENEFITS_ARRAY(footer_year_const, each_grh_issue) & "  .  .  .  $ " & ISSUED_BENEFITS_ARRAY(grh_issued_const, each_grh_issue)
                        y_pos = y_pos + 10
                    Next
                    y_pos_end = y_pos
                    If x_pos = 30 Then
                        x_pos = 260
                        y_pos = y_pos_reset
                    ElseIf x_pos = 260 Then
                        x_pos = 30
                        y_pos_reset = y_pos + 10
                        y_pos = y_pos_reset
                    End If
                End If

                ' Text 30, 155, 35, 10, "SNAP"
                ' Text 40, 165, 150, 10, "MM/YY  .  .  .  $ 1234        Recoup: $ 52"
                ' Text 40, 175, 150, 10, "MM/YY  .  .  .  $ 1234        Recoup: $ 52"
                ' Text 40, 185, 150, 10, "MM/YY  .  .  .  $ 1234        Recoup: $ 52"
                ' Text 260, 155, 35, 10, "GA"
                ' Text 270, 165, 150, 10, "MM/YY  .  .  .  $ 1234        Recoup: $ 52"
                ' Text 270, 175, 150, 10, "MM/YY  .  .  .  $ 1234        Recoup: $ 52"
                ' Text 270, 185, 150, 10, "MM/YY  .  .  .  $ 1234        Recoup: $ 52"
                ' Text 30, 205, 35, 10, "MFIP"
                ' Text 40, 215, 205, 10, "MM/YY  .  .  .  Cash $ 1234  -  Food $  4321      Recoup: $ 52"
                ' Text 40, 225, 205, 10, "MM/YY  .  .  .  Cash $ 1234  -  Food $  4321      Recoup: $ 52"
                ' Text 40, 235, 205, 10, "MM/YY  .  .  .  Cash $ 1234  -  Food $  4321      Recoup: $ 52"
                ' Text 260, 205, 35, 10, "GA"
                ' Text 270, 215, 150, 10, "MM/YY  .  .  .  $ 1234        Recoup: $ 52"
                ' Text 270, 225, 150, 10, "MM/YY  .  .  .  $ 1234        Recoup: $ 52"
                ' Text 270, 235, 150, 10, "MM/YY  .  .  .  $ 1234        Recoup: $ 52"
                Text 135, y_pos_end+5, 295, 10, "No Issuances for: " & programs_with_no_past_issuance
                If run_from_client_contact = False Then
                    PushButton 15, dlg_len-25, 160, 15, "Run NOTICES - PA Verifications Request", run_pa_verif_reqquest__btn
                    PushButton 185, dlg_len-25, 135, 15, "Run NOTES - Client Contact", run_client_contact_btn
                End If
                PushButton 330, dlg_len-25, 100, 15, "End Script Run", complete_script_run_btn
            EndDialog

            dialog Dialog1
            cancel_without_confirmation

        Loop until err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = False


end function

Call gather_case_benefits_details

Call script_end_procedure("")
