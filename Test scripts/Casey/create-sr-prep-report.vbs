'Required for statistical purposes===============================================================================
name_of_script = "ADMIN - CREATE SR PREP REPORT.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 13                      'manual run time in seconds
STATS_denomination = "C"       							'C is for each CASE
'END OF stats block==============================================================================================

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

'Open Excel
    'open DAIL sheet
        'find all cases with SR MASS dail info and add to an array
    'open ACTV sheet
    'Read cases and add any to an array with the following criteria:
        'Case has a next REVW date of CM + 2 or CM + 8 or in the past
        'Case has MF or GA active or FS if dail exists

'Find most recent approval - if reporting status is Six-Month - turn off 'assignment'

function keep_MX_passworded_in()
    ' MsgBox running_stopwatch & vbNewLine & timer
    If timer - running_stopwatch > 720 Then         'this means the script has been running for more than 12 minutes since we last popped in to MMIS
        Call navigate_to_MAXIS_screen("REPT", "ACTV")      'Going to MMIS'
        Call back_to_SELF
		'MsgBox "In MMIS"
        ' Call navigate_to_MAXIS(maxis_area)                       'going back to MAXIS'
        'MsgBox "Back to MAXIS"
        running_stopwatch = timer                                       'resetting the stopwatch'
    End If
end function

'COLUMN CONSTANTS
const actv_worker_col       = 1
const actv_case_numb_col    = 2
const actv_name_col         = 3
const actv_next_revw_col    = 4
const actv_snap_col         = 5
const actv_cash_col         = 6

const dail_x1numb_col       = 1
const dail_case_numb_col    = 2
const dail_client_name_col  = 3
const dail_type_col         = 4
const dail_month_col        = 5
const dail_message_col      = 6

const wrklst_case_numb_col      = 1
const wrklst_case_name_col      = 2
const wrklst_next_revw_col      = 3
const wrklst_dail_col           = 4
const wrklst_snap_col           = 5
const wrklst_snap_app_date_col  = 6
const wrklst_snap_rept_status_col = 7
const wrklst_mfip_col           = 8
const wrklst_mfip_app_date_col  = 9
const wrklst_mfip_rept_status_col = 10
const wrklst_ga_col             = 11
const wrklst_ga_app_date_col    = 12
const wrklst_ga_rept_status_col = 13
const wrklst_review_case_col    = 14

' const wrklst__col         = 1
' const wrklst__col         = 1
' const wrklst__col         = 1
' const wrklst__col         = 1
' const wrklst__col         = 1


const case_numb_const           = 00
const case_name_const           = 01
const MF_case_const             = 02
const GA_case_const             = 03
const FS_case_const             = 04
const next_revw_const           = 05
const MF_rept_status_const      = 06
const GA_rept_status_const      = 07
const FS_rept_status_const      = 08
const MF_approval_date_const    = 09
const GA_approval_date_const    = 10
const FS_approval_date_const    = 11
const dail_exists_const         = 12
const review_case_const         = 13
const last_const                = 20

Dim CASES_ARRAY()
ReDim CASES_ARRAY(last_const, 0)

CM_plus_2_date = CM_plus_2_mo & "/1/" & CM_plus_2_yr
CM_plus_2_date = DateAdd("d", 0, CM_plus_2_date)
CM_plus_8_date = DateAdd("m", 6, CM_plus_2_date)

file_name = CM_plus_2_mo & "-" & CM_plus_2_yr & " SR Preparation Report.xlsx"
file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\Cash Six-Month Reporting Transition\" & file_name
visible_status = True
alerts_status = False
Call excel_open(file_path, visible_status, alerts_status, ObjExcel, objWorkbook)

objExcel.worksheets("DAIL").Activate

Dim SR_MASS_DAIL_ARRAY()
ReDim SR_MASS_DAIL_ARRAY(0)
cow = 0
running_stopwatch = timer

excel_row = 2
Do
    is_mass_dail = False
    If ObjExcel.Cells(excel_row, dail_message_col).value = "UHFS MASS CHANGE NOT AUTO-APPROVED - REVIEW BUDGET SEE PF12" Then is_mass_dail = True
    If ObjExcel.Cells(excel_row, dail_message_col).value = "UHFS MASS CHANGE NOT AUTO-APPROVED - REVIEW BUDGET SEE PF12" Then is_mass_dail = True
    If ObjExcel.Cells(excel_row, dail_message_col).value = "GA: NOT AUTO-APPROVED - SEE EXPLANATION*" Then is_mass_dail = True
    If ObjExcel.Cells(excel_row, dail_month_col).value <> "08 25" Then is_mass_dail = False
    If is_mass_dail Then
        ReDim preserve SR_MASS_DAIL_ARRAY(cow)
        SR_MASS_DAIL_ARRAY(cow) = trim(objExcel.Cells(excel_row, dail_case_numb_col).Value)
        cow = cow + 1
    End If
    excel_row = excel_row + 1
Loop While objExcel.Cells(excel_row, dail_case_numb_col).Value <> ""


objExcel.worksheets("ACTV").Activate
excel_row = 2
horse = 0
Do
    add_to_worklist = False
    revw_month_to_process = False
    prog_to_process = False
    If IsDate(ObjExcel.Cells(excel_row, actv_next_revw_col).Value) Then
        next_revw_date = CDate(ObjExcel.Cells(excel_row, actv_next_revw_col).Value)
        If DateDiff("d", next_revw_date, CM_plus_2_date) = 0 Then revw_month_to_process = True
        If DateDiff("d", next_revw_date, CM_plus_8_date) = 0 Then revw_month_to_process = True
        If DateDiff("d", next_revw_date, Date) > 0 Then revw_month_to_process = True
    Else
        revw_month_to_process = True 'No next review date
    End If

    dail_there = ""
    If revw_month_to_process Then
        If InStr(ObjExcel.Cells(excel_row, actv_cash_col).Value, "GA A") or InStr(ObjExcel.Cells(excel_row, actv_cash_col).Value, "MF A") Then
            prog_to_process = True
        ElseIf ObjExcel.Cells(excel_row, actv_snap_col).Value = "A" Then
            dail_there = False
            For each duck in SR_MASS_DAIL_ARRAY
                If duck = trim(ObjExcel.Cells(excel_row, actv_case_numb_col).Value) Then
                    prog_to_process = True
                    dail_there = True
                    Exit For
                End If
            Next
        End If
    End If

    If revw_month_to_process and prog_to_process Then add_to_worklist = True

    If add_to_worklist Then
        ReDim Preserve CASES_ARRAY(last_const, horse)
        CASES_ARRAY(case_numb_const, horse) = trim(ObjExcel.Cells(excel_row, actv_case_numb_col).Value)
        CASES_ARRAY(case_name_const, horse) = trim(ObjExcel.Cells(excel_row, actv_name_col).Value)
        CASES_ARRAY(next_revw_const, horse) = ObjExcel.Cells(excel_row, actv_next_revw_col).Value
        CASES_ARRAY(review_case_const, horse) = True

        CASES_ARRAY(GA_case_const, horse) = False
        If InStr(ObjExcel.Cells(excel_row, actv_cash_col).Value, "GA A") Then CASES_ARRAY(GA_case_const, horse) = True
        CASES_ARRAY(MF_case_const, horse) = False
        If InStr(ObjExcel.Cells(excel_row, actv_cash_col).Value, "MF A") Then CASES_ARRAY(MF_case_const, horse) = True
        CASES_ARRAY(FS_case_const, horse) = False
        If ObjExcel.Cells(excel_row, actv_snap_col).Value = "A" Then CASES_ARRAY(FS_case_const, horse) = True

        If dail_there = "" Then
            CASES_ARRAY(dail_exists_const, horse) = False
            For each duck in SR_MASS_DAIL_ARRAY
                If duck = trim(ObjExcel.Cells(excel_row, actv_case_numb_col).Value) Then
                    CASES_ARRAY(dail_exists_const, horse) = True
                    Exit For
                End If
            Next
        Else
            CASES_ARRAY(dail_exists_const, horse) = dail_there
        End If
        horse = horse + 1
    End If
    ' item_info = "excel_row - " & excel_row & vbCr &_
    '        "case number - " & ObjExcel.Cells(excel_row, actv_case_numb_col).Value & vbCr &_
    '        "next review date - " & ObjExcel.Cells(excel_row, actv_next_revw_col).Value & vbCr &_
    '        "Cash col - " & ObjExcel.Cells(excel_row, actv_cash_col).Value & vbCr &_
    '        "revw_month_to_process - " & revw_month_to_process & vbCr &_
    '        "prog_to_process - " & prog_to_process & vbCr &_
    '        "add_to_worklist - " & add_to_worklist & vbCr

    ' If add_to_worklist Then
    '     item_info = item_info & "GA case - " & CASES_ARRAY(GA_case_const, horse - 1) & vbCr
    '     item_info = item_info & "MF case - " & CASES_ARRAY(MF_case_const, horse - 1) & vbCr
    '     item_info = item_info & "FS case - " & CASES_ARRAY(FS_case_const, horse - 1) & vbCr
    '     item_info = item_info & "dail exists - " & CASES_ARRAY(dail_exists_const, horse - 1) & vbCr
    '     item_info = item_info & "review case - " & CASES_ARRAY(review_case_const, horse - 1)
    ' End If
    ' MsgBox item_info
    excel_row = excel_row + 1
    call keep_MX_passworded_in
    ' If excel_row = 2002 Then
    '     ' MsgBox "Elapsed time - " & (timer - start_time) & " seconds."
    '     Exit Do
    ' End If
Loop While ObjExcel.Cells(excel_row, actv_case_numb_col).Value <> ""

objExcel.worksheets("Worklist Partial").Activate
excel_row = 2
For duck = 0 to UBound(CASES_ARRAY, 2)
    ObjExcel.Cells(excel_row, wrklst_case_numb_col).Value = CASES_ARRAY(case_numb_const, duck)
    ObjExcel.Cells(excel_row, wrklst_case_name_col).Value = CASES_ARRAY(case_name_const, duck)
    ObjExcel.Cells(excel_row, wrklst_next_revw_col).Value = CASES_ARRAY(next_revw_const, duck)
    ObjExcel.Cells(excel_row, wrklst_dail_col).Value = CASES_ARRAY(dail_exists_const, duck)
    If CASES_ARRAY(FS_case_const, duck) Then ObjExcel.Cells(excel_row, wrklst_snap_col).Value = "True"
    If CASES_ARRAY(MF_case_const, duck) Then ObjExcel.Cells(excel_row, wrklst_mfip_col).Value = "True"
    If CASES_ARRAY(GA_case_const, duck) Then ObjExcel.Cells(excel_row, wrklst_ga_col).Value = "True"
    excel_row = excel_row + 1
Next

' MsgBox "The SR Preparation Report has been created. Please check the file " & file_path & "."
Call check_for_MAXIS(False)
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr
Call back_to_SELF
For duck = 0 to UBound(CASES_ARRAY, 2)
    MAXIS_case_number = CASES_ARRAY(case_numb_const, duck)
    Call navigate_to_MAXIS_screen_review_PRIV("STAT", "SUMM", is_this_priv)
    need_review = False

    If is_this_priv = False Then
        Call navigate_to_MAXIS_screen("ELIG", "    ")
        If CASES_ARRAY(MF_case_const, duck) Then
            Call write_value_and_transmit("MFIP", 20, 71)
            Call find_last_approved_ELIG_version(20, 79, version_number, version_date, version_result, approval_found)
            CASES_ARRAY(MF_approval_date_const, duck) = version_date
            Call write_value_and_transmit("MFSM", 20, 71)
            EMReadScreen mf_rept_status, 9, 8, 31
            CASES_ARRAY(MF_rept_status_const, duck) = mf_rept_status
            If trim(mf_rept_status) = "MONTHLY" Then need_review = True
            If trim(mf_rept_status) <> "SIX MONTH" Then
                If DateDiff("d", #7/1/2025#, version_date) < 0 Then need_review = True
            End If
            PF3
        End If
        If CASES_ARRAY(GA_case_const, duck) Then
            Call write_value_and_transmit("GA  ", 20, 71)
            Call find_last_approved_ELIG_version(20, 78, version_number, version_date, version_result, approval_found)
            CASES_ARRAY(GA_approval_date_const, duck) = version_date
            Call write_value_and_transmit("GASM", 20, 70)
            EMReadScreen ga_rept_status, 9, 8, 32
            CASES_ARRAY(GA_rept_status_const, duck) = ga_rept_status
            If trim(ga_rept_status) = "MONTHLY" Then need_review = True
            If trim(ga_rept_status) <> "SIX MONTH" Then
                If DateDiff("d", #7/1/2025#, version_date) < 0 Then need_review = True
            End If
            PF3
        End If
        If CASES_ARRAY(FS_case_const, duck) Then
            Call write_value_and_transmit("FS  ", 20, 71)
            Call find_last_approved_ELIG_version(19, 78, version_number, version_date, version_result, approval_found)
            CASES_ARRAY(FS_approval_date_const, duck) =  version_date
            Call write_value_and_transmit("FSSM", 19, 70)
            EMReadScreen fs_rept_status, 9, 8, 31
            CASES_ARRAY(FS_rept_status_const, duck) = fs_rept_status
            If trim(fs_rept_status) = "MONTHLY" Then need_review = True
            ' If trim(fs_rept_status) <> "SIX MONTH" Then
            '     If DateDiff("d", #7/1/2025#, version_date) < 0 Then need_review = True
            ' End If
            PF3
        End If
        ' If all_rept_status_are_sr Then
        CASES_ARRAY(review_case_const, duck) = need_review
    End If
    Call back_to_SELF
Next

objExcel.worksheets("Worklist").Activate
excel_row = 2
For duck = 0 to UBound(CASES_ARRAY, 2)
    ObjExcel.Cells(excel_row, wrklst_case_numb_col).Value = CASES_ARRAY(case_numb_const, duck)
    ObjExcel.Cells(excel_row, wrklst_case_name_col).Value = CASES_ARRAY(case_name_const, duck)
    ObjExcel.Cells(excel_row, wrklst_next_revw_col).Value = CASES_ARRAY(next_revw_const, duck)
    ObjExcel.Cells(excel_row, wrklst_dail_col).Value = CASES_ARRAY(dail_exists_const, duck)

    If CASES_ARRAY(FS_case_const, duck) Then
        ObjExcel.Cells(excel_row, wrklst_snap_col).Value = "True"
        ObjExcel.Cells(excel_row, wrklst_snap_app_date_col).Value = CASES_ARRAY(FS_approval_date_const, duck)
        ObjExcel.Cells(excel_row, wrklst_snap_rept_status_col).Value = CASES_ARRAY(FS_rept_status_const, duck)
    End If

    If CASES_ARRAY(MF_case_const, duck) Then
        ObjExcel.Cells(excel_row, wrklst_mfip_col).Value = "True"
        ObjExcel.Cells(excel_row, wrklst_mfip_app_date_col).Value = CASES_ARRAY(MF_approval_date_const, duck)
        ObjExcel.Cells(excel_row, wrklst_mfip_rept_status_col).Value = CASES_ARRAY(MF_rept_status_const, duck)
    End If

    If CASES_ARRAY(GA_case_const, duck) Then
        ObjExcel.Cells(excel_row, wrklst_ga_col).Value = "True"
        ObjExcel.Cells(excel_row, wrklst_ga_app_date_col).Value = CASES_ARRAY(GA_approval_date_const, duck)
        ObjExcel.Cells(excel_row, wrklst_ga_rept_status_col).Value = CASES_ARRAY(GA_rept_status_const, duck)
    End If
    ObjExcel.Cells(excel_row, wrklst_review_case_col).Value = CASES_ARRAY(review_case_const, duck)

    excel_row = excel_row + 1
Next

end_msg = "Report Completed" & vbCr & vbCr & "Elapsed time - " & (timer - start_time) & " seconds."
Call script_end_procedure(end_msg)
