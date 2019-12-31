'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "BULK - PAPERLESS Review.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = "60"                'manual run time in seconds
STATS_denomination = "C"       			'C is for each CASE

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
call changelog_update("12/17/2018", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DECLARATIONS===============================================================================================================

'CONSTANTS
const case_nrb      = 0
const basket_nbr    = 1
const clt_name      = 2
const memb_on_hc    = 3
const revw_type     = 4
const hc_sr_date    = 5
const hc_er_date    = 6

const cash_revw     = 7
const SNAP_revw     = 8

const ca_sr_date    = 9
const ca_er_date    = 10
const fs_sr_date    = 11
const fs_er_date    = 12

const time_between  = 13
const hc_type       = 14
const cash_status   = 15
const SNAP_status   = 16
const waived_revw   = 17
const case_notes    = 18

'ARRAY
Dim ALL_HC_REVS_ARRAY()
ReDim ALL_HC_REVS_ARRAY(case_notes, 0)

'===========================================================================================================================
'THE SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""

'establishing variable for the script since most users are approving CM + 1
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

EMReadScreen on_revw_panel, 4, 2, 52
If on_revw_panel = "REVW" Then
    EMReadScreen basket_number, 7, 21, 6
    worker_number = trim(basket_number)
End If

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 241, 130, "PAPERLESS IR"
  EditBox 70, 10, 50, 15, worker_number
  EditBox 190, 10, 15, 15, MAXIS_footer_month
  EditBox 210, 10, 15, 15, MAXIS_footer_year
  CheckBox 10, 30, 145, 10, "Check here to run for the entire agency", whole_county_check
  ButtonGroup ButtonPressed
    OkButton 120, 110, 50, 15
    CancelButton 175, 110, 50, 15
  Text 5, 15, 60, 10, "Worker number(s):"
  Text 125, 15, 65, 10, "Footer month/year:"
  GroupBox 5, 50, 220, 55, "About the Paperless IR script:"
  Text 10, 65, 205, 35, "This script will update REVW for each starred IR, after checking JOBS/BUSI/RBIC for discrepancies. It skips cases that are also reviewing for SNAP. You will have to manually check ELIG/HC for each case and approve the results/case note."
EndDialog
DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
		If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) > 2 or len(MAXIS_footer_month) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer month."
		If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) > 2 or len(MAXIS_footer_year) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer year."
		'If trim(worker_number) <> "" AND Len(worker_number) <> 7 then err_msg = err_msg & vbNewLine & "* You must enter a valid 7 DIGIT worker number."
        If trim(worker_number) = "" AND whole_county_check = unchecked then err_msg = err_msg & vbNewLine & "* You must either list a 7 DIGIT worker number OR indicate the script should be run for the entire county."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

If whole_county_check = checked Then
    all_case_numbers_array = " "					'Creating blank variable for the future array
    get_county_code	'Determines worker county code

    call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
    worker_array = Array()
    ' MsgBox worker_number
    If len(worker_number) = 7 Then
        worker_array = Array(worker_number)
    Else
        worker_array = split(worker_number, ",")
    End If
End If

hc_reviews = 0

For each worker in worker_array
    worker = trim(worker)

    Call back_to_SELF
    EMWriteScreen "        ", 18, 43
    Call MAXIS_footer_month_confirmation
    Call navigate_to_MAXIS_screen("rept", "revs")
    EMWriteScreen worker, 21, 6
    EMWriteScreen CM_plus_2_mo, 20, 55
    EMWriteScreen CM_plus_2_yr, 20, 58
    transmit

    EMReadScreen REVW_check, 4, 2, 52
    If REVW_check <> "REVS" then script_end_procedure("You must start this script at the beginning of REPT/REVS. Navigate to the screen and try again!")

    row = 7
    Do
        MAXIS_check = ""
        last_page_check = ""

        EMReadScreen the_case_number, 8, row, 6
        EMReadScreen cash_review, 1, row, 39
        EMReadScreen snap_review, 1, row, 45
        EMReadScreen hc_review, 1, row, 49
        EMReadScreen paperless_check, 1, row, 51
        EMReadScreen hc_pop, 4, row, 55

        If hc_review = "N" Then

            ReDim Preserve ALL_HC_REVS_ARRAY(case_notes, hc_reviews)
            If trim(the_case_number) = "" Then MsgBox "Row: " & row & vbNewLine & "review code: " & hc_review
            ALL_HC_REVS_ARRAY(case_nrb, hc_reviews)     = trim(the_case_number)
            ALL_HC_REVS_ARRAY (basket_nbr, hc_reviews)   = worker
            If paperless_check = "*" Then
                ALL_HC_REVS_ARRAY(waived_revw, hc_reviews) = TRUE
            Else
                ALL_HC_REVS_ARRAY(waived_revw, hc_reviews) = FALSE
            End If
            ALL_HC_REVS_ARRAY (hc_type, hc_reviews)      = trim(hc_pop)
            If cash_review = "N" Then ALL_HC_REVS_ARRAY (cash_revw, hc_reviews) = TRUE
            If SNAP_review = "N" Then ALL_HC_REVS_ARRAY (SNAP_revw, hc_reviews) = TRUE

            hc_reviews = hc_reviews + 1
        End If

        row = row + 1

        If row = 19 then
            PF8
            row = 7
            EMReadScreen MAXIS_check, 5, 1, 39
            If MAXIS_check <> "MAXIS" then stopscript
            EMReadScreen last_page_check, 4, 24, 14
        End if
    Loop until last_page_check = "LAST" or trim(the_case_number) = ""

    pf3

Next

Call back_to_SELF

For hc_reviews = 0 to UBound(ALL_HC_REVS_ARRAY, 2)
    MAXIS_case_number = ALL_HC_REVS_ARRAY (case_nrb, hc_reviews)

    Call navigate_to_MAXIS_screen("CASE", "CURR")
    EmReadscreen priv_check, 4, 24, 46
    If priv_check <> "PRIV" Then
        Call navigate_to_MAXIS_screen ("STAT", "MEMB")
        EMReadScreen last_name, 25, 6, 30
        EMReadScreen first_name, 12, 6, 63
        last_name = replace(last_name, "_", "")
        first_name = replace(first_name, "_", "")

        ALL_HC_REVS_ARRAY (clt_name, hc_reviews) = last_name & ", " & first_name

        Call navigate_to_MAXIS_screen ("STAT", "PROG")

        EMReadScreen Cash1_code, 4, 6, 74
        EMReadScreen Cash2_code, 4, 7, 74
        EMReadScreen SNAP_code, 4, 10, 74

        If Cash1_code = "ACTV" Then ALL_HC_REVS_ARRAY (cash_status, hc_reviews)  = "Active"
        If Cash2_code = "ACTV" Then ALL_HC_REVS_ARRAY (cash_status, hc_reviews)  = "Active"
        If Cash1_code = "PEND" Then ALL_HC_REVS_ARRAY (cash_status, hc_reviews)  = "Pending"
        If Cash2_code = "PEND" Then ALL_HC_REVS_ARRAY (cash_status, hc_reviews)  = "Pending"
        If ALL_HC_REVS_ARRAY (cash_status, hc_reviews) = "" Then ALL_HC_REVS_ARRAY (cash_status, hc_reviews) = "Inactive"

        If SNAP_code = "ACTV" Then ALL_HC_REVS_ARRAY (SNAP_status, hc_reviews)  = "Active"
        If SNAP_code = "PEND" Then ALL_HC_REVS_ARRAY (SNAP_status, hc_reviews)  = "Pending"
        If ALL_HC_REVS_ARRAY (SNAP_status, hc_reviews) = "" Then ALL_HC_REVS_ARRAY (SNAP_status, hc_reviews) = "Inactive"

        Call navigate_to_MAXIS_screen("STAT", "REVW")

        EMReadScreen revw_cycle_hc, 2, 9, 79
        ALL_HC_REVS_ARRAY (revw_type, hc_reviews) = revw_cycle_hc

        EMWriteScreen "X", 5, 71
        transmit

        EMReadScreen ALL_HC_REVS_ARRAY(hc_er_date, hc_reviews), 8, 9, 27
        ALL_HC_REVS_ARRAY (hc_er_date, hc_reviews) = replace(ALL_HC_REVS_ARRAY (hc_er_date, hc_reviews), " ", "/")

        EMReadScreen inc_renw_date, 8, 8, 27
        EMReadScreen ast_renw_date, 8, 8, 71
        If inc_renw_date <> "__ 01 __" Then ALL_HC_REVS_ARRAY (hc_sr_date, hc_reviews) = replace(inc_renw_date, " ", "/")
        If ast_renw_date <> "__ 01 __" Then ALL_HC_REVS_ARRAY (hc_sr_date, hc_reviews) = replace(ast_renw_date, " ", "/")
        PF3

        If ALL_HC_REVS_ARRAY (cash_status, hc_reviews) = "Active" Then

            EMReadScreen ALL_HC_REVS_ARRAY (ca_er_date, hc_reviews), 8, 9, 37

            ALL_HC_REVS_ARRAY (ca_er_date, hc_reviews) = replace(ALL_HC_REVS_ARRAY (ca_er_date, hc_reviews), " ", "/")

        End If

        If ALL_HC_REVS_ARRAY (SNAP_status, hc_reviews) = "Active" Then
            EMWriteScreen "X", 5, 58
            transmit
            EMReadScreen ALL_HC_REVS_ARRAY (fs_sr_date, hc_reviews), 8, 9, 26
            EMReadScreen ALL_HC_REVS_ARRAY (fs_er_date, hc_reviews), 8, 9, 64

            ALL_HC_REVS_ARRAY (fs_sr_date, hc_reviews) = replace(ALL_HC_REVS_ARRAY (fs_sr_date, hc_reviews), " ", "/")
            ALL_HC_REVS_ARRAY (fs_er_date, hc_reviews) = replace(ALL_HC_REVS_ARRAY (fs_er_date, hc_reviews), " ", "/")

            If ALL_HC_REVS_ARRAY (fs_sr_date, hc_reviews) = "__/01/__" Then ALL_HC_REVS_ARRAY (fs_sr_date, hc_reviews) = ""
            PF3
        End If
    Else
        ALL_HC_REVS_ARRAY (revw_type, hc_reviews) = "PRIV"
    End If
Next

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Name for the current sheet'
ObjExcel.ActiveSheet.Name = "HC Reviews"

col_to_use = 1

'Excel headers and formatting the columns
objExcel.Cells(1, col_to_use).Value  = "BASKET"
basket_col = col_to_use
col_to_use = col_to_use + 1

objExcel.Cells(1, col_to_use).Value  = "CASE NUMBER"
case_number_col = col_to_use
col_to_use = col_to_use + 1

objExcel.Cells(1, col_to_use).Value  = "CLIENT NAME"
client_name_col = col_to_use
col_to_use = col_to_use + 1

'objExcel.Cells(1, 4).Value  = "MEMBS ON HC"
objExcel.Cells(1, col_to_use).Value  = "MAGI HC"
magi_col = col_to_use
col_to_use = col_to_use + 1

objExcel.Cells(1, col_to_use).Value  = "Current HC REVW"
current_revw_col = col_to_use
current_revw_letter_col = convert_digit_to_excel_column(current_revw_col)
col_to_use = col_to_use + 1

objExcel.Cells(1, col_to_use).Value  = "Paperless IR"
paperless_col = col_to_use
paperless_letter_col = convert_digit_to_excel_column(paperless_col)
col_to_use = col_to_use + 1

objExcel.Cells(1, col_to_use).Value  = "HC SR"
hc_sr_col = col_to_use
col_to_use = col_to_use + 1

objExcel.Cells(1, col_to_use).Value  = "HC ER"
hc_er_col = col_to_use
col_to_use = col_to_use + 1

objExcel.Cells(1, col_to_use).Value  = "Cash Status"
cash_col = col_to_use
col_to_use = col_to_use + 1

objExcel.Cells(1, col_to_use).Value = "Cash ER"
cash_er_col = col_to_use
col_to_use = col_to_use + 1

objExcel.Cells(1, col_to_use).Value = "SNAP Status"
snap_col = col_to_use
col_to_use = col_to_use + 1

objExcel.Cells(1, col_to_use).Value = "SNAP SR"
snap_sr_col = col_to_use
col_to_use = col_to_use + 1

objExcel.Cells(1, col_to_use).Value = "SNAP ER"
snap_er_col = col_to_use
col_to_use = col_to_use + 1

For i = 1 to col_to_use
    ObjExcel.Cells(1, i).Font.Bold = TRUE
Next

excel_row = 2
For hc_reviews = 0 to UBound(ALL_HC_REVS_ARRAY, 2)

    ObjExcel.Cells(excel_row, basket_col).Value         = ALL_HC_REVS_ARRAY(basket_nbr, hc_reviews)
    ObjExcel.Cells(excel_row, case_number_col).Value    = ALL_HC_REVS_ARRAY(case_nrb, hc_reviews)
    ObjExcel.Cells(excel_row, client_name_col).Value    = ALL_HC_REVS_ARRAY(clt_name, hc_reviews)
    ' ObjExcel.Cells(excel_row, 4).Value = ALL_HC_REVS_ARRAY(memb_on_hc, hc_reviews)
    ObjExcel.Cells(excel_row, magi_col).Value           = ALL_HC_REVS_ARRAY(hc_type, hc_reviews)
    ObjExcel.Cells(excel_row, current_revw_col).Value   = ALL_HC_REVS_ARRAY(revw_type, hc_reviews)
    ObjExcel.Cells(excel_row, paperless_col).Value      = ALL_HC_REVS_ARRAY(waived_revw, hc_reviews)
    ObjExcel.Cells(excel_row, hc_sr_col).Value          = ALL_HC_REVS_ARRAY(hc_sr_date, hc_reviews)
    ObjExcel.Cells(excel_row, hc_er_col).Value          = ALL_HC_REVS_ARRAY(hc_er_date, hc_reviews)
    ObjExcel.Cells(excel_row, cash_col).Value           = ALL_HC_REVS_ARRAY(cash_status, hc_reviews)
    ObjExcel.Cells(excel_row, cash_er_col).Value        = ALL_HC_REVS_ARRAY(ca_er_date, hc_reviews)
    ObjExcel.Cells(excel_row, snap_col).Value           = ALL_HC_REVS_ARRAY(SNAP_status, hc_reviews)
    ObjExcel.Cells(excel_row, snap_sr_col).Value        = ALL_HC_REVS_ARRAY(fs_sr_date, hc_reviews)
    ObjExcel.Cells(excel_row, snap_er_col).Value        = ALL_HC_REVS_ARRAY(fs_er_date, hc_reviews)

    ' const case_nrb      = 0
    ' const basket_nbr    = 1
    ' const clt_name      = 2
    ' const memb_on_hc    = 3
    ' const revw_type     = 4
    ' const hc_sr_date    = 5
    ' const hc_er_date    = 6
    '
    ' const ca_sr_date    = 7
    ' const ca_er_date    = 8
    ' const fs_sr_date    = 9
    ' const fs_er_date    = 10
    '
    ' const time_between  = 11
    ' const hc_type       = 12
    ' const cash_status   = 13
    ' const SNAP_status   = 14
    ' const waived_revw   = 15
    ' const case_notes    = 16
    excel_row = excel_row + 1
Next

col_to_use = col_to_use + 2	'Doing two because the wrap-up is two columns
letter_col_to_use = convert_digit_to_excel_column(col_to_use)

'Query date/time/runtime info
ObjExcel.Cells(1, col_to_use - 1).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, col_to_use).Value = now
ObjExcel.Cells(2, col_to_use - 1).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, col_to_use).Value = timer - query_start_time
ObjExcel.Cells(3, col_to_use - 1). Value = "Number of cases with HC ER for " & CM_plus_2_mo & "/" & CM_plus_2_yr
ObjExcel.Cells(3, col_to_use). Value = "=COUNTIF(" & current_revw_letter_col & ":" & current_revw_letter_col & ", " & chr(34) & "ER" & chr(34) & ")"
ObjExcel.Cells(4, col_to_use - 1). Value = "Number of cases with HC IR for " & CM_plus_2_mo & "/" & CM_plus_2_yr
ObjExcel.Cells(4, col_to_use). Value = "=COUNTIF(" & current_revw_letter_col & ":" & current_revw_letter_col & ", " & chr(34) & "IR" & chr(34) & ")"
ObjExcel.Cells(5, col_to_use - 1). Value = "Number of cases with HC AR for " & CM_plus_2_mo & "/" & CM_plus_2_yr
ObjExcel.Cells(5, col_to_use). Value = "=COUNTIF(" & current_revw_letter_col & ":" & current_revw_letter_col & ", " & chr(34) & "AR" & chr(34) & ")"
ObjExcel.Cells(6, col_to_use - 1). Value = "Total number of cases with either HC SR for " & CM_plus_2_mo & "/" & CM_plus_2_yr
ObjExcel.Cells(6, col_to_use). Value = "=" & letter_col_to_use & "4" & "+" & letter_col_to_use & "5"
ObjExcel.Cells(7, col_to_use - 1). Value = "Number of cases with Waived HC SR fpr " & CM_plus_2_mo & "/" & CM_plus_2_yr
ObjExcel.Cells(7, col_to_use). Value = "=COUNTIF(" & paperless_letter_col & ":" & paperless_letter_col & ", " & chr(34) & "TRUE" & chr(34) & ")"

For i = 1 to 7
    ObjExcel.Cells(i, col_to_use - 1).Font.Bold = TRUE
Next
'Autofitting columns
For col_to_autofit = 1 to col_to_use
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

script_end_procedure("Check List")
