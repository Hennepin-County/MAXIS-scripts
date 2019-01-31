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
call changelog_update("05/14/2018", "Updated the TIKL functionality to write TIKL for the current day of the month.", "Ilse Ferris, Hennepin County")
call changelog_update("12/05/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DECLARATIONS----------------------------------------------------------------------------------------------
'Variables
'establishing variable for the script since most users are approving CM + 1
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

'defaulting the cleared date in REVW to the frst of the current month
current_month = CM_mo
current_day = "01"
current_year = CM_yr

'Constants
const case_nrb      = 0
const basket_nbr    = 1
const clt_name      = 2
const memb_on_hc    = 3
const revw_type     = 4
const hc_sr_date    = 5
const hc_er_date    = 6

const new_hc_sr     = 7
const new_hc_er     = 8

const cash_revw     = 9
const SNAP_revw     = 10

const ca_sr_date    = 11
const ca_er_date    = 12
const fs_sr_date    = 13
const fs_er_date    = 14

const new_ca_er     = 15
const new_fs_sr     = 16
const new_fs_er     = 17

const time_between  = 18
const hc_type       = 19
const cash_status   = 20
const SNAP_status   = 21
const waived_revw   = 22
const actually_paperless = 23
const membs_updated = 24
const current_budg  = 25
const tikl_done     = 26
const correct_list  = 27
const revw_updated  = 28
const case_notes    = 29

'Updated Waived IRs
'Other current reviews
'Other reviews are off
'Not actually paperless
'No MEMBS with N REVW'
'ObjExcel.worksheets("").Activate

'ARRAY
Dim ALL_HC_REVS_ARRAY()
ReDim ALL_HC_REVS_ARRAY(case_notes, 0)

'----------------------------------------------------------------------------------------------------------
'DIALOG----------------------------------------------------------------------------------------------------
BeginDialog paperless_IR_dialog, 0, 0, 241, 130, "PAPERLESS IR"
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

'THE SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""

'THIS IS THE MAPPING FOR A REWRITE OF PAPERLESS IR
'THIS SCRIPT IS NOT LIVE AND IS NOT TESTED
developer_mode = FALSE

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

EMReadScreen on_revw_panel, 4, 2, 52
If on_revw_panel = "REVW" Then
    EMReadScreen basket_number, 7, 21, 6
    worker_number = trim(basket_number)
End If

DO
	DO
		err_msg = ""
		Dialog paperless_IR_dialog
		If buttonpressed = 0 then stopscript
		If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) > 2 or len(MAXIS_footer_month) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer month."
		If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) > 2 or len(MAXIS_footer_year) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer year."
		'If trim(worker_number) <> "" AND Len(worker_number) <> 7 then err_msg = err_msg & vbNewLine & "* You must enter a valid 7 DIGIT worker number."
        If trim(worker_number) = "" AND whole_county_check = unchecked then err_msg = err_msg & vbNewLine & "* You must either list a 7 DIGIT worker number OR indicate the script should be run for the entire county."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

worker_number = trim(worker_number)

If left(worker_number, 10) = "UUDDLRLRBA" Then
    developer_mode = TRUE
    worker_number = right(worker_number, len(worker_number)-10)
    worker_number = trim(worker_number)
    MsgBox "Congratulations! You are now in DEVELOPER MODE. Have Fun!"
End If
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
    EMWriteScreen MAXIS_footer_month, 20, 55
    EMWriteScreen MAXIS_footer_year, 20, 58
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

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Name for the current sheet'
ObjExcel.ActiveSheet.Name = "Updated Waived IRs"
on_loop = 1

Do
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

    objExcel.Cells(1, col_to_use).Value  = "MEMBS ON HC"
    membs_col = col_to_use
    col_to_use = col_to_use + 1

    objExcel.Cells(1, col_to_use).Value  = "MAGI HC"
    magi_col = col_to_use
    col_to_use = col_to_use + 1

    objExcel.Cells(1, col_to_use).Value  = "Current HC REVW"
    current_revw_col = col_to_use
    current_revw_letter_col = convert_digit_to_excel_column(current_revw_col)
    col_to_use = col_to_use + 1

    ' objExcel.Cells(1, col_to_use).Value  = "Paperless IR"
    ' paperless_col = col_to_use
    ' paperless_letter_col = convert_digit_to_excel_column(paperless_col)
    ' col_to_use = col_to_use + 1

    objExcel.Cells(1, col_to_use).Value  = "HC SR"
    hc_sr_col = col_to_use
    col_to_use = col_to_use + 1

    objExcel.Cells(1, col_to_use).Value  = "HC ER"
    hc_er_col = col_to_use
    col_to_use = col_to_use + 1

    ' objExcel.Cells(1, col_to_use).Value  = "NEW HC SR"
    ' new_hc_sr_col = col_to_use
    ' col_to_use = col_to_use + 1
    '
    ' objExcel.Cells(1, col_to_use).Value  = "NEW HC ER"
    ' new_hc_er_col = col_to_use
    ' col_to_use = col_to_use + 1

    objExcel.Cells(1, col_to_use).Value  = "Cash Status"
    cash_col = col_to_use
    col_to_use = col_to_use + 1

    objExcel.Cells(1, col_to_use).Value = "Cash ER"
    cash_er_col = col_to_use
    col_to_use = col_to_use + 1

    ' objExcel.Cells(1, col_to_use).Value = "NEW Cash ER"
    ' new_cash_er_col = col_to_use
    ' col_to_use = col_to_use + 1

    objExcel.Cells(1, col_to_use).Value = "SNAP Status"
    snap_col = col_to_use
    col_to_use = col_to_use + 1

    objExcel.Cells(1, col_to_use).Value = "SNAP SR"
    snap_sr_col = col_to_use
    col_to_use = col_to_use + 1

    objExcel.Cells(1, col_to_use).Value = "SNAP ER"
    snap_er_col = col_to_use
    col_to_use = col_to_use + 1

    ' objExcel.Cells(1, col_to_use).Value = "NEW SNAP SR"
    ' new_snap_sr_col = col_to_use
    ' col_to_use = col_to_use + 1

    objExcel.Cells(1, col_to_use).Value = "HH Updt on REVW"
    updates_col = col_to_use
    col_to_use = col_to_use + 1

    objExcel.Cells(1, col_to_use).Value = "Budget"
    budg_col = col_to_use
    col_to_use = col_to_use + 1

    objExcel.Cells(1, col_to_use).Value = "TIKL"
    tikl_col = col_to_use
    col_to_use = col_to_use + 1

    objExcel.Cells(1, col_to_use).Value = "NOTES"
    notes_col = col_to_use
    col_to_use = col_to_use + 1

    For i = 1 to col_to_use
        ObjExcel.Cells(1, i).Font.Bold = TRUE
    Next

    on_loop = on_loop + 1
    If on_loop = 2 Then ObjExcel.Worksheets.Add().Name = "Other current reviews"
    If on_loop = 3 Then ObjExcel.Worksheets.Add().Name = "Other reviews are off"
    If on_loop = 4 Then ObjExcel.Worksheets.Add().Name = "Not actually paperless"
    If on_loop = 5 Then ObjExcel.Worksheets.Add().Name = "No MEMBS with N REVW"

Loop until on_loop = 6
last_col = col_to_use

not_waived_excel_row = 2
curr_revw_excel_row = 2
othr_revw_excel_row = 2
paperless_excel_row = 2
not_updated_excel_row = 2

Call back_to_SELF

For hc_reviews = 0 to UBound(ALL_HC_REVS_ARRAY, 2)
    MAXIS_case_number = ALL_HC_REVS_ARRAY (case_nrb, hc_reviews)
    list_of_hc_membs = ""
    HC_PERS_ARRAY = ""

    If ALL_HC_REVS_ARRAY(waived_revw, hc_reviews) = TRUE Then
        Call navigate_to_MAXIS_screen("CASE", "CURR")
        EmReadscreen priv_check, 4, 24, 46
        If priv_check <> "PRIV" Then
            ALL_HC_REVS_ARRAY (actually_paperless, hc_reviews) = TRUE
            'figure out who is on HC

            Call navigate_to_MAXIS_screen("CASE", "PERS")
            pers_row = 10
            last_page_check = ""
            Do
                EMReadScreen pers_memb_numb, 2, pers_row, 3
                EMReadScreen pers_hc_status, 1, pers_row, 61

                If pers_memb_numb = "  " Then Exit Do
                If pers_hc_status = "A" Then list_of_hc_membs = list_of_hc_membs & "~" & pers_memb_numb

                pers_row = pers_row + 3
                If pers_row = 19 Then
                    pers_row = 10
                    PF8
                    EMReadScreen last_page_check, 9, 24, 14
                End If
            Loop until last_page_check = "LAST PAGE"
            list_of_hc_membs = right(list_of_hc_membs, len(list_of_hc_membs) - 1)
            HC_PERS_ARRAY = split(list_of_hc_membs, "~")

            For each pers_nbr in HC_PERS_ARRAY
                Do
                    Call navigate_to_MAXIS_screen("STAT", "SUMM")
                    EMReadScreen summ_check, 4, 2, 46
                Loop until summ_check = "SUMM"

                call navigate_to_MAXIS_screen ("STAT", "JOBS")
                EMWriteScreen pers_nbr, 20, 76
                transmit
                Do
                    EMReadScreen panel_check, 8, 2, 72
                    current_panel = trim(left(panel_check, 2))
                    total_panels = trim(right(panel_check, 2))
                    EMReadScreen date_check, 8, 9, 49
                    If total_panels <> "0" & date_check = "__ __ __" then
                        ALL_HC_REVS_ARRAY (actually_paperless, hc_reviews) = False
                        ALL_HC_REVS_ARRAY (case_notes, hc_reviews) = ALL_HC_REVS_ARRAY (case_notes, hc_reviews) & "; MEMB " & pers_nbr & " has an open JOBS Panel."
                    End If
                    if current_panel <> total_panels then transmit
                Loop until current_panel = total_panels

                call navigate_to_MAXIS_screen ("STAT", "BUSI")
                EMWriteScreen pers_nbr, 20, 76
                transmit
                Do
                    current_panel = trim(left(panel_check, 2))
                    EMReadScreen panel_check, 8, 2, 72
                    total_panels = trim(right(panel_check, 2))
                    EMReadScreen date_check, 8, 5, 71
                    If total_panels <> "0" & date_check = "__ __ __" then
                        ALL_HC_REVS_ARRAY (actually_paperless, hc_reviews) = False
                        ALL_HC_REVS_ARRAY (case_notes, hc_reviews) = ALL_HC_REVS_ARRAY (case_notes, hc_reviews) & "; MEMB " & pers_nbr & " has an open BUSI Panel."
                    End If
                    if current_panel <> total_panels then transmit
                Loop until current_panel = total_panels

                call navigate_to_MAXIS_screen ("STAT", "RBIC")
                EMWriteScreen pers_nbr, 20, 76
                transmit
                Do
                    EMReadScreen panel_check, 8, 2, 72
                    current_panel = trim(left(panel_check, 2))
                    total_panels = trim(right(panel_check, 2))
                    EMReadScreen date_check, 8, 6, 68
                    If total_panels <> "0" & date_check = "__ __ __" then
                        ALL_HC_REVS_ARRAY (actually_paperless, hc_reviews) = False
                        ALL_HC_REVS_ARRAY (case_notes, hc_reviews) = ALL_HC_REVS_ARRAY (case_notes, hc_reviews) & "; MEMB " & pers_nbr & " has an open JOBS Panel."
                    End If
                    if current_panel <> total_panels then transmit
                Loop until current_panel = total_panels
            Next


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

            hc_sr = ""
            hc_er = ""
            snap_sr = ""
            snap_er = ""
            cash_er = ""
            ALL_HC_REVS_ARRAY(revw_updated, hc_reviews) = FALSE
            If ALL_HC_REVS_ARRAY (actually_paperless, hc_reviews) = TRUE Then

                ALL_HC_REVS_ARRAY(correct_list, hc_reviews) = "Updated Waived IRs"

                EMReadScreen fs_revw_code, 1, 7, 60
                EMReadScreen cash_revw_code, 1, 7, 40

                other_review = FALSE
                If fs_revw_code <> "_" Then other_review = TRUE
                If cash_revw_code <> "_" Then other_review = TRUE

                If ALL_HC_REVS_ARRAY (fs_er_date, hc_reviews) <> "" Then                                                    'If there is a SNAP ER
                    If ALL_HC_REVS_ARRAY (hc_er_date, hc_reviews) <> ALL_HC_REVS_ARRAY (fs_er_date, hc_reviews) THen        'If the SNAP ER doesn't match the HC ER
                        If DateDiff("m", date, ALL_HC_REVS_ARRAY (fs_er_date, hc_reviews)) < 12 then
                            hc_er = ALL_HC_REVS_ARRAY (fs_er_date, hc_reviews)
                        Else
                            Do
                                hc_er = DateAdd("m", -12, ALL_HC_REVS_ARRAY (fs_er_date, hc_reviews))
                            Loop Until DateDiff("m", date, hc_er) < 12
                        End If
                        ALL_HC_REVS_ARRAY(new_hc_er, hc_reviews) = hc_er

                        If ALL_HC_REVS_ARRAY (ca_er_date, hc_reviews) <> "" Then
                            If hc_er <> ALL_HC_REVS_ARRAY (ca_er_date, hc_reviews) Then ALL_HC_REVS_ARRAY (new_ca_er, hc_reviews) = hc_er
                        End If
                    End If

                    If ALL_HC_REVS_ARRAY (fs_sr_date, hc_reviews) <> "" AND DateDiff("m", date, ALL_HC_REVS_ARRAY (fs_er_date, hc_reviews)) > 12 Then
                        ALL_HC_REVS_ARRAY(new_fs_sr, hc_reviews) = "__ __ __"
                    ElseIf ALL_HC_REVS_ARRAY (fs_sr_date, hc_reviews) <> "" Then
                        snap_sr_correct = FALSE
                        If DateDiff("m", ALL_HC_REVS_ARRAY (fs_sr_date, hc_reviews), ALL_HC_REVS_ARRAY (fs_er_date, hc_reviews)) <> 6 Then snap_sr_correct = TRUE
                        If DateDiff("m", ALL_HC_REVS_ARRAY (fs_sr_date, hc_reviews), ALL_HC_REVS_ARRAY (fs_er_date, hc_reviews)) <> -6 Then snap_sr_correct = TRUE


                        If snap_sr_correct = FALSE Then
                            If DateDiff("m", date, ALL_HC_REVS_ARRAY (fs_er_date, hc_reviews)) > 6 Then
                                 ALL_HC_REVS_ARRAY (new_fs_sr, hc_reviews) = DateAdd("m", -6, ALL_HC_REVS_ARRAY (fs_er_date, hc_reviews))
                            Else
                                 ALL_HC_REVS_ARRAY (new_fs_sr, hc_reviews) = DateAdd("m", 6, ALL_HC_REVS_ARRAY (fs_er_date, hc_reviews))
                            End If
                        End If

                    End If
                ElseIf ALL_HC_REVS_ARRAY (ca_er_date, hc_reviews) <> "" Then
                    If ALL_HC_REVS_ARRAY (hc_er_date, hc_reviews) <> ALL_HC_REVS_ARRAY (ca_er_date, hc_reviews) Then ALL_HC_REVS_ARRAY(new_hc_er, hc_reviews) = ALL_HC_REVS_ARRAY (ca_er_date, hc_reviews)
                End If

                If ALL_HC_REVS_ARRAY (fs_sr_date, hc_reviews) <> "" Then
                    snap_sr = ALL_HC_REVS_ARRAY (fs_sr_date, hc_reviews)
                    If ALL_HC_REVS_ARRAY(new_fs_sr, hc_reviews) <> "__ __ __" AND ALL_HC_REVS_ARRAY(new_fs_sr, hc_reviews) <> "" Then snap_sr = ALL_HC_REVS_ARRAY(new_fs_sr, hc_reviews)

                    If snap_sr <> ALL_HC_REVS_ARRAY (hc_sr_date, hc_reviews) Then ALL_HC_REVS_ARRAY (new_hc_sr, hc_reviews) = snap_sr
                Else
                    hc_sr_correct = FALSE
                    If DateDiff("m", ALL_HC_REVS_ARRAY (hc_sr_date, hc_reviews), ALL_HC_REVS_ARRAY (hc_er_date, hc_reviews)) <> 6 Then hc_sr_correct = TRUE
                    If DateDiff("m", ALL_HC_REVS_ARRAY (hc_sr_date, hc_reviews), ALL_HC_REVS_ARRAY (hc_er_date, hc_reviews)) <> -6 Then hc_sr_correct = TRUE

                    If hc_sr_correct = FALSE Then
                        If DateDiff("m", date, ALL_HC_REVS_ARRAY (hc_er_date, hc_reviews)) > 6 Then
                             ALL_HC_REVS_ARRAY (new_hc_sr, hc_reviews) = DateAdd("m", -6, ALL_HC_REVS_ARRAY (hc_er_date, hc_reviews))
                        Else
                             ALL_HC_REVS_ARRAY (new_hc_sr, hc_reviews) = DateAdd("m", 6, ALL_HC_REVS_ARRAY (hc_er_date, hc_reviews))
                        End If
                    End If
                End If

                'THIS IS THE OLD WAY
                If developer_mode = FALSE Then PF9
                EMWriteScreen "x", 5, 71                'Open HC Renewals Pop-Up
                transmit
                EMReadScreen renewal_year, 2, 8, 33
                If renewal_year = "__" then
                    EMReadScreen renewal_year, 2, 8, 77
                    renewal_year_col = 77
                Else
                    renewal_year_col = 33
                End if
                If developer_mode = FALSE Then
                    EMWriteScreen CM_mo, 6, 27
                    EMWriteScreen "01", 6, 30
                    EMWriteScreen CM_yr, 6, 33

                    new_renewal_year = cint(right(current_year, 2)) + 1
                    If current_month = 12 then new_renewal_year = new_renewal_year + 1 'Because otherwise the renewal year will be the current footer month.
                    EMWriteScreen new_renewal_year, 8, renewal_year_col
                End If

                revw_row = 13
                ALL_HC_REVS_ARRAY(revw_updated, hc_reviews) = FALSE
                Do
                    EMReadScreen renewal_status, 1, revw_row, 43
                    EMReadScreen ref_nbr, 2, revw_row, 6
                    If ref_nbr = "  " Then Exit Do

                    If renewal_status = "N" Then
                        ALL_HC_REVS_ARRAY(revw_updated, hc_reviews) = TRUE
                        If developer_mode = FALSE Then EMWriteScreen "U", revw_row, 43
                        ALL_HC_REVS_ARRAY(membs_updated, hc_reviews) = ALL_HC_REVS_ARRAY(membs_updated, hc_reviews) & ", " & ref_nbr
                    End If

                    revw_row = revw_row + 1
                    if revw_row = 21 Then
                        EMReadScreen first_ref_nbr, 2, 13, 6
                        PF20
                        EMReadScreen new_first, 2, 13, 6
                        if first_ref_nbr <> new_first Then revw_row = 13
                    End If

                Loop Until revw_row = 21


                If ALL_HC_REVS_ARRAY(revw_updated, hc_reviews) = FALSE Then
                    PF10    'undoing the date updates'
                    PF3     'leaving the HC Pop Up
                    PF10    'Leaving Edit Mode
                    ALL_HC_REVS_ARRAY(correct_list, hc_reviews) = "No MEMBS with N REVW"
                Else
                    PF3
                    EMReadScreen hc_revw_code, 1, 7, 73
                    If hc_revw_code = "N" Then
                        ALL_HC_REVS_ARRAY(correct_list, hc_reviews) = "No MEMBS with N REVW"
                        ALL_HC_REVS_ARRAY(revw_updated, hc_reviews) = FALSE
                        ALL_HC_REVS_ARRAY(membs_updated, hc_reviews) = ""
                        PF10
                    End If
                    transmit    'save and get out of pop up
                    EMReadScreen failure_check, 78, 24, 2
                    failure_check = trim(failure_check)

                    If failure_check = "FS REVIEW DATES MUST ALSO BE UPDATED" Then
                        ALL_HC_REVS_ARRAY(correct_list, hc_reviews) = "Other current reviews"
                        ALL_HC_REVS_ARRAY(revw_updated, hc_reviews) = FALSE
                        ALL_HC_REVS_ARRAY(membs_updated, hc_reviews) = ""
                        PF10
                    End If

                    If failure_check = "CASH/GRH REVIEW DATE MUST ALSO BE UPDATED" Then
                        ALL_HC_REVS_ARRAY(correct_list, hc_reviews) = "Other current reviews"
                        ALL_HC_REVS_ARRAY(revw_updated, hc_reviews) = FALSE
                        ALL_HC_REVS_ARRAY(membs_updated, hc_reviews) = ""
                        PF10
                    End If

                    If failure_check = "NEXT REVIEW DATE MUST BE AFTER THE CURRENT CALENDAR MONTH AND YEAR" Then
                        ALL_HC_REVS_ARRAY(correct_list, hc_reviews) = "Other reviews are off"
                        ALL_HC_REVS_ARRAY(revw_updated, hc_reviews) = FALSE
                        ALL_HC_REVS_ARRAY(membs_updated, hc_reviews) = ""
                        PF10
                    End If

                End If
                'THE NEW WAY
                'This should change all to an IR and use the new review dates as determined above.
                'Will also adjust budget.
                Call navigate_to_MAXIS_screen("STAT", "BUDG")

                EMReadScreen start_of_budg, 5, 10, 35
                EMReadScreen end_of_budg, 5, 10, 46

                start_of_budg = replace(start_of_budg, " ", "/")
                end_of_budg = replace(end_of_budg, " ", "/")

                ALL_HC_REVS_ARRAY(current_budg, hc_reviews) = start_of_budg & " - " & end_of_budg

            Else
                ALL_HC_REVS_ARRAY(correct_list, hc_reviews) = "Not actually paperless"
                ALL_HC_REVS_ARRAY(revw_updated, hc_reviews) = FALSE
            End IF

            ALL_HC_REVS_ARRAY(memb_on_hc, hc_reviews) = Join(HC_PERS_ARRAY, ", ")
            HC_PERS_ARRAY = ""
        Else
            ALL_HC_REVS_ARRAY (revw_type, hc_reviews) = "PRIV"
            ALL_HC_REVS_ARRAY(correct_list, hc_reviews) = "No MEMBS with N REVW"
        End If

        If developer_mode = FALSE Then
            If ALL_HC_REVS_ARRAY(revw_updated, hc_reviews) = TRUE Then
                navigate_to_MAXIS_screen "DAIL", "WRIT"
                call create_MAXIS_friendly_date(date, 0, 5, 18)
                EMWritescreen "%^% Sent through background for Paperless IR Review %^%", 9, 3
                transmit
                EMReadScreen tikl_success, 4, 24, 2
                ' MsgBox "Suc? - ''" & tikl_success & "'"
                If tikl_success <> "    " Then
                    ALL_HC_REVS_ARRAY(tikl_done, hc_reviews) = "Fail"
                    ' MsgBox "This case - " & MAXIS_case_number & " failed to have a TIKL set, track and case note manually"
                Else
                    ALL_HC_REVS_ARRAY(tikl_done, hc_reviews) = "Success"
                End If
                PF3
            Else
                ALL_HC_REVS_ARRAY(tikl_done, hc_reviews) = "REVW not Updated"
            End If
        End If

        objExcel.worksheets(ALL_HC_REVS_ARRAY(correct_list, hc_reviews)).Activate

        If ALL_HC_REVS_ARRAY(correct_list, hc_reviews) = "Updated Waived IRs" Then excel_row_to_use = not_waived_excel_row
        If ALL_HC_REVS_ARRAY(correct_list, hc_reviews) = "Other current reviews" Then excel_row_to_use = curr_revw_excel_row
        If ALL_HC_REVS_ARRAY(correct_list, hc_reviews) = "Other reviews are off" Then excel_row_to_use = othr_revw_excel_row
        If ALL_HC_REVS_ARRAY(correct_list, hc_reviews) = "Not actually paperless" Then excel_row_to_use = paperless_excel_row
        If ALL_HC_REVS_ARRAY(correct_list, hc_reviews) = "No MEMBS with N REVW" Then excel_row_to_use = not_updated_excel_row

        ObjExcel.Cells(excel_row_to_use, basket_col).Value         = ALL_HC_REVS_ARRAY(basket_nbr, hc_reviews)
        ObjExcel.Cells(excel_row_to_use, case_number_col).Value    = ALL_HC_REVS_ARRAY(case_nrb, hc_reviews)
        ObjExcel.Cells(excel_row_to_use, client_name_col).Value    = ALL_HC_REVS_ARRAY(clt_name, hc_reviews)
        ObjExcel.Cells(excel_row_to_use, membs_col).Value          = ALL_HC_REVS_ARRAY(memb_on_hc, hc_reviews)
        ObjExcel.Cells(excel_row_to_use, magi_col).Value           = ALL_HC_REVS_ARRAY(hc_type, hc_reviews)
        ObjExcel.Cells(excel_row_to_use, current_revw_col).Value   = ALL_HC_REVS_ARRAY(revw_type, hc_reviews)
        ' ObjExcel.Cells(excel_row_to_use, paperless_col).Value      = ALL_HC_REVS_ARRAY(waived_revw, hc_reviews)
        ObjExcel.Cells(excel_row_to_use, hc_sr_col).Value          = ALL_HC_REVS_ARRAY(hc_sr_date, hc_reviews)
        ObjExcel.Cells(excel_row_to_use, hc_er_col).Value          = ALL_HC_REVS_ARRAY(hc_er_date, hc_reviews)
        ' ObjExcel.Cells(excel_row_to_use, new_hc_er_col).Value      = ALL_HC_REVS_ARRAY(new_hc_er, hc_reviews)
        ' ObjExcel.Cells(excel_row_to_use, new_hc_sr_col).Value      = ALL_HC_REVS_ARRAY(new_hc_sr, hc_reviews)

        ObjExcel.Cells(excel_row_to_use, cash_col).Value           = ALL_HC_REVS_ARRAY(cash_status, hc_reviews)
        ObjExcel.Cells(excel_row_to_use, cash_er_col).Value        = ALL_HC_REVS_ARRAY(ca_er_date, hc_reviews)
        ' ObjExcel.Cells(excel_row_to_use, new_cash_er_col).Value    = ALL_HC_REVS_ARRAY(new_ca_er, hc_reviews)

        ObjExcel.Cells(excel_row_to_use, snap_col).Value           = ALL_HC_REVS_ARRAY(SNAP_status, hc_reviews)
        ObjExcel.Cells(excel_row_to_use, snap_sr_col).Value        = ALL_HC_REVS_ARRAY(fs_sr_date, hc_reviews)
        ObjExcel.Cells(excel_row_to_use, snap_er_col).Value        = ALL_HC_REVS_ARRAY(fs_er_date, hc_reviews)
        ' ObjExcel.Cells(excel_row_to_use, new_snap_sr_col).Value    = ALL_HC_REVS_ARRAY(new_fs_sr, hc_reviews)
        ' ObjExcel.Cells(excel_row_to_use, new_snap_er_col).Value    = ALL_HC_REVS_ARRAY(new_fs_er, hc_reviews)

        ObjExcel.Cells(excel_row_to_use, updates_col).Value        = ALL_HC_REVS_ARRAY(membs_updated, hc_reviews)
        ObjExcel.Cells(excel_row_to_use, budg_col).Value           = ALL_HC_REVS_ARRAY(current_budg, hc_reviews)
        ObjExcel.Cells(excel_row_to_use, tikl_col).Value           = ALL_HC_REVS_ARRAY(tikl_done, hc_reviews)
        ObjExcel.Cells(excel_row_to_use, notes_col).Value          = ALL_HC_REVS_ARRAY(case_notes, hc_reviews)

        excel_row_to_use = excel_row_to_use + 1


    End If
Next

on_loop = 1

Do
    If on_loop = 1 Then objExcel.worksheets("Other current reviews").Activate
    If on_loop = 2 Then objExcel.worksheets("Other reviews are off").Activate
    If on_loop = 3 Then objExcel.worksheets("Not actually paperless").Activate
    If on_loop = 4 Then objExcel.worksheets("No MEMBS with N REVW").Activate

    'Autofitting columns
    For col_to_autofit = 1 to last_col
    	ObjExcel.columns(col_to_autofit).AutoFit()
    Next

    on_loop = on_loop + 1
Loop until on_loop = 5


objExcel.worksheets("Updated Waived IRs").Activate

col_to_use = last_col + 2	'Doing two because the wrap-up is two columns
letter_col_to_use = convert_digit_to_excel_column(col_to_use)

'Query date/time/runtime info
ObjExcel.Cells(1, col_to_use - 1).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, col_to_use).Value = now
ObjExcel.Cells(2, col_to_use - 1).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, col_to_use).Value = timer - query_start_time


'Autofitting columns
For col_to_autofit = 1 to col_to_use
	ObjExcel.columns(col_to_autofit).AutoFit()
Next


script_end_procedure("Success! All starred (*) IRs have been sent into background, except those with current JOBS/BUSI/RBIC, those who have members other than 01 open, or those who also have SNAP up for review. You must go through and approve these results when they come through background.")
