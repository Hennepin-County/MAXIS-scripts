'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - EARNED INCOME BUDGETING.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 473                	'manual run time in seconds
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
call changelog_update("11/28/2016", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS==================================================================================================================

function sort_dates(dates_array)

    dim ordered_dates ()
    redim ordered_dates(0)

    days =  0
    do

        prev_date = ""
        for each thing in dates_array
            check_this_date = TRUE
            For each known_date in ordered_dates
                if known_date = thing Then check_this_date = FALSE
                'MsgBox "known dates is " & known_date & vbNewLine & "thing is " & thing & vbNewLine & "match - " & check_this_date
            next
            if check_this_date = TRUE Then
                if prev_date = "" Then
                    prev_date = thing
                Else
                    if DateDiff("d", prev_date, thing) <0 then
                        prev_date = thing
                    end if
                end if
            end if
        next
        if prev_date <> "" Then
            redim preserve ordered_dates(days)
            ordered_dates(days) = prev_date
            days = days + 1
        end if
    loop until days > UBOUND(dates_array)

    dates_array = ordered_dates
end function

'Declarations ==============================================================================================================
'CONSTANTS'
'Panel Constants
const panel_type        = 1
const panel_member      = 2
const panel_instance    = 3
const employer          = 4
const income_type       = 5
const income_verif      = 6
const hourly_wage       = 7
const income_start_dt   = 8
const income_end_dt     = 9
const income_list_indct = 10
const pay_freq          = 11
const date_of_calc      = 12
const hrs_per_wk        = 13
const pay_per_hr        = 14
const ave_hrs_per_pay   = 15
const ave_inc_per_pay   = 16
const SNAP_mo_inc       = 17
const reg_non_monthly   = 18
const numb_months       = 19
const self_emp_mthd     = 20
const method_date       = 21
const reptd_hours       = 22
const apply_to_SNAP     = 23
const apply_to_CASH     = 24
const apply_to_HC       = 25
const pay_weekday       = 26
const income_received   = 27
const verif_date        = 28
const verif_explain     = 29
const old_verif         = 30
const initial_month_mo  = 31
const initial_month_yr  = 32
const update_futue_chkbx = 33
const order_ubound      = 34
const self_emp_mthd_conv = 35
const cash_mos_list     = 36
const pick_one          = 37
const selection_rsn     = 38
const ignore_antic      = 39
const antic_pay_list    = 40
const update_this_month = 41
const last_paycheck     = 42
const panel_first_check = 43
const this_is_a_new_panel = 44
const days_of_verif     = 45

const spoke_to          = 46
const convo_detail      = 47

const use_actual        = 1
const use_estimate      = 2

'Income Constants
const panel_indct           = 0
const pay_date              = 1
const gross_amount          = 2
const hours                 = 3
const budget_in_SNAP_yes    = 4
const budget_in_SNAP_no     = 5
const reason_to_exclude     = 6
const exclude_amount        = 7
const check_order           = 8
const reason_amt_excluded   = 9


'Cash Months Constants
Const cash_mo_yr    = 1
const retro_mo_yr   = 2
Const retro_updtd   = 3
Const prosp_updtd   = 4
const mo_retro_pay  = 5
const mo_retro_hrs  = 6
const mo_prosp_pay  = 7
const mo_prosp_hrs  = 8

'ARRAYS'
Dim LIST_OF_INCOME_ARRAY()
ReDim LIST_OF_INCOME_ARRAY(reason_amt_excluded, 0)

Dim EARNED_INCOME_PANELS_ARRAY()
ReDim EARNED_INCOME_PANELS_ARRAY(convo_detail, 0)

Dim CASH_MONTHS_ARRAY()
ReDim CASH_MONTHS_ARRAY(8, 0)
'===========================================================================================================================


'THE SCRIPT ================================================================================================================
'Connecting to MAXIS, and grabbing the case number and footer month'
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

' 'TESTING ELEMENT REMOVAL'
' Dim TEST_ARRAY()
' ReDim TEST_ARRAY(0)
'
' Full = 6
' For the_thing = 0 to full
'     ReDim Preserve TEST_ARRAY(the_thing)
'     TEST_ARRAY(the_thing) = the_thing * the_thing
' Next
'
' For each square in TEST_ARRAY
'     MsgBox square
' Next
'
' ReDim Preserve TEST_ARRAY(5)
'
' For each square in TEST_ARRAY
'     MsgBox square
' Next
'
' MsgBox "That's it"

original_month = MAXIS_footer_month
originial_year = MAXIS_footer_year

'DIALOG TO GET CASE NUMBER
'Possibly add worker signature here and take it out of the following dialogs
BeginDialog Dialog1, 0, 0, 191, 220, "Case Number"
  EditBox 90, 5, 70, 15, MAXIS_case_number
  EditBox 100, 25, 15, 15, original_month
  EditBox 120, 25, 15, 15, originial_year
  CheckBox 10, 45, 140, 10, "Check here to have the script update all", future_months_check
  EditBox 5, 80, 175, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 80, 200, 50, 15
    CancelButton 135, 200, 50, 15
  Text 5, 10, 85, 10, "Enter your case number:"
  Text 5, 30, 90, 10, "Starting Footer Month/Year:"
  Text 20, 55, 120, 10, "future months and send through BG."
  Text 5, 70, 65, 10, "Worker Signature:"
  GroupBox 5, 100, 180, 95, "INSTRUCTIONS - PLEASE READ!!!"
  Text 10, 115, 170, 25, "This script is to help in correctly budgeting EARNED income on JOBS, BUSI, or RBIC. It will update MAXIS and CASE/NOTE the information provided. "
  Text 10, 150, 170, 40, "If a JOBS panel or BUSI panel needs to be added to MAXIS for a client or income source, the script will ask for any panels that need to be added first. Review the case now to ensure that the correct action will be taken in the correct order."
EndDialog

Do
    Do
        err_msg = ""
        dialog Dialog1
        cancel_confirmation

        If IsNumeric(MAXIS_case_number) = FALSE or Len(MAXIS_case_number) > 8 Then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
        If trim(worker_signature) = "" Then err_msg = err_msg & vbNewLine & "* Enter your worker signature for your case notes."

        If err_msg <> "" Then MsgBox "-- Please resolve the following to continue --" & vbNewLine & err_msg
    Loop until err_msg = ""
    call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false


MAXIS_footer_month = original_month
MAXIS_footer_year = originial_year

CASH_case = FALSE
SNAP_case = FALSE
HC_case = FALSE

Call Navigate_to_MAXIS_screen("STAT", "PROG")

EMReadScreen cash_one_status, 4, 6, 74
EMReadScreen cash_two_status, 4, 7, 74
EMReadScreen snap_status, 4, 10, 74
EMReadScreen hc_status, 4, 12, 74

If cash_one_status = "ACTV" OR cash_one_status = "PEND" Then CASH_case = TRUE
If cash_two_status = "ACTV" OR cash_two_status = "PEND" Then CASH_case = TRUE
If snap_status = "ACTV" OR snap_status = "PEND" Then SNAP_case = TRUE
If hc_status = "ACTV" OR hc_status = "PEND" Then HC_case = TRUE


                            '----------------------------------------------------------'
                    '---------------------------------------------------------------------------------'
'-------------------------------------------------FINDING ALL CURRENT EI PANELS --------------------------------------------------'
                    '---------------------------------------------------------------------------------'
                            '----------------------------------------------------------'


the_panel = 0
all_ei_panels_found = FALSE

call HH_member_custom_dialog(HH_member_array)

Call navigate_to_MAXIS_screen("STAT", "JOBS")
For each member in HH_member_array
    EMWriteScreen member, 20, 76
    'EMWriteScreen "01", 20, 79
    Transmit

    EMReadScreen number_of_jobs_panels, 1, 2, 78

    If number_of_jobs_panels <> "0" Then
        number_of_jobs_panels = number_of_jobs_panels * 1

        For panel = 1 to number_of_jobs_panels
            EMWriteScreen "0" & panel, 20, 79
            transmit

            save_this_panel = TRUE

            EMReadScreen end_date, 8, 9, 49

            If end_date <> "__ __ __" Then
                end_date = replace(end_date, " ", "/")
                end_date = DateValue(end_date)

                If DateDiff(end_date, date, "m") > 3 Then

                    BeginDialog Dialog1, 0, 0, 186, 140, "Dialog"
                      OptionGroup RadioGroup1
                        RadioButton 20, 45, 70, 10, "Delete this Panel", delete_panel
                        RadioButton 20, 60, 75, 10, "Leave this Panel", leave_ended_panel
                      EditBox 10, 100, 170, 15, explain_leaving_ended_panel
                      ButtonGroup ButtonPressed
                        OkButton 130, 120, 50, 15
                      Text 10, 10, 170, 25, "This JOBS panel indicates the income ended more than 3 months ago. This panel is no longer needed in this month since this income has ended."
                      Text 10, 80, 115, 20, "If this ened panel is to be left, explain why it is still needed:"
                    EndDialog

                    Do
                        Do
                            err_msg = ""

                            Dialog Dialog1

                            explain_leaving_ended_panel = trim(explain_leaving_ended_panel)

                            If leave_ended_panel = checked and explain_leaving_ended_panel = "" Then err_msg = err_msg & vbNewLine & "* If an ended panel is to be left on an active case, explain why it is still needed."

                            If err_msg <> "" Then MsgBox "** Please Resolve to Continue **" & vbNewLine & err_msg
                        Loop Until err_msg = ""
                        call check_for_password(are_we_passworded_out)
                    Loop until are_we_passworded_out = false

                    If delete_panel = checked then panels_to_delete = panels_to_delete & "~" & "JOBS " & member & " " & "0" & panel

                End If
            End If

            If save_this_panel = TRUE Then
                ReDim Preserve EARNED_INCOME_PANELS_ARRAY(convo_detail, the_panel)

                EARNED_INCOME_PANELS_ARRAY(panel_type, the_panel) = "JOBS"
                EARNED_INCOME_PANELS_ARRAY(panel_member, the_panel) = member
                EARNED_INCOME_PANELS_ARRAY(panel_instance, the_panel) = "0" & panel
                EARNED_INCOME_PANELS_ARRAY(income_received, the_panel) = FALSE
                If CASH_case = TRUE Then EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, the_panel) = checked
                If SNAP_case = TRUE Then EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, the_panel) = checked
                If HC_case = TRUE Then EARNED_INCOME_PANELS_ARRAY(apply_to_HC, the_panel) = checked

                EMReadScreen type_of_job, 1, 5, 34
                EMReadScreen job_verif, 25, 6, 34
                EMReadScreen listed_hrly_wage, 6, 6, 75
                EMReadScreen employer_name, 30, 7, 42
                EMReadScreen start_date, 8, 9, 35
                EMReadScreen end_date, 8, 9, 49
                EMReadScreen frequency, 1, 18, 35
                EMReadScreen current_verif, 27, 6, 34

                If type_of_job = "J" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "J - WIOA"
                If type_of_job = "W" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "W - Wages"
                If type_of_job = "E" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "E - EITC"
                If type_of_job = "G" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "G - Experience Works"
                If type_of_job = "F" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "F - Federal Work Study"
                If type_of_job = "S" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "S - State Work Study"
                If type_of_job = "O" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "O - Other"
                If type_of_job = "C" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "C - Contract Income"
                If type_of_job = "T" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "T - Training Program"
                If type_of_job = "P" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "P - Service Program"
                If type_of_job = "R" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "R - Rehab Program"

                EARNED_INCOME_PANELS_ARRAY(income_verif, the_panel) = trim(job_verif)
                EARNED_INCOME_PANELS_ARRAY(employer, the_panel) = replace(employer_name, "_", "")
                EARNED_INCOME_PANELS_ARRAY(hourly_wage, the_panel) = trim(listed_hrly_wage)
                EARNED_INCOME_PANELS_ARRAY(income_start_dt, the_panel) = replace(start_date, " ", "/")
                EARNED_INCOME_PANELS_ARRAY(income_end_dt, the_panel) = replace(end_date, " ", "/")
                EARNED_INCOME_PANELS_ARRAY(old_verif, the_panel) = trim(current_verif)
                If frequency = "1" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "1 - One Time Per Month"
                If frequency = "2" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "2 - Two Times Per Month"
                If frequency = "3" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "3 - Every Other Week"
                If frequency = "4" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "4 - Every Week"
                If frequency = "5" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "5 - Other"

                EARNED_INCOME_PANELS_ARRAY(income_list_indct, the_panel) = "NONE"
                'EARNED_INCOME_PANELS_ARRAY(, the_panel) =

                the_panel = the_panel + 1
            End If
        Next
    End If
Next

If panels_to_delete <> "" Then
    array_of_ended_panels = split(panels_to_delete, "~")
    For each panel in array_of_ended_panels
        Call back_to_SELF

        Call navigate_to_MAXIS_screen("STAT", "JOBS")
        EmWriteScreen Mid(panel, 6, 2), 20, 76
        EmWriteScreen right(panel, 2), 20, 79

        transmit

        EMReadScreen employer_name, 30, 7, 42

        Do
            Call navigate_to_MAXIS_screen("STAT", "JOBS")
            EmWriteScreen Mid(panel, 6, 2), 20, 76
            EmWriteScreen right(panel, 2), 20, 79

            transmit

            EMReadScreen this_employer_name, 30, 7, 42

            If this_employer_name = employer_name Then
                EmWriteScreen "DEL", 20, 71
                PF9
                transmit
            End If

            PF3

            EMReadScreen another_month, 11, 16, 3
            If another_month = "Update Next" Then
                EmWriteScreen "Y", 16, 54
                transmit
            End If
        Loop until another_month <> "Update Next"

    Next
End If

Call navigate_to_MAXIS_screen("STAT", "BUSI")
For each member in HH_member_array
    EMWriteScreen member, 20, 76
    'EMWriteScreen "01", 20, 79
    Transmit

    EMReadScreen number_of_busi_panels, 1, 2, 78

    If number_of_busi_panels <> "0" Then
        number_of_busi_panels = number_of_busi_panels * 1

        For panel = 1 to number_of_busi_panels
            EMWriteScreen "0" & panel, 20, 79
            transmit

            ReDim Preserve EARNED_INCOME_PANELS_ARRAY(convo_detail, the_panel)

            EARNED_INCOME_PANELS_ARRAY(panel_type, the_panel) = "BUSI"
            EARNED_INCOME_PANELS_ARRAY(panel_member, the_panel) = member
            EARNED_INCOME_PANELS_ARRAY(panel_instance, the_panel) = "0" & panel
            EARNED_INCOME_PANELS_ARRAY(income_received, the_panel) = FALSE
            If CASH_case = TRUE Then EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, the_panel) = checked
            If SNAP_case = TRUE Then EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, the_panel) = checked
            If HC_case = TRUE Then EARNED_INCOME_PANELS_ARRAY(apply_to_HC, the_panel) = checked

            EMReadScreen type_of_busi, 2, 5, 37
            EMReadScreen start_date, 8, 5, 55
            EMReadScreen end_date, 8, 5, 72
            EMReadScreen listed_method, 2, 16, 53
            EMReadScreen lst_mthd_date, 8, 16, 63

            If type_of_busi = "01" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "01 - Farming"
            If type_of_busi = "02" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "02 - Real Estate"
            If type_of_busi = "03" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "03 - Home Product Sales"
            If type_of_busi = "04" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "04 - Other Sales"
            If type_of_busi = "05" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "05 - Personal Services"
            If type_of_busi = "06" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "06 - Paper Route"
            If type_of_busi = "07" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "07 - In Home Daycare"
            If type_of_busi = "08" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "08 - Rental Income"
            If type_of_busi = "09" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "09 - Other"
            EARNED_INCOME_PANELS_ARRAY(income_start_dt, the_panel) = replace(start_date, " ", "/")
            EARNED_INCOME_PANELS_ARRAY(income_end_dt, the_panel) = replace(end_date, " ", "/")
            If listed_method = "01" Then EARNED_INCOME_PANELS_ARRAY(self_emp_mthd, the_panel) = "01 - 50% Gross Inc"
            If listed_method = "02" Then EARNED_INCOME_PANELS_ARRAY(self_emp_mthd, the_panel) = "02 - Tax Forms"
            EARNED_INCOME_PANELS_ARRAY(method_date, the_panel) = replace(lst_mthd_date, " ", "/")
            If EARNED_INCOME_PANELS_ARRAY(method_date, the_panel) = "__/__/__" Then EARNED_INCOME_PANELS_ARRAY(method_date, the_panel) = ""

            EmWriteScreen "X", 6, 26
            transmit

            For busi_row = 9 to 19
                EMReadScreen busi_verif, 1, busi_row, 73
                If busi_verif <> "_" Then
                    If busi_verif = "1" Then EARNED_INCOME_PANELS_ARRAY(old_verif, the_panel) = "1 - Income Tax Returns"
                    If busi_verif = "2" Then EARNED_INCOME_PANELS_ARRAY(old_verif, the_panel) = "2 - Receipts of Sales/Purchases"
                    If busi_verif = "3" Then EARNED_INCOME_PANELS_ARRAY(old_verif, the_panel) = "3 - Client BUSI Records/Ledger"
                    If busi_verif = "6" Then EARNED_INCOME_PANELS_ARRAY(old_verif, the_panel) = "6 - Other Document"
                    If busi_verif = "N" Then EARNED_INCOME_PANELS_ARRAY(old_verif, the_panel) = "N - NO Verif Provided"
                    Exit For
                End If
            Next
            PF3

            EARNED_INCOME_PANELS_ARRAY(income_list_indct, the_panel) = "NONE"
            EARNED_INCOME_PANELS_ARRAY(this_is_a_new_panel, the_panel) = FALSE

            the_panel = the_panel + 1
        Next
    End If
Next



                            '----------------------------------------------------------'
                    '---------------------------------------------------------------------------------'
'------------------------------------------------- ADDING NEW EI PANELS --------------------------------------------------'
                    '---------------------------------------------------------------------------------'
                            '----------------------------------------------------------'



Do
    y_pos = 25
    dlg_len = 15 * UBOUND(EARNED_INCOME_PANELS_ARRAY, 2) + 15 * UBOUND(HH_member_array) + 125
    BeginDialog Dialog1, 0, 0, 330, dlg_len, "Case Number"

      Text 5, 10, 105, 10, "Known JOBS and BUSI panels:"

      For ei_panel = 0 to UBOUND(EARNED_INCOME_PANELS_ARRAY, 2)
        earned_income_panel_detail = EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel) & " " & EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel) & " " & EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel) & " - "
        If EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel) = "JOBS" Then
            earned_income_panel_detail = earned_income_panel_detail & EARNED_INCOME_PANELS_ARRAY(employer, ei_panel)
        ElseIf EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel) = "BUSI" Then
            earned_income_panel_detail = earned_income_panel_detail & "TYPE: " & EARNED_INCOME_PANELS_ARRAY(income_type, ei_panel)
        End If
        earned_income_panel_detail = earned_income_panel_detail & " - Income Start: " & EARNED_INCOME_PANELS_ARRAY(income_start_dt, ei_panel) & " - Verif: " & EARNED_INCOME_PANELS_ARRAY(old_verif, ei_panel)
        Text 10, y_pos, 305, 10, earned_income_panel_detail
        'Text 10, 25, 295, 10, "JOBS 01 01 - EMPLOYER - Income Start: mm/dd/yy - Verif: N"
        'Text 10, 40, 295, 10, "BUSI 01 01 - TYPE: 04 - Other Sales - Income Start: mm/dd/yy - Verif: N"
        y_pos = y_pos + 15
      Next
      y_pos = y_pos + 5
      Text 5, y_pos, 295, 10, "These are all the panels that are currently known in MAXIS for these Household Members:"
      y_pos = y_pos + 15
      For each member in HH_member_array
        Text 10, y_pos, 45, 10, "Member " & member
        y_pos = y_pos + 15
        'Text 10, 75, 45, 10, "MEMBER 01"
      Next
      y_pos = y_pos - 10
      Text 80, y_pos, 160, 10, "Do you need to add a new JOBS or BUSI panel?"
      ButtonGroup ButtonPressed
        PushButton 85, y_pos + 15, 140, 20, "Yes - Add a new Earned Income panel", add_new_panel_button
        PushButton 85, y_pos + 40, 140, 10, "No - The panel(s) to update are in MAXIS", continue_to_update_button
    EndDialog
    'MsgBOx "Y Position is " & y_pos & vbNewLine & "Dialog length is " & dlg_len

    dialog Dialog1

    If buttonpressed = add_new_panel_button Then
        original_month = MAXIS_footer_month
        original_year = MAXIS_footer_year
        panel_to_add = "JOBS"
        'MsgBox "Add a new panel!"
        '2 different dialogs for JOBS vs BUSI and add here then add to the EARNED_INCOME_PANELS_ARRAY

        BeginDialog Dialog1, 0, 0, 191, 50, "Panel to Add"
          DropListBox 30, 30, 60, 45, "Select one..."+chr(9)+"JOBS"+chr(9)+"BUSI", panel_to_add
          ButtonGroup ButtonPressed
            OkButton 135, 10, 50, 15
            CancelButton 135, 30, 50, 15
          Text 15, 10, 85, 20, "Which type of panel would you like to add?"
        EndDialog

        Do
            Do
                err_msg = ""

                dialog Dialog1
                cancel_confirmation

                If panel_to_add = "Select one..." Then err_msg = err_msg & vbNewLine & "* Indicate which type of panel needs to be added."

                If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg

            Loop until err_msg = ""
            call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
        LOOP UNTIL are_we_passworded_out = false


        info_saved = FALSE

        If CASH_case = TRUE Then cash_checkbox = checked
        If SNAP_case = TRUE Then snap_checkbox = checked
        If HC_case = TRUE Then hc_checkbox = checked

        Select Case panel_to_add

        Case "JOBS"
            'Start on DIALOG need to keep working on it'
            BeginDialog Dialog1, 0, 0, 431, 115, "New JOBS Panel"
              EditBox 75, 10, 20, 15, enter_JOBS_clt_ref_nbr
              DropListBox 155, 10, 60, 45, "W - Wages (Incl Tips)"+chr(9)+"J - WIOA"+chr(9)+"E - EITC"+chr(9)+"G - Experience Works"+chr(9)+"F - Federal Work Study"+chr(9)+"S - State Work Study"+chr(9)+"O - Other"+chr(9)+"C - Contract Income"+chr(9)+"T - Training Program"+chr(9)+"P - Service Program"+chr(9)+"R - Rehab Program", enter_JOBS_inc_type_code
              DropListBox 330, 10, 95, 45, ""+chr(9)+"01 - Subsidized Public Sector Employer"+chr(9)+"02 - Subsidized Private Sector Employer"+chr(9)+"03 - On-The-Job Training"+chr(9)+"04 - AmeriCorps(VISTA/State/National/NCCC)", enter_JOBS_subsdzd_inc_type
              DropListBox 155, 30, 90, 45, "1 - Pay Stubs/Tip Report"+chr(9)+"2 - Empl Statement"+chr(9)+"3 - Coltrl Stmt"+chr(9)+"4 - Other Document"+chr(9)+"5 - Pend Out State Verification"+chr(9)+"N - No Ver Prvd"+chr(9)+"? - Unknown", enter_JOBS_verif_code
              EditBox 330, 30, 50, 15, enter_JOBS_hrly_wage
              EditBox 155, 50, 195, 15, enter_JOBS_employer
              EditBox 155, 70, 50, 15, enter_JOBS_start_date
              EditBox 330, 70, 50, 15, enter_JOBS_end_date
              CheckBox 105, 95, 30, 10, "SNAP", snap_checkbox
              CheckBox 145, 95, 30, 10, "CASH", cash_checkbox
              CheckBox 190, 95, 20, 10, "HC", hc_checkbox
              ButtonGroup ButtonPressed
                OkButton 320, 95, 50, 15
                CancelButton 375, 95, 50, 15
              Text 10, 15, 65, 10, "Client Ref Number:"
              Text 105, 15, 45, 10, "Income Type:"
              Text 240, 15, 85, 10, "Subsidized Income Type:"
              Text 110, 35, 40, 10, "Verification:"
              Text 280, 35, 50, 10, "Hourly Wage:"
              Text 115, 55, 35, 10, "Employer:"
              Text 105, 75, 45, 10, "Income Start:"
              Text 285, 75, 40, 10, "Income End:"
              Text 10, 95, 90, 10, "Apply Income to Programs:"
            EndDialog

            Do
                Do
                    err_msg = ""

                    dialog Dialog1
                    cancel_confirmation

                    If trim(enter_JOBS_clt_ref_nbr) = "" Then err_msg = err_msg & vbNewLine & "* Enter the member number of the client that is employed at this job."
                    If len(enter_JOBS_clt_ref_nbr) <> 2 Then err_msg = err_msg & vbNewLine & "* The member number should be 2 digits"
                    If trim(enter_JOBS_inc_type_code) = "  " Then err_msg = err_msg & vbNewLine & "* Enter the income type of the job."
                    If trim(enter_JOBS_verif_code) = "  " Then err_msg = err_msg & vbNewLine & "* Enter the verification code for this job."
                    If trim(enter_JOBS_employer) = "" Then err_msg = err_msg & vbNewLine & "* Enter the employer name for this job."
                    If IsDate(enter_JOBS_start_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date for Income start date."
                    If trim(enter_JOBS_end_date) <> "" AND IsDate(enter_JOBS_end_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date for Income end date."

                    if err_msg <> "" Then msgBox "Please resolve the following to continue:" & vbNewLine & err_msg

                Loop until err_msg = ""
                call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
            LOOP UNTIL are_we_passworded_out = false


            Call navigate_to_MAXIS_screen("CASE", "CURR")

            EMReadScreen appl_date, 8, 8, 29

            If DateDiff("m", appl_date, enter_JOBS_start_date) > 0 Then
                beginning_month = DatePart("m", enter_JOBS_start_date)
                beginning_year = DatePart("yyyy", enter_JOBS_start_date)
                first_check = enter_JOBS_start_date
            Else
                beginning_month = DatePart("m", appl_date)
                beginning_year = DatePart("yyyy", appl_date)
                first_check = beginning_month & "/01/" & beginning_year
            End If


            beginning_month = right("00"&beginning_month, 2)
            beginning_year = right(beginning_year, 2)
            'MsgBox "Begind date is " & beginning_month & "/" & beginning_year

            MAXIS_footer_month = beginning_month
            MAXIS_footer_year = beginning_year

            Call back_to_SELF
            'QUESTION - what to do about footer month'

            Call navigate_to_MAXIS_screen("STAT", "SUMM")

            Do

                If info_saved = FALSE Then
                    EMWriteScreen "JOBS", 20, 71
                    EMWriteScreen enter_JOBS_clt_ref_nbr, 20, 76
                    EMWriteScreen "NN", 20, 79

                    transmit
                Else
                    EMWriteScreen "JOBS", 20, 71
                    EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_member, the_panel-1), 20, 76
                    EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_instance, the_panel-1), 20, 79

                    transmit
                    'MsgBox "Pause"

                    EMReadScreen check_for_panel, 14, 24, 13
                    'MsgBox check_for_panel
                    If check_for_panel = "DOES NOT EXIST" Then
                        EMWriteScreen "JOBS", 20, 71
                        EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_member, the_panel-1), 20, 76
                        EMWriteScreen "NN", 20, 79

                        transmit
                    Else
                        PF9
                    End If
                End If
                'MsgBox "nav to panel"

                EMWriteScreen left(enter_JOBS_inc_type_code, 1), 5, 34
                EMWriteScreen left(enter_JOBS_subsdzd_inc_type, 2), 5, 74
                EMWriteScreen left(enter_JOBS_verif_code, 1), 6, 34
                EMWriteScreen enter_JOBS_hrly_wage, 6, 75
                EMWriteScreen enter_JOBS_employer, 7, 42

                Call write_date(enter_JOBS_start_date, "MM DD YY", 9, 35)
                If trim(enter_JOBS_end_date) <> "" Then Call write_date(enter_JOBS_end_date, "MM DD YY", 9, 49)

                Call write_date(first_check, "MM DD YY", 12, 54)
                EMWriteScreen "    0.00", 12, 67
                EMWriteScreen "0  ", 18, 72
                'MsgBox "Pause before transmit"
                transmit
                'MsgBox "Pause after transmit"

                If info_saved = FALSE Then
                    ReDim Preserve EARNED_INCOME_PANELS_ARRAY(convo_detail, the_panel)

                    EARNED_INCOME_PANELS_ARRAY(panel_type, the_panel) = "JOBS"
                    EARNED_INCOME_PANELS_ARRAY(panel_member, the_panel) = member
                    EARNED_INCOME_PANELS_ARRAY(panel_instance, the_panel) = "0" & panel
                    EARNED_INCOME_PANELS_ARRAY(income_received, the_panel) = FALSE
                    EARNED_INCOME_PANELS_ARRAY(initial_month_mo, the_panel) = MAXIS_footer_month
                    EARNED_INCOME_PANELS_ARRAY(initial_month_yr, the_panel) = MAXIS_footer_year

                    If CASH_case = TRUE Then EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, the_panel) = checked
                    If SNAP_case = TRUE Then EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, the_panel) = checked
                    If HC_case = TRUE Then EARNED_INCOME_PANELS_ARRAY(apply_to_HC, the_panel) = checked

                    EMReadScreen type_of_job, 1, 5, 34
                    EMReadScreen job_verif, 25, 6, 34
                    EMReadScreen listed_hrly_wage, 6, 6, 75
                    EMReadScreen employer_name, 30, 7, 42
                    EMReadScreen start_date, 8, 9, 35
                    EMReadScreen end_date, 8, 9, 49
                    EMReadScreen frequency, 1, 18, 35
                    EMReadScreen current_verif, 27, 6, 34

                    If type_of_job = "J" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "J - WIOA"
                    If type_of_job = "W" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "W - Wages"
                    If type_of_job = "E" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "E - EITC"
                    If type_of_job = "G" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "G - Experience Works"
                    If type_of_job = "F" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "F - Federal Work Study"
                    If type_of_job = "S" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "S - State Work Study"
                    If type_of_job = "O" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "O - Other"
                    If type_of_job = "C" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "C - Contract Income"
                    If type_of_job = "T" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "T - Training Program"
                    If type_of_job = "P" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "P - Service Program"
                    If type_of_job = "R" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "R - Rehab Program"

                    EARNED_INCOME_PANELS_ARRAY(income_verif, the_panel) = trim(job_verif)
                    EARNED_INCOME_PANELS_ARRAY(employer, the_panel) = replace(employer_name, "_", "")
                    EARNED_INCOME_PANELS_ARRAY(hourly_wage, the_panel) = trim(listed_hrly_wage)
                    EARNED_INCOME_PANELS_ARRAY(income_start_dt, the_panel) = replace(start_date, " ", "/")
                    EARNED_INCOME_PANELS_ARRAY(income_end_dt, the_panel) = replace(end_date, " ", "/")
                    EARNED_INCOME_PANELS_ARRAY(old_verif, the_panel) = trim(current_verif)
                    If frequency = "1" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "1 - One Time Per Month"
                    If frequency = "2" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "2 - Two Times Per Month"
                    If frequency = "3" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "3 - Every Other Week"
                    If frequency = "4" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "4 - Every Week"
                    If frequency = "5" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "5 - Other"

                    EARNED_INCOME_PANELS_ARRAY(income_list_indct, the_panel) = "NONE"

                    EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, the_panel) = snap_checkbox
                    EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, the_panel) = cash_checkbox
                    EARNED_INCOME_PANELS_ARRAY(apply_to_HC, the_panel) = hc_checkbox
                    'EARNED_INCOME_PANELS_ARRAY(, the_panel) =

                    the_panel = the_panel + 1

                    info_saved = TRUE
                End If

                ' EMReadScreen expire_check, 11, 24, 12
                '
                ' If expire_check = "WILL EXPIRE" THen
                '     Msgbox "need to update in other months"
                '
                '
                '
                ' End If
                'Navigates to the current month + 1 footer month, then back into the JOBS panel
                CALL write_value_and_transmit("BGTX", 20, 71)
                CALL write_value_and_transmit("y", 16, 54)

                EMReadScreen all_months_check, 24, 24, 2
                EMReadScreen MAXIS_footer_month, 2, 20, 55
                EMReadScreen MAXIS_footer_year, 2, 20, 58

                first_check = MAXIS_footer_month & "/01/" & MAXIS_footer_year
            Loop until all_months_check = "CONTINUATION NOT ALLOWED"
            PF3

            EARNED_INCOME_PANELS_ARRAY(this_is_a_new_panel, the_panel) = TRUE
        Case "BUSI"

        End Select

        MAXIS_footer_month = original_month
        MAXIS_footer_year = original_year
        Call back_to_SELF
    End If

Loop until buttonpressed = continue_to_update_button


                                '----------------------------------------------------------'
                        '---------------------------------------------------------------------------------'
'------------------------------------------------- GATHERING PAY INFORMATION FOR EACH PANEL --------------------------------------------------'
                        '---------------------------------------------------------------------------------'
                                '----------------------------------------------------------'


'STILL need some handling for scheduled income with no actual checks or cases where scheduled income is different from actual checks but we get both.


pay_item = 0
For ei_panel = 0 to UBOUND(EARNED_INCOME_PANELS_ARRAY, 2)


    If EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel) = "JOBS" Then

        Call Navigate_to_MAXIS_screen("STAT", "JOBS")
        EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel), 20, 76
        EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel), 20, 79
        transmit

        employer_check = MsgBox("Do you have income verification for this job? Employer name: " & EARNED_INCOME_PANELS_ARRAY(employer, ei_panel), vbYesNo + vbQuestion, "Select Income Panel")

        If EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel) = "" Then EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel) = MAXIS_footer_month
        If EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel) = "" Then EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel) = MAXIS_footer_year
        EARNED_INCOME_PANELS_ARRAY(update_futue_chkbx, ei_panel) = future_months_check

        If employer_check = vbYes Then
            EARNED_INCOME_PANELS_ARRAY(income_received, ei_panel) = TRUE
            review_small_dlg = TRUE
            EARNED_INCOME_PANELS_ARRAY(ignore_antic, ei_panel) = FALSE
            Do
                Do
                    big_err_msg = ""


                    If review_small_dlg = TRUE Then
                        Do
                            Do
                                dlg_factor = 0
                                Dialog1 = ""

                                LIST_OF_INCOME_ARRAY(panel_indct, pay_item) = ei_panel

                                If LIST_OF_INCOME_ARRAY(panel_indct, 0) <> "" Then
                                    For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                                        If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then dlg_factor = dlg_factor + 1
                                    Next
                                End If

                                dlg_factor = dlg_factor - 1

                                sm_err_msg = ""

                                'MsgBox "Dialog Factor: " & dlg_factor

                                BeginDialog Dialog1, 0, 0, 606, (dlg_factor * 20) + 160, "Enter ALL Paychecks Received"
                                  Text 10, 10, 265, 10, "JOBS " & EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel) & " " & EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel) & " - " & EARNED_INCOME_PANELS_ARRAY(employer, ei_panel)
                                  Text 220, 15, 40, 10, "Start Date:"
                                  EditBox 255, 10, 50, 15, EARNED_INCOME_PANELS_ARRAY (income_start_dt, ei_panel)
                                  Text 315, 15, 50, 10, "Income Type:"
                                  DropListBox 365, 10, 100, 45, "J - WIOA"+chr(9)+"W - Wages"+chr(9)+"E - EITC"+chr(9)+"G - Experience Works"+chr(9)+"F - Federal Work Study"+chr(9)+"S - State Work Study"+chr(9)+"O - Other"+chr(9)+"C - Contract Income"+chr(9)+"T - Training Program"+chr(9)+"P - Service Program"+chr(9)+"R - Rehab Program", EARNED_INCOME_PANELS_ARRAY(income_type, ei_panel)
                                  GroupBox 475, 5, 125, 25, "Apply Income to Programs:"
                                  CheckBox 485, 15, 30, 10, "SNAP", EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel)
                                  CheckBox 530, 15, 30, 10, "CASH", EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, ei_panel)
                                  CheckBox 570, 15, 20, 10, "HC", EARNED_INCOME_PANELS_ARRAY(apply_to_HC, ei_panel)
                                  Text 5, 40, 60, 10, "JOBS Verif Code:"
                                  DropListBox 65, 35, 105, 45, "1 - Pay Stubs/Tip Report"+chr(9)+"2 - Empl Statement"+chr(9)+"3 - Coltrl Stmt"+chr(9)+"4 - Other Document"+chr(9)+"5 - Pend Out State Verification"+chr(9)+"N - No Ver Prvd", EARNED_INCOME_PANELS_ARRAY(income_verif, ei_panel)
                                  Text 175, 40, 155, 10, "additional detail of verification received:"
                                  EditBox 310, 35, 290, 15, EARNED_INCOME_PANELS_ARRAY(verif_explain, ei_panel)
                                  Text 5, 60, 90, 10, "Date verification received:"
                                  EditBox 100, 55, 50, 15, EARNED_INCOME_PANELS_ARRAY(verif_date, ei_panel)
                                  Text 5, 80, 80, 10, "Pay Date (MM/DD/YY):"
                                  Text 90, 80, 50, 10, "Gross Amount:"
                                  Text 145, 80, 25, 10, "Hours:"
                                  Text 180, 65, 25, 25, "Use in SNAP budget"
                                  Text 235, 80, 85, 10, "If not used, explain why:"
                                  Text 355, 70, 245, 10, "If there is a specific amount that should be NOT budgeted from this check:"
                                  Text 355, 80, 30, 10, "Amount:"
                                  Text 410, 80, 30, 10, "Reason:"

                                  y_pos = 0
                                  For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                                      If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then
                                          EditBox 5, (y_pos * 20) + 90, 65, 15, LIST_OF_INCOME_ARRAY(pay_date, all_income) 'pay_date'
                                          EditBox 90, (y_pos * 20) + 90, 45, 15, LIST_OF_INCOME_ARRAY(gross_amount, all_income) 'gross_amount'
                                          EditBox 145, (y_pos * 20) + 90, 25, 15, LIST_OF_INCOME_ARRAY(hours, all_income) 'hours_on_check'

                                          CheckBox 180, (y_pos * 20) + 95, 50, 10, "Exclude", LIST_OF_INCOME_ARRAY(budget_in_SNAP_no, all_income)
                                          ' OptionGroup RadioGroup1
                                          '   If LIST_OF_INCOME_ARRAY(budget_in_SNAP_no, all_income) <> 1 Then LIST_OF_INCOME_ARRAY(budget_in_SNAP_yes, all_income) = 1
                                          '   RadioButton 180, (y_pos * 20) + 90, 25, 10, "Yes", LIST_OF_INCOME_ARRAY(budget_in_SNAP_yes, all_income) 'budget_yes'
                                          '   RadioButton 210, (y_pos * 20) + 90, 25, 10, "No", LIST_OF_INCOME_ARRAY(budget_in_SNAP_no, all_income) 'budget_no'
                                          EditBox 235, (y_pos * 20) + 90, 115, 15, LIST_OF_INCOME_ARRAY(reason_to_exclude, all_income) 'reason_not_budgeted'
                                          EditBox 355, (y_pos * 20) + 90, 45, 15, LIST_OF_INCOME_ARRAY(exclude_amount, all_income) 'not_budgeted_amount'
                                          EditBox 410, (y_pos * 20) + 90, 185, 15, LIST_OF_INCOME_ARRAY(reason_amt_excluded, all_income) 'amount_not_budgeted_reason'
                                          y_pos = y_pos + 1
                                      End If
                                  Next


                                  Text 5, (dlg_factor * 20) + 115, 70, 10, "Anticipated Income"
                                  Text 5, (dlg_factor * 20) + 130, 50, 10, "Rate of Pay/Hr"
                                  Text 75, (dlg_factor * 20) + 130, 35, 10, "Hours/Wk"
                                  Text 130, (dlg_factor * 20) + 130, 50, 10, "Pay Frequency"
                                  Text 225, (dlg_factor * 20) + 115, 70, 10, "Regular Non-Monthly"
                                  Text 225, (dlg_factor * 20) + 130, 25, 10, "Amount"
                                  Text 280, (dlg_factor * 20) + 130, 50, 10, "Nbr of Months"
                                  Text 340, (dlg_factor * 20) + 130, 65, 10, "Known Pay Date"
                                  EditBox 5, (dlg_factor * 20) + 140, 50, 15, EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel)
                                  EditBox 75, (dlg_factor * 20) + 140, 40, 15, EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel)
                                  DropListBox 130, (dlg_factor * 20) + 140, 85, 45, ""+chr(9)+"1 - One Time Per Month"+chr(9)+"2 - Two Times Per Month"+chr(9)+"3 - Every Other Week"+chr(9)+"4 - Every Week", EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel)
                                  EditBox 225, (dlg_factor * 20) + 140, 40, 15, EARNED_INCOME_PANELS_ARRAY(reg_non_monthly, ei_panel)
                                  EditBox 280, (dlg_factor * 20) + 140, 30, 15, EARNED_INCOME_PANELS_ARRAY(numb_months, ei_panel)
                                  EditBox 340, (dlg_factor * 20) + 140, 65, 15, known_pay_date
                                  Text 455, (dlg_factor * 20) + 115, 85, 10, "Initial Month to Update:"
                                  EditBox 535, (dlg_factor * 20) + 110, 15, 15, EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel)
                                  EditBox 560, (dlg_factor * 20) + 110, 15, 15, EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel)
                                  CheckBox 455, (dlg_factor * 20) + 125, 120, 10, "Update Future Months", EARNED_INCOME_PANELS_ARRAY(update_futue_chkbx, ei_panel)

                                  ButtonGroup ButtonPressed
                                    PushButton 455, (dlg_factor * 20) + 140, 15, 15, "+", add_another_check
                                    PushButton 475, (dlg_factor * 20) + 140, 15, 15, "-", take_a_check_away
                                    OkButton 495, (dlg_factor * 20) + 140, 50, 15
                                    CancelButton 550, (dlg_factor * 20) + 140, 50, 15
                                EndDialog

                                Dialog Dialog1
                                cancel_confirmation

                                actual_checks_provided = FALSE
                                For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                                    LIST_OF_INCOME_ARRAY(pay_date, all_income) = trim(LIST_OF_INCOME_ARRAY(pay_date, all_income))
                                    LIST_OF_INCOME_ARRAY(gross_amount, all_income) = trim(LIST_OF_INCOME_ARRAY(gross_amount, all_income))
                                    LIST_OF_INCOME_ARRAY(hours, all_income) = trim(LIST_OF_INCOME_ARRAY(hours, all_income))
                                    If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(pay_date, all_income) <> "" AND LIST_OF_INCOME_ARRAY(gross_amount, all_income) <> "" AND LIST_OF_INCOME_ARRAY(hours, all_income) <> "" Then
                                        'ADD ERROR HANDLING HERE
                                        actual_checks_provided = TRUE
                                        If IsDate(LIST_OF_INCOME_ARRAY(pay_date, all_income)) = FALSE Then sm_err_msg = sm_err_msg & vbNewLine & "* Enter a valid pay date for all checks."
                                        If IsNumeric(LIST_OF_INCOME_ARRAY(gross_amount, all_income)) = FALSE Then sm_err_msg = sm_err_msg & vbNewLine & "* Enter the Gross Amount of the check as a number."
                                        If LIST_OF_INCOME_ARRAY(budget_in_SNAP_no, all_income) = 1 AND trim(LIST_OF_INCOME_ARRAY(reason_to_exclude, all_income)) = "" Then sm_err_msg = sm_err_msg & vbNewLine & "* The check on " & LIST_OF_INCOME_ARRAY(pay_date, all_income) & " is to be excluded, list a reason for excluding this check."
                                        If IsNumeric(LIST_OF_INCOME_ARRAY(hours, all_income)) = FALSE Then sm_err_msg = sm_err_msg & vbNewLine & "* Enter the number of hours for the paycheck on " & LIST_OF_INCOME_ARRAY(pay_date, all_income) & " as a number."
                                        If IsNumeric(LIST_OF_INCOME_ARRAY(exclude_amount, all_income)) = FALSE AND trim(LIST_OF_INCOME_ARRAY(exclude_amount, all_income)) <> "" Then sm_err_msg = sm_err_msg & vbNewLine & "* Enter the amount excluded from the budget as a number."
                                        LIST_OF_INCOME_ARRAY(exclude_amount, all_income) = trim(LIST_OF_INCOME_ARRAY(exclude_amount, all_income))
                                    End If
                                Next
                                anticipated_income_provided = FALSE
                                EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel) = trim(EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel))
                                EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) = trim(EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel))
                                EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = trim(EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel))

                                EARNED_INCOME_PANELS_ARRAY(reg_non_monthly, ei_panel) = trim(EARNED_INCOME_PANELS_ARRAY(reg_non_monthly, ei_panel))
                                EARNED_INCOME_PANELS_ARRAY(numb_months, ei_panel) = trim(EARNED_INCOME_PANELS_ARRAY(numb_months, ei_panel))
                                known_pay_date = trim(known_pay_date)

                                If EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel) <> "" AND EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) <> "" AND EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) <> "" Then
                                    anticipated_income_provided = TRUE
                                    If EARNED_INCOME_PANELS_ARRAY(income_start_dt, ei_panel) = "" Then sm_err_msg = sm_err_msg & vbNewLine & "* Enter an income start date, since anticipated pay dates cannot be determined without the initial pay date."
                                End If
                                If EARNED_INCOME_PANELS_ARRAY(reg_non_monthly, ei_panel) <> "" AND EARNED_INCOME_PANELS_ARRAY(numb_months, ei_panel) <> "" Then anticipated_income_provided = TRUE

                                If anticipated_income_provided = FALSE AND actual_checks_provided = FALSE Then
                                    sm_err_msg = sm_err_msg & vbNewLine & "* Income information needs to be provided, either in the form of actual checks or anticipated income, hours, and rate of pay."
                                End If
                                If known_pay_date <> "" AND IsDate(known_pay_date) = FALSE Then sm_err_msg = sm_err_msg & vbNewLine & "* A known pay date needs to be entered as a date. Check the entry."
                                If ButtonPressed = add_another_check Then
                                    pay_item = pay_item + 1
                                    ReDim Preserve LIST_OF_INCOME_ARRAY(reason_amt_excluded, pay_item)
                                    LIST_OF_INCOME_ARRAY(panel_indct, pay_item) = ei_panel
                                    dlg_factor = dlg_factor + 1

                                    LIST_OF_INCOME_ARRAY(pay_date, all_income) = ""
                                    LIST_OF_INCOME_ARRAY(gross_amount, all_income) = ""
                                    LIST_OF_INCOME_ARRAY(hours, all_income) = ""
                                    LIST_OF_INCOME_ARRAY(reason_to_exclude, all_income) = ""
                                    LIST_OF_INCOME_ARRAY(exclude_amount, all_income) = ""
                                    LIST_OF_INCOME_ARRAY(reason_amt_excluded, all_income) = ""

                                    sm_err_msg = "LOOP" & sm_err_msg

                                End If

                                If ButtonPressed = take_a_check_away Then
                                    pay_item = pay_item - 1
                                    If pay_item < 0 Then pay_item = 0
                                    ReDim Preserve LIST_OF_INCOME_ARRAY(reason_amt_excluded, pay_item)
                                    dlg_factor = dlg_factor - 1
                                    sm_err_msg = "LOOP" & sm_err_msg
                                End If

                                If sm_err_msg <> "" AND left(sm_err_msg, 4) <> "LOOP" then MsgBox "Please resolve before continuing:" & vbNewLine & sm_err_msg

                            Loop until sm_err_msg = ""
                            call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
                        LOOP UNTIL are_we_passworded_out = false


                        actual_checks_provided = FALSE
                        For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                            If LIST_OF_INCOME_ARRAY(pay_date, all_income) = "" AND LIST_OF_INCOME_ARRAY(gross_amount, all_income) = "" AND LIST_OF_INCOME_ARRAY(hours, all_income) = "" Then LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ""
                            If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then
                                'ADD ERROR HANDLING HERE
                                actual_checks_provided = TRUE
                            End If
                        Next
                    End If

                    If actual_checks_provided = TRUE AND anticipated_income_provided = TRUE Then
                        BeginDialog Dialog1, 0, 0, 196, 165, "Reasonably Expected to Continue"
                          OptionGroup RadioGroup1
                            RadioButton 25, 70, 130, 10, "Use the actual check amounts/dates", use_actual_income
                            RadioButton 25, 85, 130, 10, "Use the anticipated hours/wage", use_anticipated_income
                          EditBox 10, 125, 180, 15, EARNED_INCOME_PANELS_ARRAY(selection_rsn, ei_panel)
                          ButtonGroup ButtonPressed
                            OkButton 140, 145, 50, 15
                          Text 10, 10, 185, 35, "Both Actual Income and Anticipated Income have been listed for a SNAP case. Since both have been reported, both will be case ntoed. For entering information to the PIC, one option should be selected."
                          GroupBox 5, 55, 185, 45, "Which is the best estimation of anticipated income?"
                          Text 10, 110, 185, 10, "Explain why this is the best estimation of future income:"
                        EndDialog

                        Do
                            Do
                                sm_err_msg = ""

                                Dialog Dialog1

                                EARNED_INCOME_PANELS_ARRAY(selection_rsn, ei_panel) = trim(EARNED_INCOME_PANELS_ARRAY(selection_rsn, ei_panel))
                                If use_actual_income = checked Then selection_pick = "ACTUAL LIST OF CHECKS."
                                If use_anticipated_income = checked Then selection_pick = "INCOME ESTIMATED FROM HOURS AND RATE OF PAY."

                                If EARNED_INCOME_PANELS_ARRAY(selection_rsn, ei_panel) = "" Then sm_err_msg = sm_err_msg & vbNewLine & "* Enter explanation of why the best way to determine future income is to use " & selection_pick
                                If len(EARNED_INCOME_PANELS_ARRAY(selection_rsn, ei_panel)) < 10 Then sm_err_msg = sm_err_msg & vbNewLine & "* Explanation is not sufficient to adequately case note information about budget. Expand."

                                If sm_err_msg <> "" Then MsgBox "** Please Resolve before Continuting **" & vbNewLine & sm_err_msg
                            Loop until sm_err_msg = ""
                            call check_for_password(are_we_passworded_out)
                        Loop until are_we_passworded_out = false

                        If use_actual_income = checked Then
                            EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_actual
                            EARNED_INCOME_PANELS_ARRAY(ignore_antic, ei_panel) = TRUE
                        End If
                        If use_anticipated_income = checked Then
                            EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_estimate

                            For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                                If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then
                                    LIST_OF_INCOME_ARRAY(budget_in_SNAP_no, all_income) = checked
                                    LIST_OF_INCOME_ARRAY(reason_to_exclude, all_income) = "Not best estimate of Anticipated Income"
                                    list_of_actual_paydates = list_of_actual_paydates & "~" & LIST_OF_INCOME_ARRAY(pay_date, all_income)
                                    list_of_excluded_pay_dates = list_of_excluded_pay_dates & ", " & LIST_OF_INCOME_ARRAY(pay_date, all_income)
                                End If
                            Next
                            list_of_actual_paydates = right(list_of_actual_paydates, len(list_of_actual_paydates) - 1)
                            If list_of_excluded_pay_dates <> "" Then list_of_excluded_pay_dates = right(list_of_excluded_pay_dates, len(list_of_excluded_pay_dates) - 2)
                        End If

                        'https://www.dhssir.cty.dhs.state.mn.us/MAXIS/trntl/snap/SNAP_Anticipating_Income.pdf
                    Else
                        If actual_checks_provided = TRUE Then EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_actual
                        If anticipated_income_provided = TRUE Then EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_estimate
                    End If

                    If EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_actual Then
                        total_of_counted_income = 0
                        total_of_included_pay_checks = 0
                        total_of_hours = 0
                        number_of_checks_budgeted = 0
                        cash_checks = 0
                        EARNED_INCOME_PANELS_ARRAY(pay_weekday, ei_panel) = ""
                        list_of_excluded_pay_dates = ""
                        previous_pay_date = ""
                        paycheck_list_title = "Paychecks Inclued in Budget:"

                        'Adding the order to the array for what the order the checks should be in
                        '-----THis block works to display in order------'
                        all_pay_dates = ""          'blanking out for each loop of different EI panels
                        array_of_pay_dates = ""
                        For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)                                   'look at each entry inthe income array
                            If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then                    'find the ones for the current panel
                                'MsgBox "Look at each date: " & LIST_OF_INCOME_ARRAY(pay_date, all_income)
                                all_pay_dates = all_pay_dates & "~" & LIST_OF_INCOME_ARRAY(pay_date, all_income)'create a list of just the pay dates
                            End If
                        Next
                        If all_pay_dates <> "" Then all_pay_dates = right(all_pay_dates, len(all_pay_dates)-1)      'make a single dimension array of the pay dates for this one panel
                        array_of_pay_dates = split(all_pay_dates, "~")

                        Call sort_dates(array_of_pay_dates)                             'use the function to re order that array into chronological order.
                        first_date = array_of_pay_dates(0)
                        last_date = array_of_pay_dates(UBOUND(array_of_pay_dates))
                        EARNED_INCOME_PANELS_ARRAY(last_paycheck, ei_panel) = last_date
                        spread_of_pay_dates = DateDiff("d", first_date, last_date)
                        If EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) = checked THen
                            using_30_days = TRUE

                            If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "1 - One Time Per Month" Then
                                If spread_of_pay_dates > 30 Then using_30_days = FALSE
                            ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "2 - Two Times Per Month" Then
                                If spread_of_pay_dates > 30 Then using_30_days = FALSE
                                If spread_of_pay_dates < 13 Then using_30_days = FALSE
                            ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "3 - Every Other Week" Then
                                If spread_of_pay_dates <> 28 Then using_30_days = FALSE
                            ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "4 - Every Week" Then
                                If spread_of_pay_dates <> 28 Then using_30_days = FALSE
                            ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "5 - Other" Then
                            End If

                            ' MsgBox "First pay date: " & first_date & vbNewLine & "Last pay date: " & last_date & vbNewLine & "Spread - " & spread_of_pay_dates & vbNewLine & "30 days of income - " & using_30_days
                        End If

                        For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)           'Now loop through all of the listed income - again
                            If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then    'find the ones for THIS PANEL ONLY
                                for index = 0 to UBOUND(array_of_pay_dates)                     'loop through the array of the pay dates only'
                                    'once the pay date in the income array matches the one in the chronological list of dates, use the index number to set an order code within the list of income array
                                    'MsgBox "Look at each index: " & index
                                    If array_of_pay_dates(index) = LIST_OF_INCOME_ARRAY(pay_date, all_income) Then
                                        LIST_OF_INCOME_ARRAY(check_order, all_income) = index + 1

                                    End If
                                    top_of_order = index + 1    'this identifies how many pay dates there are in for this panel
                                next
                            End If
                        Next
                        EARNED_INCOME_PANELS_ARRAY(order_ubound, ei_panel) = top_of_order   'setting the number of unique pay dates within the panel array because we need it for sorting correctly

                        'MsgBox top_of_order
                        ' 'this part actually looks at the income information IN ORDER
                        ' For order_number = 1 to top_of_order                        'loop through the order number lowest to highest
                        '     For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                        '         'conditional if it is the right panel AND the order matches - then do the thing you need to do
                        '         If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(check_order, all_income) = order_number Then
                        '             list_of_dates = list_of_dates & vbNewLine & "Check Date: " & LIST_OF_INCOME_ARRAY(pay_date, all_income) & " Income: $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " Hours: " & LIST_OF_INCOME_ARRAY(hours, all_income)
                        '         End If
                        '     next
                        ' next
                        ' MsgBOx list_of_dates

                        issues_with_frequency = FALSE
                        For order_number = 1 to top_of_order                        'loop through the order number lowest to highest
                            For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                'conditional if it is the right panel AND the order matches - then do the thing you need to do
                                If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(check_order, all_income) = order_number Then
                                    If EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel) = "" Then EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel) = LIST_OF_INCOME_ARRAY(pay_date, all_income)
                                    list_of_dates = list_of_dates & vbNewLine & "Check Date: " & LIST_OF_INCOME_ARRAY(pay_date, all_income) & " Income: $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " Hours: " & LIST_OF_INCOME_ARRAY(hours, all_income)

                                    If prev_date <> "" Then
                                        days_between_checks = DateDiff("d", prev_date, LIST_OF_INCOME_ARRAY(pay_date, all_income))

                                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "1 - One Time Per Month" Then
                                            If days_between_checks < 28 or days_between_checks > 31 Then issues_with_frequency = TRUE
                                        ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "2 - Two Times Per Month" Then
                                            If days_between_checks < 14 or days_between_checks > 17 Then issues_with_frequency = TRUE
                                        ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "3 - Every Other Week" Then
                                            If days_between_checks <> 14 Then issues_with_frequency = TRUE
                                        ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "4 - Every Week" Then
                                            If days_between_checks <> 7 Then issues_with_frequency = TRUE
                                        ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "5 - Other" Then

                                        Else
                                            If days_between_checks = 7 Then
                                                EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "4 - Every Week"
                                            ElseIf days_between_checks = 14 Then
                                                EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "3 - Every Other Week"
                                            ElseIf days_between_checks >= 14 AND days_between_checks <= 19 Then
                                                EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "2 - Two Times Per Month"
                                            ElseIf days_between_checks >= 28 AND days_between_checks <= 31 Then
                                                EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "1 - One Time Per Month"
                                            End If

                                        End If
                                    Else

                                    End If
                                    prev_date = LIST_OF_INCOME_ARRAY(pay_date, all_income)
                                End If
                            next
                        next

                        If issues_with_frequency = FALSE Then
                            If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "3 - Every Other Week" OR EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "4 - Every Week" Then
                                For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                                    If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then
                                        EARNED_INCOME_PANELS_ARRAY(pay_weekday, ei_panel) = WeekDayName(Weekday(LIST_OF_INCOME_ARRAY(pay_date, all_income)))
                                        'MsgBox "Look at the payday: " & EARNED_INCOME_PANELS_ARRAY(pay_weekday, ei_panel)
                                        Exit For
                                    End If
                                Next
                            End If
                        End If

                        '--------------------------------------'

                        For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                            If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then
                                cash_checks = cash_checks + 1
                                EARNED_INCOME_PANELS_ARRAY(income_list_indct, ei_panel) = EARNED_INCOME_PANELS_ARRAY(income_list_indct, ei_panel) & "~" & all_income

                                If LIST_OF_INCOME_ARRAY(budget_in_SNAP_no, all_income) = unchecked Then LIST_OF_INCOME_ARRAY(budget_in_SNAP_yes, all_income) = checked
                                If LIST_OF_INCOME_ARRAY(budget_in_SNAP_yes, all_income) = checked Then
                                    If LIST_OF_INCOME_ARRAY(exclude_amount, all_income) = "" Then LIST_OF_INCOME_ARRAY(exclude_amount, all_income) = 0
                                    LIST_OF_INCOME_ARRAY(exclude_amount, all_income) = LIST_OF_INCOME_ARRAY(exclude_amount, all_income) * 1
                                    LIST_OF_INCOME_ARRAY(gross_amount, all_income) = LIST_OF_INCOME_ARRAY(gross_amount, all_income) * 1
                                    net_amount = LIST_OF_INCOME_ARRAY(gross_amount, all_income) - LIST_OF_INCOME_ARRAY(exclude_amount, all_income)
                                    total_of_counted_income = total_of_counted_income + net_amount
                                    total_of_included_pay_checks = total_of_included_pay_checks +  LIST_OF_INCOME_ARRAY(gross_amount, all_income)
                                    number_of_checks_budgeted = number_of_checks_budgeted + 1

                                    LIST_OF_INCOME_ARRAY(hours, all_income) = LIST_OF_INCOME_ARRAY(hours, all_income) * 1
                                    total_of_hours = total_of_hours + LIST_OF_INCOME_ARRAY(hours, all_income)
                                Else
                                    list_of_excluded_pay_dates = list_of_excluded_pay_dates & ", " & LIST_OF_INCOME_ARRAY(pay_date, all_income)
                                End If
                            End If
                        Next

                        'ONCE PAY FREQUENCY IS DETERMINED, write to assess if paystubs consititute 30 days and if not, force clarification of income.
                        EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel) = total_of_hours / number_of_checks_budgeted
                        EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel) = FormatNumber(EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel), 2,,0)

                        EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel) = total_of_counted_income / number_of_checks_budgeted
                        EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel) = FormatNumber(EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel),2,,0)

                        EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel) = total_of_included_pay_checks / total_of_hours
                        EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel) = FormatNumber(EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel),2,,0)

                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "1 - One Time Per Month" Then EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) = EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel)/4.3
                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "2 - Two Times Per Month" Then EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) = (EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel)*2)/4.3
                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "3 - Every Other Week" Then EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) = EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel)/2
                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "4 - Every Week" Then EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) = EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel)
                        EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) = FormatNumber(EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel), 2,,0)

                        If list_of_excluded_pay_dates <> "" Then list_of_excluded_pay_dates = right(list_of_excluded_pay_dates, len(list_of_excluded_pay_dates) - 2)
                        If issues_with_frequency = FALSE AND EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) <> "" Then
                            pay_multiplier = 0
                            If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "1 - One Time Per Month" Then pay_multiplier = 1
                            If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "2 - Two Times Per Month" Then pay_multiplier = 2
                            If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "3 - Every Other Week" Then pay_multiplier = 2.15
                            If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "4 - Every Week" Then pay_multiplier = 4.3
                            EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel) = pay_multiplier * EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel)

                        End If
                        If EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel) = "" OR EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel) = 0 THen EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel) = "?"
                        EARNED_INCOME_PANELS_ARRAY(income_list_indct, ei_panel) = right(EARNED_INCOME_PANELS_ARRAY(income_list_indct, ei_panel), len(EARNED_INCOME_PANELS_ARRAY(income_list_indct, ei_panel))-1)


                    End If

                    If EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_estimate Then

                        using_30_days = TRUE
                        paycheck_list_title = "Anticipated Paychecks for " & EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel) & "/" & EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel) & ":"

                        ' EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel) = trim(EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel))
                        ' EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) = trim(EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel))
                        ' EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = trim(EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel))
                        '
                        ' Text 185, y_pos + 10, 200, 10, "Average hourly rate of pay: $" & EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel)
                        ' Text 185, y_pos + 25, 200, 10, "Average weekly hours: " & EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel)
                        ' Text 185, y_pos + 40, 200, 10, "Average paycheck amount: $" & EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel)
                        ' Text 185, y_pos + 55, 200, 10, "Monthly Budgeted Income: $" & EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel)
                        '
                        '
                        ' ""+chr(9)+"1 - One Time Per Month"+chr(9)+"2 - Two Times Per Month"+chr(9)+"3 - Every Other Week"+chr(9)+"4 - Every Week", EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel)

                        the_first_of_CM_2 = CM_plus_2_mo & "/1/" & CM_plus_2_yr
                        CM_2_mo = DatePart("m", the_first_of_CM_2)
                        CM_2_yr = DatePart("yyyy", the_first_of_CM_2)
                        the_initial_month = DateValue(EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel) & "/1/" & EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel))

                        EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel) = EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel)

                        days_to_add = 0
                        months_to_add = 0
                        Select Case EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel)
                            Case "1 - One Time Per Month"
                                EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel) = EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel) * EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel)
                                EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel) = EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel)
                                EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel) = EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) * 4.3
                                days_to_add = 0
                                months_to_add = 1
                                default_start_date = the_initial_month
                            Case "2 - Two Times Per Month"
                                EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel) = EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel) * EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) * 4.3 / 2
                                EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel) = EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel) * 2
                                EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel) = (EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) * 4.3)/2
                                days_to_add = 15
                                months_to_add = 1
                                default_start_date = the_initial_month
                            Case "3 - Every Other Week"
                                EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel) = EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel) * EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) * 2
                                EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel) = EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel) * 2.15
                                EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel) = EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) * 2
                                days_to_add = 14
                                months_to_add = 0
                                the_date_of_week = the_initial_month
                                Do
                                    If Weekday(the_date_of_week) = vbFriday Then
                                        default_start_date = the_date_of_week
                                        Exit Do
                                    Else
                                        the_date_of_week = DateAdd("d", 1, the_date_of_week)
                                    End If
                                Loop
                            Case "4 - Every Week"
                                EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel) = EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel) * EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel)
                                EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel) = EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel) * 4.3
                                EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel) = EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel)
                                days_to_add = 7
                                months_to_add = 0
                                the_date_of_week = the_initial_month
                                Do
                                    If Weekday(the_date_of_week) = vbFriday Then
                                        default_start_date = the_date_of_week
                                        Exit Do
                                    Else
                                        the_date_of_week =DateAdd("d", 1, the_date_of_week)
                                    End If
                                Loop
                        End Select
                        ' MsgBox "Default start date - "& default_start_date

                        snap_anticipated_pay_array = ""
                        checks_list = ""
                        ' list_of_actual_paydates
                        'Trying to figure out the ACTUAL pay dates.
                        Call Navigate_to_MAXIS_screen("STAT", "JOBS")
                        EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel), 20, 76
                        EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel), 20, 79
                        transmit
                        this_pay_date = ""
                        If list_of_actual_paydates <> "" Then
                            paydates_array = split(list_of_actual_paydates, "~")
                            this_pay_date = paydates_array(0)
                            If EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel) = "" Then EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel) = this_pay_date

                        ElseIf known_pay_date <> "" Then
                            this_pay_date = DateAdd("d", 0, known_pay_date)
                        Else
                            EMReadScreen this_pay_date, 8, 12, 25   'first check on retro side
                            If this_pay_date = "__ __ __" Then
                                this_pay_date = ""
                            Else
                                this_pay_date = replace(this_pay_date, " ", "/")
                                this_pay_date = DateValue(this_pay_date)
                            End If
                        End If
                        If this_pay_date = "" Then this_pay_date = default_start_date
                        save_dates = FALSE
                        'MsgBox "CM plus 2 - " & CM_2_mo & "/" & CM_2_yr
                        'MsgBox this_pay_date
                        'MsgBox "Initial Month " & the_initial_month
                        Do 'While DatePart("m", this_pay_date) <> CM_2_mo AND DatePart("yyyy", this_pay_date) <> CM_2_yr
                            ' MsgBox this_pay_date
                            save_dates = FALSE
                            If DatePart("m", this_pay_date) = DatePart("m", the_initial_month) AND DatePart("yyyy", this_pay_date) = DatePart("yyyy", the_initial_month) Then save_dates = TRUE
                            If save_dates = TRUE Then
                                'MsgBox "SAVE - " & this_pay_date
                                If EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel) = "" Then EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel) = this_pay_date

                                check_found = FALSE
                                For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                                    If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then
                                        If DateValue(LIST_OF_INCOME_ARRAY(pay_date, all_income)) = this_pay_date Then
                                            check_found = TRUE
                                            check_number = all_income
                                            Exit For
                                        End If
                                    End If
                                Next
                                If check_found = TRUE Then
                                    If len(this_pay_date) = 10 Then checks_list = checks_list & "%" & this_pay_date & "   ~   $" & LIST_OF_INCOME_ARRAY(gross_amount, check_number)
                                    If len(this_pay_date) = 9 Then checks_list = checks_list & "%" & this_pay_date & "    ~   $" & LIST_OF_INCOME_ARRAY(gross_amount, check_number)
                                    If len(this_pay_date) = 8 Then checks_list = checks_list & "%" & this_pay_date & "     ~   $" & LIST_OF_INCOME_ARRAY(gross_amount, check_number)
                                Else
                                    If len(this_pay_date) = 10 Then checks_list = checks_list & "%" & this_pay_date & "   ~   $" & EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel)
                                    If len(this_pay_date) = 9 Then checks_list = checks_list & "%" & this_pay_date & "    ~   $" & EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel)
                                    If len(this_pay_date) = 8 Then checks_list = checks_list & "%" & this_pay_date & "     ~   $" & EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel)
                                End If
                            End If
                            If months_to_add = 0 Then
                                this_pay_date = DateAdd("d", days_to_add, this_pay_date)
                            ElseIf days_to_add = 0 Then
                                this_pay_date = DateAdd("m", months_to_add, this_pay_date)
                            Else
                                checks_list = checks_list & "%" & DateAdd("d", days_to_add, this_pay_date) & " ~ $" & EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel)
                                this_pay_date = DateAdd("m", months_to_add, this_pay_date)

                            End If
                            'MsgBox "NEXT - " & this_pay_date
                        Loop until DatePart("m", this_pay_date) = CM_2_mo AND DatePart("yyyy", this_pay_date) = CM_2_yr

                        If left(checks_list, 1) = "%" Then checks_list = right(checks_list, len(checks_list)-1)
                        If InStr(checks_list, "%") <> 0 Then
                            snap_anticipated_pay_array = Split(checks_list,"%")
                        Else
                            snap_anticipated_pay_array = Array(checks_list)
                        End If
                    End If

                    'Script will determine pay frequency and potentially 1st check (if not listed on JOBS)
                    'Script will determine the initial footer month to change by the pay dates listed.
                    'Script will create a budget based on the program this income applies to
                    'Dialog the budget and have the worker confirm - if they decline - pull the check list dialog back up and have them adjust it there.

                    ' 'If we are applying to cash, we need to look at each month of paychecks to see if the month is prospective or retrospective
                    ' If EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, ei_panel) = checked Then
                    '
                    '     MAXIS_footer_month = EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel)
                    '     MAXIS_footer_year = EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel)
                    '
                    '     Call navigate_to_MAXIS_screen("STAT", "PROG")
                    '     EMReadScreen cash_one_status, 4, 6, 74
                    '     If cash_one_status = "ACTV" or cash_one_status = "PEND" Then
                    '         EMReadScreen cash_prog, 2, 6, 67
                    '     Else
                    '         EMReadScreen cash_two_status, 4, 7, 74
                    '         If cash_two_status = "ACTV" or cash_two_status = "PEND" Then
                    '             EMReadScreen cash_prog, 2, 7, 67
                    '         End If
                    '     End If
                    '     the_cash = cash_prog
                    '     If the_cash = "MS" Then the_cash = "MSA"
                    '
                    '     Call back_to_self
                    '
                    '     list_of_months = "~"
                    '     pay_month = MAXIS_footer_month
                    '     pay_year = MAXIS_footer_year
                    '
                    '     Do
                    '         list_of_months = list_of_months & pay_month & "/" & pay_year & "~"
                    '
                    '         pay_month = (pay_month * 1) + 1
                    '         If pay_month = 13 Then
                    '             pay_month = "01"
                    '             pay_year = right("00"&((pay_year * 1) + 1), 2)
                    '         Else
                    '             pay_month = right("00"&pay_month, 2)
                    '         End If
                    '     Loop Until pay_month = CM_plus_2_mo AND pay_year = CM_plus_2_yr
                    '
                    '     Dim CASH_MONTHS_ARRAY()
                    '     ReDim CASH_MONTHS_ARRAY(2, 0)
                    '
                    '     If left(list_of_months, 1) = "~" Then list_of_months = right(list_of_months, len(list_of_months)-1)
                    '     If right(list_of_months, 1) = "~" Then list_of_months = left(list_of_months, len(list_of_months)-1)
                    '     each_month = 0
                    '     'MsgBox list_of_months
                    '
                    '     If InStr(list_of_months, "~") <> 0 Then
                    '         array_of_months = split(list_of_months, "~")
                    '         For each elig_thing in array_of_months
                    '             ReDim Preserve CASH_MONTHS_ARRAY(2, each_month)
                    '             CASH_MONTHS_ARRAY(cash_mo_yr, each_month) = elig_thing
                    '             CASH_MONTHS_ARRAY(update_y_n, each_month) = checked
                    '             each_month = each_month + 1
                    '         Next
                    '     Else
                    '         CASH_MONTHS_ARRAY(cash_mo_yr, each_month) = list_of_months
                    '         CASH_MONTHS_ARRAY(update_y_n, each_month) = checked
                    '     End If
                    '
                    '     If the_cash <> "" Then
                    '         For update_month = 0 to UBOUND(CASH_MONTHS_ARRAY, 2)
                    '             MAXIS_footer_month = left(CASH_MONTHS_ARRAY(cash_mo_yr, update_month), 2)
                    '             MAXIS_footer_year = right(CASH_MONTHS_ARRAY(cash_mo_yr, update_month), 2)
                    '
                    '             Call back_to_SELF
                    '
                    '             Call navigate_to_MAXIS_screen("ELIG", the_cash)
                    '             If the_cash = "MF" Then
                    '                 EmWriteScreen "MFSM", 20, 71
                    '                 transmit
                    '
                    '                 EMReadScreen type_of_budget, 5, 12, 31
                    '             ElseIf the_cash = "DW" Then
                    '                 type_of_budget = "PROSP"
                    '
                    '             ElseIf the_cash = "MSA" Then
                    '                 EmWriteScreen "MSSM", 20, 71
                    '                 transmit
                    '
                    '                 EMReadScreen type_of_budget, 5, 13, 29
                    '             ElseIf the_cash = "GA" Then
                    '                 EmWriteScreen "GASM", 20, 71
                    '                 transmit
                    '
                    '                 EMReadScreen type_of_budget, 5, 12, 32
                    '             End If
                    '             CASH_MONTHS_ARRAY(budget_cycle, update_month) = type_of_budget
                    '         Next
                    '     End If
                    ' End If

                    word_for_freq = ""
                    If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "1 - One Time Per Month" Then word_for_freq = "monthly"
                    If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "2 - Two Times Per Month" Then word_for_freq = "semi-monthly"
                    If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "3 - Every Other Week" Then word_for_freq = "biweekly"
                    If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "4 - Every Week" Then word_for_freq = "weekly"

                    dlg_len = 65
                    If EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) = checked Then
                        dlg_len = dlg_len + 80
                        dlg_len = dlg_len + number_of_checks_budgeted*10
                        If number_of_checks_budgeted < 4 THen dlg_len = 180
                        If using_30_days = FALSE Then dlg_len = dlg_len + 30
                    End If
                    If EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, ei_panel) = checked Then
                        dlg_len = dlg_len + 45
                        cash_grp_len = 60
                        length_of_checks_list = cash_checks*10

                        dlg_len = dlg_len + length_of_checks_list
                        cash_grp_len = cash_grp_len + length_of_checks_list

                    End If

                    y_pos = 25
                    BeginDialog Dialog1, 0, 0, 421, dlg_len, "Confirm JOBS Budget"
                      Text 10, 10, 175, 10, "JOBS " & EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel) & " " & EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel) & " - " & EARNED_INCOME_PANELS_ARRAY(employer, ei_panel)
                      Text 245, 10, 50, 10, "Pay Frequency"
                      DropListBox 305, 5, 95, 45, ""+chr(9)+"1 - One Time Per Month"+chr(9)+"2 - Two Times Per Month"+chr(9)+"3 - Every Other Week"+chr(9)+"4 - Every Week"+chr(9)+"5 - Other", EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel)
                      ' Text 240, 30, 60, 10, "Income Start Date:"
                      ' EditBox 305, 25, 70, 15, income_start_date
                      If EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) = checked Then
                          grp_len = 100 + number_of_checks_budgeted*10
                          If number_of_checks_budgeted < 4 Then grp_len = 130
                          If using_30_days = FALSE Then grp_len = grp_len + 30
                          GroupBox 5, y_pos, 410, grp_len, "SNAP Budget"

                          Text 10, y_pos + 10, 400, 10, "Income provided covers the period " & first_date & " to " & last_date & ". This income covers " & spread_of_pay_dates & " days."
                          If using_30_days = FALSE Then
                              y_pos = y_pos + 25
                              Text 10, y_pos + 5, 175, 10, "It appears this is not 30 days of income. Explain:"
                              EditBox 185, y_pos, 200, 15, not_30_explanation
                          End If

                          y_pos = y_pos + 10

                          'GroupBox 5, y_pos, 410, 75 + number_of_checks_budgeted*10, "SNAP Budget"
                          Text 10, y_pos + 10, 150, 10, paycheck_list_title        '"Paychecks Inclued in Budget:"'
                          list_pos = 0

                          If EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_actual Then
                              ' 'this part actually looks at the income information IN ORDER
                              For order_number = 1 to top_of_order                        'loop through the order number lowest to highest
                                  For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                      'conditional if it is the right panel AND the order matches - then do the thing you need to do
                                      If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(check_order, all_income) = order_number Then
                                          'list_of_dates = list_of_dates & vbNewLine & "Check Date: " & LIST_OF_INCOME_ARRAY(pay_date, all_income) & " Income: $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " Hours: " & LIST_OF_INCOME_ARRAY(hours, all_income)
                                          If LIST_OF_INCOME_ARRAY(exclude_amount, all_income) = 0 Then Text 20, (list_pos * 10) + y_pos + 25, 90, 10, LIST_OF_INCOME_ARRAY(pay_date, all_income) & " - $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & "hrs."
                                          If LIST_OF_INCOME_ARRAY(exclude_amount, all_income) <> 0 Then Text 20, (list_pos * 10) + y_pos + 25, 90, 10, LIST_OF_INCOME_ARRAY(pay_date, all_income) & " - $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & "hrs. - $" & LIST_OF_INCOME_ARRAY(exclude_amount, all_income) & " not included."

                                          list_pos = list_pos + 1
                                      End If
                                  next
                              next
                          ElseIf EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_estimate Then
                              For each money_day in snap_anticipated_pay_array
                                  ' MsgBox money_day
                                  Text 20, (list_pos * 10) + y_pos + 25, 90, 10, money_day
                                  list_pos = list_pos + 1
                              Next
                          End If
                          If list_pos < 3 Then list_pos = 3
                          ' MsgBOx list_of_dates
                          ' For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                          '     If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then
                          '       If LIST_OF_INCOME_ARRAY(budget_in_SNAP_yes, all_income) = checked Then
                          '         Text 20, (list_pos * 10) + 65 + 25, 90, 10, LIST_OF_INCOME_ARRAY(pay_date, all_income) & " - $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & "hrs."
                          '         list_pos = list_pos + 1
                          '         'Text 20, 65, 90, 10, "01/01/2018 - $400 - 40 hrs"
                          '         'Text 20, 75, 90, 10, "01/15/2018- $400 - 40 hrs"
                          '       End If
                          '     End If
                          ' Next

                          Text 185, y_pos + 10, 200, 10, "Average hourly rate of pay: $" & EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel)
                          Text 185, y_pos + 25, 200, 10, "Average " & word_for_freq & " hours: " & EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel)
                          Text 185, y_pos + 40, 200, 10, "Average paycheck amount: $" & EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel)
                          Text 185, y_pos + 55, 200, 10, "Monthly Budgeted Income: $" & EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel)
                          If EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel) = "?" Then
                            ButtonGroup ButtonPressed
                                PushButton 305, y_pos + 55, 60, 10, "Calculate", calc_btn
                          End If

                          y_pos = y_pos + (list_pos * 10) + 35
                          Text 10, y_pos, 400, 10, "Paychecks not included: " & list_of_excluded_pay_dates

                          CheckBox 10, y_pos + 15, 330, 10, "Check here if you confirm that this budget is correct and is the best estimate of anticipated income.", confirm_budget_checkbox
                          Text 10, y_pos + 35, 60, 10, "Conversation with:"
                          ComboBox 75, y_pos + 30, 60, 45, " "+chr(9)+"Client - not employee"+chr(9)+"Employee"+chr(9)+"Employer",  EARNED_INCOME_PANELS_ARRAY(spoke_with, ei_panel)
                          Text 140, y_pos + 35, 25, 10, "clarifies"
                          EditBox 170, y_pos + 30, 235, 15, EARNED_INCOME_PANELS_ARRAY(convo_detail, ei_panel)
                          y_pos = y_pos + 55
                      Else
                        confirm_budget_checkbox = checked
                        using_30_days = TRUE
                      End If
                      'TODO deal with cash stuff - need to address retro/prosp and change this dialog to only show cash/snap if the income applies to that.
                      If EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, ei_panel) = checked Then
                          GroupBox 5, y_pos, 410, cash_grp_len, "CASH Budget"
                          y_pos = y_pos + 15
                          Text 10, y_pos, 400, 10, "Pay information will be entered on the RETRO side if provided. The script will not calculate an average for any Retro pay dates."
                          Text 10, y_pos + 10, 410, 10, "For each month to be updated, the script will use actual pay information or the average for that month on the prospective side."
                          y_pos = y_pos + 30
                          Text 10, y_pos, 30, 10, "CHECKS"

                          ' Text 150, y_pos, 30, 10, "MONTH"
                          ' Text 190, y_pos, 35, 10, "BUDGET"
                          ' Text 240, y_pos, 40, 10, "UPDATE"

                          y_pos = y_pos + 10

                          CheckBox 150, y_pos, 250, 10, "Check here if these checks are accurate and should be entered.", confirm_checks_checkbox


                          list_pos = 0
                          ' 'this part actually looks at the income information IN ORDER
                          For order_number = 1 to top_of_order                        'loop through the order number lowest to highest
                              For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                  'conditional if it is the right panel AND the order matches - then do the thing you need to do
                                  If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(check_order, all_income) = order_number Then
                                      'list_of_dates = list_of_dates & vbNewLine & "Check Date: " & LIST_OF_INCOME_ARRAY(pay_date, all_income) & " Income: $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " Hours: " & LIST_OF_INCOME_ARRAY(hours, all_income)
                                      Text 20, (list_pos * 10) + y_pos, 90, 10, LIST_OF_INCOME_ARRAY(pay_date, all_income) & " - $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & "hrs."
                                      list_pos = list_pos + 1
                                  End If
                              next
                          next

                          ' mo_list = 0
                          ' For each_month = 0 to UBOUND(CASH_MONTHS_ARRAY, 2)
                          '     Text          150, (mo_list*15) + y_pos + 5,  20, 10, CASH_MONTHS_ARRAY(cash_mo_yr, each_month)
                          '     DropListBox   190, (mo_list*15) + y_pos,      40, 45, " "+chr(9)+"Retro"+chr(9)+"Prosp", CASH_MONTHS_ARRAY(budget_cycle, each_month)
                          '     CheckBox      240, (mo_list*15) + y_pos + 5,  75, 10, "Update this month", CASH_MONTHS_ARRAY(update_y_n, each_month)
                          '     mo_list = mo_list + 1
                          '     'y_pos = y_pos + 15
                          ' Next

                          bottom_of_checks = y_pos + (list_pos * 10)
                          y_pos = bottom_of_checks + 10

                          'y_pos = y_pos + 10
                      Else
                        confirm_checks_checkbox = checked
                      End If
                      ButtonGroup ButtonPressed
                        OkButton 315, y_pos, 50, 15
                        CancelButton 365, y_pos, 50, 15
                    EndDialog


                    Dialog Dialog1
                    cancel_confirmation

                    If confirm_budget_checkbox = unchecked then
                        big_err_msg = big_err_msg & vbNewLine & "*** Since the budget is not confirmed as correct, the ENTER PAY INFORMATION DIALOG will reappear and allow information to be corrected to generate an accurate budget. ***"
                        review_small_dlg = TRUE
                    End If
                    If confirm_checks_checkbox = unchecked Then
                        big_err_msg = big_err_msg & vbNewLine & "*** If the checks are not accurate, review them and update as necessary. ***"
                        review_small_dlg = TRUE
                    End If
                    If using_30_days = FALSE Then
                        If not_30_explanation = "" Then big_err_msg = big_err_msg & vbNewLine & "** Since income received is not 30 days of income for SNAP, it must be explained why we are accepting more or less."
                    End If
                    If ButtonPressed = calc_btn Then
                        review_small_dlg = FALSE
                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "" Then
                            big_err_msg = big_err_msg & vbNewLine & "** List the pay frequency for this income."
                        Else
                            big_err_msg = "LOOP" & big_err_msg
                        End If
                    End If

                    If big_err_msg <> "" Then
                        For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                            If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then
                                If LIST_OF_INCOME_ARRAY(budget_in_SNAP_yes, all_income) = checked Then

                                    LIST_OF_INCOME_ARRAY(gross_amount, all_income) = LIST_OF_INCOME_ARRAY(gross_amount, all_income) & ""
                                    LIST_OF_INCOME_ARRAY(hours, all_income) = LIST_OF_INCOME_ARRAY(hours, all_income) & ""

                                End If
                            End If
                        Next
                        If left(big_err_msg, 4) <> "LOOP" Then MsgBox "Review JOBS Pay Information" & vbNewLine & big_err_msg
                    End If

                Loop until big_err_msg = ""
                call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
            LOOP UNTIL are_we_passworded_out = false


            If EARNED_INCOME_PANELS_ARRAY(cash_mos_list, ei_panel) <> "" Then EARNED_INCOME_PANELS_ARRAY(cash_mos_list, ei_panel) = Join(CASH_MONTHS_ARRAY, "~")
            If EARNED_INCOME_PANELS_ARRAY(antic_pay_list, ei_panel) <> "" Then EARNED_INCOME_PANELS_ARRAY(antic_pay_list, ei_panel) = Join(snap_anticipated_pay_array, "%*%")
            If EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_actual Then
                EARNED_INCOME_PANELS_ARRAY(days_of_verif, ei_panel) = "Pay verifications covers the period " & first_date & " to " & last_date & " which is " & spread_of_pay_dates & " days. "
                If using_30_days = FALSE Then
                    EARNED_INCOME_PANELS_ARRAY(days_of_verif, ei_panel) = EARNED_INCOME_PANELS_ARRAY(days_of_verif, ei_panel) & "This is not 30 days, we are not using 30 days because: " & not_30_explanation
                End If
            End If
            'Add handling to check WREG for correct coding based upon this income information (potentiall update)

            'Worker must confirm the frequency, first pay, and footer month
            'Worker will inicate if future months should be updated - default this to 'yes' as script will update retro and prospective specific to each month
            'SNAP PIC, GRH PIC, HC EI EST will be checked to be updated IF any of these programs are open on the case.

            'NEED to add handling for future/current changes - start or stop work - get policy on this from SNAP refresher - talk to Melissa.
        End If
    End If

    If EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel) = "BUSI" Then
        'NAVIGATE to BUSI for each HH MEMBER and ask if Income Information was received for this Self Employment.
        Call Navigate_to_MAXIS_screen("STAT", "BUSI")
        EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel), 20, 76
        EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel), 20, 79
        transmit

        employer_check = MsgBox("Do you have income verification for this self employment? Type of Self Employment: " & EARNED_INCOME_PANELS_ARRAY(income_type, ei_panel), vbYesNo + vbQuestion, "Select Income Panel")

        ' If EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel) = "" Then EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel) = MAXIS_footer_month
        ' If EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel) = "" Then EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel) = MAXIS_footer_year
        ' EARNED_INCOME_PANELS_ARRAY(update_futue_chkbx, ei_panel) = future_months_check

        If employer_check = vbYes Then
            EARNED_INCOME_PANELS_ARRAY(income_received, ei_panel) = TRUE
            Do
                Do
                    big_err_msg  ""

                    basic_info_gathered = FALSE
                    Do
                        Do
                            'MsgBox "Basic Info Gathered - " & basic_info_gathered & vbNewLine & "Ready Error Message: " & ready_err_msg
                            sm_err_msg = ""
                            ready_err_msg = ""

                            dlg_len = 80
                            If basic_info_gathered = TRUE Then
                                If EARNED_INCOME_PANELS_ARRAY(self_emp_mthd, ei_panel) = "01 - 50% Grosss Inc" Then
                                    dlg_len = (dlg_factor * 20) + 125
                                End If
                                If EARNED_INCOME_PANELS_ARRAY(self_emp_mthd, ei_panel) = "02 - Tax Forms" Then dlg_len = 125
                            End If
                            'MsgBox "Dialog Length: " & dlg_len
                            BeginDialog Dialog1, 0, 0, 486, dlg_len, "Enter Self Employment Information"
                              Text 10, 10, 180, 10, EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel) & " " & EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel) & " " & EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel)  ''"BUSI 01 01 - CLIENT NAME"
                              Text 200, 10, 80, 10, "Self Employment Type:"
                              DropListBox 280, 5, 125, 45, "01 - Farming"+chr(9)+"02 - Real Estate"+chr(9)+"03 - Home Product Sales"+chr(9)+"04 - Other Sales"+chr(9)+"05 - Personal Services"+chr(9)+"06 - Paper Route"+chr(9)+"07 - In Home Daycare"+chr(9)+"08 - Rental Income"+chr(9)+"09 - Other", EARNED_INCOME_PANELS_ARRAY(income_type, ei_panel)
                              Text 10, 30, 65, 10, "Verification srouce:"
                              DropListBox 90, 25, 75, 45, " "+chr(9)+"1 - Income Tax Returns"+chr(9)+"2 - Receipts of Sales/Purch"+chr(9)+"3 - Client Busi Records/Ledger"+chr(9)+"6 - Other Document"+chr(9)+"N - No Ver Prvd", EARNED_INCOME_PANELS_ARRAY(income_verif, ei_panel)
                              'QUESTION - do we need to add this option back in so that the way income is reported in is independent of the budgeting method'
                              'Text 180, 30, 100, 10, "Amount of Income Information:"
                              'DropListBox 290, 25, 80, 45, "Select One..."+chr(9)+"A Full Year Totaled"+chr(9)+"Month by Month", amount_income
                              Text 10, 50, 120, 10, "Self Employment Budgeting Method"
                              DropListBox 135, 45, 85, 45, " "+chr(9)+"01 - 50% Grosss Inc"+chr(9)+"02 - Tax Forms", EARNED_INCOME_PANELS_ARRAY(self_emp_mthd, ei_panel)
                              Text 225, 50, 50, 10, "Selection Date:"
                              EditBox 280, 45, 50, 15, EARNED_INCOME_PANELS_ARRAY(method_date, ei_panel)
                              CheckBox 30, 65, 210, 10, "Check here to confirm this method was discussed with Client.", EARNED_INCOME_PANELS_ARRAY(self_emp_mthd_conv, ei_panel)
                              GroupBox 415, 5, 65, 70, "Apply Income To"
                              CheckBox 425, 20, 35, 10, "SNAP", EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel)
                              CheckBox 425, 35, 35, 10, "CASH", EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, ei_panel)
                              CheckBox 425, 50, 25, 10, "HC", EARNED_INCOME_PANELS_ARRAY(apply_to_HC, ei_panel)
                              ButtonGroup ButtonPressed
                                PushButton 355, 60, 50, 15, "Ready", open_button
                              If basic_info_gathered = TRUE Then
                                  Text 330, 90, 50, 10, "Reported Hours"
                                  EditBox 385, 85, 30, 15, numb_hrs_reptd
                                  DropListBox 420, 85, 60, 15, ""+chr(9)+"per week"+chr(9)+"per month", hours_rate
                                  If EARNED_INCOME_PANELS_ARRAY(self_emp_mthd, ei_panel) = "01 - 50% Grosss Inc" Then
                                      Text 10, 90, 55, 10, "Month and Year"
                                      Text 70, 90, 50, 10, "Gross Income"
                                      Text 130, 80, 90, 10, "Exclude from SNAP Budget"
                                      Text 130, 90, 30, 10, "Amount"
                                      Text 190, 90, 30, 10, "Reason"
                                      y_pos = 0
                                      For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                                          If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then
                                              EditBox 10, (y_pos * 20) + 105, 40, 15, LIST_OF_INCOME_ARRAY(pay_date, all_income)
                                              EditBox 70, (y_pos * 20) + 105, 50, 15, LIST_OF_INCOME_ARRAY(gross_amount, all_income)
                                              EditBox 130, (y_pos * 20) + 105, 50, 15, LIST_OF_INCOME_ARRAY(exclude_amount, all_income)
                                              EditBox 190, (y_pos * 20) + 105, 290, 15, LIST_OF_INCOME_ARRAY(reason_amt_excluded, all_income)
                                              y_pos = y_pos + 1
                                          End If
                                      Next
                                      ButtonGroup ButtonPressed
                                        PushButton 320, (y_pos * 20) + 105, 15, 15, "+", plus_button
                                        PushButton 340, (y_pos * 20) + 105, 15, 15, "-", minus_button
                                        y_pos = (y_pos * 20) + 105

                                  ElseIf EARNED_INCOME_PANELS_ARRAY(self_emp_mthd, ei_panel) = "02 - Tax Forms" Then
                                      Text 10, 90, 35, 10, "Tax Year"
                                      Text 60, 80, 35, 20, "Months in Business"
                                      Text 110, 90, 30, 10, "Income"
                                      Text 155, 90, 35, 10, "Expenses"
                                      EditBox 10, 105, 40, 15, tax_year
                                      DropListBox 60, 105, 40, 45, "12"+chr(9)+"11"+chr(9)+"10"+chr(9)+"9"+chr(9)+"8"+chr(9)+"7"+chr(9)+"6"+chr(9)+"5"+chr(9)+"4"+chr(9)+"3"+chr(9)+"2"+chr(9)+"1", months_covered
                                      EditBox 110, 105, 40, 15, tax_income
                                      EditBox 155, 105, 40, 15, tax_expenses

                                      y_pos = 105
                                  End If
                                  ButtonGroup ButtonPressed
                                    'PushButton 320, 155, 15, 15, "+", plus_button
                                    'PushButton 340, 155, 15, 15, "-", minus_button
                                    OkButton 375, y_pos, 50, 15
                                    CancelButton 430, y_pos, 50, 15
                              End If
                            EndDialog

                            Dialog Dialog1
                            cancel_confirmation

                            If buttonpressed = open_button Then
                                basic_info_gathered = TRUE
                                If trim(EARNED_INCOME_PANELS_ARRAY(income_type, ei_panel)) = "" Then ready_err_msg = ready_err_msg & vbNewLine & "* Indicate the TYPE of self employment income."
                                If trim(EARNED_INCOME_PANELS_ARRAY(income_verif, ei_panel)) = "" Then ready_err_msg = ready_err_msg & vbNewLine & "* List the verification received for this income."
                                If trim(EARNED_INCOME_PANELS_ARRAY(self_emp_mthd, ei_panel)) = "" Then ready_err_msg = ready_err_msg & vbNewLine & "* Enter the self employment budgeting method."
                                If trim(EARNED_INCOME_PANELS_ARRAY(method_date, ei_panel)) = "" Then ready_err_msg = ready_err_msg & vbNewLine & "* List the date the self employment method was selected."

                                If ready_err_msg <> "" Then
                                    basic_info_gathered = FALSE
                                    MsgBOx "Cannot open additional details section until the income information section is completed. Please resolve the following:" & vbNewLine & ready_err_msg
                                End If

                                If EARNED_INCOME_PANELS_ARRAY(self_emp_mthd, ei_panel) = "01 - 50% Grosss Inc" Then
                                    pay_item = pay_item + 1
                                    ReDim Preserve LIST_OF_INCOME_ARRAY(reason_amt_excluded, pay_item)
                                    LIST_OF_INCOME_ARRAY(panel_indct, pay_item) = ei_panel
                                    dlg_factor = dlg_factor + 1
                                End If

                                sm_err_msg = "LOOP" & sm_err_msg

                            End If

                            If ButtonPressed = plus_button Then
                                pay_item = pay_item + 1
                                ReDim Preserve LIST_OF_INCOME_ARRAY(reason_amt_excluded, pay_item)
                                LIST_OF_INCOME_ARRAY(panel_indct, pay_item) = ei_panel
                                dlg_factor = dlg_factor + 1

                                sm_err_msg = "LOOP" & sm_err_msg

                            End If

                            If ButtonPressed = minus_button Then
                                pay_item = pay_item - 1
                                ReDim Preserve LIST_OF_INCOME_ARRAY(reason_amt_excluded, pay_item)
                                dlg_factor = dlg_factor - 1
                                sm_err_msg = "LOOP" & sm_err_msg
                            End If

                            If sm_err_msg <> "" AND left(sm_err_msg, 4) <> "LOOP" then MsgBox "Please resolve before continuing:" & vbNewLine & sm_err_msg

                        Loop until sm_err_msg = ""
                        call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
                    LOOP UNTIL are_we_passworded_out = false



                Loop until big_err_msg = ""
                call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
            LOOP UNTIL are_we_passworded_out = false

        End If
    End If

    'NAVIGATE to RBIC for each HH MEMBER and ask if Income Information was received for this RBIC

Next


                                '----------------------------------------------------------'
                    '---------------------------------------------------------------------------------'
'------------------------------------------------- DETERMINING WHICH MONTHS TO UPDATE --------------------------------------------------'
                    '---------------------------------------------------------------------------------'
                                '----------------------------------------------------------'


list_of_all_months_to_update = "~"

For ei_panel = 0 to UBOUND(EARNED_INCOME_PANELS_ARRAY, 2)
    If EARNED_INCOME_PANELS_ARRAY(income_received, ei_panel) = TRUE Then
        EARNED_INCOME_PANELS_ARRAY(update_this_month, ei_panel) = FALSE

        'TODO find handling for EARNED_INCOME_PANELS_ARRAY(cash_mos_list, ei_panel in here somewhere
        mm_1_yy = EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel) & "/1/" & EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel)
        mm_1_yy = DateValue(mm_1_yy)
        If InStr(list_of_all_months_to_update, "~" & mm_1_yy & "~") = 0 Then
            list_of_all_months_to_update = list_of_all_months_to_update & mm_1_yy & "~"
        End If

        'If EARNED_INCOME_PANELS_ARRAY(update_futue_chkbx, ei_panel) = checked AND EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel) <> CM_plus_1_mo AND EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel) <> CM_plus_1_yr Then
        If EARNED_INCOME_PANELS_ARRAY(update_futue_chkbx, ei_panel) = checked Then
            next_month = mm_1_yy
            CM_plus_2 = DateValue(CM_plus_2_mo & "/1/" & CM_plus_2_yr)
            CM_plus_2 = DateValue(CM_plus_2)
            Do
                'MsgBox next_month
                If InStr(list_of_all_months_to_update, "~" & next_month & "~") = 0 Then
                    list_of_all_months_to_update = list_of_all_months_to_update & next_month & "~"
                End If

                next_month = DateAdd("m", 1, next_month)
            Loop until next_month = CM_plus_2
        End If
    End If
    'MsgBox "2 - " & list_of_all_months_to_update
Next

list_of_all_months_to_update = right(list_of_all_months_to_update, len(list_of_all_months_to_update)-1)
list_of_all_months_to_update = left(list_of_all_months_to_update, len(list_of_all_months_to_update)-1)
If InStr(list_of_all_months_to_update, "~") <> 0 Then
    update_months_array = split(list_of_all_months_to_update, "~")
    Call sort_dates(update_months_array)
Else
    update_months_array = array(list_of_all_months_to_update)
End If



                        '----------------------------------------------------------'
                '---------------------------------------------------------------------------------'
'-------------------------------------------------GIONG TO UPDATE THE PANEL --------------------------------------------------'
                '---------------------------------------------------------------------------------'
                        '----------------------------------------------------------'


' First we are going to get all the way to SELF
' Then we will go in to each month that is in our list
' Once there we will loop through the panels, pick the one that needs updating AND determine if the month we are in is one to update FOR THAT Panel
' Then we loop through the income to actually update the panel
' Array witin and array within an array (probably some more arrays)
Call back_to_SELF
' For each active_month in update_months_array
'     MsgBox active_month
' Next
next_cash_month = 0
For each active_month in update_months_array
    MAXIS_footer_month = DatePart("m", active_month)
    MAXIS_footer_month = right("00" & MAXIS_footer_month, 2)
    MAXIS_footer_year = DatePart("yyyy", active_month)
    MAXIS_footer_year = right(MAXIS_footer_year, 2)

    RETRO_month = DateAdd("m", -2, active_month)
    RETRO_footer_month = DatePart("m", RETRO_month)
    RETRO_footer_month = right("00" & RETRO_footer_month, 2)
    RETRO_footer_year = DatePart("yyyy", RETRO_month)
    RETRO_footer_year = right(RETRO_footer_year, 2)

    EMReadScreen summ_check, 4, 2, 46
    If summ_check <> "SUMM" Then
        Call back_to_SELF

        Do
            Call navigate_to_MAXIS_screen("STAT", "SUMM")
            EMReadScreen summ_check, 4, 2, 46
        Loop until summ_check = "SUMM"

    End If

    For ei_panel = 0 to UBOUND(EARNED_INCOME_PANELS_ARRAY, 2)

        If EARNED_INCOME_PANELS_ARRAY(income_received, ei_panel) = TRUE Then
            'ALL THE JUICY BITS GO HERE
            top_of_order = EARNED_INCOME_PANELS_ARRAY(order_ubound, ei_panel)

            'Find all the checks in this month
            'TODO - create a list of all the checks for THIS month for THIS income so later we can just loop through that lsit to update JOBS'
            this_month_checks_array = ""
            checks_list = ""

            ' MsgBox EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel)
            this_month = MAXIS_footer_month & "/1/" & MAXIS_footer_year


            If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "1 - One Time Per Month" Then
                day_of_month = DatePart("d", EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel))

                the_day_of_pay = MAXIS_footer_month & "/" & day_of_month & "/" & MAXIS_footer_year
                the_day_of_pay = DateValue(the_day_of_pay)
                checks_list = checks_list & "~" & the_day_of_pay

            ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "2 - Two Times Per Month" Then


            ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "3 - Every Other Week" Then
                the_date = DateValue(EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel))
                Do
                    If DatePart("m", the_date) = DatePart("m", this_month) AND DatePart("yyyy", the_date) = DatePart("yyyy", this_month) Then
                        checks_list = checks_list & "~" & the_date
                    End If
                    the_date = DateAdd("d", 14, the_date)
                Loop until right("0" & DatePart("m", the_date), 2) = CM_plus_2_mo AND right(DatePart("yyyy", the_date), 2) = CM_plus_2_yr
            ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "4 - Every Week" Then
                the_date = DateValue(EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel))
                Do
                    If DatePart("m", the_date) = DatePart("m", this_month) AND DatePart("yyyy", the_date) = DatePart("yyyy", this_month) Then
                        checks_list = checks_list & "~" & the_date
                    End If
                    the_date = DateAdd("d", 14, the_date)
                Loop until right("0" & DatePart("m", the_date), 2) = CM_plus_2_mo AND right(DatePart("yyyy", the_date), 2) = CM_plus_2_yr

            ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "5 - Other" Then
            End If

            ' MsgBox checks_list

            If left(checks_list, 1) = "~" Then checks_list = right(checks_list, len(checks_list)-1)
            If InStr(checks_list, "~") <> 0 Then
                If left(checks_list, 1) = "~" Then checks_list = right(checks_list, len(checks_list)-1)
                this_month_checks_array = Split(checks_list,"~")
            Else
                this_month_checks_array = Array(checks_list)
            End If


            'List of the retro months for this month'
            retro_month_checks_array = ""
            checks_list = ""

            If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "1 - One Time Per Month" Then
                day_of_month = DatePart("d", EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel))

                the_day_of_pay = RETRO_footer_month & "/" & day_of_month & "/" & RETRO_footer_year
                the_day_of_pay = DateValue(the_day_of_pay)
                checks_list = checks_list & "~" & the_day_of_pay


            ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "2 - Two Times Per Month" Then


            ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "3 - Every Other Week" Then
                the_date = DateValue(EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel))
                Do
                    If DatePart("m", the_date) = DatePart("m", RETRO_month) AND DatePart("yyyy", the_date) = DatePart("yyyy", RETRO_month) Then
                        checks_list = checks_list & "~" & the_date
                    End If
                    the_date = DateAdd("d", 14, the_date)
                Loop until right("0" & DatePart("m", the_date), 2) = CM_plus_2_mo AND right(DatePart("yyyy", the_date), 2) = CM_plus_2_yr
            ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "4 - Every Week" Then
                the_date = DateValue(EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel))
                Do
                    If DatePart("m", the_date) = DatePart("m", RETRO_month) AND DatePart("yyyy", the_date) = DatePart("yyyy", RETRO_month) Then
                        checks_list = checks_list & "~" & the_date
                    End If
                    the_date = DateAdd("d", 14, the_date)
                Loop until right("0" & DatePart("m", the_date), 2) = CM_plus_2_mo AND right(DatePart("yyyy", the_date), 2) = CM_plus_2_yr

            ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "5 - Other" Then
            End If

            ' MsgBox "RETRO - " & checks_list

            If left(checks_list, 1) = "~" Then checks_list = right(checks_list, len(checks_list)-1)
            If InStr(checks_list, "~") <> 0 Then
                If left(checks_list, 1) = "~" Then checks_list = right(checks_list, len(checks_list)-1)
                retro_month_checks_array = Split(checks_list,"~")
            Else
                retro_month_checks_array = Array(checks_list)
            End If

            If EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel) = MAXIS_footer_month AND EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel) = MAXIS_footer_year Then EARNED_INCOME_PANELS_ARRAY(update_this_month, ei_panel) = TRUE
            ' MsgBox "Update this month - " & EARNED_INCOME_PANELS_ARRAY(update_this_month, ei_panel)
            If EARNED_INCOME_PANELS_ARRAY(update_this_month, ei_panel) = TRUE Then
                ' MsgBox "Panel type - " & EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel)
                If EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel) = "JOBS" Then

                    Call Navigate_to_MAXIS_screen("STAT", "JOBS")
                    EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel), 20, 76
                    EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel), 20, 79
                    transmit
                    PF9
                    ' MsgBox "In EDIT"

                    EMWriteScreen left(EARNED_INCOME_PANELS_ARRAY(income_type, ei_panel), 1), 5, 34
                    EMWriteScreen left(EARNED_INCOME_PANELS_ARRAY(income_verif, ei_panel), 1), 6, 34
                    EMWriteScreen "      ", 6, 75
                    EMWriteScreen EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel), 6, 75
                    EMWriteScreen EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel), 18, 35

                    If EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) = checked THen
                        EMWriteScreen "X", 19, 38
                        transmit
                        PF20

                        list_row = 9
                        beg_of_list_check = ""
                        Do
                            EMWriteScreen "  ", list_row, 13
                            EMWriteScreen "  ", list_row, 16
                            EMWriteScreen "  ", list_row, 19
                            EMWriteScreen "        ", list_row, 25
                            EMWriteScreen "      ", list_row, 35

                            list_row = list_row + 1

                            If list_row = 14 Then
                                PF19
                                PF19

                                EMReadScreen beg_of_list_check, 10, 20, 18
                                list_row = 9
                            End If
                        Loop until beg_of_list_check = "FIRST PAGE"

                        EMWriteScreen "      ", 8, 64
                        EMWriteScreen "        ", 9, 66
                        EMWriteScreen "        ", 13, 66
                        EMWriteScreen "  ", 14, 64

                        Call create_MAXIS_friendly_date(date, 0, 5, 34)
                        EMWriteScreen left(EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel), 1), 5, 64

                        If EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_estimate Then
                            EMWriteScreen EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel), 9, 66
                            EMWriteScreen EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel), 8, 64
                        End If
                        If EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_actual Then
                            list_row = 9
                            For order_number = 1 to top_of_order                        'loop through the order number lowest to highest
                                For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                    'conditional if it is the right panel AND the order matches - then do the thing you need to do
                                    If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(check_order, all_income) = order_number Then
                                        If LIST_OF_INCOME_ARRAY(budget_in_SNAP_no, all_income) = unchecked Then
                                            Call create_MAXIS_friendly_date(LIST_OF_INCOME_ARRAY(pay_date, all_income), 0, list_row, 13)
                                            net_amount = LIST_OF_INCOME_ARRAY(gross_amount, all_income) - LIST_OF_INCOME_ARRAY(exclude_amount, all_income)
                                            EMWriteScreen net_amount, list_row, 25
                                            EMWriteScreen LIST_OF_INCOME_ARRAY(hours, all_income), list_row, 35

                                            list_row = list_row + 1
                                            If list_row = 14 Then
                                                PF20
                                                list_row = 9
                                            End If
                                        End If
                                    End If
                                next
                            next

                        End If
                        transmit
                        transmit
                        PF3


                        jobs_row = 12
                        jobs_col = 25
                        Do
                            If jobs_col = 38 Then
                                EMWriteScreen "        ", jobs_row, jobs_col
                            ElseIf jobs_col = 67 Then
                                EMWriteScreen "        ", jobs_row, jobs_col
                            Else
                                EMWriteScreen "  ", jobs_row, jobs_col
                            End If

                            jobs_col = jobs_col + 3
                            If jobs_col = 34 then jobs_col = 38
                            If jobs_col = 41 then jobs_col = 54
                            If jobs_col = 63 then jobs_col = 67
                            If jobs_col = 70 Then
                                jobs_col = 25
                                jobs_row = jobs_row + 1
                            End If
                        Loop until jobs_row = 17



                        jobs_row = 12
                        total_hours = 0

                        For each this_date in this_month_checks_array
                            date_found = FALSE
                            For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                'conditional if it is the right panel AND the order matches - then do the thing you need to do
                                ' If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND DateValue(LIST_OF_INCOME_ARRAY(pay_date, all_income)) = this_date Then
                                If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND DateDiff("d", LIST_OF_INCOME_ARRAY(pay_date, all_income), this_date) = 0 Then
                                    date_found = TRUE
                                    Call create_MAXIS_friendly_date(LIST_OF_INCOME_ARRAY(pay_date, all_income), 0, jobs_row, 54)
                                    EMWriteScreen LIST_OF_INCOME_ARRAY(gross_amount, all_income), jobs_row, 67
                                    total_hours = total_hours + LIST_OF_INCOME_ARRAY(hours, all_income)


                                End If
                            Next

                            If date_found = FALSE Then
                                Call create_MAXIS_friendly_date(this_date, 0, jobs_row, 54)
                                EMWriteScreen EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel), jobs_row, 67
                                total_hours = total_hours + EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel)
                            End If
                            jobs_row = jobs_row + 1

                        Next
                        total_hours = Round(total_hours)
                        EMWriteScreen "   ", 18, 72
                        EMWriteScreen "   ", 18, 43
                        EMWriteScreen total_hours, 18, 72

                        ' For order_number = 1 to top_of_order                        'loop through the order number lowest to highest
                        '     For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                        '         'conditional if it is the right panel AND the order matches - then do the thing you need to do
                        '         If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(check_order, all_income) = order_number Then
                        '             pay_month = right("00" & DatePart("m", LIST_OF_INCOME_ARRAY(pay_date, all_income)), 2)
                        '             pay_year = right(DatePart("yyyy", LIST_OF_INCOME_ARRAY(pay_date, all_income)), 2)
                        '
                        '             If pay_month = MAXIS_footer_month AND pay_year = MAXIS_footer_year Then
                        '
                        '                 Call create_MAXIS_friendly_date(LIST_OF_INCOME_ARRAY(pay_date, all_income), 0, jobs_row, 13)
                        '                 EMWriteScreen LIST_OF_INCOME_ARRAY(gross_amount, all_income), jobs_row, 25
                        '                 total_hours = total_hours + LIST_OF_INCOME_ARRAY(hours, all_income)
                        '
                        '                 jobs_row = jobs_row + 1
                        '             End If
                        '         End If
                        '     next
                        ' next


                    End If

                    If EARNED_INCOME_PANELS_ARRAY(apply_to_HC, ei_panel) = checked Then

                        For jobs_row = 12 to 16
                            EMWriteScreen "  ", jobs_row, 54
                            EMWriteScreen "  ", jobs_row, 57
                            EMWriteScreen "  ", jobs_row, 60
                            EMWriteScreen "        ", jobs_row, 67
                        Next

                        For jobs_row = 12 to 16
                            EMWriteScreen "  ", jobs_row, 25
                            EMWriteScreen "  ", jobs_row, 28
                            EMWriteScreen "  ", jobs_row, 31
                            EMWriteScreen "        ", jobs_row, 38
                        Next

                        jobs_row = 12
                        total_hours = 0

                        For each this_date in this_month_checks_array
                            date_found = FALSE
                            For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                'conditional if it is the right panel AND the order matches - then do the thing you need to do
                                If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND DateDiff("d", LIST_OF_INCOME_ARRAY(pay_date, all_income), this_date) = 0 Then
                                    date_found = TRUE
                                    Call create_MAXIS_friendly_date(LIST_OF_INCOME_ARRAY(pay_date, all_income), 0, jobs_row, 54)
                                    EMWriteScreen LIST_OF_INCOME_ARRAY(gross_amount, all_income), jobs_row, 67
                                    total_hours = total_hours + LIST_OF_INCOME_ARRAY(hours, all_income)


                                End If
                            Next

                            If date_found = FALSE Then
                                Call create_MAXIS_friendly_date(this_date, 0, jobs_row, 54)
                                EMWriteScreen EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel), jobs_row, 67
                                total_hours = total_hours + EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel)
                            End If
                            jobs_row = jobs_row + 1

                        Next
                        total_hours = Round(total_hours)
                        EMWriteScreen "   ", 18, 72
                        EMWriteScreen "   ", 18, 43
                        EMWriteScreen total_hours, 18, 72

                        If MAXIS_footer_month = CM_plus_1_mo AND MAXIS_footer_year = CM_plus_1_yr Then
                            EMWriteScreen "X", 19, 48
                            transmit

                            EMWriteScreen "        ", 11, 63
                            EMWriteScreen EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel), 11, 63
                            transmit
                            transmit

                        End If
                    End If

                    If EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, ei_panel) = checked Then
                        ReDim Preserve CASH_MONTHS_ARRAY(8, next_cash_month)
                        CASH_MONTHS_ARRAY(retro_updtd, next_cash_month) = FALSE
                        CASH_MONTHS_ARRAY(prosp_updtd, next_cash_month) = FALSE
                        CASH_MONTHS_ARRAY(panel_indct, next_cash_month) = ei_panel
                        CASH_MONTHS_ARRAY(cash_mo_yr, next_cash_month) = MAXIS_footer_month & "/" & MAXIS_footer_year
                        CASH_MONTHS_ARRAY(retro_mo_yr, next_cash_month) = RETRO_footer_month & "/" & RETRO_footer_year

                        'RETROSPECTIVE SIDE'
                        For jobs_row = 12 to 16
                            EMWriteScreen "  ", jobs_row, 25
                            EMWriteScreen "  ", jobs_row, 28
                            EMWriteScreen "  ", jobs_row, 31
                            EMWriteScreen "        ", jobs_row, 38
                        Next

                        jobs_row = 12
                        total_hours = 0
                        total_pay = 0
                        count_checks = 0

                        For each this_date in retro_month_checks_array
                            date_found = FALSE

                            For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                'conditional if it is the right panel AND the order matches - then do the thing you need to do
                                If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND DateDiff("d", LIST_OF_INCOME_ARRAY(pay_date, all_income), this_date) = 0 Then
                                    date_found = TRUE
                                    CASH_MONTHS_ARRAY(retro_updtd, next_cash_month) = TRUE
                                    Call create_MAXIS_friendly_date(LIST_OF_INCOME_ARRAY(pay_date, all_income), 0, jobs_row, 25)
                                    EMWriteScreen LIST_OF_INCOME_ARRAY(gross_amount, all_income), jobs_row, 38
                                    total_hours = total_hours + LIST_OF_INCOME_ARRAY(hours, all_income)
                                    total_pay = total_pay + LIST_OF_INCOME_ARRAY(gross_amount, all_income)
                                    count_checks = count_checks + 1

                                End If
                            Next

                            jobs_row = jobs_row + 1

                        Next
                        total_hours = Round(total_hours)
                        'EMWriteScreen "   ", 18, 72
                        EMWriteScreen "   ", 18, 43
                        EMWriteScreen total_hours, 18, 43

                        If count_checks <> 0 Then
                            this_month_ave_pay = total_pay/count_checks
                            this_month_ave_pay = FormatNumber(this_month_ave_pay, 2,,0)
                            this_month_ave_hours = total_hours/count_checks

                            CASH_MONTHS_ARRAY(mo_retro_pay, next_cash_month) = FormatNumber(total_pay, 2,,0)
                            CASH_MONTHS_ARRAY(mo_retro_hrs, next_cash_month) = total_hours
                        End If


                        'PROSPECTIVE SIDE'
                        For jobs_row = 12 to 16
                            EMWriteScreen "  ", jobs_row, 54
                            EMWriteScreen "  ", jobs_row, 57
                            EMWriteScreen "  ", jobs_row, 60
                            EMWriteScreen "        ", jobs_row, 67
                        Next

                        jobs_row = 12
                        total_hours = 0
                        CASH_MONTHS_ARRAY(prosp_updtd, next_cash_month) = TRUE
                        CASH_MONTHS_ARRAY(mo_prosp_pay, next_cash_month) = 0

                        For each this_date in this_month_checks_array
                            date_found = FALSE
                            For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                'conditional if it is the right panel AND the order matches - then do the thing you need to do
                                ' If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then MsgBox "This Date - " & this_date & vbNewLine & "Array Date - " & LIST_OF_INCOME_ARRAY(pay_date, all_income)
                                If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND DateDiff("d", LIST_OF_INCOME_ARRAY(pay_date, all_income), this_date) = 0 Then
                                    date_found = TRUE
                                    Call create_MAXIS_friendly_date(LIST_OF_INCOME_ARRAY(pay_date, all_income), 0, jobs_row, 54)
                                    EMWriteScreen LIST_OF_INCOME_ARRAY(gross_amount, all_income), jobs_row, 67
                                    total_hours = total_hours + LIST_OF_INCOME_ARRAY(hours, all_income)
                                    CASH_MONTHS_ARRAY(mo_prosp_pay, next_cash_month) = CASH_MONTHS_ARRAY(mo_prosp_pay, next_cash_month) + LIST_OF_INCOME_ARRAY(gross_amount, all_income)
                                    ' MsgBox "Found Date"
                                End If
                            Next

                            If date_found = FALSE Then
                                Call create_MAXIS_friendly_date(this_date, 0, jobs_row, 54)
                                If count_checks <> 0 Then
                                    EMWriteScreen this_month_ave_pay, jobs_row, 67
                                    total_hours = total_hours + this_month_ave_hours
                                    CASH_MONTHS_ARRAY(mo_prosp_pay, next_cash_month) = CASH_MONTHS_ARRAY(mo_prosp_pay, next_cash_month) + this_month_ave_pay
                                Else
                                    EMWriteScreen EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel), jobs_row, 67
                                    total_hours = total_hours + EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel)
                                    CASH_MONTHS_ARRAY(mo_prosp_pay, next_cash_month) = CASH_MONTHS_ARRAY(mo_prosp_pay, next_cash_month) + EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel)
                                End If
                            End If
                            jobs_row = jobs_row + 1

                        Next
                        total_hours = Round(total_hours)
                        EMWriteScreen "   ", 18, 72
                        'EMWriteScreen "   ", 18, 43
                        EMWriteScreen total_hours, 18, 72
                        CASH_MONTHS_ARRAY(mo_prosp_hrs, next_cash_month) = total_hours
                        CASH_MONTHS_ARRAY(mo_prosp_pay, next_cash_month) = FormatNumber(CASH_MONTHS_ARRAY(mo_prosp_pay, next_cash_month), 2,,0)

                        next_cash_month = next_cash_month + 1

                    End If


                End If

            End If
            ' MsgBox "Does this look right?"

            iF EARNED_INCOME_PANELS_ARRAY(update_futue_chkbx, ei_panel) = unchecked Then EARNED_INCOME_PANELS_ARRAY(update_this_month, ei_panel) = FALSE
        End If
    Next

    transmit

    EmWriteScreen "BGTX", 20, 71
    transmit
    If active_month <> update_months_array(ubound(update_months_array)) Then EmWriteScreen "Y", 16, 54
    transmit

    If active_month <> update_months_array(ubound(update_months_array)) Then
        EmWriteScreen "SUMM", 20, 71
        transmit
    End If
Next

'
' 'THIS IS GOOD START FOR SOME OF THE PANEL UPDATING BUT IT NEEDS TO GO IN THE ABOVE LOOP'
' For ei_panel = 0 to UBOUND(EARNED_INCOME_PANELS_ARRAY, 2)
'     Call back_to_SELF
'     If EARNED_INCOME_PANELS_ARRAY(income_received, ei_panel) = TRUE Then
'         MAXIS_footer_month = EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel)
'         MAXIS_footer_year = EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel)
'
'         Do
'             If EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel) = "JOBS" Then
'
'                 Call Navigate_to_MAXIS_screen("STAT", "JOBS")
'                 EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel), 20, 76
'                 EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel), 20, 79
'                 transmit
'                 PF9
'
'                 'FOR SNAP BUDGETING
'                 'script will determine which dates of the current month need to be entered.
'                 'go through all of the income information and find the dates that need to be entered
'                 'if no check is found for the coresponding date the script enteres the date and the average pay
'                 'the script goes through the checks again and any that were not previously entered will be entered.
'
'                 'open the PIC
'                 'read to see if it was updated already - create a BOOLEAN for the first month to make sure that it is updated in the first month THEN use the date updated to determine if it was updated
'                 'enter all the information'
'
'                 'FOR CASH BUDGETING
'                 'determine all of the dates for prosp side
'                 'if the month we are in is retro budgeted,
'                     'go through all the checks IN ORDER and find all of the ones for CM-2 and put in retro side
'                     'go through all the checks IN ORDER and find all of the ones for CM and enter in to prosp side
'                 'read through the prosp side and determine if any of the dates were missed.
'                 'if they were missed, create an average from the retro side and enter an estimated amount.
'
'                 'if the month we are in is prosp Budgeted
'                 'go through all of the dates and find paychecks for those dates, entering estimates for missing dates.
'
'                 For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
'                     If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then
'                         'Go through all of the paychecks IN ORDER.
'                             'If the check is for the month that we are in, then it will enter that check.
'
'
'
'                     End If
'                 Next
'
'                 'after all the checks are entered, the script will go back and read the '
'
'                 If LIST_OF_INCOME_ARRAY(budget_in_SNAP_yes, all_income) = checked Then
'                     If LIST_OF_INCOME_ARRAY(exclude_amount, all_income) = "" Then LIST_OF_INCOME_ARRAY(exclude_amount, all_income) = 0
'                     LIST_OF_INCOME_ARRAY(exclude_amount, all_income) = LIST_OF_INCOME_ARRAY(exclude_amount, all_income) * 1
'                     LIST_OF_INCOME_ARRAY(gross_amount, all_income) = LIST_OF_INCOME_ARRAY(gross_amount, all_income) * 1
'                     net_amount = LIST_OF_INCOME_ARRAY(gross_amount, all_income) - LIST_OF_INCOME_ARRAY(exclude_amount, all_income)
'                     total_of_counted_income = total_of_counted_income + net_amount
'                     number_of_checks_budgeted = number_of_checks_budgeted + 1
'
'                     LIST_OF_INCOME_ARRAY(hours, all_income) = LIST_OF_INCOME_ARRAY(hours, all_income) * 1
'                     total_of_hours = total_of_hours + LIST_OF_INCOME_ARRAY(hours, all_income)
'                 End If
'
'                 If EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) Then
'
'                 End If
'                 'If EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, ei_panel)
'                 'If EARNED_INCOME_PANELS_ARRAY(apply_to_HC, ei_panel)
'
'
'             End If
'
'             If EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel) = "BUSI" Then
'                 Call Navigate_to_MAXIS_screen("STAT", "BUSI")
'                 EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel), 20, 76
'                 EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel), 20, 79
'                 transmit
'                 PF9
'
'             End If
'
'             If EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel) = "RBIC" Then
'                 Call Navigate_to_MAXIS_screen("STAT", "RBIC")
'                 EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel), 20, 76
'                 EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel), 20, 79
'                 transmit
'                 PF9
'
'             End If
'
'             If EARNED_INCOME_PANELS_ARRAY(update_futue_chkbx, ei_panel) = 0 then exit do
'
'             'Navigates to the current month + 1 footer month, then back into the JOBS panel
'             CALL write_value_and_transmit("BGTX", 20, 71)
'             CALL write_value_and_transmit("y", 16, 54)
'             EMReadScreen all_months_check, 24, 24, 2
'
'             EMReadScreen MAXIS_footer_month, 2, 20, 55
'             EMReadScreen MAXIS_footer_year, 2, 20, 58
'
'             transmit
'
'         Loop until all_months_check = "CONTINUATION NOT ALLOWED"
'         PF3
'     End If
'     MAXIS_footer_month = original_month
'     MAXIS_footer_year = original_year
' Next



                '----------------------------------------------------------'
        '---------------------------------------------------------------------------------'
'-------------------------------------------------CASE NOTING --------------------------------------------------'
        '---------------------------------------------------------------------------------'
                '----------------------------------------------------------'

For ei_panel = 0 to UBOUND(EARNED_INCOME_PANELS_ARRAY, 2)
    prog_list = ""

    If EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) = checked Then prog_list = prog_list & "/SNAP"
    If EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, ei_panel) = checked Then prog_list = prog_list & "/CASH"
    If EARNED_INCOME_PANELS_ARRAY(apply_to_HC, ei_panel) = checked Then prog_list = prog_list & "/HC"

    If left(prog_list, 1) = "/" Then prog_list = right(prog_list, len(prog_list)-1)

    top_of_order = EARNED_INCOME_PANELS_ARRAY(order_ubound, ei_panel)

    Select Case EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel)

    Case "JOBS"

        If EARNED_INCOME_PANELS_ARRAY(this_is_a_new_panel, ei_panel) = TRUE OR EARNED_INCOME_PANELS_ARRAY(income_received, ei_panel) = TRUE Then


            Call start_a_blank_CASE_NOTE

            If EARNED_INCOME_PANELS_ARRAY(income_received, ei_panel) = TRUE Then
                Call write_variable_in_CASE_NOTE("INCOME DETAIL: M" & EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel) & " - JOBS - " & EARNED_INCOME_PANELS_ARRAY(employer, ei_panel) & " - PROG: " & prog_list)
            Else
                Call write_variable_in_CASE_NOTE("NEW JOB ADDED: M" & EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel) & " - JOBS - " & EARNED_INCOME_PANELS_ARRAY(employer, ei_panel) & " - PROG: " & prog_list)
            End If

            If EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) = checked Then
                Call write_variable_in_CASE_NOTE("Income Budget for SNAP -------------------------------------")

                Call write_bullet_and_variable_in_CASE_NOTE("Monthly budgeted income", "$" & EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel))
                Call write_bullet_and_variable_in_CASE_NOTE("Average per Pay Period", "$" & EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel))
                Call write_bullet_and_variable_in_CASE_NOTE("Average hours per week", EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel))
                Call write_bullet_and_variable_in_CASE_NOTE("Average pay per hour", "$" & EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel))
                Call write_bullet_and_variable_in_CASE_NOTE("Pay Frequency", EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel))

            End If
            If EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, ei_panel) = checked Then
                Call write_variable_in_CASE_NOTE("Income Budget for CASH -------------------------------------")

                For each_cash_month = 0 to UBOUND(CASH_MONTHS_ARRAY, 2)
                    If CASH_MONTHS_ARRAY(panel_indct, each_cash_month) = ei_panel Then
                        Call write_variable_in_CASE_NOTE("* Income updated in " & CASH_MONTHS_ARRAY(cash_mo_yr, each_cash_month))
                        If CASH_MONTHS_ARRAY(retro_updtd, each_cash_month) = TRUE Then Call write_variable_in_CASE_NOTE("* --RETRO Income updated: $" & CASH_MONTHS_ARRAY(mo_retro_pay, each_cash_month) & " total income for " & CASH_MONTHS_ARRAY(retro_mo_yr, each_cash_month) & " with " & CASH_MONTHS_ARRAY(mo_retro_hrs, each_cash_month) & " total hours.")
                        If CASH_MONTHS_ARRAY(prosp_updtd, each_cash_month) = TRUE Then Call write_variable_in_CASE_NOTE("* --Prosp Income updated: $" & CASH_MONTHS_ARRAY(mo_prosp_pay, each_cash_month) & " total income for " & CASH_MONTHS_ARRAY(cash_mo_yr, each_cash_month) & " with " & CASH_MONTHS_ARRAY(mo_prosp_hrs, each_cash_month) & " total hours.")

                    End If
                Next

            End If
            If EARNED_INCOME_PANELS_ARRAY(apply_to_HC, ei_panel) = checked Then

            End If

            If EARNED_INCOME_PANELS_ARRAY(income_received, ei_panel) = TRUE Then
                Call write_variable_in_CASE_NOTE("Verification Received: " & EARNED_INCOME_PANELS_ARRAY(verif_date, ei_panel) & "-----------------------------")

                Call write_bullet_and_variable_in_CASE_NOTE("Type Received", EARNED_INCOME_PANELS_ARRAY(income_verif, ei_panel))
                Call write_bullet_and_variable_in_CASE_NOTE("Explanation of Verification", EARNED_INCOME_PANELS_ARRAY(verif_explain, ei_panel))
                Call write_bullet_and_variable_in_CASE_NOTE("Days covered by check stubs", EARNED_INCOME_PANELS_ARRAY(days_of_verif, ei_panel))

                Call write_bullet_and_variable_in_CASE_NOTE("Conversation with", EARNED_INCOME_PANELS_ARRAY(spoke_with, ei_panel))
                Call write_bullet_and_variable_in_CASE_NOTE("Conversation Details", EARNED_INCOME_PANELS_ARRAY(convo_detail, ei_panel))


                Call write_variable_in_CASE_NOTE("Income Information Received -----------------------------------")

                If EARNED_INCOME_PANELS_ARRAY(order_ubound, ei_panel) <> "" AND EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel) <> "" Then        'If there is an order ubound then there are actual checks'
                    Call write_variable_in_CASE_NOTE("* Both actual check stubs and anticipated income estimates were received for this income.")

                    If EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_actual Then Call write_variable_in_CASE_NOTE("* Actual pay amounts used to determine income to budget.")
                    If EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_estimate Then Call write_variable_in_CASE_NOTE("* Income to budget determined by anticipated hours and rate of pay.")
                    Call write_bullet_and_variable_in_CASE_NOTE("Reason for choice", EARNED_INCOME_PANELS_ARRAY(selection_rsn, ei_panel))

                End If
                If EARNED_INCOME_PANELS_ARRAY(order_ubound, ei_panel) <> "" Then

                    Call write_variable_in_CASE_NOTE("* Checks provided to agency.")
                    For order_number = 1 to top_of_order                        'loop through the order number lowest to highest
                        For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                            'conditional if it is the right panel AND the order matches - then do the thing you need to do
                            If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(check_order, all_income) = order_number Then
                                If EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) = checked Then
                                    If LIST_OF_INCOME_ARRAY(budget_in_SNAP_yes, all_income) = checked Then
                                        If LIST_OF_INCOME_ARRAY(exclude_amount, all_income) <> 0 Then
                                            Call write_bullet_and_variable_in_CASE_NOTE(LIST_OF_INCOME_ARRAY(pay_date, all_income), "Gross: $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & " hrs. Only $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) - LIST_OF_INCOME_ARRAY(exclude_amount, all_income) & " included in SNAP budget because: " & LIST_OF_INCOME_ARRAY(reason_amt_excluded, all_income) & " - $" & LIST_OF_INCOME_ARRAY(exclude_amount, all_income) " of check not included.")
                                        Else
                                            Call write_bullet_and_variable_in_CASE_NOTE(LIST_OF_INCOME_ARRAY(pay_date, all_income), "Gross: $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & " hrs.")
                                        End If
                                    Else
                                        Call write_bullet_and_variable_in_CASE_NOTE(LIST_OF_INCOME_ARRAY(pay_date, all_income), "Gross: $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & " hrs. ** THIS CHECK EXCLUDED FROM SNAP BUDGET because " & LIST_OF_INCOME_ARRAY(reason_to_exclude, all_income))
                                    End If
                                Else
                                    Call write_bullet_and_variable_in_CASE_NOTE(LIST_OF_INCOME_ARRAY(pay_date, all_income), "Gross: $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & " hrs.")
                                End If
                            End If
                        next
                    next

                    If EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) = checked AND EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_actual Then Call write_variable_in_CASE_NOTE("* All included checks have been added to the PIC. Gross amount on PIC is reflective of the included pay amount.")

                End If
                If EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel) <> "" Then
                    Call write_variable_in_CASE_NOTE("* Anticipated Income Estimate provided to Agency.")

                    Call write_bullet_and_variable_in_CASE_NOTE("Hourly Pay Rate", "$" & EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel) "/hr")
                    Call write_bullet_and_variable_in_CASE_NOTE("Hours Per Week", EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) & " hours")
                    Call write_bullet_and_variable_in_CASE_NOTE("Pay Frequency", EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel))

                    ' EditBox 5, (dlg_factor * 20) + 140, 50, 15, EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel)
                    ' EditBox 75, (dlg_factor * 20) + 140, 40, 15, EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel)
                    ' DropListBox 130, (dlg_factor * 20) + 140, 85, 45, ""+chr(9)+"1 - One Time Per Month"+chr(9)+"2 - Two Times Per Month"+chr(9)+"3 - Every Other Week"+chr(9)+"4 - Every Week", EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel)

                End If

                ' If EARNED_INCOME_PANELS_ARRAY(selection_rsn, ei_panel) <> "" Then
                ' Else
                '     If EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_actual Then Call write_variable_in_CASE_NOTE("* Actual pay amounts used to determine income to budget as client only provided income information.")
                '     If EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_estimate Then Call write_variable_in_CASE_NOTE("* Income to budget determined by anticipated hours and rate of pay as this is the only information provided.")
                ' End If

                Call write_variable_in_CASE_NOTE("ACTION TAKEN: JOBS Updated ------------------------------------")

                If EARNED_INCOME_PANELS_ARRAY(update_futue_chkbx, ei_panel) = unchecked Then
                    Call write_variable_in_CASE_NOTE("* Updated jobs for the month " & EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel) & "/" & EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel) & ".")
                ElseIf EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel) = CM_plus_1_mo AND EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel) = CM_plus_1_yr Then
                    Call write_variable_in_CASE_NOTE("* Updated jobs for the month " & EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel) & "/" & EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel) & ".")
                ElseIf EARNED_INCOME_PANELS_ARRAY(update_futue_chkbx, ei_panel) = checked Then
                    Call write_variable_in_CASE_NOTE("* Updated jobs from " & EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel) & "/" & EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel) & " to " & CM_plus_1_mo & "/" & CM_plus_1_yr & ".")
                End If

                Call write_variable_in_CASE_NOTE("---")
                Call write_variable_in_CASE_NOTE(worker_signature)
            End If


        End If

    Case "BUSI"

    End Select
Next



script_end_procedure("")


'SAVING SOME STATIC DIALOG CODE FOR MAYBE RE WORKING'

' HOLDING FOR A VERSION TO BE PUT IN TO DLG EDIT'
' BeginDialog Dialog1, 0, 0, 606, 160, "Enter ALL Paychecks Received"
'   Text 10, 10, 265, 10, "JOBS 01 01 - EMPLOYER"
'   Text 310, 10, 50, 10, "Income Type:"
'   DropListBox 365, 10, 100, 45, "J - WIOA"+chr(9)+"W - Wages (Incl Tips)"+chr(9)+"E - EITC"+chr(9)+"G - Experience Works"+chr(9)+"F - Federal Work Study"+chr(9)+"S - State Work Study"+chr(9)+"O - Other"+chr(9)+"C - Contract Income"+chr(9)+"T - Training Program"+chr(9)+"P - Service Program"+chr(9)+"R - Rehab Program", income_type
'   GroupBox 475, 5, 125, 25, "Apply Income to Programs:"
'   CheckBox 485, 15, 30, 10, "SNAP", apply_to_SNAP
'   CheckBox 530, 15, 30, 10, "CASH", apply_to_CASH
'   CheckBox 570, 15, 20, 10, "HC", apply_to_HC
'   Text 5, 40, 60, 10, "JOBS Verif Code:"
'   DropListBox 65, 35, 105, 45, "1 - Pay Stubs/Tip Report"+chr(9)+"2 - Empl Statement"+chr(9)+"3 - Coltrl Stmt"+chr(9)+"4 - Other Document"+chr(9)+"5 - Pend Out State Verification"+chr(9)+"N - No Ver Prvd", JOBS_verif_code
'   Text 175, 40, 155, 10, "additional detail of verification received:"
'   EditBox 310, 35, 290, 15, Edit2
'   Text 5, 60, 90, 10, "Date verification received:"
'   EditBox 100, 55, 50, 15, verif_date
'   Text 5, 80, 80, 10, "Pay Date (MM/DD/YY):"
'   Text 90, 80, 50, 10, "Gross Amount:"
'   Text 145, 80, 25, 10, "Hours:"
'   Text 180, 65, 25, 25, "Use in SNAP budget"
'   Text 235, 80, 85, 10, "If not used, explain why:"
'   Text 355, 70, 245, 10, "If there is a specific amount that should be NOT budgeted from this check:"
'   Text 355, 80, 30, 10, "Amount:"
'   Text 410, 80, 30, 10, "Reason:"
'   EditBox 5, 90, 65, 15, pay_date
'   EditBox 90, 90, 45, 15, gross_amount
'   EditBox 145, 90, 25, 15, hours_on_check
'   OptionGroup RadioGroup1
'     RadioButton 180, 90, 25, 10, "Yes", budget_yes
'     RadioButton 210, 90, 25, 10, "No", budget_no
'   EditBox 235, 90, 115, 15, reason_not_budgeted
'   EditBox 355, 90, 45, 15, not_budgeted_amount
'   EditBox 410, 90, 185, 15, amount_not_budgeted_reason
'   Text 5, 115, 70, 10, "Anticipated Income"
'   Text 5, 130, 50, 10, "Rate of Pay/Hr"
'   Text 75, 130, 35, 10, "Hours/Wk"
'   Text 130, 130, 50, 10, "Pay Frequency"
'   Text 225, 115, 70, 10, "Regular Non-Monthly"
'   Text 225, 130, 25, 10, "Amount"
'   Text 280, 130, 50, 10, "Nbr of Months"
'   EditBox 5, 140, 50, 15, rate_of_pay
'   EditBox 75, 140, 40, 15, hours_per_week
'   DropListBox 130, 140, 85, 45, "1 - One Time Per Month"+chr(9)+"2 - Two Times Per Month"+chr(9)+"3 - Every Other Week"+chr(9)+"4 - Every Week", pay_frequency
'   EditBox 225, 140, 40, 15, non_monthly_amt
'   EditBox 280, 140, 30, 15, number_non_reg_months
'   ButtonGroup ButtonPressed
'     PushButton 440, 140, 15, 15, "+", add_another_check
'     PushButton 460, 140, 15, 15, "-", take_a_check_away
'     OkButton 495, 140, 50, 15
'     CancelButton 550, 140, 50, 15
' EndDialog

' BeginDialog Dialog1, 0, 0, 421, 240, "Confirm JOBS Budget"
'   Text 10, 10, 175, 10, "JOBS 01 01 - EMPLOYER"
'   Text 245, 10, 50, 10, "Pay Frequency"
'   DropListBox 305, 5, 95, 45, "1 - One Time Per Month"+chr(9)+"2 - Two Times Per Month"+chr(9)+"3 - Every Other Week"+chr(9)+"4 - Every Week"+chr(9)+"5 - Other", pay_frequency
'   Text 240, 30, 60, 10, "Income Start Date:"
'   EditBox 305, 25, 70, 15, income_start_date
'   GroupBox 5, 40, 410, 105, "SNAP Budget"
'   Text 10, 50, 100, 10, "Paychecks Inclued in Budget:"
'   Text 20, 65, 90, 10, "01/01/2018 - $400 - 40 hrs"
'   Text 20, 75, 90, 10, "01/15/2018- $400 - 40 hrs"
'   Text 10, 95, 130, 10, "Paychecks not included: 12/24/2018"
'   Text 185, 50, 90, 10, "Average hourly rate of pay:"
'   Text 185, 65, 90, 10, "Average weekly hours:"
'   Text 185, 80, 90, 10, "Average paycheck amount:"
'   Text 185, 95, 90, 10, "Monthly Budgeted Income:"
'   CheckBox 10, 110, 330, 10, "Check here if you confirm that this budget is correct and is the best estimate of anticipated income.", confirm_budget_checkbox
'   Text 10, 130, 60, 10, "Conversation with:"
'   ComboBox 75, 125, 60, 45, " "+chr(9)+"Client - not employee"+chr(9)+"Employee"+chr(9)+"Employer", converstion_with
'   Text 140, 130, 25, 10, "clarifies"
'   EditBox 170, 125, 235, 15, conversation_detail
'   GroupBox 5, 150, 410, 60, "CASH Budget"
'   Text 15, 165, 110, 10, "Actual Paychecks to add to JOBS:"
'   Text 25, 180, 90, 10, "01/01/2018 - $400 - 40 hrs"
'   Text 25, 190, 90, 10, "01/15/2018- $400 - 40 hrs"
'   ButtonGroup ButtonPressed
'     OkButton 315, 220, 50, 15
'     CancelButton 370, 220, 50, 15
' EndDialog
