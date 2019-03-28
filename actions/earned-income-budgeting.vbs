'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - EARNED INCOME BUDGETING.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
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
Call changelog_update("03/26/2019", "Fixed errors when pay is twice monthly. Added better handling for reading the employer name.", "Casey Love, Hennepin County")
call changelog_update("03/05/2019", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'SCRIPT TABLE OF CONTENTS-------------------------------------------------------
'FUNCTIONS  .   .   .   .   .   .   .   .   .   .   . Line 96
    'sort_dates .   .   .   .   .   .   .   .   .   . Line 98
    'navigate_to_approved_SNAP_eligibility  .   .   . Line 135
'CONSTANTS  .   .   .   .   .   .   .   .   .   .   . Line 156
'SCRIPT START   .   .   .   .   .   .   .   .   .   . Line 264
    'INITIAL Dialog .   .   .   .   .   .   .   .   . Line 275
'FINDING ALL CURRENT EI PANELS  .   .   .   .   .   . Line 366
    'JOBS   .   .   .   .   .   .   .   .   .   .   . Line 377
    'BUSI   .   .   .   .   .   .   .   .   .   .   . Line 539
'ADDING NEW EI PANELS   .   .   .   .   .   .   .   . Line 619
    'ASK TO ADD NEW PANEL Dailog.   .   .   .   .   . Line 634
    'TYPE OF PANEL TO ADD Dialog.   .   .   .   .   . Line 683
    'NEW JOB PANEL Dialog   .   .   .   .   .   .   . Line 730
    'CONFIRM ADD PANEL MONTH Dialog .   .   .   .   . Line 815
'GATHERING PAY INFORMATION FOR EACH PANEL   .   .   . Line 1014
    'vbYesNo MsgBox - employer_check.   .   .   .   . Line 1038
    'ENTER PAY Dialog   .   .   .   .   .   .   .   . Line 1104
    'CHOOSE CORRECT METHOD Dialog   .   .   .   .   . Line 1303
    'Order checks chronological .   .   .   .   .   . Line 1372
    'Find dates for bimonthly Dialog.   .   .   .   . Line 1392
    'Looking for missing checks .   .   .   .   .   . Line 1548
    'FREQUENCY ISSUE Dialog .   .   .   .   .   .   . Line 1708
    'Use Estimate functionality .   .   .   .   .   . Line 1854
    'CONFIRM BUDGET Dialog  .   .   .   .   .   .   . Line 2067
'DETERMINING WHICH MONTHS TO UPDATE .   .   .   .   . Line 2606
'GOING TO UPDATE THE PANEL  .   .   .   .   .   .   . Line 2643
    'Updating for SNAP  .   .   .   .   .   .   .   . Line 2970
    'Updating for GRH   .   .   .   .   .   .   .   . Line 3179
    'Updating for HC.   .   .   .   .   .   .   .   . Line 3284
    'Updating for Cash  .   .   .   .   .   .   .   . Line 3347
'CASE NOTING.   .   .   .   .   .   .   .   .   .   . Line 3494

'SEARCH TAGS--------------------------------------------------------------------
'FUTURE FUNCTIONALITY        - ideas/code to be added at a future time.
'TESTING NEEDED              - code created but not tested or vetted
'NEED COMMENTS               - code that has not been commented sufficiently
'BUGGY CODE                 - code that has either been reported as having bugs or appears it may be buggy
'PROCEDURE CLARIFICATION    - possible place to confirm the script's actions with subject matter experts
'REMOVE CODE                - code that might be superfluous
'NEED POLICY REFERENCE      - process should have a reference to policy saved in it

'FUNCTIONS==================================================================================================================

function sort_dates(dates_array)
'--- Takes an array of dates and reorders them to be chronological.
'~~~~~ dates_array: an array of dates only
'===== Keywords: MAXIS, date, order, list, array
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

function navigate_to_approved_SNAP_eligibility()
'--- This function navigates to ELIG/FS and finds the most recent approved version
'===== Keywords: MAXIS, navigate, SNAP
    navigate_to_MAXIS_screen "ELIG", "FS"

    EMWriteScreen "NN", 19, 78
    transmit
    elig_row = 7
    Do
        EMReadScreen app_status, 10, elig_row, 50
        app_status = trim(app_status)
        If app_status = "APPROVED" Then
            EMReadScreen approved_version, 2, elig_row, 22
            EMWriteScreen approved_version, 18, 54
            transmit
        End If
        elig_row = elig_row + 1
    Loop until app_status = "APPROVED"
end function

'Declarations ==============================================================================================================
'CONSTANTS'
'Constants for the array that deals with each panel - EARNED_INCOME_PANELS_ARRAY
const panel_type            = 1
const panel_member          = 2
const panel_instance        = 3
const employer              = 4
const income_type           = 5
const income_verif          = 6
const hourly_wage           = 7
const income_start_dt       = 8
const income_end_dt         = 9
const income_list_indct     = 10
const pay_freq              = 11
const date_of_calc          = 12
const hrs_per_wk            = 13
const pay_per_hr            = 14
const ave_hrs_per_pay       = 15
const ave_inc_per_pay       = 16
const snap_ave_inc_per_pay  = 17
const snap_ave_hrs_per_pay  = 18
const snap_hrs_per_wk       = 19
const SNAP_mo_inc           = 20
const reg_non_monthly       = 21
const numb_months           = 22
const self_emp_mthd         = 23
const method_date           = 24
const reptd_hours           = 25
const apply_to_SNAP         = 26
const apply_to_CASH         = 27
const apply_to_HC           = 28
const apply_to_GRH          = 29
const pay_weekday           = 30
const income_received       = 31
const verif_date            = 32
const verif_explain         = 33
const old_verif             = 34
const initial_month_mo      = 35
const initial_month_yr      = 36
const update_futue_chkbx    = 37
const order_ubound          = 38
const self_emp_mthd_conv    = 39
const cash_mos_list         = 40
const pick_one              = 41
const selection_rsn         = 42
const ignore_antic          = 43
const antic_pay_list        = 44
const update_this_month     = 45
const last_paycheck         = 46
const panel_first_check     = 47
const this_is_a_new_panel   = 48
const days_of_verif         = 49
const months_updated        = 50
const income_lumped_mo      = 51
const lump_reason           = 52
const act_checks_lumped     = 53
const est_checks_lumped     = 54
const lump_gross            = 55
const lump_hrs              = 56
const excl_cash_rsn         = 57
const GRH_mo_inc            = 58
const spoke_to              = 59
const employer_with_underscores = 60
const bimonthly_first       = 61
const bimonthly_second      = 62
const hc_budg_notes         = 63
const hc_retro              = 64
const convo_detail          = 65

'Constants to make an option selection easier to read.
const use_actual        = 1
const use_estimate      = 2

'Constants for the array that handles each income - LIST_OF_INCOME_ARRAY
const panel_indct           = 0
const pay_date              = 1
const gross_amount          = 2
const hours                 = 3
const budget_in_SNAP_yes    = 4
const budget_in_SNAP_no     = 5
const reason_to_exclude     = 6
const exclude_amount        = 7
const check_order           = 8
const view_pay_date         = 9
const frequency_issue       = 10
const future_check          = 11
const reason_amt_excluded   = 12

'Constants for the array of the cash months - CASH_MONTHS_ARRAY
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

original_month = MAXIS_footer_month     'setting these to a seperate variable for the dialog
original_year = MAXIS_footer_year

future_months_check = checked           'default to having th script update future months

'INITIAL Dialog - case number, footer month, worker signature
BeginDialog Dialog1, 0, 0, 191, 220, "Case Number"
  EditBox 90, 5, 70, 15, MAXIS_case_number
  EditBox 100, 25, 15, 15, original_month
  EditBox 120, 25, 15, 15, original_year
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

'calling the dialog
Do
    Do
        err_msg = ""
        dialog Dialog1
        cancel_confirmation

        If IsNumeric(MAXIS_case_number) = FALSE or Len(MAXIS_case_number) > 8 Then err_msg = err_msg & vbNewLine & "* Enter a valid case number."       'confirming a valid case number
        If trim(worker_signature) = "" Then err_msg = err_msg & vbNewLine & "* Enter your worker signature for your case notes."                        'confirming there is a worker signature

        original_month = trim(original_month)       'cleaning up the entry here
        original_year = trim(original_year)
        If len(original_year) <> 2 or len(original_month) <> 2 Then err_msg = err_msg & vbNewLine & "* Enter a 2 digit footer month and year."          'forcing 2 digit month and year to be entered

        If err_msg <> "" Then MsgBox "-- Please resolve the following to continue --" & vbNewLine & err_msg                                             'displaying the error handling
    Loop until err_msg = ""
    call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

MAXIS_footer_month = original_month     'setting the footer month and year back to what was entered in the dialog.
MAXIS_footer_year = original_year       'this is split out for the option of having seperate handling prior to the reassignment for working in current month if needed

Call back_to_SELF                       'need to gather some detail to have the correct script run

developer_mode = FALSE                  'allowing worker to exit if started in Inquiry on accident
EMReadScreen MX_region, 12, 22, 48
MX_region = trim(MX_region)
If MX_region = "INQUIRY DB" Then
    continue_in_inquiry = MsgBox("It appears you are in INQUIRY. Income information cannot be saved to STAT and a CASE/NOTE cannot be created." & vbNewLine & vbNewLine & "Do you wish to continue?", vbQuestion + vbYesNo, "Continue in Inquiry?")
    If continue_in_inquiry = vbNo Then script_end_procedure("Script ended since it was started in Inquiry.")
    developer_mode = TRUE
End If
If developer_mode = TRUE then MsgBox "Developer Mode ACTIVATED!"        'developer mode difference is that the MAXIS update detail is shown in a messagebox instead of updating the panel
If developer_mode = TRUE Then script_run_lowdown = "Run in INQUIRY"

Call navigate_to_MAXIS_screen("CASE", "CURR")                           'Going to find the FS Application Date
curr_row = 1
curr_col = 1
EMSearch " FS:", curr_row, curr_col
If curr_row <> 0 Then
    EMReadScreen fs_prog_status, 7, curr_row, curr_col + 5
    fs_prog_status = trim(fs_prog_status)
    If fs_prog_status = "ACTIVE" OR fs_prog_status = "PENDING" Then
        EMReadScreen fs_appl_date, 8, curr_row, curr_col + 25
    End If
End If
fs_appl_footer_month = left(fs_appl_date, 2)                            'Making a footer month and year for the FS Application'
fs_appl_footer_year = right(fs_appl_date, 2)

Do                                                                      'Getting in to STA (making sure we get past background)
    Call navigate_to_MAXIS_screen ("STAT", "SUMM")
    EMReadScreen summ_check, 4, 2, 46
Loop until summ_check = "SUMM"

CASH_case = FALSE       'defining these as a default
SNAP_case = FALSE
HC_case = FALSE

Call Navigate_to_MAXIS_screen("STAT", "PROG")                           'Getting program status to identify potential programs the income should apply to

EMReadScreen cash_one_status, 4, 6, 74                                  'reading each program status
EMReadScreen cash_two_status, 4, 7, 74
EMReadScreen grh_status, 4, 9, 74
EMReadScreen snap_status, 4, 10, 74
EMReadScreen hc_status, 4, 12, 74

If cash_one_status = "ACTV" OR cash_one_status = "PEND" Then CASH_case = TRUE   'setting programs to TRUE based on PROG status
If cash_two_status = "ACTV" OR cash_two_status = "PEND" Then CASH_case = TRUE
If grh_status = "ACTV" OR grh_status = "PEND" Then GRH_case = TRUE
If snap_status = "ACTV" OR snap_status = "PEND" Then SNAP_case = TRUE
If hc_status = "ACTV" OR hc_status = "PEND" Then HC_case = TRUE

                            '----------------------------------------------------------'
                    '---------------------------------------------------------------------------------'
'-------------------------------------------------FINDING ALL CURRENT EI PANELS --------------------------------------------------'
                    '---------------------------------------------------------------------------------'
                            '----------------------------------------------------------'

the_panel = 0                       'this is our counter to add new panels to the EARNED_INCOME_PANELS_ARRAY

call HH_member_custom_dialog(HH_member_array)   'finding who should be looked at for income on the case
'FUTURE FUNCTIONALITY - Stop work should be added in before we add information to the EARNED_INCOME_PANELS_ARRAY

Call navigate_to_MAXIS_screen("STAT", "JOBS")       'Starting with JOBS panels
For each member in HH_member_array                  'We are going to look at each HH member checked in the HH_member dialog
    EMWriteScreen member, 20, 76                    'going to the member in JOBS
    Transmit

    EMReadScreen number_of_jobs_panels, 1, 2, 78    'finding the total number of panels currently existing for the current member.

    If number_of_jobs_panels <> "0" Then            'if there are 0 panels we don't need to do anything else in JOBS for this member
        number_of_jobs_panels = number_of_jobs_panels * 1       'making the number read and actual number

        For panel = 1 to number_of_jobs_panels      'we are going to cycle through each of the panels for this member
            EMWriteScreen "0" & panel, 20, 79       'navigating to the panel instance
            transmit
            'FUTURE FUNCTIONALITY - Stop work should be added in before we add information to the EARNED_INCOME_PANELS_ARRAY

            save_this_panel = TRUE                  'we are always at this point going to save the panel to the EARNED_INCOME_PANELS_ARRAY
                                                    'FUTURE FUNCTIONALITY where we may be deleting old panels, in which case we would NOT be saving the panel to the array

            EMReadScreen end_date, 8, 9, 49         'finding the end date

            If end_date <> "__ __ __" Then
                end_date = replace(end_date, " ", "/")
                end_date = DateValue(end_date)

                'FUTURE FUNCTIONALITY - here is a start with deleting an already ended panel - this has not been tested or completed
                'TESTING NEEDED
                ' If DateDiff("m", end_date, date) > 3 Then
                '
                '     BeginDialog Dialog1, 0, 0, 186, 140, "Dialog"
                '       OptionGroup RadioGroup1
                '         RadioButton 20, 45, 70, 10, "Delete this Panel", delete_panel
                '         RadioButton 20, 60, 75, 10, "Leave this Panel", leave_ended_panel
                '       EditBox 10, 100, 170, 15, explain_leaving_ended_panel
                '       ButtonGroup ButtonPressed
                '         OkButton 130, 120, 50, 15
                '       Text 10, 10, 170, 25, "This JOBS panel indicates the income ended more than 3 months ago. This panel is no longer needed in this month since this income has ended."
                '       Text 10, 80, 115, 20, "If this ened panel is to be left, explain why it is still needed:"
                '     EndDialog
                '
                '     Do
                '         Do
                '             err_msg = ""
                '
                '             Dialog Dialog1
                '
                '             explain_leaving_ended_panel = trim(explain_leaving_ended_panel)
                '
                '             If leave_ended_panel = checked and explain_leaving_ended_panel = "" Then err_msg = err_msg & vbNewLine & "* If an ended panel is to be left on an active case, explain why it is still needed."
                '
                '             If err_msg <> "" Then MsgBox "** Please Resolve to Continue **" & vbNewLine & err_msg
                '         Loop Until err_msg = ""
                '         call check_for_password(are_we_passworded_out)
                '     Loop until are_we_passworded_out = false
                '
                '     If delete_panel = checked then panels_to_delete = panels_to_delete & "~" & "JOBS " & member & " " & "0" & panel
                '
                ' End If
            End If

            If save_this_panel = TRUE Then                                      'if the panel will be saved (always for now) then we are going to read panel detail.
                ReDim Preserve EARNED_INCOME_PANELS_ARRAY(convo_detail, the_panel)          'resizing the array

                'Setting known information and defaults
                EARNED_INCOME_PANELS_ARRAY(panel_type, the_panel) = "JOBS"                  'all in this loop are JOBS
                EARNED_INCOME_PANELS_ARRAY(panel_member, the_panel) = member                'member known from member array
                EARNED_INCOME_PANELS_ARRAY(panel_instance, the_panel) = "0" & panel         'instance known from the for-next of all panels for this member
                EARNED_INCOME_PANELS_ARRAY(income_received, the_panel) = FALSE              'default this to false, user will inidcate if income is received later
                If CASH_case = TRUE Then EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, the_panel) = checked     'These are defaulted by whatever program is active or pending - will be able to be changed later
                If SNAP_case = TRUE Then EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, the_panel) = checked
                If HC_case = TRUE Then EARNED_INCOME_PANELS_ARRAY(apply_to_HC, the_panel) = checked
                If GRH_case = TRUE Then EARNED_INCOME_PANELS_ARRAY(apply_to_GRH, the_panel) = checked

                'Reading the information from the panel
                'FUTURE FUNCTIONALITY - add ability to read current income from the panel/PIC etc. so that partial work can be screen scraped instead of having to retype it
                EMReadScreen type_of_job, 1, 5, 34
                EMReadScreen job_verif, 25, 6, 34
                EMReadScreen listed_hrly_wage, 6, 6, 75
                EMReadScreen employer_name, 30, 7, 42
                EMReadScreen start_date, 8, 9, 35
                EMReadScreen end_date, 8, 9, 49
                EMReadScreen frequency, 1, 18, 35
                EMReadScreen current_verif, 27, 6, 34

                If type_of_job = "J" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "J - WIOA"       'setting the full detail to the array instead of a single letter code
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

                'formatting the information from the panel and adding it to the EARNED_INCOME_PANELS_ARRAY
                EARNED_INCOME_PANELS_ARRAY(income_verif, the_panel) = trim(job_verif)
                EARNED_INCOME_PANELS_ARRAY(employer, the_panel) = replace(employer_name, "_", "")
                EARNED_INCOME_PANELS_ARRAY(employer_with_underscores, the_panel) = employer_name
                EARNED_INCOME_PANELS_ARRAY(hourly_wage, the_panel) = trim(listed_hrly_wage)
                EARNED_INCOME_PANELS_ARRAY(income_start_dt, the_panel) = replace(start_date, " ", "/")
                EARNED_INCOME_PANELS_ARRAY(income_end_dt, the_panel) = replace(end_date, " ", "/")
                If EARNED_INCOME_PANELS_ARRAY(income_start_dt, the_panel) = "__/__/__" Then EARNED_INCOME_PANELS_ARRAY(income_start_dt, the_panel) = ""
                If EARNED_INCOME_PANELS_ARRAY(income_end_dt, the_panel) = "__/__/__" Then EARNED_INCOME_PANELS_ARRAY(income_end_dt, the_panel) = ""
                EARNED_INCOME_PANELS_ARRAY(old_verif, the_panel) = trim(current_verif)
                If frequency = "1" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "1 - One Time Per Month"      'setting full detail to the array instead of a single letter code
                If frequency = "2" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "2 - Two Times Per Month"
                If frequency = "3" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "3 - Every Other Week"
                If frequency = "4" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "4 - Every Week"
                If frequency = "5" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "5 - Other"

                EARNED_INCOME_PANELS_ARRAY(income_list_indct, the_panel) = "NONE"       'This is where all of the array items from LIST_OF_INCOME_ARRAY will be added that are associated with this panel
                EARNED_INCOME_PANELS_ARRAY(this_is_a_new_panel, the_panel) = FALSE      'identifies if a panel was created by the script or not - these are currently existing - changes CNote

                the_panel = the_panel + 1       'incrementing our counter to be ready for the next panel/member/income type
            End If      'If save_this_panel = TRUE Then
        Next            'For panel = 1 to number_of_jobs_panels
    End If              'If number_of_jobs_panels <> "0" Then
Next                    'For each member in HH_member_array

'FUTURE FUNCTIONALITY - this will remove panels if ended long enough ago.
'TESTING NEEDED
If panels_to_delete <> "" Then                              'at this time panels_to_delete wil ALWAYS = "" so no need to comment out
    array_of_ended_panels = split(panels_to_delete, "~")    'NEED COMMENTS
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

'Now we will repeat looking at each panel for each member but in BUSI
'At this time the script fully reads any BUSI panels but functionality has not been created to add new or update any BUSI panels
Call navigate_to_MAXIS_screen("STAT", "BUSI")           'going to BUSI
For each member in HH_member_array                      'looking at each member from member_array - checked by worker in HH_member dialog
    EMWriteScreen member, 20, 76                        'navigating to the correct member
    Transmit

    EMReadScreen number_of_busi_panels, 1, 2, 78        'reading the total panels to look at for that member

    If number_of_busi_panels <> "0" Then                'if there is at least 1 panel, we will loop through all of the panels to gather detail
        number_of_busi_panels = number_of_busi_panels * 1       'making this an actual number and not a string

        For panel = 1 to number_of_busi_panels          'looping through each of the panels
            EMWriteScreen "0" & panel, 20, 79           'navigating to the panel on the current loop
            transmit

            ReDim Preserve EARNED_INCOME_PANELS_ARRAY(convo_detail, the_panel)      'resizing the array

            'Setting the known information and defaults to the EARNED_INCOME_PANELS_ARRAY'
            EARNED_INCOME_PANELS_ARRAY(panel_type, the_panel) = "BUSI"              'This is a BUSI panel
            EARNED_INCOME_PANELS_ARRAY(panel_member, the_panel) = member            'the member we are currently reviewing
            EARNED_INCOME_PANELS_ARRAY(panel_instance, the_panel) = "0" & panel     'the panel instance
            EARNED_INCOME_PANELS_ARRAY(income_received, the_panel) = FALSE          'this is always defaulted to FALSE, user indicating there is income to update will change this to TRUE
            If CASH_case = TRUE Then EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, the_panel) = checked     'defaulting these based on the programs currently pending or active
            If SNAP_case = TRUE Then EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, the_panel) = checked
            If HC_case = TRUE Then EARNED_INCOME_PANELS_ARRAY(apply_to_HC, the_panel) = checked
            If GRH_case = TRUE Then EARNED_INCOME_PANELS_ARRAY(apply_to_GRH, the_panel) = checked

            'Reading information from the panel
            EMReadScreen type_of_busi, 2, 5, 37
            EMReadScreen start_date, 8, 5, 55
            EMReadScreen end_date, 8, 5, 72
            EMReadScreen listed_method, 2, 16, 53
            EMReadScreen lst_mthd_date, 8, 16, 63

            'Formatting and adding information from the panel to the EARNED_INCOME_PANELS_ARRAY'
            If type_of_busi = "01" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "01 - Farming"             'Updating BUSI type to have full detail instead of  a 2-digit code
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
            If EARNED_INCOME_PANELS_ARRAY(income_start_dt, the_panel) = "__/__/__" Then EARNED_INCOME_PANELS_ARRAY(income_start_dt, the_panel) = ""
            If EARNED_INCOME_PANELS_ARRAY(income_end_dt, the_panel) = "__/__/__" Then EARNED_INCOME_PANELS_ARRAY(income_end_dt, the_panel) = ""
            If listed_method = "01" Then EARNED_INCOME_PANELS_ARRAY(self_emp_mthd, the_panel) = "01 - 50% Gross Inc"
            If listed_method = "02" Then EARNED_INCOME_PANELS_ARRAY(self_emp_mthd, the_panel) = "02 - Tax Forms"
            EARNED_INCOME_PANELS_ARRAY(method_date, the_panel) = replace(lst_mthd_date, " ", "/")
            If EARNED_INCOME_PANELS_ARRAY(method_date, the_panel) = "__/__/__" Then EARNED_INCOME_PANELS_ARRAY(method_date, the_panel) = ""

            EmWriteScreen "X", 6, 26        'opening the GROSS INCOME CALCULATION pop-up
            transmit
            'TESTING NEEDED'
            For busi_row = 9 to 19          'BUGGY CODE - this doesn't seem quite right - maybe there are different verifs for each type
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

            EARNED_INCOME_PANELS_ARRAY(income_list_indct, the_panel) = "NONE"       'defaulting this, items from LIST_OF_INCOME_ARRAY will be added here as they are created
            EARNED_INCOME_PANELS_ARRAY(this_is_a_new_panel, the_panel) = FALSE      'this indicates if a panel was created by the script - these were all existing prior to the script run

            the_panel = the_panel + 1                   'incrementing our counter to be ready for the next panel/member/income type
        Next        'For panel = 1 to number_of_busi_panels
    End If      'If number_of_busi_panels <> "0" Then
Next        'For each member in HH_member_array

'FUTURE FUNCTIONALITY - add gethering RBIC panels to the EARNED_INCOME_PANELS_ARRAY

                            '----------------------------------------------------------'
                    '---------------------------------------------------------------------------------'
'------------------------------------------------- ADDING NEW EI PANELS --------------------------------------------------'
                    '---------------------------------------------------------------------------------'
                            '----------------------------------------------------------'

'Here we are allowing the user to add a new panel if needed. It is on a loop so as many as desired can be added
add_panel_button_pushed_count = 0
Do
    panels_exist = TRUE     'default to panels existing - this is also reset on each loop so that if a panel is added this statement is reassessed.
    'THIS VARIABLE SET TO FALSE WILL CAUSE THE SCRIPT TO END AFTER THIS LOOP

    y_pos = 25              'setting coordinates for the dialog to be created - this is the vertical position in the dialog
    dlg_len = 15 * UBOUND(EARNED_INCOME_PANELS_ARRAY, 2) + 15 * UBOUND(HH_member_array) + 125       'creating the height of the dialog

    'ASK TO ADD NEW PANEL Dailog - lists all current panels, Yes/No question about adding another
    BeginDialog Dialog1, 0, 0, 390, dlg_len, "Do you want to add a new JOBS or BUSI Panel?"

      Text 5, 10, 105, 10, "Known JOBS and BUSI panels:"        'This part lists the current panels and will change each time through the loop as new panels are added '

      'Looping through the EARNED_INCOME_PANELS_ARRAY to get all panels found earlier
      For ei_panel = 0 to UBOUND(EARNED_INCOME_PANELS_ARRAY, 2)
        'compiling the information about the panel here to make it more readble and specific to the PANEL detail to include
        earned_income_panel_detail = EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel) & " " & EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel) & " " & EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel) & " - "
        If EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel) = "JOBS" Then               'JOBS has employer information
            earned_income_panel_detail = earned_income_panel_detail & EARNED_INCOME_PANELS_ARRAY(employer, ei_panel)
        ElseIf EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel) = "BUSI" Then           'BUSI has type of self employment information
            earned_income_panel_detail = earned_income_panel_detail & "TYPE: " & EARNED_INCOME_PANELS_ARRAY(income_type, ei_panel)
        End If
        earned_income_panel_detail = earned_income_panel_detail & " - Income Start: " & EARNED_INCOME_PANELS_ARRAY(income_start_dt, ei_panel) & " - Verif: " & EARNED_INCOME_PANELS_ARRAY(old_verif, ei_panel)

        If EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel) = "" Then       'if the panel_type position is blank that means that we never incremented the array counter and there are NO JOBS or BUSI panels
            earned_income_panel_detail = "** THERE ARE NO EARNED INCOME PANELS ON THIS CASE **"     'special detial if there are no panels to make the dialog clearer
            panels_exist = FALSE        'setting this so the script knows to end outside of the loop if no panels are added.
        End If
        Text 10, y_pos, 375, 10, earned_income_panel_detail     'here is where we actually list the information about the panel on the dialog
        y_pos = y_pos + 15                                      'incrementing this placeholder so that panel information is not on top of each other in the dialog
      Next
      y_pos = y_pos + 5     'now we move down a little more in the dialog
      Text 5, y_pos, 295, 10, "These are all the panels that are currently known in MAXIS for these Household Members:" 'listing all the household members we looked at in gathering panel information
      y_pos = y_pos + 15
      For each member in HH_member_array
        Text 10, y_pos, 45, 10, "Member " & member
        y_pos = y_pos + 15
        'Text 10, 75, 45, 10, "MEMBER 01"
      Next
      y_pos = y_pos - 10
      ' Text 80, y_pos, 160, 10, "Do you need to add a new JOBS or BUSI panel?"     'FUTURE FUNCTIONALITY - Saving this for when BUSI is added
      Text 115, y_pos, 160, 10, "Do you need to add a new JOBS panel?"
      ButtonGroup ButtonPressed             'NO CANCEL button on this dialog - pressing the 'X' in the corner will default to the 'No' button pressed
        PushButton 120, y_pos + 15, 140, 20, "Yes - Add a new Earned Income panel", add_new_panel_button
        PushButton 120, y_pos + 35, 140, 20, "No - The panel(s) to update are in MAXIS", continue_to_update_button
    EndDialog

    dialog Dialog1          'displaying this dialog - no error handling because no input fields - just a Yes/No
                            'vbYesNo MsgBox NOT used because we want to allow MAXIS navigation in this time.

    'Pushing the 'Yes' button on the dialog causes this code to be used - pressing 'No' will skip this
    If buttonpressed = add_new_panel_button Then
        original_month = MAXIS_footer_month     'saving the MAXIS footer month and year because we may move around as the panels are added in the month the income started OR application month
        original_year = MAXIS_footer_year
        panel_to_add = "JOBS"                   'defaulting this to JOBS because that is actually the only option right now
        add_panel_button_pushed_count = add_panel_button_pushed_count + 1

        'TYPE OF PANEL TO ADD Dialog - Select panel type
        'FUTURE FUNCTIONALITY - add BUSI back as an option to select here
        BeginDialog Dialog1, 0, 0, 191, 50, "Panel Type to Add"
          ' DropListBox 30, 30, 60, 45, "Select one..."+chr(9)+"JOBS"+chr(9)+"BUSI", panel_to_add
          DropListBox 30, 30, 60, 45, "Select one..."+chr(9)+"JOBS", panel_to_add
          ButtonGroup ButtonPressed
            OkButton 135, 10, 50, 15
            CancelButton 135, 30, 50, 15
          Text 15, 10, 85, 20, "Which type of panel would you like to add?"
        EndDialog

        cancel_clarify = ""     'resetting this here - this supports canceling the new job add without cancelling the script
        Do
            Do
                err_msg = ""

                dialog Dialog1

                If ButtonPressed = 0 then       'this is the cancel button
                    'this is an upgrade in functionality from cancel_confirmation to asking if we should cancel the current function or the script
                    cancel_clarify = MsgBox("Do you want to stop the script entirely?" & vbNewLine & vbNewLine & "If the script is stopped no actions taken so far will be noted.", vbQuestion + vbYesNo, "Clarify Cancel")
                    If cancel_clarify = vbYes Then script_end_procedure("~PT: user pressed cancel")
                    'script_end_procedure text added for statistical purposes. If script was canceled prior to completion, the statistics will reflect this.
                End if
                If cancel_clarify = vbNo Then       'This means to cancel the current operation but not the script
                    panel_to_add = ""               'blanking this out so a panel is not added
                    Exit Do                         'leaving the loop for the dialog
                End If

                If panel_to_add = "Select one..." Then err_msg = err_msg & vbNewLine & "* Indicate which type of panel needs to be added."      'error handling - must select a panel type

                If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg                                                'showing the error message

            Loop until err_msg = ""
            call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
        LOOP UNTIL are_we_passworded_out = false

        info_saved = FALSE          'defaulting if the information about the new panel has been saved to the EARNED_INCOME_PANELS_ARRAY so we don't accidentally save it twice

        If CASH_case = TRUE Then cash_checkbox = checked
        If SNAP_case = TRUE Then snap_checkbox = checked
        If HC_case = TRUE Then hc_checkbox = checked
        If GRH_case = TRUE Then grh_checkbox = checked

        Select Case panel_to_add            'This will operate specific code based on if a JOBS or BUSI panel is to be added.

        Case "JOBS"
            'NEW JOB PANEL Dialog'
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
              CheckBox 230, 95, 30, 10, "GRH", grh_checkbox
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

            cancel_clarify = ""         'blanking this out from previous dialog or another loop
            yes_create_panel = TRUE     'defaulting to having the panel created
            Do
                Do
                    err_msg = ""

                    dialog Dialog1

                    'alternate for cancel_confirmation
                    If ButtonPressed = 0 then       'this is the cancel button
                        cancel_clarify = MsgBox("Do you want to stop the script entirely?" & vbNewLine & vbNewLine & "If the script is stopped no actions taken so far will be noted.", vbQuestion + vbYesNo, "Clarify Cancel")
                        If cancel_clarify = vbYes Then script_end_procedure("~PT: user pressed cancel")     'ends the script entirely
                        'script_end_procedure text added for statistical purposes. If script was canceled prior to completion, the statistics will reflect this.
                    End if
                    If cancel_clarify = vbNo Then       'cancels the current operation without cancelling the script
                        yes_create_panel = FALSE        'this keeps a blank panel from being created if 'Cancel' is selected
                        Exit Do
                    End If

                    'Error handling
                    enter_JOBS_employer = trim(enter_JOBS_employer)
                    If trim(enter_JOBS_clt_ref_nbr) = "" Then err_msg = err_msg & vbNewLine & "* Enter the member number of the client that is employed at this job."               'Need a member reference number
                    If len(enter_JOBS_clt_ref_nbr) <> 2 Then err_msg = err_msg & vbNewLine & "* The member number should be 2 digits"                                               'Ensuring the member reference number is accurate
                    If trim(enter_JOBS_inc_type_code) = "  " Then err_msg = err_msg & vbNewLine & "* Enter the income type of the job."                                             'Need JOB income type - usually 'W'
                    If trim(enter_JOBS_verif_code) = "  " Then err_msg = err_msg & vbNewLine & "* Enter the verification code for this job."                                        'MAXIS requires some kind of verification code - 'N' is okay here
                    If enter_JOBS_employer = "" Then err_msg = err_msg & vbNewLine & "* Enter the employer name for this job."                                                      'Must have an employer name to enter into MAXIS
                    If len(enter_JOBS_employer) > 30 Then err_msg = err_msg & vbNewLine & "* The Employer name is too long to fit on the JOBS panel, abbreviate as necessary."      'Employer name line on JOBS is only a certain length, conforming to that length
                    If IsDate(enter_JOBS_start_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date for Income start date."                                     'Requiring a start date for income since SNAP eligibility results can only be created if we have this
                    If trim(enter_JOBS_end_date) <> "" AND IsDate(enter_JOBS_end_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date for Income end date."     'IF there is an end date, it must be a valid date

                    if err_msg <> "" Then msgBox "Please resolve the following to continue:" & vbNewLine & err_msg          'displaying the error messages if there are any

                Loop until err_msg = ""
                call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
            LOOP UNTIL are_we_passworded_out = false

            If yes_create_panel = TRUE Then                     'only continues if cancel was not selected above
                Call navigate_to_MAXIS_screen("CASE", "CURR")   'We need the most recent case application date because we cannot go further back than that month/year

                EMReadScreen appl_date, 8, 8, 29                'this is where the case application date is - NOT program specific

                If DateDiff("m", appl_date, enter_JOBS_start_date) >= 0 Then    'if the application date is before the income start date
                    beginning_month = DatePart("m", enter_JOBS_start_date)      'we use the income start date as the footer month and year to enter the JOBS panel
                    beginning_year = DatePart("yyyy", enter_JOBS_start_date)
                    first_check = enter_JOBS_start_date                             'setting the date of the first check to be entered on JOBS
                Else                                                            'otherwise, if the job started before the application date
                    beginning_month = DatePart("m", appl_date)                  'we use the application date as the footer month and year to enter the JOBS panel
                    beginning_year = DatePart("yyyy", appl_date)
                    first_check = beginning_month & "/01/" & beginning_year         'setting the date of the first check to be entered on JOBS
                End If

                beginning_month = right("00"&beginning_month, 2)                'creating 2 digit month and year variables
                beginning_year = right(beginning_year, 2)

                If DateDiff("m", first_check, date) > 12 Then                   'if the first check to be entered on the panel is more that 12 months from the current date
                                                                                'script will confirm the month and year to add the panel in MAXIS
                    'PROCEDURE CLARIFICATION - allowing workers to adjust the month and year the panel is entered
                    'CONFIRM ADD PANEL MONTH Dialog'
                    BeginDialog Dialog1, 0, 0, 191, 175, "Confirm Update Month"
                      EditBox 140, 60, 15, 15, beginning_month
                      EditBox 160, 60, 15, 15, beginning_year
                      ButtonGroup ButtonPressed
                        OkButton 135, 155, 50, 15
                      Text 10, 10, 165, 10, "** This new job started at least 12 months ago. **"
                      Text 10, 30, 165, 20, "The script will go back to " & beginning_month & "/" & beginning_year & " to add this JOBS panel in the month the job started."
                      Text 10, 60, 120, 15, "If this needs to be adjusted, change the footer month and year here:"
                      GroupBox 10, 85, 170, 65, "Info"
                      Text 20, 100, 150, 40, "Best practice is to add the job information in the footer month and year the income started. An exception may be if the job was currently in STAT and the panel was deleted, only add the information in the first month of deletion."
                    EndDialog

                    dialog Dialog1

                    beginning_month = beginning_month * 1       'making these numbers for error handling
                    beginning_year = beginning_year * 1

                    If beginning_month = DatePart("m", enter_JOBS_start_date) AND beginning_year = DatePart("yyyy", enter_JOBS_start_date) Then     'setting the first check to be entered on JOBS
                        first_check = enter_JOBS_start_date
                    Else
                        first_check = beginning_month & "/01/" & beginning_year
                    End If

                    beginning_month = right("00"&beginning_month, 2)            'making these 2-digit variables
                    beginning_year = right(beginning_year, 2)

                End If

                MAXIS_footer_month = beginning_month        'setting the footer month and year to the update month so nav functions work
                MAXIS_footer_year = beginning_year

                Call back_to_SELF           'Getting out of STAT to be sure footer month is correct

                Call navigate_to_MAXIS_screen("STAT", "SUMM")

                Do                          'this loop is to update future months with the JOB information
                    STATS_manualtime = STATS_manualtime + 85
                    If info_saved = FALSE Then          'If the information has not yet been saved to the array it means we are in the first month
                        EMWriteScreen "JOBS", 20, 71                    'go to JOBS
                        EMWriteScreen enter_JOBS_clt_ref_nbr, 20, 76    'go to the right member
                        EMWriteScreen "NN", 20, 79                      'create new JOBS panel

                        transmit
                    Else                                'If the information is in the array, we will use that to navigate
                        EMWriteScreen "JOBS", 20, 71                    'go to JOBS
                        EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_member, the_panel-1), 20, 76     'here we use 'the_panel-1' because it would have been incremented on the previous loop after saving the information
                        EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_instance, the_panel-1), 20, 79

                        transmit

                        EMReadScreen check_for_panel, 14, 24, 13                'sometimes the panel does not exist in a future month because data expires, we then need to add it again
                        If check_for_panel = "DOES NOT EXIST" Then
                            EMWriteScreen "JOBS", 20, 71
                            EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_member, the_panel-1), 20, 76
                            EMWriteScreen "NN", 20, 79

                            transmit
                        Else
                            PF9             'other wise put it in EDIT MODE
                        End If
                    End If

                    EMWriteScreen left(enter_JOBS_inc_type_code, 1), 5, 34      'adding all the information from the dialog - using only the codes which are stored in the leftmost portion of the variables
                    EMWriteScreen left(enter_JOBS_subsdzd_inc_type, 2), 5, 74
                    EMWriteScreen left(enter_JOBS_verif_code, 1), 6, 34
                    EMWriteScreen "      ", 6, 75
                    EMWriteScreen enter_JOBS_hrly_wage, 6, 75
                    EMWriteScreen enter_JOBS_employer, 7, 42

                    Call write_date(enter_JOBS_start_date, "MM DD YY", 9, 35)   'entering the start date
                    If trim(enter_JOBS_end_date) <> "" Then                     'TESTING NEEDED - I do not believe this functionality has been well used in testing
                        Call write_date(enter_JOBS_end_date, "MM DD YY", 9, 49) 'entering the end date if one was listed
                        If DateDiff("d", first_check, enter_JOBS_end_date) >= 0 Then    'as long as the end date is after the date of the check to entering - the ceck is entered with $0 pay amount
                            Call write_date(first_check, "MM DD YY", 12, 54)
                            EMWriteScreen "    0.00", 12, 67
                            EMWriteScreen "0  ", 18, 72
                        Else                                                    'otherwise the pay information is blanked out
                            EMWriteScreen "  ", 12, 54
                            EMWriteScreen "  ", 12, 57
                            EMWriteScreen "  ", 12, 60
                            EMWriteScreen "        ", 12, 67
                            EMWriteScreen "   ", 18, 72
                        End If
                    Else                                                        'if there is no end date entered, the first pay date and $0 pay is entered
                        Call write_date(first_check, "MM DD YY", 12, 54)
                        EMWriteScreen "    0.00", 12, 67
                        EMWriteScreen "0  ", 18, 72
                    End If
                    transmit
                    EMReadScreen check_for_error_prone_warning, 20, 6, 43       'some times some warnings come up - need to move from them
                    If check_for_error_prone_warning = "Error Prone Warnings" Then transmit
                    'MsgBox "Pause after check for error"
                    EMReadScreen new_panel, 1, 2, 73            'reading the panel instance that was created. It is only generated AFTER transmit has saved the entered information

                    If info_saved = FALSE Then                                  'If we have not already saved the information to EARNED_INCOME_PANELS_ARRAY - this will do it here
                        ReDim Preserve EARNED_INCOME_PANELS_ARRAY(convo_detail, the_panel)      'resizing the array

                        EARNED_INCOME_PANELS_ARRAY(panel_type, the_panel) = "JOBS"
                        EARNED_INCOME_PANELS_ARRAY(panel_member, the_panel) = enter_JOBS_clt_ref_nbr
                        EARNED_INCOME_PANELS_ARRAY(panel_instance, the_panel) = "0" & new_panel
                        EARNED_INCOME_PANELS_ARRAY(income_received, the_panel) = FALSE
                        EARNED_INCOME_PANELS_ARRAY(initial_month_mo, the_panel) = MAXIS_footer_month    'defaulting the first date to update to the month/year the panel was created
                        EARNED_INCOME_PANELS_ARRAY(initial_month_yr, the_panel) = MAXIS_footer_year

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
                        EARNED_INCOME_PANELS_ARRAY(employer_with_underscores, the_panel) = employer_name
                        EARNED_INCOME_PANELS_ARRAY(hourly_wage, the_panel) = trim(listed_hrly_wage)
                        EARNED_INCOME_PANELS_ARRAY(income_start_dt, the_panel) = replace(start_date, " ", "/")
                        EARNED_INCOME_PANELS_ARRAY(income_end_dt, the_panel) = replace(end_date, " ", "/")
                        If EARNED_INCOME_PANELS_ARRAY(income_end_dt, the_panel) = "__/__/__" Then EARNED_INCOME_PANELS_ARRAY(income_end_dt, the_panel) = ""
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
                        EARNED_INCOME_PANELS_ARRAY(apply_to_GRH, the_panel) = grh_checkbox

                        EARNED_INCOME_PANELS_ARRAY(this_is_a_new_panel, the_panel) = TRUE       'since this is a new panel - it is saved here - this is mostly for the CNote

                        the_panel = the_panel + 1       'incrementing the counter

                        info_saved = TRUE               'changing this variable for the next loop
                    End If

                    'Navigates to the current month + 1 footer month without sending the case through background
                    CALL write_value_and_transmit("BGTX", 20, 71)
                    CALL write_value_and_transmit("y", 16, 54)

                    EMReadScreen all_months_check, 24, 24, 2    'this reads the error message at the bottom of STAT/WRAP if we cannot get to the next month because we are in CM+1
                    EMReadScreen MAXIS_footer_month, 2, 20, 55  'If we are successful in getting to the next month, the footer month and year are set here
                    EMReadScreen MAXIS_footer_year, 2, 20, 58

                    first_check = MAXIS_footer_month & "/01/" & MAXIS_footer_year   'need a check date in the current footer month to enter on JOBS
                Loop until all_months_check = "CONTINUATION NOT ALLOWED"
                PF3 'leaving STAT - sending the case through background
            End If      'If yes_create_panel = TRUE Then'

        Case "BUSI"     'Option if BUSI is selected
        'FUTURE FUNCTIONALITY - create a new BUSI panel

        End Select

        enter_JOBS_clt_ref_nbr = ""
        enter_JOBS_inc_type_code = ""
        enter_JOBS_subsdzd_inc_type = ""
        enter_JOBS_verif_code = ""
        enter_JOBS_hrly_wage = ""
        enter_JOBS_employer = ""
        enter_JOBS_start_date = ""
        enter_JOBS_end_date = ""
        snap_checkbox = ""
        cash_checkbox = ""
        hc_checkbox = ""
        grh_checkbox = ""

        MAXIS_footer_month = original_month     'resetting the footer month and year to what was indicated in the initial dialog
        MAXIS_footer_year = original_year
        Call back_to_SELF
    End If      'If buttonpressed = add_new_panel_button Then
    'There is nothing specific that happens if the 'NO' or continue_to_update button is pushed other than leaving this portion of the functionality to the next portion

'this loop until functionality allows for as many JOBS/BUSI panels to be added as needed.
'Also, note that the new panel will be in the array and so will be added to the dialog asking about adding a new panel
Loop until buttonpressed = continue_to_update_button
script_run_lowdown = script_run_lowdown & vbCr & "Add job button pressed " & add_panel_button_pushed_count & " times."

'If there were no panels and noe were added, the script will stop, alerting the worker that there are no panels to take action on.
'BUGGY CODE - we might have an issue if ONLY BUSI panels exist. We may not get a script end here but no other code would be enacted.
If panels_exist = FALSE Then script_end_procedure("There are no earned income panels on this case to update. Run the script again and add a panel or review the case and verifications you are updating.")
                                '----------------------------------------------------------'
                        '---------------------------------------------------------------------------------'
'------------------------------------------------- GATHERING PAY INFORMATION FOR EACH PANEL --------------------------------------------------'
                        '---------------------------------------------------------------------------------'
                                '----------------------------------------------------------'

Call back_to_SELF               'this is a good reset
Do                              'getting back in to STAT in case it went to background adding a new panel.
    Call navigate_to_MAXIS_screen ("STAT", "SUMM")
    EMReadScreen summ_check, 4, 2, 46
Loop until summ_check = "SUMM"

pay_item = 0                    'counter for adding items to the LIST_OF_INCOME_ARRAY
check_SNAP_for_UH = FALSE       'this boolean tells the script if we need to look for Uncle Harry specifically - since this is outside of a specific panel
For ei_panel = 0 to UBOUND(EARNED_INCOME_PANELS_ARRAY, 2)       'looping through all of the current JOBS or BUSI panels

    cancel_clarify = ""         'blanking this from previous use AND from loops through the panels
    If EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel) = "JOBS" Then       'functionality for JOBS panels'

        Call Navigate_to_MAXIS_screen("STAT", "JOBS")                       'navigate to the current panel in the array
        EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel), 20, 76
        EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel), 20, 79
        transmit

        'This is where the worker indicates if they have income information to budget for this panel.
        'If they click 'No' here, there is no way to go back to this question for that panel.
        employer_check = MsgBox("Do you have income verification for this job? Employer name: " & EARNED_INCOME_PANELS_ARRAY(employer, ei_panel), vbYesNo + vbQuestion, "Select Income Panel")
        If employer_check = vbYes Then script_run_lowdown = script_run_lowdown & vbCr & "Income received for - " & EARNED_INCOME_PANELS_ARRAY(employer, ei_panel) & " button pressed - YES"
        If employer_check = vbNo Then script_run_lowdown = script_run_lowdown & vbCr & "Income received for - " & EARNED_INCOME_PANELS_ARRAY(employer, ei_panel) & " button pressed - NO"

        'Some panels will have this defaulted  already but if not, this will defalt to the footer month and year inidicated in the initial dialog
        If EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel) = "" Then EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel) = MAXIS_footer_month
        If EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel) = "" Then EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel) = MAXIS_footer_year
        EARNED_INCOME_PANELS_ARRAY(update_futue_chkbx, ei_panel) = future_months_check      'defaulted to the checkbox in the initial dialog

        'If the worker indicates there is income information for this panel.
        If employer_check = vbYes Then
            'There are a bunch of variables and information to default.
            EARNED_INCOME_PANELS_ARRAY(income_received, ei_panel) = TRUE        'This is set to indicate actions are needed by the rest of the script
            review_small_dlg = TRUE                                             'The ENTER PAY INFORMATION Dialog is only shown if this is 'TRUE'
            EARNED_INCOME_PANELS_ARRAY(ignore_antic, ei_panel) = FALSE          'This will be determined TURE as needed later if appropriate
            not_30_explanation = ""                                             'blanking out for each loop of panels
            est_weekly_hrs = ""                                                 'Blaning out variables used for math
            list_of_excluded_pay_dates = ""
            known_pay_date = ""
            'HOW the LIST_OF_INCOME_ARRAY(panel_indct, pay_item) position of the array works"
                'Basically there are 2 arrays, one of all the panels and one of all the income/paychecks
                'In order to connect these two arrays - this position in the LIST_OF_INCOME_ARRAY tells you which position in the EARNED_INCOME_PANELS_ARRAY it belongs to
                'So if there is JOBS 01 01 - in the EARNED_INCOME_PANELS_ARRAY it is at item '0' then the 3 checks listed that are for that income will have a '0' in this position of the LIST_OF_INCOME_ARRAY
                'Then when we loop through the income list - there is an IF that will only pull the items from LIST_OF_INCOME_ARRAY that align with the current panel being looked at.
                'see the code at line 1100 for the first example
            If LIST_OF_INCOME_ARRAY(panel_indct, pay_item) = "" Then            'If the item in LIST_OF_INCOME_ARRAY at the latest point has not been assigned to a panel, it is defaulted to the current panel
                LIST_OF_INCOME_ARRAY(panel_indct, pay_item) = ei_panel
            Else                                                                'Otherwise, we need to add a new item to LIST_OF_INCOME_ARRAY because the item of the LIST_OF_INCOME_ARRAY should be for the current panel
                pay_item = pay_item + 1
                ReDim Preserve LIST_OF_INCOME_ARRAY(reason_amt_excluded, pay_item)
                LIST_OF_INCOME_ARRAY(panel_indct, pay_item) = ei_panel
            End If

            'This code has 2 dialogs that are connected. There are many do-loops within each other
            Do          'LOOP UNTIL are_we_passworded_out = false - for the SECOND DIALOG - CONFIRM BUDGET Dialog
                Do          'Loop until big_err_msg = "" - for the SECOND DIALOG - CONFIRM BUDGET Dialog
                    big_err_msg = ""        'blanking this out for each loop
                    confirm_budget_checkbox = unchecked                                 'these are the checkboxes to continue past CONFIRM BUDGET Dialog
                    confirm_checks_checkbox = unchecked
                    hc_confirm_checks_checkbox = unchecked
                    GRH_confirm_budget_checkbox = unchecked
                    confirm_pay_freq_checkbox = unchecked

                    If EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel) = "?" Then EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel) = 0       'removing this '?''

                    Do          'Loop Until loop_to_add_missing_checks = FALSE - for FIRST DIALOG - ENTER PAY Dialog
                        loop_to_add_missing_checks = FALSE          'Setting this as a default at the beginning of each loop - the script will determine if there are missing checks later
                        If review_small_dlg = TRUE Then             'Sometimes we want to loop and see only CONFIRM BUDGET Dialog - this skips the ENTER PAY Dialog
                            Do          'LOOP UNTIL are_we_passworded_out = false - for FIRST DIALOG - ENTER PAY Dialog
                                Do          'Loop until sm_err_msg = "" - for FIRST DIALOG - ENTER PAY Dialog
                                    dlg_factor = 0      'this will set the dialog height - dynamicall determined
                                    Dialog1 = ""        'sometimes dialogs fail when they write over each other - this helps because ALL THE DIALOGS

                                    If UBound(LIST_OF_INCOME_ARRAY, 2) = 0 Then LIST_OF_INCOME_ARRAY(panel_indct, 0) = ei_panel     'If there is only 1 item in the LIST_OF_INCOME_ARRAY - it has to belong to the current panel

                                    If LIST_OF_INCOME_ARRAY(panel_indct, 0) <> "" Then          'This looks at each income item to see if it belongs to the current panel to determine how many spaces for pay is needed in the dialog
                                        For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                                            If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then dlg_factor = dlg_factor + 1
                                        Next
                                    End If

                                    dlg_factor = dlg_factor - 1     'There is always one extra
                                    sm_err_msg = ""                 'blanking this out at the beginning of the loop for displaying the dialog

                                    'ENTER PAY Dialog - dynamic dialog to enter job checks or anticipated amounts
                                    BeginDialog Dialog1, 0, 0, 606, (dlg_factor * 20) + 160, "Enter ALL Paychecks Received"
                                      Text 10, 10, 265, 10, "JOBS " & EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel) & " " & EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel) & " - " & EARNED_INCOME_PANELS_ARRAY(employer, ei_panel)
                                      Text 200, 15, 40, 10, "Start Date:"
                                      EditBox 235, 10, 50, 15, EARNED_INCOME_PANELS_ARRAY (income_start_dt, ei_panel)
                                      Text 295, 15, 50, 10, "Income Type:"
                                      DropListBox 345, 10, 100, 45, "J - WIOA"+chr(9)+"W - Wages"+chr(9)+"E - EITC"+chr(9)+"G - Experience Works"+chr(9)+"F - Federal Work Study"+chr(9)+"S - State Work Study"+chr(9)+"O - Other"+chr(9)+"C - Contract Income"+chr(9)+"T - Training Program"+chr(9)+"P - Service Program"+chr(9)+"R - Rehab Program", EARNED_INCOME_PANELS_ARRAY(income_type, ei_panel)
                                      GroupBox 455, 5, 145, 25, "Apply Income to Programs:"
                                      CheckBox 465, 15, 30, 10, "SNAP", EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel)
                                      CheckBox 500, 15, 30, 10, "CASH", EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, ei_panel)
                                      CheckBox 535, 15, 20, 10, "HC", EARNED_INCOME_PANELS_ARRAY(apply_to_HC, ei_panel)
                                      CheckBox 560, 15, 30, 10, "GRH", EARNED_INCOME_PANELS_ARRAY(apply_to_GRH, ei_panel)
                                      Text 5, 40, 60, 10, "JOBS Verif Code:"
                                      DropListBox 65, 35, 105, 45, "1 - Pay Stubs/Tip Report"+chr(9)+"2 - Empl Statement"+chr(9)+"3 - Coltrl Stmt"+chr(9)+"4 - Other Document"+chr(9)+"5 - Pend Out State Verification"+chr(9)+"N - No Ver Prvd"+chr(9)+"? - EXPEDITED SNAP ONLY", EARNED_INCOME_PANELS_ARRAY(income_verif, ei_panel)
                                      Text 175, 40, 155, 10, "additional detail of verification received:"
                                      EditBox 310, 35, 290, 15, EARNED_INCOME_PANELS_ARRAY(verif_explain, ei_panel)
                                      Text 5, 60, 90, 10, "Date verification received:"
                                      EditBox 100, 55, 50, 15, EARNED_INCOME_PANELS_ARRAY(verif_date, ei_panel)
                                      Text 460, 60, 50, 10, "Pay Frequency"
                                      DropListBox 515, 55, 85, 45, ""+chr(9)+"1 - One Time Per Month"+chr(9)+"2 - Two Times Per Month"+chr(9)+"3 - Every Other Week"+chr(9)+"4 - Every Week", EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel)
                                      Text 5, 85, 80, 10, "Pay Date (MM/DD/YY):"
                                      Text 90, 85, 50, 10, "Gross Amount:"
                                      Text 145, 85, 25, 10, "Hours:"
                                      Text 180, 70, 25, 25, "Use in SNAP budget"
                                      Text 235, 85, 85, 10, "If not used, explain why:"
                                      Text 355, 75, 245, 10, "If there is a specific amount that should be NOT budgeted from this check:"
                                      Text 355, 85, 30, 10, "Amount:"
                                      Text 410, 85, 30, 10, "Reason:"

                                      y_pos = 0     'this is how we move things down in dynamic dialogs
                                      For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                                          If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then
                                              LIST_OF_INCOME_ARRAY(exclude_amount, all_income) = LIST_OF_INCOME_ARRAY(exclude_amount, all_income) & ""
                                              If LIST_OF_INCOME_ARRAY(exclude_amount, all_income) = "0" Then LIST_OF_INCOME_ARRAY(exclude_amount, all_income) = ""  'If this was 0, then we are going to make it blank for loops and error handling
                                              LIST_OF_INCOME_ARRAY(pay_date, all_income) = LIST_OF_INCOME_ARRAY(pay_date, all_income) & ""
                                              EditBox 5, (y_pos * 20) + 95, 65, 15, LIST_OF_INCOME_ARRAY(pay_date, all_income)          'BUGGY CODE - this will get funky if the view_pay_date is different and we loop back here
                                              EditBox 90, (y_pos * 20) + 95, 45, 15, LIST_OF_INCOME_ARRAY(gross_amount, all_income)
                                              EditBox 145, (y_pos * 20) + 95, 25, 15, LIST_OF_INCOME_ARRAY(hours, all_income)

                                              CheckBox 180, (y_pos * 20) + 100, 50, 10, "Exclude", LIST_OF_INCOME_ARRAY(budget_in_SNAP_no, all_income)  'possibly excluding a whole paycheck

                                              EditBox 235, (y_pos * 20) + 95, 115, 15, LIST_OF_INCOME_ARRAY(reason_to_exclude, all_income)
                                              EditBox 355, (y_pos * 20) + 95, 45, 15, LIST_OF_INCOME_ARRAY(exclude_amount, all_income)
                                              EditBox 410, (y_pos * 20) + 95, 185, 15, LIST_OF_INCOME_ARRAY(reason_amt_excluded, all_income)
                                              y_pos = y_pos + 1
                                          End If
                                      Next

                                      Text 5, (dlg_factor * 20) + 145, 70, 10, "Anticipated Income:"
                                      Text 80, (dlg_factor * 20) + 145, 50, 10, "Rate of Pay/Hr"
                                      EditBox 135, (dlg_factor * 20) + 140, 50, 15, EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel)
                                      Text 190, (dlg_factor * 20) + 145, 35, 10, "Hours/Wk"
                                      EditBox 230, (dlg_factor * 20) + 140, 40, 15, est_weekly_hrs
                                      Text 275, (dlg_factor * 20) + 145, 65, 10, "Known Pay Date"
                                      EditBox 340, (dlg_factor * 20) + 140, 65, 15, known_pay_date

                                      Text 385, (dlg_factor * 20) + 120, 85, 10, "Initial Month to Update:"
                                      EditBox 465, (dlg_factor * 20) + 115, 15, 15, EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel)
                                      EditBox 485, (dlg_factor * 20) + 115, 15, 15, EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel)
                                      CheckBox 510, (dlg_factor * 20) + 120, 120, 10, "Update Future Months", EARNED_INCOME_PANELS_ARRAY(update_futue_chkbx, ei_panel)

                                      ButtonGroup ButtonPressed
                                        PushButton 5, (dlg_factor * 20) + 115, 15, 15, "+", add_another_check
                                        PushButton 25, (dlg_factor * 20) + 115, 15, 15, "-", take_a_check_away
                                        OkButton 550, (dlg_factor * 20) + 140, 50, 15
                                    EndDialog

                                    Dialog Dialog1
                                    cancel_confirmation     'there is no cancel button but this will make sure that if the 'X' is pressed the worker has a way out

                                    'Here we start error handling and it gets a little messy.
                                    EARNED_INCOME_PANELS_ARRAY(verif_explain, ei_panel) = trim(EARNED_INCOME_PANELS_ARRAY(verif_explain, ei_panel))
                                    If EARNED_INCOME_PANELS_ARRAY(income_verif, ei_panel) = "? - EXPEDITED SNAP ONLY" Then      'there are special criteria for using this verification code
                                        EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, ei_panel) = unchecked     'only for SNAP
                                        EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) = checked
                                        EARNED_INCOME_PANELS_ARRAY(apply_to_HC, ei_panel) = unchecked
                                        EARNED_INCOME_PANELS_ARRAY(apply_to_GRH, ei_panel) = unchecked

                                        EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel) = fs_appl_footer_month   'month of application handling only
                                        EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel) = fs_appl_footer_year

                                        If EARNED_INCOME_PANELS_ARRAY(verif_explain, ei_panel) = "" Then sm_err_msg = sm_err_msg & vbNewLine & "* If the verification code is '?' additional information about the verification needs to be added." 'need more explanation
                                    End If

                                    actual_checks_provided = FALSE      'defaults for some logic coming up
                                    there_are_counted_checks = FALSE
                                    all_pay_in_app_month = TRUE
                                    For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                                        LIST_OF_INCOME_ARRAY(pay_date, all_income) = trim(LIST_OF_INCOME_ARRAY(pay_date, all_income))           'formatting the information'
                                        LIST_OF_INCOME_ARRAY(gross_amount, all_income) = trim(LIST_OF_INCOME_ARRAY(gross_amount, all_income))
                                        LIST_OF_INCOME_ARRAY(hours, all_income) = trim(LIST_OF_INCOME_ARRAY(hours, all_income))
                                        If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(pay_date, all_income) <> "" AND LIST_OF_INCOME_ARRAY(gross_amount, all_income) <> "" AND LIST_OF_INCOME_ARRAY(hours, all_income) <> "" Then
                                            actual_checks_provided = TRUE           'this helps us know what functionality to use a little later
                                            If LIST_OF_INCOME_ARRAY(budget_in_SNAP_no, all_income) = unchecked Then there_are_counted_checks = TRUE
                                            LIST_OF_INCOME_ARRAY(future_check, all_income) = FALSE              'only in the month of application can we use future checks - we need to see if there are any
                                            If IsDate(LIST_OF_INCOME_ARRAY(pay_date, all_income)) = FALSE Then      'this needs to be a date
                                                sm_err_msg = sm_err_msg & vbNewLine & "* Enter a valid pay date for all checks."
                                            Else
                                                If DatePart("m", fs_appl_date) = DatePart("m", LIST_OF_INCOME_ARRAY(pay_date, all_income)) AND DatePart("yyyy", fs_appl_date) = DatePart("yyyy", LIST_OF_INCOME_ARRAY(pay_date, all_income)) Then   'if the paydate is in the application month
                                                    If DateDiff("d", date, LIST_OF_INCOME_ARRAY(pay_date, all_income)) > 0 Then LIST_OF_INCOME_ARRAY(future_check, all_income) = TRUE   'this is a future check
                                                Else        'if the paydate is NOT in the application  month
                                                    If DateDiff("d", date, LIST_OF_INCOME_ARRAY(pay_date, all_income)) > 0 Then             'if the pay date is in the future we have to error
                                                        LIST_OF_INCOME_ARRAY(pay_date, all_income) = "**" & LIST_OF_INCOME_ARRAY(pay_date, all_income)
                                                        sm_err_msg = sm_err_msg & vbNewLine & "* Paydates cannot be in the future."
                                                    End If
                                                End If
                                                If EARNED_INCOME_PANELS_ARRAY(income_verif, ei_panel) = "? - EXPEDITED SNAP ONLY" Then      'if the verifi is '?' Then the pay dates must ALL be in the application month
                                                    If DatePart("m", fs_appl_date) = DatePart("m", LIST_OF_INCOME_ARRAY(pay_date, all_income)) AND DatePart("yyyy", fs_appl_date) = DatePart("yyyy", LIST_OF_INCOME_ARRAY(pay_date, all_income)) Then       'this is a little messy - BUGGY CODE - works fine but maybe needs a logic upgrade to be more elegant
                                                    Else
                                                        all_pay_in_app_month = FALSE
                                                    End If

                                                End If
                                                If IsDate(EARNED_INCOME_PANELS_ARRAY(income_start_dt, ei_panel)) = TRUE Then
                                                    If DateDiff("d", LIST_OF_INCOME_ARRAY(pay_date, all_income), EARNED_INCOME_PANELS_ARRAY(income_start_dt, ei_panel)) > 0 Then sm_err_msg = sm_err_msg & vbNewLine & "* Pay date (" & LIST_OF_INCOME_ARRAY(pay_date, all_income) & ") is listed before the income start date."
                                                End If
                                            End If
                                            If IsNumeric(LIST_OF_INCOME_ARRAY(gross_amount, all_income)) = FALSE Then sm_err_msg = sm_err_msg & vbNewLine & "* Enter the Gross Amount of the check as a number."            'pay amount should be a number
                                            If LIST_OF_INCOME_ARRAY(budget_in_SNAP_no, all_income) = 1 AND trim(LIST_OF_INCOME_ARRAY(reason_to_exclude, all_income)) = "" Then sm_err_msg = sm_err_msg & vbNewLine & "* The check on " & LIST_OF_INCOME_ARRAY(pay_date, all_income) & " is to be excluded, list a reason for excluding this check."     'need to explain excluding a check
                                            If IsNumeric(LIST_OF_INCOME_ARRAY(hours, all_income)) = FALSE Then sm_err_msg = sm_err_msg & vbNewLine & "* Enter the number of hours for the paycheck on " & LIST_OF_INCOME_ARRAY(pay_date, all_income) & " as a number."      'hours are a number
                                            If IsNumeric(LIST_OF_INCOME_ARRAY(exclude_amount, all_income)) = FALSE AND trim(LIST_OF_INCOME_ARRAY(exclude_amount, all_income)) <> "" Then sm_err_msg = sm_err_msg & vbNewLine & "* Enter the amount excluded from the budget as a number."       'amounts are numbers
                                            If IsNumeric(LIST_OF_INCOME_ARRAY(exclude_amount, all_income)) = TRUE AND trim(LIST_OF_INCOME_ARRAY(reason_amt_excluded, all_income)) = "" Then sm_err_msg = sm_err_msg & vbNewLine & "* Explain why $" & LIST_OF_INCOME_ARRAY(exclude_amount, all_income) & " is excluded from the pay on " & LIST_OF_INCOME_ARRAY(pay_date, all_income) & "." 'excluded portion needs explanation
                                            LIST_OF_INCOME_ARRAY(exclude_amount, all_income) = trim(LIST_OF_INCOME_ARRAY(exclude_amount, all_income))
                                        End If
                                    Next
                                    'FUTURE FUNCTIONALITY - read the Employer name for AmeriCorps to provide detail based upon the correct coding.
                                    'FUTURE FUNCTIONALITY - from the inomce type and subsidy code determine some notes on income (excluded etc.) for noting.
                                    If all_pay_in_app_month = FALSE Then sm_err_msg = sm_err_msg & vbNewLine & "* Only income from the month of appliation should be entered when using '?' as this is only for income that is not sufficiently verified to be used to determine Expedited."    'this only happens if '?' is the verif code
                                    anticipated_income_provided = FALSE     'default
                                    EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel) = trim(EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel))       'formatting
                                    EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) = trim(est_weekly_hrs)
                                    EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = trim(EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel))

                                    EARNED_INCOME_PANELS_ARRAY(reg_non_monthly, ei_panel) = trim(EARNED_INCOME_PANELS_ARRAY(reg_non_monthly, ei_panel))     'this is not currently in the dialog - FUTURE FUNCTIONALITY - need a lot of other handling to put this back in.
                                    EARNED_INCOME_PANELS_ARRAY(numb_months, ei_panel) = trim(EARNED_INCOME_PANELS_ARRAY(numb_months, ei_panel))
                                    known_pay_date = trim(known_pay_date)

                                    If EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel) <> "" AND EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) <> "" AND EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) <> "" Then          'If all anticiapated pay information has been provided, we look for a start date and define that anticiapated income is provided
                                        anticipated_income_provided = TRUE
                                        If EARNED_INCOME_PANELS_ARRAY(income_start_dt, ei_panel) = "" Then sm_err_msg = sm_err_msg & vbNewLine & "* Enter an income start date, since anticipated pay dates cannot be determined without the initial pay date."
                                    End If
                                    If EARNED_INCOME_PANELS_ARRAY(reg_non_monthly, ei_panel) <> "" AND EARNED_INCOME_PANELS_ARRAY(numb_months, ei_panel) <> "" Then anticipated_income_provided = TRUE
                                    If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "" Then sm_err_msg = sm_err_msg & vbNewLine & "* Select the pay frequency for this job."        'NEED to have a pay frequency

                                    If anticipated_income_provided = FALSE AND actual_checks_provided = FALSE Then          'there either needs to be checks OR anticipated income
                                        sm_err_msg = sm_err_msg & vbNewLine & "* Income information needs to be provided, either in the form of actual checks or anticipated income, hours, and rate of pay."
                                    End If
                                    If there_are_counted_checks = FALSE AND anticipated_income_provided = FALSE AND actual_checks_provided = TRUE Then sm_err_msg = sm_err_msg & vbNewLine & "* All the checks listed are excluded and no anticipated income estimate is provided. In order to udate a case and budget income there needs to be counted income."
                                    If known_pay_date <> "" AND IsDate(known_pay_date) = FALSE Then sm_err_msg = sm_err_msg & vbNewLine & "* A known pay date needs to be entered as a date. Check the entry."

                                    'the income needs to apply to at least one program
                                    If EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) = unchecked AND EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, ei_panel) = unchecked AND EARNED_INCOME_PANELS_ARRAY(apply_to_HC, ei_panel) = unchecked AND EARNED_INCOME_PANELS_ARRAY(apply_to_GRH, ei_panel) = unchecked Then sm_err_msg = sm_err_msg & vbNewLine & "* No programs have been selected that this icnome applies to. Chose at least one program that this income is budgeted for."
                                    EARNED_INCOME_PANELS_ARRAY(verif_date, ei_panel) = trim(EARNED_INCOME_PANELS_ARRAY(verif_date, ei_panel))
                                    If EARNED_INCOME_PANELS_ARRAY(verif_date, ei_panel) = "" Then sm_err_msg = sm_err_msg & "* Enter the date the pay information was received in the agency."

                                    If ButtonPressed = add_another_check Then       'functionality to add another check to the dialog using the '+' button
                                        pay_item = pay_item + 1     'incrementing the counter
                                        ReDim Preserve LIST_OF_INCOME_ARRAY(reason_amt_excluded, pay_item)      'resizing the array
                                        LIST_OF_INCOME_ARRAY(panel_indct, pay_item) = ei_panel          'setting the new LIST_OF_INCOME_ARRAY item to the current panel
                                        dlg_factor = dlg_factor + 1     'making the dialog bigger

                                        LIST_OF_INCOME_ARRAY(pay_date, pay_item) = ""               'there were weird amounts being put in so we need these to start as blank
                                        LIST_OF_INCOME_ARRAY(gross_amount, pay_item) = ""
                                        LIST_OF_INCOME_ARRAY(hours, pay_item) = ""
                                        LIST_OF_INCOME_ARRAY(reason_to_exclude, pay_item) = ""
                                        LIST_OF_INCOME_ARRAY(exclude_amount, pay_item) = ""
                                        LIST_OF_INCOME_ARRAY(reason_amt_excluded, pay_item) = ""

                                        sm_err_msg = "LOOP" & sm_err_msg            'makes the dialog loop back without displaying an error message

                                    End If

                                    If ButtonPressed = take_a_check_away Then       'functionality to take a check away from the dialog using the '-' button
                                        pay_item = pay_item - 1                 'incremnting the counter backward   possible BUGGY CODE - may be an issue with the 2nd job updates - we could potentially erase the items for another panel.
                                        If pay_item < 0 Then pay_item = 0       'making sure we don't go below 0
                                        ReDim Preserve LIST_OF_INCOME_ARRAY(reason_amt_excluded, pay_item)  'resizing the array - YES IT WORKS BOTH WAYS
                                        dlg_factor = dlg_factor - 1             'making the dialog smaller
                                        sm_err_msg = "LOOP" & sm_err_msg        'makes the dialog loop back without displaying an error message
                                    End If

                                    If sm_err_msg <> "" AND left(sm_err_msg, 4) <> "LOOP" then MsgBox "Please resolve before continuing:" & vbNewLine & sm_err_msg      'shoing the error message if there is one

                                Loop until sm_err_msg = ""
                                call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
                            LOOP UNTIL are_we_passworded_out = false

                            'I guess we do this again - maybe we can REMOVE CODE
                            actual_checks_provided = FALSE
                            For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                                If LIST_OF_INCOME_ARRAY(pay_date, all_income) = "" AND LIST_OF_INCOME_ARRAY(gross_amount, all_income) = "" AND LIST_OF_INCOME_ARRAY(hours, all_income) = "" Then LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ""
                                If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then
                                    'ADD ERROR HANDLING HERE
                                    actual_checks_provided = TRUE
                                End If
                            Next
                        End If

                        'If BOTH anticipated income AND actual checks are provided, the worker needs to chose which one should be budgeted.
                        If actual_checks_provided = TRUE AND anticipated_income_provided = TRUE Then
                            'CHOOSE CORRECT METHOD Dialog - select which (actual or anticipated) income information to budget and explain
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

                                    Dialog Dialog1      'one of the easiest dialogs in this script

                                    EARNED_INCOME_PANELS_ARRAY(selection_rsn, ei_panel) = trim(EARNED_INCOME_PANELS_ARRAY(selection_rsn, ei_panel))
                                    If use_actual_income = checked Then selection_pick = "ACTUAL LIST OF CHECKS."
                                    If use_anticipated_income = checked Then selection_pick = "INCOME ESTIMATED FROM HOURS AND RATE OF PAY."

                                    If EARNED_INCOME_PANELS_ARRAY(selection_rsn, ei_panel) = "" Then sm_err_msg = sm_err_msg & vbNewLine & "* Enter explanation of why the best way to determine future income is to use " & selection_pick
                                    If len(EARNED_INCOME_PANELS_ARRAY(selection_rsn, ei_panel)) < 10 Then sm_err_msg = sm_err_msg & vbNewLine & "* Explanation is not sufficient to adequately case note information about budget. Expand."

                                    If sm_err_msg <> "" Then MsgBox "** Please Resolve before Continuting **" & vbNewLine & sm_err_msg
                                Loop until sm_err_msg = ""
                                call check_for_password(are_we_passworded_out)
                            Loop until are_we_passworded_out = false

                            'Setting selections based on the choice made
                            If use_actual_income = checked Then
                                EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_actual
                                EARNED_INCOME_PANELS_ARRAY(ignore_antic, ei_panel) = TRUE
                            End If
                            If use_anticipated_income = checked Then
                                EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_estimate

                                For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)               'adding these to a list of excluded checks for noting purposes
                                    If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then
                                        LIST_OF_INCOME_ARRAY(budget_in_SNAP_no, all_income) = checked
                                        LIST_OF_INCOME_ARRAY(reason_to_exclude, all_income) = "Not best estimate of Anticipated Income"
                                        list_of_actual_paydates = list_of_actual_paydates & "~" & LIST_OF_INCOME_ARRAY(pay_date, all_income)
                                        list_of_excluded_pay_dates = list_of_excluded_pay_dates & ", " & LIST_OF_INCOME_ARRAY(pay_date, all_income)
                                    End If
                                Next
                                If left(list_of_actual_paydates, 1) = "~" Then list_of_actual_paydates = right(list_of_actual_paydates, len(list_of_actual_paydates) - 1)
                                If list_of_excluded_pay_dates <> "" Then list_of_excluded_pay_dates = right(list_of_excluded_pay_dates, len(list_of_excluded_pay_dates) - 2)
                            End If

                            'https://www.dhssir.cty.dhs.state.mn.us/MAXIS/trntl/snap/SNAP_Anticipating_Income.pdf - this is all about why we have to pick
                        Else        'if it isn't both then we don't need the dialog and the script sets the selection based on entry information
                            If actual_checks_provided = TRUE Then EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_actual
                            If anticipated_income_provided = TRUE Then EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_estimate
                        End If

                        If actual_checks_provided = TRUE Then       'this does not matter which is chosen - it does this if actual checks are provided even if that is not the correct budget
                            total_of_counted_income = 0             'there will be lots of counting/adding here so we need everything to start at 0 so we don't get carryover from previous loops or panels
                            total_of_gross_income = 0
                            total_of_included_pay_checks = 0
                            total_of_hours = 0
                            number_of_checks_budgeted = 0
                            cash_checks = 0
                            EARNED_INCOME_PANELS_ARRAY(pay_weekday, ei_panel) = ""
                            list_of_excluded_pay_dates = ""
                            previous_pay_date = ""
                            paycheck_list_title = "Paychecks Provided for Determination:"

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
                            first_date = array_of_pay_dates(0)                              'setting the first and last check dates
                            last_date = array_of_pay_dates(UBOUND(array_of_pay_dates))

                            list_of_days_of_checks = "~"
                            the_day_of_month = ""
                            third_paydate = FALSE
                            If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "2 - Two Times Per Month" AND (EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = "" OR EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) = "") Then
                                For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                    If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then
                                        the_day_of_month = DatePart("d", LIST_OF_INCOME_ARRAY(pay_date, all_income))
                                        the_day_of_month = "~" & the_day_of_month & "~"
                                        If InStr(list_of_days_of_checks, the_day_of_month) = 0 Then
                                            the_day_of_month = replace(the_day_of_month, "~", "")
                                            list_of_days_of_checks = list_of_days_of_checks & the_day_of_month & "~"
                                        End If
                                    End If
                                Next

                                For each_day = 1 to 31
                                    each_day_spider = "~" & each_day & "~"
                                    If InStr(list_of_days_of_checks, each_day_spider) <> 0 Then
                                        If EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) = "" Then
                                            EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) = each_day
                                        ElseIf EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = "" Then
                                            If each_day = 28 OR each_day = 29 OR each_day = 30 OR each_day = 31 Then
                                                EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = "LAST"
                                            Else
                                                EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = each_day
                                            End If
                                        Else
                                            third_paydate = TRUE
                                        End If
                                    End If
                                Next
                                If EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) = 28 OR EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) = 29 OR EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) = 30 OR EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) = 31 Then
                                    EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) = ""
                                    EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = "LAST"
                                End If

                                If third_paydate = TRUE OR EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = "" OR EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) = "" Then
                                    EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) = EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) & ""
                                    EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) & ""
                                    If EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = "LAST" Then
                                        last_day_checkbox = checked
                                        EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = ""
                                    End If
                                    BeginDialog Dialog1, 0, 0, 106, 115, "Days of Pay for Bimonthly"
                                      EditBox 55, 35, 25, 15, EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel)
                                      EditBox 55, 55, 25, 15, EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel)
                                      ButtonGroup ButtonPressed
                                        OkButton 50, 95, 50, 15
                                      Text 10, 10, 95, 20, "Dates of Pay for BiMonthly Pay Frequency"
                                      Text 10, 40, 35, 10, "First Day"
                                      Text 10, 60, 45, 10, "Second Day"
                                      CheckBox 10, 80, 95, 10, "Second Day is LAST Day", last_day_checkbox
                                    EndDialog

                                    Do
                                        the_err = ""

                                        dialog Dialog1

                                        EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) = trim(EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel))
                                        EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = trim(EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel))
                                        If EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) = "" Then the_err = the_err & vbNewLine & "* Enter the DAY of the month the first paycheck comes on."
                                        If EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = "" Then
                                            If last_day_checkbox = unchecked Then the_err = the_err & vbNewLine & "* Enter the DAY of the month the second paycheck comes on. Or check the box indicating the second check falls on the last day of the month."
                                        End If
                                        If the_err <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & the_err
                                    Loop until the_err = ""

                                    EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) = EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) * 1
                                    If IsNumeric(EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel)) = TRUE Then EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) * 1
                                    If last_day_checkbox = checked Then
                                        EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = "LAST"
                                    ElseIf EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = 28 OR EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = 29 OR EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = 30 OR EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = 31 Then
                                        EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = "LAST"
                                    End If
                                End If

                            End If

                            list_of_all_paydates_start_to_finish = ""   'Here we loop through to create a list of all the paychcks that we should see from the first listed to the last
                            next_paydate = first_date
                            dates_index = 0
                            Do
                                list_of_all_paydates_start_to_finish = list_of_all_paydates_start_to_finish & "~" & next_paydate

                                If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "1 - One Time Per Month" Then       'each next date is determined by the pay frequency
                                    next_paydate = DateAdd("m", 1, next_paydate)
                                ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "2 - Two Times Per Month" Then
                                    If DatePart("d", next_paydate) = EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) Then         'If we are at the first check of the month, we need to go to the second
                                        next_pay_month = DatePart("m", next_paydate)
                                        next_pay_year = DatePart("yyyy", next_paydate)

                                        If EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = "LAST" Then
                                            month_after = next_pay_month & "/1/" & next_pay_year
                                            month_after = DateAdd("m", 1, month_after)
                                            next_paydate = DateAdd("d", -1, month_after)
                                        Else
                                            next_paydate = next_pay_month & "/" & EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) & "/" & next_pay_year
                                        End If
                                    Else
                                        next_pay = DateAdd("m", 1, next_paydate)                                                            'go to the next month
                                        next_pay_month = DatePart("m", next_pay)
                                        next_pay_year = DatePart("yyyy", next_pay)
                                        next_paydate = next_pay_month & "/" & EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) & "/" & next_pay_year   'then go to the second pay date
                                    End If
                                ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "3 - Every Other Week" Then
                                    next_paydate = DateAdd("d", 14, next_paydate)
                                ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "4 - Every Week" Then
                                    next_paydate = DateAdd("d", 7, next_paydate)
                                End If
                            Loop until DateDiff("d", last_date, next_paydate) > 0           'We go until the loop has moved past the last pay date entered
                            If left(list_of_all_paydates_start_to_finish, 1) = "~" Then     'now we make the list an array
                                list_of_all_paydates_start_to_finish = right(list_of_all_paydates_start_to_finish, len(list_of_all_paydates_start_to_finish) - 1)
                            End If

                            expected_check_array = split(list_of_all_paydates_start_to_finish, "~")
                            If expected_check_array(UBound(expected_check_array)) <> last_date Then     'this got a little weird sometimes so it is just a double check
                                expected_check_array = ""
                                list_of_all_paydates_start_to_finish = list_of_all_paydates_start_to_finish & "~" & next_paydate
                                expected_check_array = split(list_of_all_paydates_start_to_finish, "~")
                            End If

                            EARNED_INCOME_PANELS_ARRAY(last_paycheck, ei_panel) = last_date     'saving this for the panel information
                            spread_of_pay_dates = DateDiff("d", first_date, last_date)          'this is how many days are between the 1st and last check - because 30 days of verif is still a thing
                            If EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) = checked THen   'though it is only a thing for SNAP so we are going to check the spread IF SNAP is a concern
                                using_30_days = TRUE        'this defaults to true

                                If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "1 - One Time Per Month" Then       'each pay frequency has a specific number of days that indicate 30 days of income has been reached
                                    If spread_of_pay_dates > 30 Then using_30_days = FALSE                              'spoiler alert - it is not always 30 days ... because of course not
                                ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "2 - Two Times Per Month" Then
                                    If spread_of_pay_dates > 30 Then using_30_days = FALSE
                                    If spread_of_pay_dates < 13 Then using_30_days = FALSE
                                ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "3 - Every Other Week" Then
                                    If spread_of_pay_dates <> 28 Then using_30_days = FALSE
                                ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "4 - Every Week" Then
                                    If spread_of_pay_dates <> 28 Then using_30_days = FALSE
                                ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "5 - Other" Then
                                End If
                            End If

                            'here we have a fun array loop within an array loop
                            'The array_of_pay_dates is the array with all the pay dates in chronological order
                            'Our goal here is to assign each item in LIST_OF_INCOME_ARRAY for this panel with basically their place in line chronologically
                            'The check_order position of LIST_OF_INCOME_ARRAY saves that place in line so we can call them in order later.
                            For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)           'Now loop through all of the listed income - again
                                If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then    'find the ones for THIS PANEL ONLY
                                    for index = 0 to UBOUND(array_of_pay_dates)                     'loop through the array of the pay dates only'
                                        'once the pay date in the income array matches the one in the chronological list of dates, use the index number to set an order code within the list of income array
                                        If array_of_pay_dates(index) = LIST_OF_INCOME_ARRAY(pay_date, all_income) Then
                                            LIST_OF_INCOME_ARRAY(pay_date, all_income) = DateValue(LIST_OF_INCOME_ARRAY(pay_date, all_income))
                                            LIST_OF_INCOME_ARRAY(check_order, all_income) = index + 1

                                        End If
                                        top_of_order = index + 1    'this identifies how many pay dates there are in for this panel
                                    next
                                End If
                            Next
                            EARNED_INCOME_PANELS_ARRAY(order_ubound, ei_panel) = top_of_order   'setting the number of unique pay dates within the panel array because we need it for sorting correctly

                            expected_check_index = 0        'setting up for another loop to see if all the expected checks have in fact been provided.
                            missing_checks_list = ""
                            order_number = 1
                            Do
                                For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                    date_in_range = ""
                                    'conditional if it is the right panel AND the order matches - then do the thing you need to do
                                    If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(check_order, all_income) = order_number Then
                                        missing_check = FALSE       'defaulting this for each loop

                                        'here we are comparing each check from the ENTER PAY Dialog for this panel in order to the checks we expected to see
                                        'We can only get an accurate panel update if all the checks for the time frame provided are given - they can be excluded but they should be there
                                        'There are allowances here for some variation as sometimes paydates shift (ie holidays or extenuating circumstances)
                                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "1 - One Time Per Month" Then
                                            date_in_range = DateDiff("d", LIST_OF_INCOME_ARRAY(pay_date, all_income), expected_check_array(expected_check_index))
                                            date_in_range = Abs(date_in_range)
                                            If date_in_range > 8 Then missing_check = TRUE      '8 day allowance
                                        ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "2 - Two Times Per Month" Then
                                            date_in_range = DateDiff("d", LIST_OF_INCOME_ARRAY(pay_date, all_income), expected_check_array(expected_check_index))
                                            date_in_range = Abs(date_in_range)
                                            If date_in_range > 5 Then missing_check = TRUE      '5 day allowance
                                        ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "3 - Every Other Week" Then
                                            date_in_range = DateDiff("d", LIST_OF_INCOME_ARRAY(pay_date, all_income), expected_check_array(expected_check_index))
                                            date_in_range = Abs(date_in_range)
                                            If date_in_range > 3 Then missing_check = TRUE      '3 day allowance
                                        ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "4 - Every Week" Then
                                            date_in_range = DateDiff("d", LIST_OF_INCOME_ARRAY(pay_date, all_income), expected_check_array(expected_check_index))
                                            date_in_range = Abs(date_in_range)
                                            If date_in_range > 3 Then missing_check = TRUE      '3 day allowance
                                        End If

                                        If missing_check = TRUE Then        'if the date difference was too much then we save the date to a list
                                            missing_checks_list = missing_checks_list & "~" & expected_check_array(expected_check_index)
                                        Else
                                            order_number = order_number + 1
                                        End If
                                        expected_check_index = expected_check_index + 1
                                        If order_number > top_of_order Then Exit For            'if we have reached the end of the entered checks OR the end of the expected checks, we need to leave the loop
                                        If expected_check_index > UBound(expected_check_array) Then Exit For
                                    End If
                                Next
                                If expected_check_index > UBound(expected_check_array) Then Exit Do     'if we have reached the end of the entered checks OR the end of the expected checks, we need to leave the loop
                                If order_number > top_of_order Then Exit Do
                            Loop until order_number = top_of_order

                            If missing_checks_list <> "" Then       'if there were any missing checks found
                                if left(missing_checks_list, 1) = "~" Then missing_checks_list = right(missing_checks_list, len(missing_checks_list) - 1)       'create an array of the missing checks
                                missing_checks_list = split(missing_checks_list, "~")
                                loop_to_add_missing_checks = TRUE       'these are set to show the ENTER PAY Dialog again without going to he CONFIRM BUDGET Dialog
                                review_small_dlg = TRUE

                                For each check_missed in missing_checks_list        'these missing dates get added to the LIST_OF_INCOME_ARRAY automatically
                                    pay_item = pay_item + 1
                                    ReDim Preserve LIST_OF_INCOME_ARRAY(reason_amt_excluded, pay_item)
                                    LIST_OF_INCOME_ARRAY(panel_indct, pay_item) = ei_panel
                                    dlg_factor = dlg_factor + 1

                                    LIST_OF_INCOME_ARRAY(pay_date, pay_item) = check_missed
                                    LIST_OF_INCOME_ARRAY(gross_amount, pay_item) = ""
                                    LIST_OF_INCOME_ARRAY(hours, pay_item) = ""
                                    LIST_OF_INCOME_ARRAY(reason_to_exclude, pay_item) = ""
                                    LIST_OF_INCOME_ARRAY(exclude_amount, pay_item) = ""
                                    LIST_OF_INCOME_ARRAY(reason_amt_excluded, pay_item) = ""
                                Next
                                'telling the worker why we are going back
                                MsgBox "*** It appears there are checks missing ***" & vbNewLine & vbNewLine & "All checks need to be entered to have a correct budget. If there are pay dates between the first and last date entered that were not included, include them now. If the pay was $0, list $0 income."
                            End If
                        End If      'If actual_checks_provided = TRUE Then
                    Loop Until loop_to_add_missing_checks = FALSE

                    prev_date = ""              'setting some variables for the loop
                    days_between_checks = ""
                    'here we are going to see if there are checks out of line with the expected frequency.
                    'These may be the correct paydates but later in the script we use the precise interval based on pay frequency to enter information
                    If actual_checks_provided = TRUE Then           'again, does not mater which way to budget is selected
                        issues_with_frequency = FALSE               'default to false
                        For order_number = 1 to top_of_order                        'loop through the order number lowest to highest
                            For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                'conditional if it is the right panel AND the order matches - then do the thing you need to do
                                If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(check_order, all_income) = order_number Then
                                    If EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel) = "" Then EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel) = LIST_OF_INCOME_ARRAY(pay_date, all_income)       'setting the first check to the panel if it has not been done
                                    list_of_dates = list_of_dates & vbNewLine & "Check Date: " & LIST_OF_INCOME_ARRAY(pay_date, all_income) & " Income: $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " Hours: " & LIST_OF_INCOME_ARRAY(hours, all_income)      'creating a readable list of the pay dates, amount, and hours
                                    LIST_OF_INCOME_ARRAY(view_pay_date, all_income) = LIST_OF_INCOME_ARRAY(pay_date, all_income)        'view pay date is the actual date that is always seen and typically is the same as the regular pay date
                                    LIST_OF_INCOME_ARRAY(frequency_issue, all_income) = FALSE                                           'defaulting this to false

                                    If prev_date <> "" Then     'we can't compare the first date to anything, so it skips the first date
                                        days_between_checks = DateDiff("d", prev_date, LIST_OF_INCOME_ARRAY(pay_date, all_income))      'determines how many days from one check to the next

                                        'if the number of days is more or less than exactly what we expect, we need clarification
                                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "1 - One Time Per Month" Then
                                            If days_between_checks < 28 or days_between_checks > 31 Then
                                                issues_with_frequency = TRUE
                                                LIST_OF_INCOME_ARRAY(frequency_issue, all_income) = TRUE
                                                LIST_OF_INCOME_ARRAY(pay_date, all_income) = DateAdd("m", 1, prev_date)
                                            End If
                                        ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "2 - Two Times Per Month" Then
                                            If EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = "LAST" Then
                                                If DatePart("d", LIST_OF_INCOME_ARRAY(pay_date, all_income)) <> EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) Then
                                                    day_after_pay = DateAdd("d", 1, LIST_OF_INCOME_ARRAY(pay_date, all_income))
                                                    If DatePart("d", day_after_pay) <> 1 Then
                                                        issues_with_frequency = TRUE
                                                        LIST_OF_INCOME_ARRAY(frequency_issue, all_income) = TRUE
                                                        month_to_use = DatePart("m", LIST_OF_INCOME_ARRAY(pay_date, all_income))
                                                        year_to_use = DatePart("yyyy", LIST_OF_INCOME_ARRAY(pay_date, all_income))
                                                        first_of_payMonth = month_to_use & "/1/" & year_to_use
                                                        first_of_nextMonth = DateAdd("m", 1, first_of_payMonth)
                                                        LIST_OF_INCOME_ARRAY(pay_date, all_income) = DateAdd("d", -1, first_of_nextMonth)
                                                    End If
                                                End If

                                            Else
                                                If DatePart("d", LIST_OF_INCOME_ARRAY(pay_date, all_income)) <> EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) AND DatePart("d", LIST_OF_INCOME_ARRAY(pay_date, all_income)) <> EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) Then
                                                    issues_with_frequency = TRUE
                                                    LIST_OF_INCOME_ARRAY(frequency_issue, all_income) = TRUE
                                                    month_to_use = DatePart("m", LIST_OF_INCOME_ARRAY(pay_date, all_income))
                                                    year_to_use = DatePart("yyyy", LIST_OF_INCOME_ARRAY(pay_date, all_income))
                                                    If DatePart("d", prev_date) = EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) Then LIST_OF_INCOME_ARRAY(pay_date, all_income) = month_to_use & "/" & EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) & "/" & year_to_use
                                                    If DatePart("d", prev_date) = EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) Then LIST_OF_INCOME_ARRAY(pay_date, all_income) = month_to_use & "/" & EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) & "/" & year_to_use
                                                End If
                                            End If
                                        ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "3 - Every Other Week" Then
                                            If days_between_checks <> 14 Then
                                                issues_with_frequency = TRUE
                                                LIST_OF_INCOME_ARRAY(frequency_issue, all_income) = TRUE
                                                LIST_OF_INCOME_ARRAY(pay_date, all_income) = DateAdd("d", 14, prev_date)
                                            End If
                                        ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "4 - Every Week" Then
                                            If days_between_checks <> 7 Then
                                                issues_with_frequency = TRUE
                                                LIST_OF_INCOME_ARRAY(frequency_issue, all_income) = TRUE
                                                LIST_OF_INCOME_ARRAY(pay_date, all_income) = DateAdd("d", 7, prev_date)
                                            End If
                                        ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "5 - Other" Then

                                        'REMOVE CODE
                                        Else        'this is code to determine the pay frequency for the worker but with all the other functionality - this is something the worker needs to provide
                                            If days_between_checks = 7 Then
                                                EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "4 - Every Week"
                                            ElseIf days_between_checks = 14 Then
                                                EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "3 - Every Other Week"
                                            ElseIf days_between_checks >= 14 AND days_between_checks <= 19 Then
                                                EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "2 - Two Times Per Month"
                                            ElseIf days_between_checks >= 28 AND days_between_checks <= 31 Then
                                                EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "1 - One Time Per Month"
                                            End If

                                        End If          'If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) =
                                    End If          'If prev_date <> "" Then
                                    prev_date = LIST_OF_INCOME_ARRAY(pay_date, all_income)      'saving this date as the one to compare to in the next loop
                                End If          'If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(check_order, all_income) = order_number Then
                            next            'For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                        next            'For order_number = 1 to top_of_order

                        If issues_with_frequency = TRUE Then        'if any checks did not align
                            dlg_len = 85        'setting the base height

                            For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)       'increasing the height for each date with a frequency issue
                                If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(frequency_issue, all_income) = TRUE Then dlg_len = dlg_len + 20
                            Next

                            'FREQUENCY ISSUE Dialog - the worker can update the view_pay_date to match if appropriate or they can confirm it is correct as is
                            BeginDialog Dialog1, 0, 0, 251, dlg_len, "Review Pay Dates"
                              Text 10, 10, 240, 10, "It appears one check does not fall in the expected pay schedule dates. "
                              Text 10, 25, 230, 10, "This job is paid - " & EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel)
                              Text 10, 40, 65, 10, "Reported Pay Date"
                              Text 85, 40, 75, 10, "Expected Pay Date"

                              y_pos = 55
                              For order_number = 1 to top_of_order                        'loop through the order number lowest to highest
                                  For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                      'conditional if it is the right panel AND the order matches - then do the thing you need to do
                                      If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(check_order, all_income) = order_number Then
                                          If LIST_OF_INCOME_ARRAY(frequency_issue, all_income) = TRUE Then
                                              If LIST_OF_INCOME_ARRAY(view_pay_date, all_income) = "" Then LIST_OF_INCOME_ARRAY(view_pay_date, all_income) = LIST_OF_INCOME_ARRAY(view_pay_date, all_income) & ""
                                              Text 10, y_pos, 10, 10, "**"
                                              EditBox 25, y_pos, 50, 15, LIST_OF_INCOME_ARRAY(view_pay_date, all_income)
                                              Text 95, y_pos + 5, 50, 10, LIST_OF_INCOME_ARRAY(pay_date, all_income)            'this cannot be changed here

                                              y_pos = y_pos + 20
                                          End If
                                      End If
                                  Next
                              Next
                              CheckBox 10, y_pos, 180, 10, "Check here if these pay dates are what was reported.", pay_dates_correct_checkbox
                              ButtonGroup ButtonPressed
                                OkButton 195, y_pos, 50, 15
                            EndDialog

                            Do
                                Do
                                    dialog Dialog1      'showing the dialog

                                    'worker must confirm the dates are correct
                                    If pay_dates_correct_checkbox = unchecked Then MsgBox "If the paydates reported in the paycheck information dates are incorrect. Correct them here. "
                                Loop until pay_dates_correct_checkbox = checked
                                call check_for_password(are_we_passworded_out)
                            Loop until are_we_passworded_out = false

                            For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(check_order, all_income) = order_number Then
                                    If LIST_OF_INCOME_ARRAY(frequency_issue, all_income) = TRUE Then
                                        'if there is more than a 6 day difference between the provided date and the expected date, we wil match the provided date
                                        If abs(DateDiff("d", LIST_OF_INCOME_ARRAY(view_pay_date, all_income), LIST_OF_INCOME_ARRAY(pay_date, all_income))) > 6 Then LIST_OF_INCOME_ARRAY(pay_date, all_income) = LIST_OF_INCOME_ARRAY(view_pay_date, all_income)
                                    End If
                                End If
                            Next

                        End If          'If issues_with_frequency = TRUE Then

                        'Finding the pay day of the week for biweekly and weekly pay schedules
                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "3 - Every Other Week" OR EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "4 - Every Week" Then
                            For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                                If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then
                                    EARNED_INCOME_PANELS_ARRAY(pay_weekday, ei_panel) = WeekDayName(Weekday(LIST_OF_INCOME_ARRAY(pay_date, all_income)))
                                    'MsgBox "Look at the payday: " & EARNED_INCOME_PANELS_ARRAY(pay_weekday, ei_panel)
                                    Exit For
                                End If
                            Next
                        End If
                    End If          'If actual_checks_provided = TRUE Then

                    cash_checks = 0         'setting for counting
                    number_of_checks_budgeted = 0
                    all_total_hours = 0
                    EARNED_INCOME_PANELS_ARRAY(income_list_indct, ei_panel) = "NONE"        'resetting this for the loop if it happens otherwise this list will be very large
                    If EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_actual Then     'so here if we are actually using the actual checks for determining the SNAP budget we need to do some math

                        For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)           'looking at all the income (order does not mater here)
                            If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then        'if the income is for this panel

                                EARNED_INCOME_PANELS_ARRAY(income_list_indct, ei_panel) = EARNED_INCOME_PANELS_ARRAY(income_list_indct, ei_panel) & "~" & all_income        'this adds a list of all the items from LIST_OF_INCOME_ARRAY the belong to this panel in EARNED_INCOME_PANELS_ARRAY
                                'currently this list is not being utilized but may be needed as future functionality is added

                                If LIST_OF_INCOME_ARRAY(exclude_amount, all_income) = "" Then LIST_OF_INCOME_ARRAY(exclude_amount, all_income) = 0      'making this a number
                                LIST_OF_INCOME_ARRAY(exclude_amount, all_income) = LIST_OF_INCOME_ARRAY(exclude_amount, all_income) * 1

                                If LIST_OF_INCOME_ARRAY(future_check, all_income) = FALSE Then          'future checks are not counted for determining averages/estimates
                                    total_of_gross_income = total_of_gross_income + LIST_OF_INCOME_ARRAY(gross_amount, all_income)      'this is for non-snap programs
                                    all_total_hours = all_total_hours + LIST_OF_INCOME_ARRAY(hours, all_income)
                                    cash_checks = cash_checks + 1
                                End If

                                If LIST_OF_INCOME_ARRAY(budget_in_SNAP_no, all_income) = unchecked Then LIST_OF_INCOME_ARRAY(budget_in_SNAP_yes, all_income) = checked      'there used to be radio buttons and so I had to make some connections
                                If LIST_OF_INCOME_ARRAY(future_check, all_income) = TRUE Then LIST_OF_INCOME_ARRAY(budget_in_SNAP_yes, all_income) = unchecked              'future checks are not used to make averages/budget
                                If LIST_OF_INCOME_ARRAY(budget_in_SNAP_yes, all_income) = checked Then          'if thise is counted for the SNAP budget

                                    LIST_OF_INCOME_ARRAY(gross_amount, all_income) = LIST_OF_INCOME_ARRAY(gross_amount, all_income) * 1                 'make this a number for math
                                    net_amount = LIST_OF_INCOME_ARRAY(gross_amount, all_income) - LIST_OF_INCOME_ARRAY(exclude_amount, all_income)      'reduce by any excluded amount that was listed
                                    total_of_counted_income = total_of_counted_income + net_amount                                                      'create a total of all the income to use in determinging ongoing estimates
                                    total_of_included_pay_checks = total_of_included_pay_checks +  LIST_OF_INCOME_ARRAY(gross_amount, all_income)       'another total for all of the checks gross total
                                    number_of_checks_budgeted = number_of_checks_budgeted + 1                                                           'counting the checks

                                    LIST_OF_INCOME_ARRAY(hours, all_income) = LIST_OF_INCOME_ARRAY(hours, all_income) * 1                               'making this a number
                                    total_of_hours = total_of_hours + LIST_OF_INCOME_ARRAY(hours, all_income)                                           'adding up all the hours
                                Else
                                    list_of_excluded_pay_dates = list_of_excluded_pay_dates & ", " & LIST_OF_INCOME_ARRAY(view_pay_date, all_income)    'making a list of all the checks that were not included in making the budget
                                End If
                            End If
                        Next

                        'This is a whole lot of math and formatting
                        EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel) = all_total_hours / cash_checks       'Average hours per paycheck for non-SNAP
                        EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel) = FormatNumber(EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel), 2,,0)

                        EARNED_INCOME_PANELS_ARRAY(snap_ave_hrs_per_pay, ei_panel) = total_of_hours / number_of_checks_budgeted     'average hours per pay check for SNAP
                        EARNED_INCOME_PANELS_ARRAY(snap_ave_hrs_per_pay, ei_panel) = FormatNumber(EARNED_INCOME_PANELS_ARRAY(snap_ave_hrs_per_pay, ei_panel), 2,,0)

                        EARNED_INCOME_PANELS_ARRAY(snap_ave_inc_per_pay, ei_panel) = total_of_counted_income / number_of_checks_budgeted            'average pay $ per paycheck for SNAP
                        EARNED_INCOME_PANELS_ARRAY(snap_ave_inc_per_pay, ei_panel) = FormatNumber(EARNED_INCOME_PANELS_ARRAY(snap_ave_inc_per_pay, ei_panel),2,,0)

                        EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel) = total_of_gross_income / cash_checks             'average $ per pay check for non-SNAP
                        EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel) = FormatNumber(EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel),2,,0)

                        EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel) = total_of_included_pay_checks / total_of_hours           'the $/hr
                        EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel) = FormatNumber(EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel),2,,0)

                        'determining the number of hours per week for SNAP
                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "1 - One Time Per Month" Then EARNED_INCOME_PANELS_ARRAY(snap_hrs_per_wk, ei_panel) = EARNED_INCOME_PANELS_ARRAY(snap_ave_hrs_per_pay, ei_panel)/4.3
                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "2 - Two Times Per Month" Then EARNED_INCOME_PANELS_ARRAY(snap_hrs_per_wk, ei_panel) = (EARNED_INCOME_PANELS_ARRAY(snap_ave_hrs_per_pay, ei_panel)*2)/4.3
                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "3 - Every Other Week" Then EARNED_INCOME_PANELS_ARRAY(snap_hrs_per_wk, ei_panel) = EARNED_INCOME_PANELS_ARRAY(snap_ave_hrs_per_pay, ei_panel)/2
                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "4 - Every Week" Then EARNED_INCOME_PANELS_ARRAY(snap_hrs_per_wk, ei_panel) = EARNED_INCOME_PANELS_ARRAY(snap_ave_hrs_per_pay, ei_panel)
                        EARNED_INCOME_PANELS_ARRAY(snap_hrs_per_wk, ei_panel) = FormatNumber(EARNED_INCOME_PANELS_ARRAY(snap_hrs_per_wk, ei_panel), 2,,0)

                        'determining the number of hours per week for non-SNAP
                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "1 - One Time Per Month" Then EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) = EARNED_INCOME_PANELS_ARRAY(snap_ave_hrs_per_pay, ei_panel)/4.3
                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "2 - Two Times Per Month" Then EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) = (EARNED_INCOME_PANELS_ARRAY(snap_ave_hrs_per_pay, ei_panel)*2)/4.3
                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "3 - Every Other Week" Then EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) = EARNED_INCOME_PANELS_ARRAY(snap_ave_hrs_per_pay, ei_panel)/2
                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "4 - Every Week" Then EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) = EARNED_INCOME_PANELS_ARRAY(snap_ave_hrs_per_pay, ei_panel)
                        EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) = FormatNumber(EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel), 2,,0)

                        If list_of_excluded_pay_dates <> "" Then list_of_excluded_pay_dates = right(list_of_excluded_pay_dates, len(list_of_excluded_pay_dates) - 2)        'formatting this list to remove the leading ", "
                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) <> "" Then            'identifying the multiplier to determine monthly anticipated pay
                            pay_multiplier = 0
                            If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "1 - One Time Per Month" Then pay_multiplier = 1
                            If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "2 - Two Times Per Month" Then pay_multiplier = 2
                            If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "3 - Every Other Week" Then pay_multiplier = 2.15
                            If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "4 - Every Week" Then pay_multiplier = 4.3
                            EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel) = pay_multiplier * EARNED_INCOME_PANELS_ARRAY(snap_ave_inc_per_pay, ei_panel)     'SNAP monthly income

                        End If
                        If EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel) = "" OR EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel) = 0 THen EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel) = "?"
                        EARNED_INCOME_PANELS_ARRAY(GRH_mo_inc, ei_panel) = pay_multiplier * EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel)       'GRH monthly income
                        EARNED_INCOME_PANELS_ARRAY(income_list_indct, ei_panel) = right(EARNED_INCOME_PANELS_ARRAY(income_list_indct, ei_panel), len(EARNED_INCOME_PANELS_ARRAY(income_list_indct, ei_panel))-1)

                    End If          'If EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_actual Then

                    If EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_estimate Then           'If we are going to use the estimate we need some additional detail

                        using_30_days = TRUE            'there is no need to explain if an estimate is being used
                        paycheck_list_title = "Anticipated Paychecks for " & EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel) & "/" & EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel) & ":"        'setting the wording for the CONFIRM BUDGET Dialog

                        the_first_of_CM_2 = CM_plus_2_mo & "/1/" & CM_plus_2_yr     'this is setting start and end dates for creating a list
                        CM_2_mo = DatePart("m", the_first_of_CM_2)
                        CM_2_yr = DatePart("yyyy", the_first_of_CM_2)
                        the_initial_month = DateValue(EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel) & "/1/" & EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel))

                        EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel) = EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel)        'there are 2 positions for this. I've left it as is for now
                        EARNED_INCOME_PANELS_ARRAY(snap_hrs_per_wk, ei_panel) =  EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel)

                        days_to_add = 0     'this is for counting one check to the next
                        months_to_add = 0
                        Select Case EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel)          'here we determine averages of hours and income to anticipate based on pay frequency
                            Case "1 - One Time Per Month"
                                EARNED_INCOME_PANELS_ARRAY(snap_ave_inc_per_pay, ei_panel) = EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel) * EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) * 4.3
                                EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel) = EARNED_INCOME_PANELS_ARRAY(snap_ave_inc_per_pay, ei_panel)
                                EARNED_INCOME_PANELS_ARRAY(snap_ave_hrs_per_pay, ei_panel) = EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) * 4.3
                                days_to_add = 0
                                months_to_add = 1
                                default_start_date = the_initial_month
                            Case "2 - Two Times Per Month"
                                EARNED_INCOME_PANELS_ARRAY(snap_ave_inc_per_pay, ei_panel) = EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel) * EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) * 4.3 / 2
                                EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel) = EARNED_INCOME_PANELS_ARRAY(snap_ave_inc_per_pay, ei_panel) * 2
                                EARNED_INCOME_PANELS_ARRAY(snap_ave_hrs_per_pay, ei_panel) = (EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) * 4.3)/2
                                days_to_add = 30
                                months_to_add = 1
                                default_start_date = the_initial_month
                            Case "3 - Every Other Week"
                                EARNED_INCOME_PANELS_ARRAY(snap_ave_inc_per_pay, ei_panel) = EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel) * EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) * 2
                                EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel) = EARNED_INCOME_PANELS_ARRAY(snap_ave_inc_per_pay, ei_panel) * 2.15
                                EARNED_INCOME_PANELS_ARRAY(snap_ave_hrs_per_pay, ei_panel) = EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) * 2
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
                                EARNED_INCOME_PANELS_ARRAY(snap_ave_inc_per_pay, ei_panel) = EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel) * EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel)
                                EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel) = EARNED_INCOME_PANELS_ARRAY(snap_ave_inc_per_pay, ei_panel) * 4.3
                                EARNED_INCOME_PANELS_ARRAY(snap_ave_hrs_per_pay, ei_panel) = EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel)
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
                        'the non-snap averages should be based on actual checks unless none were provided. This defines them if there were no actual checks
                        If EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel) = "" Then EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel) = EARNED_INCOME_PANELS_ARRAY(snap_ave_inc_per_pay, ei_panel)
                        If EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel) = "" Then EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel) = EARNED_INCOME_PANELS_ARRAY(snap_ave_hrs_per_pay, ei_panel)
                        If EARNED_INCOME_PANELS_ARRAY(GRH_mo_inc, ei_panel) = "" Then EARNED_INCOME_PANELS_ARRAY(GRH_mo_inc, ei_panel) = EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel)

                        snap_anticipated_pay_array = ""     'blanking these out because of looping
                        checks_list = ""

                        Call Navigate_to_MAXIS_screen("STAT", "JOBS")           'making sure that we are still at the right job - it may be possible to REMOVE CODE
                        EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel), 20, 76
                        EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel), 20, 79
                        transmit
                        this_pay_date = ""      'blaning this out for each paenl/loop
                        If list_of_actual_paydates <> "" Then       'if there were actual dates provided we will use one of those to determine an accurate date
                            paydates_array = split(list_of_actual_paydates, "~")        'making this an array and setting the first check date to the fisrt one
                            this_pay_date = paydates_array(0)
                            'this position in the array is attempted to be defined previously when making a chronological list so it should actually be filled
                            If EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel) = "" Then EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel) = this_pay_date
                        ElseIf known_pay_date <> "" Then        'this is a field on the ENTER PAY Dialog that the worker can enter
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
                        If this_pay_date = "" Then this_pay_date = default_start_date       'this is for makin gour list

                        save_dates = FALSE
                        Do      'While DatePart("m", this_pay_date) <> CM_2_mo AND DatePart("yyyy", this_pay_date) <> CM_2_yr
                            save_dates = FALSE
                            'if the date we are looking at is for the initial month - then we are going to save it to a list.
                            If DatePart("m", this_pay_date) = DatePart("m", the_initial_month) AND DatePart("yyyy", this_pay_date) = DatePart("yyyy", the_initial_month) Then save_dates = TRUE
                            If save_dates = TRUE Then

                                If EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel) = "" Then EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel) = this_pay_date    'saving this if there was nothing

                                check_found = FALSE         'looking to see if there was an actual check for this date
                                For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                                    If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then
                                        If DateValue(LIST_OF_INCOME_ARRAY(pay_date, all_income)) = this_pay_date Then
                                            check_found = TRUE          'if there was then we will save that information for our list
                                            check_number = all_income   'need to know what position it is at
                                            Exit For
                                        End If
                                    End If
                                Next
                                'BUGGY CODE - we may need better handling for creating a list for the non-snap program sidplays
                                If check_found = TRUE Then  'if the check was listed the information will include the actual amount listed
                                    'this different formatting is just to make it pretty
                                    If len(this_pay_date) = 10 Then checks_list = checks_list & "%" & this_pay_date & "   ~   $" & LIST_OF_INCOME_ARRAY(gross_amount, check_number)
                                    If len(this_pay_date) = 9 Then checks_list = checks_list & "%" & this_pay_date & "    ~   $" & LIST_OF_INCOME_ARRAY(gross_amount, check_number)
                                    If len(this_pay_date) = 8 Then checks_list = checks_list & "%" & this_pay_date & "     ~   $" & LIST_OF_INCOME_ARRAY(gross_amount, check_number)
                                Else            'otherwise it includes a paycheck average
                                    If len(this_pay_date) = 10 Then checks_list = checks_list & "%" & this_pay_date & "   ~   $" & EARNED_INCOME_PANELS_ARRAY(snap_ave_inc_per_pay, ei_panel)
                                    If len(this_pay_date) = 9 Then checks_list = checks_list & "%" & this_pay_date & "    ~   $" & EARNED_INCOME_PANELS_ARRAY(snap_ave_inc_per_pay, ei_panel)
                                    If len(this_pay_date) = 8 Then checks_list = checks_list & "%" & this_pay_date & "     ~   $" & EARNED_INCOME_PANELS_ARRAY(snap_ave_inc_per_pay, ei_panel)
                                End If
                            End If
                            If months_to_add = 0 Then       'these are defined by the pay frequency above and will increment us to the next pay date
                                this_pay_date = DateAdd("d", days_to_add, this_pay_date)
                            ElseIf days_to_add = 0 Then
                                this_pay_date = DateAdd("m", months_to_add, this_pay_date)
                            Else
                                If EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = "LAST" Then
                                    month_ahead = DateAdd("m", 1, this_pay_date)
                                    month_to_use = DatePart("m", month_ahead)
                                    year_to_use = DatePart("yyyy", month_ahead)
                                    If DatePart("d", this_pay_date) = EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) Then
                                        first_of_nextMonth = month_to_use & "/1/" & year_to_use
                                        this_pay_date = DateAdd("d", -1, first_of_nextMonth)
                                    Else
                                        this_pay_date = month_to_use & "/" & EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) & "/" & year_to_use
                                    End If
                                Else
                                    If DatePart("d", this_pay_date) = EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) Then
                                        month_to_use = DatePart("m", this_pay_date)
                                        year_to_use = DatePart("yyyy", this_pay_date)
                                        this_pay_date = month_to_use & "/" & EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) & "/" & year_to_use
                                    ElseIf DatePart("d", this_pay_date) = EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) Then
                                        month_ahead = DateAdd("m", 1, this_pay_date)
                                        month_to_use = DatePart("m", month_ahead)
                                        year_to_use = DatePart("yyyy", month_ahead)
                                        this_pay_date = month_to_use & "/" & EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) & "/" & year_to_use
                                    End If
                                End If
                            End If
                        Loop until DatePart("m", this_pay_date) = CM_2_mo AND DatePart("yyyy", this_pay_date) = CM_2_yr     'stop at current month plus 2

                        'Formatting the list and making it an array
                        If left(checks_list, 1) = "%" Then checks_list = right(checks_list, len(checks_list)-1)
                        If InStr(checks_list, "%") <> 0 Then
                            snap_anticipated_pay_array = Split(checks_list,"%")
                        Else
                            snap_anticipated_pay_array = Array(checks_list)
                        End If
                    End If

                    'This is the CONFIRM BUDGET Dialog - the height has given me problems so hopefully it is working well.
                    Do          'Loop until are_we_passworded_out = False
                        word_for_freq = ""      'for displaying in the dialog
                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "1 - One Time Per Month" Then word_for_freq = "monthly"
                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "2 - Two Times Per Month" Then word_for_freq = "semi-monthly"
                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "3 - Every Other Week" Then word_for_freq = "biweekly"
                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "4 - Every Week" Then word_for_freq = "weekly"

                        dlg_len = 35        'starting with this dialog
                        'FUTURE FUNCTIONALITY - maybe we add some information that summarizes what was entered on ENTER PAY Dialog - but we might not have reoom

                        If EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) = checked Then       'resizing the dialog and the SNAP Groupbox if income applies to SNAP
                            grp_len = 100 + number_of_checks_budgeted*10
                            If number_of_checks_budgeted < 4 Then grp_len = 125
                            If using_30_days = FALSE Then grp_len = grp_len + 25

                            dlg_len = dlg_len + grp_len + 20
                        End If
                        If EARNED_INCOME_PANELS_ARRAY(income_verif, ei_panel) = "? - EXPEDITED SNAP ONLY" Then      'adding size for XFS information tot be added to dialog
                            dlg_len = dlg_len + 40
                        End If
                        If EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, ei_panel) = checked Then       'resizing the dialog and the Cash Groupbox if income applies to Cash
                            dlg_len = dlg_len + 20
                            cash_grp_len = 50
                            length_of_checks_list = cash_checks*10
                            If cash_checks = 0 Then length_of_checks_list = (UBound(snap_anticipated_pay_array) + 1)*10
                            If length_of_checks_list < 40 Then length_of_checks_list = 35

                            cash_grp_len = cash_grp_len + length_of_checks_list
                            dlg_len = dlg_len + cash_grp_len
                        End If
                        If EARNED_INCOME_PANELS_ARRAY(apply_to_HC, ei_panel) = checked Then         'resizing the dialog and the HC Groupbox if income applies to HC
                            dlg_len = dlg_len + 20
                            hc_grp_len = 60
                            length_of_checks_list = cash_checks*10
                            If length_of_checks_list < 60 Then length_of_checks_list = 60

                            hc_grp_len = hc_grp_len + length_of_checks_list
                            dlg_len = dlg_len + hc_grp_len
                        End If
                        If EARNED_INCOME_PANELS_ARRAY(apply_to_GRH, ei_panel) = checked Then        'resizing the dialog and the GRH Groupbox if income applies to GRH
                            dlg_len = dlg_len + 30
                            grh_grp_len = 65
                            length_of_checks_list = cash_checks*10
                            If length_of_checks_list = 0 Then length_of_checks_list = 20

                            grh_grp_len = grh_grp_len + length_of_checks_list
                            dlg_len = dlg_len + grh_grp_len
                        End If

                        y_pos = 35      'incrementer to move things down
                        'CONFIRM BUDGET Dialog - mostly shows the information after being calculated for each program and makes the worker confirm this is correct
                        BeginDialog Dialog1, 0, 0, 421, dlg_len, "Confirm JOBS Budget"
                          Text 10, 10, 250, 10, "JOBS " & EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel) & " " & EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel) & " - " & EARNED_INCOME_PANELS_ARRAY(employer, ei_panel)
                          Text 10, 20, 150, 10, "Pay Frequency - " & EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel)
                          CheckBox 160, 20, 200, 10, "Check here to confirm this pay frequency.", confirm_pay_freq_checkbox

                          If EARNED_INCOME_PANELS_ARRAY(income_verif, ei_panel) = "? - EXPEDITED SNAP ONLY" Then
                              Text 10, y_pos, 400, 10, "THIS INCOME HAS NOT BEEN VERIFIED - '?' verification code used."
                              Text 10, y_pos +10, 400, 10, " -- Only SNAP can be handled this way. The script will only apply SNAP budgeting functionality.-- "
                              Text 10, y_pos + 20, 400, 10, "A note will be added that some or all of pay information is only reported by client and not verified."
                              y_pos = y_pos + 40
                          End If

                          If EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) = checked Then
                              GroupBox 5, y_pos, 410, grp_len, "SNAP Budget"

                              Text 10, y_pos + 10, 400, 10, "Income provided covers the period " & first_date & " to " & last_date & ". This income covers " & spread_of_pay_dates & " days."
                              If using_30_days = FALSE Then
                                  y_pos = y_pos + 25
                                  Text 10, y_pos + 5, 175, 10, "It appears this is not 30 days of income. Explain:"
                                  EditBox 185, y_pos, 200, 15, not_30_explanation
                              End If
                              y_pos = y_pos + 10

                              Text 10, y_pos + 10, 150, 10, paycheck_list_title        '"Paychecks Inclued in Budget:"'
                              list_pos = 0      'multiplier to move the array items down

                              If EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_actual Then
                                  ' 'this part actually looks at the income information IN ORDER
                                  For order_number = 1 to top_of_order                        'loop through the order number lowest to highest
                                      For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                          'conditional if it is the right panel AND the order matches - then do the thing you need to do
                                          If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(check_order, all_income) = order_number Then
                                              If LIST_OF_INCOME_ARRAY(exclude_amount, all_income) = 0 Then Text 15, (list_pos * 10) + y_pos + 25, 160, 10, LIST_OF_INCOME_ARRAY(view_pay_date, all_income) & " - $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & "hrs."
                                              If LIST_OF_INCOME_ARRAY(exclude_amount, all_income) <> 0 Then Text 15, (list_pos * 10) + y_pos + 25, 160, 10, LIST_OF_INCOME_ARRAY(view_pay_date, all_income) & " - $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & "hrs. - $" & LIST_OF_INCOME_ARRAY(exclude_amount, all_income) & " not included."

                                              list_pos = list_pos + 1
                                          End If
                                      next
                                  next
                              ElseIf EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_estimate Then
                                  For each money_day in snap_anticipated_pay_array      'this is the list we made above - it has pay date and amount in each item
                                      Text 20, (list_pos * 10) + y_pos + 25, 90, 10, money_day
                                      list_pos = list_pos + 1
                                  Next
                              End If
                              If list_pos < 3 Then list_pos = 3

                              Text 185, y_pos + 10, 200, 10, "Average hourly rate of pay: $" & EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel)
                              Text 185, y_pos + 25, 200, 10, "Average " & word_for_freq & " hours: " & EARNED_INCOME_PANELS_ARRAY(snap_ave_hrs_per_pay, ei_panel)
                              Text 185, y_pos + 40, 200, 10, "Average paycheck amount: $" & EARNED_INCOME_PANELS_ARRAY(snap_ave_inc_per_pay, ei_panel)
                              Text 185, y_pos + 55, 200, 10, "Monthly Budgeted Income: $" & EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel)
                              If EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel) = "?" Then       'REMOVE CODE - this should not be coming up any more - need to confirm
                                ButtonGroup ButtonPressed
                                    PushButton 305, y_pos + 55, 60, 10, "Calculate", calc_btn
                              End If
                              y_pos = y_pos + 65
                              If list_pos > 4 Then y_pos = y_pos + ((list_pos-4) * 10)
                              Text 10, y_pos, 400, 10, "Paychecks not included: " & list_of_excluded_pay_dates      'list of all excluded pay dates

                              CheckBox 10, y_pos + 15, 330, 10, "Check here if you confirm that this budget is correct and is the best estimate of anticipated income.", confirm_budget_checkbox
                              Text 10, y_pos + 35, 60, 10, "Conversation with:"
                              ComboBox 75, y_pos + 30, 60, 45, " "+chr(9)+"Client - not employee"+chr(9)+"Employee"+chr(9)+"Employer",  EARNED_INCOME_PANELS_ARRAY(spoke_with, ei_panel)
                              Text 140, y_pos + 35, 25, 10, "clarifies"
                              EditBox 170, y_pos + 30, 235, 15, EARNED_INCOME_PANELS_ARRAY(convo_detail, ei_panel)
                              y_pos = y_pos + 55
                          Else      'if income does not apply to SNAP, we have to default these to being completed
                            confirm_budget_checkbox = checked
                            using_30_days = TRUE
                          End If        'If EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) = checked Then

                          If EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, ei_panel) = checked Then
                              GroupBox 5, y_pos, 410, cash_grp_len, "CASH Budget"
                              y_pos = y_pos + 15
                              Text 10, y_pos, 400, 10, "Pay information will be entered on the RETRO side if provided. The script will not calculate an average for any Retro pay dates."
                              Text 10, y_pos + 10, 400, 10, "For each month to be updated, the script will use actual pay information or the average for that month on the prospective side."
                              y_pos = y_pos + 20
                              Text 10, y_pos, 30, 10, "CHECKS"
                              y_pos = y_pos + 10

                              CheckBox 190, y_pos, 220, 10, "Check here if these checks are accurate and should be entered.", confirm_checks_checkbox
                              Text 190, y_pos + 10, 220, 10, "If this income is excluded from the Cash budget, select the reason:"
                              DropListBox 190, y_pos + 20, 150, 15, "NONE"+chr(9)+"Caregiver under 20 - 50% in school"+chr(9)+"Child under 18 in school"+chr(9)+"Excluded Work Program"+chr(9)+"Excluded Spousal Income", income_excluded_cash_reason

                              list_pos = 0
                              ' 'this part actually looks at the income information IN ORDER
                              For order_number = 1 to top_of_order                        'loop through the order number lowest to highest
                                  For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                      'conditional if it is the right panel AND the order matches - then do the thing you need to do
                                      If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(check_order, all_income) = order_number Then
                                          'list_of_dates = list_of_dates & vbNewLine & "Check Date: " & LIST_OF_INCOME_ARRAY(pay_date, all_income) & " Income: $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " Hours: " & LIST_OF_INCOME_ARRAY(hours, all_income)
                                          Text 20, (list_pos * 10) + y_pos, 170, 10, LIST_OF_INCOME_ARRAY(view_pay_date, all_income) & " - $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & "hrs."
                                          list_pos = list_pos + 1
                                      End If
                                  next
                              next

                              If list_pos = 0 Then
                                  For each money_day in snap_anticipated_pay_array
                                      money_day_date = left(money_day, 10)      'just using the date from the snap list
                                      money_day_date = trim(money_day_date)
                                      Text 20, (list_pos * 10) + y_pos, 170, 10, money_day_date & " - $" & EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel) & " - " & EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel) & "hrs."
                                      list_pos = list_pos + 1
                                  Next
                              End If

                              if list_pos < 3 Then list_pos = 3
                              bottom_of_checks = y_pos + (list_pos * 10)
                              y_pos = bottom_of_checks + 10

                          Else
                            confirm_checks_checkbox = checked       'defaulting to this if not Cash
                          End If            'If EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, ei_panel) = checked Then

                          If EARNED_INCOME_PANELS_ARRAY(apply_to_HC, ei_panel) = checked Then
                              GroupBox 5, y_pos, 410, hc_grp_len, "Health Care Budget"
                              y_pos = y_pos + 15
                              Text 10, y_pos, 400, 10, "Pay information will be entered on the prospective side only, using actual or estimated pay amounts."
                              y_pos = y_pos + 10
                              Text 10, y_pos, 30, 10, "CHECKS"

                              y_pos = y_pos + 10
                              Text 150, y_pos, 250, 10, "Average amount per pay period: $" & EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel) & " - for HC Inc Est Pop-up."

                              Text 150, y_pos + 15, 200,10, "Notes about HC Budget:"
                              EditBox 150, y_pos + 25, 250, 15, EARNED_INCOME_PANELS_ARRAY(hc_budg_notes, ei_panel)

                              CheckBox 150, y_pos + 45, 200, 10, "Check here if HC needs a Retrospective Budget.", hc_retro_budget_checkbox

                              list_pos = 0
                              ' 'this part actually looks at the income information IN ORDER
                              For order_number = 1 to top_of_order                        'loop through the order number lowest to highest
                                  For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                      'conditional if it is the right panel AND the order matches - then do the thing you need to do
                                      If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(check_order, all_income) = order_number Then
                                          'list_of_dates = list_of_dates & vbNewLine & "Check Date: " & LIST_OF_INCOME_ARRAY(view_pay_date, all_income) & " Income: $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " Hours: " & LIST_OF_INCOME_ARRAY(hours, all_income)
                                          Text 20, (list_pos * 10) + y_pos, 90, 10, LIST_OF_INCOME_ARRAY(view_pay_date, all_income) & " - $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & "hrs."
                                          list_pos = list_pos + 1
                                      End If
                                  next
                              next

                              bottom_of_checks = y_pos + (list_pos * 10)
                              If list_pos < 6 Then bottom_of_checks = y_pos + 60
                              y_pos = bottom_of_checks + 5

                              CheckBox 10, y_pos, 250, 10, "Check here if these checks and estimated pay amount are accurate.", hc_confirm_checks_checkbox
                              y_pos = y_pos + 25
                          Else
                            hc_confirm_checks_checkbox = checked        'default if income does not apply to HC
                          End If

                          If EARNED_INCOME_PANELS_ARRAY(apply_to_GRH, ei_panel) = checked Then
                              GroupBox 5, y_pos, 410, grh_grp_len, "GRH Budget"
                              y_pos = y_pos + 10

                              Text 10, y_pos, 150, 10, paycheck_list_title
                              y_pos = y_pos + 15

                              Text 185, y_pos, 200, 10, "Average paycheck amount: $" & EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel)
                              Text 185, y_pos + 10, 200, 10, "Monthly Budgeted Income: $" & EARNED_INCOME_PANELS_ARRAY(GRH_mo_inc, ei_panel)

                              list_pos = 0
                              If EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_actual Then
                                  Text 10, y_pos, 150, 10, "Actual Check Stubs to Enter on GRH PIC"
                                  y_pos = y_pos + 15
                                  ' 'this part actually looks at the income information IN ORDER
                                  For order_number = 1 to top_of_order                        'loop through the order number lowest to highest
                                      For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                          'conditional if it is the right panel AND the order matches - then do the thing you need to do
                                          If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(check_order, all_income) = order_number Then
                                              'list_of_dates = list_of_dates & vbNewLine & "Check Date: " & LIST_OF_INCOME_ARRAY(view_pay_date, all_income) & " Income: $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " Hours: " & LIST_OF_INCOME_ARRAY(hours, all_income)
                                              Text 15, (list_pos * 10) + y_pos, 160, 10, LIST_OF_INCOME_ARRAY(view_pay_date, all_income) & " - $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & "hrs."

                                              list_pos = list_pos + 1
                                          End If
                                      next
                                  next
                                  y_pos = y_pos + (list_pos * 10) + 10
                              ElseIf EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_estimate Then
                                  Text 10, y_pos, 150, 10, "Income Estimate Enter on GRH PIC"

                                  Text 20, y_pos + 10, 90, 10, "Hours Per Week: " & EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel)
                                  Text 20, y_pos + 20, 90, 10, "Rate Of Pay/Hr: $" & EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel)

                                  For each money_day in snap_anticipated_pay_array      'this is using SNAP income information - BUGGY CODE
                                      Text 20, (list_pos * 10) + y_pos + 30, 90, 10, money_day
                                      list_pos = list_pos + 1
                                  Next
                                  y_pos = y_pos + 15
                              End If
                              'FUTURE FUNCTIONALITY - add an EditBox for entering an amount for the application month ONLY that is not counted (there is a field on JOBS)
                              CheckBox 10, y_pos, 330, 10, "Check here if you confirm that this budget is correct and is the best estimate of anticipated income.", GRH_confirm_budget_checkbox
                              y_pos = y_pos + 20
                          Else
                              GRH_confirm_budget_checkbox = checked     'default if income does not apply to GRH
                          End If

                          ButtonGroup ButtonPressed
                            OkButton 315, y_pos, 50, 15
                            CancelButton 365, y_pos, 50, 15
                        EndDialog

                        Dialog Dialog1      'calling the dialog

                        Call check_for_password(are_we_passworded_out)  'we are doing password handling before error handling because of the 2 dialogs looped together
                    Loop until are_we_passworded_out = False
                    If ButtonPressed = 0 then       'if the 'Cancel' button is pressed, the worker gets 3 options 1. cancel script, 2, cancel current job update, 3. Ooops, pressed cancel by mistake
                        cancel_clarify = MsgBox("Do you want to stop the script entirely?" & vbNewLine & vbNewLine & "If the script is stopped no information provided so far will be updated or noted. If you choose 'No' the update for THIS JOB will be cancelled and rest of the script will continue." & vbNewLine & vbNewLine & "YES - Stop the script entirely." & vbNewLine & "NO - Do not stop the script entrirely, just cancel the entry of this job information."& vbNewLine & "CANCEL - I didn't mean to cancel at all. (Cancel my cancel)", vbQuestion + vbYesNoCancel, "Clarify Cancel")
                        If cancel_clarify = vbYes Then script_end_procedure("~PT: user pressed cancel")
                        'script_end_procedure text added for statistical purposes. If script was canceled prior to completion, the statistics will reflect this.
                    End if
                    If cancel_clarify = vbNo Then           'this is to cancel the job update
                        review_small_dlg = FALSE
                        EARNED_INCOME_PANELS_ARRAY(income_received, ei_panel) = FALSE   'this makes the script skip this job in the next functions
                        Exit Do
                    End If          'there is no vbCancel handling because the script just continues at that point.

                    If EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel) = "?" Then
                        big_err_msg = big_err_msg & vbNewLine & " The script could not determine the Monthly SNAP Budget for this income. Use the 'Calculate' button to review and approve the SNAP budget. "
                    End If
                    If confirm_pay_freq_checkbox = unchecked Then       'mandating that the pay frequency is accurate
                        big_err_msg = big_err_msg & vbNewLine & "* Review the pay frequency and confirm."
                        review_small_dlg = TRUE
                    End If
                    If confirm_budget_checkbox = unchecked then         'mandating the SNAP budget was checked
                        big_err_msg = big_err_msg & vbNewLine & "*** Since the budget is not confirmed as correct, the ENTER PAY INFORMATION DIALOG will reappear and allow information to be corrected to generate an accurate budget. ***"
                        review_small_dlg = TRUE
                    End If
                    If confirm_checks_checkbox = unchecked Then         'mandating the cash information was checked
                        big_err_msg = big_err_msg & vbNewLine & "*** If the checks are not accurate, review them and update as necessary. ***"
                        review_small_dlg = TRUE
                    End If
                    If hc_confirm_checks_checkbox = unchecked Then      'mandating the hc information was checked
                        big_err_msg = big_err_msg & vbNewLine & "*** If the checks or HC Income Estimate are not accurate, review them and update as necessary. ***"
                        review_small_dlg = TRUE
                    End If
                    If GRH_confirm_budget_checkbox = unchecked Then     'mandating that the GRH budget was checked
                        big_err_msg = big_err_msg & vbNewLine & "*** Since the GRH budget is not confirmed accurate, the ENTER PAY INFORMATION DIALOG will reappear. Review pay information and update as necessary. ***"
                        review_small_dlg = TRUE
                    End If
                    If using_30_days = FALSE Then                   'if checks do not span 30 days and SNAP is selected, there must be an explanation
                        If not_30_explanation = "" Then big_err_msg = big_err_msg & vbNewLine & "** Since income received is not 30 days of income for SNAP, it must be explained why we are accepting more or less."
                    End If
                    If ButtonPressed = calc_btn Then            'REMOVE CODE - this should no longer come up. Need to review functionality
                        review_small_dlg = FALSE
                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "" Then
                            big_err_msg = big_err_msg & vbNewLine & "** List the pay frequency for this income."
                        Else
                            big_err_msg = "LOOP" & big_err_msg
                        End If
                    End If

                    If cancel_clarify = vbNo Then big_err_msg = ""      'removing any errors if this job update is being canceled
                    If big_err_msg <> "" Then                           'if there is an error, the script will display the error message and make the information ready to be viewed in the ENTER PAY Dialog
                        For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                            If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then
                                LIST_OF_INCOME_ARRAY(gross_amount, all_income) = LIST_OF_INCOME_ARRAY(gross_amount, all_income) & ""
                                LIST_OF_INCOME_ARRAY(hours, all_income) = LIST_OF_INCOME_ARRAY(hours, all_income) & ""
                                LIST_OF_INCOME_ARRAY(exclude_amount, all_income) = LIST_OF_INCOME_ARRAY(exclude_amount, all_income) & ""
                            End If
                        Next
                        If left(big_err_msg, 4) <> "LOOP" Then MsgBox "Review JOBS Pay Information" & vbNewLine & big_err_msg
                    End If

                    'If the initial month is not CM+1 and update future checkbox is unchecked, this is going to confirm that is correct
                    review_update_future_mos = TRUE
                    If big_err_msg = "" Then
                        If EARNED_INCOME_PANELS_ARRAY(update_futue_chkbx, ei_panel) = unchecked Then
                            If EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel) = CM_plus_1_mo AND EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel) = CM_plus_1_yr Then review_update_future_mos = FALSE

                            If review_update_future_mos = TRUE Then
                                confirm_do_not_update = MsgBox("You have selected to NOT update future months for this job." & vbNewLine & "You are starting the update of this job in " & EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel) & "/" & EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel) & "." & vbNewLine & vbNewLine & "Are you sure you want to continue?", vbQuestion + vbYesNo, "Confirm do NOT update future months")
                                If confirm_do_not_update = vbNo Then        'if this is not correct, the script will go back to ENTER PAY Dialog to have the box checked
                                    big_err_msg = "LOOP"
                                    review_small_dlg = TRUE
                                    For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                                        If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then
                                            LIST_OF_INCOME_ARRAY(gross_amount, all_income) = LIST_OF_INCOME_ARRAY(gross_amount, all_income) & ""
                                            LIST_OF_INCOME_ARRAY(hours, all_income) = LIST_OF_INCOME_ARRAY(hours, all_income) & ""
                                            LIST_OF_INCOME_ARRAY(exclude_amount, all_income) = LIST_OF_INCOME_ARRAY(exclude_amount, all_income) & ""
                                        End If
                                    Next
                                End If
                            End If
                        End If
                    End If

                    If hc_retro_budget_checkbox = checked Then EARNED_INCOME_PANELS_ARRAY(hc_retro, ei_panel) = TRUE
                    If hc_retro_budget_checkbox = unchecked Then EARNED_INCOME_PANELS_ARRAY(hc_retro, ei_panel) = FALSE

                Loop until big_err_msg = ""
                call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
            LOOP UNTIL are_we_passworded_out = false
            'WE ARE OUT OF THE DIALOGS

            'Setting some information and formatting and saving it fot the next functionality.
            'each and every thing has to be IN THE ARRAY to be saved
            If EARNED_INCOME_PANELS_ARRAY(antic_pay_list, ei_panel) <> "" Then EARNED_INCOME_PANELS_ARRAY(antic_pay_list, ei_panel) = Join(snap_anticipated_pay_array, "%*%")
            If EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_actual Then
                EARNED_INCOME_PANELS_ARRAY(days_of_verif, ei_panel) = "Pay verifications covers the period " & first_date & " to " & last_date & " which is " & spread_of_pay_dates & " days. "
                If using_30_days = FALSE Then
                    EARNED_INCOME_PANELS_ARRAY(days_of_verif, ei_panel) = EARNED_INCOME_PANELS_ARRAY(days_of_verif, ei_panel) & "This is not 30 days, we are not using 30 days because: " & not_30_explanation
                End If
            End If

            If EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) = checked Then check_SNAP_for_UH = TRUE
            If income_excluded_cash_reason <> "NONE" Then EARNED_INCOME_PANELS_ARRAY(excl_cash_rsn, ei_panel) = income_excluded_cash_reason

            'FUTURE FUNCTIONALITY - Add handling to check WREG for correct coding based upon this income information
            'FUTURE FUNCTIONALITY - Add handling for future/current changes - start or stop work - get policy on this from SNAP refresher - talk to Melissa.

            script_run_lowdown = script_run_lowdown & vbCr & vbCr & "Information about JOBS entered into dialog"
            script_run_lowdown = script_run_lowdown & vbCr & "Employer - " & EARNED_INCOME_PANELS_ARRAY(employer, ei_panel) & "MEMB " & EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel) & " " & EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel)

            For order_number = 1 to top_of_order                        'loop through the order number lowest to highest
                For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                    'conditional if it is the right panel AND the order matches - then do the thing you need to do
                    If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(check_order, all_income) = order_number Then
                        If LIST_OF_INCOME_ARRAY(budget_in_SNAP_yes, all_income) = unchecked Then
                            script_run_lowdown = script_run_lowdown & vbCr & LIST_OF_INCOME_ARRAY(view_pay_date, all_income) & " - $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & "hrs. CHECK EXCLUDED."
                        ElseIf LIST_OF_INCOME_ARRAY(exclude_amount, all_income) = 0 Then
                            script_run_lowdown = script_run_lowdown & vbCr & LIST_OF_INCOME_ARRAY(view_pay_date, all_income) & " - $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & "hrs."
                        ElseIf LIST_OF_INCOME_ARRAY(exclude_amount, all_income) <> 0 Then
                            script_run_lowdown = script_run_lowdown & vbCr & LIST_OF_INCOME_ARRAY(view_pay_date, all_income) & " - $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & "hrs. - $" & LIST_OF_INCOME_ARRAY(exclude_amount, all_income) & " not included."
                        End If
                    End If
                next
            next
            script_run_lowdown = script_run_lowdown & vbCr & "Programs Applied to: "
            If EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) = checked Then script_run_lowdown = script_run_lowdown & "/SNAP"       'setting the header programs
            If EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, ei_panel) = checked Then script_run_lowdown = script_run_lowdown & "/CASH"
            If EARNED_INCOME_PANELS_ARRAY(apply_to_HC, ei_panel) = checked Then script_run_lowdown = script_run_lowdown & "/HC"
            If EARNED_INCOME_PANELS_ARRAY(apply_to_GRH, ei_panel) = checked Then script_run_lowdown = script_run_lowdown & "/GRH"

            script_run_lowdown = script_run_lowdown & vbCr & "SNAP Detail"
            script_run_lowdown = script_run_lowdown & vbCr & "Average hourly rate of pay: $" & EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel)
            script_run_lowdown = script_run_lowdown & vbCr & "Average " & word_for_freq & " hours: " & EARNED_INCOME_PANELS_ARRAY(snap_ave_hrs_per_pay, ei_panel)
            script_run_lowdown = script_run_lowdown & vbCr & "Average paycheck amount: $" & EARNED_INCOME_PANELS_ARRAY(snap_ave_inc_per_pay, ei_panel)
            script_run_lowdown = script_run_lowdown & vbCr & "Monthly Budgeted Income: $" & EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel)
            script_run_lowdown = script_run_lowdown & vbCr & "All other Program detail"
            script_run_lowdown = script_run_lowdown & vbCr & "Average hourly rate of pay: $" & EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel)
            script_run_lowdown = script_run_lowdown & vbCr & "Average " & word_for_freq & " hours: " & EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel)
            script_run_lowdown = script_run_lowdown & vbCr & "Average paycheck amount: $" & EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel)
            script_run_lowdown = script_run_lowdown & vbCr & "-----------------------------"

        ElseIf EARNED_INCOME_PANELS_ARRAY(this_is_a_new_panel, ei_panel) = TRUE Then
            'FUTURE FUNCTIONALITY - may need some specialized functionality to handle for new panels that do NOT have income detail provided.
        End If

        MAXIS_footer_month = original_month     'resetting this for the next panel
        MAXIS_footer_year = original_year
    End If

    'FUTURE FUNCTIONALITY - this will never be called, the message box is disable but it would be the ENTER PAY and CONFIRM BUDGET for BUSI panels
    'TESTING NEEDED - anything that is here has not been finished or tested
    If EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel) = "BUSI" Then
        'NAVIGATE to BUSI for each HH MEMBER and ask if Income Information was received for this Self Employment.
        Call Navigate_to_MAXIS_screen("STAT", "BUSI")
        EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel), 20, 76
        EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel), 20, 79
        transmit

        employer_check = vbNo
        ' employer_check = MsgBox("Do you have income verification for this self employment? Type of Self Employment: " & EARNED_INCOME_PANELS_ARRAY(income_type, ei_panel), vbYesNo + vbQuestion, "Select Income Panel")

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

    'FUTURE FUNCTIONALITY - NAVIGATE to RBIC for each HH MEMBER and ask if Income Information was received for this RBIC
Next

'If there is SNAP - we need to see if the SNAP is UH - this can be found in ELIG/FS
UH_SNAP = FALSE
IF check_SNAP_for_UH = TRUE then

    If snap_status = "ACTV" Then
        MAXIS_footer_month = CM_mo
        MAXIS_footer_year = CM_yr

        CALL navigate_to_approved_SNAP_eligibility

        EMReadScreen type_of_SNAP, 13, 4, 3
        IF type_of_SNAP = "'UNCLE HARRY'" then
            UH_SNAP = TRUE
        Else
            UH_SNAP = FALSE
        End If
    Else
        ask_about_uncle_harry = MsgBox("It appears SNAP is not yet active on this case. Is this UNCLE HARRY SNAP?", vbQuestion + vbYesNo, "Check for UHFS")
        If ask_about_uncle_harry = vbYes then UH_SNAP = TRUE
    End If

End If
                                '----------------------------------------------------------'
                    '---------------------------------------------------------------------------------'
'------------------------------------------------- DETERMINING WHICH MONTHS TO UPDATE --------------------------------------------------'
                    '---------------------------------------------------------------------------------'
                                '----------------------------------------------------------'

list_of_all_months_to_update = "~"      'start of a list that will become and array
update_with_verifs = FALSE              'defaults to false

For ei_panel = 0 to UBOUND(EARNED_INCOME_PANELS_ARRAY, 2)                       'we will look at each panel that exists
    If EARNED_INCOME_PANELS_ARRAY(income_received, ei_panel) = TRUE Then        'IF the ENTER PAY and CONFIRM BUDGET were completed
        update_with_verifs = TRUE                                               'We have income to update and budget
        EARNED_INCOME_PANELS_ARRAY(update_this_month, ei_panel) = FALSE         'defaulting this for each panel - this will be updated for each month on each panel when the updating actually happens

        'here we look at the initial month for each panel - making a date for the first of the initial month
        mm_1_yy = EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel) & "/1/" & EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel)
        mm_1_yy = DateValue(mm_1_yy)
        If InStr(list_of_all_months_to_update, "~" & mm_1_yy & "~") = 0 Then    'looks to see if the initial month has already been added to the list on a previous loop
            list_of_all_months_to_update = list_of_all_months_to_update & mm_1_yy & "~" 'if not, it is added to the list
        End If

        If EARNED_INCOME_PANELS_ARRAY(update_futue_chkbx, ei_panel) = checked Then  'if it was indicated to add future months we will add all future months to the list as well
            next_month = mm_1_yy        'this is the initial month to start with
            CM_plus_2 = DateValue(CM_plus_2_mo & "/1/" & CM_plus_2_yr)          'setting a date for current month plus 2
            CM_plus_2 = DateValue(CM_plus_2)
            Do          'now we are going to loop to keep adding months
                If InStr(list_of_all_months_to_update, "~" & next_month & "~") = 0 Then     'if the month is NOT in the list, add it
                    list_of_all_months_to_update = list_of_all_months_to_update & next_month & "~"
                End If

                next_month = DateAdd("m", 1, next_month)                        'now increment to the next month
            Loop until next_month = CM_plus_2                                   'stop one the next month is current month plus 2 because AMXIS doesn't go there
        End If
    End If
    Call back_to_SELF
Next        'we do this for each panel so all possible months needed to be updated are there

                            '----------------------------------------------------------'
                    '---------------------------------------------------------------------------------'
    '-------------------------------------------------GOING TO UPDATE THE PANEL --------------------------------------------------'
                    '---------------------------------------------------------------------------------'
                            '----------------------------------------------------------'

If update_with_verifs = TRUE Then       'this means we have at least one panel with income to update'
    list_of_all_months_to_update = right(list_of_all_months_to_update, len(list_of_all_months_to_update)-1)     'formatting our list of months
    list_of_all_months_to_update = left(list_of_all_months_to_update, len(list_of_all_months_to_update)-1)
    If InStr(list_of_all_months_to_update, "~") <> 0 Then                       'making the list an array in chronologival order
        update_months_array = split(list_of_all_months_to_update, "~")
        Call sort_dates(update_months_array)
    Else
        update_months_array = array(list_of_all_months_to_update)
    End If
    script_run_lowdown = script_run_lowdown & vbCr & "The months to update - " & list_of_all_months_to_update

    Call back_to_SELF       'reset

    next_cash_month = 0         'setting this because we have ANOTHER ARRAY!
    For each active_month in update_months_array                'now we loop through the list of months we found before
        updates_to_display = ""                                 'this is for developer mode when in INQUIRY
        MAXIS_footer_month = DatePart("m", active_month)        'setting the footer month and year for each month in the list for NAV
        MAXIS_footer_month = right("00" & MAXIS_footer_month, 2)
        MAXIS_footer_year = DatePart("yyyy", active_month)
        MAXIS_footer_year = right(MAXIS_footer_year, 2)

        RETRO_month = DateAdd("m", -2, active_month)            'defining 2 months ago for CASH/UNCLE HARRY process
        RETRO_footer_month = DatePart("m", RETRO_month)
        RETRO_footer_month = right("00" & RETRO_footer_month, 2)
        RETRO_footer_year = DatePart("yyyy", RETRO_month)
        RETRO_footer_year = right(RETRO_footer_year, 2)

        EMReadScreen summ_check, 4, 2, 46                       'Making sure we start at SUMM
        'BUGGY CODE - need to make sure we are also in the right footer month here
        If summ_check <> "SUMM" Then                            'at the end of the loop we go to summ so we should be already there
            Call back_to_SELF
            Do
                Call navigate_to_MAXIS_screen("STAT", "SUMM")
                EMReadScreen summ_check, 4, 2, 46
            Loop until summ_check = "SUMM"
        End If

        For ei_panel = 0 to UBOUND(EARNED_INCOME_PANELS_ARRAY, 2)               'Now we look at each panel
            updates_to_display = ""
            If EARNED_INCOME_PANELS_ARRAY(income_received, ei_panel) = TRUE Then    'If we got income for this panel
                'BUGGY CODE - this may miss the first check(s) in a month if they are not listed or the pay date is after the first one for the first month
                'ALL THE JUICY BITS GO HERE
                top_of_order = EARNED_INCOME_PANELS_ARRAY(order_ubound, ei_panel)   'defining the position of the last check for this panel

                script_run_lowdown = script_run_lowdown & vbCr & "Updated JOBS - MEMB " & EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel) & " " & EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel) & " -- MONTH - " & MAXIS_footer_month & "/" & MAXIS_footer_year

                this_month_checks_array = ""        'have to plank these out at each panel/month loop
                checks_list = ""

                this_month = active_month           'BUGGY CODE - I need to get rid of the this_month variable and just use the active_month

                'Here we are making a list of all the checks that we expect for the active month - we will then make that an array
                If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "1 - One Time Per Month" Then
                    next_date = DateAdd("d", 1, EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel))
                    If DatePart("d", next_date) = 1 Then
                        first_of_mx_month = MAXIS_footer_month & "/1/" & MAXIS_footer_year
                        first_of_next_month = DateAdd("m", 1, first_of_mx_month)
                        the_day_of_pay = DateAdd("d", -1, first_of_next_month)
                    Else
                        day_of_month = DatePart("d", EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel))

                        the_day_of_pay = MAXIS_footer_month & "/" & day_of_month & "/" & MAXIS_footer_year
                        the_day_of_pay = DateValue(the_day_of_pay)
                    End If
                    If EARNED_INCOME_PANELS_ARRAY(income_end_dt, ei_panel) <> "" Then
                        If DateDiff("d", the_day_of_pay, EARNED_INCOME_PANELS_ARRAY(income_end_dt, ei_panel)) >= 0 Then checks_list = checks_list & "~" & the_day_of_pay
                    Else
                        checks_list = checks_list & "~" & the_day_of_pay
                    End If


                ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "2 - Two Times Per Month" Then
                    checks_in_month = 0
                    For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                        If DatePart("m", LIST_OF_INCOME_ARRAY(pay_date, all_income)) = DatePart("m", this_month) AND DatePart("yyyy", LIST_OF_INCOME_ARRAY(pay_date, all_income)) = DatePart("yyyy", this_month) Then
                            checks_in_month = checks_in_month + 1
                            checks_list = checks_list & "~" & LIST_OF_INCOME_ARRAY(view_pay_date, all_income)
                        End If
                    Next

                    If checks_in_month = 0 Then
                        month_to_use = DatePart("m", this_month)
                        year_to_use = DatePart("yyyy", this_month)

                        checks_list = checks_list & "~" & DateValue(month_to_use & "/" & EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) & "/" & year_to_use)
                        If EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = "LAST" Then
                            first_of_payMonth = month_to_use & "/1/" & year_to_use
                            first_of_nextMonth = DateAdd("m", 1, first_of_payMonth)
                            checks_list = checks_list & "~" & DateAdd("d", -1, first_of_nextMonth)
                        Else
                            checks_list = checks_list & "~" & DateValue(month_to_use & "/" & EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) & "/" & year_to_use)
                        End If

                    ElseIf checks_in_month = 1 Then
                        the_check = replace(checks_list, "~", "")
                        month_to_use = DatePart("m", this_month)
                        year_to_use = DatePart("yyyy", this_month)
                        If EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = "LAST" Then
                            first_of_payMonth = month_to_use & "/1/" & year_to_use
                            first_of_nextMonth = DateAdd("m", 1, first_of_payMonth)
                            If DatePart("d", the_check) = EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) Then
                                checks_list = checks_list & "~" & DateAdd("d", -1, first_of_nextMonth)
                            Else
                                the_other_check = month_to_use & "/" & EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) & "/" & year_to_use
                                checks_list = the_other_check & "~" & the_check
                            End If
                        Else
                            If DatePart("d", the_check) = EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) Then
                                the_other_check = month_to_use & "/" & EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) & "/" & year_to_use
                                checks_list = checks_list & "~" & the_other_check
                            ElseIf DatePart("d", the_check) = EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) Then
                                the_other_check = month_to_use & "/" & EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) & "/" & year_to_use
                                checks_list = the_other_check & "~" & the_check
                            End If
                        End If
                    End If


                ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "3 - Every Other Week" Then
                    the_date = DateValue(EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel))
                    Do
                        If DatePart("m", the_date) = DatePart("m", this_month) AND DatePart("yyyy", the_date) = DatePart("yyyy", this_month) Then
                            If EARNED_INCOME_PANELS_ARRAY(income_end_dt, ei_panel) <> "" Then
                                If DateDiff("d", the_date, EARNED_INCOME_PANELS_ARRAY(income_end_dt, ei_panel)) >= 0 Then checks_list = checks_list & "~" & the_date
                            Else
                                checks_list = checks_list & "~" & the_date
                            End If
                        End If
                        the_date = DateAdd("d", 14, the_date)
                    Loop until right("0" & DatePart("m", the_date), 2) = CM_plus_2_mo AND right(DatePart("yyyy", the_date), 2) = CM_plus_2_yr
                ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "4 - Every Week" Then
                    the_date = DateValue(EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel))
                    Do
                        If DatePart("m", the_date) = DatePart("m", this_month) AND DatePart("yyyy", the_date) = DatePart("yyyy", this_month) Then
                            If EARNED_INCOME_PANELS_ARRAY(income_end_dt, ei_panel) <> "" Then
                                If DateDiff("d", the_date, EARNED_INCOME_PANELS_ARRAY(income_end_dt, ei_panel)) >= 0 Then checks_list = checks_list & "~" & the_date
                            Else
                                checks_list = checks_list & "~" & the_date
                            End If
                        End If
                        the_date = DateAdd("d", 7, the_date)
                    Loop until right("0" & DatePart("m", the_date), 2) = CM_plus_2_mo AND right(DatePart("yyyy", the_date), 2) = CM_plus_2_yr

                ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "5 - Other" Then
                End If

                'formatting the list and maing it an array
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
                    next_date = DateAdd("d", 1, EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel))
                    If DatePart("d", next_date) = 1 Then
                        first_of_mx_month = RETRO_footer_month & "/1/" & RETRO_footer_year
                        first_of_next_month = DateAdd("m", 1, first_of_mx_month)
                        the_day_of_pay = DateAdd("d", -1, first_of_next_month)
                    Else
                        day_of_month = DatePart("d", EARNED_INCOME_PANELS_ARRAY(panel_first_check, ei_panel))

                        the_day_of_pay = RETRO_footer_month & "/" & day_of_month & "/" & RETRO_footer_year
                        the_day_of_pay = DateValue(the_day_of_pay)
                    End If
                    checks_list = checks_list & "~" & the_day_of_pay

                ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "2 - Two Times Per Month" Then
                    checks_in_month = 0
                    For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                        If DatePart("m", LIST_OF_INCOME_ARRAY(pay_date, all_income)) = DatePart("m", RETRO_month) AND DatePart("yyyy", LIST_OF_INCOME_ARRAY(pay_date, all_income)) = DatePart("yyyy", RETRO_month) Then
                            checks_in_month = checks_in_month + 1
                            checks_list = checks_list & "~" & LIST_OF_INCOME_ARRAY(view_pay_date, all_income)
                        End If
                    Next

                    If checks_in_month = 0 Then
                        month_to_use = DatePart("m", RETRO_month)
                        year_to_use = DatePart("yyyy", RETRO_month)

                        checks_list = checks_list & "~" & DateValue(month_to_use & "/" & EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) & "/" & year_to_use)
                        If EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = "LAST" Then
                            first_of_payMonth = month_to_use & "/1/" & year_to_use
                            first_of_nextMonth = DateAdd("m", 1, first_of_payMonth)
                            checks_list = checks_list & "~" & DateAdd("d", -1, first_of_nextMonth)
                        Else
                            checks_list = checks_list & "~" & DateValue(month_to_use & "/" & EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) & "/" & year_to_use)
                        End If

                    ElseIf checks_in_month = 1 Then
                        the_check = replace(checks_list, "~", "")
                        month_to_use = DatePart("m", RETRO_month)
                        year_to_use = DatePart("yyyy", RETRO_month)
                        If EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) = "LAST" Then
                            first_of_payMonth = month_to_use & "/1/" & year_to_use
                            first_of_nextMonth = DateAdd("m", 1, first_of_payMonth)
                            If DatePart("d", the_check) = EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) Then
                                checks_list = checks_list & "~" & DateAdd("d", -1, first_of_nextMonth)
                            Else
                                the_other_check = month_to_use & "/" & EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) & "/" & year_to_use
                                checks_list = the_other_check & "~" & the_check
                            End If
                        Else
                            If DatePart("d", the_check) = EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) Then
                                the_other_check = month_to_use & "/" & EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) & "/" & year_to_use
                                checks_list = checks_list & "~" & the_other_check
                            ElseIf DatePart("d", the_check) = EARNED_INCOME_PANELS_ARRAY(bimonthly_second, ei_panel) Then
                                the_other_check = month_to_use & "/" & EARNED_INCOME_PANELS_ARRAY(bimonthly_first, ei_panel) & "/" & year_to_use
                                checks_list = the_other_check & "~" & the_check
                            End If
                        End If
                    End If

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
                        ' MsgBox "The Date - " & the_date & vbNewLine & "RETRO - " & RETRO_month
                        If DatePart("m", the_date) = DatePart("m", RETRO_month) AND DatePart("yyyy", the_date) = DatePart("yyyy", RETRO_month) Then
                            checks_list = checks_list & "~" & the_date
                        End If
                        the_date = DateAdd("d", 7, the_date)
                    Loop until right("0" & DatePart("m", the_date), 2) = CM_plus_2_mo AND right(DatePart("yyyy", the_date), 2) = CM_plus_2_yr

                ElseIf EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "5 - Other" Then
                End If

                'formatting and making this an array
                If left(checks_list, 1) = "~" Then checks_list = right(checks_list, len(checks_list)-1)
                If InStr(checks_list, "~") <> 0 Then
                    If left(checks_list, 1) = "~" Then checks_list = right(checks_list, len(checks_list)-1)
                    retro_month_checks_array = Split(checks_list,"~")
                Else
                    retro_month_checks_array = Array(checks_list)
                End If

                'if the active_month is this panel's initial month - then we indicate this month needs to be updated for this panel
                If EARNED_INCOME_PANELS_ARRAY(initial_month_mo, ei_panel) = MAXIS_footer_month AND EARNED_INCOME_PANELS_ARRAY(initial_month_yr, ei_panel) = MAXIS_footer_year Then EARNED_INCOME_PANELS_ARRAY(update_this_month, ei_panel) = TRUE

                If EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel) = "JOBS" Then Call Navigate_to_MAXIS_screen("STAT", "JOBS") 'navigate to JOBS for the right member and instance
                EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel), 20, 76
                EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel), 20, 79
                transmit
                EMReadScreen confirm_same_employer, 30, 7, 42                  'double check the employer name because we don't want to have wrong income on the wrong panel and from month to month the instances may change
                the_new_instance = ""       'blanking this out
                If confirm_same_employer <> UCase(EARNED_INCOME_PANELS_ARRAY(employer_with_underscores, ei_panel)) Then      'if the name on the panel does not match the name in EARNED_INCOME_PANELS_ARRAY we have to figure this out
                    'BUGGY CODE - this might be causing issues as there were a few reports but I cannot get it to confirm
                    EMWriteScreen "JOBS", 20, 71        'go back to the first job for this person
                    EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel), 20, 76
                    transmit
                    try = 1         'we need an exit from the loop
                    Do
                        EMReadScreen confirm_same_employer, 30, 7, 42      'now we read this on each panel
                        If confirm_same_employer = UCase(EARNED_INCOME_PANELS_ARRAY(employer_with_underscores, ei_panel)) Then               'if the panel has the employer name, then we set the new instance to EARNED_INCOME_PANELS_ARRAY
                            EMReadScreen the_new_instance, 1, 2, 73
                            EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel) = "0" & the_new_instance
                            Exit Do
                        End If
                        transmit                                'otherwise go to the next panel and repeat
                        EMReadScreen last_jobs, 7, 24, 2
                        try = try + 1
                        If try = 15 Then Exit Do
                    Loop until last_jobs = "ENTER A"            'This is when you can't transmit any more

                    If the_new_instance = "" Then               'If they didn't matcj and we did not find it, this alerts the worker
                        EARNED_INCOME_PANELS_ARRAY(update_this_month, ei_panel) = FALSE     'setting this to NOT update
                        MsgBox "The panel for " & EARNED_INCOME_PANELS_ARRAY(employer, ei_panel) & " could not be found in the month " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". It may have been deleted. The script will not attempt to update this or any future month for this panel."
                    End If
                End If
                If EARNED_INCOME_PANELS_ARRAY(update_this_month, ei_panel) = TRUE Then              'if this panel should be update in thie month - here is where we do it
                    EARNED_INCOME_PANELS_ARRAY(months_updated, ei_panel) = EARNED_INCOME_PANELS_ARRAY(months_updated, ei_panel) & ", " & MAXIS_footer_month & "/" & MAXIS_footer_year       'keeping a list of all the panels updated for each job

                    If EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel) = "JOBS" Then

                        Call Navigate_to_MAXIS_screen("STAT", "JOBS")           'make sure we are at the right place
                        EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel), 20, 76
                        EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel), 20, 79
                        transmit
                        If developer_mode = FALSE Then PF9                      'if we are in INQUIRY, the panel is NOT put in edit mode - otherwise here is where it is put in edit mode
                        'All of these updates_to_display items are copying what the script is doing in the panel and if in INQUIRY this will be displayed for each job for each month
                        updates_to_display = "JOBS Update for " & MAXIS_footer_month & "/" & MAXIS_footer_year & vbNewLine & "MEMBER " & EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel) & " - Employer: " & EARNED_INCOME_PANELS_ARRAY(employer, ei_panel)

                        EMWriteScreen left(EARNED_INCOME_PANELS_ARRAY(income_type, ei_panel), 1), 5, 34         'income type, verif and hhourly wage are the samy for each progam
                        EMWriteScreen left(EARNED_INCOME_PANELS_ARRAY(income_verif, ei_panel), 1), 6, 34
                        EMWriteScreen "      ", 6, 75       'this blanks out the wage otherwise there is carryover and 10.00 becomes 100.00 - which is not correct
                        EMWriteScreen EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel), 6, 75
                        EMWriteScreen EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel), 18, 35
                        updates_to_display = updates_to_display & vbNewLine & "Income type: " & EARNED_INCOME_PANELS_ARRAY(income_type, ei_panel) & " - Verification: " & EARNED_INCOME_PANELS_ARRAY(income_verif, ei_panel) & vbNewLine & "Hourly wage: $" & EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel) & "/hr. Pay Frequency: " & EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel)

                        If IsDate(EARNED_INCOME_PANELS_ARRAY(income_start_dt, ei_panel)) = TRUE Then
                            job_start_month = DatePart("m", EARNED_INCOME_PANELS_ARRAY(income_start_dt, ei_panel))      'entering the start date - BUGGY CODE - this isn't buggy per se but I should change this to the function for this
                            job_start_day = DatePart("d", EARNED_INCOME_PANELS_ARRAY(income_start_dt, ei_panel))
                            job_start_year = DatePart("yyyy", EARNED_INCOME_PANELS_ARRAY(income_start_dt, ei_panel))
                            job_start_month = right("00" & job_start_month, 2)
                            job_start_day = right("00" & job_start_day, 2)
                            job_start_year = right(job_start_year, 2)
                            updates_to_display = updates_to_display & vbNewLine & "Income Start Date: " & EARNED_INCOME_PANELS_ARRAY(income_start_dt, ei_panel)

                            EMWriteScreen job_start_month, 9, 35
                            EMWriteScreen job_start_day, 9, 38
                            EMWriteScreen job_start_year, 9, 41
                        End If

                        'Here we update the panel with process and information that is specific to the program the income applies to
                        'The order is important here as some will take precedent. Currenly the order iw SNAP-GRH-HC-Cash
                        'Since SNAP and GRH budgets are determined by the PIC, the information on the main JOBS panel is less important.
                        'Cash has the most specific update requirements for JOBS, so it is last to ensure those are the changes to JOBS that are saved
                        If EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) = checked THen           'for income that applies to SNAP
                            STATS_manualtime = STATS_manualtime + 165
                            updates_to_display = updates_to_display & vbNewLine & vbNewLine & "** SNAP Budget Update **" & vbNewLine & "---- PIC ----"
                            EMWriteScreen "X", 19, 38       'opening the PIC
                            transmit
                            PF20

                            list_row = 9                    'here we clear the PIC of all previous data
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

                            Call create_MAXIS_friendly_date(date, 0, 5, 34)                     'enter the current date in date of calculation field
                            EMWriteScreen left(EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel), 1), 5, 64        'enter the pay frequency code only in the correct field

                            'If we are in the footer month that is the month of application or first month of income for a new job
                            'there is special functionality to LUMP the income in the PIC so that an accurate amount can be entered NEED POLICY REFERENCE
                            If (fs_appl_footer_month = MAXIS_footer_month AND fs_appl_footer_year = MAXIS_footer_year) OR (job_start_month = MAXIS_footer_month AND job_start_year = MAXIS_footer_year) Then
                                updates_to_display = updates_to_display & vbNewLine & "Date of Calculation: " & date & "  Pay Frequency: 1 - Monthly"

                                fs_appl_date = DateValue(fs_appl_date)      'setting some defaults - more for-nexts are coming
                                appl_month_gross = 0
                                appl_month_hours = 0

                                checks_lumped = ""
                                expected_pay_lumped = ""

                                month_lumped = MAXIS_footer_month & "/" & MAXIS_footer_year     'formatting for readability
                                If fs_appl_footer_month = MAXIS_footer_month AND fs_appl_footer_year = MAXIS_footer_year Then
                                    reason_lumped = "month of application"
                                ElseIf job_start_month = MAXIS_footer_month AND job_start_year = MAXIS_footer_year Then
                                    reason_lumped = "first month of new job"
                                Else
                                    reason_lumped = "month of application and first month of new job"
                                End If

                                For each this_date in this_month_checks_array           'this array was set at the begining of this month's loop - it will get us all our pay dates
                                    If DateDiff("d", this_date, EARNED_INCOME_PANELS_ARRAY(income_start_dt, ei_panel)) < 1 Then     'if the pay date we are looking at is on or after the income start date we will add it in to the lump

                                        date_found = FALSE      'default for before we look at the checks
                                        For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                            If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(pay_date, all_income) <> "" Then
                                                If DateDiff("d", LIST_OF_INCOME_ARRAY(pay_date, all_income), this_date) = 0 Then            'if the pay date (this is NOT the view_pay_date) then we use the amount provided on ENTER PAY
                                                    date_found = TRUE

                                                    appl_month_gross = appl_month_gross + LIST_OF_INCOME_ARRAY(gross_amount, all_income)
                                                    appl_month_hours = appl_month_hours + LIST_OF_INCOME_ARRAY(hours, all_income)
                                                    checks_lumped = checks_lumped & LIST_OF_INCOME_ARRAY(view_pay_date, all_income) & " - $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & "hrs.; "      'saving a list of the checks used
                                                End If

                                            End If
                                        Next

                                        If date_found = FALSE Then          'if the date was not provided on ENTER PAY, we are going to use an averate
                                            appl_month_gross = appl_month_gross + EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel)
                                            appl_month_hours = appl_month_hours + EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel)

                                            expected_pay_lumped = expected_pay_lumped & this_date & " - anticipated $" &EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel) & " - " & EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel) & "hrs.; "        'saving a list of the pay estimates used
                                        End If
                                    End If
                                Next

                                EMWriteScreen "1", 5, 64                        'A pay frequency of 1 (monthly) is entered on the PIC
                                EMWriteScreen MAXIS_footer_month, 9, 13         'a default date of the first of the month is entered on the PIC
                                EMWriteScreen "01", 9, 16
                                EMWriteScreen MAXIS_footer_year, 9, 19
                                appl_month_gross = FormatNumber(appl_month_gross, 2, -1, 0, 0)      'the sum of all the pay check gross amounts (or average) is formated and entered
                                EMWriteScreen appl_month_gross, list_row, 25
                                EMWriteScreen appl_month_hours, list_row, 35

                                updates_to_display = updates_to_display & vbNewLine & "Actual Pay: Date - " & MAXIS_footer_month & "/01/" & MAXIS_footer_year & " - $" & appl_month_gross & " - " & appl_month_hours & " hrs." & vbNewLine

                                If checks_lumped <> "" Then     'formatting the lists of the checks that we included for the CNote
                                    If right(checks_lumped, 2) = "; " Then checks_lumped = left(checks_lumped, len(checks_lumped)-2)
                                End If
                                If expected_pay_lumped <> "" Then
                                    If right(expected_pay_lumped, 1) = "; " Then expected_pay_lumped = left(expected_pay_lumped, len(expected_pay_lumped)-2)
                                End If

                                EARNED_INCOME_PANELS_ARRAY(income_lumped_mo, ei_panel)  = month_lumped              'saving all the information about the lumping to the array for CNoting
                                EARNED_INCOME_PANELS_ARRAY(lump_reason, ei_panel)       = reason_lumped
                                EARNED_INCOME_PANELS_ARRAY(act_checks_lumped, ei_panel) = checks_lumped
                                EARNED_INCOME_PANELS_ARRAY(est_checks_lumped, ei_panel) = expected_pay_lumped
                                EARNED_INCOME_PANELS_ARRAY(lump_gross, ei_panel)        = appl_month_gross
                                EARNED_INCOME_PANELS_ARRAY(lump_hrs, ei_panel)          = appl_month_hours

                            Else        'this is if we are in any month other than the month of application or first month of income for this job
                                updates_to_display = updates_to_display & vbNewLine & "Date of Calculation: " & date & "  Pay Frequency: " & EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel)
                                If EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_estimate Then           'if use estimate was selected - then we just plug in the hours/wk and pay/hr
                                    updates_to_display = updates_to_display & vbNewLine & "Estimate: Hourly Wage - $" & EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel) & "/hr - " & EARNED_INCOME_PANELS_ARRAY(snap_hrs_per_wk, ei_panel) & " hrs/wk" & vbNewLine
                                    EMWriteScreen EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel), 9, 66
                                    EMWriteScreen EARNED_INCOME_PANELS_ARRAY(snap_hrs_per_wk, ei_panel), 8, 64
                                End If
                                If EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_actual Then             'if we have to use actual - it is harder - HERE WE GO!
                                    updates_to_display = updates_to_display & vbNewLine & "Actual Pay: Date - "

                                    'here we do not need to compare it to expected checks because the PIC dates do not need to align with the current month
                                    list_row = 9
                                    For order_number = 1 to top_of_order                        'loop through the order number lowest to highest
                                        For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                            'conditional if it is the right panel AND the order matches - then do the thing you need to do
                                            If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(check_order, all_income) = order_number Then
                                                If LIST_OF_INCOME_ARRAY(budget_in_SNAP_no, all_income) = unchecked Then
                                                    Call create_MAXIS_friendly_date(LIST_OF_INCOME_ARRAY(view_pay_date, all_income), 0, list_row, 13)           'this is the CIEW date - the one the owrker actually entered
                                                    net_amount = LIST_OF_INCOME_ARRAY(gross_amount, all_income) - LIST_OF_INCOME_ARRAY(exclude_amount, all_income)      'taking out excluded amounts
                                                    net_amount = FormatNumber(net_amount, 2, -1, 0, 0)
                                                    EMWriteScreen net_amount, list_row, 25      'entering the pay amount to count
                                                    EMWriteScreen LIST_OF_INCOME_ARRAY(hours, all_income), list_row, 35     'enting the hours

                                                    updates_to_display = updates_to_display & LIST_OF_INCOME_ARRAY(view_pay_date, all_income) & " - $" & net_amount & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & " hrs." & vbNewLine

                                                    list_row = list_row + 1         'next line of the PIC'
                                                    If list_row = 14 Then
                                                        PF20
                                                        list_row = 9
                                                    End If
                                                End If
                                            End If
                                        next
                                    next

                                End If
                            End If
                            transmit            'saving the PIC
                            transmit
                            PF3

                            'now we are on the main JOBS panel
                            jobs_row = 12               'first we blank out what is already there
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

                            jobs_row = 12               'now we are going to enter the new information
                            total_hours = 0

                            updates_to_display = updates_to_display & "-- Main JOBS panel --"

                            For each this_date in this_month_checks_array       'now using the list we made of all the checks for THIS month
                                If DateDiff("d", this_date, EARNED_INCOME_PANELS_ARRAY(income_start_dt, ei_panel)) < 1 Then     'checking to make sure the paydate is not before the income start date - that causes a red line

                                    date_found = FALSE          'default for each loop
                                    For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                        'conditional if it is the right panel AND the order matches - then do the thing you need to do
                                        If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(pay_date, all_income) <> "" Then
                                            If DateDiff("d", LIST_OF_INCOME_ARRAY(pay_date, all_income), this_date) = 0 Then            'if the pay day matches the date in the list, we add the information to the panel
                                                date_found = TRUE               'saving this so that the information is not over written
                                                Call create_MAXIS_friendly_date(LIST_OF_INCOME_ARRAY(view_pay_date, all_income), 0, jobs_row, 54)       'pay date
                                                LIST_OF_INCOME_ARRAY(gross_amount, all_income) = FormatNumber(LIST_OF_INCOME_ARRAY(gross_amount, all_income), 2, -1, 0, 0)
                                                EMWriteScreen LIST_OF_INCOME_ARRAY(gross_amount, all_income), jobs_row, 67              'gross pay - not using the excluded amount on the main panel
                                                total_hours = total_hours + LIST_OF_INCOME_ARRAY(hours, all_income)             'running total of this
                                                updates_to_display = updates_to_display & vbNewLine & "Date - " & LIST_OF_INCOME_ARRAY(view_pay_date, all_income) & " - $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income)
                                            End If
                                        End If
                                    Next

                                    If date_found = FALSE Then                  'if the date was not found in the LIST_OF_INCOME_ARRAY - we will use an average
                                        Call create_MAXIS_friendly_date(this_date, 0, jobs_row, 54)         'entering the date
                                        EARNED_INCOME_PANELS_ARRAY(snap_ave_inc_per_pay, ei_panel) = FormatNumber(EARNED_INCOME_PANELS_ARRAY(snap_ave_inc_per_pay, ei_panel), 2, -1, 0, 0)
                                        EMWriteScreen EARNED_INCOME_PANELS_ARRAY(snap_ave_inc_per_pay, ei_panel), jobs_row, 67          'entering the average
                                        total_hours = total_hours + EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel)               'totalling hours
                                        updates_to_display = updates_to_display & vbNewLine & "Date - " & this_date & " - $" & EARNED_INCOME_PANELS_ARRAY(snap_ave_inc_per_pay, ei_panel)

                                    End If
                                    jobs_row = jobs_row + 1         'moving to the next row
                                End If
                            Next
                            total_hours = Round(total_hours)        'formatting the hours total
                            EMWriteScreen "   ", 18, 72             'blanking out BOTH hours positions
                            EMWriteScreen "   ", 18, 43
                            EMWriteScreen total_hours, 18, 72       'entering the hours on the panel
                            updates_to_display = updates_to_display & vbNewLine & "        Total Hours: " & total_hours
                        End If          'If EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) = checked THen'

                        If EARNED_INCOME_PANELS_ARRAY(apply_to_GRH, ei_panel) = checked Then            'now for GRH
                            STATS_manualtime = STATS_manualtime + 145
                            updates_to_display = updates_to_display & vbNewLine & vbNewLine & "*** GRH Budget Update ***" & vbNewLine & "---- PIC ----"
                            EMWriteScreen "X", 19, 71               'opening the GRH PIC
                            transmit

                            list_row = 7                            'blanking out all previous information on the GRH PIC
                            beg_of_list_check = ""
                            Do
                                EMWriteScreen "  ", list_row, 9
                                EMWriteScreen "  ", list_row, 12
                                EMWriteScreen "  ", list_row, 15
                                EMWriteScreen "        ", list_row, 21
                                EMWriteScreen "      ", list_row, 35
                                list_row = list_row + 1
                            Loop until list_row = 17

                            EMWriteScreen "      ", 6, 63
                            EMWriteScreen "        ", 7, 65
                            EMWriteScreen "        ", 11, 65

                            Call create_MAXIS_friendly_date(date, 0, 3, 30)     'entering today's date for date of calculation and the pay frequency
                            EMWriteScreen left(EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel), 1), 3, 63
                            updates_to_display = updates_to_display & vbNewLine & "Date of Calculation: " & date & "  Pay Frequency: " & EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel)

                            If EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_estimate Then           'if use estimate was shosen, just need the hourly wage and pay per hour
                                updates_to_display = updates_to_display & vbNewLine & "Estimate: Hourly Wage - $" & EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel) & "/hr - " & EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) & " hrs/wk" & vbNewLine
                                EMWriteScreen EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel), 7, 65
                                EMWriteScreen EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel), 6, 63
                            End If
                            If EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_actual Then             'if we are using actual, we need to put some real amounts in here
                                updates_to_display = updates_to_display & vbNewLine & "Actual Pay: Date - "

                                list_row = 9                'entering all checks on the PIC that were provided - no exclusions or reductions
                                For order_number = 1 to top_of_order                        'loop through the order number lowest to highest
                                    For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                        'conditional if it is the right panel AND the order matches - then do the thing you need to do
                                        If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(check_order, all_income) = order_number Then
                                            Call create_MAXIS_friendly_date(LIST_OF_INCOME_ARRAY(view_pay_date, all_income), 0, list_row, 9)
                                            LIST_OF_INCOME_ARRAY(gross_amount, all_income) = FormatNumber(LIST_OF_INCOME_ARRAY(gross_amount, all_income), 2, -1, 0, 0)
                                            EMWriteScreen LIST_OF_INCOME_ARRAY(gross_amount, all_income), list_row, 21

                                            updates_to_display = updates_to_display & LIST_OF_INCOME_ARRAY(view_pay_date, all_income) & " - $" & net_amount & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & " hrs." & vbNewLine

                                            list_row = list_row + 1
                                        End If
                                    next
                                next
                            End If

                            transmit            'saving the PIC
                            transmit
                            PF3

                            updates_to_display = updates_to_display & "-- Main JOBS panel --"

                            For jobs_row = 12 to 16                         'blanking out the main JOBS panel
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

                            jobs_row = 12           'setting for the next loop
                            total_hours = 0

                            For each this_date in this_month_checks_array           'entering each date for this month
                                date_found = FALSE
                                For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                    'conditional if it is the right panel AND the order matches - then do the thing you need to do
                                    If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(pay_date, all_income) <> "" Then
                                        If DateDiff("d", LIST_OF_INCOME_ARRAY(pay_date, all_income), this_date) = 0 Then            'if the pay date in the array matches the one we are looking at
                                            date_found = TRUE
                                            Call create_MAXIS_friendly_date(LIST_OF_INCOME_ARRAY(view_pay_date, all_income), 0, jobs_row, 54)           'write the date
                                            LIST_OF_INCOME_ARRAY(gross_amount, all_income) = FormatNumber(LIST_OF_INCOME_ARRAY(gross_amount, all_income), 2, -1, 0, 0)
                                            EMWriteScreen LIST_OF_INCOME_ARRAY(gross_amount, all_income), jobs_row, 67              'write the amount
                                            total_hours = total_hours + LIST_OF_INCOME_ARRAY(hours, all_income)                     'running total of the hours
                                            updates_to_display = updates_to_display & vbNewLine & "Date - " & LIST_OF_INCOME_ARRAY(view_pay_date, all_income) & " - $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income)
                                        End If
                                    End If
                                Next

                                If date_found = FALSE Then                  'if the date was not found we will enter an average for the pay amount
                                    Call create_MAXIS_friendly_date(this_date, 0, jobs_row, 54)             'write the date
                                    EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel) = FormatNumber(EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel), 2, -1, 0, 0)
                                    EMWriteScreen EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel), jobs_row, 67           'write the average check amount
                                    total_hours = total_hours + EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel)           'running total of hours - using average hours per check
                                    updates_to_display = updates_to_display & vbNewLine & "Date - " & this_date & " - $" & EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel)
                                End If
                                jobs_row = jobs_row + 1         'go to the next row
                            Next
                            total_hours = Round(total_hours)        'formatting and writing the cumulative hours
                            EMWriteScreen "   ", 18, 72
                            EMWriteScreen "   ", 18, 43
                            EMWriteScreen total_hours, 18, 72
                            updates_to_display = updates_to_display & vbNewLine & "        Total Hours: " & total_hours
                        End If              'If EARNED_INCOME_PANELS_ARRAY(apply_to_GRH, ei_panel) = checked Then

                        If EARNED_INCOME_PANELS_ARRAY(apply_to_HC, ei_panel) = checked Then         'now on to the health care
                            STATS_manualtime = STATS_manualtime + 140
                            updates_to_display = updates_to_display & vbNewLine & vbNewLine & "*** HC Budget Update ***"

                            For jobs_row = 12 to 16                             'blanking out the income information from the main JOBS panel'
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

                            jobs_row = 12           'setting these for this loop
                            total_hours = 0

                            For each this_date in this_month_checks_array       'looking at each check in the list of expected pay dates
                                date_found = FALSE
                                For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                    'conditional if it is the right panel AND the order matches - then do the thing you need to do
                                    If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(pay_date, all_income) <> "" Then
                                        If DateDiff("d", LIST_OF_INCOME_ARRAY(pay_date, all_income), this_date) = 0 Then        'if the date from LIST_OF_INCOME_ARRAY matches
                                            date_found = TRUE
                                            Call create_MAXIS_friendly_date(LIST_OF_INCOME_ARRAY(view_pay_date, all_income), 0, jobs_row, 54)       'enter the pay date
                                            LIST_OF_INCOME_ARRAY(gross_amount, all_income) = FormatNumber(LIST_OF_INCOME_ARRAY(gross_amount, all_income), 2, -1, 0, 0)
                                            EMWriteScreen LIST_OF_INCOME_ARRAY(gross_amount, all_income), jobs_row, 67      'enter the gross amount
                                            total_hours = total_hours + LIST_OF_INCOME_ARRAY(hours, all_income)             'total the hours
                                            updates_to_display = updates_to_display & vbNewLine & "Date - " & LIST_OF_INCOME_ARRAY(view_pay_date, all_income) & " - $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income)
                                        End If
                                    End If
                                Next

                                If date_found = FALSE Then                      'if the date was not found - we use averages
                                    Call create_MAXIS_friendly_date(this_date, 0, jobs_row, 54)         'enter the pay date
                                    EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel) = FormatNumber(EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel), 2, -1, 0, 0)
                                    EMWriteScreen EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel), jobs_row, 67       'enter the average per pay
                                    total_hours = total_hours + EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel)       'running total of hours
                                    updates_to_display = updates_to_display & vbNewLine & "Date - " & this_date & " - $" & EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel)
                                End If
                                jobs_row = jobs_row + 1
                            Next
                            total_hours = Round(total_hours)            'entering the cumulative hours
                            EMWriteScreen "   ", 18, 72
                            EMWriteScreen "   ", 18, 43
                            EMWriteScreen total_hours, 18, 72
                            updates_to_display = updates_to_display & vbNewLine & "        Total Hours: " & total_hours

                            If MAXIS_footer_month = CM_plus_1_mo AND MAXIS_footer_year = CM_plus_1_yr Then          'if we are in current month plus one, we need to update the HC Income Estimate Pop-up
                                EMWriteScreen "X", 19, 48           'opening the HC Inc Est
                                transmit

                                EMWriteScreen "        ", 11, 63        'blanking it out
                                EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel) = FormatNumber(EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel), 2, -1, 0, 0)
                                EMWriteScreen EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel), 11, 63         'writing it in
                                transmit        'saving the information
                                transmit
                            End If
                        End If      'If EARNED_INCOME_PANELS_ARRAY(apply_to_HC, ei_panel) = checked Then

                        If EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, ei_panel) = checked OR UH_SNAP = TRUE OR EARNED_INCOME_PANELS_ARRAY(hc_retro, ei_panel) = TRUE Then         'now on to cash
                            STATS_manualtime = STATS_manualtime + 185
                            updates_to_display = updates_to_display & vbNewLine & vbNewLine & "*** Cash Budget Update (or Uncle Harry SNAP) ***"

                            'for each month that is updated for cash, we need to track more detail since there is retro and prosp information to be concerned with
                            'this array stores that detail
                            ReDim Preserve CASH_MONTHS_ARRAY(8, next_cash_month)                'resizing the array
                            CASH_MONTHS_ARRAY(retro_updtd, next_cash_month) = FALSE             'defaults
                            CASH_MONTHS_ARRAY(prosp_updtd, next_cash_month) = FALSE
                            CASH_MONTHS_ARRAY(panel_indct, next_cash_month) = ei_panel          'need to connect it to a panel
                            CASH_MONTHS_ARRAY(cash_mo_yr, next_cash_month) = MAXIS_footer_month & "/" & MAXIS_footer_year       'save which month we are looking at
                            CASH_MONTHS_ARRAY(retro_mo_yr, next_cash_month) = RETRO_footer_month & "/" & RETRO_footer_year

                            'RETROSPECTIVE SIDE' first we worry about what is on the retro side - this may be blank in the end
                            For jobs_row = 12 to 16                 'blanking out the retro side of the panel
                                EMWriteScreen "  ", jobs_row, 25
                                EMWriteScreen "  ", jobs_row, 28
                                EMWriteScreen "  ", jobs_row, 31
                                EMWriteScreen "        ", jobs_row, 38
                            Next

                            jobs_row = 12           'set for the loop
                            total_hours = 0
                            total_pay = 0
                            count_checks = 0
                            updates_to_display = updates_to_display & vbNewLine & "--- RETRO ---"
                            For each this_date in retro_month_checks_array          'there is a seperate list of retro pay dates
                                If this_date <> "" Then
                                    date_found = FALSE      'default for the start of each loop
                                    For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                        'conditional if it is the right panel AND the order matches - then do the thing you need to do
                                        If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(pay_date, all_income) <> "" Then
                                            If DateDiff("d", LIST_OF_INCOME_ARRAY(pay_date, all_income), this_date) = 0 Then        'if the date was in the array, we use it
                                                date_found = TRUE
                                                CASH_MONTHS_ARRAY(retro_updtd, next_cash_month) = TRUE      'setting this to know there was retro information added
                                                Call create_MAXIS_friendly_date(LIST_OF_INCOME_ARRAY(view_pay_date, all_income), 0, jobs_row, 25)
                                                LIST_OF_INCOME_ARRAY(gross_amount, all_income) = FormatNumber(LIST_OF_INCOME_ARRAY(gross_amount, all_income), 2, -1, 0, 0)
                                                EMWriteScreen LIST_OF_INCOME_ARRAY(gross_amount, all_income), jobs_row, 38      'entering the pay information
                                                total_hours = total_hours + LIST_OF_INCOME_ARRAY(hours, all_income)             'running total of hours
                                                total_pay = total_pay + LIST_OF_INCOME_ARRAY(gross_amount, all_income)          'running total of pay for RETRO month only
                                                count_checks = count_checks + 1                 'need to track the number of checks that we used
                                                updates_to_display = updates_to_display & vbNewLine & "Date - " & LIST_OF_INCOME_ARRAY(view_pay_date, all_income) & " - $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income)
                                            End If
                                        End If
                                    Next
                                    'there is no functionality for if pay was not found as we don't use averages on the retro side

                                    jobs_row = jobs_row + 1
                                End If
                            Next
                            total_hours = Round(total_hours)        'entering retro hours on the retro side
                            EMWriteScreen "   ", 18, 43

                            If count_checks <> 0 Then           'if there we check found we make an average of pay and hours for the RETRO side only
                                EMWriteScreen total_hours, 18, 43
                                this_month_ave_pay = total_pay/count_checks
                                this_month_ave_pay = FormatNumber(this_month_ave_pay, 2,,0)
                                this_month_ave_hours = total_hours/count_checks

                                CASH_MONTHS_ARRAY(mo_retro_pay, next_cash_month) = FormatNumber(total_pay, 2,,0)        'save the total hours and pay to the array for CNote
                                CASH_MONTHS_ARRAY(mo_retro_hrs, next_cash_month) = total_hours
                                updates_to_display = updates_to_display & vbNewLine & "        Total Hours: " & total_hours
                            End If

                            'PROSPECTIVE SIDE'
                            For jobs_row = 12 to 16                 'blanking out the prospective side
                                EMWriteScreen "  ", jobs_row, 54
                                EMWriteScreen "  ", jobs_row, 57
                                EMWriteScreen "  ", jobs_row, 60
                                EMWriteScreen "        ", jobs_row, 67
                            Next

                            jobs_row = 12           'resetting this for the next loop
                            total_hours = 0
                            CASH_MONTHS_ARRAY(prosp_updtd, next_cash_month) = TRUE      'we will actually always update the prospective side
                            CASH_MONTHS_ARRAY(mo_prosp_pay, next_cash_month) = 0
                            updates_to_display = updates_to_display & vbNewLine & "--- PROSP ---"

                            For each this_date in this_month_checks_array           'now we are looking at the same list of checks all the other programs used'
                                date_found = FALSE
                                For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                    'conditional if it is the right panel AND the order matches - then do the thing you need to do
                                    If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(pay_date, all_income) <> "" Then
                                        If DateDiff("d", LIST_OF_INCOME_ARRAY(pay_date, all_income), this_date) = 0 Then        'if the date is found in the LIST_OF_INCOME_ARRAY'
                                            date_found = TRUE
                                            Call create_MAXIS_friendly_date(LIST_OF_INCOME_ARRAY(view_pay_date, all_income), 0, jobs_row, 54)       'enter the date
                                            LIST_OF_INCOME_ARRAY(gross_amount, all_income) = FormatNumber(LIST_OF_INCOME_ARRAY(gross_amount, all_income), 2, -1, 0, 0)
                                            EMWriteScreen LIST_OF_INCOME_ARRAY(gross_amount, all_income), jobs_row, 67          'enter the pay
                                            total_hours = total_hours + LIST_OF_INCOME_ARRAY(hours, all_income)                 'running total of the hours
                                            CASH_MONTHS_ARRAY(mo_prosp_pay, next_cash_month) = CASH_MONTHS_ARRAY(mo_prosp_pay, next_cash_month) + LIST_OF_INCOME_ARRAY(gross_amount, all_income)        'saving information for the array
                                            updates_to_display = updates_to_display & vbNewLine & "Date - " & LIST_OF_INCOME_ARRAY(view_pay_date, all_income) & " - $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income)
                                        End If
                                    End If
                                Next

                                If date_found = FALSE Then              'if there was no date in LIST_OF_INCOME_ARRAY on the prospective side we will enter an average
                                    Call create_MAXIS_friendly_date(this_date, 0, jobs_row, 54)     'enter the pay date
                                    If count_checks <> 0 Then           'if there were retro checks for this month, we enter the average of the retro checks only
                                        this_month_ave_pay = FormatNumber(this_month_ave_pay, 2, -1, 0, 0)
                                        EMWriteScreen this_month_ave_pay, jobs_row, 67          'enter the pay
                                        total_hours = total_hours + this_month_ave_hours        'total the hours
                                        CASH_MONTHS_ARRAY(mo_prosp_pay, next_cash_month) = CASH_MONTHS_ARRAY(mo_prosp_pay, next_cash_month) + this_month_ave_pay        'save information to the array
                                        updates_to_display = updates_to_display & vbNewLine & "Date - " & this_date & " - $" & this_month_ave_pay
                                    Else                'if there was no retro checks provided, we use the average of all checks provided (or determined by pay rate and hours)
                                        EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel) = FormatNumber(EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel), 2, -1, 0, 0)
                                        EMWriteScreen EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel), jobs_row, 67       'enter the average pay
                                        total_hours = total_hours + EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel)       'running total of hours
                                        CASH_MONTHS_ARRAY(mo_prosp_pay, next_cash_month) = CASH_MONTHS_ARRAY(mo_prosp_pay, next_cash_month) + EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel)     'saving information to hours
                                        updates_to_display = updates_to_display & vbNewLine & "Date - " & this_date & " - $" & EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel)
                                    End If
                                End If
                                jobs_row = jobs_row + 1     'next line on the JOBS panel
                            Next
                            total_hours = Round(total_hours)        'entering the total hours on the panel
                            EMWriteScreen "   ", 18, 72
                            EMWriteScreen total_hours, 18, 72
                            updates_to_display = updates_to_display & vbNewLine & "        Total Hours: " & total_hours
                            CASH_MONTHS_ARRAY(mo_prosp_hrs, next_cash_month) = total_hours          'saving the information to the array
                            CASH_MONTHS_ARRAY(mo_prosp_pay, next_cash_month) = FormatNumber(CASH_MONTHS_ARRAY(mo_prosp_pay, next_cash_month), 2,,0)

                            next_cash_month = next_cash_month + 1       'this is for incrementing the array for the next loop
                        End If          'If EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, ei_panel) = checked OR UH_SNAP = TRUE Then
                    End If          'If EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel) = "JOBS" Then
                End If          'If confirm_same_employer <> UCase(EARNED_INCOME_PANELS_ARRAY(employer, ei_panel)) Then
                If updates_to_display <> "" AND developer_mode = TRUE Then MsgBox updates_to_display            'this shows the information that WOULD have been updated if we were not in INQUIRY
                script_run_lowdown = script_run_lowdown & vbCr & updates_to_display
                'If this panel is should to update months after the initial month, this is saved for the next loop to have it updated
                'FUTURE FUNCTIONALITY - if we need to change how we handle the future month updates thing or dealing with STWK - this would be here
                If EARNED_INCOME_PANELS_ARRAY(update_futue_chkbx, ei_panel) = unchecked Then EARNED_INCOME_PANELS_ARRAY(update_this_month, ei_panel) = FALSE
            End If          'If EARNED_INCOME_PANELS_ARRAY(income_received, ei_panel) = TRUE Then
        Next            'For ei_panel = 0 to UBOUND(EARNED_INCOME_PANELS_ARRAY, 2)

        transmit        'after all of the panels have been reviewed we are going to STAT/WRAP to get to the next month without sending through background if possible

        EmWriteScreen "BGTX", 20, 71
        transmit
        If active_month <> update_months_array(ubound(update_months_array)) Then EmWriteScreen "Y", 16, 54      'if we are at the last month to update, then we leave
        transmit
        'BUGGY CODE - need to add a check here that we made it to STAT/WRAP and that we are in the next footer month

        If active_month <> update_months_array(ubound(update_months_array)) Then        'getting to SUMM
            EmWriteScreen "SUMM", 20, 71
            transmit
        End If
    Next            'For each active_month in update_months_array
End If          'If update_with_verifs = TRUE Then

                '----------------------------------------------------------'
        '---------------------------------------------------------------------------------'
'-------------------------------------------------CASE NOTING --------------------------------------------------'
        '---------------------------------------------------------------------------------'
                '----------------------------------------------------------'
end_msg = ""
For ei_panel = 0 to UBOUND(EARNED_INCOME_PANELS_ARRAY, 2)       'each panel will have it's own note
    prog_list = ""              'blanking out for each loop
    updates_to_display = ""

    If EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) = checked Then prog_list = prog_list & "/SNAP"       'setting the header programs
    If EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, ei_panel) = checked Then prog_list = prog_list & "/CASH"
    If EARNED_INCOME_PANELS_ARRAY(apply_to_HC, ei_panel) = checked Then prog_list = prog_list & "/HC"
    If EARNED_INCOME_PANELS_ARRAY(apply_to_GRH, ei_panel) = checked Then prog_list = prog_list & "/GRH"

    If left(prog_list, 1) = "/" Then prog_list = right(prog_list, len(prog_list)-1)

    top_of_order = EARNED_INCOME_PANELS_ARRAY(order_ubound, ei_panel)       'setting the variable for the end of the chronological list of checks

    Select Case EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel)        'FUTURE FUNCTIONALITY - add the BUSI/RBIC functionality

    Case "JOBS"     'JOBS now

        If EARNED_INCOME_PANELS_ARRAY(this_is_a_new_panel, ei_panel) = TRUE OR EARNED_INCOME_PANELS_ARRAY(income_received, ei_panel) = TRUE Then

            'updating information for when the script ends
            end_msg = end_msg & vbNewLine & "Updated JOBS for MEMB " & EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel) & " at " & EARNED_INCOME_PANELS_ARRAY(employer, ei_panel)
            If EARNED_INCOME_PANELS_ARRAY(this_is_a_new_panel, ei_panel) = TRUE Then end_msg = end_msg & " panel added eff with start date " & EARNED_INCOME_PANELS_ARRAY(income_start_dt, ei_panel)
            If EARNED_INCOME_PANELS_ARRAY(income_received, ei_panel) = TRUE Then end_msg = end_msg & " income budgeted, panel updated."

            If developer_mode = FALSE Then Call start_a_blank_CASE_NOTE        'now we start the case note

            If EARNED_INCOME_PANELS_ARRAY(income_received, ei_panel) = TRUE Then        'if we have income verification - the note is more detailed
                STATS_manualtime = STATS_manualtime + 120
                If EARNED_INCOME_PANELS_ARRAY(income_verif, ei_panel) = "? - EXPEDITED SNAP ONLY" Then      'special header for if '?' is used as verification so they are easy to find
                    Call write_variable_in_CASE_NOTE("XFS INCOME DETAIL: M" & EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel) & " - JOBS - " & EARNED_INCOME_PANELS_ARRAY(employer, ei_panel) & " - PROG: " & prog_list)
                Else
                    Call write_variable_in_CASE_NOTE("INCOME DETAIL: M" & EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel) & " - JOBS - " & EARNED_INCOME_PANELS_ARRAY(employer, ei_panel) & " - PROG: " & prog_list)
                End If

                If EARNED_INCOME_PANELS_ARRAY(this_is_a_new_panel, ei_panel) = TRUE Then            'line in note about adding the panel
                    Call write_variable_in_CASE_NOTE("* THIS IS NEW INCOME. Started on " & EARNED_INCOME_PANELS_ARRAY(income_start_dt, ei_panel))
                End If

                If EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) = checked Then           'budget detail about SNAP
                    Call write_variable_in_CASE_NOTE("Income Budget for SNAP -------------------------------------")
                    If EARNED_INCOME_PANELS_ARRAY(income_verif, ei_panel) = "? - EXPEDITED SNAP ONLY" Then          'different wording for if '?' verif code is used
                        Call write_variable_in_CASE_NOTE("*** JOBS has been updated with information that has not been verified. ***")
                        Call write_variable_in_CASE_NOTE("* Month of application: " & fs_appl_footer_month & "/" & fs_appl_footer_year & ". Income updated to determine eligibility for Expedited SNAP, which does not have to be verifed.")
                        Call write_variable_in_CASE_NOTE("-- Income for " & EARNED_INCOME_PANELS_ARRAY(income_lumped_mo, ei_panel) & " has been entered on PIC as a single monthly payment. --")
                        Call write_variable_in_CASE_NOTE("* Total income for this month - $" & EARNED_INCOME_PANELS_ARRAY(lump_gross, ei_panel) & " - " & EARNED_INCOME_PANELS_ARRAY(lump_hrs, ei_panel) & " hrs. Pay Frequency: Monthly")
                        Call write_variable_in_CASE_NOTE("* Income budgeted this way for this month because " & EARNED_INCOME_PANELS_ARRAY(lump_reason, ei_panel))
                        Call write_variable_in_CASE_NOTE("* Amount determined using following checks:")
                        Call write_bullet_and_variable_in_CASE_NOTE("Actual Checks", EARNED_INCOME_PANELS_ARRAY(act_checks_lumped, ei_panel))
                        Call write_bullet_and_variable_in_CASE_NOTE("Anticipated Checks", EARNED_INCOME_PANELS_ARRAY(est_checks_lumped, ei_panel))
                    Else
                        If UH_SNAP = TRUE Then                  'didfferent wording (with the CASH_MONTHS_ARRAY information) for Uncle Harry SNAP
                            Call write_variable_in_CASE_NOTE("-- SNAP is UNCLE HARRY --")
                            For each_cash_month = 0 to UBOUND(CASH_MONTHS_ARRAY, 2)
                                If CASH_MONTHS_ARRAY(panel_indct, each_cash_month) = ei_panel Then
                                    Call write_variable_in_CASE_NOTE("* Income updated in " & CASH_MONTHS_ARRAY(cash_mo_yr, each_cash_month))
                                    If CASH_MONTHS_ARRAY(retro_updtd, each_cash_month) = TRUE Then Call write_variable_in_CASE_NOTE("* --RETRO Income updated: $" & CASH_MONTHS_ARRAY(mo_retro_pay, each_cash_month) & " total income for " & CASH_MONTHS_ARRAY(retro_mo_yr, each_cash_month) & " with " & CASH_MONTHS_ARRAY(mo_retro_hrs, each_cash_month) & " total hours.")
                                    If CASH_MONTHS_ARRAY(prosp_updtd, each_cash_month) = TRUE Then Call write_variable_in_CASE_NOTE("* --Prosp Income updated: $" & CASH_MONTHS_ARRAY(mo_prosp_pay, each_cash_month) & " total income for " & CASH_MONTHS_ARRAY(cash_mo_yr, each_cash_month) & " with " & CASH_MONTHS_ARRAY(mo_prosp_hrs, each_cash_month) & " total hours.")

                                End If
                            Next
                        Else                    'this is the standard wording for SNAP budget information
                            Call write_bullet_and_variable_in_CASE_NOTE("Monthly budgeted income", "$" & EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel))
                            Call write_bullet_and_variable_in_CASE_NOTE("Average per Pay Period", "$" & EARNED_INCOME_PANELS_ARRAY(snap_ave_inc_per_pay, ei_panel))
                            Call write_bullet_and_variable_in_CASE_NOTE("Average hours per week", EARNED_INCOME_PANELS_ARRAY(snap_hrs_per_wk, ei_panel))
                            Call write_bullet_and_variable_in_CASE_NOTE("Average pay per hour", "$" & EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel)& "/hr")
                            Call write_bullet_and_variable_in_CASE_NOTE("Pay Frequency", EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel))

                            If EARNED_INCOME_PANELS_ARRAY(income_lumped_mo, ei_panel) <> "" Then
                                Call write_variable_in_CASE_NOTE("-- Income for " & EARNED_INCOME_PANELS_ARRAY(income_lumped_mo, ei_panel) & " has been entered on PIC as a single monthly payment. --")
                                Call write_variable_in_CASE_NOTE("* Total income for this month - $" & EARNED_INCOME_PANELS_ARRAY(lump_gross, ei_panel) & " - " & EARNED_INCOME_PANELS_ARRAY(lump_hrs, ei_panel) & " hrs. Pay Frequency: Monthly")
                                Call write_variable_in_CASE_NOTE("* Income budgeted this way for this month because " & EARNED_INCOME_PANELS_ARRAY(lump_reason, ei_panel))
                                Call write_variable_in_CASE_NOTE("* Amount determined using following checks:")
                                Call write_bullet_and_variable_in_CASE_NOTE("Actual Checks", EARNED_INCOME_PANELS_ARRAY(act_checks_lumped, ei_panel))
                                Call write_bullet_and_variable_in_CASE_NOTE("Anticipated Checks", EARNED_INCOME_PANELS_ARRAY(est_checks_lumped, ei_panel))
                            End If
                            Call write_variable_in_CASE_NOTE( "*** Income has been reviewed and is anticipated to continue at this amount.")
                        End If
                    End If
                End If
                If EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, ei_panel) = checked Then           'Cash budget detail
                    Call write_variable_in_CASE_NOTE("Income Budget for CASH -------------------------------------")
                    If EARNED_INCOME_PANELS_ARRAY(excl_cash_rsn, ei_panel) <> "" Then Call write_variable_in_CASE_NOTE("* This income is not counted in the Cash budget. Reason: " & EARNED_INCOME_PANELS_ARRAY(excl_cash_rsn, ei_panel))
                    For each_cash_month = 0 to UBOUND(CASH_MONTHS_ARRAY, 2)
                        If CASH_MONTHS_ARRAY(panel_indct, each_cash_month) = ei_panel Then
                            Call write_variable_in_CASE_NOTE("* Income updated in " & CASH_MONTHS_ARRAY(cash_mo_yr, each_cash_month))
                            If CASH_MONTHS_ARRAY(retro_updtd, each_cash_month) = TRUE Then Call write_variable_in_CASE_NOTE("* --RETRO Income updated: $" & CASH_MONTHS_ARRAY(mo_retro_pay, each_cash_month) & " total income for " & CASH_MONTHS_ARRAY(retro_mo_yr, each_cash_month) & " with " & CASH_MONTHS_ARRAY(mo_retro_hrs, each_cash_month) & " total hours.")
                            If CASH_MONTHS_ARRAY(prosp_updtd, each_cash_month) = TRUE Then Call write_variable_in_CASE_NOTE("* --Prosp Income updated: $" & CASH_MONTHS_ARRAY(mo_prosp_pay, each_cash_month) & " total income for " & CASH_MONTHS_ARRAY(cash_mo_yr, each_cash_month) & " with " & CASH_MONTHS_ARRAY(mo_prosp_hrs, each_cash_month) & " total hours.")

                        End If
                    Next

                End If
                If EARNED_INCOME_PANELS_ARRAY(apply_to_HC, ei_panel) = checked Then             'Health Care Budget Detail
                    Call write_variable_in_CASE_NOTE("Income Budget for HEALTH CARE ----------------------------------")
                    Call write_bullet_and_variable_in_CASE_NOTE("Average per Pay Period", "$" & EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel))
                    Call write_bullet_and_variable_in_CASE_NOTE("Pay Frequency", EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel))
                    If EARNED_INCOME_PANELS_ARRAY(hc_retro, ei_panel) = TRUE Then
                        Call write_variable_in_CASE_NOTE("* Income updated in " & CASH_MONTHS_ARRAY(cash_mo_yr, each_cash_month))
                        If CASH_MONTHS_ARRAY(retro_updtd, each_cash_month) = TRUE Then Call write_variable_in_CASE_NOTE("* --RETRO Income updated: $" & CASH_MONTHS_ARRAY(mo_retro_pay, each_cash_month) & " total income for " & CASH_MONTHS_ARRAY(retro_mo_yr, each_cash_month) & " with " & CASH_MONTHS_ARRAY(mo_retro_hrs, each_cash_month) & " total hours.")
                        If CASH_MONTHS_ARRAY(prosp_updtd, each_cash_month) = TRUE Then Call write_variable_in_CASE_NOTE("* --Prosp Income updated: $" & CASH_MONTHS_ARRAY(mo_prosp_pay, each_cash_month) & " total income for " & CASH_MONTHS_ARRAY(cash_mo_yr, each_cash_month) & " with " & CASH_MONTHS_ARRAY(mo_prosp_hrs, each_cash_month) & " total hours.")
                    End If
                    Call write_bullet_and_variable_in_CASE_NOTE("Notes on HC Budget", EARNED_INCOME_PANELS_ARRAY(hc_budg_notes, ei_panel))
                End If

                If EARNED_INCOME_PANELS_ARRAY(apply_to_GRH, ei_panel) = checked Then            'GRH budget detail
                    Call write_variable_in_CASE_NOTE("Income Budget for GRH ----------------------------------")
                    Call write_bullet_and_variable_in_CASE_NOTE("Monthly budgeted income", "$" & EARNED_INCOME_PANELS_ARRAY(GRH_mo_inc, ei_panel))
                    Call write_bullet_and_variable_in_CASE_NOTE("Average per Pay Period", "$" & EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel))
                End If

                'Every program gets information about what verification was provided
                'Though most of this is really for SNAP requirements - it is still relevant to other programs
                Call write_variable_in_CASE_NOTE("Verification Received: " & EARNED_INCOME_PANELS_ARRAY(verif_date, ei_panel) & "-----------------------------")

                Call write_bullet_and_variable_in_CASE_NOTE("Type Received", EARNED_INCOME_PANELS_ARRAY(income_verif, ei_panel))
                Call write_bullet_and_variable_in_CASE_NOTE("Explanation of Verification", EARNED_INCOME_PANELS_ARRAY(verif_explain, ei_panel))
                Call write_bullet_and_variable_in_CASE_NOTE("Days covered by check stubs", EARNED_INCOME_PANELS_ARRAY(days_of_verif, ei_panel))

                Call write_bullet_and_variable_in_CASE_NOTE("Conversation with", EARNED_INCOME_PANELS_ARRAY(spoke_with, ei_panel))
                Call write_bullet_and_variable_in_CASE_NOTE("Conversation Details", EARNED_INCOME_PANELS_ARRAY(convo_detail, ei_panel))

                'Basically a list of the verification that was received
                Call write_variable_in_CASE_NOTE("Income Information Received -----------------------------------")

                If EARNED_INCOME_PANELS_ARRAY(order_ubound, ei_panel) <> "" AND EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel) <> "" Then        'If there is an order ubound then there are actual checks'
                    Call write_variable_in_CASE_NOTE("* Both actual check stubs and anticipated income estimates were received for this income.")

                    If EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_actual Then Call write_variable_in_CASE_NOTE("* Actual pay amounts used to determine income to budget.")
                    If EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_estimate Then Call write_variable_in_CASE_NOTE("* Income to budget determined by anticipated hours and rate of pay.")
                    Call write_bullet_and_variable_in_CASE_NOTE("Reason for choice", EARNED_INCOME_PANELS_ARRAY(selection_rsn, ei_panel))
                End If

                If EARNED_INCOME_PANELS_ARRAY(order_ubound, ei_panel) <> "" Then            'list of checks in order
                    If EARNED_INCOME_PANELS_ARRAY(income_verif, ei_panel) = "? - EXPEDITED SNAP ONLY" Then
                        Call write_variable_in_CASE_NOTE("* Pay information reported by client")
                    Else
                        Call write_variable_in_CASE_NOTE("* Checks provided to agency.")
                    End If
                    For order_number = 1 to top_of_order                        'loop through the order number lowest to highest
                        For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                            'conditional if it is the right panel AND the order matches - then do the thing you need to do
                            If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel AND LIST_OF_INCOME_ARRAY(check_order, all_income) = order_number Then
                                If EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) = checked Then           'different formatting for different scenarios
                                    If LIST_OF_INCOME_ARRAY(budget_in_SNAP_yes, all_income) = checked Then
                                        If LIST_OF_INCOME_ARRAY(exclude_amount, all_income) <> 0 Then
                                            Call write_bullet_and_variable_in_CASE_NOTE(LIST_OF_INCOME_ARRAY(view_pay_date, all_income), "Gross: $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & " hrs. Only $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) - LIST_OF_INCOME_ARRAY(exclude_amount, all_income) & " included in SNAP budget because: " & LIST_OF_INCOME_ARRAY(reason_amt_excluded, all_income) & " - $" & LIST_OF_INCOME_ARRAY(exclude_amount, all_income) & " of check not included.")
                                        Else
                                            Call write_bullet_and_variable_in_CASE_NOTE(LIST_OF_INCOME_ARRAY(view_pay_date, all_income), "Gross: $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & " hrs.")
                                        End If
                                    Else
                                        Call write_bullet_and_variable_in_CASE_NOTE(LIST_OF_INCOME_ARRAY(view_pay_date, all_income), "Gross: $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & " hrs. ** THIS CHECK EXCLUDED FROM SNAP BUDGET because " & LIST_OF_INCOME_ARRAY(reason_to_exclude, all_income))
                                    End If
                                Else
                                    Call write_bullet_and_variable_in_CASE_NOTE(LIST_OF_INCOME_ARRAY(view_pay_date, all_income), "Gross: $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & " hrs.")
                                End If
                                If LIST_OF_INCOME_ARRAY(future_check, all_income) = TRUE Then Call write_variable_in_CASE_NOTE("        Pay Date in future - reported expected amount, Only used for SNAP budget in month of application.")
                            End If
                        next
                    next
                    'special notes about using the SNAP PIC
                    If EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) = checked AND EARNED_INCOME_PANELS_ARRAY(pick_one, ei_panel) = use_actual Then Call write_variable_in_CASE_NOTE("* All included checks have been added to the PIC. Gross amount on PIC is reflective of the included pay amount.")
                End If
                If EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel) <> "" Then
                    Call write_variable_in_CASE_NOTE("* Anticipated Income Estimate provided to Agency.")

                    Call write_bullet_and_variable_in_CASE_NOTE("Hourly Pay Rate", "$" & EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel) & "/hr")
                    Call write_bullet_and_variable_in_CASE_NOTE("Hours Per Week", EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel) & " hours")
                    Call write_bullet_and_variable_in_CASE_NOTE("Pay Frequency", EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel))
                End If

                Call write_variable_in_CASE_NOTE("ACTION TAKEN: JOBS Updated ------------------------------------")         'currently very basic
                If left(EARNED_INCOME_PANELS_ARRAY(months_updated, ei_panel), 1) = "," Then  EARNED_INCOME_PANELS_ARRAY(months_updated, ei_panel) = right(EARNED_INCOME_PANELS_ARRAY(months_updated, ei_panel), len(EARNED_INCOME_PANELS_ARRAY(months_updated, ei_panel)) - 1)

                Call write_bullet_and_variable_in_CASE_NOTE("Months updated", EARNED_INCOME_PANELS_ARRAY(months_updated, ei_panel))
            Else        'this is if there was ONLY the panel added but no verification of income was entered
                Call write_variable_in_CASE_NOTE("New Job: M" & EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel) & " - JOBS - " & EARNED_INCOME_PANELS_ARRAY(employer, ei_panel) & " - PROG: " & prog_list)
                Call write_variable_in_CASE_NOTE("* Information received that a new job has started.")
            End If

            Call write_variable_in_CASE_NOTE("---")
            Call write_variable_in_CASE_NOTE(worker_signature)
        End If
    Case "BUSI"     'FUTURE FUNCTIONALITY

    End Select
Next

If end_msg = "" Then end_msg = "Script ended with no action taken, panels not updated, no case note created. No new panels were indicated and no income verification was entered to be budgeted."

script_end_procedure_with_error_report(end_msg)
