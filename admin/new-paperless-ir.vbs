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
const actually_paperless = 18
const case_notes    = 19

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
        ALL_HC_REVS_ARRAY (actually_paperless, hc_reviews) = TRUE
        'figure out who is on HC
        For each pers_nbr in HC_PERS_ARRAY
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
                    ALL_HC_REVS_ARRAY (case_notes, hc_reviews) = ALL_HC_REVS_ARRAY (case_notes, hc_reviews) & "; " & pers_nbr
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
                If total_panels <> "0" & date_check = "__ __ __" then ALL_HC_REVS_ARRAY (actually_paperless, hc_reviews) = False
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
                If total_panels <> "0" & date_check = "__ __ __" then ALL_HC_REVS_ARRAY (actually_paperless, hc_reviews) = False
                if current_panel <> total_panels then transmit
            Loop until current_panel = total_panels

        Next

        ALL_HC_REVS_ARRAY(memb_on_hc, hc_reviews) = Join(HC_PERS_ARRAY, ", ")
        ReDim HC_PERS_ARRAY
    Else
        ALL_HC_REVS_ARRAY (revw_type, hc_reviews) = "PRIV"
    End If
Next

'Create Multidimensional Array for all of the information about each case
'Go to REVS for each worker and get all of the Exempt Cases along with other reviews due
'Go to STAT and confirm actually paperless
'Update REVW if actually paperless for ALL Members
    'double check on what to do with multiple members
    'figure out how to handle for other dates/reviews
'TIKL for each case updated
'Add to reason for why if not updated
'Confirm TIKL

'Dump array on to spreadsheet
'Some stats would be good - basics
'
script_end_procedure("Success! All starred (*) IRs have been sent into background, except those with current JOBS/BUSI/RBIC, those who have members other than 01 open, or those who also have SNAP up for review. You must go through and approve these results when they come through background.")
