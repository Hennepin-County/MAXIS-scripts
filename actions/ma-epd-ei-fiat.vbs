'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - MA-EPD EI FIAT.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 600                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
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
call changelog_update("05/07/2018", "Updated the script to identify cases at application versus review, and provide different functionality for those options. Average income will now be determined from the budget on ELIG.", "Casey Love, Hennepin County")
call changelog_update("04/23/2018", "Added functionality to allow any month to be selected as the first month to be FIATed.", "Casey Love, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS==================================================================================================================

function get_average_pay(job_frequency, job_income)
'This function was created for the initial process of determining average income'
'It is not currently being used in the script.
'Need to go to JOBS before calling this function - not going within the function to allow for correct navigation for the right person and instance

    'Setting these to blanks for each use of the function'
    anticipated_total = ""
    divider = divider = ""
    anticipated_average = ""
    hc_inc_est = ""
    use_hc_est_inc_radio = ""
    use_anticipated_inc_radio = ""

    'Reading the pay frequency and then setting it to a string that is more readable
    EMReadScreen pay_freq, 1, 18, 35
    If pay_freq = "1" then job_frequency = "1: monthly"
    If pay_freq = "2" then job_frequency = "2: twice monthly"
    If pay_freq = "3" then job_frequency = "3: every 2 weeks"
    If pay_freq = "4" then job_frequency = "4. every week"
    If pay_freq = "5" then job_frequency = "5. other (use monthly avg)"

    'Reading the anticipated total of income and defining a blank
    EMReadScreen anticipated_total, 8, 17, 67
    anticipated_total = trim(anticipated_total)
    If anticipated_total = "" Then anticipated_total = 0

    jobs_row = 12       'This is where pay information starts
    divider = 0         'This will count the number of pay dates listed in the prospective side to be used to calculate the average.
    Do
        EMReadScreen pay_date, 8, jobs_row, 54                  'Read the date of pay
        If pay_date <> "__ __ __" Then divider = divider + 1    'If a date is here - add another to the count of the number of checks
        jobs_row = jobs_row + 1                                 'Go to the next row in prospective pay
    Loop until jobs_row = 17                                    'There are only 5 rows of paychecks. Once it reaches 17 - there are no more checks to read

    anticipated_average = anticipated_total / divider           'Finding the average income per check by dividing the total listed on prospective jobs side by the number of checks.
    anticipated_average = FormatNumber(anticipated_average, 2)

    EMReadScreen HC_income_est_check, 3, 19, 63 'reading to find the HC income estimator is moving 6/1/16, to account for if it only affects future months we are reading to find the HC inc EST
    IF HC_income_est_check = "Est" Then 'this is the old position
      EMWriteScreen "x", 19, 54
    ELSE								'this is the new position
      EMWriteScreen "x", 19, 48
    END IF
    transmit                            'opening the HC Inc Estimate pop-up
    EMReadScreen hc_inc_est, 8, 11, 63  'Reading the income on this field.'
    hc_inc_est = trim(replace(hc_inc_est, "_", "")) 'Fomatting the number'
    transmit                            'closing the HC Inc Est pop-up

    if hc_inc_est = "" Then hc_inc_est = 0      'Making this a number
    hc_inc_est = FormatNumber(hc_inc_est, 2)

    If hc_inc_est <> anticipated_average Then       'These two should be equal - because HC Inc Estimate is based on the average of pay
    'If they are not equal - script will ask the worker to clarify which is correct.
        BeginDialog income_mismatch_dlg, 0, 0, 221, 105, "Income Mismatch"
          OptionGroup RadioGroup1
            RadioButton 10, 30, 205, 10, "Use the amount from the HC Inc Est Pop-Up of $" & hc_inc_est, use_hc_est_inc_radio
            RadioButton 10, 45, 225, 10, "Use the amount from anticipated income on JOBS of $" & anticipated_average, use_anticipated_inc_radio
          ButtonGroup ButtonPressed
            OkButton 105, 85, 50, 15
            CancelButton 165, 85, 50, 15
          Text 5, 5, 210, 20, "It appears that the average income listed on this JOBS panel does not match. Please indicate which amount should be used."
          Text 5, 65, 190, 10, "These amounts are both average per pay period amounts."
        EndDialog

        Dialog income_mismatch_dlg      'Running the dialog to ask for worker input on the correct income.
        Cancel_confirmation

        'This will set the average income for the job based on what the worker indicates
        If use_anticipated_inc_radio = 1 Then job_income = anticipated_average
        If use_hc_est_inc_radio = 1 Then job_income = hc_inc_est
    Else
        job_income = hc_inc_est     'If they are equal - this is just setting the income to the variable used later in the script
    End If

end function

'END FUNCTIONS==============================================================================================================

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
' current_month_plus_one = dateadd("m", 1, date)
'
' MAXIS_footer_month = datepart("m", current_month_plus_one)
' If len(MAXIS_footer_month) = 1 then MAXIS_footer_month = "0" & MAXIS_footer_month
'
' MAXIS_footer_year = datepart("yyyy", current_month_plus_one)
' MAXIS_footer_year = MAXIS_footer_year - 2000
'
' current_month = datepart("m", date)
' If len(current_month) = 1 then current_month = "0" & current_month
'
' current_year = datepart("yyyy", date)
' current_year = current_year - 2000
'
' current_month_and_year = current_month & "/" & current_year
' next_month_and_year = MAXIS_footer_month & "/" & MAXIS_footer_year

'DIALOGS--------------------------------
BeginDialog case_number_dialog, 0, 0, 161, 85, "Case number"
  EditBox 90, 5, 65, 15, MAXIS_case_number
  EditBox 90, 25, 30, 15, memb_number
  DropListBox 90, 45, 65, 45, "Select One..."+chr(9)+"Application"+chr(9)+"Recertification", case_status
  ButtonGroup ButtonPressed
    OkButton 40, 65, 50, 15
    CancelButton 100, 65, 50, 15
  Text 5, 10, 80, 10, "Enter your case number:"
  Text 20, 30, 65, 10, "HH memb number:"
  Text 50, 50, 35, 10, "Case is at"
EndDialog

'THE SCRIPT

EMConnect ""

'Autofilling information
call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

memb_number = "01" 'Setting a default

'If we have found a case number, the script will attempt to determine if the case is pending or active
If MAXIS_case_number <> "" Then
    Call Navigate_to_MAXIS_screen("CASE", "CURR")   'go to CASe/CURR to read for HC
    row = 1
    col = 1
    EMSearch "MA:", row, col        'Searhing for the MA line because it moves
    If row <> 0 Then                'If the script finds an MA line, it will read the status (pending or active)'
        EMReadScreen ma_status, 7, row, col+4   'Reading the status'
        'MsgBox ma_status
        ma_status = trim(ma_status)             'cutting blank
        If ma_status = "ACTIVE" Then case_status = "Recertification"    'If a case is alread active, it is often at review'
        If ma_status = "PENDING" Then case_status = "Application"       'If a case is pending then it is usually at application
    End If
End If

'Running a dialog to get case number, member number and if the case is at application or recertification.'
Do
    err_msg = ""

    Dialog case_number_dialog
    Cancel_confirmation

    If MAXIS_case_number = "" Then                                             err_msg = err_msg & vbNewLine & "* Enter a case number to continue."
    If IsNumeric(MAXIS_case_number) = FALSE or len(MAXIS_case_number) > 8 Then err_msg = err_msg & vbNewLine & "* Case number appears to be invalid. Check the case number and fix."
    If memb_number = "" Then                                                   err_msg = err_msg & vbNewLine & "* Enter a reference number for the member on MA-EPD."
    If case_status = "Select One..." Then                                      err_msg = err_msg & vbNewLine & "* Identify if case is at recertification or application."
    'If MAXIS_footer_month = "" OR MAXIS_footer_year = "" Then                  err_msg = err_msg & vbNewLine & "* Enter the MAXIS footer month and year that has the best income information in it."

    If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
Loop until err_msg = ""

const instance          = 0
const job_frequency     = 1
const check_date_one    = 2
const check_date_two    = 3
const check_date_three  = 4
const check_date_four   = 5
const check_date_five   = 6
const check_amt_one     = 7
const check_amt_two     = 8
const check_amt_three   = 9
const check_amt_four    = 10
const check_amt_five    = 11
const pay_average       = 12
const pay_weekday       = 13
const six_month_total   = 14
const average_monthly_inc   = 15
const employer          = 16
const verif_code        = 17

Dim JOBS_ARRAY()
ReDim JOBS_ARRAY(verif_code, 0)

If case_status = "Recertification" Then

    Call Navigate_to_MAXIS_screen("STAT", "REVW")

    EMReadScreen hc_revw, 8, 9, 70
    hc_revw = replace(hc_revw, " ", "/")

    If DateDiff("D", CM_plus_1_mo & "/01/" & CM_plus_1_yr, hc_revw) > 0 Then
        EMReadScreen hc_revw, 8, 11, 70
        hc_revw = replace(hc_revw, " ", "/")
    End If
    MAXIS_footer_month = DatePart("m", hc_revw)
    MAXIS_footer_month = right("00" & MAXIS_footer_month, 2)

    MAXIS_footer_year = DatePart("yyyy", hc_revw)
    MAXIS_footer_year = right(MAXIS_footer_year, 2)

    Call back_to_SELF

    Call Navigate_to_MAXIS_screen("STAT", "JOBS")
    EmWriteScreen memb_number, 20, 76
    EmWriteScreen "01", 20, 79
    transmit

    EMReadScreen number_of_jobs, 1, 2, 78
    number_of_jobs = number_of_jobs * 1

    end_msg = "Household Member " & member_number & " on this case does not have a JOBS panel entered. Please check the case, update JOBS if required and run the script again."
    If number_of_jobs = 0 Then script_end_procedure(end_msg)

    For each_job = 1 to number_of_jobs
        EMReadScreen job_verification, 1, 6, 34
        EMReadScreen first_check_month, 2, 12, 54

        If job_verification = "?" OR first_check_month <> MAXIS_footer_month Then script_end_procedure("It does not appear this JOBS panel has been updated with income information for the review.")

        reDim Preserve JOBS_ARRAY(verif_code, each_job-1)

        EMReadScreen verification, 25, 6, 34
        EMReadScreen freq, 1, 18, 35
        EMReadScreen title, 30, 7, 42

        JOBS_ARRAY(verif_code, each_job-1) = trim(verification)
        JOBS_ARRAY(job_frequency, each_job-1) = freq
        JOBS_ARRAY(instance, each_job-1) = right("00"&each_job, 2)
        JOBS_ARRAY(employer, each_job-1) = replace(title, "_", "")

        JOBS_ARRAY(six_month_total, each_job-1) = 0
        EMReadScreen pay_date, 8, 12, 54

        pay_date = replace(pay_date, " ", "/")

        If JOBS_ARRAY(job_frequency, each_job-1) = "3" OR JOBS_ARRAY(job_frequency, each_job-1) = "4" Then
            day_validation_needed = FALSE
            JOBS_ARRAY(pay_weekday, each_job-1) = WeekDayName(WeekDay(pay_date))

        End If

        EMReadScreen HC_income_est_check, 3, 19, 63 'reading to find the HC income estimator is moving 6/1/16, to account for if it only affects future months we are reading to find the HC inc EST
        IF HC_income_est_check = "Est" Then 'this is the old position
          EMWriteScreen "x", 19, 54
        ELSE								'this is the new position
          EMWriteScreen "x", 19, 48
        END IF
        transmit
        EMReadScreen hc_inc_est, 8, 11, 63
        hc_inc_est = trim(replace(hc_inc_est, "_", ""))
        transmit

        if hc_inc_est = "" Then hc_inc_est = 0
        hc_inc_est = FormatNumber(hc_inc_est, 2)


        transmit
    Next


End If

If case_status = "Application" Then

    Call Navigate_to_MAXIS_screen("STAT", "HCRE")

    hcre_row = 10
    Do
        EMReadScreen hcre_ref_numb, 2, hcre_row, 24
        If hcre_ref_numb = memb_number Then Exit Do

        hcre_row = hcre_row + 1
        If hcre_row = 18 Then
            PF20
            hcre_row = 10
        End If
        EMReadScreen next_client, 2, hcre_row, 24
    Loop until next_client = "  "

    EMReadScreen application_date, 8, hcre_row, 51
    EMReadScreen coverage_date, 5, hcre_row, 64

    application_date = replace(application_date, " ", "/")

    MAXIS_footer_month = DatePart("m", application_date)
    MAXIS_footer_month = right("00" & MAXIS_footer_month, 2)

    MAXIS_footer_year = DatePart("yyyy", application_date)
    MAXIS_footer_year = right(MAXIS_footer_year, 2)

    'MsgBox coverage_date
    If left(coverage_date, 2) <> MAXIS_footer_month OR right(coverage_date, 2) <> MAXIS_footer_year Then
        coverage_date = replace(coverage_date, " ", "/")
        MsgBox "This case appears to have a retro request back to " & coverage_date & "." & vbNewLine & vbNewLine & "Retro months should not be FIATed to even the income out. The premium in these months are based on actual income and will be different."
    End If

    Call back_to_SELF

    Call Navigate_to_MAXIS_screen("STAT", "JOBS")
    EmWriteScreen memb_number, 20, 76
    EmWriteScreen "01", 20, 79
    transmit

    EMReadScreen number_of_jobs, 1, 2, 78
    number_of_jobs = number_of_jobs * 1

    end_msg = "Household Member " & member_number & " on this case does not have a JOBS panel entered. Please check the case, update JOBS if required and run the script again."
    If number_of_jobs = 0 Then script_end_procedure(end_msg)

    For each_job = 1 to number_of_jobs
        reDim Preserve JOBS_ARRAY(verif_code, each_job-1)

        EMReadScreen verification, 25, 6, 34
        EMReadScreen freq, 1, 18, 35
        EMReadScreen title, 30, 7, 42

        JOBS_ARRAY(verif_code, each_job-1) = trim(verification)
        JOBS_ARRAY(job_frequency, each_job-1) = freq
        JOBS_ARRAY(instance, each_job-1) = right("00"&each_job, 2)
        JOBS_ARRAY(employer, each_job-1) = replace(title, "_", "")

        jobs_row = 12
        divider = 0
        Do
            EMReadScreen pay_date, 8, jobs_row, 54
            If pay_date <> "__ __ __" Then
                divider = divider + 1
                If JOBS_ARRAY(check_date_one, each_job-1) = "" Then
                    JOBS_ARRAY(check_date_one, each_job-1) = replace(pay_date, " ", "/")
                    EMReadScreen pay_amt, 8, jobs_row, 67
                    JOBS_ARRAY(check_amt_one, each_job-1) = trim(pay_amt) * 1

                ElseIf JOBS_ARRAY(check_date_two, each_job-1) = "" Then
                    JOBS_ARRAY(check_date_two, each_job-1) = replace(pay_date, " ", "/")
                    EMReadScreen pay_amt, 8, jobs_row, 67
                    JOBS_ARRAY(check_amt_two, each_job-1) = trim(pay_amt) * 1

                ElseIf JOBS_ARRAY(check_date_three, each_job-1) = "" Then
                    JOBS_ARRAY(check_date_three, each_job-1) = replace(pay_date, " ", "/")
                    EMReadScreen pay_amt, 8, jobs_row, 67
                    JOBS_ARRAY(check_amt_three, each_job-1) = trim(pay_amt) * 1

                ElseIf JOBS_ARRAY(check_date_four, each_job-1) = "" Then
                    JOBS_ARRAY(check_date_four, each_job-1) = replace(pay_date, " ", "/")
                    EMReadScreen pay_amt, 8, jobs_row, 67
                    JOBS_ARRAY(check_amt_four, each_job-1) = trim(pay_amt) * 1

                ElseIf JOBS_ARRAY(check_date_five, each_job-1) = "" Then
                    JOBS_ARRAY(check_date_five, each_job-1) = replace(pay_date, " ", "/")
                    EMReadScreen pay_amt, 8, jobs_row, 67
                    JOBS_ARRAY(check_amt_five, each_job-1) = trim(pay_amt) * 1

                End If
            End If

            jobs_row = jobs_row + 1
        Loop until jobs_row = 17

        EMReadScreen total_pay, 8, 17, 67
        total_pay = trim(total_pay)
        total_pay = total_pay * 1

        JOBS_ARRAY(pay_average, each_job-1) = total_pay / divider
        JOBS_ARRAY(six_month_total, each_job-1) = 0

        If JOBS_ARRAY(job_frequency, each_job-1) = "3" OR JOBS_ARRAY(job_frequency, each_job-1) = "4" Then
            day_validation_needed = FALSE
            JOBS_ARRAY(pay_weekday, each_job-1) = WeekDayName(WeekDay(JOBS_ARRAY(check_date_one, each_job-1)))
            If JOBS_ARRAY(check_date_two, each_job-1) <> "" Then
                If WeekDayName(WeekDay(JOBS_ARRAY(check_date_two, each_job-1))) <> JOBS_ARRAY(pay_weekday, each_job-1) Then day_validation_needed = TRUE
            End If
            If JOBS_ARRAY(check_date_three, each_job-1) <> "" Then
                If WeekDayName(WeekDay(JOBS_ARRAY(check_date_three, each_job-1))) <> JOBS_ARRAY(pay_weekday, each_job-1) Then day_validation_needed = TRUE
            End If
            If JOBS_ARRAY(check_date_four, each_job-1) <> "" Then
                If WeekDayName(WeekDay(JOBS_ARRAY(check_date_four, each_job-1))) <> JOBS_ARRAY(pay_weekday, each_job-1) Then day_validation_needed = TRUE
            End If
            If JOBS_ARRAY(check_date_five, each_job-1) <> "" Then
                If WeekDayName(WeekDay(JOBS_ARRAY(check_date_five, each_job-1))) <> JOBS_ARRAY(pay_weekday, each_job-1) Then day_validation_needed = TRUE
            End If
        End If

        If day_validation_needed = TRUE Then
            selected_weekday = JOBS_ARRAY(pay_weekday, each_job-1)

            BeginDialog weekday_dlg, 0, 0, 161, 80, "Weekday"
              DropListBox 15, 55, 60, 45, "Sunday"+chr(9)+"Monday"+chr(9)+"Tuesday"+chr(9)+"Wednesday"+chr(9)+"Thursday"+chr(9)+"Friday"+chr(9)+"Saturday", selected_weekday
              ButtonGroup ButtonPressed
                OkButton 105, 55, 50, 15
              Text 5, 10, 150, 35, "This job is paid either weekly or biweekly, but has different weekdays indicated for pay dates. Please select the weekday that the client is paid."
            EndDialog

            Dialog weekday_dlg

            JOBS_ARRAY(pay_weekday, each_job-1) = selected_weekday

        End If

        ' BeginDialog JOBS_dlg, 0, 0, 200, 205, "JOBS"
        '   Text 10, 10, 190, 10, "JOBS for MEMB " & member_number & "- Instance " & JOBS_ARRAY(instance, each_job-1)
        '   Text 15, 30, 150, 10, "Checks - Verif " & JOBS_ARRAY(verif_code, each_job-1)
        '   Text 55, 50, 120, 10, JOBS_ARRAY(check_date_one, each_job-1) & " - $" & JOBS_ARRAY(check_amt_one, each_job-1)
        '   Text 55, 65, 120, 10, JOBS_ARRAY(check_date_two, each_job-1) & " - $" & JOBS_ARRAY(check_amt_two, each_job-1)
        '   Text 55, 80, 120, 10, JOBS_ARRAY(check_date_three, each_job-1) & " - $" & JOBS_ARRAY(check_amt_three, each_job-1)
        '   Text 55, 95, 120, 10, JOBS_ARRAY(check_date_four, each_job-1) & " - $" & JOBS_ARRAY(check_amt_four, each_job-1)
        '   Text 55, 110, 120, 10, JOBS_ARRAY(check_date_five, each_job-1) & " - $" & JOBS_ARRAY(check_amt_five, each_job-1)
        '   Text 55, 125, 120, 10, "Total - $" & total_pay
        '   Text 10, 140, 120, 10, "Frequency - " & JOBS_ARRAY(job_frequency, each_job-1)
        '   Text 10, 155, 120, 10, "Payday is on " & JOBS_ARRAY(pay_weekday, each_job-1)
        '   Text 10, 170, 120, 10, "Average - $" & JOBS_ARRAY(pay_average, each_job-1)
        '   ButtonGroup ButtonPressed
        '     OkButton 10, 185, 50, 15
        ' EndDialog
        '
        ' Dialog JOBS_dlg

        transmit
    Next

    app_month = MAXIS_footer_month
    app_year = MAXIS_footer_year

    first_of_this_month = MAXIS_footer_month & "/1/" & MAXIS_footer_year
    'MsgBox "FIRST - " &first_of_this_month
    next_month = DateAdd("m", 1, first_of_this_month)

    Do
        next_month_mo = DatePart("m", next_month)
        next_month_mo = right("00"&next_month_mo, 2)
        next_month_yr = DatePart("yyyy", next_month)
        next_month_yr = right(next_month_yr, 2)

        list_of_months = list_of_months & "~" & next_month_mo & "/" & next_month_yr
        'MsgBox "List " & list_of_months

        first_of_this_month = next_month_mo & "/1/" & next_month_yr
        next_month = DateAdd("m", 1, first_of_this_month)

        'MsgBox "Start month and year - " & next_month & vbNewLine & "DIFF " & DateDiff("d", date, next_month)
    Loop until DateDiff("d", date, next_month) > 0

    'MsgBox "Complete List " & list_of_months
    list_of_months = right(list_of_months, len(list_of_months)-1)
    month_array = split(list_of_months, "~")

    ' For each thingy in month_array
    '     MsgBox thingy
    ' Next
    paychecks_align_with_weekday = TRUE
    For each footer in month_array
        Call back_to_SELF

        MAXIS_footer_month = left(footer, 2)
        MAXIS_footer_year = right(footer, 2)

        Call Navigate_to_MAXIS_screen("STAT", "JOBS")
        EmWriteScreen member_number, 20, 76
        transmit

        For the_job = 0 to UBOUND(JOBS_ARRAY, 2)
            If JOBS_ARRAY(job_frequency, the_job) = "3" OR JOBS_ARRAY(job_frequency, the_job) = "4" Then
                EmWriteScreen JOBS_ARRAY(instance, the_job), 20, 79
                transmit

                jobs_row = 12
                divider = 0
                Do
                    EMReadScreen pay_date, 8, jobs_row, 54
                    If pay_date <> "__ __ __" Then
                        pay_date = replace(pay_date, " ", "/")
                        day_of_pay = WeekDayName(Weekday(pay_date))

                        If day_of_pay <> JOBS_ARRAY(pay_weekday, the_job) Then

                            confirm_off_schedule_pay = MsgBox("The pay listed on theis JOBS panel does note match the day of the week pay was received in the initial month of applicaton." & vbNewLine & vbNewLine &_
                             "Pay date of " & pay_date & " listed is on a " & day_of_pay & "." & vbNewLine & "This job appears to have a regular pay date of " & JOBS_ARRAY(pay_weekday, this_job) & "." & vbNewLine & vbNewLine &_
                             "Health care budget requires any income entered on JOBS be the actual pay dates expected, even if the income is calculated by average pay. Review the case and make sure that check dates have been updated in every month." & vbNewLine & vbNewLine &_
                             "Has the budget been correctly determined, using actual pay dates for each month that can be updated?", vbYesNo + vbImportant, "Confirm paycheck budgeting")
                            'TODO add dialog here to have the paycheck confirmed.'

                            if confirm_off_schedule_pay = vbNo Then script_end_procedure("Update STAT/JOBS with all actual pay dates to get a correct HC budget.")
                        End If

                    End If
                    divider = divider + 1
                    jobs_row = jobs_row + 1
                Loop until jobs_row = 17
            End If
        Next
    Next

    MAXIS_footer_month = app_month
    MAXIS_footer_year = app_year

    Call Navigate_to_MAXIS_screen("ELIG", "HC__")


End If

start_month_and_year = MAXIS_footer_month & "/" & MAXIS_footer_year

'SECTION 04: NOW IT GOES TO ELIG/HC TO FIAT THE AMOUNTS
Call navigate_to_MAXIS_screen("ELIG", "HC__")

row = 1
col = 1
EMSearch memb_number & " ", row, col 'finding the member number
If row = 0 then script_end_procedure("Member number not found. You may have entered an incorrect member number on the first screen. Try the script again.")

EMWriteScreen "x", row, 26
transmit

EMReadScreen elig_type_check_first_month, 2, 12, 17
EMReadScreen elig_type_check_second_month, 2, 12, 28
EMReadScreen elig_type_check_third_month, 2, 12, 39
EMReadScreen elig_type_check_fourth_month, 2, 12, 50
EMReadScreen elig_type_check_fifth_month, 2, 12, 61
EMReadScreen elig_type_check_sixth_month, 2, 12, 72

If elig_type_check_first_month <> "DP" and elig_type_check_second_month <> "DP" and elig_type_check_third_month <> "DP" and elig_type_check_fourth_month <> "DP" and elig_type_check_fifth_month <> "DP" and elig_type_check_sixth_month <> "DP" then MsgBox "Not all of the months of this case are MA-EPD. Process manually."
If elig_type_check_first_month <> "DP" and elig_type_check_second_month <> "DP" and elig_type_check_third_month <> "DP" and elig_type_check_fourth_month <> "DP" and elig_type_check_fifth_month <> "DP" and elig_type_check_sixth_month <> "DP" then stopscript

row = 6
col = 1
EMSearch start_month_and_year, row, col

end_msg = "The selected month " & start_month_and_year & " is not in the current version of HC, review the month selected and try the script again."
If col = 0 Then script_end_procedure(end_msg)

number_of_months = 0
budg_pd_wages = 0
Do
    EMWriteScreen "x", 9, col + 2
    transmit
    EMWriteScreen "x", 13, 03
    transmit

    this_job = 0
    budg_row = 8
    Do
        EMReadScreen inc_type, 2, budg_row, 8
        If inc_type = "02" Then
            EMReadScreen month_total, 11, budg_row, 43
            month_total = replace(month_total, "_", "")
            month_total = trim(month_total)
            'MsgBox month_total
            month_total = month_total * 1
            JOBS_ARRAY(six_month_total, this_job) = JOBS_ARRAY(six_month_total, this_job) + month_total
            this_job = this_job + 1
        End If
        budg_row = budg_row + 1
    Loop until inc_type = "__"

    number_of_months = number_of_months + 1
    col = col + 11
    transmit
    transmit
loop until col > 76

' For the_job = 0 to UBOUND(JOBS_ARRAY, 2)
'     MsgBox "The total income for 6 months for this job is $" & JOBS_ARRAY(six_month_total, the_job) & vbNewLine & "Number of months is " & number_of_months
' Next

y_pos = 60
BeginDialog average_income_dlg, 0, 0, 480, 105 + (UBOUND(JOBS_ARRAY, 2) *20), "Average Monthly JOBS Income"
  Text 10, 10, 95, 10, "This case is at " & case_status
  If case_status = "Application" Then Text 10, 25, 210, 10, "The date of application is " & application_date & " and the first month to FIAT is"
  If case_status = "Recertification" Then Text 10, 25, 210, 10, "The recertificiation is for " & MAXIS_footer_month & "/" & MAXIS_footer_year & " and the first month to FIAT is"
  EditBox 225, 20, 15, 15, MAXIS_footer_month
  EditBox 245, 20, 15, 15, MAXIS_footer_year
  Text 10, 45, 125, 10, "The script found the following  job(s):"
  For the_job = 0 to UBOUND(JOBS_ARRAY, 2)
      If JOBS_ARRAY(job_frequency, the_job) = "1" then JOBS_ARRAY(job_frequency, the_job) = "monthly"
      If JOBS_ARRAY(job_frequency, the_job) = "2" then JOBS_ARRAY(job_frequency, the_job) = "semi-monthly"
      If JOBS_ARRAY(job_frequency, the_job) = "3" then JOBS_ARRAY(job_frequency, the_job) = "biweekly"
      If JOBS_ARRAY(job_frequency, the_job) = "4" then JOBS_ARRAY(job_frequency, the_job) = "weekly"
      If JOBS_ARRAY(job_frequency, the_job) = "5" then JOBS_ARRAY(job_frequency, the_job) = "other"
    JOBS_ARRAY(average_monthly_inc, the_job) = FormatNumber(JOBS_ARRAY(six_month_total, the_job)/number_of_months, 2,,,0) & ""
    Text 20, y_pos + 5, 395, 10, JOBS_ARRAY(employer, the_job) & " - paid " & JOBS_ARRAY(job_frequency, the_job) & " on " & JOBS_ARRAY(pay_weekday, the_job) & " - total income for six-month budget period - $" & JOBS_ARRAY(six_month_total, the_job) & " - Average monthly income $"
    EditBox 415, y_pos, 55, 15, JOBS_ARRAY(average_monthly_inc, the_job)
    y_pos = y_pos + 20
  Next
  ButtonGroup ButtonPressed
    OkButton 370, y_pos + 5, 50, 15
    CancelButton 425, y_pos + 5, 50, 15
EndDialog

Do
    err_msg = ""
    Dialog average_income_dlg
    cancel_confirmation

    If trim(MAXIS_footer_month) = "" or trim(MAXIS_footer_year) = "" Then
        err_msg = err_msg & vbNewLine & "* Enter the footer month and year in which the FIATing should start."
        If case_status = "Application" Then err_msg = err_msg & vbNewLine & "  - This case is at application and most cases at application should be FIATed starting in the month of application."
        If case_status = "Recertification" Then err_msg = err_msg & vbNewLine & "  - This case is at recertification and most cases at recertification should be FIATed starting the first month of the next budget period."
    End If

    For the_job = 0 to UBOUND(JOBS_ARRAY, 2)
        If JOBS_ARRAY(average_monthly_inc, the_job) = "" Then err_msg = err_msg & vbNewLine & "* Enter the average monthly income for " & JOBS_ARRAY(employer, the_job) & "."
    Next

    If err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
Loop until err_msg = ""

Call back_to_SELF

start_month_and_year = MAXIS_footer_month & "/" & MAXIS_footer_year

'SECTION 04: NOW IT GOES TO ELIG/HC TO FIAT THE AMOUNTS
Call navigate_to_MAXIS_screen("ELIG", "HC__")

row = 1
col = 1
EMSearch memb_number & " ", row, col 'finding the member number
If row = 0 then script_end_procedure("Member number not found. You may have entered an incorrect member number on the first screen. Try the script again.")

EMWriteScreen "x", row, 26
transmit

row = 6
col = 1
EMSearch start_month_and_year, row, col

end_msg = "The selected month " & start_month_and_year & " is not in the current version of HC, review the month selected and try the script again."
If col = 0 Then script_end_procedure(end_msg)

PF9
EMReadScreen FIAT_check, 4, 24, 45
If FIAT_check <> "FIAT" then
  EMSendKey "05"
  transmit
End if

Do
    EMWriteScreen "x", 9, col + 2
    transmit
    EMWriteScreen "x", 13, 03
    transmit

    budg_row = 8
    For the_job = 0 to UBOUND(JOBS_ARRAY, 2)
        EmWriteScreen "___________", budg_row, 43
        EmWriteScreen JOBS_ARRAY(average_monthly_inc, the_job), budg_row, 43
        budg_row = budg_row + 1
    Next
    'MsgBox ("Budget updated.")
    col = col + 11
    transmit
    transmit
    transmit
loop until col > 76



script_end_procedure("Success! Please make sure to check eligibility for any Medicare savings programs such as QMB or SLMB.")
script_end_procedure("Did the FIATing Happen?")


MAXIS_footer_month = right("00" & MAXIS_footer_month, 2)
MAXIS_footer_year  = right("00" & MAXIS_footer_year, 2)
memb_number        = right("00" & memb_number, 2)

call navigate_to_MAXIS_screen("STAT", "JOBS")
EMReadScreen jobs_memb, 2, 4, 33  'checking if the current jobs panel is the memb, if not it will nav to member.
IF jobs_memb <> memb_number THEN
	EMWriteScreen memb_number, 20, 76
	transmit
END IF
EMReadScreen jobs_total, 1, 2, 78
EMReadScreen jobs_current, 1, 2, 73

If jobs_total = "0" then MsgBox "No JOBS panel is known for this client. You will have to enter income amounts manually."

If jobs_current = "1" then

    CALL get_average_pay(frequency_job_01, income_job_01)
  ' EMReadScreen pay_freq_01, 1, 18, 35
  ' If pay_freq_01 = "1" then frequency_job_01 = "1: monthly"
  ' If pay_freq_01 = "2" then frequency_job_01 = "2: twice monthly"
  ' If pay_freq_01 = "3" then frequency_job_01 = "3: every 2 weeks"
  ' If pay_freq_01 = "4" then frequency_job_01 = "4. every week"
  ' If pay_freq_01 = "5" then frequency_job_01 = "5. other (use monthly avg)"
  ' EMReadScreen HC_income_est_check, 3, 19, 63 'reading to find the HC income estimator is moving 6/1/16, to account for if it only affects future months we are reading to find the HC inc EST
  ' IF HC_income_est_check = "Est" Then 'this is the old position
	' EMWriteScreen "x", 19, 54
  ' ELSE								'this is the new position
	' EMWriteScreen "x", 19, 48
  ' END IF
  ' transmit
  ' EMReadScreen income_job_01, 8, 11, 63
  ' income_job_01 = trim(replace(income_job_01, "_", ""))
  ' transmit
  transmit
  EMReadScreen jobs_current, 1, 2, 73
End if

If jobs_current = "2" then
    CALL get_average_pay(frequency_job_02, income_job_02)
  ' EMReadScreen pay_freq_02, 1, 18, 35
  ' If pay_freq_02 = "1" then frequency_job_02 = "1: monthly"
  ' If pay_freq_02 = "2" then frequency_job_02 = "2: twice monthly"
  ' If pay_freq_02 = "3" then frequency_job_02 = "3: every 2 weeks"
  ' If pay_freq_02 = "4" then frequency_job_02 = "4. every week"
  ' If pay_freq_02 = "5" then frequency_job_02 = "5. other (use monthly avg)"
  ' EMReadScreen HC_income_est_check, 3, 19, 63 'reading to find the HC income estimator is moving 6/1/16, to account for if it only affects future months we are reading to find the HC inc EST
  ' IF HC_income_est_check = "Est" Then 'this is the old position
	' EMWriteScreen "x", 19, 54
  ' ELSE								'this is the new position
	' EMWriteScreen "x", 19, 48
  ' END IF
  ' transmit
  ' EMReadScreen income_job_02, 8, 11, 63
  ' income_job_02 = trim(replace(income_job_02, "_", ""))
  ' transmit
  transmit
  EMReadScreen jobs_current, 1, 2, 73
End if

If jobs_current = "3" then
    CALL get_average_pay(frequency_job_03, income_job_03)
  ' EMReadScreen pay_freq_03, 1, 18, 35
  ' If pay_freq_03 = "1" then frequency_job_03 = "1: monthly"
  ' If pay_freq_03 = "2" then frequency_job_03 = "2: twice monthly"
  ' If pay_freq_03 = "3" then frequency_job_03 = "3: every 2 weeks"
  ' If pay_freq_03 = "4" then frequency_job_03 = "4. every week"
  ' If pay_freq_03 = "5" then frequency_job_03 = "5. other (use monthly avg)"
  ' EMReadScreen HC_income_est_check, 3, 19, 63 'reading to find the HC income estimator is moving 6/1/16, to account for if it only affects future months we are reading to find the HC inc EST
  ' IF HC_income_est_check = "Est" Then 'this is the old position
	' EMWriteScreen "x", 19, 54
  ' ELSE								'this is the new position
	' EMWriteScreen "x", 19, 48
  ' END IF
  ' transmit
  ' EMReadScreen income_job_03, 8, 11, 63
  ' income_job_03 = trim(replace(income_job_03, "_", ""))
  ' transmit
  transmit
  EMReadScreen jobs_current, 1, 2, 73
End if

If income_job_01 = "" then
  income_job_01 = income_job_02
  frequency_job_01 = frequency_job_02
  income_job_02 = ""
  frequency_job_02 = ""
End if

If income_job_02 = "" then
  income_job_02 = income_job_03
  frequency_job_02 = frequency_job_03
  income_job_03 = ""
  frequency_job_03 = ""
End if

start_mo_month = MAXIS_footer_month & ""
start_mo_year = MAXIS_footer_year & ""

BeginDialog MA_EPD_dialog, 0, 0, 186, 140, "MA-EPD dialog"
  EditBox 30, 20, 40, 15, income_job_01
  DropListBox 85, 20, 90, 15, "1: monthly"+chr(9)+"2: twice monthly"+chr(9)+"3: every 2 weeks"+chr(9)+"4. every week"+chr(9)+"5. other (use monthly avg)", frequency_job_01
  EditBox 30, 40, 40, 15, income_job_02
  DropListBox 85, 40, 90, 15, "1: monthly"+chr(9)+"2: twice monthly"+chr(9)+"3: every 2 weeks"+chr(9)+"4. every week"+chr(9)+"5. other (use monthly avg)", frequency_job_02
  EditBox 30, 60, 40, 15, income_job_03
  DropListBox 85, 60, 90, 15, "1: monthly"+chr(9)+"2: twice monthly"+chr(9)+"3: every 2 weeks"+chr(9)+"4. every week"+chr(9)+"5. other (use monthly avg)", frequency_job_03
  EditBox 125, 80, 15, 15, start_mo_month
  EditBox 145, 80, 15, 15, start_mo_year
  ButtonGroup ButtonPressed
    OkButton 40, 120, 50, 15
    CancelButton 100, 120, 50, 15
  Text 35, 5, 40, 10, "Income amt"
  Text 115, 5, 30, 10, "Pay freq."
  Text 5, 25, 25, 10, "Job 1:"
  Text 5, 45, 25, 10, "Job 2:"
  Text 5, 65, 25, 10, "Job 3:"
  Text 5, 85, 110, 10, "Script will FIAT starting in month:"
  Text 10, 95, 115, 20, "The script will update this month and future months in ELIG."
EndDialog

Do
    err_msg = ""

    Dialog MA_EPD_dialog
    cancel_confirmation

    If start_mo_month = "" or start_mo_year = "" Then err_msg = err_msg & vbNewLine & "* Enter footer month and year."

    If err_msg <> "" Then MsgBox "Please resolve to cotinue:" & vbNewLine & err_msg
Loop until err_msg = ""

start_mo_month = right("00" & start_mo_month, 2)
start_mo_year = right("00" & start_mo_year, 2)

start_month_and_year = start_mo_month & "/" & start_mo_year

'SECTION 04: NOW IT GOES TO ELIG/HC TO FIAT THE AMOUNTS
Call navigate_to_MAXIS_screen("ELIG", "HC__")

row = 1
col = 1
EMSearch memb_number & " ", row, col 'finding the member number
If row = 0 then script_end_procedure("Member number not found. You may have entered an incorrect member number on the first screen. Try the script again.")

EMWriteScreen "x", row, 26
transmit

EMReadScreen elig_type_check_first_month, 2, 12, 17
EMReadScreen elig_type_check_second_month, 2, 12, 28
EMReadScreen elig_type_check_third_month, 2, 12, 39
EMReadScreen elig_type_check_fourth_month, 2, 12, 50
EMReadScreen elig_type_check_fifth_month, 2, 12, 61
EMReadScreen elig_type_check_sixth_month, 2, 12, 72

If elig_type_check_first_month <> "DP" and elig_type_check_second_month <> "DP" and elig_type_check_third_month <> "DP" and elig_type_check_fourth_month <> "DP" and elig_type_check_fifth_month <> "DP" and elig_type_check_sixth_month <> "DP" then MsgBox "Not all of the months of this case are MA-EPD. Process manually."
If elig_type_check_first_month <> "DP" and elig_type_check_second_month <> "DP" and elig_type_check_third_month <> "DP" and elig_type_check_fourth_month <> "DP" and elig_type_check_fifth_month <> "DP" and elig_type_check_sixth_month <> "DP" then stopscript

PF9
EMReadScreen FIAT_check, 4, 24, 45
If FIAT_check <> "FIAT" then
  EMSendKey "05"
  transmit
End if
' If radio1 = 1 then
'   row = 6
'   col = 1
'   EMSearch current_month_and_year, row, col
' End if
'
' If radio2 = 1 or row = 0 then
'   row = 6
'   col = 1
'   EMSearch next_month_and_year, row, col
' End if

row = 6
col = 1
EMSearch start_month_and_year, row, col

end_msg = "The selected month " & start_month_and_year & " is not in the current version of HC, review the month selected and try the script again."
If col = 0 Then script_end_procedure(end_msg)

'Multiplier calculations
If frequency_job_01 = "1: monthly" or frequency_job_01 = "5. other (use monthly avg)" then multiplier_01 = 1
If frequency_job_02 = "1: monthly" or frequency_job_02 = "5. other (use monthly avg)" then multiplier_02 = 1
If frequency_job_03 = "1: monthly" or frequency_job_03 = "5. other (use monthly avg)" then multiplier_03 = 1

If frequency_job_01 = "2: twice monthly" then multiplier_01 = 2
If frequency_job_02 = "2: twice monthly" then multiplier_02 = 2
If frequency_job_03 = "2: twice monthly" then multiplier_03 = 2

If frequency_job_01 = "3: every 2 weeks" then multiplier_01 = 2.16
If frequency_job_02 = "3: every 2 weeks" then multiplier_02 = 2.16
If frequency_job_03 = "3: every 2 weeks" then multiplier_03 = 2.16

If frequency_job_01 = "4. every week" then multiplier_01 = 4.3
If frequency_job_02 = "4. every week" then multiplier_02 = 4.3
If frequency_job_03 = "4. every week" then multiplier_03 = 4.3

Do
  EMWriteScreen "x", 9, col + 2
  transmit
  EMWriteScreen "x", 13, 03
  transmit
  EMWriteScreen "___________", 8, 43
  EMWriteScreen income_job_01 * multiplier_01, 8, 43
  If income_job_02 <> "" then
    EMWriteScreen "___________", 9, 43
    EMWriteScreen income_job_02 * multiplier_02, 9, 43
  End if
  If income_job_03 <> "" then
    EMWriteScreen "___________", 10, 43
    EMWriteScreen income_job_03 * multiplier_03, 10, 43
  End if
  col = col + 11
  transmit
  transmit
  transmit
loop until col > 76

script_end_procedure("Success! Please make sure to check eligibility for any Medicare savings programs such as QMB or SLMB.")
