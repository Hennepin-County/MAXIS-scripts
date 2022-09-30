'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - MA-EPD EI FIAT.vbs"
start_time = timer
STATS_counter = 0                     	'sets the stats counter at one
STATS_manualtime = 100                	'manual run time in seconds
STATS_denomination = "I"       		'I is for Item - this is each MONTH that is FIATed
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
call changelog_update("09/30/2022", "Added handling so the FIATer will work for budgets that are only partially MA-EPD. The script will recognize which months are MA-EPD and only average the income over the months that are coded 'DP'.##~##", "Casey Love, Henneppin County")
call changelog_update("05/25/2022", "Updated the handling of the Footer Month for more stable script operation. For active MA EPD cases (the Update process) the script will read JOBS and UNEA informaiton from Current Month plus One as that is the most recent information available.", "Casey Love, Hennepin County")
call changelog_update("06/25/2020", "Added handling to stop the script run if the income information is not updated fully or does not meet requiements for needing an MA-EPD FIAT.##~## If you have questions about use of the FIAT or necessary updates to JOBS and UNEA panels, please contact the BlueZone Script Team and we will direct you to the best resources.", "Casey Love, Hennepin County")
call changelog_update("06/04/2020", "BUG FIX - Script was failing with the new functionality trying to correctly navigate to UNEA panels. It was causing an error on some cases that prevented the script from continuing. Bug should now be resolved.##~##", "Casey Love, Hennepin County")
call changelog_update("05/22/2020", "Added functionality so the script can FIAT income from Unemployment as well as JOBS income. As UI income is received weekly, it can cause the premium to vary from month to month. This income also requires a FIAT to be balanced across the budget.##~## ##~## The functionality for UNEA panels coded with UI income works at the same time and in the same manner as the JOBS functionality.", "Casey Love, Hennepin County")
call changelog_update("11/27/2018", "Changed the case options to 'Initial' and 'Update' for the type of approval being made.", "Casey Love, Hennepin County")
call changelog_update("08/24/2018", "Fixed script to accommodate a $0 income job.", "Casey Love, Hennepin County")
call changelog_update("05/16/2018", "Added a place to input the footer month and year for the start of MA EPD.", "Casey Love, Hennepin County")
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
      EMWriteScreen "X", 19, 54
    ELSE								'this is the new position
      EMWriteScreen "X", 19, 48
    END IF
    transmit                            'opening the HC Inc Estimate pop-up
    EMReadScreen hc_inc_est, 8, 11, 63  'Reading the income on this field.'
    hc_inc_est = trim(replace(hc_inc_est, "_", "")) 'Fomatting the number'
    transmit                            'closing the HC Inc Est pop-up

    if hc_inc_est = "" Then hc_inc_est = 0      'Making this a number
    hc_inc_est = FormatNumber(hc_inc_est, 2)

    If hc_inc_est <> anticipated_average Then       'These two should be equal - because HC Inc Estimate is based on the average of pay
    'If they are not equal - script will ask the worker to clarify which is correct.
        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 221, 105, "Income Mismatch"
          OptionGroup RadioGroup1
            RadioButton 10, 30, 205, 10, "Use the amount from the HC Inc Est Pop-Up of $" & hc_inc_est, use_hc_est_inc_radio
            RadioButton 10, 45, 225, 10, "Use the amount from anticipated income on JOBS of $" & anticipated_average, use_anticipated_inc_radio
          ButtonGroup ButtonPressed
            OkButton 105, 85, 50, 15
            CancelButton 165, 85, 50, 15
          Text 5, 5, 210, 20, "It appears that the average income listed on this JOBS panel does not match. Please indicate which amount should be used."
          Text 5, 65, 190, 10, "These amounts are both average per pay period amounts."
        EndDialog

        Do
            Dialog Dialog1      'Running the dialog to ask for worker input on the correct income.
            Cancel_confirmation
            Call check_for_password(are_we_passworded_out)
        Loop until are_we_passworded_out = FALSE

        'This will set the average income for the job based on what the worker indicates
        If use_anticipated_inc_radio = 1 Then job_income = anticipated_average
        If use_hc_est_inc_radio = 1 Then job_income = hc_inc_est
    Else
        job_income = hc_inc_est     'If they are equal - this is just setting the income to the variable used later in the script
    End If

end function

'END FUNCTIONS==============================================================================================================

'THE SCRIPT--------------------------------
EMConnect ""
Call check_for_MAXIS(False)
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
        ma_status = trim(ma_status)             'cutting blank
        If ma_status = "ACTIVE" Then case_status = "Update"    'If a case is alread active, it is often at review'
        If ma_status = "PENDING" Then case_status = "Initial"       'If a case is pending then it is usually at Initial
    End If
End If

Call back_to_SELF
EMReadScreen MX_environment, 7, 22, 48
If MX_environment = "INQUIRY" Then script_end_procedure("FIATER scripts do not work in Inquiry. This is currently in inuiry, the script will now end. Switch to production and run the script again.")

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 161, 85, "Case number"
  EditBox 90, 5, 65, 15, MAXIS_case_number
  EditBox 90, 25, 30, 15, memb_number
  DropListBox 90, 45, 65, 45, "Select One..."+chr(9)+"Initial"+chr(9)+"Update", case_status
  ButtonGroup ButtonPressed
    OkButton 40, 65, 50, 15
    CancelButton 100, 65, 50, 15
  Text 5, 10, 80, 10, "Enter your case number:"
  Text 20, 30, 65, 10, "HH memb number:"
  Text 30, 50, 55, 10, "Approval will be "
EndDialog
'Running a dialog to get case number, member number and if the case is at Initial or Update.'
Do
    Do
        err_msg = ""

        Dialog Dialog1
        Cancel_confirmation

        If MAXIS_case_number = "" Then                                             err_msg = err_msg & vbNewLine & "* Enter a case number to continue."
        If IsNumeric(MAXIS_case_number) = FALSE or len(MAXIS_case_number) > 8 Then err_msg = err_msg & vbNewLine & "* Case number appears to be invalid. Check the case number and fix."
        If memb_number = "" Then                                                   err_msg = err_msg & vbNewLine & "* Enter a reference number for the member on MA-EPD."
        If case_status = "Select One..." Then                                      err_msg = err_msg & vbNewLine & "* Identify if approval is update or initial."
        'If MAXIS_footer_month = "" OR MAXIS_footer_year = "" Then                  err_msg = err_msg & vbNewLine & "* Enter the MAXIS footer month and year that has the best income information in it."

        If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

'Setting constants for the array of JOB information
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
const unea_type         = 16
const est_pop_up        = 17
const verif_code        = 18

Dim JOBS_ARRAY()                    'setting up the array
ReDim JOBS_ARRAY(verif_code, 0)

Dim UNEA_ARRAY()
ReDim UNEA_ARRAY(verif_code, 0)

Call navigate_to_MAXIS_screen_review_PRIV("STAT", "SUMM", is_this_priv)
If is_this_priv = True Then script_end_procedure("This case is privileged and the script cannot access the CASE. Reqquest access to the case and rerun the script.")
EMReadScreen pw_county_code, 2, 21, 19
If pw_county_code <> "27" Then script_end_procedure("This case is not in Hennepin County and cannot be updated. The script will now end.")

'Cases at Update have a different information to look at
If case_status = "Update" Then
    Call Navigate_to_MAXIS_screen("STAT", "REVW")       'Going to find the REVW month as that this the relevant JOBS information

    EMReadScreen hc_revw, 8, 9, 70                      'Reading the current REVW date
    hc_revw = replace(hc_revw, " ", "/")

    If DateDiff("D", CM_plus_1_mo & "/01/" & CM_plus_1_yr, hc_revw) > 0 Then        'If the case has been U coded, the review date will switch to the next review... in six months
        EMReadScreen hc_revw, 8, 11, 70                                             'So this will read the last review date that is listed below it.
        hc_revw = replace(hc_revw, " ", "/")
    End If

    If hc_revw <> "__/__/__" Then
        MAXIS_footer_month = DatePart("m", hc_revw)                     'Setting the dates to month and year variables in a 2 digit format
        MAXIS_footer_month = right("00" & MAXIS_footer_month, 2)

        MAXIS_footer_year = DatePart("yyyy", hc_revw)
        MAXIS_footer_year = right(MAXIS_footer_year, 2)
    Else
        MAXIS_footer_month = CM_mo
        MAXIS_footer_year = CM_yr
    End If

    Call back_to_SELF       'Getting out of STAT so that we can switch months if needed
End If

'For Initials, there are months with pay already put in it.
If case_status = "Initial" Then
    Call Navigate_to_MAXIS_screen("STAT", "HCRE")   'Going to HCRE to get information about when to start the FIATing

    hcre_row = 10       'Setting the row to find the correct member to read the application date and possible retro months
    Do
        EMReadScreen hcre_ref_numb, 2, hcre_row, 24     'reading the reference number
        If hcre_ref_numb = memb_number Then Exit Do     'Once the member number has been matched, this will exit to do because the row is set already and we will use the same row variable.

        hcre_row = hcre_row + 1         'Incrementing the row
        If hcre_row = 18 Then           'Scrolling through the list if needed
            PF20
            hcre_row = 10
        End If
        EMReadScreen next_client, 2, hcre_row, 24   'Finding the end of the list
    Loop until next_client = "  "

    EMReadScreen application_date, 8, hcre_row, 51      'reading the application date using the row found previously
    EMReadScreen coverage_date, 5, hcre_row, 64         'reading the coverage date to look for retro requests

    application_date = replace(application_date, " ", "/")      'making this variable actually a date

    MAXIS_footer_month = DatePart("m", application_date)        'Setting the footer month and year as 2 digit variables
    MAXIS_footer_month = right("00" & MAXIS_footer_month, 2)

    MAXIS_footer_year = DatePart("yyyy", application_date)
    MAXIS_footer_year = right(MAXIS_footer_year, 2)

    'if there is a retro request, a quick reminder that retro months have different budgeting processing.
    If left(coverage_date, 2) <> MAXIS_footer_month OR right(coverage_date, 2) <> MAXIS_footer_year Then
        coverage_date = replace(coverage_date, " ", "/")
        MsgBox "This case appears to have a retro request back to " & coverage_date & "." & vbNewLine & vbNewLine & "Retro months should not be FIATed to even the income out. The premium in these months are based on actual income and will be different."
    End If

    Call back_to_SELF       'Going out of STAT to switch months
End If

'Getting the footer month
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 196, 100, "Select Beginning of the Budget"
  EditBox 140, 5, 15, 15, MAXIS_footer_month
  EditBox 160, 5, 15, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 135, 75, 50, 15
  Text 5, 10, 120, 10, "Beginning month of MA-EPD budget"
  Text 5, 30, 185, 20, "Review the footer month listed here. This should be the first month of the budget period that you need to FIAT."
  Text 5, 55, 145, 20, "CRITICAL REVIEW - With the REVW Waiver for HC this date may be wrong."
EndDialog

Do
    Do
        err_msg = ""
        Dialog Dialog1
		cancel_without_confirmation

        If trim(MAXIS_footer_month) = "" or trim(MAXIS_footer_year) = "" Then err_msg = err_msg & vbNewLine & "* Enter the footer month and year."
        If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

MAXIS_footer_month = right("00"&MAXIS_footer_month, 2)
MAXIS_footer_year = right("00"&MAXIS_footer_year, 2)
script_run_lowdown = script_run_lowdown & vbCr & "Footer Month - " & MAXIS_footer_month & vbCr & "Footer Year - " & MAXIS_footer_year
Original_MAXIS_footer_month = MAXIS_footer_month
Original_MAXIS_footer_year = MAXIS_footer_year

Call Navigate_to_MAXIS_screen("STAT", "JOBS")       'Going to look at jobs
EMWriteScreen memb_number, 20, 76
EMWriteScreen "01", 20, 79
transmit

EMReadScreen number_of_jobs, 1, 2, 78               'Reading the number of JOBS panels for this client.
number_of_jobs = number_of_jobs * 1

Call Navigate_to_MAXIS_screen("STAT", "UNEA")       'Going to look at jobs
EMWriteScreen memb_number, 20, 76
EMWriteScreen "01", 20, 79
transmit

EMReadScreen number_of_unea, 1, 2, 78               'Reading the number of JOBS panels for this client.
number_of_unea = number_of_unea * 1

list_of_unea_income_to_fiat = ""
For each_unea = 1 to number_of_unea
    each_unea = "0" & each_unea                     'Navigate to each of the UNEA panels to see if they are Unemployment
    EMWriteScreen each_unea, 20, 79
    transmit

    EMReadScreen income_type, 2, 5, 37              'Reading the income type
    'Only used for Unemployment Income
    If income_type = "14" Then list_of_unea_income_to_fiat = list_of_unea_income_to_fiat & "~" & each_unea     'saving the panel instance in a list if paid weekly or biweekly
Next

'If no jobs, there is no income to FIAT and script will end.
If number_of_jobs = 0 AND list_of_unea_income_to_fiat = "" Then
    end_msg = "Household Member " & member_number & " on this case has no JOBS panel and no UI UNEA panel. Please check the case, update JOBS if required and run the script again."
    script_end_procedure(end_msg)
End If

If list_of_unea_income_to_fiat <> "" Then           'If there was a panel of unemployment income
    If left(list_of_unea_income_to_fiat, 1) = "~" Then list_of_unea_income_to_fiat = right(list_of_unea_income_to_fiat, len(list_of_unea_income_to_fiat) - 1)
    If InStr(list_of_unea_income_to_fiat, "~") = 0 Then
        unea_panels_array = ARRAY(list_of_unea_income_to_fiat)
    Else
        unea_panels_array = split(list_of_unea_income_to_fiat, "~")
    End If
Else
    unea_panels_array = ARRAY("")
End If

If case_status = "Update" Then
	MAXIS_footer_month = CM_plus_1_mo
	MAXIS_footer_year = CM_plus_1_yr
    Call Navigate_to_MAXIS_screen("STAT", "JOBS")       'Going back to look at jobs
    EMWriteScreen memb_number, 20, 76
    EMWriteScreen "01", 20, 79
    transmit
    For each_job = 1 to number_of_jobs              'This will loop through each of the jobs
        EMReadScreen job_verification, 1, 6, 34     'reading information that should be updated for the reveiw to be processed
        EMReadScreen first_check_month, 2, 12, 54

        'If these have not been updated then the script will end because STAT needs to be updated first
        end_msg = "It does not appear this JOBS panel has been updated with income information for the review." & vbNewLine & vbNewLine & "If this job has ended and has no income in this month, STWK should be updated and this JOBS panel deleted." & vbNewLine & vbNewLine & "Cases should be fully processed prior to fiating eligibility results."
        If job_verification = "?" OR first_check_month <> MAXIS_footer_month Then script_end_procedure(end_msg)

        reDim Preserve JOBS_ARRAY(verif_code, each_job-1)   'Updating the array with JOB information

        'Gathering data and adding it to the array
        EMReadScreen verification, 25, 6, 34
        EMReadScreen freq, 1, 18, 35
        EMReadScreen title, 30, 7, 42

        JOBS_ARRAY(verif_code, each_job-1) = trim(verification)
        JOBS_ARRAY(job_frequency, each_job-1) = freq
        JOBS_ARRAY(instance, each_job-1) = right("00"&each_job, 2)
        JOBS_ARRAY(employer, each_job-1) = replace(title, "_", "")

        JOBS_ARRAY(six_month_total, each_job-1) = 0     'setting this as 0 because it will be added to later
        EMReadScreen pay_date, 8, 12, 54

        pay_date = replace(pay_date, " ", "/")

        'Setting the pay day
        If JOBS_ARRAY(job_frequency, each_job-1) = "3" OR JOBS_ARRAY(job_frequency, each_job-1) = "4" Then JOBS_ARRAY(pay_weekday, each_job-1) = WeekDayName(WeekDay(pay_date))

        'Looking at the pop-up for income information
        EMReadScreen HC_income_est_check, 3, 19, 63 'reading to find the HC income estimator is moving 6/1/16, to account for if it only affects future months we are reading to find the HC inc EST
        IF HC_income_est_check = "Est" Then 'this is the old position
          EMWriteScreen "X", 19, 54
        ELSE								'this is the new position
          EMWriteScreen "X", 19, 48
        END IF
        transmit
        EMReadScreen hc_inc_est, 8, 11, 63
        hc_inc_est = trim(replace(hc_inc_est, "_", ""))
        transmit

        if hc_inc_est = "" Then hc_inc_est = 0
        hc_inc_est = FormatNumber(hc_inc_est, 2)        'This formats the number with 2 decimal places

        JOBS_ARRAY(est_pop_up, each_job-1) = hc_inc_est

        transmit    'Going to the next JOBS panel
    Next

    Call Navigate_to_MAXIS_screen("STAT", "UNEA")       'Going back to UNEA
    EMWriteScreen memb_number, 20, 76
    transmit
    counter = 0
    For each each_unea in unea_panels_array
		each_unea = trim(each_unea)
        If each_unea <> "" Then
            EMWriteScreen each_unea, 20, 79
            transmit

            EMReadScreen unea_verification, 1, 5, 65     'reading information that should be updated for the reveiw to be processed
            EMReadScreen first_check_month, 2, 13, 54

            'If these have not been updated then the script will end because STAT needs to be updated first
            end_msg = "It does not appear this UNEA panel has been updated with income information for the review." & vbNewLine & vbNewLine & "If this unea has ended and has no income in this month, the information should be noted and the panel deleted." & vbNewLine & vbNewLine & "Cases should be fully processed prior to fiating eligibility results."
            If unea_verification = "?" OR first_check_month <> MAXIS_footer_month Then script_end_procedure(end_msg)

            reDim Preserve UNEA_ARRAY(verif_code, counter)   'Updating the array with JOB information

            'Gathering data and adding it to the array
            EMReadScreen verification, 16, 5, 65
            EMReadScreen income_source, 2, 5, 37

            UNEA_ARRAY(verif_code, counter) = trim(verification)
            UNEA_ARRAY(instance, counter) = right("00"&each_unea, 2)
            UNEA_ARRAY(unea_type, counter) = "Unemployment Insurance"

            UNEA_ARRAY(six_month_total, counter) = 0     'setting this as 0 because it will be added to later
            EMReadScreen pay_date, 8, 13, 54
            pay_date = replace(pay_date, " ", "/")

            'Looking at the pop-up for income information
            EMWriteScreen "X", 6, 56
            transmit
            EMReadScreen hc_inc_est, 8, 9, 65
            hc_inc_est = trim(replace(hc_inc_est, "_", ""))
            transmit

            if hc_inc_est = "" Then hc_inc_est = 0
            hc_inc_est = FormatNumber(hc_inc_est, 2)        'This formats the number with 2 decimal places

            UNEA_ARRAY(est_pop_up, counter) = hc_inc_est
            UNEA_ARRAY(job_frequency, counter) = "4"
            'Setting the pay day
            If UNEA_ARRAY(job_frequency, counter) = "3" OR UNEA_ARRAY(job_frequency, counter) = "4" Then UNEA_ARRAY(pay_weekday, counter) = WeekDayName(WeekDay(pay_date))
            counter = counter + 1
        End If
    Next
	MAXIS_footer_month = Original_MAXIS_footer_month
	MAXIS_footer_year = Original_MAXIS_footer_year
End If

If case_status = "Initial" Then
    Call Navigate_to_MAXIS_screen("STAT", "JOBS")       'Going back to look at jobs
    EMWriteScreen memb_number, 20, 76
    EMWriteScreen "01", 20, 79
    transmit
    'reading each JOBS panel and adding the information to the array
    For each_job = 1 to number_of_jobs
        reDim Preserve JOBS_ARRAY(verif_code, each_job-1)       'resizing the array

        'reading information from the panel and adding it to the array
        EMReadScreen verification, 25, 6, 34
        EMReadScreen freq, 1, 18, 35
        EMReadScreen title, 30, 7, 42

        JOBS_ARRAY(verif_code, each_job-1) = trim(verification)
        JOBS_ARRAY(job_frequency, each_job-1) = freq
        JOBS_ARRAY(instance, each_job-1) = right("00"&each_job, 2)
        JOBS_ARRAY(employer, each_job-1) = replace(title, "_", "")

        jobs_row = 12           'this is where the the prospective pay information starts
        divider = 0             'this is going to count the number of paychecks
        Do
            EMReadScreen pay_date, 8, jobs_row, 54      'reading the date
            If pay_date <> "__ __ __" Then              'if the date is not blank then gathering all the information
                                                        'this will read each already stored information in the array and will find the first empty position within the array
                divider = divider + 1                                                       'increase the count of checks listed
                If JOBS_ARRAY(check_date_one, each_job-1) = "" Then                         'this reads if this position in the array already has data
                    JOBS_ARRAY(check_date_one, each_job-1) = replace(pay_date, " ", "/")    'formatting the date
                    EMReadScreen pay_amt, 8, jobs_row, 67                                   'reading the pay amount and then formats it
                    JOBS_ARRAY(check_amt_one, each_job-1) = trim(pay_amt) * 1

                ElseIf JOBS_ARRAY(check_date_two, each_job-1) = "" Then                         'this reads if this position in the array already has data
                    JOBS_ARRAY(check_date_two, each_job-1) = replace(pay_date, " ", "/")        'formatting the date
                    EMReadScreen pay_amt, 8, jobs_row, 67                                       'reading the pay amount and then formats it
                    JOBS_ARRAY(check_amt_two, each_job-1) = trim(pay_amt) * 1

                ElseIf JOBS_ARRAY(check_date_three, each_job-1) = "" Then                       'this reads if this position in the array already has data
                    JOBS_ARRAY(check_date_three, each_job-1) = replace(pay_date, " ", "/")      'formatting the date
                    EMReadScreen pay_amt, 8, jobs_row, 67                                       'reading the pay amount and then formats it
                    total_prosp = total_prosp + trim(pay_amt) * 1
                    JOBS_ARRAY(check_amt_three, each_job-1) = trim(pay_amt) * 1

                ElseIf JOBS_ARRAY(check_date_four, each_job-1) = "" Then                        'this reads if this position in the array already has data
                    JOBS_ARRAY(check_date_four, each_job-1) = replace(pay_date, " ", "/")       'formatting the date
                    EMReadScreen pay_amt, 8, jobs_row, 67                                       'reading the pay amount and then formats it
                    JOBS_ARRAY(check_amt_four, each_job-1) = trim(pay_amt) * 1

                ElseIf JOBS_ARRAY(check_date_five, each_job-1) = "" Then                        'this reads if this position in the array already has data
                    JOBS_ARRAY(check_date_five, each_job-1) = replace(pay_date, " ", "/")       'formatting the date
                    EMReadScreen pay_amt, 8, jobs_row, 67                                       'reading the pay amount and then formats it
                    JOBS_ARRAY(check_amt_five, each_job-1) = trim(pay_amt) * 1
                End If
            End If

            jobs_row = jobs_row + 1         'going to the next row in the list of paychecks
        Loop until jobs_row = 17

        EMReadScreen total_pay, 8, 17, 67   'reading the total of the pay listed in the prospective side of the JOBS panel and formatting
        total_pay = trim(total_pay)
        If total_pay = "" Then total_pay = 0
        total_pay = total_pay * 1

        JOBS_ARRAY(pay_average, each_job-1) = total_pay / divider       'Finding the average of all of the paychecks listed
        JOBS_ARRAY(six_month_total, each_job-1) = 0                     'setting this equal to 0 as we will be adding to it later
        day_validation_needed = FALSE       'default for this variable
        If JOBS_ARRAY(job_frequency, each_job-1) = "3" OR JOBS_ARRAY(job_frequency, each_job-1) = "4" Then  'if this JOB is paid weekly or biweekly
            JOBS_ARRAY(pay_weekday, each_job-1) = WeekDayName(WeekDay(JOBS_ARRAY(check_date_one, each_job-1)))  'finding the day of the week of the first paycheck
            If JOBS_ARRAY(check_date_two, each_job-1) <> "" Then        'if this paycheck was found, it will find the day of the week paycheck and compare it to the first, if they do not match - validation is needed
                If WeekDayName(WeekDay(JOBS_ARRAY(check_date_two, each_job-1))) <> JOBS_ARRAY(pay_weekday, each_job-1) Then day_validation_needed = TRUE
            End If
            If JOBS_ARRAY(check_date_three, each_job-1) <> "" Then        'if this paycheck was found, it will find the day of the week paycheck and compare it to the first, if they do not match - validation is needed
                If WeekDayName(WeekDay(JOBS_ARRAY(check_date_three, each_job-1))) <> JOBS_ARRAY(pay_weekday, each_job-1) Then day_validation_needed = TRUE
            End If
            If JOBS_ARRAY(check_date_four, each_job-1) <> "" Then        'if this paycheck was found, it will find the day of the week paycheck and compare it to the first, if they do not match - validation is needed
                If WeekDayName(WeekDay(JOBS_ARRAY(check_date_four, each_job-1))) <> JOBS_ARRAY(pay_weekday, each_job-1) Then day_validation_needed = TRUE
            End If
            If JOBS_ARRAY(check_date_five, each_job-1) <> "" Then        'if this paycheck was found, it will find the day of the week paycheck and compare it to the first, if they do not match - validation is needed
                If WeekDayName(WeekDay(JOBS_ARRAY(check_date_five, each_job-1))) <> JOBS_ARRAY(pay_weekday, each_job-1) Then day_validation_needed = TRUE
            End If
        End If

        If day_validation_needed = TRUE Then        'If any of the paychecks listed do not match the first, then worker needs to identify the correct pay day
            selected_weekday = JOBS_ARRAY(pay_weekday, each_job-1)

            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 161, 80, "Weekday"
              DropListBox 15, 55, 60, 45, "Sunday"+chr(9)+"Monday"+chr(9)+"Tuesday"+chr(9)+"Wednesday"+chr(9)+"Thursday"+chr(9)+"Friday"+chr(9)+"Saturday", selected_weekday
              ButtonGroup ButtonPressed
                OkButton 105, 55, 50, 15
              Text 5, 10, 150, 35, "This job is paid either weekly or biweekly, but has different weekdays indicated for pay dates. Please select the weekday that the client is paid."
            EndDialog

            Do
                Dialog Dialog1
                Call check_for_password(are_we_passworded_out)
            Loop until are_we_passworded_out = FALSE

            JOBS_ARRAY(pay_weekday, each_job-1) = selected_weekday

        End If

        total_pay = FormatNumber(total_pay, 2)
        JOBS_ARRAY(est_pop_up, each_job-1) = total_pay

        transmit    'Giong to the next JOBS panel
    Next

    Call Navigate_to_MAXIS_screen("STAT", "UNEA")       'Going back to UNEA
    EMWriteScreen memb_number, 20, 76
    transmit
    counter = 0
    If IsArray(unea_panels_array) = TRUE Then
        For each each_unea in unea_panels_array
			If each_unea <> "" Then
	            EMWriteScreen each_unea, 20, 79
	            transmit

	            reDim Preserve UNEA_ARRAY(verif_code, counter)   'Updating the array with JOB information

	            'Gathering data and adding it to the array
	            EMReadScreen verification, 16, 5, 65
	            EMReadScreen income_source, 2, 5, 37

	            UNEA_ARRAY(verif_code, counter) = trim(verification)
	            UNEA_ARRAY(instance, counter) = right("00"&each_unea, 2)
	            UNEA_ARRAY(unea_type, counter) = "Unemployment Insurance"

	            'Looking at the pop-up for income information
	            EMWriteScreen "X", 6, 56
	            transmit
	            EMReadScreen hc_inc_est, 8, 9, 65
	            hc_inc_est = trim(replace(hc_inc_est, "_", ""))
	            transmit

	            if hc_inc_est = "" Then hc_inc_est = 0
	            hc_inc_est = FormatNumber(hc_inc_est, 2)        'This formats the number with 2 decimal places

	            UNEA_ARRAY(est_pop_up, counter) = hc_inc_est
	            UNEA_ARRAY(job_frequency, counter) = "4"

	            unea_row = 13
	            divider = 0
	            Do
	                EMReadScreen pay_date, 8, unea_row, 54      'reading the date
	                If pay_date <> "__ __ __" Then              'if the date is not blank then gathering all the information
	                                                            'this will read each already stored information in the array and will find the first empty position within the array
	                    divider = divider + 1                                                       'increase the count of checks listed
	                    If UNEA_ARRAY(check_date_one, counter) = "" Then                         'this reads if this position in the array already has data
	                        UNEA_ARRAY(check_date_one, counter) = replace(pay_date, " ", "/")    'formatting the date
	                        EMReadScreen pay_amt, 8, unea_row, 68                                   'reading the pay amount and then formats it
	                        UNEA_ARRAY(check_amt_one, counter) = trim(pay_amt) * 1

	                    ElseIf UNEA_ARRAY(check_date_two, counter) = "" Then                         'this reads if this position in the array already has data
	                        UNEA_ARRAY(check_date_two, counter) = replace(pay_date, " ", "/")        'formatting the date
	                        EMReadScreen pay_amt, 8, unea_row, 68                                       'reading the pay amount and then formats it
	                        UNEA_ARRAY(check_amt_two, counter) = trim(pay_amt) * 1

	                    ElseIf UNEA_ARRAY(check_date_three, counter) = "" Then                       'this reads if this position in the array already has data
	                        UNEA_ARRAY(check_date_three, counter) = replace(pay_date, " ", "/")      'formatting the date
	                        EMReadScreen pay_amt, 8, unea_row, 68                                       'reading the pay amount and then formats it
	                        total_prosp = total_prosp + trim(pay_amt) * 1
	                        UNEA_ARRAY(check_amt_three, counter) = trim(pay_amt) * 1

	                    ElseIf UNEA_ARRAY(check_date_four, counter) = "" Then                        'this reads if this position in the array already has data
	                        UNEA_ARRAY(check_date_four, counter) = replace(pay_date, " ", "/")       'formatting the date
	                        EMReadScreen pay_amt, 8, unea_row, 68                                       'reading the pay amount and then formats it
	                        UNEA_ARRAY(check_amt_four, counter) = trim(pay_amt) * 1

	                    ElseIf UNEA_ARRAY(check_date_five, counter) = "" Then                        'this reads if this position in the array already has data
	                        UNEA_ARRAY(check_date_five, counter) = replace(pay_date, " ", "/")       'formatting the date
	                        EMReadScreen pay_amt, 8, unea_row, 68                                       'reading the pay amount and then formats it
	                        UNEA_ARRAY(check_amt_five, counter) = trim(pay_amt) * 1
	                    End If
	                End If

	                unea_row = unea_row + 1         'going to the next row in the list of paychecks
	            Loop until unea_row = 18

	            EMReadScreen total_pay, 8, 18, 68   'reading the total of the pay listed in the prospective side of the JOBS panel and formatting
	            total_pay = trim(total_pay)
	            If total_pay = "" Then total_pay = 0
	            total_pay = total_pay * 1

	            UNEA_ARRAY(pay_average, counter) = total_pay / divider       'Finding the average of all of the paychecks listed
	            UNEA_ARRAY(six_month_total, counter) = 0                     'setting this equal to 0 as we will be adding to it later
	            day_validation_needed = FALSE       'default for this variable
	            If UNEA_ARRAY(job_frequency, counter) = "3" OR UNEA_ARRAY(job_frequency, counter) = "4" Then  'if this JOB is paid weekly or biweekly
	                UNEA_ARRAY(pay_weekday, counter) = WeekDayName(WeekDay(UNEA_ARRAY(check_date_one, counter)))  'finding the day of the week of the first paycheck
	                If UNEA_ARRAY(check_date_two, counter) <> "" Then        'if this paycheck was found, it will find the day of the week paycheck and compare it to the first, if they do not match - validation is needed
	                    If WeekDayName(WeekDay(UNEA_ARRAY(check_date_two, counter))) <> UNEA_ARRAY(pay_weekday, counter) Then day_validation_needed = TRUE
	                End If
	                If UNEA_ARRAY(check_date_three, counter) <> "" Then        'if this paycheck was found, it will find the day of the week paycheck and compare it to the first, if they do not match - validation is needed
	                    If WeekDayName(WeekDay(UNEA_ARRAY(check_date_three, counter))) <> UNEA_ARRAY(pay_weekday, counter) Then day_validation_needed = TRUE
	                End If
	                If UNEA_ARRAY(check_date_four, counter) <> "" Then        'if this paycheck was found, it will find the day of the week paycheck and compare it to the first, if they do not match - validation is needed
	                    If WeekDayName(WeekDay(UNEA_ARRAY(check_date_four, counter))) <> UNEA_ARRAY(pay_weekday, counter) Then day_validation_needed = TRUE
	                End If
	                If UNEA_ARRAY(check_date_five, counter) <> "" Then        'if this paycheck was found, it will find the day of the week paycheck and compare it to the first, if they do not match - validation is needed
	                    If WeekDayName(WeekDay(UNEA_ARRAY(check_date_five, counter))) <> UNEA_ARRAY(pay_weekday, counter) Then day_validation_needed = TRUE
	                End If
	            End If

	            If day_validation_needed = TRUE Then        'If any of the paychecks listed do not match the first, then worker needs to identify the correct pay day
	                selected_weekday = UNEA_ARRAY(pay_weekday, counter)

	                Dialog1 = ""
	                BeginDialog Dialog1, 0, 0, 161, 80, "Weekday"
	                  DropListBox 15, 55, 60, 45, "Sunday"+chr(9)+"Monday"+chr(9)+"Tuesday"+chr(9)+"Wednesday"+chr(9)+"Thursday"+chr(9)+"Friday"+chr(9)+"Saturday", selected_weekday
	                  ButtonGroup ButtonPressed
	                    OkButton 105, 55, 50, 15
	                  Text 5, 10, 150, 35, "This unearned income is paid either weekly or biweekly, but has different weekdays indicated for pay dates. Please select the weekday that the client is paid."
	                EndDialog

	                Do
	                    Dialog Dialog1
	                    Call check_for_password(are_we_passworded_out)
	                Loop until are_we_passworded_out = FALSE

	                UNEA_ARRAY(pay_weekday, counter) = selected_weekday

	            End If

	            total_pay = FormatNumber(total_pay, 2)
	            UNEA_ARRAY(est_pop_up, counter) = total_pay
	            counter = counter + 1
			End If
        Next
    End If

    app_month = MAXIS_footer_month      'saving the footer month in a seperate variable because we need to navigate to other months
    app_year = MAXIS_footer_year

    first_of_this_month = MAXIS_footer_month & "/1/" & MAXIS_footer_year    'setting to a date so we can use date functionality
    next_month = DateAdd("m", 1, first_of_this_month)       'going to the next month

    Do
        next_month_mo = DatePart("m", next_month)       'finding the next month and creating a 2 digit variable for the month and year
        next_month_mo = right("00"&next_month_mo, 2)
        next_month_yr = DatePart("yyyy", next_month)
        next_month_yr = right(next_month_yr, 2)

        list_of_months = list_of_months & "~" & next_month_mo & "/" & next_month_yr     'creating a list of month/years that need to be looked at

        first_of_this_month = next_month_mo & "/1/" & next_month_yr     'and to the next month
        next_month = DateAdd("m", 1, first_of_this_month)
    Loop until DateDiff("d", date, next_month) > 0      'doing this until the next month variable is after the current date

    list_of_months = right(list_of_months, len(list_of_months)-1)   'creating an array of all the months to check
    month_array = split(list_of_months, "~")

    paychecks_align_with_weekday = TRUE             'setting the default of this variable
    For each footer in month_array                  'we are going to each month
        Call back_to_SELF                           'need to back out to SELF so we can change months

        MAXIS_footer_month = left(footer, 2)        'setting the footer month and year to these global variables because those are used for the navigation functions
        MAXIS_footer_year = right(footer, 2)

        If MAXIS_footer_month = CM_plus_2_mo AND MAXIS_footer_year = CM_plus_2_yr Then Exit For

        Call Navigate_to_MAXIS_screen("STAT", "JOBS")   'going to JOBS in that month for the member
        EmWriteScreen member_number, 20, 76
        transmit

        For the_job = 0 to UBOUND(JOBS_ARRAY, 2)        'Now this is going to each of the JOBS previously found
            If JOBS_ARRAY(job_frequency, the_job) = "3" OR JOBS_ARRAY(job_frequency, the_job) = "4" Then    'If the pay is weekly or biweekly
                EmWriteScreen JOBS_ARRAY(instance, the_job), 20, 79     'navigating to the right instance of the JOB
                transmit

                jobs_row = 12       'Setting this to the beginning of the list of paychecks
                Do
                    EMReadScreen pay_date, 8, jobs_row, 54          'read the pay date and make it a date if not blank
                    If pay_date <> "__ __ __" Then
                        pay_date = replace(pay_date, " ", "/")
                        day_of_pay = WeekDayName(Weekday(pay_date))     'finding the weekday of this pay check

                        If day_of_pay <> JOBS_ARRAY(pay_weekday, the_job) Then      'the weekday paid should match the weekday already determined when reading the JOBS panel

                            'This message will have the worker confirm that the JOBS panel has the correct pay dates.
                            'In some instances there is a reason why a check date does not align with the rest of the pay dates, but if the check dates have not been properly updated in each month, the messagebox comes up for EVERY MONTH - encouraging correction
                            confirm_off_schedule_pay = MsgBox("The pay listed on this JOBS panel does not match the day of the week pay was received in the initial month of applicaton." & vbNewLine & vbNewLine &_
                             "Pay date of " & pay_date & " listed is on a " & day_of_pay & "." & vbNewLine & "This job appears to have a regular pay date of " & JOBS_ARRAY(pay_weekday, this_job) & "." & vbNewLine & vbNewLine &_
                             "Health care budget requires any income entered on JOBS be the actual pay dates expected, even if the income is calculated by average pay. Review the case and make sure that check dates have been updated in every month." & vbNewLine & vbNewLine &_
                             "Has the budget been correctly determined, using actual pay dates for each month that can be updated?", vbYesNo + vbImportant, "Confirm paycheck budgeting")

                            if confirm_off_schedule_pay = vbNo Then script_end_procedure("Update STAT/JOBS with all actual pay dates to get a correct HC budget.")  'If the paychecks or not complete the script will end
                        End If
                    End If
                    jobs_row = jobs_row + 1 'looking at the next check
                Loop until jobs_row = 17
            End If
        Next

        Call Navigate_to_MAXIS_screen("STAT", "UNEA")   'going to JOBS in that month for the member
        EmWriteScreen member_number, 20, 76
        transmit

        For the_unea = 0 to UBOUND(UNEA_ARRAY, 2)        'Now this is going to each of the JOBS previously found
			If UNEA_ARRAY(unea_type, 0) <> "" Then
	            If UNEA_ARRAY(job_frequency, the_unea) = "3" OR UNEA_ARRAY(job_frequency, the_unea) = "4" Then    'If the pay is weekly or biweekly
	                EmWriteScreen UNEA_ARRAY(instance, the_unea), 20, 79     'navigating to the right instance of the JOB
	                transmit

	                unea_row = 13       'Setting this to the beginning of the list of paychecks
	                Do
	                    EMReadScreen pay_date, 8, unea_row, 54          'read the pay date and make it a date if not blank
	                    If pay_date <> "__ __ __" Then
	                        pay_date = replace(pay_date, " ", "/")
	                        day_of_pay = WeekDayName(Weekday(pay_date))     'finding the weekday of this pay check

	                        If day_of_pay <> UNEA_ARRAY(pay_weekday, the_unea) Then      'the weekday paid should match the weekday already determined when reading the JOBS panel

	                            'This message will have the worker confirm that the JOBS panel has the correct pay dates.
	                            'In some instances there is a reason why a check date does not align with the rest of the pay dates, but if the check dates have not been properly updated in each month, the messagebox comes up for EVERY MONTH - encouraging correction
	                            confirm_off_schedule_pay = MsgBox("The pay listed on this UNEA panel does not match the day of the week pay was received in the initial month of applicaton." & vbNewLine & vbNewLine &_
	                             "Pay date of " & pay_date & " listed is on a " & day_of_pay & "." & vbNewLine & "This income appears to have a regular pay date of " & UNEA_ARRAY(pay_weekday, the_unea) & "." & vbNewLine & vbNewLine &_
	                             "Health care budget requires any income entered on UNEA be the actual pay dates expected, even if the income is calculated by average pay. Review the case and make sure that check dates have been updated in every month." & vbNewLine & vbNewLine &_
	                             "Has the budget been correctly determined, using actual pay dates for each month that can be updated?", vbYesNo + vbImportant, "Confirm paycheck budgeting")

	                            if confirm_off_schedule_pay = vbNo Then script_end_procedure("Update STAT/UNEA with all actual pay dates to get a correct HC budget.")  'If the paychecks or not complete the script will end
	                        End If
	                    End If
	                    unea_row = unea_row + 1 'looking at the next check
	                Loop until unea_row = 18
	            End If
			End If
        Next
    Next

    'Resetting the footer month and year to what it was saved at before looking at each month
    MAXIS_footer_month = app_month
    MAXIS_footer_year = app_year
End If

'This variable will be used to find the first footer month
start_month_and_year = MAXIS_footer_month & "/" & MAXIS_footer_year

Call navigate_to_MAXIS_screen("ELIG", "HC__")       'Going to ELIG/HC

'Finding the member in the list of all members on ELIG/HC
row = 1
col = 1
EMSearch memb_number & " ", row, col 'finding the member number
If row = 0 then script_end_procedure("Member number not found. You may have entered an incorrect member number on the first screen. Try the script again.")

'Opening the eligibility span of the client
EMWriteScreen "X", row, 26
transmit

'Reading the elig type of all the months. They should be DP because that is MA-EPD
EMReadScreen elig_type_check_first_month, 2, 12, 17
EMReadScreen elig_type_check_second_month, 2, 12, 28
EMReadScreen elig_type_check_third_month, 2, 12, 39
EMReadScreen elig_type_check_fourth_month, 2, 12, 50
EMReadScreen elig_type_check_fifth_month, 2, 12, 61
EMReadScreen elig_type_check_sixth_month, 2, 12, 72

'There needs to be at least 1 month of MA-EPD Elig Results - the script will check each month and if all are NOT DP - the script will end.
If elig_type_check_first_month <> "DP" and elig_type_check_second_month <> "DP" and elig_type_check_third_month <> "DP" and elig_type_check_fourth_month <> "DP" and elig_type_check_fifth_month <> "DP" and elig_type_check_sixth_month <> "DP" then script_end_procedure("Not all of the months of this case are MA-EPD. Process manually.")

'This is determining the first month that MA-EPD is coded in the budget.
If elig_type_check_first_month = "DP" Then
	EMReadScreen first_month_and_year, 5, 6, 19
ElseIf elig_type_check_second_month = "DP" Then
	EMReadScreen first_month_and_year, 5, 6, 30
ElseIf elig_type_check_third_month = "DP" Then
	EMReadScreen first_month_and_year, 5, 6, 41
ElseIf elig_type_check_fourth_month = "DP" Then
	EMReadScreen first_month_and_year, 5, 6, 52
ElseIf elig_type_check_fifth_month = "DP" Then
	EMReadScreen first_month_and_year, 5, 6, 63
ElseIf elig_type_check_sixth_month = "DP" then
	EMReadScreen first_month_and_year, 5, 6, 74
End If
'Worker should have indicated the first month of MA - EPD here. If that was not accurate, the STAT information already gathered may be incorrect
'the script will end if the entered month and found month do not match
If start_month_and_year <> first_month_and_year Then
	end_msg = "The Beginning Month of MA-EPD that you entered into the start of the script run: " & start_month_and_year & " does not appear to be the first month of MA-EPD in this HC Budget." & vbCr & vbCr & "It appears the first month of MA-EPD is " & first_month_and_year & "." & vbCr & vbCr & "Please review the case and rerun the script, entering the correct Beginning Month of MA-EPD."
	call script_end_procedure_with_error_report(end_msg)
End If

'Looking for the first month to FIAT
row = 6
col = 1
EMSearch first_month_and_year, row, col

end_msg = "The selected month " & first_month_and_year & " is not in the current version of HC, review the month selected and try the script again."
If col = 0 Then script_end_procedure(end_msg)

'Now looking at each month in ELIG
number_of_months = 0        'setting this at - it will count the number of months to be FIATed
Do
	EMReadScreen elig_type_check, 2, 12, col-2	'ensuring the month is MA-EPD elig type'
	If elig_type_check = "DP" Then
	    EMWriteScreen "X", 9, col + 2       'opening the Budget for the month
	    transmit
	    EMWriteScreen "X", 13, 03           'opening the earned income pop-up
	    transmit

	    this_job = 0                        'setting to loop through the rows and jobs
	    budg_row = 8
	    Do
	        EMReadScreen inc_type, 2, budg_row, 8           'looking for wage information
	        If inc_type = "__" Then Exit Do
	        If JOBS_ARRAY(est_pop_up, this_job) <> 0.00 Then    ''
	            If inc_type = "02" Then
	                EMReadScreen month_total, 11, budg_row, 43      'finding the income in that row of wages and formatting the amount
	                month_total = replace(month_total, "_", "")
	                month_total = trim(month_total)
	                month_total = month_total * 1
	                JOBS_ARRAY(six_month_total, this_job) = JOBS_ARRAY(six_month_total, this_job) + month_total     'adding this amount to the array - to create a sum of all the income listed for the job
	                this_job = this_job + 1     'going to the next job in the array
	            End If
	            budg_row = budg_row + 1     'looking at the next budget row
	        Else
	            this_job = this_job + 1     'going to the next job in the array
	        End If
	    Loop until this_job > UBOUND(JOBS_ARRAY, 2)
	    transmit

	    EMWriteScreen "X", 9, 03           'opening the unearned income pop-up
	    transmit

	    this_unea = 0                        'setting to loop through the rows and jobs
	    budg_row  = 8
	    Do
	        EMReadScreen inc_type, 2, budg_row, 8           'looking for wage information
	        If inc_type = "__" Then Exit Do
	        If UNEA_ARRAY(est_pop_up, this_unea) <> 0.00 Then    ''
	            If inc_type = "12" Then
	                EMReadScreen month_total, 11, budg_row, 43      'finding the income in that row of wages and formatting the amount
	                month_total = replace(month_total, "_", "")
	                month_total = trim(month_total)
	                month_total = month_total * 1
	                UNEA_ARRAY(six_month_total, this_unea) = UNEA_ARRAY(six_month_total, this_unea) + month_total     'adding this amount to the array - to create a sum of all the income listed for the job
	                this_unea = this_unea + 1     'going to the next job in the array
	            End If
	            budg_row = budg_row + 1     'looking at the next budget row
	        Else
	            this_unea = this_unea + 1     'going to the next job in the array
	        End If
	    Loop until this_unea > UBOUND(UNEA_ARRAY, 2)
	    transmit
		number_of_months = number_of_months + 1	'counting how many months meet the crieria for MA-EPD and should be used to average the income.
		transmit
	End If
	col = col + 11
    ' transmit
loop until col > 76
'this should never be 0 with the code checking for DP BUT if it is, there will always be an overflow error and we don't want that.
If number_of_months = 0 Then script_end_procedure("The script could not find MA-EPD months in the currently available HC Budget Span. The script will now end. Review STAT and ELIG to resovle.")
'This will set the word to use in the dialog to indicate how many months the income has been averaged over.
If number_of_months = 1 Then months_in_average = "one"
If number_of_months = 2 Then months_in_average = "two"
If number_of_months = 3 Then months_in_average = "three"
If number_of_months = 4 Then months_in_average = "four"
If number_of_months = 5 Then months_in_average = "five"
If number_of_months = 6 Then months_in_average = "six"

job_msg = ""
continue_fiat = FALSE
jobs_to_fiat = FALSE
unea_to_fiat = FALSE
fail_message = ""
For the_job = 0 to UBOUND(JOBS_ARRAY, 2)
    job_msg = job_msg & vbNewLine & JOBS_ARRAY(employer, the_job) & " - paid " & JOBS_ARRAY(job_frequency, the_job) & " on " & JOBS_ARRAY(pay_weekday, the_job) & " - total income for " & months_in_average & "-month budget period - $" & JOBS_ARRAY(six_month_total, the_job) & " - Average monthly income $" & JOBS_ARRAY(average_monthly_inc, the_job)
	If JOBS_ARRAY(job_frequency, the_job) = "3" then
		continue_fiat = TRUE
		jobs_to_fiat = TRUE
	End If
	If JOBS_ARRAY(job_frequency, the_job) = "4" then
		continue_fiat = TRUE
		jobs_to_fiat = TRUE
	End If

	If JOBS_ARRAY(job_frequency, the_job) = "2" then fail_message = fail_message & vbNewLine & "- Job: " & JOBS_ARRAY(employer, the_job) & "is paid semi-monthly and does not need a FIAT to balance the inocme."
	If JOBS_ARRAY(job_frequency, the_job) = "1" then fail_message = fail_message & vbNewLine & "- Job: " & JOBS_ARRAY(employer, the_job) & "is paid monthly and does not need a FIAT to balance the inocme."

	If UNEA_ARRAY(unea_type, 0) <> "" Then
		continue_fiat = TRUE
		unea_to_fiat = TRUE
	Else
		fail_message = fail_message & vbNewLine & "There is no Unemployment UNEA to FIAT."
	End If
Next
For the_job = 0 to UBOUND(JOBS_ARRAY, 2)
	If JOBS_ARRAY(job_frequency, the_job) = "_" then
		continue_fiat = FALSE
		fail_message = fail_message & vbNewLine & "- Job: " & JOBS_ARRAY(employer, the_job) & " does not have a pay frequency listed on the JOBS panel"
	End If
Next
If jobs_to_fiat = FALSE AND unea_to_fiat = FALSE Then continue_fiat = FALSE
If jobs_to_fiat = FALSE Then fail_message = fail_message & vbNewLine & "There are no JOBS to FIAT."
fail_message = "Script ended because FIAT could not continue due to:" & vbNewLine & fail_message
if continue_fiat = FALSE Then script_end_procedure_with_error_report(fail_message)

'Dynamic dialog to have the worker confirm the average income
Dialog1 = ""
y_pos = 60
BeginDialog Dialog1, 0, 0, 490, 140 + (UBOUND(JOBS_ARRAY, 2)*20) + (UBOUND(UNEA_ARRAY, 2)*20), "Average Monthly JOBS Income"
  Text 10, 10, 95, 10, "This case is at " & case_status     'identify if at Initial or Update
  If case_status = "Initial" Then Text 10, 25, 210, 10, "The date of application is " & application_date & " and the first month to FIAT is"
  If case_status = "Update" Then Text 10, 25, 210, 10, "The ongoing case is for " & MAXIS_footer_month & "/" & MAXIS_footer_year & " and the first month to FIAT is"
  EditBox 225, 20, 15, 15, MAXIS_footer_month
  EditBox 245, 20, 15, 15, MAXIS_footer_year
  Text 10, 45, 125, 10, "The script found the following  job(s):"
  If JOBS_ARRAY(employer, 0) <> "" Then
      For the_job = 0 to UBOUND(JOBS_ARRAY, 2)
          If JOBS_ARRAY(job_frequency, the_job) = "1" then JOBS_ARRAY(job_frequency, the_job) = "monthly"
          If JOBS_ARRAY(job_frequency, the_job) = "2" then JOBS_ARRAY(job_frequency, the_job) = "semi-monthly"
          If JOBS_ARRAY(job_frequency, the_job) = "3" then JOBS_ARRAY(job_frequency, the_job) = "biweekly"
          If JOBS_ARRAY(job_frequency, the_job) = "4" then JOBS_ARRAY(job_frequency, the_job) = "weekly"
          If JOBS_ARRAY(job_frequency, the_job) = "5" then JOBS_ARRAY(job_frequency, the_job) = "other"
        'setting the average monthly income by taking the total of the income listed in the budget on ELIG, then dividing it by the number of months
        'then using format number to make this a number with 2 decimel points and then removing the commas from the number because ELIG/HC doesn't like commas
        JOBS_ARRAY(average_monthly_inc, the_job) = FormatNumber(JOBS_ARRAY(six_month_total, the_job)/number_of_months, 2,,,0) & ""
        Text 20, y_pos + 5, 395, 10, JOBS_ARRAY(employer, the_job) & " - paid " & JOBS_ARRAY(job_frequency, the_job) & " on " & JOBS_ARRAY(pay_weekday, the_job) & " - total income for " & months_in_average & "-month budget period - $" & JOBS_ARRAY(six_month_total, the_job) & " - Average monthly income $"
        EditBox 415, y_pos, 55, 15, JOBS_ARRAY(average_monthly_inc, the_job)
        y_pos = y_pos + 20
      Next
  Else
      Text 20, y_pos, 300, 10, "No JOBS found."
      y_pos = y_pos + 20
  End If
  Text 10, y_pos, 125, 10, "The script found the following unemployment income:"
  y_pos = y_pos + 15
  If UNEA_ARRAY(unea_type, 0) <> "" Then
      For the_unea = 0 to UBOUND(UNEA_ARRAY, 2)
          If UNEA_ARRAY(job_frequency, the_unea) = "1" then UNEA_ARRAY(job_frequency, the_unea) = "monthly"
          If UNEA_ARRAY(job_frequency, the_unea) = "2" then UNEA_ARRAY(job_frequency, the_unea) = "semi-monthly"
          If UNEA_ARRAY(job_frequency, the_unea) = "3" then UNEA_ARRAY(job_frequency, the_unea) = "biweekly"
          If UNEA_ARRAY(job_frequency, the_unea) = "4" then UNEA_ARRAY(job_frequency, the_unea) = "weekly"
          If UNEA_ARRAY(job_frequency, the_unea) = "5" then UNEA_ARRAY(job_frequency, the_unea) = "other"
        'setting the average monthly income by taking the total of the income listed in the budget on ELIG, then dividing it by the number of months
        'then using format number to make this a number with 2 decimel points and then removing the commas from the number because ELIG/HC doesn't like commas
        UNEA_ARRAY(average_monthly_inc, the_unea) = FormatNumber(UNEA_ARRAY(six_month_total, the_unea)/number_of_months, 2,,,0) & ""
        Text 20, y_pos + 5, 405, 10, UNEA_ARRAY(unea_type, the_unea) & " - paid " & UNEA_ARRAY(job_frequency, the_unea) & " on " & UNEA_ARRAY(pay_weekday, the_unea) & " - total income for " & months_in_average & "-month budget period - $" & UNEA_ARRAY(six_month_total, the_unea) & " - Average monthly income $"
        EditBox 425, y_pos, 55, 15, UNEA_ARRAY(average_monthly_inc, the_unea)
        y_pos = y_pos + 20
      Next
  Else
      Text 20, y_pos, 300, 10, "No UNEA found."
      y_pos = y_pos + 20
  End If
  ButtonGroup ButtonPressed
    OkButton 370, y_pos + 5, 50, 15
    CancelButton 425, y_pos + 5, 50, 15
EndDialog

Do
    Do
        'showing the dialog that workers will indicate the monthly income to be used.
        err_msg = ""
        Dialog Dialog1
        cancel_confirmation

        If trim(MAXIS_footer_month) = "" or trim(MAXIS_footer_year) = "" Then
            err_msg = err_msg & vbNewLine & "* Enter the footer month and year in which the FIATing should start."
            If case_status = "Initial" Then err_msg = err_msg & vbNewLine & "  - This case is at application and most cases at application should be FIATed starting in the month of application."
            If case_status = "Update" Then err_msg = err_msg & vbNewLine & "  - This case is ongoing and most ongoing cases should be FIATed starting the first month of the next budget period."
        End If

        If JOBS_ARRAY(employer, 0) <> "" Then
            For the_job = 0 to UBOUND(JOBS_ARRAY, 2)
                If JOBS_ARRAY(average_monthly_inc, the_job) = "" Then err_msg = err_msg & vbNewLine & "* Enter the average monthly income for " & JOBS_ARRAY(employer, the_job) & "."
            Next
        End If

        If UNEA_ARRAY(unea_type, 0) <> "" Then
            For the_unea = 0 to UBOUND(UNEA_ARRAY, 2)
                If UNEA_ARRAY(average_monthly_inc, the_unea) = "" Then err_msg = err_msg & vbNewLine & "* Enter the average monthly income for " & UNEA_ARRAY(unea_type, the_unea) & "."
            Next
        End If

        If err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

Call back_to_SELF           'going back to SELF because footer month may had been changed

'NOW IT GOES TO ELIG/HC TO FIAT THE AMOUNTS
Call navigate_to_MAXIS_screen("ELIG", "HC__")

row = 1
col = 1
EMSearch memb_number & " ", row, col 'finding the member number
If row = 0 then script_end_procedure("Member number not found. You may have entered an incorrect member number on the first screen. Try the script again.")

EMWriteScreen "X", row, 26
transmit

row = 6
col = 1
EMSearch first_month_and_year, row, col 'finding the right month to start with

end_msg = "The selected month " & first_month_and_year & " is not in the current version of HC, review the month selected and try the script again."
If col = 0 Then script_end_procedure(end_msg)

'putting in the budget in edit mode to FIAT
PF9
EMReadScreen FIAT_check, 4, 24, 45
If FIAT_check <> "FIAT" then
    EMSendKey "05"
    transmit
End if

Do
	EMReadScreen elig_type_check, 2, 12, col-2		'checking to be sure the month is MA-EPD before FIATing.
	If elig_type_check = "DP" Then
		STATS_counter = STATS_counter + 1	'we count each month that is FIATed for statistics
	    EMWriteScreen "X", 9, col + 2       'opening the budget
	    transmit
	    EMWriteScreen "X", 13, 03           'opening the Earned Income line
	    transmit

	    budg_row = 8                        'reading each row to enter the information for each job
	    update_made = FALSE
	    For the_job = 0 to UBOUND(JOBS_ARRAY, 2)
	        JOBS_ARRAY(average_monthly_inc, the_job) = FormatNumber(JOBS_ARRAY(average_monthly_inc, the_job), 2,,,0)    'making sure the number is formatted correctly, 2 decimal places and no commas
	        If JOBS_ARRAY(average_monthly_inc, the_job) <> 0.00 Then
	            EmWriteScreen "___________", budg_row, 43       'blanking out the current income amount
	            EmWriteScreen JOBS_ARRAY(average_monthly_inc, the_job), budg_row, 43        'writing in the new averaged amount
	            budg_row = budg_row + 1
	            update_made = TRUE
	        End If
	    Next
	    If update_made = TRUE Then transmit            'saving the earned income amount
	    transmit            'closing the earned income pop-up

	    EMWriteScreen "X", 9, 03           'opening the unearned income pop-up
	    transmit

	    budg_row  = 8
	    update_made = FALSE
	    For the_unea = 0 to UBOUND(UNEA_ARRAY, 2)
	        Do
	            EMReadScreen inc_type, 2, budg_row, 8           'looking for wage information
	            If inc_type = "12" Then
	                UNEA_ARRAY(average_monthly_inc, the_unea) = FormatNumber(UNEA_ARRAY(average_monthly_inc, the_unea), 2,,,0)    'making sure the number is formatted correctly, 2 decimal places and no commas
	                If UNEA_ARRAY(average_monthly_inc, the_unea) <> 0.00 Then
	                    EmWriteScreen "__________", budg_row, 43       'blanking out the current income amount
	                    EmWriteScreen UNEA_ARRAY(average_monthly_inc, the_unea), budg_row, 43        'writing in the new averaged amount
	                    budg_row = budg_row + 1
	                    update_made = TRUE
	                End If
	            Else
	                budg_row = budg_row + 1
	            End If
	        Loop until inc_type = "__"
	    Next
	    If update_made = TRUE Then transmit            'saving the earned income amount
	    transmit            'closing the earned income pop-up
		transmit            'closing the budget pop-up
	End If
    col = col + 11      'going to next month
loop until col > 76

script_end_procedure_with_error_report("Success! Please make sure to check eligibility for any Medicare savings programs such as QMB or SLMB.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------05/25/2022
'--Tab orders reviewed & confirmed----------------------------------------------05/25/2022
'--Mandatory fields all present & Reviewed--------------------------------------05/25/2022
'--All variables in dialog match mandatory fields-------------------------------05/25/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------05/25/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------05/25/2022
'--Out-of-County handling reviewed----------------------------------------------05/25/2022
'--script_end_procedures (w/ or w/o error messaging)----------------------------05/25/2022
'--BULK - review output of statistics and run time/count (if applicable)--------09/30/2022
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---05/25/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------05/25/2022
'--Incrementors reviewed (if necessary)-----------------------------------------05/25/2022
'--Denomination reviewed -------------------------------------------------------05/25/2022
'--Script name reviewed---------------------------------------------------------05/25/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------09/30/2022

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------05/25/2022
'--comment Code-----------------------------------------------------------------09/30/2022
'--Update Changelog for release/update------------------------------------------05/25/2022
'--Remove testing message boxes-------------------------------------------------05/25/2022
'--Remove testing code/unnecessary code-----------------------------------------05/25/2022
'--Review/update SharePoint instructions----------------------------------------09/30/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------N/A
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------N/A
