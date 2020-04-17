'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - JOB CHANGE REPORTED.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 345                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================
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
call changelog_update("01/05/2018", "Updated coordinates in STAT/JOBS for income type and verification codes.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS ================================================================================================================
function determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case)
' fills a seris of booleans for case and programs status'
    Call navigate_to_MAXIS_screen("CASE", "CURR")
    family_cash_case = FALSE
    adult_cash_case = FALSE
    ga_case = FALSE
    msa_case = FALSE
    mfip_case = FALSE
    dwp_case = FALSE
    grh_case = FALSE
    snap_case = FALSE
    ma_case = FALSE
    msp_case = FALSE
    case_active = FALSE
    case_pending = FALSE
    row = 1
    col = 1
    EMSearch "FS:", row, col
    MsgBox "FS Row - " & row
    If row <> 0 Then
        EMReadScreen fs_status, 9, row, col + 4
        fs_status = trim(fs_status)
        If fs_status = "ACTIVE" or fs_status = "APP CLOSE" or fs_status = "APP OPEN" Then
            snap_case = TRUE
            case_active = TRUE
        End If
        If fs_status = "PENDING" Then
            snap_case = TRUE
            case_pending = TRUE
        ENd If
    End If

    row = 1
    col = 1
    EMSearch "GRH:", row, col
    If row <> 0 Then
        EMReadScreen grh_status, 9, row, col + 5
        grh_status = trim(grh_status)
        If grh_status = "ACTIVE" or grh_status = "APP CLOSE" or grh_status = "APP OPEN" Then
            grh_case = TRUE
            case_active = TRUE
        End If
        If grh_status = "PENDING" Then
            grh_case = TRUE
            case_pending = TRUE
        ENd If
    End If

    row = 1
    col = 1
    EMSearch "MSA:", row, col
    If row <> 0 Then
        EMReadScreen ms_status, 9, row, col + 5
        ms_status = trim(ms_status)
        If ms_status = "ACTIVE" or ms_status = "APP CLOSE" or ms_status = "APP OPEN" Then
            msa_case = TRUE
            adult_cash_case = TRUE
            case_active = TRUE
        End If
        If ms_status = "PENDING" Then
            msa_case = TRUE
            adult_cash_case = TRUE
            case_pending = TRUE
        ENd If
    End If

    row = 1
    col = 1
    EMSearch "GA:", row, col
    If row <> 0 Then
        EMReadScreen ga_status, 9, row, col + 4
        ga_status = trim(ga_status)
        If ga_status = "ACTIVE" or ga_status = "APP CLOSE" or ga_status = "APP OPEN" Then
            ga_case = TRUE
            adult_cash_case = TRUE
            case_active = TRUE
        End If
        If ga_status = "PENDING" Then
            ga_case = TRUE
            adult_cash_case = TRUE
            case_pending = TRUE
        ENd If
    End If

    row = 1
    col = 1
    EMSearch "DWP:", row, col
    If row <> 0 Then
        EMReadScreen dw_status, 9, row, col + 4
        dw_status = trim(dw_status)
        If dw_status = "ACTIVE" or dw_status = "APP CLOSE" or dw_status = "APP OPEN" Then
            dwp_case = TRUE
            family_cash_case = TRUE
            case_active = TRUE
        End If
        If dw_status = "PENDING" Then
            dwp_case = TRUE
            family_cash_case = TRUE
            case_pending = TRUE
        ENd If
    End If

    row = 1
    col = 1
    EMSearch "MFIP:", row, col
    If row <> 0 Then
        EMReadScreen mf_status, 9, row, col + 6
        mf_status = trim(mf_status)
        If mf_status = "ACTIVE" or mf_status = "APP CLOSE" or mf_status = "APP OPEN" Then
            mfip_case = TRUE
            family_cash_case = TRUE
            case_active = TRUE
        End If
        If mf_status = "PENDING" Then
            mfip_case = TRUE
            family_cash_case = TRUE
            case_pending = TRUE
        ENd If
    End If

    row = 1
    col = 1
    EMSearch "MA:", row, col
    If row <> 0 Then
        EMReadScreen ma_status, 9, row, col + 4
        ma_status = trim(ma_status)
        If ma_status = "ACTIVE" or ma_status = "APP CLOSE" or ma_status = "APP OPEN" Then
            ma_case = TRUE
            case_active = TRUE
        End If
        If ma_status = "PENDING" Then
            ma_case = TRUE
            case_pending = TRUE
        ENd If
    End If

    row = 1
    col = 1
    EMSearch "QMB:", row, col
    If row <> 0 Then
        EMReadScreen qm_status, 9, row, col + 5
        qm_status = trim(qm_status)
        If qm_status = "ACTIVE" or qm_status = "APP CLOSE" or qm_status = "APP OPEN" Then
            msp_case = TRUE
            case_active = TRUE
        End If
        If qm_status = "PENDING" Then
            msp_case = TRUE
            case_pending = TRUE
        ENd If
    End If
    row = 1
    col = 1
    EMSearch "SLMB:", row, col
    If row <> 0 Then
        EMReadScreen sl_status, 9, row, col + 6
        sl_status = trim(sl_status)
        If sl_status = "ACTIVE" or sl_status = "APP CLOSE" or sl_status = "APP OPEN" Then
            msp_case = TRUE
            case_active = TRUE
        End If
        If sl_status = "PENDING" Then
            msp_case = TRUE
            case_pending = TRUE
        ENd If
    End If
    row = 1
    col = 1
    EMSearch "QMB:", row, col
    If row <> 0 Then
        EMReadScreen qm_status, 9, row, col + 5
        qm_status = trim(qm_status)
        If qm_status = "ACTIVE" or qm_status = "APP CLOSE" or qm_status = "APP OPEN" Then
            msp_case = TRUE
            case_active = TRUE
        End If
        If qm_status = "PENDING" Then
            msp_case = TRUE
            case_pending = TRUE
        ENd If
    End If

End Function
'===========================================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
'connecting to MAXIS & grabbing the case number
EMConnect ""
call MAXIS_case_number_finder(MAXIS_case_number)

'Shows and defines the case number dialog
BeginDialog Dialog1, 0, 0, 276, 110, "Job Change Selection"
  EditBox 130, 10, 60, 15, MAXIS_case_number
  DropListBox 130, 30, 140, 45, "Select One ..."+chr(9)+"New Job Reported"+chr(9)+"Income/Hours Chnage for Current Job"+chr(9)+"Job Ended", job_change_type
  CheckBox 130, 50, 130, 10, "Check here if the JOBS panel exists.", job_panel_exists
  EditBox 130, 65, 140, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 160, 90, 50, 15
    CancelButton 220, 90, 50, 15
  Text 40, 15, 85, 10, "Enter your case number:"
  Text 10, 35, 110, 10, "What is the nature of the change?"
  Text 60, 70, 60, 10, "Worker Signature:"
EndDialog


Do
    Do
        err_msg = ""
	    dialog Dialog1					'Calling a dialog without a assigned variable will call the most recently defined dialog
	    cancel_without_confirmation

        Call validate_MAXIS_case_number(err_msg, "*")
        If job_change_type = "Select One ..." Then err_msg = err_msg & vbNewLine & "* Indicate what type of change is being reported (stop, start, or change)."
        If trim(worker_signature) = "" Then err_msg = err_msg & vbNewLine & "* Enter your name for the CASE/NOTE."

        If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
    Loop until err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

If job_change_type = "New Job Reported" Then change_text = "new job"
If job_change_type = "Income/Hours Chnage for Current Job" Then change_text = "change job income"
If job_change_type = "Job Ended" Then change_text = "stwk"

BeginDialog Dialog1, 0, 0, 251, 130, "Verification Received?"
  DropListBox 10, 35, 235, 45, "Select One..."+chr(9)+"Yes - sufficient verification received to budget income accurately."+chr(9)+"No - we need to request additional verification.", verifs_received_selection
  ButtonGroup ButtonPressed
    OkButton 140, 110, 50, 15
    CancelButton 195, 110, 50, 15
  Text 10, 10, 175, 20, "Have we received verification sufficient verification to budget the income for this " & change_text & "?"
  Text 25, 55, 190, 50, "This script is meant for noting and updating when a change in job is reported but has not been sufficiently verified. If we have sufficient verification to budget the new income, we need to use the script ACTIONS - Earned Income Budgeting as it has been created for the purpose of updating and docudmenting verified job changes."
EndDialog
Do
    Do
        err_msg = ""

        Dialog Dialog1
        cancel_without_confirmation

        If verifs_received_selection = "Select One..." Then err_msg = err_msg & vbNewLine & "* Indicate if we have sufficient verification or are requesting more."

        If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

If verifs_received_selection = "Yes - sufficient verification received to budget income accurately." Then
    MsgBox "Since verifications have been received for a " & change_text & ", the " & vbNewLine & vbNewLine & "             ACTIONS - Earned Income Budgeting       " & vbNewLine & vbNewLine &"script will run to accurately update and CASE/NOTE the budgeted income."
    Call run_from_GitHub(script_repository & "actions/earned-income-budgeting.vbs")
End If

Call back_to_SELF
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

MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr

Call generate_client_list(list_of_employees, "Select One ...")
call generate_client_list(list_of_members, "Type or Select")
client_name_array = split(list_of_members, chr(9))

Call MAXIS_background_check

CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.

DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
    EMReadscreen ref_nbr, 3, 4, 33

    client_string = client_string & "|" & ref_nbr

    transmit
    Emreadscreen edit_check, 7, 24, 2
LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

client_string = trim(client_string)
client_string = right(client_string, len(client_string) - 1)
ref_nbr_array = split(client_string, "|")

If job_panel_exists = checked Then
    job_selected = FALSE
    For each clt_number in ref_nbr_array

        Call Navigate_to_MAXIS_screen("STAT", "JOBS")                       'navigate to the current panel in the array

        EMWriteScreen clt_number, 20, 76
        transmit

        EMReadScreen total_jobs, 1, 2, 78
        If total_jobs <> "0" Then
            Do
                EMReadScreen employer, 30, 7, 42
                employer = replace(employer, "_", "")
                employer_check = MsgBox("Is this the job reported? Employer name: " & employer, vbYesNo + vbQuestion, "Select Income Panel")
                If employer_check = vbYes Then
                    job_selected = TRUE
                    Exit Do
                End If
                transmit
                EMReadScreen last_job, 7, 24, 2
            Loop until last_job = "ENTER A"
        End If
        If job_selected = TRUE Then Exit For
    Next

    If job_selected = TRUE Then
        EMReadScreen job_update_date, 8, 21, 55
        EMReadScreen job_instance, 1, 2, 73

        EMReadScreen job_ref_number, 2, 4, 33
        EMReadScreen job_income_type, 1, 5, 34
        EMReadScreen job_verification, 1, 6, 34
        EMReadScreen job_subsidized_income_type, 2, 5, 74
        EMReadScreen job_hourly_wage, 6, 6, 75
        EMReadScreen job_employer_full_name, 30, 7, 42
        EMReadScreen job_income_start, 8, 9, 35
        EMReadScreen job_income_end, 8, 9, 49
        EMReadScreen job_pay_frequency, 1, 18, 35

        EMWriteScreen "X", 19, 38           'opening the PIC'
        transmit

        EMReadScreen job_pic_calculation_date, 8, 5, 34
        EMReadScreen job_pic_pay_frequency, 1, 5, 64
        EMReadScreen job_pic_hours_per_week, 6, 8, 64
        EMReadScreen job_pic_pay_per_hour, 8, 9, 66

        PF3

        job_update_date = replace(job_update_date, " ", "/")
        job_instance = "0" & job_instance
        For each person in client_name_array
            If left(person, 2) = job_ref_number Then job_employee =  person
        Next
        If job_income_type = "J" Then job_oncome_type = "J - WIOA"
        If job_income_type = "W" Then job_income_type = "W - Wages"
        If job_income_type = "E" Then job_income_type = "E - EITC"
        If job_income_type = "G" Then job_income_type = "G - Experience Works"
        If job_income_type = "F" Then job_income_type = "F - Federal Work Study"
        If job_income_type = "S" Then job_income_type = "S - State Work Study"
        If job_income_type = "O" Then job_income_type = "O - TOher"
        If job_income_type = "C" Then job_income_type = "C - Contract Income"
        If job_income_type = "T" Then job_income_type = "T - Training Program"
        If job_income_type = "P" Then job_income_type = "P - Service Program"
        If job_income_type = "R" Then job_income_type = "R - Rehab Program"

        If job_verification = "1" Then job_verification = "1 - Pay Stubs/Tip Report"
        If job_verification = "2" Then job_verification = "2 - Employer Statement"
        If job_verification = "3" Then job_verification = "3 - Collateral Statement"
        If job_verification = "4" Then job_verification = "4 - Other Document"
        If job_verification = "5" Then job_verification = "5 - Pend Out State Verif"
        If job_verification = "N" Then job_verification = "N - No Verif Provided"
        If job_verification = "?" Then job_verification = "? - Delayed Verification"

        If job_subsidized_income_type = "__" Then job_subsidized_income_type = ""
        If job_subsidized_income_type = "01" Then job_subsidized_income_type = "01 - Subsidized Public Sector Employer"
        If job_subsidized_income_type = "02" Then job_subsidized_income_type = "02 - Subsidized Private Sector Employer"
        If job_subsidized_income_type = "03" Then job_subsidized_income_type = "03 - On-the-Job-Training"
        If job_subsidized_income_type = "04" Then job_subsidized_income_type = "04 - AmeriCorps"

        job_hourly_wage = trim(job_hourly_wage)
        job_hourly_wage = replace(job_hourly_wage, "_", "")
        job_employer_name = replace(job_employer_full_name, "_", "")
        job_income_start = replace(job_income_start, " ", "/")
        If job_income_start = "__/__/__" Then job_income_start = ""
        job_income_end = replace(job_income_end, " ", "/")
        If job_income_end = "__/__/__" Then job_income_end = ""
        If job_pay_frequency = "1" Then job_pay_frequency = "1 - Monthly"
        If job_pay_frequency = "2" Then job_pay_frequency = "2 - Semi-Monthly"
        If job_pay_frequency = "3" Then job_pay_frequency = "3 - Biweekly"
        If job_pay_frequency = "4" Then job_pay_frequency = "4 - Weekly"
        If job_pay_frequency = "5" Then job_pay_frequency = "5 - Other"

        job_pic_calculation_date = replace(job_pic_calculation_date, " ", "/")
        If job_pic_calculation_date = "__/__/__" Then job_pic_calculation_date = ""

        If job_pic_pay_frequency = "1" Then job_pic_pay_frequency = "1 - Monthly"
        If job_pic_pay_frequency = "2" Then job_pic_pay_frequency = "2 - Semi-Monthly"
        If job_pic_pay_frequency = "3" Then job_pic_pay_frequency = "3 - Biweekly"
        If job_pic_pay_frequency = "4" Then job_pic_pay_frequency = "4 - Weekly"

        job_pic_hours_per_week = trim(job_pic_hours_per_week)
        job_pic_hours_per_week = replace(job_pic_hours_per_week, "_", "")
        job_pic_pay_per_hour = trim(job_pic_pay_per_hour)
        job_pic_pay_per_hour = replace(job_pic_pay_per_hour, "_", "")

        job_update_date = DateValue(job_update_date)
        If job_update_date = date THen script_update_stat = "No - Update of JOBS not needed"
    End If
End If

Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case)

MsgBox "Case Information" & vbNewLine & vbNewLine & "Case Active - " & case_active & vbNewLine & "Case Pending - " & case_pending & vbNewLine & "Family Cash - " & family_cash_case & vbNewLine &_
       "MFIP - " & mfip_case & vbNewLine & "DWP - " & dwp_case & vbNewLine & "Adult Cash - " & adult_cash_case & vbNewLine & "GA - " & ga_case & vbNewLine & "MSA - " & msa_case & vbNewLine & "GRH - " & grh_case & vbNewLine &_
       "SNAP - " & snap_case & vbNewLine & "MA - " & ma_case & vbNewLine & "MSP - " & msp_case

For each clt_number in ref_nbr_array

    Call Navigate_to_MAXIS_screen("STAT", "WREG")                       'navigate to the current panel in the array

    EMWriteScreen clt_number, 20, 76
    transmit

    EMReadScreen pwe_code, 1, 6, 68
    If pwe_code = "Y" Then
        pwe_ref = clt_number
        Exit For
    End If
Next
MsgBox "PWE Reference Number is " & pwe_ref

If job_change_type = "New Job Reported" Then
    If IsDate(job_income_start) = TRUE Then new_job_income_start = job_income_start
End If
If job_change_type = "Job Ended" Then
    If IsDate(job_income_end) = TRUE Then job_end_income_end_date = job_income_end
End If
If job_verification = "" Then job_verification = "N - No Verif Provided"

' job_update_date
' job_instance
' job_ref_number
' job_employee
' job_income_type
' job_verification
' job_subsidized_income_type
' job_hourly_wage
' job_employer_name
' job_employer_full_name
' job_income_start
' job_income_end
' job_pay_frequency
'
' job_pic_calculation_date
' job_pic_pay_frequency
' job_pic_hours_per_week
' job_pic_pay_per_hour
reported_date = date & ""
script_update_stat = "Yes - Create a new JOBS Panel"
If job_panel_exists = checked Then script_update_stat = "Yes - Update an existing JOBS Panel"
If developer_mode = TRUE Then script_update_stat = "No - Update of JOBS not needed"

BeginDialog Dialog1, 0, 0, 536, 370, "New job reported dialog"
  GroupBox 5, 5, 525, 45, "Reporting Information"
  Text 7, 20, 93, 10, "* How was the job reported?:"
  Text 212, 20, 53, 10, "* Who reported?"
  Text 410, 20, 55, 10, "* Date Reported:"
  ComboBox 105, 15, 100, 15, "Type or Select"+chr(9)+"phone call"+chr(9)+"Change Report Form"+chr(9)+"office visit"+chr(9)+"mailing"+chr(9)+"fax"+chr(9)+"ES counselor"+chr(9)+"CCA worker"+chr(9)+"scanned document", job_report_type
  ComboBox 265, 15, 135, 45, list_of_members, person_who_reported_job
  EditBox 470, 15, 50, 15, reported_date
  CheckBox 105, 35, 275, 10, "Check here if the employee gave verbal authorization to check the Work Number.", work_number_verbal_checkbox

  GroupBox 5, 55, 525, 90, "Job Information"
  Text 7, 70, 83, 10, "* Employee (HH Member):"
  Text 290, 70, 60, 10, "Work Number ID:"
  If job_ref_number <> "" Then Text 475, 65, 50, 10, "JOBS " & job_ref_number & " "  & job_instance
  Text 45, 90, 45, 10, "Income Type:"
  Text 245, 90, 85, 10, "Subsidized Income Type:"
  Text 52, 110, 38, 10, "* Employer:"
  Text 242, 110, 43, 10, "* Verification:"
  Text 395, 110, 55, 10, "Pay Frequency:"
  Text 30, 130, 60, 10, "Income start date:"
  Text 165, 130, 60, 10, "Income End Date:"
  Text 300, 130, 75, 10, "Contract through date:"
  DropListBox 95, 65, 180, 45, list_of_employees, job_employee
  EditBox 350, 65, 50, 15, work_numb_id
  DropListBox 95, 85, 110, 15, "W - Wages"+chr(9)+"J - WIOA"+chr(9)+"E - EITC"+chr(9)+"G - Experience Works"+chr(9)+"F - Federal Work Study"+chr(9)+"S - State Work Study"+chr(9)+"O - Other"+chr(9)+"C - Contract Income"+chr(9)+"T - Training Program"+chr(9)+"P - Service Program"+chr(9)+"R - Rehab Program", job_income_type
  DropListBox 335, 85, 150, 15, ""+chr(9)+"01 - Subsidized Public Sector Employer"+chr(9)+"02 - Subsidized Private Sector Employer"+chr(9)+"03 - On-the-Job-Training"+chr(9)+"04 - AmeriCorps", job_subsidized_income_type
  EditBox 95, 105, 140, 15, job_employer_name
  DropListBox 290, 105, 95, 45, ""+chr(9)+"N - No Verif Provided"+chr(9)+"? - Delayed Verification", job_verification
  DropListBox 450, 105, 70, 45, " "+chr(9)+"1 - Monthly"+chr(9)+"2 - Semi-Monthly"+chr(9)+"3 - Biweekly"+chr(9)+"4 - Weekly"+chr(9)+"5 - Other", job_pay_frequency
  EditBox 95, 125, 55, 15, job_income_start
  EditBox 230, 125, 55, 15, job_income_end
  EditBox 380, 125, 55, 15, job_contract_through_date

  Select Case job_change_type
      Case "New Job Reported"
          GroupBox 5, 150, 525, 130, "Update Reported - NEW JOB"
          Text 15, 170, 65, 10, "Date Work Started:"
          Text 142, 170, 103, 10, "* Date Income started/will start:"
          Text 370, 170, 95, 10, "Initial check GROSS amount:"
          Text 10, 190, 65, 10, "Anticipated Income:"
          Text 80, 190, 60, 10, "Hours per Week:"
          Text 200, 190, 50, 10, "Hourly Wage:"
          Text 15, 210, 65, 10, "Conversation with"
          Text 225, 210, 25, 10, "details:"
          Text 12, 260, 93, 10, "* Impact on WREG/ABAWD:"
          EditBox 85, 165, 55, 15, date_work_started
          EditBox 250, 165, 50, 15, new_job_income_start
          EditBox 470, 165, 55, 15, initial_check_gross_amount
          EditBox 140, 185, 50, 15, new_job_hours_per_week
          EditBox 250, 185, 50, 15, new_job_hourly_wage
          ComboBox 80, 205, 135, 45, list_of_members, conversation_with_person
          EditBox 260, 205, 265, 15, conversation_detail
          CheckBox 80, 225, 190, 10, "Check here if you sent a Work Number request.", work_number_checkbox
          EditBox 105, 255, 420, 15, wreg_abawd_notes
      Case "Income/Hours Chnage for Current Job"
          GroupBox 5, 150, 525, 130, "Update Reported - JOB CHANGE"
          Text 15, 165, 55, 10, "Date of Change:"
          Text 145, 165, 65, 10, "Change Reported:"
          Text 15, 185, 80, 10, "Old Anticipated Income:"
          Text 100, 185, 60, 10, "Hours per Week:"
          Text 230, 185, 50, 10, "Hourly Wage:"
          Text 340, 185, 55, 10, "Income Change:"
          Text 10, 205, 85, 10, "* New Anticipated Income:"
          Text 100, 205, 60, 10, "Hours per Week:"
          Text 230, 205, 50, 10, "Hourly Wage:"
          Text 377, 205, 83, 10, "* First Pay Date Impacted:"
          Text 15, 225, 65, 10, "Conversation with"
          Text 225, 225, 25, 10, "details:"
          Text 12, 260, 93, 10, "* Impact on WREG/ABAWD:"
          EditBox 70, 160, 55, 15, job_change_date
          EditBox 210, 160, 315, 15, job_change_details
          EditBox 165, 180, 50, 15, job_change_old_hours_per_week
          EditBox 280, 180, 50, 15, job_change_old_hourly_wage
          ComboBox 395, 180, 130, 45, "Select or Type"+chr(9)+"Increase"+chr(9)+"Decrease", income_change_type
          EditBox 165, 200, 50, 15, job_change_new_hours_per_week
          EditBox 280, 200, 50, 15, job_change_new_hourly_wage
          EditBox 465, 200, 60, 15, first_pay_date_of_change
          ComboBox 80, 220, 135, 45, list_of_members, conversation_with_person
          EditBox 260, 220, 265, 15, conversation_detail
          CheckBox 80, 240, 190, 10, "Check here if you sent a Work Number request.", work_number_checkbox
          EditBox 105, 255, 420, 15, wreg_abawd_notes
      Case "Job Ended"
          GroupBox 5, 150, 525, 130, "Update Reported - JOB ENDED"
          Text 12, 170, 68, 10, "* Date Work Ended:"
          Text 147, 170, 103, 10, "* Date Income ended/will end:"
          Text 312, 170, 63, 10, "* Last pay amount:"
          Text 15, 190, 60, 10, "Reason for STWK:"
          Text 217, 190, 53, 10, "* Voluntary Quit:"
          Text 350, 190, 100, 10, "Does this meet Good Cause?"
          Text 15, 210, 65, 10, "Conversation with"
          Text 225, 210, 25, 10, "details:"
          Text 15, 230, 160, 10, "Is the client applying for Unemployment Income?"
          Text 232, 230, 93, 10, "* Impact on WREG/ABAWD:"
          Text 12, 265, 68, 10, "* STWK Verification:"
          Text 235, 265, 50, 10, "Refused Empl:"
          Text 350, 265, 65, 10, "Refused Empl Date:"
          EditBox 80, 165, 55, 15, date_work_ended
          EditBox 250, 165, 50, 15, job_end_income_end_date
          EditBox 380, 165, 50, 15, last_pay_amount
          EditBox 80, 185, 130, 15, stwk_reason
          DropListBox 275, 185, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", vol_quit_yn
          DropListBox 450, 185, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", good_cause_yn
          ComboBox 80, 205, 135, 45, list_of_members, conversation_with_person
          EditBox 260, 205, 260, 15, conversation_detail
          DropListBox 180, 225, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", uc_yn
          EditBox 330, 225, 190, 15, wreg_abawd_notes
          CheckBox 80, 245, 190, 10, "Check here if you sent a Work Number request.", work_number_checkbox
          DropListBox 80, 260, 125, 45, "Select One..."+chr(9)+"N - No Verif Provided"+chr(9)+"? - Delayed Verification", stwk_verif
          DropListBox 290, 260, 30, 45, "?"+chr(9)+"Yes"+chr(9)+"No", refused_empl_yn
          EditBox 420, 260, 65, 15, refused_empl_date
          ' ComboBox 80, 260, 125, 45, "Select One..."+chr(9)+"1 - Employers Statement"+chr(9)+"2 - Seperation Notice"+chr(9)+"3 - Collateral Statement"+chr(9)+"4 - Other Document"+chr(9)+"N - No Verif Provided", stwk_verif

  End Select

  GroupBox 5, 280, 525, 65, "Actions"
  Text 10, 295, 105, 10, "Date verification Request Sent:"
  Text 185, 295, 105, 10, "Time frame of verifs requested:"
  Text 20, 330, 90, 10, "Have Script Update Panel:"
  Text 10, 355, 25, 10, "Notes:"
  EditBox 120, 290, 50, 15, verif_form_date
  EditBox 295, 290, 75, 15, verif_time_frame
  CheckBox 385, 295, 105, 10, "Check here to TIKL for return.", TIKL_checkbox
  CheckBox 120, 310, 165, 10, "Check here if you are requesting CEI/OHI docs.", requested_CEI_OHI_docs_checkbox
  DropListBox 120, 325, 135, 45, "Select One..."+chr(9)+"No - Update of JOBS not needed"+chr(9)+"Yes - Update an existing JOBS Panel"+chr(9)+"Yes - Create a new JOBS Panel", script_update_stat
  CheckBox 365, 325, 165, 10, "Check here if you sent a status update to CCA.", CCA_checkbox
  CheckBox 365, 310, 160, 10, "Check here if you sent a status update to ES.", ES_checkbox
  EditBox 40, 350, 370, 15, notes
  ButtonGroup ButtonPressed
    OkButton 425, 350, 50, 15
    CancelButton 480, 350, 50, 15
EndDialog

Do
    Do
        err_msg = ""

        dialog Dialog1
        cancel_confirmation

        If job_change_type = "New Job Reported" Then
            If IsDate(job_income_start) = TRUE Then
                If IsDate(new_job_income_start) = TRUE Then
                    If job_income_start <> new_job_income_start Then err_msg = err_msg & vbNewLine & "* The income start dates do not match. Review the income start dates"
                Else
                    new_job_income_start = job_income_start
                End If
            Else
                If IsDate(new_job_income_start) = TRUE Then job_income_start = new_job_income_start
            End If
        End If

        If job_report_type = "Select or Type" OR trim(job_report_type) = "" Then err_msg = err_msg & vbNewLine & "* Indicate how the income was reported."
        If person_who_reported_job = "Select or Type" or trim(person_who_reported_job) = "" Then err_msg = err_msg & vbNewLine & "* Select the household member, or type in the name of the person who reported the job."
        If IsDate(reported_date)  = False Then err_msg = err_msg & vbNewLine & "* Enter a valid date for when this job was reported."

        If job_employee = "Select One ..." Then err_msg = err_msg & vbNewLine & "* Select the household member who is the employee."
        If trim(job_employer_name) = "" Then err_msg = err_msg & vbNewLine & "* Enter the name of the employer."
        If len(trim(job_employer_name)) > 30 Then err_msg = err_msg & vbNewLine & "* The employer name is more than 30 characters and MAXIS only allows for 30 characters on the employer line. Change the employer name to fit MAIXS."
        If job_verification = "" Then err_msg = err_msg & vbNewLine & "* Enter the verification of the JOBS panel."
        If job_pay_frequency = " " Then err_msg = err_msg & vbNewLine & "* "
        If IsDate(job_income_start) = False Then err_msg = err_msg & vbNewLine & "* Enter the date the income will or has started."

        Select Case job_change_type
            Case "New Job Reported"
                If IsDate(new_job_income_start) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the date the income will or has started."
                If trim(new_job_hourly_wage) <> "" AND IsNumeric(new_job_hourly_wage) = False Then err_msg = err_msg & vbNewLine & "* Enter the hourly wage as a number."
                If trim(new_job_hours_per_week) <> "" AND IsDate(new_job_hours_per_week) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the hours per week as a number."

            Case "Income/Hours Chnage for Current Job"
                If IsDate(job_change_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the date the change will be reflected in income (the first paycheck that will be affected by this change)."
                If trim(job_change_details) = "" Then err_msg = err_msg & vbNewLine & "* Enter the information about the change reported."
                If trim(job_change_new_hourly_wage) <> "" AND IsNumeric(job_change_new_hourly_wage) = False Then err_msg = err_msg & vbNewLine & "* Enter the hourly wage as a number."
                If trim(job_change_new_hours_per_week) <> "" AND IsNumeric(job_change_new_hours_per_week) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the hours per week as a number."
                If IsDate(first_pay_date_of_change) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the date that this change will be reflected in a change of pay."

            Case "Job Ended"
                If IsDate(job_end_income_end_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the date of the last paycheck."
                If IsDate(date_work_ended) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the date the client last worked."
                If IsNumeric(last_pay_amount) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the amount of the last pay. This can be an estimate amount."
                If vol_quit_yn = "?" Then err_msg = err_msg & vbNewLine & "* Select Voluntary Quit information - Yes or No"
                If stwk_verif = "Select One..." Then err_msg = err_msg & vbNewLine & "* Enter the verification code for the STWK Panel."
                If refused_empl_yn = "Yes" AND IsDate(refused_empl_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Since you indicated that the client refused employment, you must enter the date employment was refused as a valid date."
                If vol_quit_yn = "Yes" AND good_cause_yn = "No" Then err_msg = err_msg & vbNewLine & "* Currently there voluntary quit sanctions and processing is suspended due to COVID-19. If the client quit voluntarily, code good cause as 'Yes'."
                If vol_quit_yn = "Yes" AND good_cause_yn = "?" Then err_msg = err_msg & vbNewLine & "* Since this job has ended voluntarily, indicate if this voluntary quit meets good cause."
        End Select
        If trim(wreg_abawd_notes) = "" Then err_msg = err_msg & vbNewLine & "* Enter how this change will impact WREG/ABWAD for this pmember. NOTE: The imact may be 'NO CHANGE' or something similar. This information provided here should be thorough and complete, more is better when explaining the WREG and ABAWD status."
        If conversation_with_person = "Type or Select" Then
            If trim(conversation_detail) <> "" Then err_msg = err_msg & vbNewLine & "* There is information added to the conversation detail but no information about who the conversation was with has been provided. Enter the member or name of the person the conversationwas completed with."
        Else
            If trim(conversation_detail) = "" Then err_msg = err_msg & vbNewLine & "* There is information provided about who the conversation has been completed with but no details about what was discussed in the conversation. Add details about the conversation."
        End If

        If trim(verif_form_date) <> "" Then
            If IsDate(verif_form_date) = FALSE Then
                err_msg = err_msg & vbNewLine & "* Enter a valid date for the date the verification request form has been sent."
            Else
                If trim(verif_time_frame) = "" Then err_msg = err_msg & vbNewLine & "* For the verification request, indicate the time frame requested of the verifications. (This could be date specific or a general month.)"
            End If
        Else
            If TIKL_checkbox = checked Then err_msg = err_msg & vbNewLine & "* You have requested to TIKL for the verification request return but no verification request date has been entered. Either update the verifications request date or uncheck the TIKL for return box."
            If requested_CEI_OHI_docs_checkbox = checked Then err_msg = err_msg & vbNewLine & "* You have indicated that CEI/OHI documents are being requested but have not indicated when the verification request form was sent. Either update the verification request date or uncheck the CEI/OHI box."
        End If

        If script_update_stat = "Select One..." Then err_msg = err_msg & vbNewLine & "* Indicate if the script should update the STAT panels or not."

        If trim(verif_time_frame) <> "" AND IsDate(verif_form_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Indate the time frame for the verification request."

        If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg

    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

review_vol_quit = FALSE
Select Case job_change_type
    Case "New Job Reported"
        start_month = DatePart("m", new_job_income_start)
        start_year = DatePart("yyyy", new_job_income_start)
    Case "Income/Hours Chnage for Current Job"
        start_month = DatePart("m", first_pay_date_of_change)
        start_year = DatePart("yyyy", first_pay_date_of_change)
    Case "Job Ended"
        start_month = DatePart("m", job_end_income_end_date)
        start_year = DatePart("yyyy", job_end_income_end_date)
End Select
MAXIS_footer_month = right("00" & start_month, 2)
MAXIS_footer_year = right(start_year, 2)

If job_change_type = "Income/Hours Chnage for Current Job" Then
    If IsNumeric(job_change_old_hours_per_week) = TRUE Then
        job_change_old_hours_per_week = job_change_old_hours_per_week * 1
        job_change_new_hours_per_week = job_change_new_hours_per_week * 1

        If job_change_old_hours_per_week > 30 and job_change_new_hourly_wage <30 Then
            review_vol_quit = TRUE
            Vol_quit_type = "Hours Reduction"
        End If
    End If
End If

If vol_quit_yn = "Yes" Then
    review_vol_quit = TRUE
    Vol_quit_type = "Quit Job"
End If

If review_vol_quit = TRUE Then
    If mfip_case = TRUE Then mfip_vol_quit_checkbox = checked
    If dwp_case = TRUE Then dwp_vol_quit_checkbox = checked
    If snap_case = TRUE Then snap_vol_quit_checkbox = checked
    'TODO - If SNAP then we need to go to DISQ to determine if we can tell which sanction it is.
    Do
        Do
            err_msg = ""

            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 236, 210, "Voluntary Quit Detail"
              DropListBox 120, 25, 105, 45, "Select One..."+chr(9)+"Quit Job"+chr(9)+"Hours Reduction", vol_quit_type
              DropListBox 170, 45, 55, 45, "?"+chr(9)+"Yes"+chr(9)+"No", vol_quit_yn
              DropListBox 170, 65, 55, 45, "?"+chr(9)+"Yes"+chr(9)+"No", good_cause_yn
              EditBox 15, 100, 210, 15, vol_quit_reason
              CheckBox 35, 150, 30, 10, "MFIP", mfip_vol_quit_checkbox
              CheckBox 75, 150, 30, 10, "DWP", dwp_vol_quit_checkbox
              CheckBox 120, 150, 30, 10, "SNAP", snap_vol_quit_checkbox
              DropListBox 130, 170, 95, 45, "Zelect One..."+chr(9)+"First Sanction      - 1st"+chr(9)+"Second Sanction - 2nd"+chr(9)+"Third Sanction     - 3rd"+chr(9)+"More than Three", snap_vol_quit_occurance
              ButtonGroup ButtonPressed
                OkButton 130, 190, 50, 15
                CancelButton 185, 190, 50, 15
              Text 10, 10, 255, 10, "This case appears to have or potentially have a voluntary quit situation."
              Text 15, 30, 105, 10, "What kind of action was taken? "
              Text 15, 50, 150, 10, "Was this voluntary on the part of the Client?"
              Text 15, 70, 120, 10, "If  voluntary, is there  good cause?"
              Text 15, 90, 135, 10, "Explain the cuase as provided by client:"
              Text 90, 120, 135, 10, "Explain if client meets good cause or not."
              GroupBox 15, 135, 140, 30, "Program Impacted by Voluntary Quit"
              Text 15, 175, 110, 10, "If SNAP, what occurance is this?"
            EndDialog

            dialog Dialog1
            cancel_confirmation

            vol_quit_reason = trim(vol_quit_reason)
            If vol_quit_yn = "?" Then err_msg = err_msg & vbNewLine & "* Indicate if the job was voluntarily quit/hours reduced."
            If vol_quit_yn = "Yes" Then
                If vol_quit_type = "Select One..." Then err_msg = err_msg & vbNewLine & "* Indicate if the job was voluntarily ended or voluntarily reduced hours."
                If len(vol_quit_reason) < 20 Then err_msg = err_msg & vbNewLine & "* Enter full detail of information about the reason the job was ended. IF the client has not provided a reason, list that here and indicate how we are going to determine possible good cause."
                If snap_vol_quit_checkbox = checked AND snap_vol_quit_occurance = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since this voluntary quit impacts the SNAP program, indicate which occurance of this type of sanction this client is on."
            End If
            If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg

        Loop until err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE
End If

Call back_to_SELF
Call MAXIS_background_check
MsgBox "Start month - " & MAXIS_footer_month & vbNewLine & "Start year - " & MAXIS_footer_year

Initial_footer_month = MAXIS_footer_month
Initial_footer_year = MAXIS_footer_year
second_loop = FALSE
counter = 1

If developer_mode = FALSE
    Do
        EMWriteScreen "SUMM", 20, 71
        transmit

        EMReadScreen MAXIS_footer_month, 2, 20, 55
        EMReadScreen MAXIS_footer_year, 2, 20, 58

        MsgBox "Counter - " & counter & vbNewLine & "Second loop? - " & second_loop & vbNewLine & "MAXIS footer year - " & MAXIS_footer_month & "/" & MAXIS_footer_year & vbNewLine & "Should be on SUMM"

        If second_loop = TRUE Then
            EMWriteScreen "JOBS", 20, 71                    'go to JOBS
            EMWriteScreen ref_nbr, 20, 76    'go to the right member
            EMWriteScreen job_instance, 20, 79
            transmit

            EMReadScreen this_panel_job, 30, 7, 42

            If this_panel_job <> original_full_jobs_name Then
                MsgBox "They don't match!" & vbNewLine & "THIS PANEL-" & this_panel_job & "-" & vbNewLine & "ORIGINAL-" & original_full_jobs_name & "-"
            Else
                PF9
            End If

        ElseIf script_update_stat = "Yes - Update an existing JOBS Panel" Then
            MsgBox "Update - " & script_update_stat & " - 1"
            If job_instance = "" Then
                ref_nbr = left(job_employee, 2)
                EMWriteScreen ref_nbr, 20, 76    'go to the right member
                transmit

                EMReadScreen total_jobs, 1, 2, 78
                If total_jobs = "0" Then
                    Call script_end_procedure("Update and Note NOT Completed. There are no jobs for " & job_employee & " listed in MAXIS and you have selected to have the script Update and existing JOBS panel.")
                Else
                    job_selected = FALSE
                    Do
                        EMReadScreen employer, 30, 7, 42
                        employer = replace(employer, "_", "")
                        employer_check = MsgBox("Is this the job reported? Employer name: " & employer, vbYesNo + vbQuestion, "Select Income Panel")
                        If employer_check = vbYes Then
                            job_selected = TRUE
                            Exit Do
                        End If
                        transmit
                        EMReadScreen last_job, 7, 24, 2
                    Loop until last_job = "ENTER A"
                    If job_selected = FALSE Then Call script_end_procedure("Update and Note NOT completed. You did not select any of the JOBS for " & job_employee & " but indicated the script should update a JOBS panel.")
                End If
                EMReadScreen this_instance, 1, 2, 73
                job_instance = "0" & this_instance
            Else
                EMWriteScreen "JOBS", 20, 71                    'go to JOBS
                ref_nbr = left(job_employee, 2)
                EMWriteScreen ref_nbr, 20, 76    'go to the right member
                EMWriteScreen job_instance, 20, 79

                transmit
            End If
            PF9

        ElseIf script_update_stat = "Yes - Create a new JOBS Panel" Then
            MsgBox "Update - " & script_update_stat & " - 2"
            EMWriteScreen "JOBS", 20, 71                    'go to JOBS
            ref_nbr = left(job_employee, 2)
            EMWriteScreen ref_nbr, 20, 76    'go to the right member
            EMWriteScreen "NN", 20, 79                      'create new JOBS panel

            transmit

            EMReadScreen this_instance, 1, 2, 73
            job_instance = "0" & this_instance
        End If

        If script_update_stat = "Yes - Update an existing JOBS Panel" OR script_update_stat = "Yes - Create a new JOBS Panel" Then
            MsgBox "Update - " & script_update_stat & " - 3"
            EMWriteScreen left(job_income_type, 1), 5, 34       'income type
            If job_subsidized_income_type <> "" Then EMWriteScreen left(job_subsidized_income_type, 2), 5, 74           'subsidized type
            If job_verification <> " " Then EMWriteScreen left(job_verification, 1), 6, 34              'job verification
            EMWriteScreen "                              ", 7, 42           'blank out the employer name
            EMWriteScreen job_employer_name, 7, 42                          'enter the employer name
            If IsDate(job_income_start) = TRUE Then Call create_mainframe_friendly_date(job_income_start, 9, 35, "YY")      'income start date
            If IsDate(job_income_end) = TRUE Then Call create_mainframe_friendly_date(job_income_end, 9, 49, "YY")          'income end date
            If IsDate(contract_through_date) = TRUE then call create_mainframe_friendly_date(contract_through_date, 9, 73, "YY")
            EMWriteScreen left(job_pay_frequency, 1), 18, 35                'pay frequency
            ' If original_full_jobs_name = "" Then EMReadScreen original_full_jobs_name, 30, 7, 42             'read the employer name as it originally exists on the panel

            If job_change_type = "Income/Hours Chnage for Current Job" Then
                EMReadScreen job_amount_one, 8, 12, 67
                EMReadScreen job_amount_two, 8, 13, 67
                EMReadScreen job_amount_three, 8, 14, 67
                EMReadScreen job_amount_four, 8, 15, 67
                EMReadScreen job_amount_five, 8, 16, 67
                EMReadScreen job_hours, 3, 18, 72

                numb_of_checks = 0
                check_amount = ""
                total_hours = ""
                If job_amount_one <> "________" then
                    numb_of_checks = 1
                    check_amount = trim(job_amount_one)
                End If
                If job_amount_two <> "________" then
                    numb_of_checks = 2
                    check_amount = trim(job_amount_two)
                End If
                If job_amount_three <> "________" then
                    numb_of_checks = 3
                    check_amount = trim(job_amount_three)
                End If
                If job_amount_four <> "________" then
                    numb_of_checks = 4
                    check_amount = trim(job_amount_four)
                End If
                If job_amount_five <> "________" then
                    numb_of_checks = 5
                    check_amount = trim(job_amount_five)
                End If
                job_hours = job_hours * 1
                total_hours = job_hours
                hours_per_check = job_hours / numb_of_checks
            End If
            'Blank the pay information
            If job_change_type = "Job Ended" Then
                EMReadScreen known_pay_amount, 8, 12, 67
                known_pay_amount = trim(known_pay_amount)
            End If
            jobs_row = 12
            Do
                EMWriteScreen "  ", jobs_row, 25            'retro side
                EMWriteScreen "  ", jobs_row, 28
                EMWriteScreen "  ", jobs_row, 31
                EMWriteScreen "        ", jobs_row, 38

                EMWriteScreen "  ", jobs_row, 54            'prospective side
                EMWriteScreen "  ", jobs_row, 57
                EMWriteScreen "  ", jobs_row, 60
                EMWriteScreen "        ", jobs_row, 67

                jobs_row = jobs_row + 1
            Loop until jobs_row = 17
            EMWriteScreen "  ", 18, 43                      'blanking out hours
            EMWriteScreen "  ", 18, 72

            Select Case job_change_type

            ' DropListBox 125, 65, 150, 45, list_of_employees, job_employee
            ' EditBox 350, 65, 50, 15, work_numb_id
            ' DropListBox 125, 85, 110, 15, ""+chr(9)+"W - Wages"+chr(9)+"J - WIOA"+chr(9)+"E - EITC"+chr(9)+"G - Experience Works"+chr(9)+"F - Federal Work Study"+chr(9)+"S - State Work Study"+chr(9)+"O - Other"+chr(9)+"C - Contract Income"+chr(9)+"T - Training Program"+chr(9)+"P - Service Program"+chr(9)+"R - Rehab Program", job_income_type
            ' DropListBox 335, 85, 150, 15, ""+chr(9)+"01 - Subsidized Public Sector Employer"+chr(9)+"02 - Subsidized Private Sector Employer"+chr(9)+"03 - On-the-Job-Training"+chr(9)+"04 - AmeriCorps", job_subsidized_income_type
            ' EditBox 125, 105, 140, 15, job_employer_name
            ' DropListBox 315, 105, 55, 45, " "+chr(9)+"N - No Verif Provided"+chr(9)+"? - Delayed Verification", job_verification
            ' DropListBox 435, 105, 85, 45, ""+chr(9)+"1 - Monthly"+chr(9)+"2 - Semi-Monthly"+chr(9)+"3 - Biweekly"+chr(9)+"4 - Weekly"+chr(9)+"5 - Other", job_pay_frequency
            ' EditBox 125, 125, 55, 15, job_income_start
            ' EditBox 280, 125, 55, 15, job_income_end
            ' EditBox 430, 125, 55, 15, job_contract_through_date


                Case "New Job Reported"
                    If trim(new_job_hourly_wage) <> "" Then EMWriteScreen new_job_hourly_wage, 6, 75       'hourly wage

                    the_last_pay_date = new_job_income_start
                    the_first_pay_date = new_job_income_start
                    end_of_pay = "99/99/99"
                    pay_amt = "0"
                    ' prosp_hours = "000"
                Case "Income/Hours Chnage for Current Job"
                    If trim(job_change_new_hourly_wage) <> "" Then EMWriteScreen job_change_new_hourly_wage, 6, 75       'hourly wage

                    the_last_pay_date = first_pay_date_of_change
                    the_first_pay_date = first_pay_date_of_change
                    end_of_pay = "99/99/99"
                    job_change_new_hourly_wage = job_change_new_hourly_wage * 1
                    job_change_new_hours_per_week = job_change_new_hours_per_week * 1
                    If job_pay_frequency = "1 - Monthly" Then
                        pay_amt = job_change_new_hourly_wage * job_change_new_hours_per_week * 4.3

                    ElseIf job_pay_frequency = "2 - Semi-Monthly" Then
                        pay_amt = job_change_new_hourly_wage * job_change_new_hours_per_week * 2.15
                    ElseIf job_pay_frequency = "3 - Biweekly" Then
                        pay_amt = job_change_new_hourly_wage * job_change_new_hours_per_week * 2
                    ElseIf job_pay_frequency = "4 - Weekly" Then
                        pay_amt = job_change_new_hourly_wage * job_change_new_hours_per_week
                    End If
                    ' pay_amt =
                    prosp_hours = 0
                Case "Job Ended"
                    ' EMWriteScreen "      ", 6, 75       'hourly wage
                    end_of_pay = "99/99/99"
                    If IsDate(job_end_income_end_date) = TRUE Then end_of_pay = job_end_income_end_date
                    the_last_pay_date = job_end_income_end_date
                    the_first_pay_date = job_end_income_end_date
                    prosp_hours = 0
                    pay_amt = known_pay_amount

                    ' EditBox 80, 165, 55, 15, date_work_ended
                    ' EditBox 250, 165, 50, 15, job_end_income_end_date
                    ' EditBox 380, 165, 50, 15, last_pay_amount
                    ' DropListBox 490, 165, 30, 45, "Yes"+chr(9)+"No", refused_empl_yn
                    ' EditBox 80, 185, 130, 15, stwk_reason
                    ' DropListBox 275, 185, 45, 45, "Yes"+chr(9)+"No", vol_quit_yn
                    ' DropListBox 450, 185, 45, 45, "Yes"+chr(9)+"No", good_cause_yn
                    ' ComboBox 80, 205, 135, 45, list_of_members, conversation_with_person
                    ' EditBox 260, 205, 260, 15, conversation_detail
                    ' DropListBox 180, 225, 45, 45, "Yes"+chr(9)+"No", uc_yn
                    ' EditBox 330, 225, 190, 15, wreg_abawd_notes
                    ' CheckBox 80, 245, 190, 10, "Check here if you sent a Work Number request.", work_number_checkbox
                    ' ComboBox 80, 260, 125, 45, "Select One..."+chr(9)+"1 - Employers Statement"+chr(9)+"2 - Seperation Notice"+chr(9)+"3 - Collateral Statement"+chr(9)+"4 - Other Document"+chr(9)+"N - No Verif Provided", stwk_verif
            End Select

            If end_of_pay = "99/99/99" Then
                jobs_row = 12
            Else
                jobs_row = 16
                Call create_mainframe_friendly_date(job_end_income_end_date, 9, 49, "YY")
            End If
            If Initial_footer_month = MAXIS_footer_month AND Initial_footer_year = MAXIS_footer_year Then
                Call create_mainframe_friendly_date(the_first_pay_date, jobs_row, 54, "YY")
                If end_of_pay = "99/99/99" Then
                    EMWriteScreen pay_amt, jobs_row, 67
                    jobs_row = jobs_row + 1
                Else
                    EMWriteScreen last_pay_amount, jobs_row, 67
                    jobs_row = jobs_row - 1
                End If
            End If
            next_month_mo = ""
            next_month_yr = ""
            the_month_here = DateValue(MAXIS_footer_month & "/01/" & MAXIS_footer_year)
            the_next_month = DateAdd("m", 1, the_month_here)
            Call convert_date_into_MAXIS_footer_month(the_next_month, next_month_mo, next_month_yr)

            MsgBox "The next month - " & next_month_mo & "/" & next_month_yr & vbCr & vbCr & "MAXIS month - " & MAXIS_footer_month & "/" & MAXIS_footer_year
            Do
                the_pay_date = ""
                the_pay_date_two = ""
                If end_of_pay = "99/99/99" Then
                    If job_pay_frequency = "1 - Monthly" Then
                        the_pay_date = DateAdd("m", 1, the_last_pay_date)
                        prosp_hours = prosp_hours + (job_change_new_hours_per_week * 4.3)
                    ElseIf job_pay_frequency = "2 - Semi-Monthly" Then
                        the_pay_date = DateAdd("d", 15, the_last_pay_date)
                        the_pay_date_two = DateAdd("m", 1, the_last_pay_date)
                        prosp_hours = prosp_hours + (job_change_new_hours_per_week * 2.15)
                    ElseIf job_pay_frequency = "3 - Biweekly" Then
                        the_pay_date = DateAdd("d", 14, the_last_pay_date)
                        prosp_hours = prosp_hours + (job_change_new_hours_per_week * 2)
                    ElseIf job_pay_frequency = "4 - Weekly" Then
                        the_pay_date = DateAdd("d", 7, the_last_pay_date)
                        prosp_hours = prosp_hours + job_change_new_hours_per_week
                    End If
                Else
                    If job_pay_frequency = "1 - Monthly" Then
                        the_pay_date = DateAdd("m", -1, the_last_pay_date)
                    ElseIf job_pay_frequency = "2 - Semi-Monthly" Then
                        the_pay_date = DateAdd("d", -15, the_last_pay_date)
                        the_pay_date_two = DateAdd("m", -1, the_last_pay_date)
                    ElseIf job_pay_frequency = "3 - Biweekly" Then
                        the_pay_date = DateAdd("d", -14, the_last_pay_date)
                    ElseIf job_pay_frequency = "4 - Weekly" Then
                        the_pay_date = DateAdd("d", -7, the_last_pay_date)
                    End If
                End If

                Call convert_date_into_MAXIS_footer_month(the_pay_date, pay_date_mo, pay_date_yr)
                If IsDate(the_pay_date_two) = TRUE Then Call convert_date_into_MAXIS_footer_month(the_pay_date_two, pay_date_two_mo, pay_date_two_yr)
                MsgBox "The pay date - " & the_pay_date & vbCr & vbCr & "Pay month - " & pay_date_mo & "/" & pay_date_yr & vbCr & "MAXIS month - " & MAXIS_footer_month & "/" & MAXIS_footer_year & vbCr & vbCr & "JOBS Row - " & jobs_row
                If pay_date_mo = MAXIS_footer_month AND pay_date_yr = MAXIS_footer_year Then
                    Call create_mainframe_friendly_date(the_pay_date, jobs_row, 54, "YY")
                    EMWriteScreen pay_amt, jobs_row, 67
                    If end_of_pay = "99/99/99" Then
                        jobs_row = jobs_row + 1
                    Else
                        jobs_row = jobs_row - 1
                        pay_amt = pay_amt * 1
                        job_hourly_wage = job_hourly_wage * 1
                        hours_of_pay = pay_amt/job_hourly_wage
                        prosp_hours = prosp_hours + hours_of_pay
                    End If
                End If
                If pay_date_two_mo = MAXIS_footer_month AND pay_date_two_yr = MAXIS_footer_year Then
                    Call create_mainframe_friendly_date(the_pay_date_two, jobs_row, 54, "YY")
                    EMWriteScreen pay_amt, jobs_row, 67
                    If end_of_pay = "99/99/99" Then
                        jobs_row = jobs_row + 1
                    Else
                        jobs_row = jobs_row - 1
                        pay_amt = pay_amt * 1
                        job_hourly_wage = job_hourly_wage * 1
                        hours_of_pay = pay_amt/job_hourly_wage
                        prosp_hours = prosp_hours + hours_of_pay
                    End If
                End If
                the_last_pay_date = the_pay_date
                If the_pay_date_two <> "" Then the_last_pay_date = the_pay_date_two
                ' job_end_income_end_date
                If job_change_type = "Job Ended" Then
                    If pay_date_mo <> MAXIS_footer_month Then Exit Do
                End If
            Loop until pay_date_mo = next_month_mo AND pay_date_yr = next_month_yr OR pay_date_two_mo = next_month_mo AND pay_date_two_yr = next_month_yr
            checks_entered = jobs_row - 12
            If end_of_pay = "99/99/99" Then
                checks_entered = jobs_row - 12
            Else
                checks_entered = abs(jobs_row - 16)
            End If
            If job_change_type = "New Job Reported" Then prosp_hours = "000"
            EMWriteScreen "   ", 18, 72
            EMWriteScreen prosp_hours, 18, 72
            If SNAP_active = TRUE Then          'Open the SNAP PIC'
                EMWriteScreen "X", 19, 38
                transmit

                Call create_mainframe_friendly_date(date, 5, 34, "YY")
                EMWriteScreen left(job_pay_frequency, 1), 18, 35

                Select Case job_change_type
                    Case "New Job Reported"
                        EMWriteScreen "      ", 8, 64
                        If trim(new_job_hourly_wage) <> "" Then
                            EMWriteScreen "0", 8, 64
                            EMWriteScreen new_job_hourly_wage, 9, 66       'hourly wage
                        End If

                    Case "Income/Hours Chnage for Current Job"
                        If trim(job_change_old_hourly_wage) <> "" Then
                            EMWriteScreen job_change_old_hours_per_week, 8, 64
                            EMWriteScreen job_change_old_hourly_wage, 9, 66       'hourly wage
                        End If

                End Select
            End If
        End If
        transmit        'after all of the panels have been reviewed we are going to STAT/WRAP to get to the next month without sending through background if possible
        If original_full_jobs_name = "" Then EMReadScreen original_full_jobs_name, 30, 7, 42             'read the employer name as it originally exists on the panel

        If script_update_stat = "Yes - Update an existing JOBS Panel" OR script_update_stat = "Yes - Create a new JOBS Panel" Then
            If job_change_type = "Job Ended" Then
                Call navigate_to_MAXIS_screen("STAT", "STWK")
                EMWriteScreen ref_nbr, 20, 76
                transmit

                EMReadScreen version_of_stwk, 1, 2, 73
                If version_of_stwk = "1" Then
                    PF9
                End If
                If version_of_stwk = "0" Then
                    EMWriteScreen "NN", 20, 79
                End If

                EMWriteScreen "                              ", 6, 46
                EMWriteScreen job_employer_name, 6, 46

                Call create_mainframe_friendly_date(date_work_ended, 7, 46, "YY")
                Call create_mainframe_friendly_date(job_end_income_end_date, 8, 46, "YY")

                EMWriteScreen stwk_verif, 7, 63
                EMWriteScreen refused_empl_yn, 8, 78
                If IsDate(refused_empl_date) = TRUE Then Call create_mainframe_friendly_date(refused_empl_date, 10, 72, "YY")
                EMWriteScreen vol_quit_yn, 10, 46
                If adult_cash_case = TRUE or family_cash_case = TRUE Then EMWriteScreen good_cause_yn, 12, 52
                If grh_case = TRUE Then EMWriteScreen good_cause_yn, 12, 60
                If snap_case = TRUE Then EMWriteScreen good_cause_yn, 12, 67

                If pwe_ref = ref_nbr Then
                    EMWriteScreen "Y", 14, 46
                Else
                    EMWriteScreen "N", 14, 46
                End If
                transmit
                EMReadScreen error_prone_warning, 20, 6, 43
                If error_prone_warning = "Error Prone Warnings" Then transmit

            End If

            If vol_quit_yn = "Yes" AND good_cause_yn = "No" Then
                'TODO - add functionality to update DISQ once vol quit is a thing again.'
            End If
        End If

        transmit
        EmWriteScreen "BGTX", 20, 71
        transmit

        EMReadScreen wrap_check, 4, 2, 46
        If wrap_check <> "WRAP" Then

        End If
        EMWriteScreen "Y", 16, 54
        If MAXIS_footer_month = CM_plus_1_mo AND MAXIS_footer_year = CM_plus_1_yr Then EMWriteScreen "N", 16, 54
        transmit
        MsgBox "Pause here - should be in the next month."

        second_loop = TRUE
    Loop until MAXIS_footer_month = CM_plus_1_mo AND MAXIS_footer_year = CM_plus_1_yr
Else

    Do
        Do
            err_msg = ""

            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 316, 175, "Send Job Change Information as Email"
              CheckBox 15, 40, 245, 10, "Check here to have the script send the job change information via email", send_email_checkbox
              EditBox 25, 75, 115, 15, email_address_one
              EditBox 25, 95, 115, 15, email_address_two
              EditBox 25, 115, 115, 15, email_address_three
              EditBox 25, 135, 115, 15, email_address_four
              EditBox 25, 155, 115, 15, email_address_five
              ButtonGroup ButtonPressed
                OkButton 205, 155, 50, 15
                CancelButton 260, 155, 50, 15
              Text 10, 10, 305, 20, "Since this script was run in Inquiry, if this information should be provided to another worker within Hennepin County the scrapt can send an email to up to five individuals or teams."
              Text 25, 60, 150, 10, "Email Addresses to send the information to:"
              Text 145, 80, 50, 10, "@hennepin.us"
              Text 145, 100, 50, 10, "@hennepin.us"
              Text 145, 120, 50, 10, "@hennepin.us"
              Text 145, 140, 50, 10, "@hennepin.us"
              Text 145, 160, 50, 10, "@hennepin.us"
            EndDialog

            dialog Dialog1
            cancel_confirmation

            If err_msg <> "" Then MsgBox "Please resolve to continue: " & vbNewLine & err_msg

        Loop until err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE

    If send_email_checkbox = checked Then
        all_email_recipients = email_address_one
        all_email_recipients = all_email_recipients & "; " & email_address_two
        all_email_recipients = all_email_recipients & "; " & email_address_three
        all_email_recipients = all_email_recipients & "; " & email_address_four
        all_email_recipients = all_email_recipients & "; " & email_address_five
        Do
            end_characters = right(all_email_recipients, 2)
            If end_characters = "; " Then all_email_recipients = left(all_email_recipients, len(all_email_recipients) - 2)
        Loop until

        Call create_outlook_email(all_email_recipients, "", "Job Change Reported for MX Case", email_body, email_attachment, send_email)
    End If
End If
'EMAIL BETWEEN ES
'MEMO WITH EARNED INCOME DETAIL ON VERIFYING



If job_change_type = "New Job Reported" Then verif_type_requested = "new job"
If job_change_type = "Income/Hours Chnage for Current Job" Then verif_type_requested = "change in current job"
If job_change_type = "Job Ended" Then verif_type_requested = "job ended"
If TIKL_checkbox = checked Then Call create_TIKL("Verification of " & verif_type_requested & " due.", 10, verif_form_date, TURE, TIKL_note_text)

Call start_a_blank_CASE_NOTE

Call write_variable_in_CASE_NOTE("Change in Income Reported: " & UCase(verif_type_requested) & " - Reported on: " & reported_date)
Call write_variable_in_CASE_NOTE("=== Details of the Reported Change ===")
Call write_variable_in_CASE_NOTE("* Job Name: " & job_employer_name & " - Income Type: " & job_income_type)
Call write_variable_in_CASE_NOTE("* Employee: " & job_employee)
Call write_variable_in_CASE_NOTE("* Income Start Date: " & job_income_start & " - End Date: " & job_income_end)
Call write_variable_in_CASE_NOTE("* Verification: " & job_verification)
Call write_variable_in_CASE_NOTE("*** Type of Change: " & job_change_type)
Select Case job_change_type
    Case "New Job Reported"
        Call write_variable_with_indent_in_CASE_NOTE("Work Start Date: " & date_work_started)
        Call write_variable_with_indent_in_CASE_NOTE("Income Start Date: " & new_job_income_start & " - Initial Gross Pay: $" & initial_check_gross_amount)
        Call write_variable_with_indent_in_CASE_NOTE("Anticipated Income: Hours per Week: " & new_job_hours_per_week & " - Hourly Wage: " & new_job_hourly_wage)
    Case "Income/Hours Chnage for Current Job"
        Call write_variable_with_indent_in_CASE_NOTE("Change happened on " & job_change_date & " change will cause " & income_change_type)
        Call write_variable_with_indent_in_CASE_NOTE("Date of pay first impacted: " & first_pay_date_of_change)
        Call write_variable_with_indent_in_CASE_NOTE("Change: " & job_change_details)
        Call write_variable_with_indent_in_CASE_NOTE("Previous Income: Hours per Week: " & job_change_old_hours_per_week & " - Hourly Wage: " & job_change_old_hourly_wage)
        Call write_variable_with_indent_in_CASE_NOTE("New Income: Hours per Week: " & job_change_new_hours_per_week & " - Hourly Wage: " & job_change_new_hourly_wage)
    Case "Job Ended"
        Call write_variable_with_indent_in_CASE_NOTE("Income End Date: " & job_end_income_end_date & " - Final pay amount: "& last_pay_amount)
        Call write_variable_with_indent_in_CASE_NOTE("Work stoped on: " & date_work_ended)
        Call write_variable_with_indent_in_CASE_NOTE("Quit details:")
        Call write_variable_with_indent_in_CASE_NOTE(" - Employee refused employment: " & refused_empl_yn)
        Call write_variable_with_indent_in_CASE_NOTE(" - Was this a voluntary quit? " & vol_quit_yn)
        Call write_variable_with_indent_in_CASE_NOTE(" - Reason for STWK: " & stwk_reason)
        Call write_variable_with_indent_in_CASE_NOTE("    - Meets good cause? " & good_cause_yn)
End Select
Call write_variable_in_CASE_NOTE("* Impact on WREG/ABAWD: " & wreg_abawd_notes)
Call write_variable_in_CASE_NOTE("Information about job gathered in conversation with " & conversation_with_person)
Call write_variable_in_CASE_NOTE("  - Details of conversation: " & conversation_detail)
Call write_variable_in_CASE_NOTE("=== Reporting Information ===")
Call write_variable_in_CASE_NOTE("* Reported via " & job_report_type & " by " & person_who_reported_job & " on " & reported_date)
If work_number_verbal_checkbox = checked Then Call write_variable_in_CASE_NOTE("*** Verbal authorization to check the Work Number received.")
Call write_variable_in_CASE_NOTE("=== Verification ===")
If work_number_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Sent Work Number request for income verification.")
Call write_variable_in_CASE_NOTE("* Verification request sent on " & verif_form_date)
Call write_variable_in_CASE_NOTE("* Time frame of income verification requested: " & verif_time_frame)
If requested_CEI_OHI_docs_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Requested Health Insurance information available from employer.")
Call write_variable_in_CASE_NOTE("---")
Call write_bullet_and_variable_in_CASE_NOTE("NOTES", notes)
Call write_variable_in_CASE_NOTE(TIKL_note_text)
Call write_variable_in_CASE_NOTE(worker_signature)

call script_end_procedure_with_error_report("Success! Job change noted and updates made.")
' job_update_date
' job_instance
' job_ref_number
' job_employee
' job_income_type
' job_verification
' job_subsidized_income_type
' job_hourly_wage
' job_employer_name
' job_employer_full_name
' job_income_start
' job_income_end
' job_pay_frequency
'
' job_pic_calculation_date
' job_pic_pay_frequency
' job_pic_hours_per_week
' job_pic_pay_per_hour


' Select Case job_change_type
'     Case "New Job Reported"
'
'     Case "Income/Hours Chnage for Current Job"
'
'     Case "Job Ended"
'
' End Select

'Navigating to DAIL/WRIT
If TIKL_checkbox = 1 then
	script_end_procedure("Success! MAXIS updated for job change, a case note made, and a TIKL has been sent for 10 days from now. An EV should now be sent. The job is at: " & employer & ".")
Else
	script_end_procedure("Success! MAXIS updated for job change, and a case note has been made. An EV should now be sent. The job is at: " & employer & ".")
END IF
