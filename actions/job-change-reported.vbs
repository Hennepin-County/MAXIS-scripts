'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - JOB CHANGE REPORTED.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 345                	'manual run time in seconds
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
call changelog_update("05/28/2020", "Added virtual drop box information to SPEC/MEMO.", "MiKayla Handley, Hennepin County")
call changelog_update("04/24/2020", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
'connecting to MAXIS & grabbing the case number
EMConnect ""
call MAXIS_case_number_finder(MAXIS_case_number)

'This initial dialog is to get the rest of the run set up. We need to know what kind of change is being reported.
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 276, 110, "Job Change Selection"
  EditBox 130, 10, 60, 15, MAXIS_case_number
  DropListBox 130, 30, 140, 45, "Select One ..."+chr(9)+"New Job Reported"+chr(9)+"Income/Hours Change for Current Job"+chr(9)+"Job Ended", job_change_type
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
	    dialog Dialog1
	    cancel_without_confirmation

        Call validate_MAXIS_case_number(err_msg, "*")               'case number mandatory
        If job_change_type = "Select One ..." Then err_msg = err_msg & vbNewLine & "* Indicate what type of change is being reported (stop, start, or change)."         'change type mandatory
        If trim(worker_signature) = "" Then err_msg = err_msg & vbNewLine & "* Enter your name for the CASE/NOTE."          'worker signature mandatory

        If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
    Loop until err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

If job_change_type = "New Job Reported" Then change_text = "new job"                                        'creating text for easy view in the next dialog - this is not functional - just formatting.
If job_change_type = "Income/Hours Change for Current Job" Then change_text = "change job income"
If job_change_type = "Job Ended" Then change_text = "stwk"

'This dialog ensures we are processing non-verified reports - if we have verification the script will redirect to Earned Income Budgeting
'Earned Income budgeting has functionality to CORRECTLY update JOBS and CNOTE with complete information
Dialog1 = ""
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

'Here, if there are verifications indicated, we will be using Earned Income Budgeting. This will alert the worker to this fact AND redirect.
If verifs_received_selection = "Yes - sufficient verification received to budget income accurately." Then
    MsgBox "Since verifications have been received for a " & change_text & ", the " & vbNewLine & vbNewLine & "             ACTIONS - Earned Income Budgeting       " & vbNewLine & vbNewLine &"script will run to accurately update and CASE/NOTE the budgeted income."
    Call run_from_GitHub(script_repository & "actions/earned-income-budgeting.vbs")
End If

Call back_to_SELF                       'go back to self to read the region.
developer_mode = FALSE                  'defaulting developer to false because we are usually not in developer mode
EMReadScreen MX_region, 12, 22, 48      'reading for the region
MX_region = trim(MX_region)             'formatting the region
If MX_region = "INQUIRY DB" Then        'This is what INQUIRY looks like on SELF.
    'We are going to confirm HERE that the worker meant to run this in inquiry If not, the script run will end.
    continue_in_inquiry = MsgBox("It appears you are in INQUIRY. Income information cannot be saved to STAT and a CASE/NOTE cannot be created." & vbNewLine & vbNewLine & "Do you wish to continue?", vbQuestion + vbYesNo, "Continue in Inquiry?")
    If continue_in_inquiry = vbNo Then script_end_procedure("Script ended since it was started in Inquiry.")
    developer_mode = TRUE               'If thes cript didn't end and we were in inquiry, we are automatically in developer mode
End If
If worker_signature = "UUDDLRLRBA" Then developer_mode = TRUE           'Use of the konami code in worker_signature will also cause the script to run in developer mode
If developer_mode = TRUE then MsgBox "Developer Mode ACTIVATED!"        'developer mode prevents actions from being taken in MAXIS
If developer_mode = TRUE Then script_run_lowdown = "Run in INQUIRY" & vbCr & vbCr     'adding this to any error reporting

script_run_lowdown = script_run_lowdown & "Change type: " & job_change_type & vbCr                      'adding information to pass through to any possible error report
If job_panel_exists = checked Then
    script_run_lowdown = script_run_lowdown & "The panel for the job already exists in MAXIS." & vbCr
Else
    script_run_lowdown = script_run_lowdown & "The panel for the job DOES NOT exist in MAXIS." & vbCr
End If

MAXIS_footer_month = CM_mo                      'Defaulting the MAXIS footer month and year to the current month
MAXIS_footer_year = CM_yr
Call MAXIS_background_check                     'Making sure we can get into STAT

Call generate_client_list(list_of_employees, "Select One ...")          'Using the client list functionality the script will read STAT for all the household members to populate droplist box
call generate_client_list(list_of_members, "Type or Select")
client_name_array = split(list_of_members, chr(9))                      'creating an array of the HH Members from the list created for the droplist box

CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.

DO								              'reading all the reference number to make an array of the reference numbers.
    EMReadscreen ref_nbr, 3, 4, 33

    client_string = client_string & "|" & ref_nbr

    transmit            'this goes to the nest MEMB panel
    Emreadscreen edit_check, 7, 24, 2       'reading for if we made it to the last of the MEMB panels.
LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

client_string = trim(client_string)         'turning the list of reference numbers into the array'
client_string = right(client_string, len(client_string) - 1)
ref_nbr_array = split(client_string, "|")

'If the client indicated the JOB panel already exists, the script will try to determine which JOBS panel should be used.
If job_panel_exists = checked Then
    job_selected = FALSE                        'defaulting a variable to indicate if the panel has been found to FALSE
    For each clt_number in ref_nbr_array        'looping through all the reference numbers so that we can check all the members for JOBS panels.

        Call Navigate_to_MAXIS_screen("STAT", "JOBS")                       'Go to JOBS
        EMWriteScreen clt_number, 20, 76                                    'Enter the reference for the clients in turn to check each.
        EMWriteScreen "01", 20, 79
        transmit

        EMReadScreen total_jobs, 1, 2, 78                                   'look for how many JOBS panels there are so we can loop through them all
        If total_jobs <> "0" Then       'if there are no JOBS panels listed for this member, we should't try to read them.
            Do
                EMReadScreen employer, 30, 7, 42                'reading the employer name
                employer = replace(employer, "_", "")           'taking out the underscores

                'We have to ask the worker if this is the JOB. This messagebox will show up for EVERY job found until they click 'Yes'
                employer_check = MsgBox("Is this the job reported? Employer name: " & employer, vbYesNo + vbQuestion, "Select Income Panel")
                If employer_check = vbYes Then      'If we find the job - then we leave all the loops
                    job_selected = TRUE             'This is how we know we found the job now.'
                    Exit Do
                End If
                transmit
                EMReadScreen last_job, 7, 24, 2
            Loop until last_job = "ENTER A"
        End If
        If job_selected = TRUE Then Exit For
    Next

    'Once we find the job, we are going to gather all the detail from the panel.
    If job_selected = TRUE Then
        EMReadScreen job_update_date, 8, 21, 55                 'Date listed from MAXIS that list when the panel was last updated.
        EMReadScreen job_instance, 1, 2, 73                     'Reading the instance indicator so we can go back to this panel when we need to.

        EMReadScreen job_ref_number, 2, 4, 33                   'Reference number listed on the panel
        EMReadScreen job_income_type, 1, 5, 34                  'Job income type - usually Wages
        EMReadScreen job_verification, 1, 6, 34                 'Reading teh current verification code.
        EMReadScreen job_subsidized_income_type, 2, 5, 74       'Reading the subsidized income type in case one is listed.
        EMReadScreen job_hourly_wage, 6, 6, 75                  'The hourly wage can be entered on the panel - reading if it is listed there.
        EMReadScreen job_employer_full_name, 30, 7, 42          'Reading the employer name without removing the underscores so that we can make sure we are matching
        EMReadScreen job_income_start, 8, 9, 35                 'Reading the date the income started
        EMReadScreen job_income_end, 8, 9, 49                   'Reading for the date the income ended if it already has a date end listed.
        EMReadScreen job_pay_frequency, 1, 18, 35               'Pay frequency from main JOBS panel

        EMWriteScreen "X", 19, 38           'opening the PIC'
        transmit

        EMReadScreen job_pic_calculation_date, 8, 5, 34         'Getting the detail already listed on the PIC
        EMReadScreen job_pic_pay_frequency, 1, 5, 64
        EMReadScreen job_pic_hours_per_week, 6, 8, 64
        EMReadScreen job_pic_pay_per_hour, 8, 9, 66

        PF3                                 'leaving the PIC

        job_update_date = replace(job_update_date, " ", "/")        'formatting the date to look like a date
        job_instance = "0" & job_instance                           'making the insance 2 digits
        For each person in client_name_array                        'looping through the people in the client array that also has reference numbers listed in it to make the employer name have the reference number AND full name from STAT
            If left(person, 2) = job_ref_number Then job_employee =  person
        Next
        If job_income_type = "J" Then job_oncome_type = "J - WIOA"              'formatting the income type to have the code and PF1 information
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

        If job_verification = "1" Then job_verification = "1 - Pay Stubs/Tip Report"        'formatting the verification to have the code and PF1 information
        If job_verification = "2" Then job_verification = "2 - Employer Statement"
        If job_verification = "3" Then job_verification = "3 - Collateral Statement"
        If job_verification = "4" Then job_verification = "4 - Other Document"
        If job_verification = "5" Then job_verification = "5 - Pend Out State Verif"
        If job_verification = "N" Then job_verification = "N - No Verif Provided"
        If job_verification = "?" Then job_verification = "? - Delayed Verification"

        If job_subsidized_income_type = "__" Then job_subsidized_income_type = ""               'formatting the verification to have the code and PF1 information
        If job_subsidized_income_type = "01" Then job_subsidized_income_type = "01 - Subsidized Public Sector Employer"
        If job_subsidized_income_type = "02" Then job_subsidized_income_type = "02 - Subsidized Private Sector Employer"
        If job_subsidized_income_type = "03" Then job_subsidized_income_type = "03 - On-the-Job-Training"
        If job_subsidized_income_type = "04" Then job_subsidized_income_type = "04 - AmeriCorps"

        job_hourly_wage = trim(job_hourly_wage)                     'formatting the information from the panel to be readable in a dialog/note
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

'This function will read what programs are on the case and what the case status is.
Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status)

For each clt_number in ref_nbr_array                        'Now we are going to check WREG for the PWE

    Call Navigate_to_MAXIS_screen("STAT", "WREG")           'Go to WREG for each client on the case in turn
    EMWriteScreen clt_number, 20, 76
    transmit

    EMReadScreen pwe_code, 1, 6, 68                         'Read PWE code
    If pwe_code = "Y" Then                                  'If it is 'Y' - save the reference number to identify the PWE
        pwe_ref = clt_number
        Exit For
    End If
Next

'Here are some presets for the main dialog
If job_change_type = "New Job Reported" Then                'If the panel already exists and it has an income start date, we can default to that date
    If IsDate(job_income_start) = TRUE Then new_job_income_start = job_income_start
End If
If job_change_type = "Job Ended" Then                       'If the panel exists and there is already an end date, we can default to that date
    If IsDate(job_income_end) = TRUE Then job_end_income_end_date = job_income_end
End If
If job_verification = "" Then job_verification = "N - No Verif Provided"        'Setting the verification to 'N'

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
reported_date = date & ""                                   'defaulting the reported date to today
script_update_stat = "Yes - Create a new JOBS Panel"        'Defaulting to having the script create and update a new panel
If job_panel_exists = checked Then script_update_stat = "Yes - Update an existing JOBS Panel"   'If the worker indicated the panel exists then defaulting updating an existing panel
If developer_mode = TRUE Then script_update_stat = "No - Update of JOBS not needed"             'If we are in developer mode, then we cannot update

'This didalog has a different middle part based on the type of report that is happening
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 536, 370, "Job Change Details Dialog"
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
      Case "Income/Hours Change for Current Job"
          GroupBox 5, 150, 525, 130, "Update Reported - JOB CHANGE"
          Text 15, 165, 55, 10, "Date of Change:"
          Text 145, 165, 70, 10, "* Change Reported:"
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

Do                      'Showing the main dialog
    Do
        err_msg = ""

        dialog Dialog1
        cancel_confirmation

        If job_change_type = "New Job Reported" Then                'Income start date shows up twice on this dialog when the job change is new job.
            If IsDate(job_income_start) = TRUE Then                 'This code will autofill one from the other if only one is completed an make sure they match
                If IsDate(new_job_income_start) = TRUE Then
                    If job_income_start <> new_job_income_start Then err_msg = err_msg & vbNewLine & "* The income start dates do not match. Review the income start dates"
                Else
                    new_job_income_start = job_income_start
                End If
            Else
                If IsDate(new_job_income_start) = TRUE Then job_income_start = new_job_income_start
            End If
        End If

        'The rest of the error handing. There is a lot here because we need to gether more information about job changes.
        If job_report_type = "Type or Select" OR trim(job_report_type) = "" Then err_msg = err_msg & vbNewLine & "* Indicate how the income was reported."
        If person_who_reported_job = "Type or Select" or trim(person_who_reported_job) = "" Then err_msg = err_msg & vbNewLine & "* Select the household member, or type in the name of the person who reported the job."
        If IsDate(reported_date)  = False Then err_msg = err_msg & vbNewLine & "* Enter a valid date for when this job was reported."
        err_msg = err_msg & vbNewLine

        If job_employee = "Select One ..." Then err_msg = err_msg & vbNewLine & "* Select the household member who is the employee."
        If trim(job_employer_name) = "" Then err_msg = err_msg & vbNewLine & "* Enter the name of the employer."
        If len(trim(job_employer_name)) > 30 Then err_msg = err_msg & vbNewLine & "* The employer name is more than 30 characters and MAXIS only allows for 30 characters on the employer line. Change the employer name to fit MAIXS."
        If job_verification = "" Then err_msg = err_msg & vbNewLine & "* Enter the verification of the JOBS panel."
        If job_pay_frequency = " " Then err_msg = err_msg & vbNewLine & "* Enter the frequency of the pay (weekly, biweekly, monthly)."
        If IsDate(job_income_start) = False Then err_msg = err_msg & vbNewLine & "* Enter the date the income will or has started."
        err_msg = err_msg & vbNewLine

        Select Case job_change_type
            Case "New Job Reported"
                If trim(new_job_hourly_wage) <> "" AND IsNumeric(new_job_hourly_wage) = False Then err_msg = err_msg & vbNewLine & "* Enter the hourly wage as a number."
                If trim(new_job_hours_per_week) <> "" AND IsNumeric(new_job_hours_per_week) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the hours per week as a number."

            Case "Income/Hours Change for Current Job"
                If IsDate(first_pay_date_of_change) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the date the change will be reflected in income (the first paycheck that will be affected by this change)."
                If trim(job_change_details) = "" Then err_msg = err_msg & vbNewLine & "* Enter the information about the change reported."
                If trim(job_change_new_hourly_wage) = "" Then
                    err_msg = err_msg & vbNewLine & "* The new hourly wage must be entered."
                Else
                    If IsNumeric(job_change_new_hourly_wage) = False Then err_msg = err_msg & vbNewLine & "* Enter the hourly wage as a number."
                End If
                If trim(job_change_new_hours_per_week) = "" Then
                    err_msg = err_msg & vbNewLine & "* The new hours worked per week must be entered."
                Else
                    If IsNumeric(job_change_new_hours_per_week) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the hours per week as a number."
                End If

            Case "Job Ended"
                If IsDate(job_end_income_end_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the date of the last paycheck."
                If IsDate(date_work_ended) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the date the client last worked."
                If IsNumeric(last_pay_amount) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the amount of the last pay. This can be an estimate amount."
                If vol_quit_yn = "?" Then err_msg = err_msg & vbNewLine & "* Select Voluntary Quit information - Yes or No"
                If stwk_verif = "Select One..." Then err_msg = err_msg & vbNewLine & "* Enter the verification code for the STWK Panel."
                If refused_empl_yn = "Yes" AND IsDate(refused_empl_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Since you indicated that the client refused employment, you must enter the date employment was refused as a valid date."
                If vol_quit_yn = "Yes" AND good_cause_yn = "No" Then err_msg = err_msg & vbNewLine & "* Currently there voluntary quit sanctions and processing is suspended due to COVID-19. If the client quit voluntarily, code good cause as 'Yes'."            'This is here because good cause is always true - COVID WAIVER
                If vol_quit_yn = "Yes" AND good_cause_yn = "?" Then err_msg = err_msg & vbNewLine & "* Since this job has ended voluntarily, indicate if this voluntary quit meets good cause."
        End Select
        err_msg = err_msg & vbNewLine
        If trim(wreg_abawd_notes) = "" AND SNAP_case = TRUE Then err_msg = err_msg & vbNewLine & vbNewLine & "* Enter how this change will impact WREG/ABWAD for this pmember. NOTE: The imact may be 'NO CHANGE' or something similar. This information provided here should be thorough and complete, more is better when explaining the WREG and ABAWD status." & vbNewLine
        If conversation_with_person = "Type or Select" Then     'handling to deal with conversation information - thiere should be a person listed AND conversation detil if either are listed
            If trim(conversation_detail) <> "" Then err_msg = err_msg & vbNewLine & "* There is information added to the conversation detail but no information about who the conversation was with has been provided. Enter the member or name of the person the conversationwas completed with."
        Else
            If trim(conversation_detail) = "" Then err_msg = err_msg & vbNewLine & "* There is information provided about who the conversation has been completed with but no details about what was discussed in the conversation. Add details about the conversation."
        End If

        If trim(verif_form_date) <> "" Then                     'Getting detail if a verif date is listed and ensuring the verif date is a date.
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

        If err_msg = vbNewLine & vbNewLine & vbNewLine Then err_msg = ""
        'Displaying the error message
        If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg

    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE
job_income_start = DateAdd("d", 0, job_income_start)

Select Case job_change_type                                         'here we are finding the MAXIS Footer Month and Year based on when the change was made
    Case "New Job Reported"
        CALL convert_date_into_MAXIS_footer_month(new_job_income_start, MAXIS_footer_month, MAXIS_footer_year)
    Case "Income/Hours Change for Current Job"
        CALL convert_date_into_MAXIS_footer_month(first_pay_date_of_change, MAXIS_footer_month, MAXIS_footer_year)
    Case "Job Ended"
        CALL convert_date_into_MAXIS_footer_month(job_end_income_end_date, MAXIS_footer_month, MAXIS_footer_year)
End Select

If conversation_with_person = "Type or Select" Then conversation_with_person = ""

'VOLUNTARY QUIT functionality needs TESTING when this policy goes back into effect
review_vol_quit = FALSE                     'this is a a defaulted variable to indicate if we need to review more detail about possible voluntary quit
If job_change_type = "Income/Hours Change for Current Job" Then                 'Job change does not have a vol quit option but can still meet vol quit/reduction
    If IsNumeric(job_change_old_hours_per_week) = TRUE Then                     'we will default to asking more if the hours decrease sufficiently
        job_change_old_hours_per_week = job_change_old_hours_per_week * 1
        job_change_new_hours_per_week = job_change_new_hours_per_week * 1

        If job_change_old_hours_per_week > 30 and job_change_new_hourly_wage <30 Then
            review_vol_quit = TRUE
            Vol_quit_type = "Hours Reduction"
        End If
    End If
End If
If vol_quit_yn = "Yes" Then         'job end has vol quit field that if indicated yes will force additional information
    review_vol_quit = TRUE
    Vol_quit_type = "Quit Job"
End If
review_vol_quit = FALSE         'This is going to false now because there is no voluntary quit right now - COVID WAIVER
code_disq = FALSE               'defaulting this variable before the additional vol quit detail
'This functionality provides additional information gathering about Volundary Quit and the ability to take action/create seperate notes
If review_vol_quit = TRUE Then
    If mfip_case = TRUE Then mfip_vol_quit_checkbox = checked           'Autochecking the program boxes based upon which programs were found to be active
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

    If vol_quit_yn = "Yes" AND good_cause_yn = "No" Then code_disq = TRUE       'This is the variable that will cause the DISQ updating later
End If

Call back_to_SELF               'making sure we are in the right place
Call MAXIS_background_check
' MsgBox "Start month - " & MAXIS_footer_month & vbNewLine & "Start year - " & MAXIS_footer_year
employee_name_only = right(job_employee, len(job_employee) - 5)

Initial_footer_month = MAXIS_footer_month           'We are going to loop through months, so we need to set the initial footer month so we remember it later
Initial_footer_year = MAXIS_footer_year
second_loop = FALSE         'Knowing where we are

If developer_mode = FALSE Then                      'If we are in developer mode then we are going to skip the update parts
    Do
        EMWriteScreen "SUMM", 20, 71                'Getting into STAT
        transmit

        EMReadScreen MAXIS_footer_month, 2, 20, 55          'Setting the right month and year
        EMReadScreen MAXIS_footer_year, 2, 20, 58

        If second_loop = TRUE Then                          'If we are past the first loop, we know the job member and instance so we can naviage directly there
            EMWriteScreen "JOBS", 20, 71                    'go to JOBS
            EMWriteScreen ref_nbr, 20, 76                   'go to the right member
            EMWriteScreen job_instance, 20, 79
            transmit

            EMReadScreen panel_exists, 14, 24, 13
            If panel_exists = "DOES NOT EXIST" Then
                EMWriteScreen "JOBS", 20, 71                                        'go to JOBS
                ref_nbr = left(job_employee, 2)
                EMWriteScreen ref_nbr, 20, 76                                       'go to the right member
                EMWriteScreen "NN", 20, 79                                          'create new JOBS panel
                transmit

                EMReadScreen this_instance, 1, 2, 73                                'Reading the instance because we need it on the next loop
                job_instance = "0" & this_instance
            Else
                EMReadScreen this_panel_job, 30, 7, 42          'reading the name of job

                If this_panel_job <> original_full_jobs_name Then       'if the panel name isn't what we read earlier then we have a problem
                MsgBox "They don't match!" & vbNewLine & "THIS PANEL-" & this_panel_job & "-" & vbNewLine & "ORIGINAL-" & original_full_jobs_name & "-"
                Else
                PF9         'If they match - put it in edit mode
                End If
            End If


        ElseIf script_update_stat = "Yes - Update an existing JOBS Panel" Then          'if we are in the first loop then the action is going to change based on the the update detail
            ' MsgBox "Update - " & script_update_stat & " - 1"
            If job_instance = "" Then                                           'If we don't know the instance yet - the script will facilitate finding the existing panel
                ref_nbr = left(job_employee, 2)                                 'Finding the reference number and going to the right member
                EMWriteScreen ref_nbr, 20, 76
                transmit

                EMReadScreen total_jobs, 1, 2, 78                               'Reading the number of jobs that are on this case for this member
                If total_jobs = "0" Then                                        'If there are none and the worker indicated to update an existing panel, thes cript will end.
                    Call script_end_procedure("Update and Note NOT Completed. There are no jobs for Memb " & job_employee & " listed in MAXIS and you have selected to have the script Update and existing JOBS panel.")
                Else    'If they have found panels for this member
                    job_selected = FALSE                                        'setting variable for job as unkown
                    Do      'This is to loop through all the JOBS panels
                        EMReadScreen employer, 30, 7, 42                        'Reading the employer name
                        employer = replace(employer, "_", "")                   'Trimming/formatting the name
                        employer_check = MsgBox("Is this the job reported? Employer name: " & employer, vbYesNo + vbQuestion, "Select Income Panel")        'Now we ask the worker if this is the employer
                        If employer_check = vbYes Then                          'If the worker says 'Yes' to the message box then we found the job and we will set the information
                            job_selected = TRUE                                 'letting the script know we have the job'\
                            Exit Do                                             'stop looking at more jobs
                        End If
                        transmit                                                'If we didn't find the job then it is going to transmit to the next job for this member
                        EMReadScreen last_job, 7, 24, 2                         'checking to see if we found the last job for the member
                    Loop until last_job = "ENTER A"
                    'If we leave the loop and we haven't found the job the script will end because we didn't find a job
                    If job_selected = FALSE Then Call script_end_procedure("Update and Note NOT completed. You did not select any of the JOBS for Memb " & job_employee & " but indicated the script should update a JOBS panel.")
                End If
                EMReadScreen this_instance, 1, 2, 73            'reading the instance of the job we found because we need it on the next loop
                job_instance = "0" & this_instance
            Else
                EMWriteScreen "JOBS", 20, 71                    'If we know the instance we can just navigate to the JOBS panel
                ref_nbr = left(job_employee, 2)
                EMWriteScreen ref_nbr, 20, 76                   'go to the right member
                EMWriteScreen job_instance, 20, 79

                transmit
            End If
            PF9                                                 'Put the JOBS in edit mode

        ElseIf script_update_stat = "Yes - Create a new JOBS Panel" Then        'If this is a new panel - the script will now create it
            ' MsgBox "Update - " & script_update_stat & " - 2"
            EMWriteScreen "JOBS", 20, 71                                        'go to JOBS
            ref_nbr = left(job_employee, 2)
            EMWriteScreen ref_nbr, 20, 76                                       'go to the right member
            EMWriteScreen "NN", 20, 79                                          'create new JOBS panel
            transmit

            EMReadScreen this_instance, 1, 2, 73                                'Reading the instance because we need it on the next loop
            job_instance = "0" & this_instance
        End If          'Now we are done with the code for if we are in the first loop of updating
        new_hours_per_check = 0
        old_hours_per_check = 0
        'If the script indicates that we need to update at all here is the part where we actually put the information in the panel
        'The panel should already be in EDIT MODE
        If script_update_stat = "Yes - Update an existing JOBS Panel" OR script_update_stat = "Yes - Create a new JOBS Panel" Then
            ' MsgBox "Update - " & script_update_stat & " - 3"
            EMWriteScreen left(job_income_type, 1), 5, 34                                                               'income type
            If job_subsidized_income_type <> "" Then EMWriteScreen left(job_subsidized_income_type, 2), 5, 74           'subsidized type
            If job_verification <> " " Then EMWriteScreen left(job_verification, 1), 6, 34                              'job verification
            EMWriteScreen "                              ", 7, 42                                                       'blank out the employer name
            EMWriteScreen job_employer_name, 7, 42                                                                      'enter the employer name
            If IsDate(job_income_start) = TRUE Then Call create_mainframe_friendly_date(job_income_start, 9, 35, "YY")      'income start date
            If IsDate(job_income_end) = TRUE Then Call create_mainframe_friendly_date(job_income_end, 9, 49, "YY")          'income end date
            If IsDate(contract_through_date) = TRUE then call create_mainframe_friendly_date(contract_through_date, 9, 73, "YY")
            EMWriteScreen left(job_pay_frequency, 1), 18, 35                                                            'pay frequency
            ' If original_full_jobs_name = "" Then EMReadScreen original_full_jobs_name, 30, 7, 42             'read the employer name as it originally exists on the panel

            If job_change_type = "Income/Hours Change for Current Job" Then         'If we are changing a job, we need to find the check amount that is already there s we can put it back in
                EMReadScreen job_amount_one, 8, 12, 67
                EMReadScreen job_amount_two, 8, 13, 67
                EMReadScreen job_amount_three, 8, 14, 67
                EMReadScreen job_amount_four, 8, 15, 67
                EMReadScreen job_amount_five, 8, 16, 67
                EMReadScreen job_hours, 3, 18, 72
                If job_hours = "___" Then job_hours = 0

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
                job_hours = trim(job_hours)
                job_hours = job_hours * 1
                total_hours = job_hours
                If numb_of_checks <> 0 Then old_hours_per_check = job_hours / numb_of_checks
            End If
            If job_change_type = "Job Ended" Then           'If job ended we need to know what pay amounts were known.
                EMReadScreen known_pay_amount, 8, 12, 67
                known_pay_amount = trim(known_pay_amount)
            End If
            'Now that we have read the information we need from the panel as it already exists, we will blank out all the dates and check amounts.
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

            Select Case job_change_type                 'What we enter and how we enter is baed upon what change type is reported.
                'This doesn't actually do the updates but sets up all the variables so we can use the same update code
                Case "New Job Reported"
                    If trim(new_job_hourly_wage) <> "" Then
                        EMWriteScreen "      ", 6, 75
                        EMWriteScreen new_job_hourly_wage, 6, 75            'setting the hourly wage
                    End If

                    the_last_pay_date = new_job_income_start                    'these are used for the loops through paychecks - we start when the income actually starts
                    the_first_pay_date = new_job_income_start
                    end_of_pay = "99/99/99"                                     'If we have a new job, we can't code the end of income
                    pay_amt = "0"                                               'We do not enter an amount for new jobs since it isn't verified
                Case "Income/Hours Change for Current Job"
                    If trim(job_change_new_hourly_wage) <> "" Then
                        EMWriteScreen "      ", 6, 75
                        EMWriteScreen job_change_new_hourly_wage, 6, 75       'setting the hourly wage
                    End If

                    date_jump_back = first_pay_date_of_change
                    change_month = DatePart("m", first_pay_date_of_change)
                    change_year = DatePart("yyyy", first_pay_date_of_change)
                    the_last_pay_date = first_pay_date_of_change                'used for loops - we start with when the income actually changes

                    end_of_pay = "99/99/99"                                     'For job change we do not have an end of income
                    job_change_new_hourly_wage = job_change_new_hourly_wage * 1         'making hourly wage and hours per week actual numbers instead of strings
                    job_change_new_hours_per_week = job_change_new_hours_per_week * 1

                    Do
                        If prev_date <> "" Then date_jump_back = prev_date
                        If job_pay_frequency = "1 - Monthly" Then                   'Setting multiplier based on the pay frequency and wage to determine the pay amount to enter
                            pay_amt = job_change_new_hourly_wage * job_change_new_hours_per_week * 4.3
                            prev_date = DateAdd("m", -1, date_jump_back)
                            If DatePart("m", prev_date) = change_month AND DatePart("yyyy", prev_date) = change_year Then
                                If DateDiff("d", prev_date, job_income_start) =< 0 Then the_last_pay_date = prev_date
                            End If
                        ElseIf job_pay_frequency = "2 - Semi-Monthly" Then
                            pay_amt = job_change_new_hourly_wage * job_change_new_hours_per_week * 2.15
                            prev_date = DateAdd("d", -14, date_jump_back)
                            If DatePart("m", prev_date) = change_month AND DatePart("yyyy", prev_date) = change_year Then
                                If DateDiff("d", prev_date, job_income_start) =< 0 Then the_last_pay_date = prev_date
                            End If
                        ElseIf job_pay_frequency = "3 - Biweekly" Then
                            pay_amt = job_change_new_hourly_wage * job_change_new_hours_per_week * 2
                            prev_date = DateAdd("d", -14, date_jump_back)
                            If DatePart("m", prev_date) = change_month AND DatePart("yyyy", prev_date) = change_year Then
                                If DateDiff("d", prev_date, job_income_start) =< 0 Then the_last_pay_date = prev_date
                            End If
                        ElseIf job_pay_frequency = "4 - Weekly" Then
                            pay_amt = job_change_new_hourly_wage * job_change_new_hours_per_week
                            prev_date = DateAdd("d", -7, date_jump_back)
                            If DatePart("m", prev_date) = change_month AND DatePart("yyyy", prev_date) = change_year Then
                                ' MsgBox "Prev date: " & prev_date & vbNewLine  & "Diff - " & DateDiff("d", prev_date, job_income_start)
                                If DateDiff("d", prev_date, job_income_start) =< 0 Then the_last_pay_date = prev_date
                            End If
                        End If
                    Loop until DatePart("m", prev_date) <> change_month

                    the_first_pay_date = the_last_pay_date                      'used for loops - we start with when the income actually changes
                    prosp_hours = 0                                             'setting the pospecive hours to be a number'
                Case "Job Ended"
                    end_of_pay = "99/99/99"                                     'defaulting the end of pay to being blank
                    If IsDate(job_end_income_end_date) = TRUE Then end_of_pay = job_end_income_end_date         'If the end of of income is listed as a date, it will be saved here
                    the_last_pay_date = job_end_income_end_date                 'Sving these for the loops, we start with the last day of pay
                    the_first_pay_date = job_end_income_end_date
                    prosp_hours = 0                                             'setting the prospective hourse to be a number
                    pay_amt = known_pay_amount                                  'setting the pay amount to the amount we found earlier when reading the panel
            End Select

            If end_of_pay = "99/99/99" Then     'Based on the end date, we will set the row to start with
                jobs_row = 12       'If we have no end date we start at the top of the list of pay dates
            Else
                jobs_row = 16       'If we do have an end date, we start at the bottom
                Call create_mainframe_friendly_date(job_end_income_end_date, 9, 49, "YY")       'If we know the end date, this will enter the income end date in the forect spot on the panel.
            End If
            prosp_hours = 0                 'setting the pospecive hours to be a number - this has to reset at the beginning of every month

            'If we are in the first month to update
            If Initial_footer_month = MAXIS_footer_month AND Initial_footer_year = MAXIS_footer_year Then
                Call create_mainframe_friendly_date(the_first_pay_date, jobs_row, 54, "YY")     'Entering the first pay date
                If end_of_pay = "99/99/99" Then             'If the job is not ended
                    If job_change_type = "Income/Hours Change for Current Job" Then                     'Entering the pay amount -we write the pay amount determined by job change type
                        If DateDiff("d", the_first_pay_date, first_pay_date_of_change) > 0 Then
                            EMWriteScreen check_amount, jobs_row, 67
                            prosp_hours = prosp_hours + old_hours_per_check
                        Else
                            EMWriteScreen pay_amt, jobs_row, 67
                            prosp_hours = prosp_hours + new_hours_per_check
                        End If
                    Else
                        EMWriteScreen pay_amt, jobs_row, 67
                        prosp_hours = prosp_hours + new_hours_per_check
                    End If
                    jobs_row = jobs_row + 1                     'going to the next row down
                Else                                        'If the job IS ended
                    EMWriteScreen last_pay_amount, jobs_row, 67 'Enter the last gross pay
                    prosp_hours = last_pay_amount/job_hourly_wage
                    jobs_row = jobs_row - 1                     'go to the next row above
                End If
            End If
            next_month_mo = ""      'blanking out the next footer month
            next_month_yr = ""
            the_month_here = DateValue(MAXIS_footer_month & "/01/" & MAXIS_footer_year)     'making an actual date using the footer month
            the_next_month = DateAdd("m", 1, the_month_here)                                'getting a date in the next month
            Call convert_date_into_MAXIS_footer_month(the_next_month, next_month_mo, next_month_yr)     'getting a footer month and year from the next month found

            ' MsgBox "The next month - " & next_month_mo & "/" & next_month_yr & vbCr & vbCr & "MAXIS month - " & MAXIS_footer_month & "/" & MAXIS_footer_year
            'Here we are going to loop through the pay dates to enter any additional ones for the current month
            Do
                the_pay_date = ""               'clearing the pay date variables
                the_pay_date_two = ""
                If end_of_pay = "99/99/99" Then                                 'If the job has not ended the paycheck dates go forward to find the next check from the starting check
                    If job_pay_frequency = "1 - Monthly" Then
                        the_pay_date = DateAdd("m", 1, the_last_pay_date)
                        new_hours_per_check = job_change_new_hours_per_week * 4.3
                    ElseIf job_pay_frequency = "2 - Semi-Monthly" Then
                        the_pay_date = DateAdd("d", 15, the_last_pay_date)
                        the_pay_date_two = DateAdd("m", 1, the_last_pay_date)
                        new_hours_per_check = job_change_new_hours_per_week * 2.15
                    ElseIf job_pay_frequency = "3 - Biweekly" Then
                        the_pay_date = DateAdd("d", 14, the_last_pay_date)
                        new_hours_per_check = job_change_new_hours_per_week * 2
                    ElseIf job_pay_frequency = "4 - Weekly" Then
                        the_pay_date = DateAdd("d", 7, the_last_pay_date)
                        new_hours_per_check = job_change_new_hours_per_week
                    End If
                Else                                                            'If the job HAS ended the paycheck dates go backward to find the next check from the starting check
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
                ' MsgBox "Old hours per check: " & old_hours_per_check & vbNewLine & "New hours per check: " & new_hours_per_check

                Call convert_date_into_MAXIS_footer_month(the_pay_date, pay_date_mo, pay_date_yr)       'Getting the footer month and year from the pay date
                If IsDate(the_pay_date_two) = TRUE Then Call convert_date_into_MAXIS_footer_month(the_pay_date_two, pay_date_two_mo, pay_date_two_yr)
                ' MsgBox "The pay date - " & the_pay_date & vbCr & vbCr & "Pay month - " & pay_date_mo & "/" & pay_date_yr & vbCr & "MAXIS month - " & MAXIS_footer_month & "/" & MAXIS_footer_year & vbCr & vbCr & "JOBS Row - " & jobs_row
                If pay_date_mo = MAXIS_footer_month AND pay_date_yr = MAXIS_footer_year Then            'If the pay date is in the current month, we are going to writ it into JOBS
                    Call create_mainframe_friendly_date(the_pay_date, jobs_row, 54, "YY")               'Entering the date
                    If job_change_type = "Income/Hours Change for Current Job" Then                     'Entering the pay amount -
                        If DateDiff("d", the_pay_date, first_pay_date_of_change) > 0 Then
                            EMWriteScreen check_amount, jobs_row, 67
                            prosp_hours = prosp_hours + old_hours_per_check
                        Else
                            EMWriteScreen pay_amt, jobs_row, 67
                            prosp_hours = prosp_hours + new_hours_per_check
                        End If
                    Else
                        EMWriteScreen pay_amt, jobs_row, 67
                        prosp_hours = prosp_hours + new_hours_per_check
                    End If
                    If end_of_pay = "99/99/99" Then                 'going to the next row
                        jobs_row = jobs_row + 1                         'go down one for job is not ended
                    Else
                        jobs_row = jobs_row - 1                         'go up one for job IS ended
                        pay_amt = pay_amt * 1                           'This determines how many hours to put on the panel if job has ended - We need to calculate for each pay amount
                        job_hourly_wage = job_hourly_wage * 1
                        hours_of_pay = pay_amt/job_hourly_wage
                        prosp_hours = prosp_hours + hours_of_pay        'totaling the hours calculated here
                    End If
                End If
                If pay_date_two_mo = MAXIS_footer_month AND pay_date_two_yr = MAXIS_footer_year Then        'Same functionality for a second pay date - this is used for semi monthly pay frequesny
                    Call create_mainframe_friendly_date(the_pay_date_two, jobs_row, 54, "YY")
                    If job_change_type = "Income/Hours Change for Current Job" Then                     'Entering the pay amount -
                        If DateDiff("d", the_pay_date, first_pay_date_of_change) > 0 Then
                            EMWriteScreen check_amount, jobs_row, 67
                        Else
                            EMWriteScreen pay_amt, jobs_row, 67
                        End If
                    Else
                        EMWriteScreen pay_amt, jobs_row, 67
                    End If
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
                the_last_pay_date = the_pay_date        'setting the variable that stores the previous pay date for the next loop
                If the_pay_date_two <> "" Then the_last_pay_date = the_pay_date_two

                If job_change_type = "Job Ended" Then       'Since end job dates move backwards, we would get stuck in a loop if we don't exit out once we move from teh maxis month
                    If pay_date_mo <> MAXIS_footer_month Then Exit Do
                End If
            Loop until pay_date_mo = next_month_mo AND pay_date_yr = next_month_yr OR pay_date_two_mo = next_month_mo AND pay_date_two_yr = next_month_yr       'This is getting out of the update loop using footer month

            If end_of_pay = "99/99/99" Then             'counting how many checks have been enterter for this month
                checks_entered = jobs_row - 12
            Else
                checks_entered = abs(jobs_row - 16)
            End If
			If IsNumeric(prosp_hours) = TRUE Then prosp_hours = round(prosp_hours)
            If job_change_type = "New Job Reported" Then prosp_hours = "000"        'If new job then set the hours to 0
            If job_change_type = "Job Ended" AND prosp_hours = 0 Then prosp_hours = ""
            EMWriteScreen "   ", 18, 72             'blanking out hours and entering the new ones
            EMWriteScreen prosp_hours, 18, 72

            'Now we update the PIC
            If snap_case = TRUE Then
                EMWriteScreen "X", 19, 38           'Open the SNAP PIC'
                transmit

                Call create_mainframe_friendly_date(date, 5, 34, "YY")      'Entering the date of calculateion - TODAY
                EMWriteScreen left(job_pay_frequency, 1), 5, 64            'Entering the job frequency code

                Select Case job_change_type                                 'We only update the PIC for new job and job change - job end takes care of itself.
                    Case "New Job Reported"
                        EMWriteScreen "      ", 8, 64
                        EMWriteScreen "        ", 9, 66
                        If trim(new_job_hourly_wage) <> "" Then
                            EMWriteScreen "0", 8, 64                        '0 for hours per week'
                            EMWriteScreen new_job_hourly_wage, 9, 66        'entering the hourly wage
                        End If

                    Case "Income/Hours Change for Current Job"
                        EMWriteScreen "      ", 8, 64
                        EMWriteScreen "        ", 9, 66
                        EMWriteScreen job_change_new_hours_per_week, 8, 64      'entering the OLD hours per week on the PIC
                        EMWriteScreen job_change_new_hourly_wage, 9, 66         'entering the OLD hourly wage on the PIC
                End Select
                transmit
                PF3
            End If
        End If
        transmit        'saving the panel information
        ' MsgBox "JOBS saved"

        If original_full_jobs_name = "" Then EMReadScreen original_full_jobs_name, 30, 7, 42             'read the employer name as it originally exists on the panel

        If script_update_stat = "Yes - Update an existing JOBS Panel" OR script_update_stat = "Yes - Create a new JOBS Panel" Then      'If we are updating
            If job_change_type = "Job Ended" Then                                                                                       'and we are coding a job end change type
                Call navigate_to_MAXIS_screen("STAT", "STWK")                   'We also have to code end of employmnt on STWK
                EMWriteScreen ref_nbr, 20, 76                                   'navigate to STWK for the right person
                transmit

                EMReadScreen version_of_stwk, 1, 2, 73                          'Seeing if there is already a STWK panel in existence
                If version_of_stwk = "1" Then                                   'If one exists, put in edit mode
                    PF9
                End If
                If version_of_stwk = "0" Then                                   'If one does not exist, create a new panel
                    EMWriteScreen "NN", 20, 79
                    transmit
                End If

                EMWriteScreen "                              ", 6, 46           'blanking out employer name and entering a new one
                EMWriteScreen job_employer_name, 6, 46

                Call create_mainframe_friendly_date(date_work_ended, 7, 46, "YY")               'Entering work end date
                Call create_mainframe_friendly_date(job_end_income_end_date, 8, 46, "YY")       'Entering income end date

                EMWriteScreen stwk_verif, 7, 63                 'Entering the verification of STWK
                If refused_empl_yn = "?" Then refused_empl_yn = ""
                If good_cause_yn = "?" Then good_cause_yn = ""
                EMWriteScreen refused_empl_yn, 8, 78            'Coding if refused employment
                If IsDate(refused_empl_date) = TRUE Then Call create_mainframe_friendly_date(refused_empl_date, 10, 72, "YY")       'Entering the date refused employment
                EMWriteScreen vol_quit_yn, 10, 46               'Entering the voluntary Quit yes/no
                If adult_cash_case = TRUE or family_cash_case = TRUE Then EMWriteScreen good_cause_yn, 12, 52       'Coding good cause by programs
                If grh_case = TRUE Then EMWriteScreen good_cause_yn, 12, 60
                If snap_case = TRUE Then EMWriteScreen good_cause_yn, 12, 67

                If pwe_ref = ref_nbr Then           'Entering if the member is PWE based on WREG
                    EMWriteScreen "Y", 14, 46
                Else
                    EMWriteScreen "N", 14, 46
                End If
                transmit                'Saving the panel
                EMReadScreen error_prone_warning, 20, 6, 43
                If error_prone_warning = "Error Prone Warnings" Then transmit

                If refused_empl_yn = "" Then refused_empl_yn = "?"
                If good_cause_yn = "" Then good_cause_yn = "?"
            End If

            If code_disq = TRUE Then        'This will code a DISQ panel in the case of Voluntary Quit
                Call navigate_to_MAXIS_screen("STAT", "DISQ")
                EMWriteScreen ref_nbr, 20, 76
                transmit

                'TODO - add functionality to update DISQ once vol quit is a thing again.'
            End If
        End If

        transmit                            'Now we send the case through background and go to STAT/WRAP
        EmWriteScreen "BGTX", 20, 71
        transmit

        EMReadScreen wrap_check, 4, 2, 46
        If wrap_check <> "WRAP" Then

        End If
        EMWriteScreen "Y", 16, 54           'This goes into the next footer month without leaving STAT
        If MAXIS_footer_month = CM_plus_1_mo AND MAXIS_footer_year = CM_plus_1_yr Then EMWriteScreen "N", 16, 54        'If we are already in CM+1 then we enter N because we can't update the next month
        transmit
        ' MsgBox "Pause here - should be in the next month."

        second_loop = TRUE                  'setting this veriable to knkow we aren't in the first go round any more
    Loop until MAXIS_footer_month = CM_plus_1_mo AND MAXIS_footer_year = CM_plus_1_yr
Else            'If we are in developer mode, we will go here to allow for some display and output options.
    Do
        Do
            err_msg = ""

            'This dialog allows workers to send an email with detail gathered to up to five recipients.
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

            email_address_one = trim(email_address_one)
            email_address_two = trim(email_address_two)
            email_address_three = trim(email_address_three)
            email_address_four = trim(email_address_four)
            email_address_five = trim(email_address_five)

            If send_email_checkbox = checked Then
                If email_address_one = "" AND email_address_two = "" AND email_address_three = "" AND email_address_four = "" AND email_address_five = "" Then err_msg = err_msg & vbNewLine & "* You have selected to have the script email the Job Change information but haven't provided email addresses to send the information to."
            Else
                If email_address_one <> "" OR email_address_two <> "" OR email_address_three <> "" OR email_address_four <> "" OR email_address_five = "" Then err_msg = err_msg & vbNewLine & "* You have entered at least one email address but have not indicated you want the script to send an email."
            End If
            If InStr(email_address_one, " ") <> 0 Then err_msg = err_msg & vbNewLine & "* The FIRST Email Address has a space in it. This is not a valid email address. Please review the first email address entered."
            If InStr(email_address_two, " ") <> 0 Then err_msg = err_msg & vbNewLine & "* The SECOND Email Address has a space in it. This is not a valid email address. Please review the second email address entered."
            If InStr(email_address_three, " ") <> 0 Then err_msg = err_msg & vbNewLine & "* The THIRD Email Address has a space in it. This is not a valid email address. Please review the third email address entered."
            If InStr(email_address_four, " ") <> 0 Then err_msg = err_msg & vbNewLine & "* The FOURTH Email Address has a space in it. This is not a valid email address. Please review the fourth email address entered."
            If InStr(email_address_five, " ") <> 0 Then err_msg = err_msg & vbNewLine & "* The FIFTH Email Address has a space in it. This is not a valid email address. Please review the fifth email address entered."
            If InStr(email_address_one, "@") <> 0 Then err_msg = err_msg & vbNewLine & "* The FIRST Email Address appears to have the email domain (the '@somewhere.com' part) listed in it. The scriot will fill in '@hennepin.us' for each email address. Other domains are not allowed at this time as we do not have functionality built in to safely email outside of the county."
            If InStr(email_address_two, "@") <> 0 Then err_msg = err_msg & vbNewLine & "* The SECOND Email Address appears to have the email domain (the '@somewhere.com' part) listed in it. The scriot will fill in '@hennepin.us' for each email address. Other domains are not allowed at this time as we do not have functionality built in to safely email outside of the county."
            If InStr(email_address_three, "@") <> 0 Then err_msg = err_msg & vbNewLine & "* The THIRD Email Address appears to have the email domain (the '@somewhere.com' part) listed in it. The scriot will fill in '@hennepin.us' for each email address. Other domains are not allowed at this time as we do not have functionality built in to safely email outside of the county."
            If InStr(email_address_four, "@") <> 0 Then err_msg = err_msg & vbNewLine & "* The FOURTH Email Address appears to have the email domain (the '@somewhere.com' part) listed in it. The scriot will fill in '@hennepin.us' for each email address. Other domains are not allowed at this time as we do not have functionality built in to safely email outside of the county."
            If InStr(email_address_five, "@") <> 0 Then err_msg = err_msg & vbNewLine & "* The FIFTH Email Address appears to have the email domain (the '@somewhere.com' part) listed in it. The scriot will fill in '@hennepin.us' for each email address. Other domains are not allowed at this time as we do not have functionality built in to safely email outside of the county."


            email_address_two = trim(email_address_two)
            email_address_three = trim(email_address_three)
            email_address_four = trim(email_address_four)
            email_address_five = trim(email_address_five)

            If err_msg <> "" Then MsgBox "Please resolve to continue: " & vbNewLine & err_msg

        Loop until err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE

    If refused_empl_yn = "?" Then refused_empl_yn = "N/A"
    If good_cause_yn = "?" Then good_cause_yn = "N/A"
    'The email checkbox will start this option to send emails
    If send_email_checkbox = checked Then
        email_address_one = trim(email_address_one)         'formatting the email address information that was entered
        email_address_two = trim(email_address_two)
        email_address_three = trim(email_address_three)
        email_address_four = trim(email_address_four)
        email_address_five = trim(email_address_five)

        'Now the email addresses will add the email suffix to each email address that is not blank
        'Then if not blank, it will be added to a string of all of the email addresses
        If email_address_one <> "" Then
            email_address_one = email_address_one & "@hennepin.us"
            all_email_recipients = email_address_one
        End If
        If email_address_two <> "" Then
            email_address_two = email_address_two & "@hennepin.us"
            all_email_recipients = all_email_recipients & "; " & email_address_two
        End If
        If email_address_three <> "" Then
            email_address_three = email_address_three & "@hennepin.us"
            all_email_recipients = all_email_recipients & "; " & email_address_three
        End If
        If email_address_four <> "" Then
            email_address_four = email_address_four & "@hennepin.us"
            all_email_recipients = all_email_recipients & "; " & email_address_four
        End If
        If email_address_five <> "" Then
            email_address_five = email_address_five & "@hennepin.us"
            all_email_recipients = all_email_recipients & "; " & email_address_five
        End If

        If right(all_email_recipients, 2) = "; " Then all_email_recipients = left(all_email_recipients, len(all_email_recipients) - 2)      'removing the extra seperators so the email doesn't error out

        'Setting up the information to send in the email.
        'This closely follows the CASE:NOTE at the end of the scrip
        email_body = "Job change information reported. Update to MAXIS not made." & vbCR
        email_body = email_body & "Case Number: " & MAXIS_case_number & " - Job chaange type: " & job_change_type & vbCr
        email_body = email_body & "Change reported on " & date
        email_body = email_body & "Script was run when MAXIS was in INQUIRY and information could not be added to MAXIS." & vbCr & vbCr

        email_body = email_body & "=== Details of the Reported Change ===" & vbCr
        email_body = email_body & "Job Name: " & job_employer_name & " - Income Type: " & job_income_type & " - Employee: Memb " & job_employee & vbCr
        If job_income_end <> "" Then email_body = email_body & "Income Start Date: " & job_income_start & " - End Date: " & job_income_end & vbCr
        If job_income_end = "" Then email_body = email_body & "Income Start Date: " & job_income_start & vbCr
        email_body = email_body & vbCr

        Select Case job_change_type
            Case "New Job Reported"
                email_body = email_body & "*** New Job Reported ***" & vbCr
                email_body = email_body & "Work Start Date: " & date_work_started & vbCr
                email_body = email_body & "Income Start Date: " & new_job_income_start & " - Initial Gross Pay: $" & initial_check_gross_amount & vbCr
                email_body = email_body & "Anticipated Income: Hours per Week: " & new_job_hours_per_week & " - Hourly Wage: " & new_job_hourly_wage & vbCr
            Case "Income/Hours Change for Current Job"
                email_body = email_body & "*** Income/Hours Change for Current Job ***" & vbCr
                email_body = email_body & "Change happened on " & job_change_date & " change will cause " & income_change_type & vbCr
                email_body = email_body & "Date of pay first impacted: " & first_pay_date_of_change & vbCr
                email_body = email_body & "Change: " & job_change_details & vbCr
                email_body = email_body & "Previous Income: Hours per Week: " & job_change_old_hours_per_week & " - Hourly Wage: " & job_change_old_hourly_wage & vbCr
                email_body = email_body & "New Income: Hours per Week: " & job_change_new_hours_per_week & " - Hourly Wage: " & job_change_new_hourly_wage & vbCr
            Case "Job Ended"
                email_body = email_body & "*** Job Ended ***" & vbCr
                email_body = email_body & "Income End Date: " & job_end_income_end_date & " - Final pay amount: "& last_pay_amount & vbCr
                email_body = email_body & "Work stoped on: " & date_work_ended & vbCr
                email_body = email_body & "* Quit details:" & vbCr
                email_body = email_body & " - Employee refused employment: " & refused_empl_yn & vbCr
                email_body = email_body & " - Was this a voluntary quit? " & vol_quit_yn & vbCr
                If trim(stwk_reason) <> "" Then email_body = email_body & " - Reason for STWK: " & stwk_reason & vbCr
                email_body = email_body & " - Meets good cause? " & good_cause_yn & vbCr
        End Select
        email_body = email_body & vbCr
        email_body = email_body & "Impact on WREG/ABAWD: " & wreg_abawd_notes & vbCr
        If conversation_with_person <> "" Then
            email_body = email_body & "Information about job gathered in conversation with " & conversation_with_person & vbCr
            email_body = email_body & "  - Details of conversation: " & conversation_detail & vbCr
        End If
        email_body = email_body & "=== Reporting Information ===" & vbCr
        email_body = email_body & "Reported via " & job_report_type & " by " & person_who_reported_job & " on " & reported_date & vbCr
        If work_number_verbal_checkbox = checked Then email_body = email_body & "*** Verbal authorization to check the Work Number received." & vbCr
        email_body = email_body & vbCr
        email_body = email_body & "=== Verification ===" & vbCr
        If work_number_checkbox = checked Then email_body = email_body & "Sent Work Number request for income verification." & vbCr
        If verif_form_date <> "" Then
            email_body = email_body & "* Verification request sent on " & verif_form_date & vbCr
            email_body = email_body & "* Time frame of income verification requested: " & verif_time_frame & vbCr
            If requested_CEI_OHI_docs_checkbox = checked Then email_body = email_body & "* Requested Health Insurance information available from employer." & vbCr
        End If
        email_body = email_body & vbCr
        If notes <> "" Then email_body = email_body & "NOTES: " & ntoes & vbCr
        If worker_signature <> "UUDDLRLRBA" Then
            email_body = email_body & vbCr
            email_body = email_body & worker_signature & vbCr
        End If

        Call create_outlook_email(all_email_recipients, "", "Job Change Reported for MX Case", email_body, "", TRUE)
    End If
End If
If refused_empl_yn = "?" Then refused_empl_yn = "N/A"
If good_cause_yn = "?" Then good_cause_yn = "N/A"

If job_change_type = "New Job Reported" Then verif_type_requested = "new job"                                       'setting a variale for entry into heaters/notes/TIKL
If job_change_type = "Income/Hours Change for Current Job" Then verif_type_requested = "change in current job"
If job_change_type = "Job Ended" Then verif_type_requested = "job ended"
'This sets a TIKL if requested and NOT in developer mode
If TIKL_checkbox = checked and developer_mode = FALSE Then Call create_TIKL("Verification of " & verif_type_requested & " due.", 10, verif_form_date, TURE, TIKL_note_text)

If IsDate(verif_form_date) = TRUE and developer_mode = FALSE Then
    'Send a SPEC/MEMO to help support the verification needed from the client.
    'THIS DOES NOT REPLACE THE VERIFICATION REQUEST FORM
    Call back_to_SELF
    Call MAXIS_background_check
    CALL start_a_new_spec_memo

    CALL write_variable_in_SPEC_MEMO("---Verification Needed of Job Change for " & employee_name_only & "---")
    CALL write_variable_in_SPEC_MEMO("")
    CALL write_variable_in_SPEC_MEMO("You reported that " & employee_name_only & " has had a change in employment detail. Type - " & job_change_type)
    CALL write_variable_in_SPEC_MEMO("Job that changed: " & job_employer_name & ".")
    CALL write_variable_in_SPEC_MEMO("Verification needed for the time period: " & verif_time_frame & ". This means we need all income or work hours verification during this time.")
	CALL write_variable_in_SPEC_MEMO("")
    Select Case job_change_type
        Case "New Job Reported"
            CALL write_variable_in_SPEC_MEMO("Since a new job has started, we need verification of any pay you have received so far. If this income covers a 30 day span from the first pay to the most recent, and provides income that reflects the amount you anticipate being paid, this verification should be sufficient. Otherwise, provide verification of the anticipated rate of pay and hours scheduled per week.")
        Case "Income/Hours Change for Current Job"
            CALL write_variable_in_SPEC_MEMO("Since this is a change in job income and/or hours we need proof of this change in income. Provide all income verification from the first pay impacted by this change. If this income covers a 30 day span, this verification should be sufficient. Otherwise, provide verification of the anticipated rate of pay and hours scheduled per week.")
        Case "Job Ended"
            CALL write_variable_in_SPEC_MEMO("To verify the end of employment, you must provide verification of the end of work, including the last day of work, and the date and amount of your last pay.")
            CALL write_variable_in_SPEC_MEMO("We also need to know the nature of the end of employment. If you left the job voluntarily (quit) we need to know the reason for you leaving.")
    End Select
    CALL write_variable_in_SPEC_MEMO("")
    CALL write_variable_in_SPEC_MEMO("If you have questions about verifications needed, or if your job change requires explanation, please call as much of the clarification can be provided verbally and is our best means to correctly budget your income.")
	CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us")
    PF4
    Call back_to_SELF
End If

'If we are indeveloper mode, the script will show the updates in a Dialog
'This will allow the information to be output to a Word Document if desired
If developer_mode = TRUE Then
    y_pos = 15
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 550, 400, "Message Display of CASE NOTE"
      Text 10, y_pos, 530, 10, "Change in Income Reported: " & UCase(verif_type_requested) & " - Reported on: " & reported_date
      y_pos = y_pos + 15
      Text 10, y_pos, 530, 10, "=== Details of the Reported Change ==="
      y_pos = y_pos + 10
      Text 10, y_pos, 530, 10, "* Job Name: " & job_employer_name & " - Income Type: " & job_income_type
      y_pos = y_pos + 10
      Text 10, y_pos, 530, 10, "* Employee: Memb " & job_employee
      y_pos = y_pos + 10
      If job_income_end <> "" Then Text 10, y_pos, 530, 10, "* Income Start Date: " & job_income_start & " - End Date: " & job_income_end
      If job_income_end = "" Then Text 10, y_pos, 530, 10, "* Income Start Date: " & job_income_start
      y_pos = y_pos + 10
      Text 10, y_pos, 530, 10, "* Verification: " & job_verification
      y_pos = y_pos + 15
      Text 10, y_pos, 530, 10, "* Type of Change: " & job_change_type
      y_pos = y_pos + 10
      Select Case job_change_type
          Case "New Job Reported"
              Text 10, y_pos, 530, 10, "Work Start Date: " & date_work_started
              y_pos = y_pos + 10
              If initial_check_gross_amount <> "" Then Text 10, y_pos, 530, 10, "Income Start Date: " & new_job_income_start & " - Initial Gross Pay: $" & initial_check_gross_amount
              If initial_check_gross_amount = "" Then Text 10, y_pos, 530, 10, "Income Start Date: " & new_job_income_start
              y_pos = y_pos + 10
              Text 10, y_pos, 530, 10, "Anticipated Income: Hours per Week: " & new_job_hours_per_week & " - Hourly Wage: $" & new_job_hourly_wage & "/hour"
          Case "Income/Hours Change for Current Job"
              Text 10, y_pos, 530, 10, "Change happened on " & job_change_date & " change will cause " & income_change_type
              y_pos = y_pos + 10
              Text 10, y_pos, 530, 10, "Date of pay first impacted: " & first_pay_date_of_change
              y_pos = y_pos + 10
              Text 10, y_pos, 530, 10, "Change: " & job_change_details
              y_pos = y_pos + 10
              Text 10, y_pos, 530, 10, "Previous Income: Hours per Week: " & job_change_old_hours_per_week & " - Hourly Wage: $" & job_change_old_hourly_wage & "/hour"
              y_pos = y_pos + 10
              Text 10, y_pos, 530, 10, "New Income: Hours per Week: " & job_change_new_hours_per_week & " - Hourly Wage: $" & job_change_new_hourly_wage & "/hour"
          Case "Job Ended"
              Text 10, y_pos, 530, 10, "Income End Date: " & job_end_income_end_date & " - Final pay amount: "& last_pay_amount
              y_pos = y_pos + 10
              Text 10, y_pos, 530, 10, "Work stoped on: " & date_work_ended
              y_pos = y_pos + 10
              Text 10, y_pos, 530, 10, "Quit details:"
              y_pos = y_pos + 10
              Text 10, y_pos, 530, 10, " - Employee refused employment: " & refused_empl_yn
              y_pos = y_pos + 10
              Text 10, y_pos, 530, 10, " - Was this a voluntary quit? " & vol_quit_yn
              y_pos = y_pos + 10
              Text 10, y_pos, 530, 10, " - Reason for STWK: " & stwk_reason
              y_pos = y_pos + 10
              Text 10, y_pos, 530, 10, "    - Meets good cause? " & good_cause_yn
      End Select
      y_pos = y_pos + 15
      Text 10, y_pos, 530, 10, "* Impact on WREG/ABAWD: " & wreg_abawd_notes
      y_pos = y_pos + 10
      Text 10, y_pos, 530, 10, "Information about job gathered in conversation with " & conversation_with_person
      y_pos = y_pos + 10
      Text 10, y_pos, 530, 10, "  - Details of conversation: " & conversation_detail
      y_pos = y_pos + 15
      Text 10, y_pos, 530, 10, "=== Reporting Information ==="
      y_pos = y_pos + 10
      Text 10, y_pos, 530, 10, "* Reported via " & job_report_type & " by " & person_who_reported_job & " on " & reported_date
      If work_number_verbal_checkbox = checked Then
        y_pos = y_pos + 10
        Text 10, y_pos, 530, 10, "*** Verbal authorization to check the Work Number received."
      End If
      y_pos = y_pos + 15
      Text 10, y_pos, 530, 10, "=== Verification ==="
      y_pos = y_pos + 10
      If work_number_checkbox = checked Then Text 10, y_pos, 530, 10, "* Sent Work Number request for income verification."
      If verif_form_date <> "" Then
          y_pos = y_pos + 10
          Text 10, y_pos, 530, 10, "* Verification request sent on " & verif_form_date
          y_pos = y_pos + 10
          Text 10, y_pos, 530, 10, "* Time frame of income verification requested: " & verif_time_frame
          y_pos = y_pos + 10
          If requested_CEI_OHI_docs_checkbox = checked Then Text 10, y_pos, 530, 10, "* Requested Health Insurance information available from employer."
      End If
      y_pos = y_pos + 15
      Text 10, y_pos, 530, 10, "---"
      y_pos = y_pos + 10
      Text 10, y_pos, 530, 10, "NOTES" & notes
      y_pos = y_pos + 10
      Text 10, y_pos, 530, 10, TIKL_note_text
      y_pos = y_pos + 10
      If worker_signature <> "UUDDLRLRBA" Then
          Text 10, y_pos, 530, 10, "---"
          y_pos = y_pos + 10
          Text 10, y_pos, 530, 10, worker_signature
      End If
      ButtonGroup ButtonPressed
        OkButton 495, 380, 50, 15
      CheckBox 10, 385, 290, 10, "Check here to have the script add the information listed here into a word document.", export_note_to_word
    EndDialog

    Do

        dialog Dialog1      'showing the dialog

        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE

    If export_note_to_word = checked Then                   'If the checkbox to export to a Word Document - this is wehre it creates the word document.
        Set objWord = CreateObject("Word.Application")
        Const wdDialogFilePrint = 88
        Const end_of_doc = 6
        objWord.Caption = "Outside Resource Information"
        objWord.Visible = True

        Set objDoc = objWord.Documents.Add()
        Set objSelection = objWord.Selection

        objSelection.PageSetup.LeftMargin = 50
        objSelection.PageSetup.RightMargin = 50
        objSelection.PageSetup.TopMargin = 30
        objSelection.PageSetup.BottomMargin = 25

        todays_date = date & ""
        objSelection.Font.Name = "Arial"
        objSelection.Font.Size = "14"
        objSelection.Font.Bold = TRUE
        objSelection.TypeText "Change in Income Reported: " & UCase(verif_type_requested) & " - Reported on: " & reported_date
        objSelection.TypeParagraph()
        objSelection.ParagraphFormat.SpaceAfter = 0

        objSelection.Font.Size = "12"
        objSelection.Font.Bold = FALSE

        objSelection.TypeText "=== Details of the Reported Change ===" & vbCr
        objSelection.TypeText "* Job Name: " & job_employer_name & " - Income Type: " & job_income_type & vbCr
        objSelection.TypeText "* Employee: Memb " & job_employee & vbCr
        If job_income_end <> "" Then objSelection.TypeText "* Income Start Date: " & job_income_start & " - End Date: " & job_income_end & vbCr
        If job_income_end = "" Then objSelection.TypeText "* Income Start Date: " & job_income_start & vbCr
        objSelection.TypeText "* Verification: " & job_verification & vbCr
        objSelection.TypeText "* Type of Change: " & job_change_type & vbCr
        Select Case job_change_type
            Case "New Job Reported"
                objSelection.TypeText "Work Start Date: " & date_work_started & vbCr
                If initial_check_gross_amount <> "" Then objSelection.TypeText "Income Start Date: " & new_job_income_start & " - Initial Gross Pay: $" & initial_check_gross_amount & vbCr
                If initial_check_gross_amount = "" Then objSelection.TypeText "Income Start Date: " & new_job_income_start & vbCr
                objSelection.TypeText "Anticipated Income: Hours per Week: " & new_job_hours_per_week & " - Hourly Wage: $" & new_job_hourly_wage & "/hour" & vbCr
            Case "Income/Hours Change for Current Job"
                objSelection.TypeText "Change happened on " & job_change_date & " change will cause " & income_change_type & vbCr
                objSelection.TypeText "Date of pay first impacted: " & first_pay_date_of_change & vbCr
                objSelection.TypeText "Change: " & job_change_details & vbCr
                objSelection.TypeText "Previous Income: Hours per Week: " & job_change_old_hours_per_week & " - Hourly Wage: $" & job_change_old_hourly_wage & "/hour" & vbCr
                objSelection.TypeText "New Income: Hours per Week: " & job_change_new_hours_per_week & " - Hourly Wage: $" & job_change_new_hourly_wage & "/hour" & vbCr
            Case "Job Ended"
                objSelection.TypeText "Income End Date: " & job_end_income_end_date & " - Final pay amount: "& last_pay_amount & vbCr
                objSelection.TypeText "Work stoped on: " & date_work_ended & vbCr
                objSelection.TypeText "* Quit details:" & vbCr
                objSelection.TypeText " - Employee refused employment: " & refused_empl_yn & vbCr
                objSelection.TypeText " - Was this a voluntary quit? " & vol_quit_yn & vbCr
                If trim(stwk_reason) <> "" Then objSelection.TypeText " - Reason for STWK: " & stwk_reason & vbCr
                objSelection.TypeText " - Meets good cause? " & good_cause_yn & vbCr
        End Select
        objSelection.TypeText "* Impact on WREG/ABAWD: " & wreg_abawd_notes & vbCr
        If conversation_with_person <> "" Then
            objSelection.TypeText "Information about job gathered in conversation with " & conversation_with_person & vbCr
            objSelection.TypeText "  - Details of conversation: " & conversation_detail & vbCr
        End If
        objSelection.TypeText "=== Reporting Information ===" & vbCr
        objSelection.TypeText "* Reported via " & job_report_type & " by " & person_who_reported_job & " on " & reported_date & vbCr
        If work_number_verbal_checkbox = checked Then
            objSelection.TypeText "*** Verbal authorization to check the Work Number received." & vbCr
        End If
        objSelection.TypeText "=== Verification ===" & vbCr
        If work_number_checkbox = checked Then objSelection.TypeText "* Sent Work Number request for income verification." & vbCr
        If verif_form_date <> "" Then
            objSelection.TypeText "* Verification request sent on " & verif_form_date & vbCr
            objSelection.TypeText "* Time frame of income verification requested: " & verif_time_frame & vbCr
            If requested_CEI_OHI_docs_checkbox = checked Then objSelection.TypeText "* Requested Health Insurance information available from employer." & vbCr
        End If
        objSelection.TypeText "---" & vbCr
        objSelection.TypeText "NOTES: " & notes & vbCr
        objSelection.TypeText TIKL_note_text & vbCr
        If worker_signature <> "UUDDLRLRBA" Then
            objSelection.TypeText "---" & vbCr
            objSelection.TypeText worker_signature & vbCr
        End If
    End If
Else        'If we are NOT in developer mode the script will create a CASE:NOTE now

    Call start_a_blank_CASE_NOTE

    Call write_variable_in_CASE_NOTE("Change in Income Reported: " & UCase(verif_type_requested) & " - Reported on: " & reported_date)
    Call write_variable_in_CASE_NOTE("=== Details of the Reported Change ===")
    Call write_variable_in_CASE_NOTE("* Job Name: " & job_employer_name & " - Income Type: " & job_income_type)
    Call write_variable_in_CASE_NOTE("* Employee: Memb " & job_employee)
    If job_income_end <> "" Then Call write_variable_in_CASE_NOTE("* Income Start Date: " & job_income_start & " - End Date: " & job_income_end)
    If job_income_end = "" Then Call write_variable_in_CASE_NOTE("* Income Start Date: " & job_income_start)
    Call write_variable_in_CASE_NOTE("* Verification: " & job_verification)
    Call write_variable_in_CASE_NOTE("* Type of Change: " & job_change_type)
    Select Case job_change_type
        Case "New Job Reported"
            Call write_variable_with_indent_in_CASE_NOTE("Work Start Date: " & date_work_started)
            If initial_check_gross_amount <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Income Start Date: " & new_job_income_start & " - Initial Gross Pay: $" & initial_check_gross_amount)
            If initial_check_gross_amount = "" Then Call write_variable_with_indent_in_CASE_NOTE("Income Start Date: " & new_job_income_start)
            Call write_variable_with_indent_in_CASE_NOTE("Anticipated Income: Hours per Week: " & new_job_hours_per_week & " - Hourly Wage: $" & new_job_hourly_wage & "/hour")
        Case "Income/Hours Change for Current Job"
            Call write_variable_with_indent_in_CASE_NOTE("Change happened on " & job_change_date & " change will cause " & income_change_type)
            Call write_variable_with_indent_in_CASE_NOTE("Date of pay first impacted: " & first_pay_date_of_change)
            Call write_variable_with_indent_in_CASE_NOTE("Change: " & job_change_details)
            Call write_variable_with_indent_in_CASE_NOTE("Previous Income: Hours per Week: " & job_change_old_hours_per_week & " - Hourly Wage: $" & job_change_old_hourly_wage & "/hour")
            Call write_variable_with_indent_in_CASE_NOTE("New Income: Hours per Week: " & job_change_new_hours_per_week & " - Hourly Wage: $" & job_change_new_hourly_wage & "/hour")
        Case "Job Ended"
            Call write_variable_with_indent_in_CASE_NOTE("Income End Date: " & job_end_income_end_date & " - Final pay amount: $"& last_pay_amount)
            Call write_variable_with_indent_in_CASE_NOTE("Work stoped on: " & date_work_ended)
            Call write_variable_in_CASE_NOTE("* Quit details:")
            Call write_variable_with_indent_in_CASE_NOTE("Employee refused employment: " & refused_empl_yn)
            Call write_variable_with_indent_in_CASE_NOTE("Was this a voluntary quit? " & vol_quit_yn)
            If trim(stwk_reason) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Reason for STWK: " & stwk_reason)
            Call write_variable_with_indent_in_CASE_NOTE("Meets good cause? " & good_cause_yn)
    End Select
    Call write_variable_in_CASE_NOTE("* Impact on WREG/ABAWD: " & wreg_abawd_notes)
    If conversation_with_person <> "" Then
        Call write_variable_in_CASE_NOTE("Information about job gathered in conversation with " & conversation_with_person)
        Call write_variable_in_CASE_NOTE("  - Details of conversation: " & conversation_detail)
    End If
    Call write_variable_in_CASE_NOTE("=== Reporting Information ===")
    Call write_variable_in_CASE_NOTE("* Reported via " & job_report_type & " by " & person_who_reported_job & " on " & reported_date)
    If work_number_verbal_checkbox = checked Then Call write_variable_in_CASE_NOTE("*** Verbal authorization to check the Work Number received.")
    If verif_form_date <> "" Then
        Call write_variable_in_CASE_NOTE("=== Verification ===")
        If work_number_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Sent Work Number request for income verification.")
        Call write_variable_in_CASE_NOTE("* Verification request sent on " & verif_form_date)
        Call write_variable_in_CASE_NOTE("* Time frame of income verification requested: " & verif_time_frame)
        If requested_CEI_OHI_docs_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Requested Health Insurance information available from employer.")
        Call write_variable_in_CASE_NOTE("---")
    End If
    Call write_bullet_and_variable_in_CASE_NOTE("NOTES", notes)
    Call write_variable_in_CASE_NOTE(TIKL_note_text)
    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)
End If

'Creation of detail in the end message
end_msg = "Success! Job change information received. Change reported: " & job_change_type & vbNewLine & vbNewLine
If developer_mode = TRUE Then
    If send_email_checkbox = checked Then end_msg = end_msg & "Email sent with job change information. " & vbNewLine
    If export_note_to_word = checked Then end_mdg = end_msg & "Word document created with job detail information. " & vbNewLine
    end_msg = end_msg & vbNewLine
End If
If script_update_stat = "Yes - Update an existing JOBS Panel" Then  end_msg = end_msg & "Existing panels updated with job change information. " & vbNewLine
If script_update_stat = "Yes - Create a new JOBS Panel" Then end_msg = end_msg & "New panel created and updated with job change information. " & vbNewLine
If TIKL_checkbox = checked and developer_mode = FALSE Then end_msg = end_msg & "TIKL set for return of verification." & vbNewLine

call script_end_procedure_with_error_report(end_msg)
