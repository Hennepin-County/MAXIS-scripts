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
call changelog_update("05/24/2024", "Initial version.", "Mark Riegel and Megan Geissler, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
'connecting to MAXIS & grabbing the case number
EMConnect ""
get_county_code
Call check_for_MAXIS(False)
CALL MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

testing_status = false	'Testing Status: Update status to true to display all testing msgbox or false to hide all testing msgbox
'Initial dialog to gather case details


'Initial Case Number Dialog 
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 221, 125, "Job Change Reported Case Number Dialog"
  EditBox 75, 5, 50, 15, MAXIS_case_number
  EditBox 75, 25, 20, 15, MAXIS_footer_month
  EditBox 105, 25, 20, 15, MAXIS_footer_year
  EditBox 75, 45, 140, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 65, 100, 45, 15
    PushButton 115, 100, 50, 15, "Instructions", msg_show_instructions_btn
    CancelButton 170, 100, 45, 15
  Text 20, 10, 50, 10, "Case Number:"
  Text 20, 30, 45, 10, "Footer month:"
  Text 10, 50, 60, 10, "Worker Signature:"
  Text 10, 75, 200, 20, "Script Purpose: Captures details and updates STAT panels of a change to a job reported, but not verified to the agency."
EndDialog

Do 
    Do
        err_msg = ""
        Dialog Dialog1
        cancel_without_confirmation
        Call validate_MAXIS_case_number(err_msg, "*")
        Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
        If trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Sign your case note."
        
        If ButtonPressed = msg_show_instructions_btn Then 
            err_msg = "LOOP"
            run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/ACTIONS/ACTIONS%20-%20JOB%20CHANGE%20REPORTED.docx"
        End If
        IF err_msg <> "" and err_msg <> "LOOP" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'MAXIS and Case Verifications 
Call check_for_MAXIS(False)
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "ADDR", is_this_priv)
If is_this_priv = True then script_end_procedure("This case is privileged and cannot be accessed. The script will now stop.")
EmReadscreen county_code, 4, 21, 21
If county_code <> worker_county_code then script_end_procedure("This case is out-of-county, and cannot access CASE/NOTE. The script will now stop.")


'Verification Dialog - Potential Redirect to EIB
Do
    Do
        err_msg = ""
        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 240, 80, "Verification Received?"
            DropListBox 10, 35, 225, 15, "Select One..."+chr(9)+"Yes - sufficient verification received to budget income accurately."+chr(9)+"No - we need to request additional verification.", verifs_received_selection
            ButtonGroup ButtonPressed
                OkButton 125, 55, 50, 15
                CancelButton 180, 55, 50, 15
            Text 10, 10, 200, 20, "Have we received sufficient verification to budget the income for this job change?"
        EndDialog
        Dialog Dialog1
        cancel_without_confirmation
        If verifs_received_selection = "Select One..." Then err_msg = err_msg & vbNewLine & "* Indicate if we have sufficient verification or are requesting more."
        If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

'Here, if there are verifications indicated, we will be using Earned Income Budgeting. This will alert the worker to this fact AND redirect.
If verifs_received_selection = "Yes - sufficient verification received to budget income accurately." Then
    MsgBox "~~~Redirecting to ACTIONS - Earned Income Budgeting Script~~~" & vbcr & vbcr & "Script Purpose: Determines an accurate budget for case with Earned Income once verifications are received in the agency."
    Call run_from_GitHub(script_repository & "actions/earned-income-budgeting.vbs")
End If


'Defining Member lists/dropdowns
call generate_client_list(list_of_hh_membs_with_jobs, "New Job Reported")
client_name_array = split(list_of_hh_membs_with_jobs, chr(9))
call generate_client_list(list_of_all_hh_members, "Select or Type")
Call Generate_Client_List(HH_Memb_DropDown, "Select One:")

hh_memb_and_current_jobs =  "*"
hh_memb_5_jobs_panels = "*"

'looping through all the reference numbers so that we can compile a list of members and respective JOBS panels.
For i = 1 to ubound(client_name_array) 	
    Do 
    	Call Navigate_to_MAXIS_screen("STAT", "JOBS")                       'Go to JOBS
        EMReadScreen nav_check, 4, 2, 45
        EMWaitReady 0, 0
    Loop until nav_check = "JOBS"
	EMWriteScreen left(client_name_array(i), 2), 20, 76                 'Enter the reference for the clients in turn to check each.
	EMWriteScreen "01", 20, 79
	transmit

	EMReadScreen total_jobs, 1, 2, 78                                   'look for how many JOBS panels there are so we can loop through them all
	If total_jobs <> "0" Then                                           'if there are no JOBS panels listed for this member, we should't try to read them.
		Do
			EMReadScreen employer_name, 30, 7, 42                       'reading the employer name
			employer_name = replace(employer_name, "_", "")             'taking out the underscores
			EMReadScreen jobs_panel_number, 1, 2, 73                    'reading job panel number 
			jobs_panel_number = "0" & jobs_panel_number 

			If jobs_panel_number = "05" then hh_memb_5_jobs_panels = hh_memb_5_jobs_panels & client_name_array(i) & "*"     'Adding job/member to string to that we can handle for members with 5 JOBS panels 
			If testing_status = True Then msgbox "HH memebers with 5 jobs panels:" & hh_memb_5_jobs_panels

			hh_memb_and_current_jobs = hh_memb_and_current_jobs & client_name_array(i) &  " (" & employer_name & " - " & "JOBS " & jobs_panel_number & ")*"         'Creating a list of members/jobs for the next dialog 
			transmit
			EMReadScreen last_job, 7, 24, 2
		Loop until last_job = "ENTER A"
	End If
Next

'Formatting member/job 
hh_memb_and_current_jobs = replace(hh_memb_and_current_jobs, "*", chr(9))
hh_memb_and_current_jobs = left(hh_memb_and_current_jobs, len(hh_memb_and_current_jobs) - 1)
If testing_status = True Then msgbox "current jobs" & hh_memb_and_current_jobs


'Job Change Selection Dialog 
Dialog1 = ""

BeginDialog Dialog1, 0, 0, 411, 135, "Job Change Selection - Case: " & MAXIS_case_number
  DropListBox 140, 5, 265, 15, "Select One ..."+chr(9)+"No JOBS Panel Exists"+ hh_memb_and_current_jobs, hh_member_current_jobs
  DropListBox 295, 25, 110, 15, HH_Memb_DropDown, hh_memb_with_new_job
  DropListBox 70, 50, 140, 15, "Select One ..."+chr(9)+"New Job Reported"+chr(9)+"Income/Hours Change for Current Job"+chr(9)+"Job Ended", job_change_type
  ComboBox 295, 50, 110, 15, list_of_all_hh_members, person_who_reported_job
  ComboBox 100, 70, 110, 15, "Type or Select"+chr(9)+"phone call"+chr(9)+"Change Report Form"+chr(9)+"office visit"+chr(9)+"mailing"+chr(9)+"fax"+chr(9)+"ES counselor"+chr(9)+"CCA worker"+chr(9)+"scanned document", job_report_type
  EditBox 295, 70, 55, 15, reported_date
  CheckBox 10, 90, 275, 10, "Check here if the employee gave verbal authorization to check the Work Number", work_number_verbal_checkbox
  ButtonGroup ButtonPressed
    OkButton 295, 115, 50, 15
    CancelButton 355, 115, 50, 15
  Text 245, 75, 50, 10, "Date reported:"
  Text 35, 30, 260, 10, "If 'No JOBS Panel Exists' selected above, select member to create JOBS panel:"
  Text 10, 10, 130, 10, "Select Member/JOBS Panel impacted:"
  Text 10, 55, 60, 10, "Job Change Type:"
  Text 245, 55, 50, 10, "Who reported?"
  Text 10, 75, 90, 10, "How was the job reported?"
EndDialog

Do
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
        If hh_member_current_jobs = "Select One:" Then err_msg = err_msg & vbNewLine & "* Must make selection for Select Member/JOBS Panel impacted"
        If hh_memb_with_new_job <>  "Select One:" AND hh_member_current_jobs <> "No JOBS Panel Exists" Then err_msg = err_msg & vbNewLine & "* Member/JOBS Panel impacted must equal 'No JOBS Panel Exists' if Member to create JOBS panel is selected"
        If hh_member_current_jobs = "No JOBS Panel Exists" AND hh_memb_with_new_job = "Select One:" Then err_msg = err_msg & vbNewLine & "* Member to create JOBS panel must be selected if Member/JOBS Panel impacted is 'No JOBS Panel Exists'" 
        If job_change_type = "New Job Reported" AND (hh_memb_with_new_job = "Select One:" OR hh_member_current_jobs <> "No JOBS Panel Exists") Then err_msg = err_msg & vbNewLine & "* Change Type: New Job Reported, Member/JOBS panel impacted must be 'No JOBS Panel Exists' and Member to create JOBS panel must be selected."
        If job_change_type = "Select One ..." Then err_msg = err_msg & vbNewLine & "* Select valid Job Change Type"
		If person_who_reported_job = "Select or Type" Then err_msg = err_msg & vbNewLine & "* Select who reported the job change"
		If job_report_type = "Type or Select" Then err_msg = err_msg & vbNewLine & "* Select how the job was reported"
		If IsDate(reported_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter valid date reported." 
        If err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = "" 
    If instr(hh_memb_5_jobs_panels, hh_memb_with_new_job) then Call script_end_procedure("Script is unable to add another JOBS panel for " & hh_memb_with_new_job & " as there are 5 JOBS panels already. The script will now end." & vbcr & vbcr &  "Please delete a JOBS panel for the Household Member and then restart the script.")
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in




If hh_member_current_jobs <> "No JOBS Panel Exists" Then            'Reading JOBS panel details for the job selected in the JOb Selection Dialog 
    Do 
    	Call Navigate_to_MAXIS_screen("STAT", "JOBS")                       'Go to JOBS
        EMReadScreen nav_check, 4, 2, 45
        EMWaitReady 0, 0
    Loop until nav_check = "JOBS"
	EmWriteScreen left(hh_member_current_jobs, 2), 20, 76
    job_version_only = left(hh_member_current_jobs, Len(hh_member_current_jobs) - 1)    'Pulling out job instance from hh_member_current_jobs
	call write_value_and_transmit(right(job_version_only, 2), 20, 79)

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

	EMWriteScreen "X", 19, 38                               'opening the PIC'
	transmit

	EMReadScreen job_pic_calculation_date, 8, 5, 34         'Getting the detail already listed on the PIC
	EMReadScreen job_pic_pay_frequency, 1, 5, 64
	EMReadScreen job_pic_hours_per_week, 6, 8, 64
	EMReadScreen job_pic_pay_per_hour, 8, 9, 66

	PF3                                                     'leaving the PIC

	job_update_date = replace(job_update_date, " ", "/")    'formatting the date to look like a date
	job_instance = "0" & job_instance                       'making the instance 2 digits
    ref_nbr =  left(hh_member_current_jobs, 2)
    If testing_status = True Then msgbox "job_instance" & job_instance & vbcr & "ref_nbr" & ref_nbr
  
	If job_income_type = "J" Then job_income_type = "J - WIOA"              'formatting the income type to have the code and PF1 information
	If job_income_type = "W" Then job_income_type = "W - Wages"
	If job_income_type = "E" Then job_income_type = "E - EITC"
	If job_income_type = "G" Then job_income_type = "G - Experience Works"
	If job_income_type = "F" Then job_income_type = "F - Federal Work Study"
	If job_income_type = "S" Then job_income_type = "S - State Work Study"
	If job_income_type = "O" Then job_income_type = "O - Other"
	If job_income_type = "C" Then job_income_type = "C - Contract Income"
	If job_income_type = "T" Then job_income_type = "T - Training Program"
	If job_income_type = "P" Then job_income_type = "P - Service Program"
	If job_income_type = "R" Then job_income_type = "R - Rehab Program"
	If job_income_type = "N" Then job_income_type = "N - Census Income"

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
	If job_update_date = date Then 
        script_update_stat = "No - Update of JOBS not needed"       'If job panel has an updated date of today, default script_update_status = no (since in theory it's already been updated)
    Else 
        script_update_stat = "Yes - Update an existing JOBS Panel"  'Defaulting to Yes update, if jobs panel updated date is not today 
    End If
Else 
    script_update_stat = "Yes - Create a new JOBS Panel"            'If user did not select an existing job, it will default to creating a new JOBS panel
End If


'This function will read what programs are on the case and what the case status is.
Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
EMReadScreen case_appl_date, 8, 8, 29											'reading the date the case was APPLd so we know how far back we can go for new jobs start date
case_appl_date = DateAdd("d", 0, case_appl_date)


If job_change_type = "Job Ended" Then
    For i = 1 to Ubound(client_name_array)                         'Now we are going to check WREG for the PWE
        Do 
            Call Navigate_to_MAXIS_screen("STAT", "WREG")                       'Go to JOBS
            EMReadScreen nav_check, 4, 2, 48
            EMWaitReady 0, 0
        Loop until nav_check = "WREG"
        EMWriteScreen left(client_name_array(i), 2), 20, 76
        transmit

        If testing_status = True Then msgbox "Should be at left(client_name_array(i), 2) " & left(client_name_array(i), 2)

        EMReadScreen pwe_code, 1, 6, 68                            'Read PWE code
        If pwe_code = "Y" Then                                     'If it is 'Y' - save the reference number to identify the PWE
            pwe_ref = left(client_name_array(i), 2)
            Exit For
        End If
    Next
End If

'Here are some presets for the main dialog
If job_change_type = "New Job Reported" Then                'If the panel already exists and it has an income start date, we can default to that date
    If IsDate(job_income_start) = TRUE Then new_job_income_start = job_income_start
End If
If job_change_type = "Job Ended" Then                       'If the panel exists and there is already an end date, we can default to that date
    If IsDate(job_income_end) = TRUE Then job_end_income_end_date = job_income_end
End If

If job_verification = "" Then job_verification = "N - No Verif Provided"        'Setting the verification to 'N'

reported_date = date & ""                                   'defaulting the reported date to today
If developer_mode = TRUE Then script_update_stat = "No - Update of JOBS not needed"             'If we are in developer mode, then we cannot update

'Determine the HH memb
If hh_member_current_jobs <> "No JOBS Panel Exists" then
    hh_memb_with_job_change = mid(hh_member_current_jobs, 1, (instr(hh_member_current_jobs, "(") - 2))
Else
    hh_memb_with_job_change = hh_memb_with_new_job
End If 


'This dialog has a different middle part based on the type of report that is happening
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 521, 370, "Job Change Details Dialog"
  GroupBox 5, 5, 250, 250, "Job Information"
  Text 10, 20, 85, 10, "Employee (HH Member)*:"
  Text 10, 40, 60, 10, "Work Number ID:"
'   If job_ref_number <> "" Then Text 475, 65, 50, 10, "JOBS " & job_ref_number & " "  & job_instance
  Text 10, 60, 45, 10, "Income Type:"
  Text 10, 80, 85, 10, "Subsidized Income Type:"
  Text 10, 100, 40, 10, "Employer*:"
  Text 10, 120, 45, 10, "Verification*:"
  Text 10, 140, 55, 10, "Pay Frequency:*"
  Text 10, 160, 60, 10, "Income Start Date:"
  Text 10, 180, 60, 10, "Income End Date:"
  Text 10, 200, 75, 10, "Contract through date:"
  Text 95, 20, 140, 10, hh_memb_with_job_change
'   DropListBox 95, 15, 140, 45, list_of_employees, job_employee
  EditBox 95, 35, 50, 15, work_numb_id
  DropListBox 95, 55, 140, 15, ""+chr(9)+"W - Wages"+chr(9)+"J - WIOA"+chr(9)+"E - EITC"+chr(9)+"G - Experience Works"+chr(9)+"F - Federal Work Study"+chr(9)+"S - State Work Study"+chr(9)+"O - Other"+chr(9)+"C - Contract Income"+chr(9)+"T - Training Program"+chr(9)+"P - Service Program"+chr(9)+"R - Rehab Program", job_income_type
  DropListBox 95, 75, 140, 15, ""+chr(9)+"01 - Subsidized Public Sector Employer"+chr(9)+"02 - Subsidized Private Sector Employer"+chr(9)+"03 - On-the-Job-Training"+chr(9)+"04 - AmeriCorps", job_subsidized_income_type
  EditBox 95, 95, 140, 15, job_employer_name
  DropListBox 95, 115, 95, 45, ""+chr(9)+"N - No Verif Provided"+chr(9)+"? - Delayed Verification", job_verification
  DropListBox 95, 135, 95, 45, ""+chr(9)+"1 - Monthly"+chr(9)+"2 - Semi-Monthly"+chr(9)+"3 - Biweekly"+chr(9)+"4 - Weekly"+chr(9)+"5 - Other", job_pay_frequency
  EditBox 95, 155, 55, 15, job_income_start
  EditBox 95, 175, 55, 15, job_income_end
  EditBox 95, 195, 55, 15, job_contract_through_date

  Select Case job_change_type
      Case "New Job Reported"
        GroupBox 260, 5, 250, 250, "Update Reported - NEW JOB"
        Text 270, 20, 65, 10, "Date work started:"
        Text 270, 40, 105, 10, "Date income started/will start*:"
        Text 270, 60, 95, 10, "Initial check GROSS amount:"
        Text 270, 80, 90, 10, "Anticipated hours per week:"
        Text 270, 100, 80, 10, "Anticipated hourly wage:"
        Text 270, 120, 65, 10, "Conversation with:"
        Text 270, 140, 25, 10, "Details:"
        Text 270, 160, 95, 10, "WREG/ABAWD Impact*:"
        EditBox 375, 15, 55, 15, date_work_started
        EditBox 375, 35, 55, 15, new_job_income_start
        EditBox 375, 55, 55, 15, initial_check_gross_amount
        EditBox 375, 75, 55, 15, new_job_hours_per_week
        EditBox 375, 95, 55, 15, new_job_hourly_wage
        ComboBox 375, 115, 130, 45, list_of_all_hh_members, conversation_with_person
        EditBox 375, 135, 130, 15, conversation_detail
        EditBox 375, 155, 125, 15, wreg_abawd_notes
        CheckBox 270, 175, 145, 10, "Check here if Work Number request sent", work_number_checkbox
      Case "Income/Hours Change for Current Job"
        GroupBox 260, 5, 245, 250, "Update Reported - JOB CHANGE"
        Text 265, 20, 55, 10, "Date of Change:"
        Text 265, 40, 70, 10, "Change Reported*:"
        Text 265, 65, 90, 10, "-- Old Anticipated Income --"
        Text 265, 80, 60, 10, "Hours per Week:"
        Text 395, 80, 50, 10, "Hourly Wage:"
        Text 265, 100, 55, 10, "Income Change:"
        Text 265, 125, 95, 10, "-- New Anticipated Income* --"
        Text 265, 140, 60, 10, "Hours per Week:"
        Text 400, 140, 50, 10, "Hourly Wage:"
        Text 265, 160, 85, 10, "First Pay Date Impacted*:"
        Text 265, 190, 65, 10, "Conversation with:"
        Text 265, 205, 25, 10, "Details:"
        Text 265, 225, 85, 10, "WREG/ABAWD Impact*:"
        EditBox 350, 15, 55, 15, job_change_date
        EditBox 350, 35, 150, 15, job_change_details
        EditBox 350, 75, 40, 15, job_change_old_hours_per_week
        EditBox 455, 75, 45, 15, job_change_old_hourly_wage
        ComboBox 350, 95, 150, 45, "Select or Type"+chr(9)+"Increase"+chr(9)+"Decrease", income_change_type
        EditBox 350, 135, 35, 15, job_change_new_hours_per_week
        EditBox 455, 135, 45, 15, job_change_new_hourly_wage
        EditBox 350, 155, 60, 15, first_pay_date_of_change
        ComboBox 350, 185, 150, 45, list_of_all_hh_members, conversation_with_person
        EditBox 350, 200, 150, 15, conversation_detail
        EditBox 350, 220, 150, 15, wreg_abawd_notes
        CheckBox 265, 240, 190, 10, "Check here if you sent a Work Number request.", work_number_checkbox  
      Case "Job Ended"
        GroupBox 260, 5, 250, 250, "Update Reported - JOB ENDED"
        Text 270, 20, 70, 10, "Date Work Ended*:"
        Text 270, 40, 65, 10, "Last pay amount*:"
        Text 270, 60, 105, 10, "Date income ended/will end*:"
        Text 270, 80, 60, 10, "Reason for STWK:"
        Text 270, 100, 70, 10, "STWK Verification*:"
        Text 270, 115, 55, 10, "Voluntary Quit*:"
        Text 410, 115, 60, 10, "Good cause met?"
        Text 270, 130, 50, 10, "Refused Empl:"
        Text 410, 130, 45, 10, "Refusal Date:"
        Text 270, 150, 160, 10, "Is the client applying for Unemployment Income?"
        Text 270, 170, 65, 10, "Conversation with:"
        Text 270, 185, 25, 10, "Details:"
        Text 270, 205, 85, 10, "WREG/ABAWD Impact*:"
        EditBox 365, 15, 55, 15, date_work_ended
        EditBox 365, 35, 55, 15, last_pay_amount
        EditBox 365, 55, 55, 15, job_end_income_end_date
        EditBox 365, 75, 135, 15, stwk_reason
        DropListBox 365, 95, 135, 45, "Select One..."+chr(9)+"N - No Verif Provided"+chr(9)+"? - Delayed Verification", stwk_verif
        DropListBox 365, 110, 30, 45, "?"+chr(9)+"Yes"+chr(9)+"No", vol_quit_yn
        DropListBox 470, 110, 30, 45, "?"+chr(9)+"Yes"+chr(9)+"No", good_cause_yn
        DropListBox 365, 125, 30, 45, "?"+chr(9)+"Yes"+chr(9)+"No", refused_empl_yn
        EditBox 460, 125, 40, 15, refused_empl_date
        DropListBox 435, 145, 65, 45, "?"+chr(9)+"Yes"+chr(9)+"No", uc_yn
        ComboBox 365, 165, 135, 45, list_of_all_hh_members, conversation_with_person
        EditBox 365, 180, 135, 15, conversation_detail
        EditBox 365, 200, 135, 15, wreg_abawd_notes
        CheckBox 270, 220, 190, 10, "Check here if you sent a Work Number request.", work_number_checkbox
        ' ComboBox 80, 260, 125, 45, "Select One..."+chr(9)+"1 - Employers Statement"+chr(9)+"2 - Seperation Notice"+chr(9)+"3 - Collateral Statement"+chr(9)+"4 - Other Document"+chr(9)+"N - No Verif Provided", stwk_verif
  End Select

  GroupBox 5, 265, 505, 75, "Actions"
  Text 10, 280, 105, 10, "Date verification Request Sent:"
  Text 10, 300, 105, 10, "Time frame of verifs requested:"
  Text 10, 320, 90, 10, "Have Script Update Panel:"
  Text 10, 350, 25, 10, "Notes:"
  EditBox 120, 275, 75, 15, verif_form_date
  EditBox 120, 295, 75, 15, verif_time_frame
  CheckBox 270, 275, 105, 10, "Check here to TIKL for return.", TIKL_checkbox
  CheckBox 270, 320, 165, 10, "Check here if you are requesting CEI/OHI docs.", requested_CEI_OHI_docs_checkbox
  DropListBox 100, 315, 145, 15, "No - Update of JOBS not needed" +chr(9)+ "Yes - Update an existing JOBS Panel" +chr(9)+ "Yes - Create a new JOBS Panel", script_update_stat
  CheckBox 270, 305, 165, 10, "Check here if you sent a status update to CCA.", CCA_checkbox
  CheckBox 270, 290, 160, 10, "Check here if you sent a status update to ES.", ES_checkbox
  EditBox 40, 345, 340, 15, notes
  ButtonGroup ButtonPressed
    OkButton 405, 345, 50, 15
    CancelButton 460, 345, 50, 15
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
        err_msg = err_msg & vbNewLine
        If job_income_type = "" OR trim(job_employer_name) = "" OR job_verification = "" OR job_pay_frequency = "" OR  IsDate(job_income_start) = False Then err_msg = err_msg & vbNewLine & "JOB INFORMATION SECTION"
        If job_income_type = "" Then err_msg = err_msg & vbNewLine & "* Select Income type"
        If trim(job_employer_name) = "" Then err_msg = err_msg & vbNewLine & "* Enter the name of the employer."
        If len(trim(job_employer_name)) > 30 Then err_msg = err_msg & vbNewLine & "* The employer name is more than 30 characters and MAXIS only allows for 30 characters on the employer line. Change the employer name to fit MAIXS."
        If job_verification = "" Then err_msg = err_msg & vbNewLine & "* Enter the verification of the JOBS panel."
        If job_pay_frequency = "" Then err_msg = err_msg & vbNewLine & "* Enter the frequency of the pay (weekly, biweekly, monthly)." 
        If IsDate(job_income_start) = False Then err_msg = err_msg & vbNewLine & "* Enter the date the income will or has started."
        err_msg = err_msg & vbNewLine

        Select Case job_change_type
            Case "New Job Reported"
                IF (trim(new_job_hourly_wage) <> "" AND IsNumeric(new_job_hourly_wage) = False) OR (trim(new_job_hours_per_week) <> "" AND IsNumeric(new_job_hours_per_week) = FALSE) Then err_msg = err_msg & vbNewLine & "UPDATE REPORTED SECTION"
                If trim(new_job_hourly_wage) <> "" AND IsNumeric(new_job_hourly_wage) = False Then err_msg = err_msg & vbNewLine & "* Enter the hourly wage as a number."
                If trim(new_job_hours_per_week) <> "" AND IsNumeric(new_job_hours_per_week) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the hours per week as a number."

            Case "Income/Hours Change for Current Job"
                If trim(job_change_details) = "" OR trim(job_change_new_hourly_wage) = "" OR IsNumeric(job_change_new_hourly_wage) = False OR trim(job_change_new_hours_per_week) = "" OR IsNumeric(job_change_new_hours_per_week) = FALSE OR IsDate(first_pay_date_of_change) = FALSE Then err_msg = err_msg & vbNewLine & "UPDATE REPORTED SECTION"
                If trim(job_change_details) = "" Then err_msg = err_msg & vbNewLine & "* Enter the information about the change reported."
                If trim(job_change_new_hourly_wage) = "" Then
                    err_msg = err_msg & vbNewLine & "* The anticipated hourly wage must be entered."
                Else
                    If IsNumeric(job_change_new_hourly_wage) = False Then err_msg = err_msg & vbNewLine & "* Enter the hourly wage as a number."
                End If
                If trim(job_change_new_hours_per_week) = "" Then
                    err_msg = err_msg & vbNewLine & "* The anticipated hours worked per week must be entered."
                Else
                    If IsNumeric(job_change_new_hours_per_week) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the hours per week as a number."
                End If
                If IsDate(first_pay_date_of_change) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the date the change will be reflected in income (the first paycheck that will be affected by this change)."

            Case "Job Ended"
                If IsDate(date_work_ended) = FALSE OR IsNumeric(last_pay_amount) = FALSE OR  IsDate(job_end_income_end_date) = FALSE OR stwk_verif = "Select One..." OR vol_quit_yn = "?" OR (vol_quit_yn = "Yes" AND good_cause_yn = "?") OR refused_empl_yn = "Yes" AND IsDate(refused_empl_date) = FALSE Then err_msg = err_msg & vbNewLine & "UPDATE REPORTED SECTION"
                If IsDate(date_work_ended) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the date the client last worked."
                If IsNumeric(last_pay_amount) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the amount of the last pay. This can be an estimate amount."
                If IsDate(job_end_income_end_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the date of the last paycheck."
                If stwk_verif = "Select One..." Then err_msg = err_msg & vbNewLine & "* Enter the verification code for the STWK Panel."
                If vol_quit_yn = "?" Then err_msg = err_msg & vbNewLine & "* Select Voluntary Quit information - Yes or No"
                If vol_quit_yn = "Yes" AND good_cause_yn = "?" Then err_msg = err_msg & vbNewLine & "* Since this job has ended voluntarily, indicate if this voluntary quit meets good cause."
                If refused_empl_yn = "Yes" AND IsDate(refused_empl_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Since you indicated that the client refused employment, you must enter the date employment was refused as a valid date."
        End Select
        err_msg = err_msg & vbNewLine
        If trim(wreg_abawd_notes) = "" AND SNAP_case = TRUE Then err_msg = err_msg & vbNewLine & vbNewLine & "* Enter how this change will impact WREG/ABWAD for this member. NOTE: The impact may be 'NO CHANGE' or something similar." & vbNewLine
        If conversation_with_person = "Select or Type" Then     'handling to deal with conversation information - there should be a person listed AND conversation detail if either are listed
            If trim(conversation_detail) <> "" Then err_msg = err_msg & vbNewLine & "* There is information added to the conversation detail but no information about who the conversation was with has been provided. Enter the member or name of the person the conversation was completed with."
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
        If job_change_type = "New Job Reported" AND script_update_stat = "Yes - Update an existing JOBS Panel" Then err_msg = err_msg & vbNewLine & "* You selected 'New Job Reported' on the initial dialog, update 'Have Script Update Panel' to 'Yes - Create a new JOBS Panel'"
        If job_change_type = "Income/Hours Change for Current Job" AND script_update_stat = "Yes - Create a new JOBS Panel" Then err_msg = err_msg & vbNewLine & "* You selected update an existing job on the initial dialog, update 'Have Script Update Panel' to 'Yes - Update an existing JOBS Panel'"

        If trim(verif_time_frame) <> "" AND IsDate(verif_form_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the date the verification request form has been sent."

        If err_msg = vbNewLine & vbNewLine & vbNewLine Then err_msg = ""
        'Displaying the error message
        If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg

    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

job_income_start = DateAdd("d", 0, job_income_start)

Select Case job_change_type                                         'here we are finding the MAXIS Footer Month and Year based on when the change was made
    Case "New Job Reported"
		If DateDiff("m", new_job_income_start, case_appl_date) > 0 Then			'new job start is before case application
			CALL convert_date_into_MAXIS_footer_month(case_appl_date, MAXIS_footer_month, MAXIS_footer_year)			'so we set the initial month to the case appl date
		Else
        	CALL convert_date_into_MAXIS_footer_month(new_job_income_start, MAXIS_footer_month, MAXIS_footer_year)		'otherwise the initial month is the job start month'
		End If
    Case "Income/Hours Change for Current Job"
        CALL convert_date_into_MAXIS_footer_month(first_pay_date_of_change, MAXIS_footer_month, MAXIS_footer_year)
    Case "Job Ended"
        CALL convert_date_into_MAXIS_footer_month(job_end_income_end_date, MAXIS_footer_month, MAXIS_footer_year)
End Select

If conversation_with_person = "Select or Type" Then conversation_with_person = ""

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

'review_vol_quit = FALSE         'This is going to false now because there is no voluntary quit right now - COVID WAIVER
code_disq = FALSE               'defaulting this variable before the additional vol quit detail
'This functionality provides additional information gathering about Volundary Quit and the ability to take action/create seperate notes
If review_vol_quit = TRUE Then
    If mfip_case = TRUE Then mfip_vol_quit_checkbox = checked           'Autochecking the program boxes based upon which programs were found to be active
    If dwp_case = TRUE Then dwp_vol_quit_checkbox = checked
    If snap_case = TRUE Then snap_vol_quit_checkbox = checked
    Do
        Do
            err_msg = ""

            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 186, 265, "Voluntary Quit Detail"
                DropListBox 95, 25, 70, 45, "Select One..."+chr(9)+"Quit Job"+chr(9)+"Hours Reduction", vol_quit_type
                DropListBox 110, 45, 55, 45, "?"+chr(9)+"Yes"+chr(9)+"No", vol_quit_yn
                DropListBox 135, 65, 30, 45, "?"+chr(9)+"Yes"+chr(9)+"No", good_cause_yn
                EditBox 50, 100, 115, 15, vol_quit_reason
                CheckBox 40, 140, 30, 10, "MFIP", mfip_vol_quit_checkbox
                CheckBox 40, 150, 30, 10, "DWP", dwp_vol_quit_checkbox
                CheckBox 40, 160, 30, 10, "SNAP", snap_vol_quit_checkbox
                DropListBox 90, 175, 75, 45, "Select One..."+chr(9)+"1st Sanction"+chr(9)+"2nd Sanction"+chr(9)+"3rd+ Sanctions", snap_vol_quit_occurance
                EditBox 90, 195, 75, 15, disq_begin_date
                EditBox 90, 215, 75, 15, disq_end_date
                ButtonGroup ButtonPressed
                    OkButton 75, 245, 50, 15
                    CancelButton 130, 245, 50, 15
                Text 10, 10, 255, 10, "This case appears to have a voluntary quit situation."
                Text 20, 30, 75, 10, "Type of action taken?"
                Text 20, 50, 90, 10, "Resident voluntarily quit?"
                Text 20, 70, 110, 10, "If voluntary, is there good cause?"
                Text 20, 90, 140, 10, "Explain the cause/if they meet good cause:"
                Text 20, 125, 140, 10, "Programs Impacted by Voluntary Quit"
                Text 20, 180, 70, 10, "If SNAP, occurance:"
                Text 20, 200, 60, 10, "Disq Begin Date:"
                Text 20, 220, 50, 10, "Disq End Date:"
            EndDialog

            dialog Dialog1
            cancel_confirmation

            vol_quit_reason = trim(vol_quit_reason)
            If vol_quit_yn = "?" Then err_msg = err_msg & vbNewLine & "* Indicate if the job was voluntarily quit/hours reduced."
            If vol_quit_yn = "Yes" Then
                If vol_quit_type = "Select One..." Then err_msg = err_msg & vbNewLine & "* Indicate if the job was voluntarily ended or voluntarily reduced hours."
                If len(vol_quit_reason) < 20 Then err_msg = err_msg & vbNewLine & "* Enter full detail of information about the reason the job was ended. If the resident has not provided a reason, list that here and indicate how we are going to determine possible good cause."
                If snap_vol_quit_checkbox = checked AND snap_vol_quit_occurance = "Select One..." Then err_msg = err_msg & vbNewLine & "* This voluntary quit impacts the SNAP program, therefore indicate which occurance of this type of sanction this resident is on."
            End If
            If IsDate(disq_begin_date) = False then err_msg = err_msg & vbNewLine & "* Enter XX/XX/XX date for Disq begin date"
            If IsDate(disq_end_date) = False then err_msg = err_msg & vbNewLine & "* Enter XX/XX/XX date for Disq end date"
            If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg

        Loop until err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE

    If vol_quit_yn = "Yes" AND good_cause_yn = "No" Then code_disq = TRUE       'This is the variable that will cause the DISQ updating later
End If

Call back_to_SELF               'making sure we are in the right place
Call MAXIS_background_check
' MsgBox "Start month - " & MAXIS_footer_month & vbNewLine & "Start year - " & MAXIS_footer_year
employee_name_only = right(hh_memb_with_job_change, len(hh_memb_with_job_change) - 5)

Initial_footer_month = MAXIS_footer_month           'We are going to loop through months, so we need to set the initial footer month so we remember it later
Initial_footer_year = MAXIS_footer_year
second_loop = FALSE         'Knowing where we are

If developer_mode = FALSE Then                      'If we are in developer mode then we are going to skip the update parts
    Do
        EMWriteScreen "SUMM", 20, 71                'Getting into STAT
        transmit
        EMReadScreen MAXIS_footer_month, 2, 20, 55          'Setting the right month and year
        EMReadScreen MAXIS_footer_year, 2, 20, 58

        If testing_status = TRUE Then MsgBox "are we on stat summ?"

        If second_loop = TRUE Then                          'If we are past the first loop, we know the job member and instance so we can naviage directly there
            EMWriteScreen "JOBS", 20, 71                    'go to JOBS
            EMWriteScreen ref_nbr, 20, 76                   'go to the right member
            transmit

            EMReadScreen panel_exists, 14, 24, 13

            If panel_exists = "DOES NOT EXIST" Then
                ref_nbr = left(hh_memb_with_job_change, 2) 
                EMWriteScreen ref_nbr, 20, 76                                       'go to the right member
                EMWriteScreen "NN", 20, 79                                          'create new JOBS panel
                transmit

                EMReadScreen job_instance, 1, 2, 73                                'Reading the instance because we need it on the next loop
                job_instance = "0" & job_instance
                If testing_status = TRUE Then MsgBox "Panel did not exist so one was created with instance: " & job_instance
                
            Else 
                Do 
                    EMReadScreen company_name, 30, 7, 42          'reading the name of job
                    company_name = ucase(replace(company_name, "_", ""))
                    original_full_jobs_name = job_employer_name
                    original_full_jobs_name = ucase(original_full_jobs_name)
                    If testing_status = TRUE Then MsgBox company_name & vbcr & original_full_jobs_name

                    If company_name = original_full_jobs_name Then
                        EMReadScreen job_instance, 1, 2, 73
                        job_instance = "0" & job_instance
                        PF9
                        Exit do
                    Else 
                        transmit
                        If testing_status = TRUE Then MsgBox "we are transmitting to the next panel"
                    End If

                    EMReadScreen reached_last_JOBS_panel, 13, 24, 2
                Loop until reached_last_JOBS_panel = "ENTER A VALID"
                
                If reached_last_JOBS_panel = "ENTER A VALID" Then 
                    Do 
                        EMReadScreen panel_number, 1, 2, 73
                        If panel_number = "5" Then
                            Dialog1 = ""
                            BeginDialog Dialog1, 0, 0, 201, 65, "JOBS PANEL at MAX"
                                ButtonGroup ButtonPressed
                                    OkButton 45, 45, 50, 15
                                    CancelButton 105, 45, 50, 15
                                Text 50, 5, 160, 10, "5 Jobs panels already exist!"
                                Text 20, 20, 165, 10, " Review MAXIS, delete a job panel, then select OK. "
                            EndDialog

                            dialog Dialog1
                            cancel_confirmation
                        Else 
                            If testing_status = TRUE Then MsgBox "less than 5 panles, good to udpate"
                            Exit do
                        End If
                    Loop until panel_number <> "5"
                    If testing_status = TRUE Then MsgBox "Pausing before we fill out the panel"
                    EMWriteScreen Left(hh_memb_with_job_change, 2), 20, 76               'go to the right member
                    EMWriteScreen "NN", 20, 79                                          'create new JOBS panel
                    transmit
                    EMReadScreen job_instance, 1, 2, 73         'Reading job instance so that when we update the next month we have this value
                    job_instance = "0" & job_instance
                    ref_nbr = Left(hh_memb_with_job_change, 2)  'defining this variable so that when we update the next month we have this value
                    If testing_status = TRUE Then msgbox "Did we successfully create a new panel for: " & ref_nbr & "with instance: " & job_instance
                End If
            End If

        ElseIf script_update_stat = "Yes - Update an existing JOBS Panel" Then          'if we are in the first loop then the action is going to change based on the the update detail
            'Navigate to STAT/JOBS
            EMReadScreen total_jobs, 1, 2, 78                               'Reading the number of jobs that are on this case for this member
            If total_jobs = "0" Then                                        'If there are none and the worker indicated to update an existing panel, thes cript will end.
                Call script_end_procedure("Update and Note NOT Completed. There are no jobs for Memb " & hh_memb_with_job_change & " listed in MAXIS and you have selected to have the script Update and existing JOBS panel.")
            Else    
                EMWriteScreen "JOBS", 20, 71                    'If we know the instance we can just navigate to the JOBS panel
                ref_nbr = left(hh_memb_with_job_change, 2) 
                EMWriteScreen ref_nbr, 20, 76                   'go to the right member
                EMWriteScreen job_instance, 20, 79
                transmit 
                If testing_status = TRUE Then MsgBox "are we on jobs with correct member?"
            End If
        
            PF9     'placing JOBS panel in edit mode

        ElseIf script_update_stat = "Yes - Create a new JOBS Panel" Then
            Do 
                Call Navigate_to_MAXIS_screen("STAT", "JOBS")                       'Go to JOBS
                EMReadScreen nav_check, 4, 2, 45
                EMWaitReady 0, 0
            Loop until nav_check = "JOBS"
            EMWriteScreen Left(hh_memb_with_job_change, 2), 20, 76               'go to the right member
            EMWriteScreen "NN", 20, 79                                          'create new JOBS panel
            transmit
            EMReadScreen job_instance, 1, 2, 73         'Reading job instance so that when we update the next month we have this value
            job_instance = "0" & job_instance
            ref_nbr = Left(hh_memb_with_job_change, 2)  'defining this variable so that when we update the next month we have this value
            If testing_status = TRUE Then msgbox "Did we successfully create a new panel for: " & ref_nbr & "with instance: " & job_instance
        End If          'Now we are done with the code for if we are in the first loop of updating


        new_hours_per_check = 0
        old_hours_per_check = 0
        'If the script indicates that we need to update at all here is the part where we actually put the information in the panel
        'The panel should already be in EDIT MODE
        If script_update_stat = "Yes - Update an existing JOBS Panel" OR script_update_stat = "Yes - Create a new JOBS Panel" Then
            ' MsgBox "Update - " & script_update_stat & " - 3"
            EMWriteScreen left(job_income_type, 1), 5, 34                                                               'income type
            If job_subsidized_income_type <> "" Then EMWriteScreen left(job_subsidized_income_type, 2), 5, 74           'subsidized type
            If job_verification <> "" Then EMWriteScreen left(job_verification, 1), 6, 34                              'job verification
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
                IF testing_status = TRUE Then msgbox "Reading old checks" & numb_of_checks & " old job hours" & job_hours 
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
            EMWriteScreen "   ", 18, 43                      'blanking out hours
            EMWriteScreen "   ", 18, 72
            Select Case job_change_type                 'What we enter and how we enter is baed upon what change type is reported.
                'This doesn't actually do the updates but sets up all the variables so we can use the same update code
                Case "New Job Reported"
                    If trim(new_job_hourly_wage) <> "" Then
                        EMWriteScreen "      ", 6, 75
                        EMWriteScreen new_job_hourly_wage, 6, 75            'setting the hourly wage
                    End If

					the_last_pay_date = new_job_income_start                    'these are used for the loops through paychecks - we start when the income actually starts
					the_first_pay_date = new_job_income_start
					If DateDiff("m", new_job_income_start, case_appl_date) > 0 Then			'new job start is before case application
						Do														'we have to find the first check date in the initial month (application month)
							If job_pay_frequency = "1 - Monthly" Then
		                        the_last_pay_date = DateAdd("m", 1, the_last_pay_date)
								the_first_pay_date = DateAdd("m", 1, the_first_pay_date)
		                    ElseIf job_pay_frequency = "2 - Semi-Monthly" Then
		                        the_last_pay_date = DateAdd("d", 15, the_last_pay_date)
								the_first_pay_date = DateAdd("d", 15, the_first_pay_date)
		                    ElseIf job_pay_frequency = "3 - Biweekly" Then
		                        the_last_pay_date = DateAdd("d", 14, the_last_pay_date)
								the_first_pay_date = DateAdd("d", 14, the_first_pay_date)
		                    ElseIf job_pay_frequency = "4 - Weekly" Then
		                        the_last_pay_date = DateAdd("d", 7, the_last_pay_date)
								the_first_pay_date = DateAdd("d", 7, the_first_pay_date)
		                    End If
						Loop until DateDiff("m", the_first_pay_date, case_appl_date) <= 0
					End If
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
                    the_last_pay_date = job_end_income_end_date                 'Saving these for the loops, we start with the last day of pay
                    the_first_pay_date = job_end_income_end_date
                    prosp_hours = 0                                             'setting the prospective hourse to be a number
                    pay_amt = known_pay_amount                                  'setting the pay amount to the amount we found earlier when reading the panel
            End Select

            If end_of_pay = "99/99/99" Then     'Based on the end date, we will set the row to start with
                jobs_row = 12       'If we have no end date we start at the top of the list of pay dates
            Else
                jobs_row = 16       'If we do have an end date, we start at the bottom
                Call create_mainframe_friendly_date(job_end_income_end_date, 9, 49, "YY")       'If we know the end date, this will enter the income end date in the correct spot on the panel.
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
                    If job_hourly_wage = "" or job_hourly_wage = "0.00" Then 
                        prosp_hours = 0 
                        jobs_row = jobs_row - 1  
                    Else 
                        prosp_hours = last_pay_amount/job_hourly_wage
                        jobs_row = jobs_row - 1                     'go to the next row above
                    End If
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
                        If job_hourly_wage = "" or job_hourly_wage = "0.00" Then 
                            hours_of_pay = 0 
                            prosp_hours = prosp_hours + hours_of_pay        'totaling the hours calculated here 
                        Else 
                            job_hourly_wage = job_hourly_wage * 1
                            hours_of_pay = pay_amt/job_hourly_wage
                            prosp_hours = prosp_hours + hours_of_pay        'totaling the hours calculated here
                        End If
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
                        If job_hourly_wage = "" or job_hourly_wage = "0.00" Then 
                            hours_of_pay = 0 
                            prosp_hours = prosp_hours + hours_of_pay        'totaling the hours calculated here 
                        Else 
                            job_hourly_wage = job_hourly_wage * 1
                            hours_of_pay = pay_amt/job_hourly_wage
                            prosp_hours = prosp_hours + hours_of_pay        'totaling the hours calculated here
                        End If
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
                EMWaitReady 0, 0
                EMReadScreen warning_popup, 69, 6, 7
                If warning_popup = "WARNING: PROSPECTIVE EARNINGS EXIST WITH NO HOURS OR HOURS EXIST WITH" then 
                    warning_check = MsgBox("Warning: Prospective earnings exist with no hours or hours exist with no prospective earnings. Is this correct for this job?", vbYesNo) 
                    If warning_check = vbYes Then 
                        EMWriteScreen "Y", 9, 58
                        transmit 
                    ElseIf warning_check = vbNo Then 
                        EMWriteScreen "N", 9, 58
                        transmit
                    End If
                End If
                EMWaitReady 0, 0
                EMWaitReady 0, 0
                EMReadScreen error_warning, 20, 6, 43
                If error_warning = "Error Prone Warnings" Then
                    PF3
                End If
            End If
        End If
        transmit        'saving the panel information

        If testing_status = TRUE Then msgbox "JOBS saved"

        If original_full_jobs_name = "" Then EMReadScreen original_full_jobs_name, 30, 7, 42             'read the employer name as it originally exists on the panel

        If script_update_stat = "Yes - Update an existing JOBS Panel" OR script_update_stat = "Yes - Create a new JOBS Panel" Then      'If we are updating
            If job_change_type = "Job Ended" Then                                                                                       'and we are coding a job end change type
                Do 
                    Call navigate_to_MAXIS_screen("STAT", "STWK")                   'We also have to code end of employmnt on STWK
                    EMReadScreen nav_check, 4, 2, 45
                    EMWaitReady 0, 0
                Loop until nav_check = "STWK"
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

                disq_updated = FALSE 
                If code_disq = TRUE Then        'This will code a DISQ panel in the case of Voluntary Quit
                    Call convert_date_into_MAXIS_footer_month(disq_begin_date, disq_footer_month, disq_footer_year)
                    If MAXIS_footer_month = disq_footer_month AND MAXIS_footer_year = disq_footer_year Then 
                        Do 
                            Call navigate_to_MAXIS_screen("STAT", "DISQ")                   'We also have to code end of employmnt on STWK
                            EMReadScreen nav_check, 4, 2, 48
                            EMWaitReady 0, 0
                        Loop until nav_check = "DISQ"
                        EMWriteScreen ref_nbr, 20, 76
                        transmit
                        EMReadScreen disq_total, 1 , 2, 78
                        If disq_total < "5" Then 
                            EMWriteScreen "NN", 20, 79
                            transmit
                            EMWriteScreen "FS", 6, 54
                            EMWriteScreen "11", 6, 64
                            Call create_MAXIS_friendly_date_with_YYYY(disq_begin_date, 0, 8, 64)
                            Call create_MAXIS_friendly_date_with_YYYY(disq_end_date, 0, 9, 64)
                            transmit 
                            disq_updated = TRUE 
                        Else 
                            code_disq = FALSE 
                            disq_updated = FALSE
                            msgbox "DISQ panel is maxed out, unable to update. Manually update after script run"
                        End If 
                    End If
                End If
            End If
        End If
        transmit 
        CALL write_value_and_transmit("BGTX", 20, 71)   'Now we send the case through background and go to STAT/WRAP
        Do 
            EMReadScreen wrap_check, 4, 2, 46
            If wrap_check <> "WRAP" Then 
                EMReadScreen warning_popup, 69, 6, 7
                If warning_popup = "WARNING: PROSPECTIVE EARNINGS EXIST WITH NO HOURS OR HOURS EXIST WITH" then 

                    warning_check = MsgBox("Warning: Prospective earnings exist with no hours or hours exist with no prospective earnings. Is this correct for this job?", vbYesNo) 
                    If warning_check = vbYes Then 
                        EMWriteScreen "Y", 9, 58
                        transmit 
                    ElseIf warning_check = vbNo Then 
                        EMWriteScreen "N", 9, 58
                        transmit
                    End If
                    CALL write_value_and_transmit("BGTX", 20, 71)   'Now we send the case through background and go to STAT/WRAP
                End If
            Else 
                Exit do
            End If
        Loop until wrap_check = "WRAP"

        If MAXIS_footer_month = CM_plus_1_mo AND MAXIS_footer_year = CM_plus_1_yr Then 
            EMWriteScreen "N", 16, 54        'If we are already in CM+1 then we enter N because we can't update the next month
        Else 
            EMWriteScreen "Y", 16, 54           'This goes into the next footer month without leaving STAT
        End If
        transmit
        If testing_status = TRUE Then msgbox "Pause here - should be in the next month."
        second_loop = TRUE                  'setting this veriable to knkow we aren't in the first go round any more
    Loop until MAXIS_footer_month = CM_plus_1_mo AND MAXIS_footer_year = CM_plus_1_yr
Else            'If we are in developer mode, we will go here to allow for some display and output options.
    Do
        Do
            err_msg = ""

            'This dialog allows workers to send an email with detail gathered to up to five recipients.
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 216, 180, "Send Job Change Information as Email"
                CheckBox 10, 40, 245, 10, "Check here to email the job change information", send_email_checkbox
                EditBox 25, 55, 115, 15, email_address_one
                EditBox 25, 75, 115, 15, email_address_two
                EditBox 25, 95, 115, 15, email_address_three
                EditBox 25, 115, 115, 15, email_address_four
                EditBox 25, 135, 115, 15, email_address_five
                ButtonGroup ButtonPressed
                    OkButton 105, 160, 50, 15
                    CancelButton 160, 160, 50, 15
                Text 10, 10, 180, 20, "To email job change information collected by the script, select checkbox and enter email(s) below."
                Text 145, 60, 50, 10, "@hennepin.us"
                Text 145, 80, 50, 10, "@hennepin.us"
                Text 145, 100, 50, 10, "@hennepin.us"
                Text 145, 120, 50, 10, "@hennepin.us"
                Text 145, 140, 50, 10, "@hennepin.us"
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
        email_body = email_body & "Job Name: " & job_employer_name & " - Income Type: " & job_income_type & " - Employee: Memb " & hh_memb_with_job_change & vbCr
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
                email_body = email_body & " - Voluntary quit reason: " & vol_quit_reason & vbCr
                If trim(stwk_reason) <> "" Then email_body = email_body & " - Reason for STWK: " & stwk_reason & vbCr
                email_body = email_body & " - Meets good cause? " & good_cause_yn & vbCr
                If disq_updated = TRUE then 
                     email_body = email_body & " - DISQ panel updated for " & disq_begin_date & "-" & disq_end_date & vbCr
                End If
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
        Call create_outlook_email("", all_email_recipients, "", "", "Job Change Reported for MX Case", 1, False, "", "", False, "", email_body, False, "", True)

    End If
End If

If refused_empl_yn = "?" Then refused_empl_yn = "N/A"
If good_cause_yn = "?" Then good_cause_yn = "N/A"

If job_change_type = "New Job Reported" Then verif_type_requested = "new job"                                       'setting a variale for entry into headers/notes/TIKL
If job_change_type = "Income/Hours Change for Current Job" Then verif_type_requested = "change in current job"
If job_change_type = "Job Ended" Then verif_type_requested = "job ended"
'This sets a TIKL if requested and NOT in developer mode
If TIKL_checkbox = checked and developer_mode = FALSE Then Call create_TIKL("Verification of " & verif_type_requested & " due.", 10, verif_form_date, True, TIKL_note_text)

If IsDate(verif_form_date) = TRUE and developer_mode = FALSE Then
    'Send a SPEC/MEMO to help support the verification needed from the client.
    'THIS DOES NOT REPLACE THE VERIFICATION REQUEST FORM
    Call back_to_SELF
    Call MAXIS_background_check
    CALL start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, True)

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
	CALL digital_experience
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
      Text 10, y_pos, 530, 10, "* Employee: Memb " & hh_memb_with_job_change
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
              Text 10, y_pos, 530, 10, " - Voluntary quit reason " & vol_quit_reason
              y_pos = y_pos + 10
              Text 10, y_pos, 530, 10, " - Reason for STWK: " & stwk_reason
              y_pos = y_pos + 10
              Text 10, y_pos, 530, 10, "    - Meets good cause? " & good_cause_yn
              y_pos = y_pos + 10
              Text 10, y_pos, 530, 10, " - DISQ panel updated " & DISQ_updated
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
        objSelection.TypeText "* Employee: Memb " & hh_memb_with_job_change & vbCr
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
                objSelection.TypeText " - Voluntary quit reason: " & vol_quit_reason & vbCr
                If trim(stwk_reason) <> "" Then objSelection.TypeText " - Reason for STWK: " & stwk_reason & vbCr
                objSelection.TypeText " - Meets good cause? " & good_cause_yn & vbCr
                If disq_updated = TRUE Then objSelection.TypeText " - DISQ panel updated for " & disq_begin_date & "-" & disq_end_date & vbCr
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
    Call write_variable_in_CASE_NOTE("* Employee: Memb " & hh_memb_with_job_change)
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
            Call write_variable_with_indent_in_CASE_NOTE("Voluntary quit reason: " & vol_quit_reason)
            If trim(stwk_reason) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Reason for STWK: " & stwk_reason)
            Call write_variable_with_indent_in_CASE_NOTE("Meets good cause? " & good_cause_yn)
            If disq_updated = TRUE Then Call write_variable_with_indent_in_CASE_NOTE("DISQ panel updated for: " & disq_begin_date & "-" & disq_end_date)
            Call write_variable_with_indent_in_CASE_NOTE("Sanction Occurance: " & snap_vol_quit_occurance)
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