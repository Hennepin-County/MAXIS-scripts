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



testing_status = True		'Testing Status: Update status to true to display all testing msgbox or false to hide all testing msgbox

'Initial dialog to gather case details
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 176, 65, "Case Number Dialog"
  ButtonGroup ButtonPressed
    OkButton 75, 45, 45, 15
    CancelButton 125, 45, 45, 15
  EditBox 75, 5, 45, 15, MAXIS_case_number
  EditBox 75, 25, 95, 15, worker_signature
  Text 20, 10, 50, 10, "Case Number:"
  Text 10, 30, 60, 10, "Worker Signature:"
EndDialog

'Runs the first dialog - which confirms the case number
Do
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
      	Call validate_MAXIS_case_number(err_msg, "*")
        If trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Sign your case note."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call check_for_MAXIS(False)
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "ADDR", is_this_priv)
If is_this_priv = True then script_end_procedure("This case is privileged and cannot be accessed. The script will now stop.")

EmReadscreen county_code, 4, 21, 21
If county_code <> worker_county_code then script_end_procedure("This case is out-of-county, and cannot access CASE/NOTE. The script will now stop.")

If job_change_type = "New Job Reported" Then change_text = "new job"                                        'creating text for easy view in the next dialog - this is not functional - just formatting.
If job_change_type = "Income/Hours Change for Current Job" Then change_text = "change job income"
If job_change_type = "Job Ended" Then change_text = "stwk"

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

call generate_client_list(list_of_hh_membs_with_jobs, "New Job Reported")
client_name_array = split(list_of_hh_membs_with_jobs, chr(9))
call generate_client_list(list_of_all_hh_members, "Select or Type")
Call Generate_Client_List(HH_Memb_DropDown, "Select One:")


hh_memb_and_current_jobs =  "*"
hh_memb_5_jobs_panels = "*"

'If the client indicated the JOB panel already exists, the script will try to determine which JOBS panel should be used.
For i = 1 to ubound(client_name_array) 	'looping through all the reference numbers so that we can check all the members for JOBS panels.

	Call Navigate_to_MAXIS_screen("STAT", "JOBS")                       'Go to JOBS
	EMWriteScreen left(client_name_array(i), 2), 20, 76                                    'Enter the reference for the clients in turn to check each.
	EMWriteScreen "01", 20, 79
	transmit

	EMReadScreen total_jobs, 1, 2, 78                                   'look for how many JOBS panels there are so we can loop through them all
	If total_jobs <> "0" Then       'if there are no JOBS panels listed for this member, we should't try to read them.
		Do
			EMReadScreen employer_name, 30, 7, 42                'reading the employer name
			employer_name = replace(employer_name, "_", "")           'taking out the underscores
			EMReadScreen jobs_panel_number, 1, 2, 73
			jobs_panel_number = "0" & jobs_panel_number 

			If jobs_panel_number = "05" then hh_memb_5_jobs_panels = hh_memb_5_jobs_panels & client_name_array(i) & "*"
			msgbox "hh_memb_5_jobs_panels " & hh_memb_5_jobs_panels

			hh_memb_and_current_jobs = hh_memb_and_current_jobs & client_name_array(i) &  "-" & employer_name & "-" & jobs_panel_number & "*"

			transmit
			EMReadScreen last_job, 7, 24, 2
		Loop until last_job = "ENTER A"
	End If
Next

' hh_memb_and_current_jobs = split(hh_memb_and_current_jobs, "*")
' call generate_client_list(hh_memb_and_current_jobs, "Type or Select")
hh_memb_and_current_jobs = replace(hh_memb_and_current_jobs, "*", chr(9))
hh_memb_and_current_jobs = left(hh_memb_and_current_jobs, len(hh_memb_and_current_jobs) - 1)
If testing_status = True Then msgbox hh_memb_and_current_jobs

BeginDialog Dialog1, 0, 0, 386, 125, "Job Change Selection - Case: " & MAXIS_case_number
  DropListBox 125, 5, 255, 15, "Select One ..."+chr(9) + "No JOBS Panel Exists" + hh_memb_and_current_jobs, hh_member_current_jobs
  DropListBox 245, 20, 135, 15, HH_Memb_DropDown, hh_memb_with_new_job
  DropListBox 75, 45, 140, 15, "Select One ..."+chr(9)+"New Job Reported"+chr(9)+"Income/Hours Change for Current Job"+chr(9)+"Job Ended", job_change_type
  ComboBox 290, 45, 90, 15, list_of_all_hh_members, person_who_reported_job
  ComboBox 105, 65, 110, 15, "Type or Select"+chr(9)+"phone call"+chr(9)+"Change Report Form"+chr(9)+"office visit"+chr(9)+"mailing"+chr(9)+"fax"+chr(9)+"ES counselor"+chr(9)+"CCA worker"+chr(9)+"scanned document", job_report_type
  EditBox 290, 65, 55, 15, reported_date
  CheckBox 15, 90, 305, 10, "Check here if the employee gave verbal authorization to check the Work Number", work_number_verbal_checkbox
  ButtonGroup ButtonPressed
    OkButton 270, 105, 50, 15
    CancelButton 330, 105, 50, 15
  Text 235, 70, 55, 10, "Date reported:"
  Text 85, 25, 155, 10, "If no corresponding JOBS panel exists, select Member for new JOBS panel:"
  Text 15, 10, 110, 10, "Member(s)-Job(s) listed on case: "
  Text 15, 50, 60, 10, "Job Change Type:"
  Text 235, 50, 55, 10, "Who reported?"
  Text 15, 70, 90, 10, "How was the job reported?"
EndDialog

Do
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
		' If (hh_member_current_jobs = "New Job Reported" AND (hh_memb_with_new_job = "Select One:" OR job_change_type <> "New Job Reported")) OR (job_change_type = "New Job Reported" AND (hh_member_current_jobs <> "New Job Reported" OR hh_memb_with_new_job = "Select One:")) Then err_msg = err_msg & vbNewLine & "* To add a new job: Member-Jobs list and Job Change Type must both say New Job Reported and Member of New Job must be selected"
		If hh_memb_with_new_job <>  "Select One:" AND (hh_member_current_jobs <> "New Job Reported" AND job_change_type <> "New Job Reported") Then err_msg = err_msg & vbNewLine & "* Cannot select Member of New Job if New Job Reported is not selected for Member-Jobs list and Job Change Type"
		If job_change_type = "Select One ..." Then err_msg = err_msg & vbNewLine & "* Select valid Job Change Type"
		If person_who_reported_job = "Select or Type" Then err_msg = err_msg & vbNewLine & "* Select who reported the job change"
		If job_report_type = "Type or Select" Then err_msg = err_msg & vbNewLine & "* Select how the job was reported"
		If IsDate(reported_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter valid date reported." 
		If instr(hh_memb_5_jobs_panels, hh_memb_with_new_job) then Call script_end_procedure("Script is unable to add another JOBS panel for " & hh_memb_with_new_job & " as there are 5 JOBS panels already. Please delete a JOBS panel for the Household Member and then restart the script.")
        If err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = "" 
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in


call generate_client_list(list_of_members, "Type or Select")
client_name_array = split(list_of_members, chr(9))

'We are going to look at each HH member checked in the HH_member dialog if the job already exists
If hh_member_current_jobs <> "No JOBS Panel Exists" Then
	Call navigate_to_MAXIS_screen("STAT", "JOBS")       'Starting with JOBS panels
	EmWriteScreen left(hh_member_current_jobs, 2), 20, 76
	call write_value_and_transmit(right(hh_member_current_jobs, 2), 20, 79)

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
	' For each person in client_name_array                        'looping through the people in the client array that also has reference numbers listed in it to make the employer name have the reference number AND full name from STAT
	'     If left(person, 2) = job_ref_number Then job_employee =  person
	' Next
	If job_income_type = "J" Then job_income_type = "J - WIOA"              'formatting the income type to have the code and PF1 information
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
	If job_update_date = date Then script_update_stat = "No - Update of JOBS not needed"
End If

'This function will read what programs are on the case and what the case status is.
Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
EMReadScreen case_appl_date, 8, 8, 29											'reading the date the case was APPLd so we know how far back we can go for new jobs start date
case_appl_date = DateAdd("d", 0, case_appl_date)

'To do - if select job ended, then evaluate this
For i = 1 to Ubound(client_name_array)                        'Now we are going to check WREG for the PWE

    Call Navigate_to_MAXIS_screen("STAT", "WREG")           'Go to WREG for each client on the case in turn
    EMWriteScreen left(client_name_array(i), 2), 20, 76
    transmit

	msgbox "Should be at left(client_name_array(i), 2) " & left(client_name_array(i), 2)

    EMReadScreen pwe_code, 1, 6, 68                         'Read PWE code
    If pwe_code = "Y" Then                                  'If it is 'Y' - save the reference number to identify the PWE
        pwe_ref = left(client_name_array(i), 2)
        Exit For
    End If
Next

msgbox "did it work?"


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

'To do - needs to be validated

Call generate_client_list(list_of_employees, "Select One ...")
Call generate_client_list(list_of_members, "Select One ...")

'This didalog has a different middle part based on the type of report that is happening
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 520, 335, "Job Change Details Dialog"
  GroupBox 5, 5, 250, 220, "Job Information"
  Text 10, 20, 85, 10, "Employee (HH Member)*:"
  Text 10, 40, 60, 10, "Work Number ID:"
  If job_ref_number <> "" Then Text 475, 65, 50, 10, "JOBS " & job_ref_number & " "  & job_instance
  Text 10, 60, 45, 10, "Income Type:"
  Text 10, 80, 85, 10, "Subsidized Income Type:"
  Text 10, 100, 40, 10, "Employer*:"
  Text 10, 120, 45, 10, "Verification*:"
  Text 10, 140, 55, 10, "Pay Frequency:"
  Text 10, 160, 60, 10, "Income start date:"
  Text 10, 180, 60, 10, "Income End Date:"
  Text 10, 200, 75, 10, "Contract through date:"
  DropListBox 95, 15, 140, 45, list_of_employees, job_employee
  EditBox 95, 35, 50, 15, work_numb_id
  DropListBox 95, 55, 140, 15, "W - Wages"+chr(9)+"J - WIOA"+chr(9)+"E - EITC"+chr(9)+"G - Experience Works"+chr(9)+"F - Federal Work Study"+chr(9)+"S - State Work Study"+chr(9)+"O - Other"+chr(9)+"C - Contract Income"+chr(9)+"T - Training Program"+chr(9)+"P - Service Program"+chr(9)+"R - Rehab Program", job_income_type
  DropListBox 95, 75, 140, 15, "01 - Subsidized Public Sector Employer"+chr(9)+"02 - Subsidized Private Sector Employer"+chr(9)+"03 - On-the-Job-Training"+chr(9)+"04 - AmeriCorps", job_subsidized_income_type
  EditBox 95, 95, 140, 15, job_employer_name
  DropListBox 95, 115, 95, 45, "N - No Verif Provided"+chr(9)+"? - Delayed Verification", job_verification
  DropListBox 95, 135, 95, 45, "1 - Monthly"+chr(9)+"2 - Semi-Monthly"+chr(9)+"3 - Biweekly"+chr(9)+"4 - Weekly"+chr(9)+"5 - Other", job_pay_frequency
  EditBox 95, 155, 55, 15, job_income_start
  EditBox 95, 175, 55, 15, job_income_end
  EditBox 95, 195, 55, 15, job_contract_through_date

  Select Case job_change_type
      Case "New Job Reported"
        GroupBox 260, 5, 250, 220, "Update Reported - NEW JOB"
        Text 270, 20, 65, 10, "Date work started:"
        Text 270, 40, 105, 10, "Date income started/will start*:"
        Text 270, 60, 95, 10, "Initial check GROSS amount:"
        Text 270, 80, 90, 10, "Anticipated hours per week:"
        Text 270, 100, 80, 10, "Anticipated hourly wage:"
        Text 270, 120, 65, 10, "Conversation with:"
        Text 270, 135, 25, 10, "Details:"
        Text 270, 175, 95, 10, "* Impact on WREG/ABAWD:"
        EditBox 375, 15, 55, 15, date_work_started
        EditBox 375, 35, 55, 15, new_job_income_start
        EditBox 375, 55, 55, 15, initial_check_gross_amount
        EditBox 375, 75, 55, 15, new_job_hours_per_week
        EditBox 375, 95, 55, 15, new_job_hourly_wage
        ComboBox 375, 120, 130, 45, list_of_members, conversation_with_person
        EditBox 375, 135, 130, 15, conversation_detail
        CheckBox 270, 155, 145, 10, "Check here if Work Number request sent", work_number_checkbox
        EditBox 375, 170, 125, 15, wreg_abawd_notes
      Case "Income/Hours Change for Current Job"
        GroupBox 260, 5, 250, 220, "Update Reported - JOB CHANGE"
        Text 265, 20, 55, 10, "Date of Change:"
        Text 265, 40, 70, 10, "* Change Reported:"
        Text 265, 55, 90, 10, "-- Old Anticipated Income --"
        Text 265, 70, 60, 10, "Hours per Week:"
        Text 395, 70, 50, 10, "Hourly Wage:"
        Text 265, 85, 55, 10, "Income Change:"
        Text 265, 105, 95, 10, "-- New Anticipated Income* --"
        Text 265, 120, 60, 10, "Hours per Week:"
        Text 400, 120, 50, 10, "Hourly Wage:"
        Text 265, 140, 85, 10, "First Pay Date Impacted*:"
        Text 265, 155, 65, 10, "Conversation with:"
        Text 265, 175, 25, 10, "Details:"
        Text 265, 210, 95, 10, "Impact on WREG/ABAWD*:"
        EditBox 340, 15, 55, 15, job_change_date
        EditBox 340, 35, 165, 15, job_change_details
        EditBox 340, 65, 50, 15, job_change_old_hours_per_week
        EditBox 455, 65, 50, 15, job_change_old_hourly_wage
        ComboBox 340, 85, 165, 45, "Select or Type"+chr(9)+"Increase"+chr(9)+"Decrease", income_change_type
        EditBox 355, 115, 35, 15, job_change_new_hours_per_week
        EditBox 455, 115, 50, 15, job_change_new_hourly_wage
        EditBox 355, 135, 60, 15, first_pay_date_of_change
        ComboBox 355, 155, 150, 45, list_of_members, conversation_with_person
        EditBox 355, 170, 150, 15, conversation_detail
        CheckBox 265, 190, 190, 10, "Check here if you sent a Work Number request.", work_number_checkbox
        EditBox 365, 205, 140, 15, wreg_abawd_notes      
      Case "Job Ended"
        GroupBox 260, 5, 250, 220, "Update Reported - JOB ENDED"
        Text 270, 20, 70, 10, "Date Work Ended*:"
        Text 270, 40, 105, 10, "Date income ended/will end*:"
        Text 400, 20, 65, 10, "Last pay amount*:"
        Text 270, 60, 60, 10, "Reason for STWK:"
        Text 270, 90, 55, 10, "Voluntary Quit*:"
        Text 395, 90, 60, 10, "Good cause met?"
        Text 270, 145, 65, 10, "Conversation with:"
        Text 270, 165, 25, 10, "Details:"
        Text 270, 125, 160, 10, "Is the client applying for Unemployment Income?"
        Text 270, 185, 95, 10, "Impact on WREG/ABAWD*:"
        Text 270, 75, 70, 10, "STWK Verification*:"
        Text 270, 110, 50, 10, "Refused Empl:"
        Text 380, 110, 65, 10, "Refused Empl Date:"
        EditBox 340, 15, 55, 15, date_work_ended
        EditBox 375, 35, 50, 15, job_end_income_end_date
        EditBox 465, 15, 40, 15, last_pay_amount
        EditBox 340, 55, 130, 15, stwk_reason
        DropListBox 340, 90, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", vol_quit_yn
        DropListBox 460, 90, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", good_cause_yn
        ComboBox 340, 145, 165, 45, list_of_members, conversation_with_person
        EditBox 340, 160, 165, 15, conversation_detail
        DropListBox 440, 125, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", uc_yn
        EditBox 365, 180, 140, 15, wreg_abawd_notes
        CheckBox 270, 200, 190, 10, "Check here if you sent a Work Number request.", work_number_checkbox
        DropListBox 340, 75, 125, 45, "Select One..."+chr(9)+"N - No Verif Provided"+chr(9)+"? - Delayed Verification", stwk_verif
        DropListBox 340, 105, 30, 45, "?"+chr(9)+"Yes"+chr(9)+"No", refused_empl_yn
        EditBox 460, 105, 45, 15, refused_empl_date      
        ' ComboBox 80, 260, 125, 45, "Select One..."+chr(9)+"1 - Employers Statement"+chr(9)+"2 - Seperation Notice"+chr(9)+"3 - Collateral Statement"+chr(9)+"4 - Other Document"+chr(9)+"N - No Verif Provided", stwk_verif
  End Select

  GroupBox 5, 230, 505, 75, "Actions"
  Text 10, 245, 105, 10, "Date verification Request Sent:"
  Text 10, 265, 105, 10, "Time frame of verifs requested:"
  Text 10, 285, 90, 10, "Have Script Update Panel:"
  Text 10, 315, 25, 10, "Notes:"
  EditBox 120, 240, 75, 15, verif_form_date
  EditBox 120, 260, 75, 15, verif_time_frame
  CheckBox 270, 240, 105, 10, "Check here to TIKL for return.", TIKL_checkbox
  CheckBox 270, 285, 165, 10, "Check here if you are requesting CEI/OHI docs.", requested_CEI_OHI_docs_checkbox
  DropListBox 120, 285, 135, 45, "Select One..."+chr(9)+"No - Update of JOBS not needed"+chr(9)+"Yes - Update an existing JOBS Panel"+chr(9)+"Yes - Create a new JOBS Panel", script_update_stat
  CheckBox 270, 270, 165, 10, "Check here if you sent a status update to CCA.", CCA_checkbox
  CheckBox 270, 255, 160, 10, "Check here if you sent a status update to ES.", ES_checkbox
  EditBox 40, 310, 340, 15, notes
  ButtonGroup ButtonPressed
    OkButton 405, 310, 50, 15
    CancelButton 460, 310, 50, 15
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

msgbox "script will now show values"