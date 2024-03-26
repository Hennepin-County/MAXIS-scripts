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
call changelog_update("07/21/2023", "Updated function that sends an email through Outlook", "Mark Riegel, Hennepin County")
call changelog_update("01/12/2023", "BUG FIX to handle for New Jobs that started prior to the initial application for the case. The script would get stuck trying to go too far in the past. The script will now initially try to update only starting in the application month.", "Casey Love, Hennepin County")
call changelog_update("05/28/2020", "Added virtual drop box information to SPEC/MEMO.", "MiKayla Handley, Hennepin County")
call changelog_update("04/24/2020", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
'connecting to MAXIS & grabbing the case number
EMConnect ""
call MAXIS_case_number_finder(MAXIS_case_number)

get_county_code
Call check_for_MAXIS(False)
CALL MAXIS_case_number_finder(MAXIS_case_number)

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

hh_memb_and_current_jobs =  "*"

'If the client indicated the JOB panel already exists, the script will try to determine which JOBS panel should be used.
For i = 1 to ubound(client_name_array) 	'looping through all the reference numbers so that we can check all the members for JOBS panels.

	Call Navigate_to_MAXIS_screen("STAT", "JOBS")                       'Go to JOBS
	EMWriteScreen left(client_name_array(i), 2), 20, 76                                    'Enter the reference for the clients in turn to check each.
	EMWriteScreen "01", 20, 79
	transmit

	EMReadScreen total_jobs, 1, 2, 78                                   'look for how many JOBS panels there are so we can loop through them all
	If total_jobs <> "0" Then       'if there are no JOBS panels listed for this member, we should't try to read them.
		Do
			EMReadScreen employer, 30, 7, 42                'reading the employer name
			employer = replace(employer, "_", "")           'taking out the underscores

			hh_memb_and_current_jobs = hh_memb_and_current_jobs & client_name_array(i) &  "-" & employer & "*"

			transmit
			EMReadScreen last_job, 7, 24, 2
		Loop until last_job = "ENTER A"
	End If
Next

' hh_memb_and_current_jobs = split(hh_memb_and_current_jobs, "*")
' call generate_client_list(hh_memb_and_current_jobs, "Type or Select")
hh_memb_and_current_jobs = replace(hh_memb_and_current_jobs, "*", chr(9))
hh_memb_and_current_jobs = left(hh_memb_and_current_jobs, len(hh_memb_and_current_jobs) - 1)
msgbox hh_memb_and_current_jobs



BeginDialog Dialog1, 0, 0, 386, 210, "Job Change Selection"
  DropListBox 110, 20, 255, 15, "New Job Reported" + hh_memb_and_current_jobs, hh_memb_and_current_jobs
  DropListBox 180, 45, 150, 15, list_of_all_hh_members, hh_memb_new_job
  DropListBox 85, 70, 140, 15, "Select One ..."+chr(9)+"New Job Reported"+chr(9)+"Income/Hours Change for Current Job"+chr(9)+"Job Ended", job_change_type
  ComboBox 110, 90, 115, 15, "Type or Select"+chr(9)+"phone call"+chr(9)+"Change Report Form"+chr(9)+"office visit"+chr(9)+"mailing"+chr(9)+"fax"+chr(9)+"ES counselor"+chr(9)+"CCA worker"+chr(9)+"scanned document", job_report_type
  ComboBox 85, 115, 140, 15, list_of_all_hh_members, person_who_reported_job
  EditBox 80, 135, 55, 15, reported_date
  CheckBox 5, 165, 305, 10, "Check here if the employee gave verbal authorization to check the Work Number", work_number_verbal_checkbox
  ButtonGroup ButtonPressed
    OkButton 265, 190, 50, 15
    CancelButton 325, 190, 50, 15
  Text 15, 140, 55, 10, "Date reported?"
  Text 5, 5, 45, 10, "Case number:"
  Text 45, 45, 130, 20, "If memb/job combination does not exist, select HH memb to create new job:"
  Text 55, 5, 55, 10, MAXIS_case_number
  Text 15, 25, 95, 10, "Member/Jobs listed on case: "
  Text 15, 75, 60, 10, "Job Change Type:"
  Text 15, 120, 55, 10, "Who reported?"
  Text 15, 95, 90, 10, "How was the job reported?"
EndDialog

Do
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
      	' Call validate_MAXIS_case_number(err_msg, "*")
        ' If trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Sign your case note."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

call generate_client_list(list_of_members, "Type or Select")
client_name_array = split(list_of_members, chr(9))

' call HH_member_custom_dialog(HH_member_array)   'finding who should be looked at for income on the case
'FUTURE FUNCTIONALITY - Stop work should be added in before we add information to the EARNED_INCOME_PANELS_ARRAY

Call navigate_to_MAXIS_screen("STAT", "JOBS")       'Starting with JOBS panels
For i = 1 to ubound(client_name_array)                  'We are going to look at each HH member checked in the HH_member dialog
	EMWriteScreen left(client_name_array(i), 2) , 20, 76                    'going to the member in JOBS
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
				' If frequency = "1" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "1 - One Time Per Month"      'setting full detail to the array instead of a single letter code
				' If frequency = "2" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "2 - Two Times Per Month"
				' If frequency = "3" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "3 - Every Other Week"
				' If frequency = "4" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "4 - Every Week"
				' If frequency = "5" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "5 - Other"

				EARNED_INCOME_PANELS_ARRAY(income_list_indct, the_panel) = "NONE"       'This is where all of the array items from LIST_OF_INCOME_ARRAY will be added that are associated with this panel
				EARNED_INCOME_PANELS_ARRAY(this_is_a_new_panel, the_panel) = FALSE      'identifies if a panel was created by the script or not - these are currently existing - changes CNote

				the_panel = the_panel + 1       'incrementing our counter to be ready for the next panel/member/income type
			End If      'If save_this_panel = TRUE Then
		Next            'For panel = 1 to number_of_jobs_panels
	End If              'If number_of_jobs_panels <> "0" Then
Next                    'For each member in HH_member_array
