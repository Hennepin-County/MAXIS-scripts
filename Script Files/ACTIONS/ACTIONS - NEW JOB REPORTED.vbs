'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - NEW JOB REPORTED.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 345                	'manual run time in seconds
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

'THIS SCRIPT IS BEING USED IN A WORKFLOW SO DIALOGS ARE NOT NAMED
'DIALOGS MAY NOT BE DEFINED AT THE BEGINNING OF THE SCRIPT BUT WITHIN THE SCRIPT FILE

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year =  CM_plus_1_yr

'THE SCRIPT----------------------------------------------------------------------------------------------------
'connecting to MAXIS & grabbing the case number
EMConnect ""
call MAXIS_case_number_finder(MAXIS_case_number)

'Shows and defines the case number dialog
BeginDialog , 0, 0, 161, 65, "Case number and footer month"
  Text 5, 10, 85, 10, "Enter your case number:"
  EditBox 95, 5, 60, 15, MAXIS_case_number
  Text 15, 30, 50, 10, "Footer month:"
  EditBox 65, 25, 25, 15, MAXIS_footer_month
  Text 95, 30, 20, 10, "Year:"
  EditBox 120, 25, 25, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 25, 45, 50, 15
    CancelButton 85, 45, 50, 15
EndDialog

Do 
	Dialog 					'Calling a dialog without a assigned variable will call the most recently defined dialog
	cancel_confirmation
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'Checks footer month and year. If footer month and year do not match the worker entry, it'll back out and get there manually.
Call MAXIS_footer_month_confirmation

'NAV to stat/jobs
call navigate_to_MAXIS_screen("stat", "jobs")

'Declaring some variables to create defaults for the new_job_reported_dialog.
create_JOBS_checkbox = 1
HH_memb = "01"
HH_memb_row = 5 'This helps the navigation buttons work!

'Shows and defines the main dialog.
BeginDialog , 0, 0, 291, 300, "New job reported dialog"
  EditBox 80, 5, 25, 15, HH_memb
  DropListBox 55, 25, 110, 15, "W Wages (Incl Tips)"+chr(9)+"J WIOA"+chr(9)+"E EITC"+chr(9)+"G Experience Works"+chr(9)+"F Federal Work Study"+chr(9)+"S State Work Study"+chr(9)+"O Other"+chr(9)+"I Infrequent < 30 N/Recur"+chr(9)+"M Infreq <= 10 MSA Exclusion"+chr(9)+"C Contract Income"+chr(9)+"T Training Program"+chr(9)+"P Service Program"+chr(9)+"R Rehab Program", income_type_dropdown
  DropListBox 135, 45, 150, 15, "not applicable"+chr(9)+"01 Subsidized Public Sector Employer"+chr(9)+"02 Subsidized Private Sector Employer"+chr(9)+"03 On-the-Job-Training"+chr(9)+"04 AmeriCorps (VISTA/State/National/NCCC)", subsidized_income_type_dropdown
  EditBox 45, 65, 240, 15, employer
  EditBox 125, 85, 55, 15, income_start_date
  EditBox 125, 105, 55, 15, contract_through_date
  EditBox 100, 125, 100, 15, who_reported_job
  ComboBox 100, 145, 100, 15, "phone call"+chr(9)+"office visit"+chr(9)+"mailing"+chr(9)+"fax"+chr(9)+"ES counselor"+chr(9)+"CCA worker"+chr(9)+"scanned document", job_report_type
  EditBox 30, 165, 255, 15, notes
  CheckBox 5, 185, 190, 10, "Check here to have the script make a new JOBS panel.", create_JOBS_checkbox
  CheckBox 5, 200, 190, 10, "Check here if you sent a status update to CCA.", CCA_checkbox
  CheckBox 5, 215, 190, 10, "Check here if you sent a status update to ES.", ES_checkbox
  CheckBox 5, 230, 190, 10, "Check here if you sent a Work Number request.", work_number_checkbox
  CheckBox 5, 245, 165, 10, "Check here if you are requesting CEI/OHI docs.", requested_CEI_OHI_docs_checkbox
  CheckBox 5, 260, 235, 10, "Check here to have the script send a TIKL to return proofs in 10 days.", TIKL_checkbox 
  EditBox 65, 275, 110, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 180, 275, 50, 15
    CancelButton 235, 275, 50, 15
    PushButton 175, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 175, 25, 45, 10, "next panel", next_panel_button
    PushButton 235, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 235, 25, 45, 10, "next memb", next_memb_button
  GroupBox 170, 5, 115, 35, "STAT-based navigation"
  Text 5, 30, 45, 10, "Income Type: "
  Text 5, 50, 130, 10, "Subsidized Income Type (if applicable):"
  Text 5, 70, 40, 10, "Employer:"
  Text 30, 90, 95, 10, "Income start date (if known):"
  Text 5, 110, 120, 10, "Contract through date (if applicable):"
  Text 5, 130, 80, 10, "Who reported the job?:"
  Text 5, 150, 90, 10, "How was the job reported?:"
  Text 5, 170, 25, 10, "Notes:"
  Text 5, 280, 60, 10, "Worker signature:"
  Text 5, 10, 70, 10, "HH member number:"
EndDialog

'Defaulting the "set TIKL" variable to checked
TIKL_checkbox = checked

DO
	Do
		Do
			Do
				Do
					Do
						Do
							Dialog 					'Calling a dialog without a assigned variable will call the most recently defined dialog
							cancel_confirmation
							MAXIS_dialog_navigation
							If isdate(income_start_date) = True then		'Logic to determine if the income start date is functional
								If (datediff("m", MAXIS_footer_month & "/01/20" & MAXIS_footer_year, income_start_date) > 0) then
									MsgBox "Your income start date is after your footer month. If the income start date is after this month, exit the script and try again in the correct footer month."
									pass_through_inc_date_loop = False
								Else
									pass_through_inc_date_loop = True
								End if
							Else
								If income_start_date <> "" then MsgBox "You must type a date in the Income Start Date field, or leave it blank."
							End if
						Loop until income_start_date = "" or pass_through_inc_date_loop = True
						If employer = "" then MsgBox "You must type an employer!"
					Loop until employer <> ""
					If isdate(contract_through_date) = True or income_type_dropdown = "C Contract Income" then
						If income_type_dropdown <> "C Contract Income" then
							MsgBox "You should not put a ''contract through'' date in, unless the income type is ''C Contract Income''."
							pass_through_contract_date_loop = False
						Elseif income_type_dropdown = "C Contract Income" and isdate(contract_through_date) = False then
							MsgBox "You should not put a ''C Contract Income'' code in, unless there is a ''contract through'' date."
							pass_through_contract_date_loop = False
						Else
							pass_through_contract_date_loop = True
						End if
					Else
						If contract_through_date <> "" then MsgBox "You must type a date in the Contract Through date field, or leave it blank."
					End if
				Loop until (contract_through_date = "" and income_type_dropdown <> "C Contract Income") or pass_through_contract_date_loop = True
				If who_reported_job = "" then MsgBox "You must type out who reported the job!"
			Loop until who_reported_job <> ""
			If job_report_type = "" then MsgBox "You must select how you heard about the job, or write something in that field yourself."
		Loop until job_report_type <> ""
		If worker_signature = "" then MsgBox "You must sign your case note!"
	Loop until worker_signature <> ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'Creates a new JOBS panel if that was selected.
If create_JOBS_checkbox = checked then
	EMWriteScreen HH_memb, 20, 76
	EMWriteScreen "nn", 20, 79
	transmit
	EMReadScreen edit_mode_check, 1, 20, 8
	If edit_mode_check = "D" then script_end_procedure("Unable to create a new JOBS panel. Check which member number you provided. Otherwise you may be in inquiry mode. If so shut down inquiry and try again. Or try closing BlueZone.")
	IF ((MAXIS_footer_month * 1) >= 10 AND (MAXIS_footer_year * 1) >= "16") OR (MAXIS_footer_year = "17") THEN  'handling for changes to jobs panel for bene month 10/16
		EMWriteScreen left(income_type_dropdown, 1), 5, 34
		If subsidized_income_type_dropdown <> "not applicable" then EMWriteScreen left(subsidized_income_type_dropdown, 2), 5, 74
		EMWriteScreen "n", 6, 34
	ELSE
		EMWriteScreen left(income_type_dropdown, 1), 5, 38
		IF left(income_type_dropdown, 1) = "J" OR left(income_type_dropdown, 1) = "G" OR left(income_type_dropdown, 1) = "T" OR left(income_type_dropdown, 1) = "P" OR left(income_type_dropdown, 1) = "R" THEN EMWriteScreen "0", 6, 75 'Adding in a temporary hourly wage for special income types which require it.
		If subsidized_income_type_dropdown <> "not applicable" then EMWriteScreen left(subsidized_income_type_dropdown, 2), 5, 71
		EMWriteScreen "n", 6, 38
	END IF
	EMWriteScreen employer, 7, 42
	If income_start_date <> "" then call create_MAXIS_friendly_date(income_start_date, 0, 9, 35)
	If contract_through_date <> "" then call create_MAXIS_friendly_date(contract_through_date, 0, 9, 73)
	EMReadScreen MAXIS_footer_month, 2, 20, 55
	EMReadScreen MAXIS_footer_year, 2, 20, 58
	If isdate(income_start_date) = True then
		If datediff("d", income_start_date, MAXIS_footer_month & "/01/20" & MAXIS_footer_year) > 0 then
			call create_MAXIS_friendly_date(MAXIS_footer_month & "/01/20" & MAXIS_footer_year, 0, 12, 54)
		Else
			call create_MAXIS_friendly_date(income_start_date, 0, 12, 54)
		End if
	Else
		call create_MAXIS_friendly_date(MAXIS_footer_month & "/01/20" & MAXIS_footer_year, 0, 12, 54)
	End if
	EMWriteScreen "0", 12, 67
	EMWriteScreen "0", 18, 72
	Do
		transmit
		EMReadScreen edit_mode_check, 1, 20, 8
	Loop until edit_mode_check = "D"
End if

If TIKL_checkbox = 1 then 
	call navigate_to_MAXIS_screen("dail", "writ")
	'The following will generate a TIKL formatted date for 10 days from now.
	call create_MAXIS_friendly_date(date, 10, 5, 18)
	'Writing in the rest of the TIKL.
	call write_variable_in_TIKL("Verification of " & employer & " job change should have returned by now. If not received and processed, take appropriate action. (TIKL auto-generated from script)." )
	transmit
	PF3
End if 

'Now the script will case note what's happened.
start_a_blank_CASE_NOTE
EMSendKey ">>>New job for MEMB " & HH_memb & " reported by " & who_reported_job & " via " & job_report_type & "<<<" & "<newline>"
call write_bullet_and_variable_in_case_note("Employer", employer)
call write_bullet_and_variable_in_case_note("Income type", income_type_dropdown)
If subsidized_income_type_dropdown <> "not applicable" then call write_bullet_and_variable_in_case_note("Subsidized income type", subsidized_income_type_dropdown)
call write_bullet_and_variable_in_case_note("Income start date", income_start_date)
if contract_through_date <> "" then call write_bullet_and_variable_in_case_note("Contract through date", contract_through_date)
if CCA_checkbox = 1 then call write_variable_in_case_note("* Sent status update to CCA.")
if ES_checkbox = 1 then call write_variable_in_case_note("* Sent status update to ES.")
if work_number_checkbox = 1 then call write_variable_in_case_note("* Sent Work Number request.")
If requested_CEI_OHI_docs_checkbox = checked then call write_variable_in_case_note("* Requested CEI/OHI docs.")
If TIKL_checkbox = checked then call write_variable_in_case_note("* TIKLed for 10-day return.")
call write_bullet_and_variable_in_case_note("Notes", notes)
call write_variable_in_case_note("---")
call write_variable_in_case_note(worker_signature)

'Navigating to DAIL/WRIT
If TIKL_checkbox = 1 then 
	script_end_procedure("Success! MAXIS updated for job change, a case note made, and a TIKL has been sent for 10 days from now. An EV should now be sent. The job is at: " & employer & ".")
Else 
	script_end_procedure("Success! MAXIS updated for job change, and a case note has been made. An EV should now be sent. The job is at: " & employer & ".")
END IF 