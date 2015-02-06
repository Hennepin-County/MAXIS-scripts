'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - NEW JOB REPORTED.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, FALSE									'Attempts to open the URL
req.send													'Sends request
IF req.Status = 200 THEN									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			script_end_procedure("Script ended due to error connecting to GitHub.")
END IF

'DIALOGS----------------------------------------------------------------------------------------------------

BeginDialog case_number_and_footer_month_dialog, 0, 0, 161, 65, "Case number and footer month"
  Text 5, 10, 85, 10, "Enter your case number:"
  EditBox 95, 5, 60, 15, case_number
  Text 15, 30, 50, 10, "Footer month:"
  EditBox 65, 25, 25, 15, footer_month
  Text 95, 30, 20, 10, "Year:"
  EditBox 120, 25, 25, 15, footer_year
  ButtonGroup ButtonPressed
    OkButton 25, 45, 50, 15
    CancelButton 85, 45, 50, 15
EndDialog


BeginDialog new_job_reported_dialog, 0, 0, 286, 280, "New job reported dialog"
  EditBox 80, 5, 25, 15, HH_memb
  DropListBox 55, 25, 110, 15, "W Wages (Incl Tips)"+chr(9)+"J WIA (JTPA)"+chr(9)+"E EITC"+chr(9)+"G Experience Works"+chr(9)+"F Federal Work Study"+chr(9)+"S State Work Study"+chr(9)+"O Other"+chr(9)+"I Infrequent < 30 N/Recur"+chr(9)+"M Infreq <= 10 MSA Exclusion"+chr(9)+"C Contract Income", income_type_dropdown
  DropListBox 135, 45, 150, 15, "not applicable"+chr(9)+"01 Subsidized Public Sector Employer"+chr(9)+"02 Subsidized Private Sector Employer"+chr(9)+"03 On-the-Job-Training"+chr(9)+"04 AmeriCorps (VISTA/State/National/NCCC)", subsidized_income_type_dropdown
  EditBox 45, 65, 195, 15, employer
  EditBox 100, 85, 55, 15, income_start_date
  EditBox 125, 105, 55, 15, contract_through_date
  EditBox 90, 125, 100, 15, who_reported_job
  ComboBox 100, 145, 90, 15, "phone call"+chr(9)+"office visit"+chr(9)+"mailing"+chr(9)+"fax"+chr(9)+"ES counselor"+chr(9)+"CCA worker"+chr(9)+"scanned document", job_report_type
  EditBox 30, 165, 210, 15, notes
  CheckBox 5, 185, 190, 10, "Check here to have the script make a new JOBS panel.", create_JOBS_checkbox
  CheckBox 5, 200, 190, 10, "Check here if you sent a status update to CCA.", CCA_checkbox
  CheckBox 5, 215, 190, 10, "Check here if you sent a status update to ES.", ES_checkbox
  CheckBox 5, 230, 190, 10, "Check here if you sent a Work Number request.", work_number_checkbox
  CheckBox 5, 245, 165, 10, "Check here if you are requesting CEI/OHI docs.", requested_CEI_OHI_docs_checkbox
  EditBox 70, 260, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 175, 260, 50, 15
    CancelButton 230, 260, 50, 15
    PushButton 175, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 175, 25, 45, 10, "next panel", next_panel_button
    PushButton 235, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 235, 25, 45, 10, "next memb", next_memb_button
  Text 5, 10, 70, 10, "HH member number:"
  GroupBox 170, 5, 115, 35, "STAT-based navigation"
  Text 5, 30, 45, 10, "Income Type: "
  Text 5, 50, 130, 10, "Subsidized Income Type (if applicable):"
  Text 5, 70, 40, 10, "Employer:"
  Text 5, 90, 95, 10, "Income start date (if known):"
  Text 5, 110, 120, 10, "Contract through date (if applicable):"
  Text 5, 130, 80, 10, "Who reported the job?:"
  Text 5, 150, 90, 10, "How was the job reported?:"
  Text 5, 170, 25, 10, "Notes:"
  Text 5, 265, 60, 10, "Worker signature:"
EndDialog

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
footer_month = datepart("m", dateadd("m", 1, date))
If len(footer_month) = 1 then footer_month = "0" & footer_month
footer_year = "" & datepart("yyyy", dateadd("m", 1, date)) - 2000




'THE SCRIPT----------------------------------------------------------------------------------------------------

EMConnect ""

'Finds a case number
call MAXIS_case_number_finder(case_number)

'Shows the case number dialog
Dialog case_number_and_footer_month_dialog
If ButtonPressed = 0 then stopscript

'It sends an enter to force the screen to refresh, in order to check for MAXIS. If MAXIS isn't found the script will stop.
transmit
EMReadScreen MAXIS_check, 5, 1, 39
IF MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then script_end_procedure("MAXIS not found. Are you in MAXIS on the screen you started the script? Check and try again. If it still doesn't work try shutting down BlueZone and starting it up again.")

'Checks footer month and year. If footer month and year do not match the worker entry, it'll back out and get there manually.
EMReadScreen footer_month_year_check, 5, 20, 55
If left(footer_month_year_check, 2) <> footer_month or right(footer_month_year_check, 2) <> footer_year then
	back_to_self
	EMWriteScreen "________", 18, 43
	EMWriteScreen footer_month, 20, 43
	EMWriteScreen footer_year, 20, 46
	transmit
End if

'Now it enters stat/jobs. It'll check to make sure it gets past the SELF menu and gets onto the JOBS panel.
call navigate_to_screen("stat", "jobs")
EMReadScreen SELF_check, 27, 2, 28
If SELF_check = "Select Function Menu (SELF)" then script_end_procedure("Unable to navigate past the SELF menu. Is your case in background? Wait a few seconds and try again.")

'Declaring some variables to create defaults for the new_job_reported_dialog.
create_JOBS_checkbox = 1
HH_memb = "01"
HH_memb_row = 5 'This helps the navigation buttons work!

'Shows the dialog.
Do
	Do
		Do
			Do
				Do
					Do
						Do
							Dialog new_job_reported_dialog
							If ButtonPressed = cancel then stopscript
							EMReadScreen STAT_check, 4, 20, 21
							If STAT_check = "STAT" then call stat_navigation
							transmit 'Forces a screen refresh, to keep MAXIS from erroring out in the event of a password prompt.
							EMReadScreen MAXIS_check, 5, 1, 39
							If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then script_end_procedure("You do not appear to be in MAXIS. Are you passworded out? Or in MMIS? Check these and try again.")
						Loop until ButtonPressed = OK
						If isdate(income_start_date) = True then		'Logic to determine if the income start date is functional
							If (datediff("m", footer_month & "/01/20" & footer_year, income_start_date) > 0) then
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

'Creates a new JOBS panel if that was selected.
If create_JOBS_checkbox = checked then
	EMWriteScreen HH_memb, 20, 76
	EMWriteScreen "nn", 20, 79
	transmit
	EMReadScreen edit_mode_check, 1, 20, 8
	If edit_mode_check = "D" then script_end_procedure("Unable to create a new JOBS panel. Check which member number you provided. Otherwise you may be in inquiry mode. If so shut down inquiry and try again. Or try closing BlueZone.")
	EMWriteScreen left(income_type_dropdown, 1), 5, 38
	If subsidized_income_type_dropdown <> "not applicable" then EMWriteScreen left(subsidized_income_type_dropdown, 2), 5, 71
	EMWriteScreen "n", 6, 38
	EMWriteScreen employer, 7, 42
	If income_start_date <> "" then call create_MAXIS_friendly_date(income_start_date, 0, 9, 35)
	If contract_through_date <> "" then call create_MAXIS_friendly_date(contract_through_date, 0, 9, 73)
	EMReadScreen footer_month, 2, 20, 55
	EMReadScreen footer_year, 2, 20, 58
	If isdate(income_start_date) = True then
		If datediff("d", income_start_date, footer_month & "/01/20" & footer_year) > 0 then
			call create_MAXIS_friendly_date(footer_month & "/01/20" & footer_year, 0, 12, 54)
		Else
			call create_MAXIS_friendly_date(income_start_date, 0, 12, 54)
		End if
	Else
		call create_MAXIS_friendly_date(footer_month & "/01/20" & footer_year, 0, 12, 54)	
	End if
	EMWriteScreen "0", 12, 67
	EMWriteScreen "0", 18, 72
	Do
		transmit
		EMReadScreen edit_mode_check, 1, 20, 8
	Loop until edit_mode_check = "D"
End if

'Jumps to case note the info.
call navigate_to_screen("case", "note")
PF9
EMReadScreen edit_mode_check, 1, 20, 9
If edit_mode_check = "D" then script_end_procedure("Unable to create a new case note. Your case may be in inquiry. If so shut down inquiry and try again. Or try closing BlueZone.")

'Now the script will case note what's happened.
EMSendKey ">>>New job for MEMB " & HH_memb & " reported by " & who_reported_job & " via " & job_report_type & "<<<" & "<newline>" 
call write_bullet_and_variable_in_case_note("Employer", employer)
call write_bullet_and_variable_in_case_note("Income type", income_type_dropdown)
If subsidized_income_type_dropdown <> "not applicable" then call write_bullet_and_variable_in_case_note("Subsidized income type", subsidized_income_type_dropdown)
if income_start_date <> "" then call write_bullet_and_variable_in_case_note("Income start date", income_start_date)
if contract_through_date <> "" then call write_bullet_and_variable_in_case_note("Contract through date", contract_through_date)
if CCA_checkbox = 1 then call write_variable_in_case_note("* Sent status update to CCA.")
if ES_checkbox = 1 then call write_variable_in_case_note("* Sent status update to ES.")
if work_number_checkbox = 1 then call write_variable_in_case_note("* Sent Work Number request.")
If requested_CEI_OHI_docs_checkbox = checked then call write_variable_in_case_note("* Requested CEI/OHI docs.")
if notes <> "" then call write_bullet_and_variable_in_case_note("Notes", notes)
call write_variable_in_case_note("* Sending employment verification. TIKLed for 10-day return.")
call write_variable_in_case_note("---")
call write_variable_in_case_note(worker_signature)

'Navigating to DAIL/WRIT
call navigate_to_screen("dail", "writ")

'The following will generate a TIKL formatted date for 10 days from now.
call create_MAXIS_friendly_date(date, 10, 5, 18)

'Writing in the rest of the TIKL.
call write_variable_in_TIKL("Verification of job change should have returned by now. If not received and processed, take appropriate action. (TIKL auto-generated from script)." )

transmit
PF3
MsgBox "Success! MAXIS updated for job change, a case note made, and a TIKL has been sent for 10 days from now. An EV should now be sent. The job is at " & employer & "."

script_end_procedure("")






