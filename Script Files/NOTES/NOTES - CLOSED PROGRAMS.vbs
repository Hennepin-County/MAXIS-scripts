'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - CLOSED PROGRAMS.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
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
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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

'VARIABLE REQUIRED TO RESIZE DIALOG BASED ON A GLOBAL VARIABLE IN FUNCTIONS FILE
If case_noting_intake_dates = False then dialog_shrink_amt = 45

'THE DIALOG----------------------------------------------------------------------------------------------------
BeginDialog closed_dialog, 0, 0, 421, 240 - dialog_shrink_amt, "Closed progs dialog"
  EditBox 65, 5, 55, 15, case_number
  EditBox 195, 5, 55, 15, closure_date
  CheckBox 315, 10, 35, 10, "SNAP", SNAP_check
  CheckBox 355, 10, 35, 10, "cash", cash_check
  CheckBox 395, 10, 25, 10, "HC", HC_check
  EditBox 80, 25, 340, 15, reason_for_closure
  EditBox 140, 45, 280, 15, verifs_needed
  EditBox 170, 65, 250, 15, open_progs
  GroupBox 5, 90, 410, 45, "Elements that affect the REIN date/docs needed"
  CheckBox 15, 100, 360, 15, "Case is at renewal (monthy, six-month, annual. Client gets entire next month after closure for REIN.)", CSR_check
  CheckBox 15, 120, 360, 10, "Case is at HC annual renewal (can turn in HC ER instead of new HCAPP, but still counts as new application)", HC_ER_check
  CheckBox 5, 145, 65, 10, "Updated MMIS?", updated_MMIS_check
  CheckBox 85, 145, 95, 10, "WCOM added to notice?", WCOM_check
  CheckBox 5, 160, 220, 10, "Check here if closure is due to client death and enter death date.", death_check
  EditBox 230, 160, 55, 15, hc_close_for_death_date
  EditBox 345, 140, 75, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 315, 160, 50, 15
    CancelButton 370, 160, 50, 15
    PushButton 180, 145, 50, 10, "SPEC/WCOM", SPEC_WCOM_button
  Text 5, 10, 50, 10, "Case number:"
  Text 135, 10, 55, 10, "Date of closure:"
  Text 265, 10, 50, 10, "Progs closed:"
  Text 5, 30, 70, 10, "Reason for closure:"
  Text 5, 50, 130, 10, "Verifs/docs/apps needed (if applicable):"
  Text 5, 70, 165, 10, "Are any programs still open? If so, list them here:"
  Text 280, 145, 60, 10, "Worker signature: "
  GroupBox 10, 175, 400, 60, "Please note for SNAP:"
  Text 20, 185, 380, 25, "Per CM 0005.09.06, we no longer require completion of a CAF when the unit has been closed for less than one month AND the reason for closing has not changed, if the unit fully resolves the reason for the SNAP case closure given on the closing notice sent in MAXIS."
  Text 20, 215, 380, 20, "As a result, SNAP cases who turn in proofs required (or otherwise become eligible for their reamining budget period) can be REINed (with proration) up until the end of the next month. If you have questions, consult a supervisor."
EndDialog


'Connects to BlueZone
EMConnect ""

'Resets variables in case this is run from docs received.
SNAP_check = 0
cash_check = 0
HC_check = 0

'Autofills case number
call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If isnumeric(case_number) = False then case_number = ""


'Dialog starts. Checks for MAXIS, includes nav button for SPEC/WCOM, validates the date of closure, confirms that date 
'    of closure is last day of a month, checks that a program was selected for closure, and navigates to CASE/NOTE.
Do
  Do
    Do
      Do
        Do
          Do
            Dialog closed_dialog
            If buttonpressed = 0 then stopscript
            transmit
            EMReadScreen MAXIS_check, 5, 1, 39
            If MAXIS_check <> "MAXIS" then MsgBox "You do not appear to be in MAXIS. You may be passworded out. Please check your MAXIS screen and try again."
          Loop until MAXIS_check = "MAXIS"
          If ButtonPressed = SPEC_WCOM_button then call navigate_to_screen("spec", "wcom")
        Loop until ButtonPressed = -1
        If isdate(closure_date) = False then MsgBox "You need to enter a valid date of closure (MM/DD/YYYY)."
        IF (death_check = 1 AND isdate(hc_close_for_death_date) = FALSE) THEN MsgBox "Please enter a date in the correct format (MM/DD/YYYY)."
	  IF (death_check <> 1 AND hc_close_for_death_date <> "") THEN MsgBox "Please check the box for client death."
      Loop until isdate(closure_date) = True AND ((death_check = 1 AND isdate(hc_close_for_death_date) = TRUE) OR (death_check <> 1 AND hc_close_for_death_date = ""))
      If datepart("d", dateadd("d", 1, closure_date)) <> 1 then MsgBox "Please use the last date of eligibility, which for an open case, should be the last day of the month. If this is a denial, use the denial script."
    Loop until datepart("d", dateadd("d", 1, closure_date)) = 1
    If SNAP_check = 0 and HC_check = 0 and cash_check = 0 then MsgBox "You need to select a program to close."
  Loop until SNAP_check = 1 or HC_check = 1 or cash_check = 1
  call navigate_to_screen("case", "note")
  PF9
  EMReadScreen mode_check, 7, 20, 3
  If mode_check <> "Mode: A" and mode_check <> "Mode: E" then MsgBox "You do not appear to be able to edit a case note. This case could have errored out, or might be in another county. Or you could be on inquiry. Check the case number, and try again."
Loop until mode_check = "Mode: A" or mode_check = "Mode: E"

'Converting dates for intake/REIN/reapp date calculations
closure_date = cdate(closure_date)                                                       'just running a cdate on the closure_date variable
closure_month_first_day = dateadd("d", 1, closure_date)                                  'This is the first day of the closure month
closure_month_last_day = dateadd("d", -1, (dateadd("m", 1, closure_month_first_day)))    'This is the last day of the closure month
intake_date = dateadd("m", 1, closure_month_first_day)                                   'Becomes an intake the first of the month after closure (case would be assigned to a new worker if they reapply)

'If the case is at the CSR then the last REIN date is automatically ahead one month. Otherwise it's the date of closure for all programs except SNAP.
If CSR_check = 1 then
  If HC_check = 1 then HC_last_REIN_date = closure_month_last_day
  If cash_check = 1 then cash_last_REIN_date = closure_month_last_day
Else
  If HC_check = 1 then HC_last_REIN_date = closure_date
  If cash_check = 1 then cash_last_REIN_date = closure_date
End if

'For HC, the client can turn in a HC ER before the end of the next month, or turn in a HCAPP anytime. Either way it's just treated as a new app.
If HC_check = 1 then
  progs_closed = progs_closed & "HC/"
  If HC_ER_check = 1 then 
    HC_last_REIN_date = closure_date 'This will force the HC_last_REIN_date variable to show the closure date, in case a SNAP CSR messes with the variable. HC ERs are always treated like a new application if turned in late.
    HC_followup_text = ", after which a new HC renewal or HCAPP is required. A new HCAPP is required after " & closure_month_last_day & "."
  End if
  If HC_ER_check = 0 then HC_followup_text = ", after which a new HCAPP is required."
End if

'SNAP closures have different logic for when to REIN. For SNAP the client gets an additional month to turn in proofs, and can be REINed without a new app.
If SNAP_check = 1 then
  progs_closed = progs_closed & "SNAP/"
  SNAP_last_REIN_date = closure_month_last_day
  SNAP_followup_text = ", after which a new CAF is required."
End if

'Cash cases use similar logic to HC but don't have the "HC renewal can be used as a new app" issue.
If cash_check = 1 then
  progs_closed = progs_closed & "cash/"
  cash_followup_text = ", after which a new CAF is required."
End if

'Removing the last character of the progs_closed variable, as it is always a trailing slash
progs_closed = left(progs_closed, len(progs_closed) - 1)

'The dialog navigated to CASE/NOTE. This will write the info into the case note.
IF death_check = 1 THEN
	case_note_header = "---Closed " & progs_closed & " due to client death---"
ELSE
	case_note_header = "---Closed " & progs_closed & " " & closure_date & "---"
END IF
call write_variable_in_case_note(case_note_header)
IF death_check = 1 AND HC_check = 1 THEN call write_bullet_and_variable_in_case_note("HC Closure Date", hc_close_for_death_date)
IF death_check = 1 AND snap_check = 1 THEN call write_bullet_and_variable_in_case_note("SNAP Closure Date", closure_date)
IF death_check = 1 AND cash_check = 1 THEN call write_bullet_and_variable_in_case_note("CASH Closure Date", closure_date)
call write_bullet_and_variable_in_case_note("Reason for closure", reason_for_closure)
If verifs_needed <> "" then call write_bullet_and_variable_in_case_note("Verifs needed", verifs_needed)
If updated_MMIS_check = 1 then call write_variable_in_case_note("* Updated MMIS.")
If WCOM_check = 1 then call write_variable_in_case_note("* Added WCOM to notice.")
If CSR_check = 1 then call write_bullet_and_variable_in_case_note("Case is at renewal", "client has an additional month to turn in the document and any required proofs.")
If HC_ER_check = 1 then call write_variable_in_case_note("* Case is at HC ER.")
If case_noting_intake_dates = True then
	call write_variable_in_case_note("---")
	IF death_check <> 1 and HC_check = 1 then call write_bullet_and_variable_in_case_note("Last HC REIN date", HC_last_REIN_date & HC_followup_text)
	If death_check <> 1 and SNAP_check = 1 then call write_bullet_and_variable_in_case_note("Last SNAP REIN date", SNAP_last_REIN_date & SNAP_followup_text)
	If death_check <> 1 and cash_check = 1 then call write_bullet_and_variable_in_case_note("Last cash REIN date", cash_last_REIN_date & cash_followup_text)
	If open_progs <> "" and len(open_progs) > 1 and open_progs <> "no" and open_progs <> "NO" and open_progs <> "No" and open_progs <> "n/a" and open_progs <> "N/A" and open_progs <> "NA" and open_progs <> "na" then
		call write_bullet_and_variable_in_case_note("Open programs", open_progs)
	Else
		IF death_check = 1 THEN call write_variable_in_case_note("* All programs closed.")
		IF death_check <> 1 THEN call write_variable_in_case_note("* All programs closed. Case becomes intake again on " & intake_date & ".")
	End if
Else
	If open_progs <> "" and len(open_progs) > 1 and open_progs <> "no" and open_progs <> "NO" and open_progs <> "No" and open_progs <> "n/a" and open_progs <> "N/A" and open_progs <> "NA" and open_progs <> "na" then call write_bullet_and_variable_in_case_note("Open programs", open_progs)
End if
call write_variable_in_case_note("---")
call write_variable_in_case_note(worker_signature)

'Runs denied progs if selected
If denied_progs_check = 1 then run_another_script("C:\DHS-MAXIS-Scripts\Script Files\NOTE - denied progs.vbs")

script_end_procedure("Success! Please remember to check the generated notice to make sure it reads correctly. If not please add WCOMs to make notice read correctly.")


