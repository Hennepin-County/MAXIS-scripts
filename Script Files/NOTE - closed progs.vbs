'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - closed progs"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

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
  If case_noting_intake_dates = True then
    GroupBox 5, 90, 410, 45, "Elements that affect the REIN date/docs needed"
    CheckBox 15, 100, 230, 15, "Case is at CSR recert (gets entire next month after closure for REIN)", CSR_check
    CheckBox 15, 120, 360, 10, "Case is at HC annual renewal (can turn in HC ER instead of new HCAPP, but still counts as new application)", HC_ER_check
  End if
  CheckBox 5, 145 - dialog_shrink_amt, 65, 10, "Updated MMIS?", updated_MMIS_check
  CheckBox 85, 145 - dialog_shrink_amt, 95, 10, "WCOM added to notice?", WCOM_check
  EditBox 345, 140 - dialog_shrink_amt, 75, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 315, 160 - dialog_shrink_amt, 50, 15
    CancelButton 370, 160 - dialog_shrink_amt, 50, 15
    PushButton 180, 145 - dialog_shrink_amt, 50, 10, "SPEC/WCOM", SPEC_WCOM_button
  Text 5, 10, 50, 10, "Case number:"
  Text 135, 10, 55, 10, "Date of closure:"
  Text 265, 10, 50, 10, "Progs closed:"
  Text 5, 30, 70, 10, "Reason for closure:"
  Text 5, 50, 130, 10, "Verifs/docs/apps needed (if applicable):"
  Text 5, 70, 165, 10, "Are any programs still open? If so, list them here:"
  Text 280, 145 - dialog_shrink_amt, 60, 10, "Worker signature: "
  GroupBox 10, 175 - dialog_shrink_amt, 400, 60, "Please note for SNAP:"
  Text 20, 185 - dialog_shrink_amt, 380, 25, "Per CM 0005.09.06, we no longer require completion of a CAF when the unit has been closed for less than one month AND the reason for closing has not changed, if the unit fully resolves the reason for the SNAP case closure given on the closing notice sent in MAXIS."
  Text 20, 215 - dialog_shrink_amt, 380, 20, "As a result, SNAP cases who turn in proofs required (or otherwise become eligible for their reamining budget period) can be REINed (with proration) up until the end of the next month. If you have questions, consult a supervisor."
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
      Loop until isdate(closure_date) = True
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

'The dialog navigated to CASE/NOTE. This will write the info into the case note.
EMSendKey "---Closed " & progs_closed & "<backspace>" & " " & closure_date & "---" & "<newline>"
call write_editbox_in_case_note("Reason for closure", reason_for_closure, 6)
If verifs_needed <> "" then call write_editbox_in_case_note("Verifs needed", verifs_needed, 6)
If updated_MMIS_check = 1 then call write_new_line_in_case_note("* Updated MMIS.")
If WCOM_check = 1 then call write_new_line_in_case_note("* Added WCOM to notice.")
If CSR_check = 1 then call write_editbox_in_case_note("Case is at CSR", "client has an additional month to turn in the CSR and any required proofs.", 6)
If HC_ER_check = 1 then call write_new_line_in_case_note("* Case is at HC ER.")
If case_noting_intake_dates = True then
	call write_new_line_in_case_note("---")
	If HC_check = 1 then call write_editbox_in_case_note("Last HC REIN date", HC_last_REIN_date & HC_followup_text, 6)
	If SNAP_check = 1 then call write_editbox_in_case_note("Last SNAP REIN date", SNAP_last_REIN_date & SNAP_followup_text, 6)
	If cash_check = 1 then call write_editbox_in_case_note("Last cash REIN date", cash_last_REIN_date & cash_followup_text, 6)
	If open_progs <> "" and len(open_progs) > 1 and open_progs <> "no" and open_progs <> "NO" and open_progs <> "No" and open_progs <> "n/a" and open_progs <> "N/A" and open_progs <> "NA" and open_progs <> "na" then
		call write_editbox_in_case_note("Open programs", open_progs, 6)
	Else
		call write_new_line_in_case_note("* All programs closed. Case becomes intake again on " & intake_date & ".")
	End if
Else
	If open_progs <> "" and len(open_progs) > 1 and open_progs <> "no" and open_progs <> "NO" and open_progs <> "No" and open_progs <> "n/a" and open_progs <> "N/A" and open_progs <> "NA" and open_progs <> "na" then call write_editbox_in_case_note("Open programs", open_progs, 6)
End if
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

'Runs denied progs if selected
If denied_progs_check = 1 then run_another_script("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\NOTE - denied progs.vbs")

script_end_procedure("")

