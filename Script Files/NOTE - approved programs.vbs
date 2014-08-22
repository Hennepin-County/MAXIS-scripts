'Created by Robert Kalb and Charles Potter from Anoka County.

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - approved programs"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS----------------------------------------------------------------------------------------------------

BeginDialog benefits_approved, 0, 0, 271, 190, "Benefits Approved"
  CheckBox 15, 25, 35, 10, "SNAP", snap_approved_check
  CheckBox 75, 25, 55, 10, "Health Care", hc_approved_check
  CheckBox 155, 25, 35, 10, "Cash", cash_approved_check
  CheckBox 210, 25, 55, 10, "Emergency", emer_approved_check
  EditBox 120, 70, 145, 15, benefit_breakdown
  EditBox 55, 90, 210, 15, other_notes
  EditBox 85, 110, 180, 15, programs_pending
  EditBox 65, 130, 200, 15, docs_needed
  EditBox 65, 150, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 155, 170, 50, 15
    CancelButton 210, 170, 50, 15
  Text 5, 5, 70, 10, "Approved Programs:"
  Text 5, 65, 110, 20, "Benefit Breakdown (Issuance/Spenddown/Premium):"
  Text 5, 95, 45, 10, "Other Notes:"
  Text 5, 115, 75, 10, "Pending Program(s):"
  Text 5, 135, 55, 10, "Verifs Needed:"
  Text 5, 155, 60, 10, "Worker Signature: "
  Text 5, 50, 55, 10, "Case Number:"
  EditBox 65, 45, 70, 15, case_number
EndDialog



'THE SCRIPT----------------------------------------------------------------------------------------------------

EMConnect ""

'Finds the case number
call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""

'Displays the dialog and navigates to case note
Do
  Do
    Do
      Dialog benefits_approved
      If buttonpressed = 0 then stopscript
      If case_number = "" then MsgBox "You must have a case number to continue!"
	If worker_signature = "" then Msgbox "Please sign your case note"
    Loop until case_number <> "" and worker_signature <> ""
    transmit
    EMReadScreen MAXIS_check, 5, 1, 39
    If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You appear to be locked out of MAXIS. Are you passworded out? Did you navigate away from MAXIS?"
  Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "
  call navigate_to_screen("case", "note")
  PF9
  EMReadScreen mode_check, 7, 20, 3
  If mode_check <> "Mode: A" and mode_check <> "Mode: E" then MsgBox "For some reason, the script can't get to a case note. Did you start the script in inquiry by mistake? Navigate to MAXIS production, or shut down the script and try again."
Loop until mode_check = "Mode: A" or mode_check = "Mode: E"

'Case notes
  IF snap_approved_check = 1 THEN approved_programs = approved_programs & "SNAP/"
  IF hc_approved_check = 1 THEN approved_programs = approved_programs & "HC/"
  IF cash_approved_check = 1 THEN approved_programs = approved_programs & "CASH/"
  IF emer_approved_check = 1 THEN approved_programs = approved_programs & "EMER/"
  EMSendKey "---Approved " & approved_programs & "<backspace>" & "---" & "<newline>"
  IF benefit_breakdown <> "" THEN call write_editbox_in_case_note("Benefit Breakdown", benefit_breakdown, 6)
  IF other_notes <> "" THEN call write_editbox_in_case_note("Approval Notes", other_notes, 6)
  IF programs_pending <> "" THEN call write_editbox_in_case_note("Programs Pending", programs_pending, 6)
  If docs_needed <> "" then call write_editbox_in_case_note("Verifs needed", docs_needed, 6) 
  call write_new_line_in_case_note("---")
  call write_new_line_in_case_note(worker_signature)

'Runs denied progs if selected
If closed_progs_check = 1 then run_another_script("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\NOTE - closed progs.vbs")

'Runs denied progs if selected
If denied_progs_check = 1 then run_another_script("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\NOTE - denied progs.vbs")


script_end_procedure("")



