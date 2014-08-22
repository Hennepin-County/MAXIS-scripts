'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - verifs needed"
start_time = timer
'
'LOADING ROUTINE FUNCTIONS
'<<DELETE REDUNDANCIES!
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'SECTION 02: DIALOGS

BeginDialog verifs_needed_dialog, 0, 0, 346, 292, "Verifs needed"
  EditBox 55, 5, 70, 15, case_number
  EditBox 30, 35, 315, 15, ADDR
  EditBox 70, 55, 275, 15, SCHL
  EditBox 30, 75, 315, 15, DISA
  EditBox 30, 95, 315, 15, JOBS
  EditBox 30, 115, 315, 15, BUSI
  EditBox 30, 135, 315, 15, UNEA
  EditBox 30, 155, 315, 15, ACCT
  EditBox 55, 175, 290, 15, other_assets
  EditBox 30, 195, 315, 15, SHEL
  EditBox 30, 215, 315, 15, INSA
  EditBox 50, 235, 295, 15, other_proofs
  CheckBox 5, 255, 95, 10, "Signature page needed?", signature_page_needed_check
  CheckBox 5, 270, 130, 10, "Check here to TIKL out for this case.", TIKL_check
  EditBox 285, 255, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 200, 275, 50, 15
    CancelButton 255, 275, 50, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 25, 300, 10, "If you aren't requesting something, leave that section blank. That way it doesn't case note."
  Text 5, 40, 25, 10, "ADDR:"
  Text 5, 60, 60, 10, "SCHL/STIN/STEC:"
  Text 5, 80, 25, 10, "DISA:"
  Text 5, 100, 25, 10, "JOBS:"
  Text 5, 120, 20, 10, "BUSI:"
  Text 5, 140, 25, 10, "UNEA:"
  Text 5, 160, 25, 10, "ACCT:"
  Text 5, 180, 50, 10, "Other assets:"
  Text 5, 200, 25, 10, "SHEL:"
  Text 5, 220, 25, 10, "INSA:"
  Text 5, 240, 45, 10, "Other proofs:"
  Text 215, 260, 70, 10, "Sign your case note:"
EndDialog


'SECTION 03: THE SCRIPT

EMConnect ""

call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""

Do
  Do
    Do
      Dialog verifs_needed_dialog
      If buttonpressed = 0 then stopscript
      If case_number = "" then MsgBox "You must have a case number to continue!"
    Loop until case_number <> ""
    transmit
    EMReadScreen MAXIS_check, 5, 1, 39
    If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You appear to be locked out of MAXIS. Are you passworded out? Did you navigate away from MAXIS?"
  Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "
  call navigate_to_screen("case", "note")
  PF9
  EMReadScreen mode_check, 7, 20, 3
  If mode_check <> "Mode: A" and mode_check <> "Mode: E" then MsgBox "For some reason, the script can't get to a case note. Did you start the script in inquiry by mistake? Navigate to MAXIS production, or shut down the script and try again."
Loop until mode_check = "Mode: A" or mode_check = "Mode: E"

call write_new_line_in_case_note(">>>Verifications Requested<<<")
If ADDR <> "" then call write_editbox_in_case_note("ADDR", ADDR, 6)
If SCHL <> "" then call write_editbox_in_case_note("SCHL/STIN/STEC", SCHL, 6)
If DISA <> "" then call write_editbox_in_case_note("DISA", DISA, 6)
If JOBS <> "" then call write_editbox_in_case_note("JOBS", JOBS, 6)
If BUSI <> "" then call write_editbox_in_case_note("BUSI", BUSI, 6)
If UNEA <> "" then call write_editbox_in_case_note("UNEA", UNEA, 6)
If ACCT <> "" then call write_editbox_in_case_note("ACCT", ACCT, 6)
If other_assets <> "" then call write_editbox_in_case_note("Other assets", other_assets, 6)
If SHEL <> "" then call write_editbox_in_case_note("SHEL", SHEL, 6)
If INSA <> "" then call write_editbox_in_case_note("INSA", INSA, 6)
If other_proofs <> "" then call write_editbox_in_case_note("Other proofs", other_proofs, 6)
If signature_page_needed_check = 1 then call write_new_line_in_case_note("* Signature page is needed.")
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

If TIKL_check = 0 then script_end_procedure("")

call navigate_to_screen("dail", "writ")

script_end_procedure("")






