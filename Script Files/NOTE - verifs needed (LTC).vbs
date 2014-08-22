'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - verifs needed (LTC)"
start_time = timer

'LOADING ROUTINE FUNCTIONS
'<<DELETE REDUNDANCIES!
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'SECTION 02: DIALOGS

BeginDialog verifs_requested_dialog, 0, 0, 351, 392, "Verifs requested"
  EditBox 55, 5, 70, 15, case_number
  EditBox 30, 40, 315, 15, FACI
  EditBox 30, 60, 315, 15, JOBS
  EditBox 50, 80, 295, 15, BUSI_RBIC
  EditBox 45, 100, 300, 15, UNEA_01
  EditBox 75, 120, 270, 15, UNEA_other_membs
  EditBox 45, 140, 300, 15, ACCT_01
  EditBox 75, 160, 270, 15, ACCT_other_membs
  EditBox 45, 180, 300, 15, SECU_01
  EditBox 75, 200, 270, 15, SECU_other_membs
  EditBox 30, 220, 315, 15, CARS
  EditBox 30, 240, 315, 15, REST
  EditBox 30, 260, 315, 15, OTHR
  EditBox 30, 280, 315, 15, SHEL
  EditBox 30, 300, 315, 15, INSA
  EditBox 75, 320, 270, 15, medical_expenses
  EditBox 50, 340, 295, 15, other_proofs
  CheckBox 5, 360, 335, 10, "Check here to TIKL out for this case.", TIKL_check
  EditBox 75, 370, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 225, 375, 50, 15
    CancelButton 280, 375, 50, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 25, 300, 10, "If you aren't requesting something, leave that section blank. That way it doesn't case note."
  Text 5, 45, 25, 10, "FACI:"
  Text 5, 65, 25, 10, "JOBS:"
  Text 5, 85, 40, 10, "BUSI/RBIC:"
  Text 5, 105, 35, 10, "UNEA 01:"
  Text 5, 125, 65, 10, "UNEA other membs:"
  Text 5, 145, 35, 10, "ACCT 01:"
  Text 5, 165, 65, 10, "ACCT other membs:"
  Text 5, 185, 35, 10, "SECU 01:"
  Text 5, 205, 70, 10, "SECU other membs:"
  Text 5, 225, 25, 10, "CARS:"
  Text 5, 245, 25, 10, "REST:"
  Text 5, 265, 25, 10, "OTHR:"
  Text 5, 285, 25, 10, "SHEL:"
  Text 5, 305, 25, 10, "INSA:"
  Text 5, 325, 65, 10, "Medical expenses:"
  Text 5, 345, 45, 10, "Other proofs:"
  Text 5, 375, 70, 10, "Sign your case note:"
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
      Dialog verifs_requested_dialog
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
If FACI <> "" then call write_editbox_in_case_note("FACI", FACI, 6)
If JOBS <> "" then call write_editbox_in_case_note("JOBS", JOBS, 6)
If BUSI_RBIC <> "" then call write_editbox_in_case_note("BUSI/RBIC", BUSI_RBIC, 6)
If UNEA_01 <> "" then call write_editbox_in_case_note("UNEA 01", UNEA_01, 6)
If UNEA_other_membs <> "" then call write_editbox_in_case_note("UNEA other membs", UNEA_other_membs, 6)
If ACCT_01 <> "" then call write_editbox_in_case_note("ACCT 01", ACCT_01, 6)
If ACCT_other_membs <> "" then call write_editbox_in_case_note("ACCT other membs", ACCT_other_membs, 6)
If SECU_01 <> "" then call write_editbox_in_case_note("SECU 01", SECU_01, 6)
If SECU_other_membs <> "" then call write_editbox_in_case_note("SECU other membs", SECU_other_membs, 6)
If CARS <> "" then call write_editbox_in_case_note("CARS", CARS, 6)
If REST <> "" then call write_editbox_in_case_note("REST", REST, 6)
If OTHR <> "" then call write_editbox_in_case_note("OTHR", OTHR, 6)
If SHEL <> "" then call write_editbox_in_case_note("SHEL", SHEL, 6)
If INSA <> "" then call write_editbox_in_case_note("INSA", INSA, 6)
If medical_expenses <> "" then call write_editbox_in_case_note("Medical expenses", medical_expenses, 6)
If other_proofs <> "" then call write_editbox_in_case_note("Other proofs", other_proofs, 6)
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

If TIKL_check = 0 then script_end_procedure("")

call navigate_to_screen("dail", "writ")

script_end_procedure("")






