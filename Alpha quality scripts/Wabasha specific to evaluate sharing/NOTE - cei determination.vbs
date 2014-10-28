'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - CEI Determination"
start_time = timer
'
'LOADING ROUTINE FUNCTIONS
'<<DELETE REDUNDANCIES!
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Z:\Scripts\BlueZone Script Directory\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'SECTION 02: DIALOGS

BeginDialog verifs_needed_dialog, 0, 0, 346, 250, "CEI Determination"
  EditBox 65, 5, 70, 15, case_number
  EditBox 75, 35, 50, 15, INFO_PROVIDED
  EditBox 75, 55, 250, 15, APPROVED_FOR
  EditBox 75, 75, 250, 15, EMPLOYER
  EditBox 90, 95, 50, 15, AMOUNT
  EditBox 75, 115, 50, 15, OB_ID
  EditBox 85, 135, 30, 15, MMIS_HH_EXCLUDED
  EditBox 85, 155, 50, 15, DHS_3880_MAILED
  EditBox 85, 175, 295, 15, OTHER_COMMENTS
  CheckBox 5, 195, 225, 10, "Check here to TIKL out for this case if client needs to enroll in CEI.", TIKL_check
  EditBox 285, 205, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 200, 230, 50, 15
    CancelButton 255, 230, 50, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 25, 300, 10, "If you aren't requesting something, leave that section blank. That way it doesn't case note."
  Text 5, 40, 50, 10, "Info Rec'd:"
  Text 5, 60, 50, 10, "Approved For:"
  Text 5, 80, 50, 10, "Employer:"
  Text 5, 100, 75, 10, "Amount (MEDI Only):"
  Text 5, 120, 50, 10, "OB ID #:"
  Text 5, 140, 75, 10, "MMIS HH Exc? (Y/N)"
  Text 5, 160, 75, 10, "DHS-3880 Mailed:"
   Text 5, 180, 80, 10, "Other Comments:"
  Text 215, 210, 70, 10, "Sign your case note:"
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

call write_new_line_in_case_note(">>>CEI Determination<<<")
If INFO_PROVIDED <> "" then call write_editbox_in_case_note("Info Rec'd", INFO_PROVIDED, 6)
If APPROVED_FOR <> "" then call write_editbox_in_case_note("Approved For", APPROVED_FOR, 6)
If EMPLOYER <> "" then call write_editbox_in_case_note("employer", EMPLOYER, 6)
If AMOUNT <> "" then call write_editbox_in_case_note("Amount (MEDI Only)", AMOUNT, 6)
If OB_ID <> "" then call write_editbox_in_case_note("OB ID #", OB_ID, 6)
If MMIS_HH_EXCLUDED <> "" then call write_editbox_in_case_note("MMIS HH Exc? (Y/N)", MMIS_HH_EXCLUDED, 6)
If DHS_3880_MAILED <> "" then call write_editbox_in_case_note("DHS-3880 Mailed", DHS_3880_MAILED, 6)
If OTHER_COMMENTS <> "" then call write_editbox_in_case_note("Other Comments", OTHER_COMMENTS, 6)
If signature_page_needed_check = 1 then call write_new_line_in_case_note("* Signature page is needed.")
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

If TIKL_check = 0 then script_end_procedure("")

call navigate_to_screen("dail", "writ")
      call create_MAXIS_friendly_date(date, 30, 5, 18) 
      EMSetCursor 9, 3
      EMSendKey "IF CLIENT HAS NOT ENROLLED IN CEI, EVALUATE ONGOING ELIG"
      transmit
      PF3


script_end_procedure("")
