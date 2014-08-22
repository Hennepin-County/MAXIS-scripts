'Informational front-end message, date dependent.
If datediff("d", "04/02/2013", now) < 5 then MsgBox "This script has been updated as of 04/02/2013! There's now a checkbox for starting the denied programs script right from this one."

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - MNsure docs req'd"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS----------------------------------------------------------------------------------------------------

BeginDialog MNsure_docs_reqd_dialog, 0, 0, 301, 105, "MNsure Docs Req'd Dialog"
  EditBox 75, 5, 70, 15, case_number
  EditBox 225, 5, 70, 15, MNsure_app_date
  EditBox 45, 25, 70, 15, MNsure_ID
  EditBox 225, 25, 70, 15, application_case_number
  EditBox 50, 45, 245, 15, docs_reqd
  EditBox 50, 65, 245, 15, other_notes
  EditBox 70, 85, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 190, 85, 50, 15
    CancelButton 245, 85, 50, 15
  Text 5, 10, 70, 10, "MAXIS case number:"
  Text 165, 10, 60, 10, "MNsure app date:"
  Text 5, 30, 40, 10, "MNsure ID:"
  Text 140, 30, 85, 10, "Application Case Number:"
  Text 5, 50, 40, 10, "Doc's req'd:"
  Text 5, 70, 45, 10, "Other notes:"
  Text 5, 90, 60, 10, "Worker signature:"
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
      Dialog MNsure_docs_reqd_dialog
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

'Case notes
EMSendKey ">>>>>MNSURE DOCS REQ'D<<<<<" & "<newline>"
If MNsure_app_date <> "" then call write_editbox_in_case_note("MNsure application date", MNsure_app_date, 6)
If MNsure_ID <> "" then call write_editbox_in_case_note("MNsure ID", MNsure_ID, 6)
If application_case_number <> "" then call write_editbox_in_case_note("Application case number", application_case_number, 6)
If docs_reqd <> "" then call write_editbox_in_case_note("Docs requested", docs_reqd, 6)
If other_notes <> "" then call write_editbox_in_case_note("Other notes", other_notes, 6)
call write_editbox_in_case_note("Please note", "If these docs come into your ''My documents received'' queue in OnBase, please create a copy of the document and re-index it to the appropriate MNsure doc type, and send to the proper workflow. If you have questions, consult a member of the MNsure team.", 6)
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

script_end_procedure("")






