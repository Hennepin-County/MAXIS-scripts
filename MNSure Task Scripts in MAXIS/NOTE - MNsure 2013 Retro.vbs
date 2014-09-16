'Informational front-end message, date dependent.
If datediff("d", "04/02/2013", now) < 5 then MsgBox "This script has been updated as of 04/02/2013! There's now a checkbox for starting the denied programs script right from this one."

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - MNsure 2013 Retro"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("M:\FAD-South\Test scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS----------------------------------------------------------------------------------------------------

BeginDialog MNsure_2013_Coverage_Retro_Dialog, 0, 0, 288, 65, "MNsure 2013 Coverage/Retro"
  EditBox 75, 4, 60, 12, case_number
  EditBox 224, 4, 60, 12, MNsure_app_date
  EditBox 75, 19, 60, 12, worker_signature
  EditBox 224, 19, 60, 12, integrated_case_number
  EditBox 56, 34, 228, 12, worker_notes
  ButtonGroup ButtonPressed
    OkButton 231, 50, 20, 12
    CancelButton 254, 50, 30, 12
  Text 5, 5, 70, 10, "MAXIS case number:"
  Text 140, 5, 60, 10, "MNsure app date:"
  Text 5, 20, 60, 10, "Worker signature:"
  Text 140, 20, 81, 10, "Integrated Case Number:"
  Text 5, 35, 46, 10, "Worker notes:"
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
      Dialog MNsure_2013_Coverage_Retro_Dialog
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
EMSendKey "MNsure Task: Retro Determination Required" & "<newline>"
If MNsure_app_date <> "" then call write_editbox_in_case_note("MNsure application date", MNsure_app_date, 6)
If integrated_case_number <> "" then call write_editbox_in_case_note("Integrated case number", integrated_case_number, 6)
call write_new_line_in_case_note("-Sent out DHS-6696A")
call write_new_line_in_case_note("-Sent out DHS-3271")
call write_new_line_in_case_note("-Requested pay stub verification for the months in which applicant may be able to receive coverage.")
call write_new_line_in_case_note("-Updated APPL and case is in PND2 status, case will pend for supplement to MNsure.")
call write_new_line_in_case_note("-Updated TYPE, HCRE, and PROG Panels; MSURE panel updated;")
call write_new_line_in_case_note("-Closed: Retro Determination Task.")
call write_new_line_in_case_note("-Create Task: 6696A Due")
call write_editbox_in_case_note("Please note", "If these docs come into your ''My documents received'' queue in OnBase, please create a copy of the document and re-index it to the appropriate MNsure doc type, and send to the proper workflow. If you have questions, consult a member of the MNsure team.", 6)
call write_new_line_in_case_note("**********")
If worker_notes <> "" then call write_editbox_in_case_note("Worker notes", worker_notes, 6)
If worker_notes <> "" then call write_new_line_in_case_note("**********")
call write_new_line_in_case_note(worker_signature)

script_end_procedure("")
