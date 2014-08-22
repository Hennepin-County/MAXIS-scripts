'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - ApplyMN received"
start_time = timer

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS----------------------------------------------------------------------------------------------------

BeginDialog apply_MN_dialog, 0, 0, 291, 125, "Apply MN"
  EditBox 60, 5, 75, 15, case_number
  EditBox 90, 25, 75, 15, app_date
  EditBox 185, 25, 40, 15, app_time
  DropListBox 235, 25, 35, 15, "AM"+chr(9)+"PM", AM_PM
  EditBox 65, 45, 100, 15, confirmation_number
  EditBox 50, 65, 205, 15, progs_applied_for
  DropListBox 55, 85, 80, 15, "N/A"+chr(9)+"known to EBT"+chr(9)+"unknown to EBT", EBT_status
  DropListBox 180, 85, 105, 15, "SPEC/XFERed to worker."+chr(9)+"Indexed to worker.", actions_taken
  EditBox 70, 105, 75, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 175, 105, 50, 15
    CancelButton 235, 105, 50, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 30, 80, 10, "Apply MN app rec'd on"
  Text 170, 30, 10, 10, "at"
  Text 5, 50, 55, 10, "Confirmation #:"
  Text 5, 70, 45, 10, "Applying for:"
  Text 5, 90, 40, 10, "EBT status:"
  Text 145, 90, 30, 10, "Actions:"
  Text 5, 110, 60, 10, "Worker signature:"
EndDialog




'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

'Finds case number
call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If isnumeric(case_number) = False then case_number = ""

'Shows dialog and navigates to case note
Do
  Do
    Dialog apply_MN_dialog
    If ButtonPressed = 0 then stopscript
    transmit
    EMReadScreen MAXIS_check, 5, 1, 39
    If MAXIS_check <> "MAXIS" then MsgBox "You do not appear to be in MAXIS on the screen you started this script. Are you passworded out? Press OK, and navigate to MAXIS before proceeding."
  Loop until MAXIS_check = "MAXIS"
  call navigate_to_screen("case", "note")
  PF9
  EMReadScreen mode_check, 7, 20, 3
  If mode_check <> "Mode: A" and mode_check <> "Mode: E" then MsgBox "The script could not navigate to a case note on edit mode. Are you in inquiry? Or an old case note? Perhaps the case is out of county? Try backing out of case note and trying again."
Loop until mode_check = "Mode: A" or mode_check = "Mode: E"

'Case notes information
EMSendKey "ApplyMN app rec'd on " & app_date & " at " & app_time & " " & AM_PM & "<newline>"
call write_editbox_in_case_note("Confirmation #", confirmation_number, 6) 'x is the header, y is the variable for the edit box which will be put in the case note.
call write_editbox_in_case_note("Applying for", progs_applied_for, 6) 'x is the header, y is the variable for the edit box which will be put in the case note.
If EBT_status <> "N/A" then call write_new_line_in_case_note("* Client is " & EBT_status & ".")
call write_new_line_in_case_note("* " & actions_taken)
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

script_end_procedure("")






