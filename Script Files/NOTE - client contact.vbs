'Informational front-end message, date dependent.
If datediff("d", "05/15/2014", now) < 2 then MsgBox "This script has been updated as of 05/15/2014! Here's what's new:" & chr(13) & chr(13) & "1. Issue/subject is now ''reason for contact''." & chr(13) & "2. A new ''helpful info for the call center'' section has gone in, with verifs needed, instructions for client, and case status sections. This will be helpful for the call center." & chr(13) & "3. ''Left generic message'' has been removed." & chr(13) & chr(13) & "If you have any issues running this script, please email Charles Potter or Robert Kalb."


'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - client contact"
start_time = timer

'FUNCTIONS----------------------------------------------------------------------------------------------------
'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script


'THE DIALOG

EMConnect ""

row = 1
col = 1

EMSearch "Case Nbr:", row, col
If row <> 0 then 
	EMReadScreen case_number, 8, row, col + 10
	case_number = replace(case_number, "_", "")
	case_number = trim(case_number)
End if

'<<<<REMOVED CHARLES AND ROBERT'S DIALOG, AS I HAD ALREADY CREATED ONE (MINE LINES UP WITH CS)
'BeginDialog contact_dialog, 0, 0, 386, 205, "Client contact"
'  DropListBox 50, 5, 60, 15, "Phone call"+chr(9)+"Voicemail"+chr(9)+"EZ info msg"+chr(9)+"Office visit"+chr(9)+"Letter", contact_type
'  DropListBox 115, 5, 45, 10, "from"+chr(9)+"to", contact_direction
'  ComboBox 165, 5, 85, 15, "client"+chr(9)+"AREP"+chr(9)+"SWKR", who_contacted
'  EditBox 280, 5, 100, 15, regarding
'  EditBox 95, 35, 60, 15, phone_number
'  EditBox 285, 35, 85, 15, when_voicemail_was_left
'  EditBox 55, 60, 85, 15, case_number
'  EditBox 75, 80, 305, 15, contact_reason
'  EditBox 55, 100, 325, 15, actions_taken
'  EditBox 105, 120, 275, 15, cl_instructions
'  EditBox 55, 140, 325, 15, verifs_needed
'  EditBox 80, 160, 70, 15, worker_signature
'  CheckBox 5, 185, 255, 10, "Check here if you want to TIKL out for this case after the case note is done.", TIKL_check
'  ButtonGroup ButtonPressed
'    OkButton 270, 180, 50, 15
'    CancelButton 330, 180, 50, 15
'  Text 5, 85, 65, 10, "Reason for Contact:"
'  Text 40, 40, 50, 10, "Phone number: "
'  Text 5, 105, 50, 10, "Actions taken: "
'  Text 260, 10, 15, 10, "Re:"
'  Text 5, 165, 70, 10, "Sign your case note: "
'  Text 5, 10, 45, 10, "Contact type:"
'  Text 5, 65, 50, 10, "Case number: "
'  GroupBox 5, 25, 370, 30, "Optional info:"
'  Text 145, 60, 235, 15, "NOTE: YOU CAN NOW PUT A MCRE NUMBER IN HERE IF YOU'RE CASE NOTING IN MMIS."
'  Text 5, 125, 100, 10, "Instructions/Message for CL:"
'  Text 195, 40, 85, 10, "When was contact made: "
'  Text 5, 145, 50, 10, "Verifs Needed: "
'EndDialog


BeginDialog contact_dialog, 0, 0, 386, 255, "Client contact"
  DropListBox 50, 5, 60, 15, "Phone call"+chr(9)+"Voicemail"+chr(9)+"EZ info msg"+chr(9)+"Office visit"+chr(9)+"Letter", contact_type
  DropListBox 115, 5, 45, 10, "from"+chr(9)+"to", contact_direction
  ComboBox 165, 5, 85, 15, "client"+chr(9)+"AREP"+chr(9)+"SWKR", who_contacted
  EditBox 280, 5, 100, 15, regarding
  EditBox 95, 35, 60, 15, phone_number
  EditBox 285, 35, 85, 15, when_voicemail_was_left
  EditBox 55, 60, 85, 15, case_number
  EditBox 70, 80, 310, 15, contact_reason
  EditBox 55, 100, 325, 15, actions_taken
  EditBox 65, 135, 310, 15, verifs_needed
  EditBox 125, 155, 250, 15, cl_instructions
  EditBox 65, 175, 310, 15, case_status
  CheckBox 5, 200, 255, 10, "Check here if you want to TIKL out for this case after the case note is done.", TIKL_check
  EditBox 310, 215, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 270, 235, 50, 15
    CancelButton 330, 235, 50, 15
  Text 5, 10, 45, 10, "Contact type:"
  Text 260, 10, 15, 10, "Re:"
  GroupBox 5, 25, 370, 30, "Optional info:"
  Text 40, 40, 50, 10, "Phone number: "
  Text 195, 40, 85, 10, "When was contact made: "
  Text 5, 65, 50, 10, "Case number: "
  Text 5, 85, 65, 10, "Reason for contact:"
  Text 5, 105, 50, 10, "Actions taken: "
  GroupBox 5, 120, 375, 75, "Helpful info for the call center/front desk to pass on to clients"
  Text 15, 140, 50, 10, "Verifs needed: "
  Text 15, 160, 105, 10, "Instructions/message for client:"
  Text 15, 180, 45, 10, "Case status: "
  Text 235, 220, 70, 10, "Sign your case note: "
EndDialog





DO
  Do
    Do
     Do
      Dialog contact_dialog
      If buttonpressed = 0 then stopscript
	IF contact_reason = "" THEN MSGBox("You must enter a reason for CL contact.")
     Loop until contact_reason <> ""
      If (isnumeric(case_number) = False and len(case_number) <> 8) then MsgBox "You must enter either a valid MAXIS or MCRE case number."
    Loop until (isnumeric(case_number) = True) or (isnumeric(case_number) = False and len(case_number) = 8)
    transmit
    If isnumeric(case_number) = True then
      EMReadScreen MAXIS_check, 5, 1, 39
      If MAXIS_check <> "MAXIS" then MsgBox "You are not in MAXIS. Navigate your screen to MAXIS and try again. You might be passworded out."
    Else
      MMIS_row = 1
      MMIS_col = 1
      EMSearch "MMIS", MMIS_row, MMIS_col
      If MMIS_row <> 1 then
        EMReadScreen OSLT_check, 4, 1, 52 'Because cases that are on the "OSLT" screen in MMIS don't contain the characters "MMIS" in the top line.
        If OSLT_check = "OSLT" then MMIS_row = 1
      End if
      If MMIS_row <> 1 then MsgBox "You are not in MMIS. Navigate your screen to MMIS and try again. You might be passworded out."
    End if
  Loop until MAXIS_check = "MAXIS" or MMIS_row = 1
  If isnumeric(case_number) = True then
    call navigate_to_screen("case", "note")
    PF9
    EMReadScreen mode_check, 7, 20, 3
    If mode_check <> "Mode: A" and mode_check <> "Mode: E" then MsgBox "The script doesn't appear to be able to find your case note. Are you in inquiry? If so, navigate to production on the screen where you clicked the script button, and try again. Otherwise, you might have forgotten to type a valid case number."
  Else
    call MMIS_RKEY_finder
    EMWriteScreen "c", 2, 19
    EMWriteScreen case_number, 9, 19
    transmit
    EMReadScreen RKEY_check, 4, 1, 52 'CHECKING FOR RKEY, IF RENEWAL IS DUE A WARNING MESSAGE WILL NEED TO BE MOVED PAST.
    If RKEY_check = "RKEY" then transmit
    PF4
    PF11
    EMReadScreen MMIS_edit_mode_check, 5, 5, 2
    If MMIS_edit_mode_check <> "'''''" then script_end_procedure("MMIS edit mode not found. Are you in inquiry? Is MMIS not functioning? Shut down this script and try again. If it continues to not work, email your script administrator the case number, and a screenshot of MMIS.")
  End if
Loop until (mode_check = "Mode: A" or mode_check = "Mode: E") or (MMIS_edit_mode_check = "'''''") 

If isnumeric(case_number) = True then
  EMSendKey contact_type & " " & contact_direction & " " & who_contacted
  If regarding <> "" then EMSendKey " re: " & regarding 
  EMSendKey "<newline>"
  If when_voicemail_was_left <> "" then Call write_editbox_in_case_note("Contact made", when_voicemail_was_left, 6)
  If phone_number <> "" then Call write_editbox_in_case_note("Phone number", phone_number, 6)
  If contact_reason <> "" then Call write_editbox_in_case_note("Reason for Contact", contact_reason, 6)
  If actions_taken <> "" then Call write_editbox_in_case_note("Actions taken", actions_taken, 6)
  IF verifs_needed <> "" then Call write_editbox_in_case_note("Verifs Needed", verifs_needed, 6)
  If cl_instructions <> "" then Call write_editbox_in_case_note("Instructions/Message for CL", cl_instructions, 6)
  If case_status <> "" then Call write_editbox_in_case_note("Case status", case_status, 6)
  Call write_new_line_in_case_note("---")
  Call write_new_line_in_case_note(worker_signature)
  If TIKL_check = 0 then script_end_procedure("")

  'TIKLING
  MsgBox "The script will now navigate to a TIKL."
  call navigate_to_screen("dail", "writ")
Else
  EMSendKey contact_type & " " & contact_direction & " " & who_contacted
  If regarding <> "" then EMSendKey " re: " & regarding 
  PF11
  If when_voicemail_was_left <> "" then EMSendKey "* Contact made: " & when_voicemail_was_left
  PF11
  If phone_number <> "" then EMSendKey "* Phone number: " & phone_number
  PF11
  If issue <> "" then EMSendKey "* Reason for Contact: " & issue
  PF11
  If actions_taken <> "" then EMSendKey "* Actions taken: " & actions_taken
  PF11
  If cl_instructions <> "" then EMSendKey "* Instructions/Message for CL: " & cl_instructions
  PF11
  If verifs_needed <> "" then EMSendKey "* Verifs Needed: " & verifs_needed
  PF11
  EMSendKey "---"
  PF11
  EMSendKey worker_signature
  PF11
  EMSendKey "************************************************************************"
  If TIKL_check = 1 then script_end_procedure("Unable to TIKL for MCRE case. Find the MAXIS case and TIKL manually.")
End if

script_end_procedure("")






