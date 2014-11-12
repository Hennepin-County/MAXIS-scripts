'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - client contact (call center)"
start_time = timer

'FUNCTIONS----------------------------------------------------------------------------------------------------
'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script


'THE DIALOG----------------------------------------------------------------------------------------------------
BeginDialog contact_dialog, 0, 0, 386, 175, "Client contact"
  ComboBox 65, 5, 85, 15, "client"+chr(9)+"AREP"+chr(9)+"non-AREP"+chr(9)+"SWKR", who_contacted
  EditBox 180, 5, 200, 15, regarding
  EditBox 85, 35, 60, 15, phone_number
  EditBox 255, 35, 115, 15, when_contact_was_made
  EditBox 55, 60, 85, 15, case_number
  EditBox 55, 80, 325, 15, issue
  CheckBox 10, 100, 95, 10, "Answered question", answered_question_check
  CheckBox 110, 100, 95, 10, "Transferred question", transferred_question_check
  CheckBox 200, 100, 150, 10, "Reminded Client re: Importance of CAF I", caf_1_check
  EditBox 55, 115, 325, 15, other_action
  EditBox 310, 135, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 270, 155, 50, 15
    CancelButton 330, 155, 50, 15
  Text 5, 10, 55, 10, "Phone call from:"
  Text 160, 10, 15, 10, "Re:"
  GroupBox 5, 25, 370, 30, "Optional info:"
  Text 30, 40, 50, 10, "Phone number: "
  Text 165, 40, 85, 10, "When was contact made: "
  Text 5, 65, 50, 10, "Case number: "
  Text 5, 85, 50, 10, "Issue/subject: "
  Text 5, 120, 45, 10, "Other action: "
  Text 235, 140, 70, 10, "Sign your case note: "
EndDialog
