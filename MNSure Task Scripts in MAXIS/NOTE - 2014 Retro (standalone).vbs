'Informational front-end message, date dependent.
If datediff("d", "04/02/2013", now) < 5 then MsgBox "This script has been updated as of 04/02/2013! There's now a checkbox for starting the denied programs script right from this one."

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "SCRIPT NAME HERE"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("M:\Income-Maintence-Share\Bluezone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

Set MNSURE_FUNCTIONS_fso = CreateObject("Scripting.FileSystemObject")
Set fso_MNSURE_FUNCTIONS_command = MNSURE_FUNCTIONS_fso.OpenTextFile("M:\Income-Maintence-Share\Bluezone Scripts\Script Files\MNSure\MNSURE FUNCTIONS FILE.vbs")
MNSURE_FUNCTIONS_contents = fso_MNSURE_FUNCTIONS_command.ReadAll
fso_MNSURE_FUNCTIONS_command.Close
Execute MNSURE_FUNCTIONS_contents 

'DIALOGS----------------------------------------------------------------------------------------------------

BeginDialog case_note_decision_dialog, 0, 0, 230, 187, "Review and Approve Results"
  EditBox 98, 49, 75, 12, case_number
  EditBox 98, 64, 75, 12, integrated_case_number
  DropListBox 99, 79, 75, 13, "Approved"+chr(9)+"Denied"+chr(9)+"Pending"+chr(9)+"Withdrawn", case_note_status
  DropListBox 99, 95, 75, 13, "RFI Received"+chr(9)+"Phone Call"+chr(9)+"Not Received", retro_received_via
  DropListBox 99, 111, 75, 13, "1"+chr(9)+"2"+chr(9)+"3", number_of_retro_months
  EditBox 98, 127, 75, 12, case_note_worker_signature
  EditBox 10, 152, 210, 12, case_note_additional_comments
  ButtonGroup ButtonPressed
    PushButton 117, 174, 55, 12, "Refresh Dialog", refresh_dialog
    OkButton 175, 174, 20, 12
    CancelButton 198, 174, 30, 12
  Text 3, 3, 224, 25, "Please review the results and make any changes/corrections as necessary. Once complete, approve your results and you may select below to automatically case note."
  GroupBox 5, 35, 220, 136, "Case Note"
  Text 13, 50, 48, 8, "Case Number:"
  Text 13, 65, 83, 8, "Integrated Case Number:"
  Text 13, 81, 42, 8, "Case Status:"
  Text 13, 97, 70, 8, "Retro Received Via:"
  Text 13, 113, 84, 8, "Retro Months Requested:"
  Text 13, 128, 62, 8, "Worker Signature :"
  Text 75, 142, 78, 8, "Additional worker notes"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------

EMConnect ""

'Finds the case number in a case
call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""

Dialog case_note_decision_dialog
If buttonpressed = 0 then stopscript
If buttonpressed = refresh_dialog then
Do
	Dialog case_note_decision_dialog
		If buttonpressed = 0 then stopscript
Loop until buttonpressed = -1
End If

call navigate_to_screen("STAT","HCRE")

EMReadScreen appl_month, 2, 10, 51
EMReadScreen appl_day, 2, 10, 54
EMReadScreen appl_year, 2, 10, 57

EMReadScreen info_received_month, 2, 10, 73
EMReadScreen info_received_day, 2, 10, 76
EMReadScreen info_received_year, 2, 10, 79

Do
	call navigate_to_screen("case", "note")
	PF9
	EMReadScreen mode_check, 7, 20, 3
	If mode_check <> "Mode: A" and mode_check <> "Mode: E" then MsgBox "For some reason, the script can't get to a case note. Did you start the script in inquiry by mistake? Navigate to MAXIS production, or shut down the script and try again."
Loop until mode_check = "Mode: A" or mode_check = "Mode: E"
    
EMSendKey "MNSURE: 2014 Retro " & case_note_status & "<newline>"
If integrated_case_number <> "" then call write_new_line_in_case_note("-Integrated case # "&integrated_case_number)
If case_number <> "" then call write_new_line_in_case_note("-MAXIS case # "&case_number)
If appl_day <> "" and appl_month <> "" and appl_year <> "" then call write_new_line_in_case_note("-MNsure application date: "&appl_month&"/"&appl_day&"/"&appl_year)
If info_received_day <> "" and info_received_month <> "" and info_received_year <> "" and retro_received_via <> "Not Received" then call write_new_line_in_case_note("-Retro Info recvd via "&retro_received_via&" on "&info_received_month&"/"&info_received_day&"/"&info_received_year&".")
If retro_received_via = "Not Received" then call write_new_line_in_case_note("-Retro Info was not recvd. Case has been " & case_note_status)
call write_new_line_in_case_note("-Retro months requested: "&number_of_retro_months)
call write_new_line_in_case_note("**********")
If case_note_additional_comments <> "" then call write_editbox_in_case_note("Worker notes", case_note_additional_comments, 6)
If case_note_additional_comments <> "" then call write_new_line_in_case_note("**********")
call write_new_line_in_case_note(case_note_worker_signature)

EMReadScreen case_note_edit_test, 5, 20, 3

If case_note_edit_test = "Mode:" then
	MsgBox "You have completed this task. Please make any changes necessary and upon pressing OK you will be taken to this cases dails for clean-up."
	call navigate_to_screen("dail","dail")
End If

script_end_procedure("")