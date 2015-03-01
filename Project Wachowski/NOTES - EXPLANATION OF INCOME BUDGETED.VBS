'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - Explanation Of Income Budgeted"
start_time = timer

''LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script





'Check for MAXIS
'WE USE THIS PIECE OF CODE A LOT
MAXIS_check_function





'DIALOG---------------------------------------------------------------

'Explanation Of Income Budgeted Dialog
BeginDialog explanation_of_income_budgeted_dialog, 0, 0, 246, 225, "Explanation Of Income Budgeted Dialog"
  EditBox 55, 5, 65, 15, case_number
  EditBox 120, 25, 125, 15, explanation_of_income_budgeted
  EditBox 120, 45, 125, 15, income_not_budgeted_and_why
  DropListBox 95, 65, 150, 15, "Paystubs"+chr(9)+"Employer's Statement"+chr(9)+"Other (Specified Under 'Other Notes')", type_of_verification_used
  DropListBox 70, 85, 175, 15, "Most Recent 30 Days"+chr(9)+"30 Days Prior To Date Of Application/Recertification"+chr(9)+"Retrospective Month (MFIP)"+chr(9)+"Other (Specified Under 'Other Notes')", time_period_used
  EditBox 85, 105, 160, 15, other_notes_on_income
  CheckBox 5, 130, 250, 10, "Agency/client are ''reasonably certain'' budgeted income", reasonably_certain_check
  CheckBox 5, 155, 195, 10, "Verification/s match what client reported they anticipate.", verifications_match_check
  EditBox 75, 175, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 135, 200, 50, 15
    CancelButton 190, 200, 50, 15
  Text 5, 5, 50, 10, "Case Number:"
  Text 5, 25, 110, 10, "Explanation Of Income Budgeted:"
  Text 5, 45, 110, 10, "Income NOT Budgeted and Why:"
  Text 5, 65, 85, 10, "Type Of Verification Used:"
  Text 5, 85, 65, 10, "Time Period Used:"
  Text 5, 105, 80, 10, "Other Notes On Income:"
  Text 10, 140, 170, 10, "will continue during the certification period."
  Text 5, 175, 75, 15, "Sign your case note:"
EndDialog



'THE SCRIPT----------------------------------------------------------------------

'Connect to BlueZone
EMConnect ""


'Check for MAXIS
'WE USE THIS PIECE OF CODE A LOT
MAXIS_check_function



'Does the dialog
Dialog explanation_of_income_budgeted_dialog
If ButtonPressed = 0 then StopScript



'Navigates to CASE/NOTE for the case number entered.
call navigate_to_screen("CASE", "NOTE")



'PF9 to open new case note
PF9



'Writes to the CASE NOTE
call write_new_line_in_case_note(">>>Explanation Of Income Budgeted<<<")
If explanation_of_income_budgeted <> "" then call write_editbox_in_case_note("Explanation Of Income Budgeted", explanation_of_income_budgeted, 6)
If income_not_budgeted_and_why <> "" then call write_editbox_in_case_note("Income Not Budgeted And Why", income_not_budgeted_and_why, 6)
call write_editbox_in_case_note("Type Of Verification Used", type_of_verification_used, 6)
call write_editbox_in_case_note("Time Period Used", time_period_used, 6)
If explanation_of_income_budgeted <> "" then call write_editbox_in_case_note("Other Notes On Income", other_notes_on_income, 6)
If reasonably_certain_check = 1 then 
	call write_new_line_in_case_note("* Agency/client are 'reasonably certain' budgeted income will continue during")
	call write_new_line_in_case_note("      the certification period.")
	end if
If verifications_match_check = 1 then call write_new_line_in_case_note("* Verification/s match what client reported they anticipate.")
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

