'STATS----------------------------------------------------------------------------------------------------
name_of_script = "MEMO - PA Verif Request"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
next_month = dateadd("m", + 1, date)
footer_month = datepart("m", next_month)
If len(footer_month) = 1 then footer_month = "0" & footer_month
footer_year = datepart("yyyy", next_month)
footer_year = "" & footer_year - 2000


'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 191, 75, "PA Verification Request"
  ButtonGroup ButtonPressed
    OkButton 75, 50, 50, 15
    CancelButton 130, 50, 50, 15
  EditBox 105, 10, 70, 15, case_number
  EditBox 105, 30, 20, 15, footer_month
  EditBox 130, 30, 20, 15, footer_year
  Text 30, 10, 50, 15, "Case Number"
  Text 30, 30, 60, 15, "Footer Month"
EndDialog

BeginDialog PA_verif_dialog, 0, 0, 236, 230, "PA Verif Dialog"
  EditBox 55, 25, 25, 15, snap_grant
  EditBox 55, 45, 25, 15, MSA_Grant
  EditBox 55, 65, 25, 15, GA_grant
  EditBox 160, 25, 20, 15, MFIP_food
  EditBox 190, 25, 20, 15, MFIP_cash
  EditBox 160, 45, 20, 15, relative_food
  EditBox 190, 45, 20, 15, relative_cash
  EditBox 190, 65, 20, 15, foster_care
  EditBox 55, 90, 175, 15, other_income
  CheckBox 55, 110, 35, 10, "Yes", subsidy_check
  EditBox 75, 135, 20, 15, cash_members
  EditBox 185, 135, 20, 15, household_members
  EditBox 60, 160, 50, 15, completed_by
  EditBox 185, 160, 45, 15, worker_phone
  EditBox 165, 185, 65, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 125, 210, 50, 15
    CancelButton 180, 210, 50, 15
  GroupBox 20, 5, 195, 80, "PA grant info:"
  Text 160, 15, 25, 10, "Food"
  Text 190, 15, 25, 10, "Cash"
  Text 30, 30, 25, 10, "SNAP:"
  Text 30, 50, 20, 10, "MSA:"
  Text 35, 70, 15, 10, "GA:"
  Text 135, 30, 25, 10, "MFIP:"
  Text 110, 50, 50, 10, "Relative Care:"
  Text 115, 70, 40, 10, "Foster Care:"
  Text 5, 90, 45, 20, "Other income and type"
  Text 5, 110, 45, 20, "$50 subsidy deduction?"
  Text 5, 135, 70, 20, "Number of members on cash grant:"
  Text 5, 165, 50, 10, "Completed by:"
  Text 125, 135, 55, 20, "Total members in household:"
  Text 130, 165, 50, 10, "Worker phone:"
  Text 100, 190, 60, 10, "Worker Signature:"
EndDialog


BeginDialog cancel_dialog, 0, 0, 141, 51, "Cancel dialog"
  Text 5, 5, 135, 10, "Are you sure you want to end this script?"
  ButtonGroup ButtonPressed
    PushButton 10, 20, 125, 10, "No, take me back to the script dialog.", no_cancel_button
    PushButton 20, 35, 105, 10, "Yes, close this script.", yes_cancel_button
EndDialog

'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
HH_memb_row = 5
Dim row
Dim col

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connecting to BlueZone
EMConnect ""

'Grabbing case number
call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""

'Grabbing footer month
call find_variable("Month: ", MAXIS_footer_month, 2)
If row <> 0 then 
  footer_month = MAXIS_footer_month
  call find_variable("Month: " & footer_month & " ", MAXIS_footer_year, 2)
  If row <> 0 then footer_year = MAXIS_footer_year
End if

'Showing case number dialog
Do
  Dialog case_number_dialog
  If ButtonPressed = 0 then stopscript
  If case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 then MsgBox "You need to type a valid case number."
Loop until case_number <> "" and IsNumeric(case_number) = True and len(case_number) <= 8

'Checking for MAXIS
transmit
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then call script_end_procedure("You are not in MAXIS or you are locked out of your case.")

'Jumping to STAT
call navigate_to_screen("stat", "memb")
EMReadScreen STAT_check, 4, 20, 21
If STAT_check <> "STAT" then call script_end_procedure("Can't get in to STAT. This case may be in background. Wait a few seconds and try again. If the case is not in background email your script administrator the case number and footer month.")
EMReadScreen ERRR_check, 4, 2, 52
If ERRR_check = "ERRR" then transmit 'For error prone cases.

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)


'Pulling household and worker info for the letter
call navigate_to_screen("stat", "addr") 
EMReadScreen addr_line1, 21, 6, 43
EMReadScreen addr_line2, 21, 7, 43
hh_address = addr_line1 & " " & addr_line2 'Finding and Formatting household address (
hh_address = replace(hh_address, "_", "")
household_members = UBound(HH_member_array) + 1 'Total members in household
household_members = cStr(household_members)

'Collecting and formatting client name
call navigate_to_screen("stat", "memb")
call find_variable("Last: ", last_name, 24)
call find_variable("First: ", first_name, 11)
client_name = first_name & " " & last_name
client_name = replace(client_name, "_", "")




'Autofilling info for case note

call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)

'Cleaning up info for case note

earned_income = trim(earned_income)
if right(earned_income, 1) = ";" then earned_income = left(earned_income, len(earned_income) - 1)
other_income = earned_income & " " & unearned_income

'This function looks for an approved version of elig
Function approved_version 
		EMReadScreen version, 2, 2, 12
		For approved = version to 0 Step -1
		EMReadScreen approved_check, 8, 3, 3
		If approved_check = "APPROVED" then Exit Function
		version = version -1
		EMWriteScreen version, 20, 79
		transmit
		Next
	End Function
'This finds the number of members on a DWP/MFIP grant
Function cash_members_finder
		call find_variable("Caregivers......", caregivers, 4)
		call find_variable("Children........", children, 4)
		cash_members = cInt(caregivers) + cInt(children)
		cash_members = cStr(cash_members)
		
End Function


'Pulling the elig amounts.
call navigate_to_screen("case", "curr")
row = 7
EMSearch "MFIP", row, col
    If row <> 0 then
      call navigate_to_screen("elig", "mfip") 
      EMReadScreen MFPR_check, 4, 3, 47
      If MFPR_check <> "MFPR" then MsgBox "no mfip results" 'need to determine this
        'Needs readscreen here for "approved" and logic to check for earlier approved version
	  call approved_version
		EMWriteScreen version, 20, 79
		transmit
        EMWriteScreen "MFSM", 20, 71
        transmit
        EMReadScreen MFIP_cash, 6, 15, 74
        EMReadScreen MFIP_food, 6, 16, 74
 		'rental subsidy check
		EMWriteScreen "MFB1", 20, 71
		EMReadScreen subsidy, 2, 17, 37
		If subsidy = "50" then subsidy_check = 1
		'add readscreen for sanction 
		'Finding the number of members on cash grant
		call cash_members_finder
		Call navigate_to_screen("case", "curr")
	End if
	If MFIP_check = "APP CL" then msgbox "MFIP is set to close, please enter amounts manually to avoid errors."
	If MFIP_check = "PENDIN" then msgbox "MFIP is pending, please enter amounts manually to avoid errors."
	row = 7

	call find_variable("FS: ", fs_check, 6)
	If fs_check = "ACTIVE" then
		call navigate_to_screen("elig", "fs")
		call approved_version
		EMWriteScreen version, 20, 78
		transmit
		EMWriteScreen "FSB2", 19, 70
		transmit
		EMReadScreen SNAP_grant, 7, 10, 75
	    call navigate_to_screen ("case", "curr")
	End if
	If fs_check = "APP CL" then msgbox "SNAP is set to close, please enter amounts manually to avoid errors."
	If fs_check = "PENDIN" then msgbox "SNAP is pending, please enter amounts manually to avoid errors."
	row = 7
	
	call find_variable("DWP: ", DWP_check, 6)
	If DWP_check = "ACTIVE" then
		call navigate_to_screen("elig", "dwp")
		call approved_version
		EMWriteScreen version, 20, 79
		transmit
		EMWriteScreen "DWB2", 20, 71
		transmit
		EMReadScreen DWP_grant, 7, 5, 37
	    EMWriteScreen "DWSM", 20, 71
		transmit
		call cash_members_finder
		call navigate_to_screen ("case", "curr")
	 End if
    If DWP_check = "APP CL" then msgbox "DWP is set to close, please enter amounts manually to avoid errors."
	If DWP_check = "PENDIN" then msgbox "DWP is pending, please enter amounts manually to avoid errors."
	row = 7
	
	call find_variable("GA: ", GA_check, 6)
	If GA_check = "ACTIVE" then
		call navigate_to_screen("elig", "GA")
		call approved_version
		EMWriteScreen version, 20, 78
		transmit
		EMWriteScreen "GAB2", 20, 70
		transmit
		EMReadScreen GA_grant, 7, 13, 75
	    EMReadScreen ga_members, 1, 13, 32
		If ga_members = 1 then cash_members = 1
		If ga_members = 6 then cash_members = 2
		call navigate_to_screen ("case", "curr")
	End If
	If GA_check = "APP CL" then msgbox "GA is set to close, please enter amounts manually to avoid errors."
	If GA_check = "PENDIN" then msgbox "GA is pending, please enter amounts manually to avoid errors."
 
	call find_variable("MSA: ", MSA_check, 6)
	If MSA_check = "ACTIVE" then
		call navigate_to_screen("elig", "msa")
		call approved_version
		EMWriteScreen version, 20, 79
		transmit
		EMWriteScreen "MSSM", 20, 71
		transmit
		EMReadScreen MSA_Grant, 7, 11, 74
		EMReadScreen cash_members, 1, 14, 29
		call navigate_to_screen ("case", "curr")
	End If
	If MSA_check = "APP CL" then MsgBox "MSA is set to close, please enter amounts manually to avoid errors."
	If MSA_check = "PENDIN" then MsgBox "MSA is pending, please enter amounts manually to avoid errors."

		
'calling the main dialog	
Do
	Dialog PA_verif_dialog
	If ButtonPressed = 0 then stopscript
  If worker_signature = ""  then MsgBox "Please sign your case note."
Loop until worker_signature <> "" 

	
	
    
	'****writing the word document

Set objWord = CreateObject("Word.Application")
Const wdDialogFilePrint = 88
Const end_of_doc = 6
objWord.Caption = "PA Verif Request"
objWord.Visible = True

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

objSelection.Font.Name = "Arial"
objSelection.Font.Size = "14"
objSelection.TypeText "Your agency requested information about public assistance from "  
objSelection.TypeText county_name 
objSelection.TypeText " for the following client:"
objSelection.TypeParagraph()
objSelection.TypeText client_name
objSelection.TypeParagraph()
objSelection.TypeText hh_address
objSelection.TypeParagraph()
objSelection.TypeText "The following grant amounts are active for this household:"

Set objRange = objSelection.Range
objDoc.Tables.Add objRange, 6, 3
set objTable = objDoc.Tables(1)

objTable.Cell(1, 2).Range.Text = "Cash  "
objTable.Cell(1, 3).Range.Text = "Food Portion"
objTable.Cell(2, 1).Range.Text = "MFIP  "
objTable.Cell(3, 1).Range.Text = "GA    "
objTable.Cell(4, 1).Range.Text = "MSA   "
objTable.Cell(5, 1).Range.Text = "SNAP  "
objTable.Cell(2, 2).Range.Text = MFIP_cash
objTable.Cell(2, 3).Range.Text = MFIP_food
objTable.Cell(3, 2).Range.Text = GA_grant
objTable.Cell(4, 2).Range.Text = MSA_Grant
objTable.Cell(5, 3).Range.Text = SNAP_grant
objTable.Cell(6, 1).Range.Text = "DWP   "
objTable.Cell(6, 2).Range.Text = DWP_grant

objTable.AutoFormat(16)

objSelection.EndKey end_of_doc
objSelection.TypeParagraph()

objSelection.TypeText "Other income known to agency: "
objSelection.TypeText other_income
objSelection.TypeParagraph()
objSelection.TypeText "Number of family members on cash grant: "
objSelection.TypeText cash_members
objSelection.TypeParagraph()
objSelection.TypeText "Number of persons in household: "
objSelection.TypeText household_members
objSelection.TypeParagraph()
objSelection.TypeText "Completed By: "
objSelection.TypeText completed_by
objSelection.TypeParagraph()
objSelection.TypeText "Worker phone: "
objSelection.TypeText worker_phone



Do	
	
	call navigate_to_screen("case", "note")
    PF9
    EMReadScreen case_note_check, 17, 2, 33
    EMReadScreen mode_check, 1, 20, 09
    If case_note_check <> "Case Notes (NOTE)" or mode_check <> "A" then MsgBox "The script can't open a case note. Are you in inquiry? Check MAXIS and try again."

Loop until case_note_check = "Case Notes (NOTE)" and mode_check = "A"

'Enters the case note
EMSendKey "<home>" & "PA verification request completed and sent to requesting agency." & "<newline>"
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)


'Starts the print dialog
objword.dialogs(wdDialogFilePrint).Show