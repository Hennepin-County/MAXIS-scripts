'STATS----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - PA VERIF REQUEST.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, FALSE									'Attempts to open the URL
req.send													'Sends request
IF req.Status = 200 THEN									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			script_end_procedure("Script ended due to error connecting to GitHub.")
END IF

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

BeginDialog PA_verif_dialog, 0, 0, 190, 250, "PA Verif Dialog"
  ButtonGroup ButtonPressed
    OkButton 85, 230, 50, 15
    CancelButton 140, 230, 50, 15
 
  EditBox 50, 15, 25, 15, snap_grant
  EditBox 125, 15, 25, 15, MFIP_food
  EditBox 155, 15, 25, 15, MFIP_cash  
  EditBox 50, 35, 25, 15, MSA_Grant
  EditBox 155, 35, 25, 15, GA_grant
  EditBox 155, 55, 25, 15, DWP_grant
  EditBox 50, 75, 130, 15, other_notes
  EditBox 50, 100, 130, 15, other_income
  CheckBox 50, 120, 35, 10, "Yes", subsidy_check
  EditBox 50, 140, 20, 15, cash_members
  EditBox 150, 140, 20, 15, household_members
  CheckBox 10, 170, 200, 10, "Include screenshot of last 3 months' benefits", inqd_check
  EditBox 40, 190, 55, 15, completed_by
  EditBox 140, 190, 45, 15, worker_phone
  EditBox 120, 210, 65, 15, worker_signature
  
  Text 5, 15, 40, 15, "SNAP:"
  Text 100, 55, 20, 15, "DWP:"
  Text 5, 75, 40, 15, "Other notes:"
  Text 100, 35, 35, 15, "GA:"
  Text 5, 35, 35, 15, "MSA:"
  Text 100, 15, 25, 15, "MFIP:"
  Text 5, 100, 45, 20, "Other income and type:"
  Text 5, 120, 45, 20, "$50 subsidy deduction?"
  Text 5, 140, 45, 30, "Number of members on cash grant:"
  Text 90, 140, 55, 25, "Total members in household:"
  Text 130, 5, 25, 10, "Food:"
  Text 160, 5, 25, 10, "Cash:"
  Text 110, 190, 25, 20, "Worker Phone:"
  Text 5, 190, 35, 20, "Completed by:"
  Text 20, 210, 90, 15, "Worker Signature (For case note):"
 
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
EMReadScreen addr_city, 14, 8, 43
EMReadScreen addr_state, 2, 8, 66
EMReadScreen addr_zip, 5, 9, 43
hh_address = addr_line1 & " " & addr_line2 'Finding and Formatting household address 
hh_address_line2 = addr_city & " " & addr_state & " " & addr_zip
hh_address = replace(hh_address, "_", "") & vbCrLf & replace(hh_address_line2, "_", "")


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

'Pulling the elig amounts for all open progs on case / curr
call navigate_to_screen("case", "curr")
 call find_variable("MFIP: ", MFIP_check, 6)
   If MFIP_check = "ACTIVE" OR MFIP_check = "APP CL" then
   call navigate_to_screen("elig", "mfip")    
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
	If MFIP_check = "PENDIN" then msgbox "MFIP is pending, please enter amounts manually to avoid errors."

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
		call find_variable("Caregivers....", caregivers, 5)
		call find_variable("Children......", children, 5)
		cash_members = cInt(caregivers) + cInt(children)
		cash_members = cStr(cash_members)
		call navigate_to_screen ("case", "curr")
	 End if
	If DWP_check = "PENDIN" then msgbox "DWP is pending, please enter amounts manually to avoid errors."
	
	call find_variable("GA: ", GA_check, 6)
	If GA_check = "ACTIVE" then
		call navigate_to_screen("elig", "GA")
		call approved_version
		EMWriteScreen version, 20, 78
		transmit
		EMWriteScreen "GASM", 20, 70
		transmit
		EMReadScreen GA_grant, 7, 9, 73
	    EMReadScreen ga_members, 1, 13, 32 'Reading file unit type to determine members on cash grant
		If ga_members = "1" then cash_members = "1"
		If ga_members = "6" then cash_members = "2"
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
	
	call find_variable("Cash: ", cash_check, 6)
	If cash_check = "PENDIN" then MsgBox "Cash is pending for this household, please explain in additional notes."
		
'calling the main dialog	
Do
	Dialog PA_verif_dialog
	If ButtonPressed = 0 then stopscript
	If worker_signature = ""  then MsgBox "Please sign your case note."
	If completed_by = "" then MsgBox "Please fill out the completed by field."
	If worker_phone = "" then MsgBox "Please fill out the worker phone field."
Loop until worker_signature <> "" and completed_by <> "" and worker_phone <> ""



  
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
ObjSelection.TypeText "Other Notes: "
objSelection.TypeText other_notes
objSelection.TypeParagraph()

'Writing INQD to the doc if selected
IF inqd_check = checked THEN
	objSelection.TypeText "Benefits Issued for last 3 months:"
	objSelection.TypeParagraph()
	objSelection.TypeText "Issue Date	    Benefit               Amount                            Benefit Period"
	objSelection.TypeParagraph()
	call navigate_to_screen("MONY", "INQD")
	output_array = "" 'resetting array
	Dim screenarray(12)	'12 line array (leaves out the header and function info)
	row = 6
	For each line in screenarray
		EMReadScreen reading_line, 80, row, 1
		output_array = output_array & reading_line & "UUDDLRLRBA"
		row = row + 1
	Next
	output_array = split(output_array, "UUDDLRLRBA")
	FOR EACH line in output_array
		IF line <> "                                                                                " THEN
			objSelection.TypeText line & Chr(11)
		END IF
	NEXT
END IF
		
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
call write_new_line_in_case_note("PA verification request completed and sent to requesting agency.")
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

'Starts the print dialog
objword.dialogs(wdDialogFilePrint).Show



