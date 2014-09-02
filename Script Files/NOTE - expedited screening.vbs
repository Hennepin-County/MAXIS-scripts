'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - expedited screening"
start_time = timer


'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog exp_screening_dialog, 0, 0, 216, 155, "Expedited Screening Dialog"
  EditBox 55, 5, 80, 15, case_number
  EditBox 100, 25, 50, 15, income
  EditBox 100, 45, 50, 15, assets
  EditBox 115, 65, 50, 15, rent
  CheckBox 15, 95, 55, 10, "Heat (or AC)", heat_AC_check
  CheckBox 75, 95, 45, 10, "Electricity", electric_check
  CheckBox 130, 95, 35, 10, "Phone", phone_check
  DropListBox 70, 115, 120, 15, "intake"+chr(9)+"add-a-program", application_type
  EditBox 130, 135, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 160, 10, 50, 15
    CancelButton 160, 30, 50, 15
  Text 5, 10, 50, 10, "Case number: "
  Text 5, 30, 95, 10, "Income received this month:"
  Text 5, 50, 95, 10, "Cash, checking, or savings: "
  Text 5, 70, 105, 10, "Amounts paid for rent/mortgage:"
  GroupBox 5, 85, 170, 25, "Utilities claimed (check below):"
  Text 5, 120, 65, 10, "Application is for:"
  Text 55, 140, 70, 10, "Sign your case note:"
EndDialog

'DATE BASED LOGIC FOR UTILITY AMOUNTS------------------------------------------------------------------------------------------
If date >= cdate("10/01/2014") then			'these variables need to change in October 2014, and subsequently every October
	heat_AC_amt = 450
	electric_amt = 150
	phone_amt = 38
Else
	heat_AC_amt = 459
	electric_amt = 140
	phone_amt = 40
End if

'Connecting to BlueZone
EMConnect ""

'It will search for a case number.
call MAXIS_case_number_finder(case_number)

'Shows the dialog
Do
	Do
		Do
			Do
				Dialog exp_screening_dialog
				If ButtonPressed = 0 then stopscript
				If isnumeric(case_number) = False then MsgBox "You must enter a valid case number."
			Loop until isnumeric(case_number) = True
			If (income <> "" and isnumeric(income) = false) or (assets <> "" and isnumeric(assets) = false) or (rent <> "" and isnumeric(rent) = false) then MsgBox "The income/assets/rent fields must be numeric only. Do not put letters or symbols in these sections."
		Loop until (income = "" or isnumeric(income) = True) and (assets = "" or isnumeric(assets) = True) and(rent = "" or isnumeric(rent) = True)
		If worker_signature = "" then MsgBox "You must sign your case note."
	Loop until worker_signature <> ""
	transmit 'to check for MAXIS status
	EMReadScreen MAXIS_check, 5, 1, 39
	If MAXIS_check <> "MAXIS" then MsgBox "MAXIS is not found. You may need to enter a password. If you are in MAXIS, you may have had a configuration error. To fix this, restart BlueZone."
Loop until MAXIS_check = "MAXIS"

'Logic for figuring out utils. The highest priority for the if...then is heat/AC, followed by electric and phone, followed by phone and electric separately.
If heat_AC_check = checked then
	utilities = heat_AC_amt
ElseIf electric_check = checked and phone_check = checked then 
	utilities = phone_amt + electric_amt					'Phone standard plus electric standard.
ElseIf phone_check = checked and electric_check = unchecked then
	utilities = phone_amt
ElseIf electric_check = checked and phone_check = unchecked then
	utilities = electric_amt
End if

'in case no options are clicked, utilities are set to zero.
If phone_check = unchecked and electric_check = unchecked and heat_AC_check = unchecked then utilities = 0

'If nothing is written for income/assets/rent info, we set to zero.
If income = "" then income = 0
If assets = "" then assets = 0
If rent = "" then rent = 0


'Calculates expedited status based on above numbers
If (int(income) < 150 and int(assets) <= 100) or ((int(income) + int(assets)) < (int(rent) + cint(utilities))) then expedited_status = "client appears expedited"
If (int(income) + int(assets) >= int(rent) + cint(utilities)) and (int(income) >= 150 or int(assets) > 100) then expedited_status = "client does not appear expedited"

'SECTION 03
'This jumps back to SELF
back_to_SELF

'Navigates to STAT/DISQ using current month as footer month. If it can't get in to the current month due to CAF received in a different month, it'll find that month and navigate to it.
EMWriteScreen "stat", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
footer_month = datepart("m", date) 'This is so the footer month is correct for a CAF1 case.
If len(footer_month) = 1 then footer_month = "0" & footer_month
footer_year = datepart("yyyy", date) - 2000
EMWriteScreen footer_month, 20, 43
EMWriteScreen footer_year, 20, 46
EMWriteScreen "disq", 21, 70
transmit
EMReadScreen benefit_period_CAF1_check, 14, 24, 2
If benefit_period_CAF1_check = "BENEFIT PERIOD" then
	EMReadScreen footer_month, 2, 24, 62
	EMReadScreen footer_year, 2, 24, 65
	EMWriteScreen "stat", 16, 43
	EMWriteScreen footer_month, 20, 43
	EMWriteScreen footer_year, 20, 46
	EMWriteScreen "disq", 21, 70
	transmit
End if
EMReadScreen SELF_check, 4, 2, 50
If SELF_check = "SELF" then script_end_procedure("Can't get past SELF. There may be a glitch on this case. Check your case number and try again. You may need to restart BlueZone.")

'Reads the DISQ info for the case note.
EMReadScreen DISQ_member_check, 34, 24, 2
If DISQ_member_check = "DISQ DOES NOT EXIST FOR ANY MEMBER" then 
	has_DISQ = False
Else
	has_DISQ = True
End if

'Navigates to a blank case note
call navigate_to_screen("case", "note")
PF9
EMReadScreen read_only_check, 41, 24, 2
If read_only_check = "YOU HAVE 'READ ONLY' ACCESS FOR THIS CASE" then script_end_procedure("You have read-only access to this case! You may be in inquiry, or this may be out of county. Expedited status is indicated as: " & expedited_status & ". Try again or process/track manually.")

'Enters data into the case note
EMSendKey "<home>" 'To get to the top of the case note.
EMSendKey "Received " & application_type & ", " & expedited_status & "<newline>"
call write_new_line_in_case_note("---")
call write_new_line_in_case_note("     CAF 1 income claimed this month: $" & income)
call write_new_line_in_case_note("         CAF 1 liquid assets claimed: $" & assets)
call write_new_line_in_case_note("         CAF 1 rent/mortgage claimed: $" & rent)
call write_new_line_in_case_note("        Utilities (amt/HEST claimed): $" & utilities)
call write_new_line_in_case_note("---")
If has_DISQ = True then call write_new_line_in_case_note("A DISQ panel exists for someone on this case.")
If has_DISQ = False then call write_new_line_in_case_note("No DISQ panels were found for this case.")
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)
If expedited_status = "client appears expedited" then
	MsgBox "This client appears expedited. A same day interview needs to be offered."
End if
If expedited_status = "client does not appear expedited" then
	MsgBox "This client does not appear expedited. A same day interview does not need to be offered."
End if

script_end_procedure("")








