'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - expedited screening (FAS)"
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




'SECTION 01
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

'Assigns numbers to the income/asset/rent/utilities variables.
If income = "" then income = 0
If assets = "" then assets = 0
If rent = "" then rent = 0
If phone_check = 1 then utilities = 40                                               '$40 is the phone standard for utility calculation as of November 2013.
If electric_check = 1 then utilities = 141                                           '$141 is the electric standard for utility calculation as of November 2013.
If electric_check = 1 and phone_check = 1 then utilities = 181                       'Phone standard plus electric standard.
If heat_AC_check = 1 then utilities = 459                                            '$459 is the maximum utility standard as of November 2013. If a client qualifies for this, they do not get the other two.
If phone_check = 0 and electric_check = 0 and heat_AC_check = 0 then utilities = 0   'in case no options are clicked, utilities is set to zero.


'Calculates expedited status based on above numbers
If (cint(income) < 150 and cint(assets) <= 100) or ((cint(income) + cint(assets)) < (cint(rent) + cint(utilities))) then expedited_status = "client appears expedited"
If (cint(income) + cint(assets) >= cint(rent) + cint(utilities)) and (cint(income) >= 150 or cint(assets) > 100) then expedited_status = "client does not appear expedited"

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
EMSendKey "Received " & application_type & ", " & expedited_status + "<newline>"
EMSendKey "---" + "<newline>"
EMSendKey "     CAF 1 income claimed this month: $" & income & "<newline>"
EMSendKey "         CAF 1 liquid assets claimed: $" & assets & "<newline>"
EMSendKey "         CAF 1 rent/mortgage claimed: $" & rent & "<newline>"
EMSendKey "        Utilities (amt/HEST claimed): $" & utilities & "<newline>"
EMSendKey "---" + "<newline>"
If has_DISQ = True then EMSendKey "A DISQ panel exists for someone on this case." + "<newline>"
If has_DISQ = False then EMSendKey "No DISQ panels were found for this case." + "<newline>"
EMSendKey "---" + "<newline>"
EMSendKey worker_signature
If expedited_status = "client appears expedited" then
  MsgBox "This client appears expedited. A same day interview needs to be offered."
End if
If expedited_status = "client does not appear expedited" then
  MsgBox "This client does not appear expedited. A same day interview does not need to be offered."
End if

script_end_procedure("")






