'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - MA-EPD EI FIAT"
start_time = timer


'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------

current_month_plus_one = dateadd("m", 1, date)

footer_month = datepart("m", current_month_plus_one)
If len(footer_month) = 1 then footer_month = "0" & footer_month

footer_year = datepart("yyyy", current_month_plus_one)
footer_year = footer_year - 2000

current_month = datepart("m", date)
If len(current_month) = 1 then current_month = "0" & current_month

current_year = datepart("yyyy", date)
current_year = current_year - 2000

current_month_and_year = current_month & "/" & current_year
next_month_and_year = footer_month & "/" & footer_year

'THE SCRIPT

EMConnect ""

BeginDialog case_number_dialog, 0, 0, 156, 61, "Case number"
  Text 5, 5, 85, 10, "Enter your case number:"
  EditBox 90, 0, 60, 15, case_number
  Text 25, 25, 65, 10, "HH memb number:"
  EditBox 90, 20, 30, 15, memb_number
  ButtonGroup ButtonPressed
    OkButton 25, 40, 50, 15
    CancelButton 85, 40, 50, 15
EndDialog

MAXIS_case_number_finder(case_number)

memb_number = "01" 'Setting a default

Dialog case_number_dialog
If buttonpressed = 0 then stopscript

back_to_self

EMWriteScreen "stat", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen footer_month, 20, 43
EMWriteScreen footer_year, 20, 46
EMWriteScreen "jobs", 21, 70
EMWriteScreen memb_number, 21, 75
transmit

EMReadScreen SELF_check, 4, 2, 50
If SELF_check = "SELF" then stopscript

EMReadScreen ERRR_check, 4, 2, 52
If ERRR_check = "ERRR" then transmit

EMReadScreen jobs_total, 1, 2, 78
EMReadScreen jobs_current, 1, 2, 73

If jobs_total = "0" then MsgBox "No JOBS panel is known for this client. You will have to enter income amounts manually."

If jobs_current = "1" then
  EMReadScreen pay_freq_01, 1, 18, 35
  If pay_freq_01 = "1" then frequency_job_01 = "1: monthly"
  If pay_freq_01 = "2" then frequency_job_01 = "2: twice monthly"
  If pay_freq_01 = "3" then frequency_job_01 = "3: every 2 weeks"
  If pay_freq_01 = "4" then frequency_job_01 = "4. every week"
  If pay_freq_01 = "5" then frequency_job_01 = "5. other (use monthly avg)"
  EMWriteScreen "x", 19, 54
  transmit
  EMReadScreen income_job_01, 8, 11, 63
  income_job_01 = trim(replace(income_job_01, "_", ""))
  transmit
  transmit
  EMReadScreen jobs_current, 1, 2, 73
End if

If jobs_current = "2" then
  EMReadScreen pay_freq_02, 1, 18, 35
  If pay_freq_02 = "1" then frequency_job_02 = "1: monthly"
  If pay_freq_02 = "2" then frequency_job_02 = "2: twice monthly"
  If pay_freq_02 = "3" then frequency_job_02 = "3: every 2 weeks"
  If pay_freq_02 = "4" then frequency_job_02 = "4. every week"
  If pay_freq_02 = "5" then frequency_job_02 = "5. other (use monthly avg)"
  EMWriteScreen "x", 19, 54
  transmit
  EMReadScreen income_job_02, 8, 11, 63
  income_job_02 = trim(income_job_02)
  transmit
  transmit
  EMReadScreen jobs_current, 1, 2, 73
End if

If jobs_current = "3" then
  EMReadScreen pay_freq_03, 1, 18, 35
  If pay_freq_03 = "1" then frequency_job_03 = "1: monthly"
  If pay_freq_03 = "2" then frequency_job_03 = "2: twice monthly"
  If pay_freq_03 = "3" then frequency_job_03 = "3: every 2 weeks"
  If pay_freq_03 = "4" then frequency_job_03 = "4. every week"
  If pay_freq_03 = "5" then frequency_job_03 = "5. other (use monthly avg)"
  EMWriteScreen "x", 19, 54
  transmit
  EMReadScreen income_job_03, 8, 11, 63
  income_job_03 = trim(income_job_03)
  transmit
  transmit
  EMReadScreen jobs_current, 1, 2, 73
End if

If income_job_01 = "" then
  income_job_01 = income_job_02
  frequency_job_01 = frequency_job_02
  income_job_02 = ""
  frequency_job_02 = ""
End if

If income_job_02 = "" then
  income_job_02 = income_job_03
  frequency_job_02 = frequency_job_03
  income_job_03 = ""
  frequency_job_03 = ""
End if

BeginDialog MA_EPD_dialog, 0, 0, 186, 156, "MA-EPD dialog"
  Text 35, 5, 40, 10, "Income amt"
  Text 115, 5, 30, 10, "Pay freq."
  Text 5, 25, 25, 10, "Job 1:"
  EditBox 30, 20, 40, 15, income_job_01
  DropListBox 85, 20, 90, 15, " "+chr(9)+"1: monthly"+chr(9)+"2: twice monthly"+chr(9)+"3: every 2 weeks"+chr(9)+"4. every week"+chr(9)+"5. other (use monthly avg)", frequency_job_01
  Text 5, 45, 25, 10, "Job 2:"
  EditBox 30, 40, 40, 15, income_job_02
  DropListBox 85, 40, 90, 15, " "+chr(9)+"1: monthly"+chr(9)+"2: twice monthly"+chr(9)+"3: every 2 weeks"+chr(9)+"4. every week"+chr(9)+"5. other (use monthly avg)", frequency_job_02
  Text 5, 65, 25, 10, "Job 3:"
  EditBox 30, 60, 40, 15, income_job_03
  DropListBox 85, 60, 90, 15, " "+chr(9)+"1: monthly"+chr(9)+"2: twice monthly"+chr(9)+"3: every 2 weeks"+chr(9)+"4. every week"+chr(9)+"5. other (use monthly avg)", frequency_job_03
  GroupBox 20, 85, 140, 40, "Script should update:"
  OptionGroup RadioGroup1
    RadioButton 25, 95, 110, 10, "Current and future months", Radio1
    RadioButton 25, 110, 100, 10, "Just future months", Radio2
  ButtonGroup ButtonPressed
    OkButton 40, 135, 50, 15
    CancelButton 100, 135, 50, 15
EndDialog


Dialog MA_EPD_dialog
If ButtonPressed = 0 then stopscript

'SECTION 04: NOW IT GOES TO ELIG/HC TO FIAT THE AMOUNTS


back_to_SELF

EMWriteScreen "elig", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen "hc", 21, 70
transmit

row = 1
col = 1
EMSearch memb_number & " ", row, col 'finding the member number
If row = 0 then 
  MsgBox "Member number not found. You may have entered an incorrect member number on the first screen. Try the script again."
  StopScript
End if

EMWriteScreen "x", row, 26
transmit

EMReadScreen elig_type_check_first_month, 2, 12, 17
EMReadScreen elig_type_check_second_month, 2, 12, 28
EMReadScreen elig_type_check_third_month, 2, 12, 39
EMReadScreen elig_type_check_fourth_month, 2, 12, 50
EMReadScreen elig_type_check_fifth_month, 2, 12, 61
EMReadScreen elig_type_check_sixth_month, 2, 12, 72

If elig_type_check_first_month <> "DP" and elig_type_check_second_month <> "DP" and elig_type_check_third_month <> "DP" and elig_type_check_fourth_month <> "DP" and elig_type_check_fifth_month <> "DP" and elig_type_check_sixth_month <> "DP" then MsgBox "Not all of the months of this case are MA-EPD. Process manually."
If elig_type_check_first_month <> "DP" and elig_type_check_second_month <> "DP" and elig_type_check_third_month <> "DP" and elig_type_check_fourth_month <> "DP" and elig_type_check_fifth_month <> "DP" and elig_type_check_sixth_month <> "DP" then stopscript

PF9
EMReadScreen FIAT_check, 4, 24, 45
If FIAT_check <> "FIAT" then
  EMSendKey "05"
  transmit
End if
If radio1 = 1 then
  row = 6
  col = 1
  EMSearch current_month_and_year, row, col
End if

If radio2 = 1 or row = 0 then
  row = 6
  col = 1
  EMSearch next_month_and_year, row, col
End if

'Multiplier calculations
If frequency_job_01 = "1: monthly" or frequency_job_01 = "5. other (use monthly avg)" then multiplier_01 = 1
If frequency_job_02 = "1: monthly" or frequency_job_02 = "5. other (use monthly avg)" then multiplier_02 = 1
If frequency_job_03 = "1: monthly" or frequency_job_03 = "5. other (use monthly avg)" then multiplier_03 = 1

If frequency_job_01 = "2: twice monthly" then multiplier_01 = 2
If frequency_job_02 = "2: twice monthly" then multiplier_02 = 2
If frequency_job_03 = "2: twice monthly" then multiplier_03 = 2

If frequency_job_01 = "3: every 2 weeks" then multiplier_01 = 2.16
If frequency_job_02 = "3: every 2 weeks" then multiplier_02 = 2.16
If frequency_job_03 = "3: every 2 weeks" then multiplier_03 = 2.16

If frequency_job_01 = "4. every week" then multiplier_01 = 4.3
If frequency_job_02 = "4. every week" then multiplier_02 = 4.3
If frequency_job_03 = "4. every week" then multiplier_03 = 4.3

Do
  EMWriteScreen "x", 9, col + 2
  transmit
  EMWriteScreen "x", 13, 03
  transmit
  EMWriteScreen "___________", 8, 43
  EMWriteScreen income_job_01 * multiplier_01, 8, 43
  If income_job_02 <> "" then
    EMWriteScreen "___________", 9, 43
    EMWriteScreen income_job_02 * multiplier_02, 9, 43
  End if
  If income_job_03 <> "" then
    EMWriteScreen "___________", 10, 43
    EMWriteScreen income_job_03 * multiplier_03, 10, 43
  End if
  col = col + 11
  transmit
  transmit
  transmit
loop until col > 76




script_end_procedure("")






