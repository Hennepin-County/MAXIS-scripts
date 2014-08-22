'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - SAVE (LEP)"
start_time = timer


'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
'>>>>NOTE: these were added as a batch process. Check below for any 'StopScript' functions and convert manually to the script_end_procedure("") function
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

EMConnect ""

BeginDialog SAVE_dialog, 0, 0, 196, 301, "SAVE Dialog"
  OptionGroup RadioGroup1
    RadioButton 5, 5, 45, 10, "SAVE 1", SAVE_1
    RadioButton 5, 20, 45, 10, "SAVE 2", SAVE_2
  GroupBox 5, 35, 185, 105, "SAVE 1"
  Text 10, 50, 50, 10, "Current status:"
  EditBox 65, 45, 120, 15, current_status
  Text 10, 70, 65, 10, "LPR adjusted from:"
  EditBox 80, 65, 105, 15, LPR_adjusted_from
  Text 10, 90, 50, 10, "Date of entry:"
  EditBox 60, 85, 125, 15, date_of_entry
  Text 10, 110, 60, 10, "Country of origin:"
  EditBox 70, 105, 115, 15, country_of_origin
  CheckBox 10, 125, 75, 10, "SAVE 2 requested?", SAVE_2_requested_check
  GroupBox 5, 145, 185, 55, "SAVE 2"
  Text 10, 155, 135, 10, "Sponsored on I-864 Affidavit of Support?"
  OptionGroup RadioGroup2
    RadioButton 15, 170, 35, 10, "No", not_sponsored
    RadioButton 15, 185, 75, 10, "Yes, sponsored by:", sponsored
  EditBox 95, 180, 90, 15, sponsor
  Text 10, 210, 85, 10, "Imigration doc received:"
  EditBox 95, 205, 95, 15, imig_doc_received
  Text 10, 230, 35, 10, "Exp date:"
  EditBox 45, 225, 40, 15, exp_date
  CheckBox 10, 245, 170, 10, "TIKL out to re-request 90 days before expiration.", TIKL_check
  Text 10, 265, 25, 10, "Notes:"
  EditBox 40, 260, 150, 15, notes
  Text 10, 285, 70, 10, "Sign your case note:"
  EditBox 80, 280, 110, 15, worker_sig
  ButtonGroup ButtonPressed
    OkButton 75, 10, 50, 15
    CancelButton 135, 10, 50, 15
EndDialog

Sub find_case_note
  EMReadScreen case_note_ready, 17, 2, 33
  EMReadScreen case_note_mode, 7, 20, 3
  If case_note_ready <> "Case Notes (NOTE)" then msgbox "You aren't in a case note on edit mode. You need to be in a case note on edit mode."
  If case_note_mode <> "Mode: A" and case_note_mode <> "Mode: E" then msgbox "You aren't in a case note on edit mode. You need to be in a case note on edit mode."
  If case_note_mode <> "Mode: A" and case_note_mode <> "Mode: E" then Dialog SAVE_dialog
  If buttonpressed = 0 then stopscript
End Sub

Do
  Dialog SAVE_dialog
  If TIKL_check = 1 and IsDate(exp_date) = False then MsgBox "You must enter a proper date (MM/DD/YYYY) if you want the script to TIKL out. Try again."
  If buttonpressed = 0 then stopscript
Loop until ButtonPressed = -1 and (TIKL_check = 0 or (TIKL_check = 1 and IsDate(exp_date) = True))


Do
  find_case_note
Loop until case_note_ready = "Case Notes (NOTE)" and case_note_mode = "Mode: A" or case_note_mode = "Mode: E"

EMSendKey "<enter>"
Do
  EMReadScreen password_prompt, 38, 2, 23
  IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then MsgBox "You are locked out of your case note. Type your password then try again."
  IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then Dialog SAVE_dialog
  IF buttonpressed = 0 then stopscript
Loop until password_prompt <> "ACF2/CICS PASSWORD VERIFICATION PROMPT"

IF SAVE_1 = 1 then 
  EMSendKey "**SAVE 1**" & "<newline>"
  EMSendKey "* Current status: " & current_status & "<newline>"
  EMSendKey "* LPR adjusted from: " & LPR_adjusted_from & "<newline>"
  EMSendKey "* Date of entry: " & date_of_entry & "<newline>"
  EMSendKey "* Country of origin: " & country_of_origin & "<newline>"
End if
IF SAVE_2 = 1 then 
  EMSendKey "**SAVE 2**" & "<newline>"
  If not_sponsored = 1 then EMSendKey "* No sponsor indicated on SAVE." & "<newline>"
  If sponsored = 1 then EMSendKey "* Client is sponsored. Sponsor is indicated as " & sponsor & "." & "<newline>"
End if

EMSendKey "* Immigration document received: " & imig_doc_received & "<newline>"
EMSendKey "* Exp date: " + exp_date 
If TIKL_check = 1 then EMSendKey ", TIKLed to re-request " & dateadd("d", -90, exp_date) & "."
EMSendKey "<newline>"
If SAVE_2_requested_check = 1 then EMSendKey "* SAVE 2 requested." & "<newline>"
If notes <> "" then EMSendKey "* Notes: " + notes + "<newline>"
EMSendKey "---" + "<newline>"
EMSendKey worker_sig

If TIKL_check <> 1 then stopscript

EMReadScreen case_number, 8, 20, 38

Do
  EMSendKey "<PF3>"
  EMWaitReady 1, 1
  EMReadScreen SELF_check, 4, 2, 50
Loop until SELF_check = "SELF"

EMWriteScreen "dail", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen "writ", 21, 70
EMSendKey "<enter>"
EMWaitReady 1, 1

TIKL_date = dateadd("d", -90, exp_date)
TIKL_month = datepart("m", TIKL_date)
If len(TIKL_month) = 1 then TIKL_month = "0" & TIKL_month
TIKL_day = datepart("d", TIKL_date)
If len(TIKL_day) = 1 then TIKL_day = "0" & TIKL_day
TIKL_year = datepart("yyyy", TIKL_date) - 2000

EMWriteScreen TIKL_month, 5, 18
EMWriteScreen TIKL_day, 5, 21
EMWriteScreen TIKL_year, 5, 24

EMSetCursor 9, 3
EMSendKey "Check on immigration documentation. If it hasn't been updated, request updated info, as what we have expires " & exp_date & ". TIKL generated via script."
EMSendKey "<enter>"
EMWaitReady 1, 1
EMSendKey "<PF3>"
EMWaitReady 1, 1
MsgBox "TIKL sent for " & TIKL_date & ", 90 days prior to document expiration."
script_end_procedure("")






