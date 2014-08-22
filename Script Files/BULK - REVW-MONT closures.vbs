'warning_box = Inputbox ("Enter password to continue:", 1)
'If warning_box <> "testitout" then stopscript

name_of_script = "BULK - REVW-MONT closures"
start_time = timer

'LOADING ROUTINE FUNCTIONS------------------------------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script


'----------------------IT NEEDS TO CALCULATE SOME DATES AND DATE INFO
'current_date = "07/27/2013" 'This setting should be commented out unless testing.
current_date = date         'This should be the default setting for production.

If datepart("m", dateadd("d", 8, current_date)) = datepart("m", current_date) then script_end_procedure("This script cannot be run until the last week of the month.")

footer_month = datepart("m", dateadd("m", 1, date))
if len(footer_month) = 1 then footer_month = "0" & footer_month
footer_year = datepart("yyyy", dateadd("m", 1, date))
footer_year = footer_year - 2000


'----------------------THIS IS THE DIALOG FOR THE SCRIPT
BeginDialog REVW_MONT_closures_dialog, 0, 0, 256, 110, "REVW/MONT closures"
  EditBox 195, 15, 55, 15, worker_signature
  EditBox 205, 35, 45, 15, worker_number
  CheckBox 15, 75, 120, 10, "REPT/MONT? (HRFs)", MONT_check
  CheckBox 15, 90, 120, 10, "REPT/REVW? (CSRs and ARs)", REVW_check
  ButtonGroup ButtonPressed
    OkButton 200, 65, 50, 15
    CancelButton 200, 90, 50, 15
  Text 5, 5, 185, 25, "This script will case note all of your renewals that are closing/incomplete. You'll need to sign your case notes:"
  Text 5, 40, 195, 10, "Enter the last three digits of your x1# here (e.g. ''X100###''):"
  GroupBox 5, 60, 150, 45, "Case note closing/incomplete cases from:"
EndDialog




'----------------------CONNECTING TO BLUEZONE, RUNNING THE DIALOG, AND NAVIGATING TO REPT/REVW
EMConnect ""
Do
  Do
    Dialog REVW_MONT_closures_dialog
    If ButtonPressed = 0 then StopScript 'Cancel button
    If worker_number <> "" then worker_number = ucase(worker_number)
    If len(worker_number) <> 3 then MsgBox "You must enter the last three digits of your " & county_worker_code & "# (and just the last three digits)."
  Loop until len(worker_number) = 3
  If worker_signature = "" then MsgBox "You must sign your case note."
Loop until worker_signature <> ""

transmit 'It transmits to check for MAXIS.
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then script_end_procedure("MAXIS is not found. You may be passworded out. Try it again.")

'THIS PART DOES THE REPT REVW----------------------------------------------------------------------------------------------------
If revw_check = 1 then 
  call navigate_to_screen("rept", "revw")
  EMReadScreen default_worker_number, 3, 21, 10
  If worker_number <> default_worker_number then
    EMWriteScreen worker_county_code & worker_number, 21, 6
    transmit
  End if
  EMReadScreen current_footer_month, 2, 20, 55
  EMReadScreen current_footer_year, 2, 20, 58
  If (current_footer_month <> footer_month) or (current_footer_year <> footer_year) then
    EMWriteScreen footer_month, 20, 55
    EMWriteScreen footer_year, 20, 58
    transmit
  End if
  row = 7
  Do
    EMReadScreen case_number, 8, row, 6
    EMReadScreen program_status, 21, row, 35
    are_programs_closing = instr(program_status, "N") <> 0 or instr(program_status, "I") <> 0
    If are_programs_closing = True then case_number_array = trim(case_number_array & " " & trim(case_number))
    row = row + 1
    If row = 19 then
      PF8
      EMReadScreen last_check, 4, 24, 14
      row = 7
    End if
  Loop until trim(case_number) = "" or last_check = "LAST"


  case_number_array = split(case_number_array)
  
  
  '-----------------------NAVIGATING TO EACH CASE AND CASE NOTING THE ONES THAT ARE CLOSING
  For each case_number in case_number_array
    call navigate_to_screen("stat", "revw")
    EMReadScreen ERRR_check, 4, 2, 52
    If ERRR_check = "ERRR" then call navigate_to_screen("stat", "revw") 'In case of error prone cases
    EMReadScreen cash_review_code, 1, 7, 40
    EMReadScreen WB_review_code, 1, 7, 50
    EMReadScreen FS_review_code, 1, 7, 60
    EMReadScreen HC_review_code, 1, 7, 73
    If cash_review_code = "N" then cash_review_status = "closing for no renewal CAF."
    If cash_review_code = "I" then cash_review_status = "closing for incomplete review. See previous case notes for details on what's needed."
    If WB_review_code = "N" then WB_review_status = "closing for no renewal."
    If WB_review_code = "I" then WB_review_status = "closing for incomplete review. See previous case notes for details on what's needed."
    If FS_review_code = "N" then 
      EMWriteScreen "x", 5, 58
      transmit
      EMReadScreen recertification_date, 8, 9, 64
      recertification_date = cdate(replace(recertification_date, " ", "/"))
      If datepart("m", recertification_date) = datepart("m", dateadd("m", 1, now)) then
        FS_review_document = "renewal CAF"
      Else
        FS_review_document = "CSR"
      End if
      FS_review_status = "closing for no " & FS_review_document & "."
      transmit
    End if
    If FS_review_code = "I" then FS_review_status = "closing for incomplete review. See previous case notes for details on what's needed."
    If HC_review_code = "N" then 
      EMWriteScreen "x", 5, 71
      transmit
      EMReadScreen recertification_date, 8, 9, 27
      recertification_date = cdate(replace(recertification_date, " ", "/"))
      If datepart("m", recertification_date) = datepart("m", dateadd("m", 1, now)) then
        If FS_review_code = "_" and cash_review_code = "_" then
          HC_review_document = "renewal HC ER"
        Else
          HC_review_document = "renewal CAF"
        End If
      Else
        HC_review_document = "CSR"
      End if
      HC_review_status = "closing for no " & HC_review_document & "."
      transmit
    End if
    If HC_review_code = "I" then HC_review_status = "closing for incomplete review. See previous case notes for details on what's needed."
  
  '---------------THIS SECTION FIGURES OUT WHEN PROGRAMS CAN TURN IN NEW RENEWALS AND WHEN THEY BECOME INTAKES AGAIN
    If cash_review_status <> "" or WB_review_status <> "" or FS_review_status <> "" or HC_review_status <> "" then
      EMReadScreen first_of_working_month, 5, 20, 55
      first_of_working_month = cdate(replace(first_of_working_month, " ", "/01/"))
      last_day_to_turn_in_docs = dateadd("d", -1, (dateadd("m", 1, first_of_working_month)))
      intake_date = dateadd("m", 1, first_of_working_month)
    End If
  
  '---------------NOW IT CASE NOTES
    PF4
    PF9
  
    If HC_review_code = "I" or FS_review_code = "I" or WB_review_code = "I" or cash_review_code = "I" then
      call write_new_line_in_case_note("---Programs closing for incomplete review---")
    Else
      call write_new_line_in_case_note("---Programs closing for no review---")
    End if
    If cash_review_status <> "" then call write_editbox_in_case_note("Cash", cash_review_status, 5)
    If WB_review_status <> "" then call write_editbox_in_case_note("WB", WB_review_status, 5)
    If FS_review_status <> "" then call write_editbox_in_case_note("SNAP", FS_review_status, 5)
    If HC_review_status <> "" then call write_editbox_in_case_note("HC", HC_review_status, 5)
    If last_day_to_turn_in_docs <> "" then call write_new_line_in_case_note("* Client has until " & last_day_to_turn_in_docs & " to turn in review doc and/or proofs.")
    If intake_date <> "" then call write_new_line_in_case_note("* Client needs to reapply after " & intake_date & ".")
    call write_new_line_in_case_note("---")
    call write_new_line_in_case_note(worker_signature & ", via automated script.")
  
  '----------------NOW IT RESETS THE VARIABLES FOR THE REVIEW CODES, STATUS, AND DATES
    cash_review_code = ""
    WB_review_code = ""
    FS_review_code = ""
    HC_review_code = ""
    cash_review_status = ""
    WB_review_status = ""
    FS_review_status = ""
    HC_review_status = ""
    first_of_working_month = ""
    last_day_to_turn_in_docs = ""
    intake_date = ""
  
  Next
  
  call navigate_to_screen("rept", "revw")
  EMReadScreen default_worker_number, 3, 21, 10
  If worker_number <> default_worker_number then
    EMWriteScreen worker_county_code & worker_number, 21, 10
    transmit
  End if
End If  

'Resetting the case number array
case_number_array = ""

'THIS PART DOES THE REPT MONT----------------------------------------------------------------------------------------------------
If mont_check = 1 then 
  'Navigating to MONT
  call navigate_to_screen("rept", "mont")

  'Checking the current worker number. If it's not the selected one it will enter the selected one.
  EMReadScreen default_worker_number, 3, 21, 10
  If worker_number <> default_worker_number then
    EMWriteScreen worker_county_code & worker_number, 21, 10
    transmit
  End if

  'Checking the footer month/year. If it's incorrect it will adjust.
  EMReadScreen current_footer_month, 2, 20, 54
  EMReadScreen current_footer_year, 2, 20, 57
  If (current_footer_month <> footer_month) or (current_footer_year <> footer_year) then
    EMWriteScreen footer_month, 20, 54
    EMWriteScreen footer_year, 20, 57
    transmit
  End if

  'Setting the variable for the following do...loop
  row = 7

  'This reads the case number and program status. If an "N" or "I" is detected it will add to the case_number_array variable.
  Do
    EMReadScreen case_number, 8, row, 6
    EMReadScreen program_status, 9, row, 45
    are_programs_closing = instr(program_status, "N") <> 0 or instr(program_status, "I") <> 0
    If are_programs_closing = True then case_number_array = trim(case_number_array & " " & trim(case_number))
    row = row + 1
    If row = 19 then
      PF8
      EMReadScreen last_check, 4, 24, 14
      row = 7
    End if
  Loop until trim(case_number) = "" or last_check = "LAST"

  'Creating an array out of the case number array
  case_number_array = split(case_number_array)
  
  
  'Navigating to each case, and case noting the ones that are closing.
  For each case_number in case_number_array
    'Going to the case, checking for error prone
    call navigate_to_screen("stat", "mont")
    EMReadScreen ERRR_check, 4, 2, 52
    If ERRR_check = "ERRR" then call navigate_to_screen("stat", "mont") 'In case of error prone cases

    'Reading the review codes, converting them to a status update for the case note
    EMReadScreen cash_review_code, 1, 11, 43
    EMReadScreen FS_review_code, 1, 11, 53
    EMReadScreen GRH_review_code, 1, 11, 63
    EMReadScreen HC_review_code, 1, 11, 73


  '---------------NOW IT CASE NOTES
    PF4
    PF9


  
    If HC_review_code = "I" or FS_review_code = "I" or GRH_review_code = "I" or cash_review_code = "I" then
      call write_new_line_in_case_note("---Incomplete HRF---")
    Else
      call write_new_line_in_case_note("---HRF not provided---")
    End if
    call write_new_line_in_case_note("---")
    call write_new_line_in_case_note(worker_signature & ", via automated script.")


  
  '----------------NOW IT RESETS THE VARIABLES FOR THE REVIEW CODES, STATUS, AND DATES
    cash_review_code = ""
    GRH_review_code = ""
    FS_review_code = ""
    HC_review_code = ""
    cash_review_status = ""
    GRH_review_status = ""
    FS_review_status = ""
    HC_review_status = ""
    first_of_working_month = ""
    last_day_to_turn_in_docs = ""
    intake_date = ""
  
  Next
  
  call navigate_to_screen("rept", "mont")
  EMReadScreen default_worker_number, 3, 21, 10
  If worker_number <> default_worker_number then
    EMWriteScreen worker_county_code & worker_number, 21, 10
    transmit
  End if
End If  


MsgBox "Success! All cases that are coded in REPT/REVW and/or REPT/MONT as either an ''N'' or an ''I'' have been case noted for why they're closing, and what documents need to get turned in."
script_end_procedure("")






