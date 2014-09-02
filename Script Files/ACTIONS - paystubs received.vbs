'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - paystubs received"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'CUSTOM FUNCTIONS

Function PIC_paystubs_info_adder(pay_date, gross_amt, hours)
  If isdate(pay_date) = True then
    If len(datepart("m", pay_date)) = 2 then
      EMWriteScreen datepart("m", pay_date), PIC_row, 13
    Else
      EMWriteScreen "0" & datepart("m", pay_date), PIC_row, 13
    End if
    If len(datepart("d", pay_date)) = 2 then
      EMWriteScreen datepart("d", pay_date), PIC_row, 16
    Else
      EMWriteScreen "0" & datepart("d", pay_date), PIC_row, 16
    End if
    EMWriteScreen right(datepart("yyyy", pay_date), 2), PIC_row, 19
    EMWriteScreen gross_amt, PIC_row, 25
    EMWriteScreen hours, PIC_row, 35
    PIC_row = PIC_row + 1
  End If
End function

Function prospective_averager(pay_date, gross_amt, hours) 'Creates variables for total_prospective_pay and total_prospective_hours
  If isdate(pay_date) = True then
    total_prospective_pay = total_prospective_pay + abs(gross_amt)
    total_prospective_hours = total_prospective_hours + abs(hours)
    paystubs_received = paystubs_received + 1
  Else
    pay_date = "01/01/2000"
  End if
End function

Function prospective_pay_analyzer(pay_date, gross_amt)
  If datediff("m", pay_date, footer_month & "/01/" & footer_year) = 0 then
    If len(datepart("m", pay_date)) = 2 then
      EMWriteScreen datepart("m", pay_date), MAXIS_row, 54
    Else
      EMWriteScreen "0" & datepart("m", pay_date), MAXIS_row, 54
    End if
    If len(datepart("d", pay_date)) = 2 then
      EMWriteScreen datepart("d", pay_date), MAXIS_row, 57
    Else
      EMWriteScreen "0" & datepart("d", pay_date), MAXIS_row, 57
    End if
    EMWriteScreen right(datepart("yyyy", pay_date), 2), MAXIS_row, 60
    EMWriteScreen gross_amt, MAXIS_row, 67
    MAXIS_row = MAXIS_row + 1
  End if
End function

Function retro_paystubs_info_adder(pay_date, gross_amt, hours)
  If isdate(pay_date) = True then
    If datediff("m", pay_date, footer_month & "/01/" & footer_year) = 2 then 
      If len(datepart("m", pay_date)) = 2 then
        EMWriteScreen datepart("m", pay_date), MAXIS_row, 25
      Else
        EMWriteScreen "0" & datepart("m", pay_date), MAXIS_row, 25
      End if
      If len(datepart("d", pay_date)) = 2 then
        EMWriteScreen datepart("d", pay_date), MAXIS_row, 28
      Else
        EMWriteScreen "0" & datepart("d", pay_date), MAXIS_row, 28
      End if
      EMWriteScreen right(datepart("yyyy", pay_date), 2), MAXIS_row, 31
      EMWriteScreen gross_amt, MAXIS_row, 38
      retro_hours = abs(retro_hours + abs(hours))
      MAXIS_row = MAXIS_row + 1
    End if
  End if
End function

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog paystubs_received_dialog, 0, 0, 256, 220, "Paystubs Received Dialog"
  DropListBox 100, 5, 100, 15, "(select one)"+chr(9)+"One Time Per Month"+chr(9)+"Two Times Per Month"+chr(9)+"Every Other Week"+chr(9)+"Every Week", pay_frequency
  EditBox 15, 45, 65, 15, pay_date_01
  EditBox 95, 45, 65, 15, gross_amt_01
  EditBox 175, 45, 65, 15, hours_01
  EditBox 15, 65, 65, 15, pay_date_02
  EditBox 95, 65, 65, 15, gross_amt_02
  EditBox 175, 65, 65, 15, hours_02
  EditBox 15, 85, 65, 15, pay_date_03
  EditBox 95, 85, 65, 15, gross_amt_03
  EditBox 175, 85, 65, 15, hours_03
  EditBox 15, 105, 65, 15, pay_date_04
  EditBox 95, 105, 65, 15, gross_amt_04
  EditBox 175, 105, 65, 15, hours_04
  EditBox 15, 125, 65, 15, pay_date_05
  EditBox 95, 125, 65, 15, gross_amt_05
  EditBox 175, 125, 65, 15, hours_05
  EditBox 55, 155, 190, 15, explanation_of_income
  DropListBox 75, 180, 120, 15, "(select one)"+chr(9)+"1 Pay Stubs/Tip Report"+chr(9)+"2 Empl Statement"+chr(9)+"3 Coltrl Stmt"+chr(9)+"4 Other Document"+chr(9)+"5 Pend Out State Verification"+chr(9)+"N No Ver Prvd", JOBS_verif_code
  EditBox 75, 200, 115, 15, worker_signature
  ButtonGroup buttonpressed
    OkButton 200, 180, 50, 15
    CancelButton 200, 200, 50, 15
  Text 40, 10, 55, 10, "Pay frequency:"
  Text 10, 30, 80, 10, "Pay date (MM/DD/YY):"
  Text 105, 30, 50, 10, "Gross amount:"
  Text 195, 30, 30, 10, "Hours:"
  GroupBox 5, 145, 245, 30, "Explain how income was calculated:"
  Text 10, 160, 45, 10, "Explanation:"
  Text 10, 185, 60, 10, "JOBS verif code:"
  Text 10, 205, 60, 10, "Worker signature:"
EndDialog




BeginDialog paystubs_received_case_number_dialog, 0, 0, 376, 170, "Case number"
  EditBox 100, 5, 60, 15, case_number
  EditBox 70, 25, 25, 15, footer_month
  EditBox 125, 25, 25, 15, footer_year
  EditBox 110, 45, 25, 15, HH_member
  CheckBox 15, 75, 110, 10, "Update and case note the PIC?", update_PIC_check
  CheckBox 15, 90, 75, 10, "Update HC popup?", update_HC_popup_check
  CheckBox 15, 105, 140, 10, "Check here to have the script update all", future_months_check
  CheckBox 15, 130, 135, 10, "Case note info about paystubs?", case_note_check
  ButtonGroup ButtonPressed
    OkButton 265, 150, 50, 15
    CancelButton 320, 150, 50, 15
  Text 10, 10, 85, 10, "Enter your case number:"
  GroupBox 175, 5, 195, 140, "INSTRUCTIONS!!! PLEASE READ!!!"
  Text 185, 20, 180, 35, "This script, by default, will update retro/pro in the footer month specified only. It can update multiple months and send through background if you select that to the left. It can also update the PIC or HC pop-ups."
  Text 185, 60, 180, 50, "PLEASE NOTE: you should already have a JOBS panel made for this client. If you haven't made a JOBS panel yet, make it and send the case through background before using this script. The script only does one job at a time, so you may need to run it more than once if you have multiple jobs."
  Text 185, 115, 180, 25, "You should also have all of the paystubs you need to update MAXIS. If you aren't ready to update STAT/JOBS, don't use this script."
  Text 20, 30, 50, 10, "Footer month:"
  Text 100, 30, 20, 10, "Year:"
  Text 35, 50, 75, 10, "HH memb # for JOBS:"
  GroupBox 10, 65, 150, 80, "Options"
  Text 30, 115, 120, 10, "future months and send through BG."
EndDialog




'THE SCRIPT----------------------------------------------------------------------------------------------------

EMConnect ""

'Finds the case number
call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""

'Default footer month is the month the worker is in, and if that isn't found, it's the current month.
call find_variable("Month: ", footer_month, 5)
footer_year = right(footer_month, 2)
footer_month = left(footer_month, 2)
If isnumeric(footer_month) = False or isnumeric(footer_year) = False then
  footer_month = datepart("m", date)
  If len(footer_month) = 1 then footer_month = "0" & footer_month
  footer_year = right(datepart("yyyy", date), 2)
End if

'Default member is member 01
HH_member = "01"

'Shows the case number dialog
Dialog paystubs_received_case_number_dialog
If buttonpressed = 0 then stopscript

'Shows the paystub dialog. Includes logic to prevent paydates from being entered incorrectly.
Do
  Do
    Do
      Do
        Do
          Do
            Do
              Do
                Dialog paystubs_received_dialog
                If ButtonPressed = 0 then stopscript
                If pay_frequency = "(select one)" then MsgBox "You must select a pay frequency."
              Loop until pay_frequency <> "(select one)"
              If JOBS_verif_code = "(select one)" then MsgBox "You must select a JOBS verif code."
            Loop until JOBS_verif_code <> "(select one)"
            If explanation_of_income = "" then MsgBox "You must explain how you calculated this income (ie: ''all paystubs from last 30 days'')"
          Loop until explanation_of_income <> ""
          If isdate(pay_date_01) = False then MsgBox "Your pay date must be ''MM/DD/YYYY'' format. Please try again."
          If isdate(pay_date_01) = True and (Isnumeric(gross_amt_01) = False or Isnumeric(hours_01) = False) then MsgBox "You must include a gross pay amount as well as an hours amount."
        Loop until (isdate(pay_date_01) = True and (Isnumeric(gross_amt_01) = True and Isnumeric(hours_01) = True))
        pay_date_02 = trim(pay_date_02)
        If isdate(pay_date_02) = False and pay_date_02 <> "" then MsgBox "Your pay date must be ''MM/DD/YYYY'' format. Please try again."
        If isdate(pay_date_02) = True and (Isnumeric(gross_amt_02) = False or Isnumeric(hours_02) = False) then MsgBox "You must include a gross pay amount as well as an hours amount."
      Loop until (isdate(pay_date_02) = True and (Isnumeric(gross_amt_02) = True and Isnumeric(hours_02) = True)) or pay_date_02 = ""
      pay_date_03 = trim(pay_date_03)
      If isdate(pay_date_03) = False and pay_date_03 <> "" then MsgBox "Your pay date must be ''MM/DD/YYYY'' format. Please try again."
      If isdate(pay_date_03) = True and (Isnumeric(gross_amt_03) = False or Isnumeric(hours_03) = False) then MsgBox "You must include a gross pay amount as well as an hours amount."
    Loop until (isdate(pay_date_03) = True and (Isnumeric(gross_amt_03) = True and Isnumeric(hours_03) = True)) or pay_date_03 = ""
    pay_date_04 = trim(pay_date_04)
    If isdate(pay_date_04) = False and pay_date_04 <> "" then MsgBox "Your pay date must be ''MM/DD/YYYY'' format. Please try again."
    If isdate(pay_date_04) = True and (Isnumeric(gross_amt_04) = False or Isnumeric(hours_04) = False) then MsgBox "You must include a gross pay amount as well as an hours amount."
  Loop until (isdate(pay_date_04) = True and (Isnumeric(gross_amt_04) = True and Isnumeric(hours_04) = True)) or pay_date_04 = ""
  pay_date_05 = trim(pay_date_05)
  If isdate(pay_date_05) = False and pay_date_05 <> "" then MsgBox "Your pay date must be ''MM/DD/YYYY'' format. Please try again."
  If isdate(pay_date_05) = True and (Isnumeric(gross_amt_05) = False or Isnumeric(hours_05) = False) then MsgBox "You must include a gross pay amount as well as an hours amount."
Loop until(isdate(pay_date_05) = True and (Isnumeric(gross_amt_05) = True and Isnumeric(hours_05) = True)) or pay_date_05 = ""

'Sends a transmit to refresh screen, then checks for MAXIS status. Does this on a loop so as not to lose pay information, and includes a cancel button.
Do
  transmit
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then
    MAXIS_check_msgbox = MsgBox("MAXIS not found. You may be passworded out.", 1)
    If MAXIS_check_msgbox = 2 then stopscript
  End if
Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "

'Checks to see if it's in STAT, and checks footer month/year. If it isn't in STAT or the right footer month/year, the script will leave the case.
EMReadScreen STAT_check, 4, 20, 21
EMReadScreen STAT_case_number, 8, 20, 37
EMReadScreen STAT_footer_month_check, 2, 20, 55
EMReadScreen STAT_footer_year_check, 2, 20, 58
If STAT_check <> "STAT" or trim(replace(STAT_case_number, "_", "")) <> case_number or STAT_footer_month_check <> footer_month or STAT_footer_year_check <> footer_year then back_to_SELF

call navigate_to_screen("stat", "jobs")

'Heads into the case/curr screen, checks to make sure the case number is correct before proceeding. If it can't get beyond the SELF menu the script will stop.
EMReadScreen SELF_check, 4, 2, 50
If SELF_check = "SELF" then stopscript

'Navigates to the JOBS panel for the right person
If HH_member <> "01" then 
  EMWriteScreen HH_member, 20, 76
  transmit
End if

'Checks to make sure there are JOBS panels for this member. If none exist the script will close
EMReadScreen total_amt_of_panels, 1, 2, 78
If total_amt_of_panels = "0" then script_end_procedure("No JOBS panels exist for this client. Please add a JOBS panel and run through background before trying again. The script will now stop.")

'If there is more than one panel, this part will grab employer info off of them and present it to the worker to decide which one to use.
If total_amt_of_panels <> "1" then
  Do
    EMReadScreen current_panel_number, 1, 2, 73
    EMReadScreen employer_name, 30, 7, 42
    employer_check = MsgBox("Is this your employer? Employer name: " & trim(replace(employer_name, "_", "")), 3)
    If employer_check = 2 then stopscript
    If employer_check = 6 then exit do
    If employer_check = 7 and current_panel_number = total_amt_of_panels then 
      EMWriteScreen "01", 20, 79
      current_panel_number = "1"
    End if
    transmit
  Loop until current_panel_number = total_amt_of_panels
End if

'Turns on edit mode
PF9

'Totals the prospective amounts, inserts "01/01/2000" for dates that were left blank, using function.
Call prospective_averager(pay_date_01, gross_amt_01, hours_01)
Call prospective_averager(pay_date_02, gross_amt_02, hours_02)
Call prospective_averager(pay_date_03, gross_amt_03, hours_03)
Call prospective_averager(pay_date_04, gross_amt_04, hours_04)
Call prospective_averager(pay_date_05, gross_amt_05, hours_05)

'Creates averages
dim paystubs_received
dim total_prospective_pay 
dim total_prospective_hours
average_pay_per_paystub = formatnumber(total_prospective_pay / paystubs_received, 2, 0, 0, 0)
average_hours_per_paystub = abs(total_prospective_hours / paystubs_received)


Do
  'If SNAP was active the script must update the PIC.
  If update_PIC_check = 1 then
    EMWriteScreen "x", 19, 38
    transmit
    'Clears existing info off PIC
    EMSendKey "<home>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>"
    'The following will generate a MAXIS formatted date for today. 
    current_day = DatePart("D", date)
    If len(current_day) = 1 then current_day = "0" & current_day
    current_month = DatePart("M", date)
    If len(current_month) = 1 then current_month = "0" & current_month
    current_year = right(DatePart("yyyy", date), 2)
    'Puts current date and pay frequency in PIC.
    EMWriteScreen current_month, 5, 34
    EMWriteScreen current_day, 5, 37
    EMWriteScreen current_year, 5, 40
    If pay_frequency = "One Time Per Month" then EMWriteScreen "1", 5, 64
    If pay_frequency = "Two Times Per Month" then EMWriteScreen "2", 5, 64
    If pay_frequency = "Every Other Week" then EMWriteScreen "3", 5, 64
    If pay_frequency = "Every Week" then EMWriteScreen "4", 5, 64
    'Sets PIC row for the next functions
    PIC_row = 9
    'Uses function to add each PIC pay date, income, and hours. Doesn't add any if they show "01/01/2000" as those are dummy numbers
    If pay_date_01 <> "01/01/2000" then call PIC_paystubs_info_adder(pay_date_01, gross_amt_01, hours_01)
    If pay_date_02 <> "01/01/2000" then call PIC_paystubs_info_adder(pay_date_02, gross_amt_02, hours_02)
    If pay_date_03 <> "01/01/2000" then call PIC_paystubs_info_adder(pay_date_03, gross_amt_03, hours_03)
    If pay_date_04 <> "01/01/2000" then call PIC_paystubs_info_adder(pay_date_04, gross_amt_04, hours_04)
    If pay_date_05 <> "01/01/2000" then call PIC_paystubs_info_adder(pay_date_05, gross_amt_05, hours_05)
    'Transmits in order to format the PIC
    transmit
    transmit  
    'Reads the contents of the PIC for case noting.
    EMReadScreen PIC_line_01, 26, 5, 49
    EMReadScreen PIC_line_02, 28, 8, 13
    EMReadScreen PIC_line_03, 28, 9, 13
    EMReadScreen PIC_line_04, 28, 10, 13
    EMReadScreen PIC_line_05, 28, 11, 13
    EMReadScreen PIC_line_06, 28, 12, 13
    EMReadScreen PIC_line_07, 28, 13, 13
    EMReadScreen PIC_line_08, 28, 14, 13
    EMReadScreen PIC_line_09, 50, 16, 22
    EMReadScreen PIC_line_10, 50, 17, 22
    EMReadScreen PIC_line_11, 50, 18, 22
    transmit
  End if
  
  
  'Clears JOBS data before updating the JOBS panel
  EMSetCursor 12, 25
  EMSendKey "___________________________________________________________________________________________________________________________________________________"
  
  'Updates for retrospective income by checking each pay date's month against the footer month using a function. If the footer month is two months ahead of the pay month it will add to JOBS and keep a tally of hours.
  MAXIS_row = 12 'Needs this for the following functions
  Dim retro_hours
  call retro_paystubs_info_adder(pay_date_01, gross_amt_01, hours_01)
  call retro_paystubs_info_adder(pay_date_02, gross_amt_02, hours_02)
  call retro_paystubs_info_adder(pay_date_03, gross_amt_03, hours_03)
  call retro_paystubs_info_adder(pay_date_04, gross_amt_04, hours_04)
  call retro_paystubs_info_adder(pay_date_05, gross_amt_05, hours_05)
  
  'Must convert retro hours into an integer for MAXIS
  retro_hours = retro_hours + .00000000000001 'This will force rounding to go half-up, as the CINT function rounds half down, which goes against procedure.
  retro_hours = cint(retro_hours)
  
  'Puts hours worked in the retro months in. This was determined using the previous functions.
  If retro_hours > 999 then retro_hours = 999 'In case there are over 999 hours, this is the procedure
  If retro_hours <> "" and retro_hours <> 0 then EMWriteScreen retro_hours, 18, 43
  retro_hours = 0 'Clears variable so it can be used in multiple months if needed
  
  'Determines the paydate to put in the prospective side. It moves forward for instances where the footer month is ahead of the first paydate, otherwise it moves backward until it lands on the right date.
  first_prospective_pay_date = pay_date_01
  If datediff("m", first_prospective_pay_date, footer_month & "/01/" & footer_year) > 0 then 'For instances where the footer month is ahead of the first paydate.
    Do
      If datediff("m", first_prospective_pay_date, footer_month & "/01/" & footer_year) = 0 then exit do
      If pay_frequency = "One Time Per Month" then first_prospective_pay_date = dateadd("m", 1, first_prospective_pay_date)
      If pay_frequency = "Two Times Per Month" then first_prospective_pay_date = dateadd("m", 1, first_prospective_pay_date)
      If pay_frequency = "Every Other Week" then first_prospective_pay_date = dateadd("d", 14, first_prospective_pay_date)
      If pay_frequency = "Every Week" then first_prospective_pay_date = dateadd("d", 7, first_prospective_pay_date)
    Loop until datediff("m", first_prospective_pay_date, footer_month & "/01/" & footer_year) = 0
  Elseif datediff("m", first_prospective_pay_date, footer_month & "/01/" & footer_year) < 0 then 'For instances where the footer month is behind the first paydate (ex: paydate is 06/26/2013 but footer month is 05/13).
    Do
      If datediff("m", first_prospective_pay_date, footer_month & "/01/" & footer_year) = 0 then exit do
      If pay_frequency = "One Time Per Month" then first_prospective_pay_date = dateadd("m", -1, first_prospective_pay_date)
      If pay_frequency = "Two Times Per Month" then first_prospective_pay_date = dateadd("m", -1, first_prospective_pay_date)
      If pay_frequency = "Every Other Week" then first_prospective_pay_date = dateadd("d", -14, first_prospective_pay_date)
      If pay_frequency = "Every Week" then first_prospective_pay_date = dateadd("d", -7, first_prospective_pay_date)
    Loop until datediff("m", first_prospective_pay_date, footer_month & "/01/" & footer_year) = 0
  End if
  
  'This checks to make sure the earliest possible paydate is selected in each prospective month.
  If pay_frequency = "Two Times Per Month" or pay_frequency = "Every Other Week" or pay_frequency = "Every Week" then 
    Do
      If pay_frequency = "Two Times Per Month" and datepart("d", first_prospective_pay_date) > 15 then first_prospective_pay_date = dateadd("d", -15, first_prospective_pay_date)
      If pay_frequency = "Every Other Week" and datepart("d", first_prospective_pay_date) > 14 then first_prospective_pay_date = dateadd("d", -14, first_prospective_pay_date)
      If pay_frequency = "Every Week" and datepart("d", first_prospective_pay_date) > 7 then first_prospective_pay_date = dateadd("d", -7, first_prospective_pay_date)
    Loop until (pay_frequency = "Two Times Per Month" and datepart("d", first_prospective_pay_date) <= 15) or (pay_frequency = "Every Other Week" and datepart("d", first_prospective_pay_date) <= 14) or (pay_frequency = "Every Week" and datepart("d", first_prospective_pay_date) <= 7)
  End if


  'Analyzes the paystubs received using a function, puts any actual paystubs received in the footer month into the JOBS panel on the prospective side.
  MAXIS_row = 12 'This variable is needed for the script to know which line to put the prospective info on
  call prospective_pay_analyzer(pay_date_01, gross_amt_01)
  call prospective_pay_analyzer(pay_date_02, gross_amt_02)
  call prospective_pay_analyzer(pay_date_03, gross_amt_03)
  call prospective_pay_analyzer(pay_date_04, gross_amt_04)
  call prospective_pay_analyzer(pay_date_05, gross_amt_05)
  total_prospective_dates = MAXIS_row - 12
  
  'Adds the remaining weeks in using a do...loop to determine all of the anticipated pay dates for the client.
  If pay_frequency = "One Time Per Month" then pay_multiplier = 31
  If pay_frequency = "Two Times Per Month" then pay_multiplier = 15
  If pay_frequency = "Every Other Week" then pay_multiplier = 14
  If pay_frequency = "Every Week" then pay_multiplier = 7
  Do
    If pay_frequency = "One Time Per Month" and total_prospective_dates >= 1 then exit do 'Shouldn't be more than one entry if pay is once per month.
    If pay_frequency = "Two Times Per Month" and total_prospective_dates >= 2 then exit do 'Shouldn't be more than two entries if pay is twice per month.
    prospective_pay_date = dateadd("d", total_prospective_dates * pay_multiplier, first_prospective_pay_date)
    If datediff("m", prospective_pay_date, footer_month & "/01/" & footer_year) = 0 then
      If len(datepart("m", prospective_pay_date)) = 2 then
        EMWriteScreen datepart("m", prospective_pay_date), MAXIS_row, 54
      Else
        EMWriteScreen "0" & datepart("m", prospective_pay_date), MAXIS_row, 54
      End if
      If len(datepart("d", prospective_pay_date)) = 2 then
        EMWriteScreen datepart("d", prospective_pay_date), MAXIS_row, 57
      Else
        EMWriteScreen "0" & datepart("d", prospective_pay_date), MAXIS_row, 57
      End if
      EMWriteScreen right(datepart("yyyy", prospective_pay_date), 2), MAXIS_row, 60
      EMWriteScreen average_pay_per_paystub, MAXIS_row, 67
      MAXIS_row = MAXIS_row + 1
      total_prospective_dates = total_prospective_dates + 1
    End if
  Loop until datediff("m", prospective_pay_date, footer_month & "/01/" & footer_year) <> 0
  
  
  'Updates pay frequency
  If pay_frequency = "One Time Per Month" then EMWriteScreen "1", 18, 35
  If pay_frequency = "Two Times Per Month" then EMWriteScreen "2", 18, 35
  If pay_frequency = "Every Other Week" then EMWriteScreen "3", 18, 35
  If pay_frequency = "Every Week" then EMWriteScreen "4", 18, 35
  
  'Puts average hours in. Added a small imperfection ".0000000000001" so that if any hourly amounts land on exactly ".5", they will round half-up instead of half down.
  If pay_frequency = "One Time Per Month" then EMWriteScreen cint(average_hours_per_paystub + .0000000000001), 18, 72
  If pay_frequency = "Two Times Per Month" then EMWriteScreen cint((average_hours_per_paystub + .0000000000001) * total_prospective_dates), 18, 72
  If pay_frequency = "Every Other Week" then EMWriteScreen cint((average_hours_per_paystub + .0000000000001) * total_prospective_dates), 18, 72
  If pay_frequency = "Every Week" then EMWriteScreen cint((average_hours_per_paystub + .0000000000001) * total_prospective_dates), 18, 72
  
  'Puts pay verification type in
  EMWriteScreen left(JOBS_verif_code, 1), 6, 38
  
  'If the footer month is the current month + 1, the script needs to update the HC popup for HC cases.
  If update_HC_popup_check = 1 and datediff("m", date, footer_month & "/01/" & footer_year) = 1 then 
    EMWriteScreen "x", 19, 54
    transmit
    EMWriteScreen "________", 11, 63
    EMWriteScreen average_pay_per_paystub, 11, 63
    Do 'Doing this as a pop-up since there are times when a warning message changes the amount of times this plays.
      transmit
      EMReadScreen HC_popup_check, 18, 9, 43
      If HC_popup_check <> "HC Income Estimate" then updated_HC_popup = True
    Loop until HC_popup_check <> "HC Income Estimate"
  End if
  
  'Transmits after ending the JOBS panel updating
  Do
    transmit
    EMReadScreen display_mode_check, 1, 20, 8
  Loop until display_mode_check = "D"
  
  If datediff("m", date, footer_month & "/01/" & footer_year) = 1 then in_future_month = True
  
  'If just on SNAP, the case does not have to update future months, so the script can now case note.
  If future_months_check = 0 or in_future_month = True then exit do
  
  'Navigates to the current month + 1 footer month, then back into the JOBS panel
  EMWriteScreen "bgtx", 20, 71
  transmit
  EMWriteScreen "y", 16, 54
  transmit
  EMReadScreen footer_month, 2, 20, 55
  EMReadScreen footer_year, 2, 20, 58
  EMWriteScreen "jobs", 20, 71
  EMWriteScreen HH_member, 20, 76
  If len(current_panel_number) = 1 then current_panel_number = "0" & current_panel_number
  EMWriteScreen current_panel_number, 20, 79
  transmit
  PF9

Loop until in_future_month = True

'Determines if the case note should add additional info about which HH member had the paystubs
If HH_member <> "01" then
  HH_memb_for_case_note = " for memb " & HH_member 
Else
  HH_memb_for_case_note = ""
End if

'Case noting section
If update_PIC_check = 1 then
  PF4
  PF9
  EMSendKey "~~~SNAP PIC" & HH_memb_for_case_note & ": " & date & "~~~" & "<newline>"
  EMSendKey PIC_line_02 & "<newline>"
  EMSendKey PIC_line_03 & "                 " & "<newline>"
  EMSendKey PIC_line_04 & "                 " & "<newline>"
  EMSendKey PIC_line_05 & "                 " & "<newline>"
  EMSendKey PIC_line_06 & "                 " & "<newline>"
  EMSendKey PIC_line_07 & "<newline>"
  EMSendKey PIC_line_08 & "<newline>"
  EMWriteScreen PIC_line_01, 6, 48
  EMWriteScreen PIC_line_09, 7, 35
  EMWriteScreen PIC_line_10, 8, 35
  EMWriteScreen PIC_line_11, 9, 35
  If explanation_of_income <> "" then 
    EMSendKey "---" & "<newline>"
    call write_editbox_in_case_note("How income was calculated", explanation_of_income, 6)
  End if
  call write_new_line_in_case_note("---")
  call write_new_line_in_case_note(worker_signature)
  PF3
  PF3
End if

If case_note_check = 1 then
  PF4
  PF9
  EMSendKey "Paystubs Received" & HH_memb_for_case_note & ": updated JOBS w/script" & "<newline>"
  call write_three_columns_in_case_note(14, "DATE", 29, "AMT", 39, "HOURS")
  If pay_date_01 <> "01/01/2000" then call write_three_columns_in_case_note(12, pay_date_01, 27, "$" & gross_amt_01, 39, hours_01)
  If pay_date_02 <> "01/01/2000" then call write_three_columns_in_case_note(12, pay_date_02, 27, "$" & gross_amt_02, 39, hours_02)
  If pay_date_03 <> "01/01/2000" then call write_three_columns_in_case_note(12, pay_date_03, 27, "$" & gross_amt_03, 39, hours_03)
  If pay_date_04 <> "01/01/2000" then call write_three_columns_in_case_note(12, pay_date_04, 27, "$" & gross_amt_04, 39, hours_04)
  If pay_date_05 <> "01/01/2000" then call write_three_columns_in_case_note(12, pay_date_05, 27, "$" & gross_amt_05, 39, hours_05)
  If explanation_of_income <> "" then 
    EMSendKey "---" & "<newline>"
    call write_editbox_in_case_note("How income was calculated", explanation_of_income, 6)
  End if
  call write_new_line_in_case_note("---")
  call write_new_line_in_case_note(worker_signature)
  PF3
  PF3
End if

MsgBox "Success! Your JOBS panel has been updated with the paystubs you've entered in. Send your case through background, review the results, and take action as appropriate. Don't forget to case note!" 
script_end_procedure("")

