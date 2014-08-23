'---------------------------------------------------------------------------------------------------
'HOW THIS SCRIPT WORKS:
'
'This script contains functions that the other BlueZone scripts use very commonly. The
'other BlueZone scripts contain a few lines of code that run this script and get the 
'functions. This saves time in writing and copy/pasting the same functions in
'many different places. Only add functions to this script if they've been tested by
'the workgroups. This document is actively used by live scripts, so it needs to be
'functionally complete at all times.
'
'Here's the code to add, including stats gathering pieces (without comments of course):
'
'GATHERING STATS----------------------------------------------------------------------------------------------------
'name_of_script = ""
'start_time = timer
'
''LOADING ROUTINE FUNCTIONS
'Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
'Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
'text_from_the_other_script = fso_command.ReadAll
'fso_command.Close
'Execute text_from_the_other_script

'----------------------------------------------------------------------------------------------------

'COUNTY CUSTOM VARIABLES----------------------------------------------------------------------------------------------------

worker_county_code = "x102"
collecting_statistics = False
EDMS_choice = "Compass Forms"
county_name = "Anoka"
county_address_line_01 = "1234 Anoka Road"
county_address_line_02 = "Anoka, MN 55555"
case_noting_intake_dates = True
move_verifs_needed = False

is_county_collecting_stats = collecting_statistics	'IT DOES THIS BECAUSE THE SETUP SCRIPT WILL OVERWRITE LINES BELOW WHICH DEPEND ON THIS, BY SEPARATING THE VARIABLES WE PREVENT ISSUES

'SHARED VARIABLES----------------------------------------------------------------------------------------------------
checked = 1		'Value for checked boxes
unchecked = 0		'Value for unchecked boxes
cancel = 0		'Value for cancel button in dialogs
OK = -1			'Value for OK button in dialogs

'Some screens require the two digit county code, and this determines what that code is
two_digit_county_code = right(worker_county_code, 2)
If two_digit_county_code = "PW" then two_digit_county_code = "91"	'For DHS purposes



'----------------------------------------------------------------------------------------------------

Function add_ACCI_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen ACCI_date, 8, 6, 73
  ACCI_date = replace(ACCI_date, " ", "/")
  If datediff("yyyy", ACCI_date, now) < 5 then
    EMReadScreen ACCI_type, 2, 6, 47
    If ACCI_type = "01" then ACCI_type = "Auto"
    If ACCI_type = "02" then ACCI_type = "Workers Comp"
    If ACCI_type = "03" then ACCI_type = "Homeowners"
    If ACCI_type = "04" then ACCI_type = "No Fault"
    If ACCI_type = "05" then ACCI_type = "Other Tort"
    If ACCI_type = "06" then ACCI_type = "Product Liab"
    If ACCI_type = "07" then ACCI_type = "Med Malprac"
    If ACCI_type = "08" then ACCI_type = "Legal Malprac"
    If ACCI_type = "09" then ACCI_type = "Diving Tort"
    If ACCI_type = "10" then ACCI_type = "Motorcycle"
    If ACCI_type = "11" then ACCI_type = "MTC or Other Bus Tort"
    If ACCI_type = "12" then ACCI_type = "Pedestrian"
    If ACCI_type = "13" then ACCI_type = "Other"
    x = x & ACCI_type & " on " & ACCI_date & ".; "
  End if
End function

Function add_ACCT_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen ACCT_amt, 8, 10, 46
  ACCT_amt = trim(ACCT_amt)
  ACCT_amt = "$" & ACCT_amt
  EMReadScreen ACCT_type, 2, 6, 44
  EMReadScreen ACCT_location, 20, 8, 44
  ACCT_location = replace(ACCT_location, "_", "")
  ACCT_location = split(ACCT_location)
  For each a in ACCT_location
    If a <> "" then
      b = ucase(left(a, 1))
      c = LCase(right(a, len(a) -1))
      If len(a) > 3 then
        new_ACCT_location = new_ACCT_location & b & c & " "
      Else
        new_ACCT_location = new_ACCT_location & a & " "
      End if
    End if
  Next
  EMReadScreen ACCT_ver, 1, 10, 63
  If ACCT_ver = "N" then 
    ACCT_ver = ", no proof provided"
  Else
    ACCT_ver = ""
  End if
  x = x & ACCT_type & " at " & new_ACCT_location & "(" & ACCT_amt & ")" & ACCT_ver & ".; "
  new_ACCT_location = ""
End function

Function add_BUSI_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen BUSI_type, 2, 5, 37
  If BUSI_type = "01" then BUSI_type = "Farming"
  If BUSI_type = "02" then BUSI_type = "Real Estate"
  If BUSI_type = "03" then BUSI_type = "Home Product Sales"
  If BUSI_type = "04" then BUSI_type = "Other Sales"
  If BUSI_type = "05" then BUSI_type = "Personal Services"
  If BUSI_type = "06" then BUSI_type = "Paper Route"
  If BUSI_type = "07" then BUSI_type = "InHome Daycare"
  If BUSI_type = "08" then BUSI_type = "Rental Income"
  If BUSI_type = "09" then BUSI_type = "Other"
  EMWriteScreen "x", 7, 26
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  If cash_check = 1 then
    EMReadScreen BUSI_ver, 1, 9, 73
  ElseIf HC_check = 1 then 
    EMReadScreen BUSI_ver, 1, 12, 73
    If BUSI_ver = "_" then EMReadScreen BUSI_ver, 1, 13, 73
  ElseIf SNAP_check = 1 then
    EMReadScreen BUSI_ver, 1, 11, 73
  End if
  EMSendKey "<PF3>"
  EMWaitReady 0, 0
  If SNAP_check = 1 then
    EMReadScreen BUSI_amt, 8, 11, 68
    BUSI_amt = trim(BUSI_amt)
  ElseIf cash_check = 1 then 
    EMReadScreen BUSI_amt, 8, 9, 54
    BUSI_amt = trim(BUSI_amt)
  ElseIf HC_check = 1 then 
    EMWriteScreen "x", 17, 29
    EMSendKey "<enter>"
    EMWaitReady 0, 0
    EMReadScreen BUSI_amt, 8, 15, 54
    If BUSI_amt = "    0.00" then EMReadScreen BUSI_amt, 8, 16, 54
    BUSI_amt = trim(BUSI_amt)
    EMSendKey "<PF3>"
    EMWaitReady 0, 0
  End if
  x = x & trim(BUSI_type) & " BUSI"
  EMReadScreen BUSI_income_end_date, 8, 5, 71
  If BUSI_income_end_date <> "__ __ __" then BUSI_income_end_date = replace(BUSI_income_end_date, " ", "/")
  If IsDate(BUSI_income_end_date) = True then
    x = x & " (ended " & BUSI_income_end_date & ")"
  Else
    If BUSI_amt <> "" then x = x & ", ($" & BUSI_amt & "/monthly)"
  End if
  If BUSI_ver = "N" or BUSI_ver = "?" then 
    x = x & ", no proof provided.; "
  Else
    x = x & ".; "
  End if
End function

Function add_CARS_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen CARS_year, 4, 8, 31
  EMReadScreen CARS_make, 15, 8, 43
  CARS_make = replace(CARS_make, "_", "")
  EMReadScreen CARS_model, 15, 8, 66
  CARS_model = replace(CARS_model, "_", "")
  CARS_type = CARS_year & " " & CARS_make & " " & CARS_model
  CARS_type = split(CARS_type)
  For each a in CARS_type
    If len(a) > 1 then
      b = ucase(left(a, 1))
      c = LCase(right(a, len(a) -1))
      new_CARS_type = new_CARS_type & b & c & " "
    End if
  Next
  EMReadScreen CARS_amt, 8, 9, 45
  CARS_amt = trim(CARS_amt)
  CARS_amt = "$" & CARS_amt
  x = x & trim(new_CARS_type) & ", (" & CARS_amt & "); "
  new_CARS_type = ""
End function

Function add_JOBS_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen JOBS_type, 30, 7, 42
  JOBS_type = replace(JOBS_type, "_", ""	)
  JOBS_type = trim(JOBS_type)
  JOBS_type = split(JOBS_type)
  For each a in JOBS_type
    If a <> "" then
      b = ucase(left(a, 1))
      c = LCase(right(a, len(a) -1))
      new_JOBS_type = new_JOBS_type & b & c & " "
    End if
  Next
  If SNAP_check = 1 then
    EMWriteScreen "x", 19, 38
    EMSendKey "<enter>"
    EMWaitReady 0, 0
    EMReadScreen SNAP_JOBS_amt, 8, 17, 56
    SNAP_JOBS_amt = trim(SNAP_JOBS_amt)
    EMReadScreen pay_frequency, 1, 5, 64
    EMSendKey "<enter>"
    EMWaitReady 0, 0
  ElseIf cash_check = 1 then
    EMReadScreen retro_JOBS_amt, 8, 17, 38
    retro_JOBS_amt = trim(retro_JOBS_amt)
  ElseIf HC_check = 1 then 
    EMReadScreen pay_frequency, 1, 18, 35
    EMWriteScreen "x", 19, 54
    EMSendKey "<enter>"
    EMWaitReady 0, 0
    EMReadScreen HC_JOBS_amt, 8, 11, 63
    HC_JOBS_amt = trim(HC_JOBS_amt)
    EMSendKey "<enter>"
    EMWaitReady 0, 0
  End If
  EMReadScreen JOBS_ver, 1, 6, 38
  EMReadScreen JOBS_income_end_date, 8, 9, 49
  If JOBS_income_end_date <> "__ __ __" then JOBS_income_end_date = replace(JOBS_income_end_date, " ", "/")
  If IsDate(JOBS_income_end_date) = True then
    x = x & new_JOBS_type & "(ended " & JOBS_income_end_date
  Else
    If pay_frequency = "1" then pay_frequency = "monthly"
    If pay_frequency = "2" then pay_frequency = "semimonthly"
    If pay_frequency = "3" then pay_frequency = "biweekly"
    If pay_frequency = "4" then pay_frequency = "weekly"
    If pay_frequency = "_" or pay_frequency = "5" then pay_frequency = "non-monthly"
    x = x & "EI from " & trim(new_JOBS_type)
    If SNAP_check = 1 then
      x = x & ", ($" & SNAP_JOBS_amt & "/" & pay_frequency
    ElseIf cash_check = 1 then
      x = x & ", ($" & retro_JOBS_amt & " budgeted"
    ElseIf HC_check = 1 then 
      x = x & ", ($" & HC_JOBS_amt & "/" & pay_frequency 
    End if
  End if
  If JOBS_ver = "N" or JOBS_ver = "?" then
    x = x & ", no proof provided).; "
  Else
    x = x & ").; "
  End if
End function

Function add_OTHR_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen OTHR_type, 16, 6, 43
  OTHR_type = trim(OTHR_type)
  EMReadScreen OTHR_amt, 10, 8, 40
  OTHR_amt = trim(OTHR_amt)
  OTHR_amt = "$" & OTHR_amt
  x = x & trim(OTHR_type) & ", (" & OTHR_amt & ").; "
  new_OTHR_type = ""
End function

Function add_RBIC_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen RBIC_type, 16, 5, 48
  RBIC_type = trim(RBIC_type)
  EMReadScreen RBIC_amt, 8, 10, 62
  RBIC_amt = trim(RBIC_amt)
  EMReadScreen RBIC_ver, 1, 10, 76
  If RBIC_ver = "N" then RBIC_ver = ", no proof provided"
  EMReadScreen RBIC_end_date, 8, 6, 68
  RBIC_end_date = replace(RBIC_end_date, " ", "/")
  If isdate(RBIC_end_date) = True then
    x = x & trim(RBIC_type) & " RBIC, ended " & RBIC_end_date & RBIC_ver & ".; "
  Else
    x = x & trim(RBIC_type) & " RBIC, ($" & RBIC_amt & RBIC_ver & ").; "
  End if
End function

Function add_REST_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen REST_type, 16, 6, 41
  REST_type = trim(REST_type)
  EMReadScreen REST_amt, 10, 8, 41
  REST_amt = trim(REST_amt)
  REST_amt = "$" & REST_amt
  x = x & trim(REST_type) & ", (" & REST_amt & ").; "
  new_REST_type = ""
End function


Function add_SECU_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen SECU_amt, 8, 10, 52
  SECU_amt = trim(SECU_amt)
  SECU_amt = "$" & SECU_amt
  EMReadScreen SECU_type, 2, 6, 50
  EMReadScreen SECU_location, 20, 8, 50
  SECU_location = replace(SECU_location, "_", "")
  SECU_location = split(SECU_location)
  For each a in SECU_location
    If a <> "" then
      b = ucase(left(a, 1))
      c = LCase(right(a, len(a) -1))
      If len(a) > 3 then
        new_SECU_location = new_SECU_location & b & c & " "
      Else
        new_SECU_location = new_SECU_location & a & " "
      End if
    End if
  Next
  EMReadScreen SECU_ver, 1, 11, 50
  If SECU_ver = "1" then SECU_ver = "agency form provided"
  If SECU_ver = "2" then SECU_ver = "source doc provided"
  If SECU_ver = "3" then SECU_ver = "verified via phone"
  If SECU_ver = "5" then SECU_ver = "other doc verified"
  If SECU_ver = "N" then SECU_ver = "no proof provided"
  x = x & SECU_type & " at " & new_SECU_location & " (" & SECU_amt & "), " & SECU_ver & ".; "
  new_SECU_location = ""
End function

Function add_UNEA_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen UNEA_type, 16, 5, 40
  If UNEA_type = "Unemployment Ins" then UNEA_type = "UC"
  If UNEA_type = "Disbursed Child " then UNEA_type = "CS"
  If UNEA_type = "Disbursed CS Arr" then UNEA_type = "CS arrears"
  UNEA_type = trim(UNEA_type)
  EMReadScreen UNEA_ver, 1, 5, 65
  EMReadScreen UNEA_income_end_date, 8, 7, 68
  If UNEA_income_end_date <> "__ __ __" then UNEA_income_end_date = replace(UNEA_income_end_date, " ", "/")
  If IsDate(UNEA_income_end_date) = True then
    x = x & UNEA_type & " (ended " & UNEA_income_end_date
  Else
    EMReadScreen UNEA_amt, 8, 18, 68
    UNEA_amt = trim(UNEA_amt)
    If SNAP_check = 1 then
      EMWriteScreen "x", 10, 26
      EMSendKey "<enter>"
      EMWaitReady 0, 0
      EMReadScreen SNAP_UNEA_amt, 8, 17, 56
      SNAP_UNEA_amt = trim(SNAP_UNEA_amt)
      EMReadScreen pay_frequency, 1, 5, 64
      EMSendKey "<enter>"
      EMWaitReady 0, 0
    ElseIf cash_check = 1 then
      EMReadScreen retro_UNEA_amt, 8, 18, 39
      retro_UNEA_amt = trim(retro_UNEA_amt)
      if retro_UNEA_amt = "" then retro_UNEA_amt = "0"
    ElseIf HC_check = 1 then 
      EMWriteScreen "x", 6, 56
      EMSendKey "<enter>"
      EMWaitReady 0, 0
      EMReadScreen HC_UNEA_amt, 8, 9, 65
      HC_UNEA_amt = trim(HC_UNEA_amt)
      EMReadScreen pay_frequency, 1, 10, 63
      EMSendKey "<enter>"
      EMWaitReady 0, 0
      If HC_UNEA_amt = "________" then
        EMReadScreen HC_UNEA_amt, 8, 18, 68
        HC_UNEA_amt = trim(HC_UNEA_amt)
        pay_frequency = "mo budgeted prospectively"
      End if
    End If
    If pay_frequency = "1" then pay_frequency = "monthly"
    If pay_frequency = "2" then pay_frequency = "semimonthly"
    If pay_frequency = "3" then pay_frequency = "biweekly"
    If pay_frequency = "4" then pay_frequency = "weekly"
    If pay_frequency = "_" then pay_frequency = "non-monthly"
    x = x & trim(UNEA_type)
    If SNAP_check = 1 then
      x = x & ", ($" & SNAP_UNEA_amt & "/" & pay_frequency
    ElseIf cash_check = 1 then
      if retro_UNEA_amt = "0" then
        EMReadScreen pro_UNEA_amt, 8, 18, 68
        pro_UNEA_amt = trim(pro_UNEA_amt)
        If pro_UNEA_amt = "" then pro_UNEA_amt = "0"
        x = x & ", ($" & pro_UNEA_amt & " budgeted prospectively"
      Else
        x = x & ", ($" & retro_UNEA_amt & " budgeted retrospectively"
      End if
    ElseIf HC_check = 1 then 
      x = x & ", ($" & HC_UNEA_amt & "/" & pay_frequency
    End if
  End if
  If UNEA_ver = "N" or UNEA_ver = "?" then
    x = x & ", no proof provided).; "
  Else
    x = x & ").; "
  End if
End function

Function attn
  EMSendKey "<attn>"
  EMWaitReady -1, 0
End function

Function autofill_editbox_from_MAXIS(HH_member_array, panel_read_from, variable_written_to)
  If panel_read_from = "ABPS" then '--------------------------------------------------------------------------------------------------------ABPS
    call navigate_to_screen("stat", "ABPS")
    EMReadScreen ABPS_total_pages, 1, 2, 78
    If ABPS_total_pages <> 0 then 
      Do
        'First it checks the support coop. If it's "N" it'll add a blurb about it to the support_coop variable
        EMReadScreen support_coop_code, 1, 4, 73
        If support_coop_code = "N" then
          EMReadScreen caregiver_ref_nbr, 2, 4, 47
          If instr(support_coop, "Memb " & caregiver_ref_nbr & " not cooperating with child support; ") = 0 then support_coop = support_coop & "Memb " & caregiver_ref_nbr & " not cooperating with child support; "'the if...then statement makes sure the info isn't duplicated. 
        End if
        'Then it gets info on the ABPS themself.
        EMReadScreen ABPS_current, 45, 10, 30
        If ABPS_current = "________________________  First: ____________" then ABPS_current = "Parent unknown"
        ABPS_current = replace(ABPS_current, "  First:", ",")
        ABPS_current = replace(ABPS_current, "_", "")
        ABPS_current = split(ABPS_current)
        For each a in ABPS_current
          b = ucase(left(a, 1))
          c = LCase(right(a, len(a) -1))
          If len(a) > 1 then
            new_ABPS_current = new_ABPS_current & b & c & " "
          Else
            new_ABPS_current = new_ABPS_current & a & " "
          End if
        Next
        ABPS_row = 15 'Setting variable for do...loop
        Do 'Using a do...loop to determine which MEMB numbers are with this parent
          EMReadScreen child_ref_nbr, 2, ABPS_row, 35
          If child_ref_nbr <> "__" then
            amt_of_children_for_ABPS = amt_of_children_for_ABPS + 1
            children_for_ABPS = children_for_ABPS & child_ref_nbr & ", "
          End if
          ABPS_row = ABPS_row + 1
        Loop until ABPS_row > 17
        'Cleaning up the "children_for_ABPS" variable to be more readable
        children_for_ABPS = left(children_for_ABPS, len(children_for_ABPS) - 2) 'cleaning up the end of the variable (removing the comma for single kids)
        children_for_ABPS = strreverse(children_for_ABPS)                       'flipping it around to change the last comma to an "and"
        children_for_ABPS = replace(children_for_ABPS, ",", "dna ", 1, 1)        'it's backwards, replaces just one comma with an "and"
        children_for_ABPS = strreverse(children_for_ABPS)                       'flipping it back around 
        if amt_of_children_for_ABPS > 1 then HH_memb_title = " for membs "
        if amt_of_children_for_ABPS <= 1 then HH_memb_title = " for memb "
        variable_written_to = variable_written_to & trim(new_ABPS_current) & HH_memb_title & children_for_ABPS & "; "
        'Resetting variables for the do...loop in case this function runs again
        new_ABPS_current = "" 
        amt_of_children_for_ABPS = 0
        children_for_ABPS = ""
        'Checking to see if it needs to run again, if it does it transmits or else the loop stops
        EMReadScreen ABPS_current_page, 1, 2, 73
        If ABPS_current_page <> ABPS_total_pages then transmit
      Loop until ABPS_current_page = ABPS_total_pages
      'Combining the two variables (support coop and the variable written to)
      variable_written_to = support_coop & variable_written_to
    End if
  Elseif panel_read_from = "ACCI" then '----------------------------------------------------------------------------------------------------ACCI
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "ACCI")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen ACCI_total, 1, 2, 78
      If ACCI_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_ACCI_to_variable(variable_written_to)
          EMReadScreen ACCI_panel_current, 1, 2, 73
          If cint(ACCI_panel_current) < cint(ACCI_total) then transmit
        Loop until cint(ACCI_panel_current) = cint(ACCI_total)
      End if
    Next
  Elseif panel_read_from = "ACCT" then '----------------------------------------------------------------------------------------------------ACCT
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "acct")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen ACCT_total, 1, 2, 78
      If ACCT_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_ACCT_to_variable(variable_written_to)
          EMReadScreen ACCT_panel_current, 1, 2, 73
          If cint(ACCT_panel_current) < cint(ACCT_total) then transmit
        Loop until cint(ACCT_panel_current) = cint(ACCT_total)
      End if
    Next
  Elseif panel_read_from = "ADDR" then '----------------------------------------------------------------------------------------------------ADDR
    call navigate_to_screen("stat", "addr")
    EMReadScreen addr_line_01, 22, 6, 43
    EMReadScreen addr_line_02, 22, 7, 43
    EMReadScreen city_line, 15, 8, 43
    EMReadScreen state_line, 2, 8, 66
    EMReadScreen zip_line, 12, 9, 43
    variable_written_to = replace(addr_line_01, "_", "") & "; " & replace(addr_line_02, "_", "") & "; " & replace(city_line, "_", "") & ", " & state_line & " " & replace(zip_line, "__ ", "-")
    variable_written_to = replace(variable_written_to, "; ; ", "; ") 'in case there's only one line on ADDR
  Elseif panel_read_from = "AREP" then '----------------------------------------------------------------------------------------------------AREP
    call navigate_to_screen("stat", "arep")
    EMReadScreen AREP_name, 37, 4, 32
    AREP_name = replace(AREP_name, "_", "")
    AREP_name = split(AREP_name)
    For each word in AREP_name
      If word <> "" then
        first_letter_of_word = ucase(left(word, 1))
        rest_of_word = LCase(right(word, len(word) -1))
        If len(word) > 2 then
          variable_written_to = variable_written_to & first_letter_of_word & rest_of_word & " "
        Else
          variable_written_to = variable_written_to & word & " "
        End if
      End if
    Next
  Elseif panel_read_from = "BILS" then '----------------------------------------------------------------------------------------------------BILS
    call navigate_to_screen("stat", "bils")
    EMReadScreen BILS_amt, 1, 2, 78
    If BILS_amt <> 0 then variable_written_to = "BILS known to MAXIS."
  Elseif panel_read_from = "BUSI" then '----------------------------------------------------------------------------------------------------BUSI
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "busi")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen BUSI_total, 1, 2, 78
      If BUSI_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_BUSI_to_variable(variable_written_to)
          EMReadScreen BUSI_panel_current, 1, 2, 73
          If cint(BUSI_panel_current) < cint(BUSI_total) then transmit
        Loop until cint(BUSI_panel_current) = cint(BUSI_total)
      End if
    Next
  Elseif panel_read_from = "CARS" then '----------------------------------------------------------------------------------------------------CARS
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "cars")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen CARS_total, 1, 2, 78
      If CARS_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_CARS_to_variable(variable_written_to)
          EMReadScreen CARS_panel_current, 1, 2, 73
          If cint(CARS_panel_current) < cint(CARS_total) then transmit
        Loop until cint(CARS_panel_current) = cint(CARS_total)
      End if
    Next
  Elseif panel_read_from = "CASH" then '----------------------------------------------------------------------------------------------------CASH
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "cash")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen cash_amt, 8, 8, 39
      cash_amt = trim(cash_amt)
      If cash_amt <> "________" then
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Cash ($" & cash_amt & "); "
      End if
    Next
  Elseif panel_read_from = "COEX" then '----------------------------------------------------------------------------------------------------COEX
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "coex")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen support_amt, 8, 10, 63
      support_amt = trim(support_amt)
      If support_amt <> "________" then
        EMReadScreen support_ver, 1, 10, 36
        If support_ver = "?" or support_ver = "N" then
          support_ver = ", no proof provided"
        Else
          support_ver = ""
        End if
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Support ($" & support_amt & "/mo" & support_ver & "); "
      End if
      EMReadScreen alimony_amt, 8, 11, 63
      alimony_amt = trim(alimony_amt)
      If alimony_amt <> "________" then
        EMReadScreen alimony_ver, 1, 11, 36
        If alimony_ver = "?" or alimony_ver = "N" then
          alimony_ver = ", no proof provided"
        Else
          alimony_ver = ""
        End if
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Alimony ($" & alimony_amt & "/mo" & alimony_ver & "); "
      End if
      EMReadScreen tax_dep_amt, 8, 12, 63
      tax_dep_amt = trim(tax_dep_amt)
      If tax_dep_amt <> "________" then
        EMReadScreen tax_dep_ver, 1, 12, 36
        If tax_dep_ver = "?" or tax_dep_ver = "N" then
          tax_dep_ver = ", no proof provided"
        Else
          tax_dep_ver = ""
        End if
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Tax dep ($" & tax_dep_amt & "/mo" & tax_dep_ver & "); "
      End if
      EMReadScreen other_COEX_amt, 8, 13, 63
      other_COEX_amt = trim(other_COEX_amt)
      If other_COEX_amt <> "________" then
        EMReadScreen other_COEX_ver, 1, 13, 36
        If other_COEX_ver = "?" or other_COEX_ver = "N" then
          other_COEX_ver = ", no proof provided"
        Else
          other_COEX_ver = ""
        End if
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Other ($" & other_COEX_amt & "/mo" & other_COEX_ver & "); "
      End if
    Next
  Elseif panel_read_from = "DCEX" then '----------------------------------------------------------------------------------------------------DCEX
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "dcex")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      DCEX_row = 11
      Do
      EMReadScreen expense_amt, 8, DCEX_row, 63
      expense_amt = trim(expense_amt)
      If expense_amt <> "________" then
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        EMReadScreen child_ref_nbr, 2, DCEX_row, 29
        EMReadScreen expense_ver, 1, DCEX_row, 41
        If expense_ver = "?" or expense_ver = "N" or expense_ver = "_" then
          expense_ver = ", no proof provided"
        Else
          expense_ver = ""
        End if
        variable_written_to = variable_written_to & "Child " & child_ref_nbr & " ($" & expense_amt & "/mo DCEX" & expense_ver & "); "
      End if
      DCEX_row = DCEX_row + 1
      Loop until DCEX_row = 17
    Next
  Elseif panel_read_from = "DIET" then '----------------------------------------------------------------------------------------------------DIET
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "diet")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      DIET_row = 8 'Setting this variable for the next do...loop
      EMReadScreen DIET_total, 1, 2, 78
      If DIET_total <> 0 then 
        If HH_member <> "01" then DIET = DIET & "Member " & HH_member & "- "
        Do
          EMReadScreen diet_type, 2, DIET_row, 40
          EMReadScreen diet_proof, 1, DIET_row, 51
          If diet_proof = "_" or diet_proof = "?" or diet_proof = "N" then 
            diet_proof = ", no proof provided"
          Else
            diet_proof = ""
          End if
          If diet_type = "01" then diet_type = "High Protein"
          If diet_type = "02" then diet_type = "Cntrl Protein (40-60 g/day)"
          If diet_type = "03" then diet_type = "Cntrl Protein (<40 g/day)"
          If diet_type = "04" then diet_type = "Lo Cholesterol"
          If diet_type = "05" then diet_type = "High Residue"
          If diet_type = "06" then diet_type = "Preg/Lactation"
          If diet_type = "07" then diet_type = "Gluten Free"
          If diet_type = "08" then diet_type = "Lactose Free"
          If diet_type = "09" then diet_type = "Anti-Dumping"
          If diet_type = "10" then diet_type = "Hypoglycemic"
          If diet_type = "11" then diet_type = "Ketogenic"
          If diet_type <> "__" and diet_type <> "  " then variable_written_to = variable_written_to & diet_type & diet_proof & "; "
          DIET_row = DIET_row + 1
        Loop until DIET_row = 19
      End if
    Next
  Elseif panel_read_from = "DISA" then '----------------------------------------------------------------------------------------------------DISA
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "disa")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen DISA_status, 2, 13, 59
      If DISA_status = "01" or DISA_status = "02" or DISA_status = "03" or DISA_status = "04" then DISA_status = "RSDI/SSI certified"
      If DISA_status = "06" then DISA_status = "SMRT/SSA pends"
      If DISA_status = "08" then DISA_status = "Certified blind"
      If DISA_status = "10" then DISA_status = "Certified disabled"
      If DISA_status = "11" then DISA_status = "Spec cat- disa child"
      If DISA_status = "20" then DISA_status = "TEFRA- disabled"
      If DISA_status = "21" then DISA_status = "TEFRA- blind"
      If DISA_status = "22" then DISA_status = "MA-EPD"
      If DISA_status = "23" then DISA_status = "MA/waiver"
      If DISA_status = "24" then DISA_status = "SSA/SMRT appeal pends"
      If DISA_status = "26" then DISA_status = "SSA/SMRT disa deny"
      If DISA_status = "__" then
        DISA_status = ""
      Else
        EMReadScreen DISA_ver, 1, 13, 69
        If DISA_ver = "?" or DISA_ver = "N" then
          DISA_proof_type = ", no proof provided"
        Else
          DISA_proof_type = ""
        End if
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & DISA_status & DISA_proof_type & "; "
      End if
    Next
  Elseif panel_read_from = "EATS" then '----------------------------------------------------------------------------------------------------EATS
    call navigate_to_screen("stat", "eats")
    row = 14
    Do
      EMReadScreen reference_numbers_current_row, 40, row, 39
      reference_numbers = reference_numbers + reference_numbers_current_row  
      row = row + 1
    Loop until row = 18
    reference_numbers = replace(reference_numbers, "  ", " ")
    reference_numbers = split(reference_numbers)
    For each member in reference_numbers
      If member <> "__" and member <> "" then EATS_info = EATS_info & member & ", "
    Next
    EATS_info = trim(EATS_info)
    if right(EATS_info, 1) = "," then EATS_info = left(EATS_info, len(EATS_info) - 1)
    If EATS_info <> "" then variable_written_to = variable_written_to & ", p/p sep from memb(s) " & EATS_info & "."
  Elseif panel_read_from = "FACI" then '----------------------------------------------------------------------------------------------------FACI
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "faci")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen FACI_total, 1, 2, 78
      If FACI_total <> 0 then
        row = 14
        Do
          EMReadScreen date_in_check, 4, row, 53
          EMReadScreen date_out_check, 4, row, 77
          If (date_in_check <> "____" and date_out_check <> "____") or (date_in_check = "____" and date_out_check = "____") then row = row + 1
          If row > 18 then
            EMReadScreen FACI_page, 1, 2, 73
            If FACI_page = FACI_total then 
              FACI_status = "Not in facility"
            Else
              transmit
              row = 14
            End if
          End if
        Loop until (date_in_check <> "____" and date_out_check = "____") or FACI_status = "Not in facility"
        EMReadScreen client_FACI, 30, 6, 43
        client_FACI = replace(client_FACI, "_", "")
        FACI_array = split(client_FACI)
        For each a in FACI_array
          If a <> "" then
            b = ucase(left(a, 1))
            c = LCase(right(a, len(a) -1))
            new_FACI = new_FACI & b & c & " "
          End if
        Next
        client_FACI = new_FACI
        If FACI_status = "Not in facility" then
          client_FACI = ""
        Else
          If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
          variable_written_to = variable_written_to & client_FACI & "; "
        End if
      End if
    Next
  Elseif panel_read_from = "FMED" then '----------------------------------------------------------------------------------------------------FMED
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "fmed")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      fmed_row = 9 'Setting this variable for the next do...loop
      EMReadScreen fmed_total, 1, 2, 78
      If fmed_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          EMReadScreen fmed_type, 2, fmed_row, 25
          EMReadScreen fmed_proof, 2, fmed_row, 32
          EMReadScreen fmed_amt, 8, fmed_row, 70
          If fmed_proof = "__" or fmed_proof = "?_" or fmed_proof = "NO" then 
            fmed_proof = ", no proof provided"
          Else
            fmed_proof = ""
          End if
          If fmed_amt = "________" then
            fmed_amt = ""
          Else
            fmed_amt = " ($" & trim(fmed_amt) & ")"
          End if
          If fmed_type = "01" then fmed_type = "Nursing Home"
          If fmed_type = "02" then fmed_type = "Hosp/Clinic"
          If fmed_type = "03" then fmed_type = "Physicians"
          If fmed_type = "04" then fmed_type = "Prescriptions"
          If fmed_type = "05" then fmed_type = "Ins Premiums"
          If fmed_type = "06" then fmed_type = "Dental"
          If fmed_type = "07" then fmed_type = "Medical Trans/Flat Amt"
          If fmed_type = "08" then fmed_type = "Vision Care"
          If fmed_type = "09" then fmed_type = "Medicare Prem"
          If fmed_type = "10" then fmed_type = "Mo. Spdwn Amt/Waiver Obl"
          If fmed_type = "11" then fmed_type = "Home Care"
          If fmed_type = "12" then fmed_type = "Medical Trans/Mileage Calc"
          If fmed_type = "15" then fmed_type = "Medi Part D premium"
          If fmed_type <> "__" then variable_written_to = variable_written_to & fmed_type & fmed_amt & fmed_proof & "; "
          fmed_row = fmed_row + 1
          If fmed_row = 15 then
            PF20
            fmed_row = 9
            EMReadScreen last_page_check, 21, 24, 2
            If last_page_check <> "THIS IS THE LAST PAGE" then last_page_check = ""
          End if
        Loop until fmed_type = "__" or last_page_check = "THIS IS THE LAST PAGE"
      End if
    Next
  Elseif panel_read_from = "HCRE" then '----------------------------------------------------------------------------------------------------HCRE
    call navigate_to_screen("stat", "hcre")
    EMReadScreen variable_written_to, 8, 10, 51
    variable_written_to = replace(variable_written_to, " ", "/")
    If variable_written_to = "__/__/__" then EMReadScreen variable_written_to, 8, 11, 51
    variable_written_to = replace(variable_written_to, " ", "/")
    If isdate(variable_written_to) = True then variable_written_to = cdate(variable_written_to) & ""
    If isdate(variable_written_to) = False then variable_written_to = ""
  Elseif panel_read_from = "HCRE-retro" then '----------------------------------------------------------------------------------------------HCRE-retro
    call navigate_to_screen("stat", "hcre")
    EMReadScreen variable_written_to, 5, 10, 64
    If isdate(variable_written_to) = True then
      variable_written_to = replace(variable_written_to, " ", "/01/")
      If DatePart("m", variable_written_to) <> DatePart("m", CAF_datestamp) or DatePart("yyyy", variable_written_to) <> DatePart("yyyy", CAF_datestamp) then
        variable_written_to = variable_written_to
      Else
        variable_written_to = ""
      End if
    End if
  Elseif panel_read_from = "HEST" then '----------------------------------------------------------------------------------------------------HEST
    call navigate_to_screen("stat", "hest")
    EMReadScreen HEST_total, 1, 2, 78
    If HEST_total <> 0 then 
      EMReadScreen heat_air_check, 6, 13, 75
      If heat_air_check <> "      " then variable_written_to = variable_written_to & "Heat/AC.; "
      EMReadScreen electric_check, 6, 14, 75
      If electric_check <> "      " then variable_written_to = variable_written_to & "Electric.; "
      EMReadScreen phone_check, 6, 15, 75
      If phone_check <> "      " then variable_written_to = variable_written_to & "Phone.; "
    End if
  Elseif panel_read_from = "IMIG" then '----------------------------------------------------------------------------------------------------IMIG
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "IMIG")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen IMIG_total, 1, 2, 78
      If IMIG_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        EMReadScreen IMIG_type, 30, 6, 48
        variable_written_to = variable_written_to & trim(IMIG_type) & "; "
      End if
    Next
  Elseif panel_read_from = "INSA" then '----------------------------------------------------------------------------------------------------INSA
    call navigate_to_screen("stat", "insa")
    EMReadScreen INSA_amt, 1, 2, 78
    If INSA_amt <> 0 then
      EMReadScreen INSA_name, 38, 10, 38
      INSA_name = replace(INSA_name, "_", "")
      INSA_name = split(INSA_name)
      For each word in INSA_name
        first_letter_of_word = ucase(left(word, 1))
        rest_of_word = LCase(right(word, len(word) -1))
        If len(word) > 4 then
          variable_written_to = variable_written_to & first_letter_of_word & rest_of_word & " "
        Else
          variable_written_to = variable_written_to & word & " "
        End if
      Next
      variable_written_to = trim(variable_written_to) & "; "
    End if
  Elseif panel_read_from = "JOBS" then '----------------------------------------------------------------------------------------------------JOBS
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "jobs")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen JOBS_total, 1, 2, 78
      If JOBS_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_JOBS_to_variable(variable_written_to)
          EMReadScreen JOBS_panel_current, 1, 2, 73
          If cint(JOBS_panel_current) < cint(JOBS_total) then transmit
        Loop until cint(JOBS_panel_current) = cint(JOBS_total)
      End if
    Next
  Elseif panel_read_from = "MEDI" then '----------------------------------------------------------------------------------------------------MEDI
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "MEDI")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen MEDI_amt, 1, 2, 78
      If MEDI_amt <> "0" then variable_written_to = variable_written_to & "Medicare for member " & HH_member & ".; "
    Next
  Elseif panel_read_from = "MEMB" then '----------------------------------------------------------------------------------------------------MEMB
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "memb")
      EMWriteScreen HH_member, 20, 76
      transmit
      EMReadScreen rel_to_applicant, 2, 10, 42
      EMReadScreen client_age, 3, 8, 76
      If client_age = "   " then client_age = 0
      If cint(client_age) >= 21 or rel_to_applicant = "02" then
        number_of_adults = number_of_adults + 1
      Else
        number_of_children = number_of_children + 1
      End if
    Next
    If number_of_adults > 0 then variable_written_to = number_of_adults & "a"
    If number_of_children > 0 then variable_written_to = variable_written_to & ", " & number_of_children & "c"
    If left(variable_written_to, 1) = "," then variable_written_to = right(variable_written_to, len(variable_written_to) - 1)
  Elseif panel_read_from = "MEMI" then '----------------------------------------------------------------------------------------------------MEMI
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "memi")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen citizen, 1, 10, 49
      If citizen = "Y" then citizen = "US citizen"
      If citizen = "N" then citizen = "non-citizen"
      EMReadScreen citizenship_ver, 2, 10, 78
      EMReadScreen SSA_MA_citizenship_ver, 1, 11, 49
      If citizenship_ver = "__" or citizenship_ver = "NO" then cit_proof_indicator = ", no verifs provided"
      If SSA_MA_citizenship_ver = "R" then cit_proof_indicator = ", MEMI infc req'd"
      If (citizenship_ver <> "__" and citizenship_ver <> "NO") or (SSA_MA_citizenship_ver = "A") then cit_proof_indicator = ""
      If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
      variable_written_to = variable_written_to & citizen & cit_proof_indicator & "; "
    Next
  ElseIf panel_read_from = "MONT" then '----------------------------------------------------------------------------------------------------MONT
    call navigate_to_screen("stat", "mont")
    EMReadScreen variable_written_to, 8, 6, 39
    variable_written_to = replace(variable_written_to, " ", "/")
    If isdate(variable_written_to) = True then
      variable_written_to = cdate(variable_written_to) & ""
    Else
      variable_written_to = ""
    End if
  Elseif panel_read_from = "OTHR" then '----------------------------------------------------------------------------------------------------OTHR
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "othr")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen OTHR_total, 1, 2, 78
      If OTHR_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_OTHR_to_variable(variable_written_to)
          EMReadScreen OTHR_panel_current, 1, 2, 73
          If cint(OTHR_panel_current) < cint(OTHR_total) then transmit
        Loop until cint(OTHR_panel_current) = cint(OTHR_total)
      End if
    Next
  Elseif panel_read_from = "PBEN" then '----------------------------------------------------------------------------------------------------PBEN
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "pben")
      EMWriteScreen HH_member, 20, 76
      transmit
      EMReadScreen panel_amt, 1, 2, 78
      If panel_amt <> "0" then
        If HH_member <> "01" then PBEN = PBEN & "Member " & HH_member & "- "
        row = 8
        Do
          EMReadScreen PBEN_type, 12, row, 28
          EMReadScreen PBEN_disp, 1, row, 77
          If PBEN_disp = "A" then PBEN_disp = " appealing"
          If PBEN_disp = "D" then PBEN_disp = " denied"
          If PBEN_disp = "E" then PBEN_disp = " eligible"
          If PBEN_disp = "P" then PBEN_disp = " pends"
          If PBEN_disp = "N" then PBEN_disp = " not applied yet"
          If PBEN_disp = "R" then PBEN_disp = " refused"
          If PBEN_type <> "            " then PBEN = PBEN & trim(PBEN_type) & PBEN_disp & "; "
          row = row + 1
        Loop until row = 14
      End if
    Next
    If PBEN <> "" then variable_written_to = variable_written_to & PBEN
  Elseif panel_read_from = "PREG" then '----------------------------------------------------------------------------------------------------PREG
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "PREG")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen PREG_total, 1, 2, 78
      If PREG_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        EMReadScreen PREG_due_date, 8, 10, 53
        If PREG_due_date = "__ __ __" then
          PREG_due_date = "unknown"
        Else
          PREG_due_date = replace(PREG_due_date, " ", "/")
        End if
        variable_written_to = variable_written_to & "Due date is " & PREG_due_date & ".; "
      End if
    Next
  Elseif panel_read_from = "PROG" then '----------------------------------------------------------------------------------------------------PROG
    call navigate_to_screen("stat", "prog") 'THIS WILL DETERMINE THE LAST DATESTAMP ON THE PROG PANEL
    row = 6
    Do
      EMReadScreen appl_prog_date, 8, row, 33
      If appl_prog_date <> "__ __ __" then appl_prog_date_array = appl_prog_date_array & replace(appl_prog_date, " ", "/") & " "
      row = row + 1
    Loop until row = 13
    appl_prog_date_array = split(appl_prog_date_array)
    variable_written_to = CDate(appl_prog_date_array(0))
    for i = 0 to ubound(appl_prog_date_array) - 1
      if CDate(appl_prog_date_array(i)) > variable_written_to then 
        variable_written_to = CDate(appl_prog_date_array(i))
      End if
    next
    If isdate(variable_written_to) = True then
      variable_written_to = cdate(variable_written_to) & ""
    Else
      variable_written_to = ""
    End if
  Elseif panel_read_from = "RBIC" then '----------------------------------------------------------------------------------------------------RBIC
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "rbic")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen RBIC_total, 1, 2, 78
      If RBIC_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_RBIC_to_variable(variable_written_to)
          EMReadScreen RBIC_panel_current, 1, 2, 73
          If cint(RBIC_panel_current) < cint(RBIC_total) then transmit
        Loop until cint(RBIC_panel_current) = cint(RBIC_total)
      End if
    Next
  Elseif panel_read_from = "REST" then '----------------------------------------------------------------------------------------------------REST
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "rest")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen REST_total, 1, 2, 78
      If REST_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_REST_to_variable(variable_written_to)
          EMReadScreen REST_panel_current, 1, 2, 73
          If cint(REST_panel_current) < cint(REST_total) then transmit
        Loop until cint(REST_panel_current) = cint(REST_total)
      End if
    Next
  Elseif panel_read_from = "REVW" then '----------------------------------------------------------------------------------------------------REVW
    call navigate_to_screen("stat", "revw")
    EMReadScreen variable_written_to, 8, 13, 37
    variable_written_to = replace(variable_written_to, " ", "/")
    If isdate(variable_written_to) = True then
      variable_written_to = cdate(variable_written_to) & ""
    Else
      variable_written_to = ""
    End if
  Elseif panel_read_from = "SCHL" then '----------------------------------------------------------------------------------------------------SCHL
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "schl")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen school_type, 2, 7, 40
      If school_type = "01" then school_type = "elementary school"
      If school_type = "11" then school_type = "middle school"
      If school_type = "02" then school_type = "high school"
      If school_type = "03" then school_type = "GED"
      If school_type = "07" then school_type = "IEP"
      If school_type = "08" or school_type = "09" or school_type = "10" then school_type = "post-secondary"
      If school_type = "06" or school_type = "__" or school_type = "?_" then
        school_type = ""
      Else
        EMReadScreen SCHL_ver, 2, 6, 63
        If SCHL_ver = "?_" or SCHL_ver = "NO" then
          school_proof_type = ", no proof provided"
        Else
          school_proof_type = ""
        End if
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & school_type & school_proof_type & "; "
      End if
    Next
  Elseif panel_read_from = "SECU" then '----------------------------------------------------------------------------------------------------SECU
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "secu")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen SECU_total, 1, 2, 78
      If SECU_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_SECU_to_variable(variable_written_to)
          EMReadScreen SECU_panel_current, 1, 2, 73
          If cint(SECU_panel_current) < cint(SECU_total) then transmit
        Loop until cint(SECU_panel_current) = cint(SECU_total)
      End if
    Next
  Elseif panel_read_from = "SHEL" then '----------------------------------------------------------------------------------------------------SHEL
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "shel")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen SHEL_total, 1, 2, 78
      If SHEL_total <> 0 then 
        If HH_member <> "01" then member_number_designation = "Member " & HH_member & "- "
        row = 11
        Do
          EMReadScreen SHEL_amount, 8, row, 56
          If SHEL_amount <> "________" then
            EMReadScreen SHEL_type, 9, row, 24
            EMReadScreen SHEL_proof_check, 2, row, 67
            If SHEL_proof_check = "NO" or SHEL_proof_check = "?_" then 
              SHEL_proof = ", no proof provided"
            Else
              SHEL_proof = ""
            End if
            SHEL_expense = SHEL_expense & "$" & trim(SHEL_amount) & "/mo " & lcase(trim(SHEL_type)) & SHEL_proof & ". ;"
          End if
          row = row + 1
        Loop until row = 19
        variable_written_to = variable_written_to & member_number_designation & SHEL_expense
      End if
      SHEL_expense = ""
    Next
  Elseif panel_read_from = "STWK" then '----------------------------------------------------------------------------------------------------STWK
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "STWK")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen STWK_total, 1, 2, 78
      If STWK_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        EMReadScreen STWK_verification, 1, 7, 63
        If STWK_verification = "N" then
          STWK_verification = ", no proof provided"
        Else
          STWK_verification = ""
        End if
        EMReadScreen STWK_employer, 30, 6, 46
        STWK_employer = replace(STWK_employer, "_", "")
        STWK_employer = split(STWK_employer)
        For each a in STWK_employer
          If a <> "" then
            b = ucase(left(a, 1))
            c = LCase(right(a, len(a) -1))
            If len(a) > 3 then
              new_STWK_employer = new_STWK_employer & b & c & " "
            Else
              new_STWK_employer = new_STWK_employer & a & " "
            End if
          End if
        Next
        EMReadScreen STWK_income_stop_date, 8, 8, 46
        If STWK_income_stop_date = "__ __ __" then
          STWK_income_stop_date = "at unknown date"
        Else
          STWK_income_stop_date = replace(STWK_income_stop_date, " ", "/")
        End if
      EMReadScreen voluntary_quit, 1, 10, 46
	vol_quit_info = ", Vol. Quit " & voluntary_quit
	  IF voluntary_quit = "Y" THEN
		EMReadScreen good_cause, 1, 12, 67
		EMReadScreen fs_pwe, 1, 14, 46
		vol_quit_info = ", Vol Quit " & voluntary_quit & ", Good Cause " & good_cause & ", FS PWE " & fs_pwe
	  END IF
        variable_written_to = variable_written_to & new_STWK_employer & "income stopped " & STWK_income_stop_date & STWK_verification & vol_quit_info & ".; "
      End if
      new_STWK_employer = "" 'clearing variable to prevent duplicates
    Next
  Elseif panel_read_from = "UNEA" then '----------------------------------------------------------------------------------------------------UNEA
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "unea")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen UNEA_total, 1, 2, 78
      If UNEA_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_UNEA_to_variable(variable_written_to)
          EMReadScreen UNEA_panel_current, 1, 2, 73
          If cint(UNEA_panel_current) < cint(UNEA_total) then transmit
        Loop until cint(UNEA_panel_current) = cint(UNEA_total)
      End if
    Next
  Elseif panel_read_from = "WREG" then '---------------------------------------------------------------------------------------------------WREG
    For each HH_member in HH_member_array
	call navigate_to_screen("stat", "wreg")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
    EMReadScreen wreg_total, 1, 2, 78
    EMReadScreen snap_case_yn, 1, 6, 50
    IF wreg_total <> "0" and snap_case_yn = "Y" THEN 
	EmWriteScreen "x", 13, 57
	transmit
	 bene_mo_col = (15 + (4*cint(footer_month)))
	  bene_yr_row = 10
       abawd_counted_months = 0
       second_abawd_period = 0
 	 month_count = 0
 	   DO
  		  EMReadScreen is_counted_month, 1, bene_yr_row, bene_mo_col
  		    IF is_counted_month = "X" or is_counted_month = "M" THEN abawd_counted_months = abawd_counted_months + 1
		    IF is_counted_month = "Y" or is_counted_month = "N" THEN second_abawd_period = second_abawd_period + 1
   		  bene_mo_col = bene_mo_col - 4
    		    IF bene_mo_col = 15 THEN
        		bene_yr_row = bene_yr_row - 1
   	     		bene_mo_col = 63
   	   	    END IF
    		  month_count = month_count + 1
  	   LOOP until month_count = 36
  	PF3
	EmreadScreen read_abawd_status, 2, 13, 50
	If read_abawd_status = 10 or read_abawd_status = 11 or read_abawd_status = 13 then
	  abawd_status = "Client is ABAWD and has used " & abawd_counted_months & " months"
	else
	  abawd_status = "Client is not ABAWD"
	end if
      IF second_abawd_period <> 0 THEN abawd_status = "CL is ABAWD and has used second 3-month period"
	variable_written_to = variable_written_to & "Member " & HH_member & "- " & abawd_status & ".; "
     END IF
    Next
  End if
  variable_written_to = trim(variable_written_to) '-----------------------------------------------------------------------------------------cleaning up editbox
  if right(variable_written_to, 1) = ";" then variable_written_to= left(variable_written_to, len(variable_written_to) - 1)
  variable_written_to = replace(variable_written_to, "$________/non-monthly", "amt unknown")
  variable_written_to = replace(variable_written_to, "$________/monthly", "amt unknown")
  variable_written_to = replace(variable_written_to, "$________/weekly", "amt unknown")
  variable_written_to = replace(variable_written_to, "$________/biweekly", "amt unknown")
  variable_written_to = replace(variable_written_to, "$________/semimonthly", "amt unknown")
End function


function back_to_SELF
  Do
    EMSendKey "<PF3>"
    EMWaitReady 0, 0
    EMReadScreen SELF_check, 4, 2, 50
  Loop until SELF_check = "SELF"
End function

Function create_array_of_all_active_x_numbers_in_county(array_name, county_code)
	'Getting to REPT/USER
	call navigate_to_screen("rept", "user")

	'Hitting PF5 to force sorting, which allows directly selecting a county
	PF5

	'Inserting county
	EMWriteScreen county_code, 21, 6
	transmit

	'Declaring the MAXIS row
	MAXIS_row = 7

	'Blanking out array_name in case this has been used already in the script
	array_name = ""

	Do
		Do
			'Reading MAXIS information for this row, adding to spreadsheet
			EMReadScreen worker_ID, 8, MAXIS_row, 5					'worker ID
			If worker_ID = "        " then exit do					'exiting before writing to array, in the event this is a blank (end of list)
			array_name = trim(array_name & " " & worker_ID)				'writing to variable
			MAXIS_row = MAXIS_row + 1
		Loop until MAXIS_row = 19

		'Seeing if there are more pages. If so it'll grab from the next page and loop around, doing so until there's no more pages.
		EMReadScreen more_pages_check, 7, 19, 3
		If more_pages_check = "More: +" then 
			PF8			'getting to next screen
			MAXIS_row = 7	'redeclaring MAXIS row so as to start reading from the top of the list again
		End if
	Loop until more_pages_check = "More:  " or more_pages_check = "       "	'The or works because for one-page only counties, this will be blank
	array_name = split(array_name)
End function

Function create_MAXIS_friendly_date(date_variable, variable_length, screen_row, screen_col) 
  var_month = datepart("m", dateadd("d", variable_length, date_variable))
  If len(var_month) = 1 then var_month = "0" & var_month
  EMWriteScreen var_month, screen_row, screen_col
  var_day = datepart("d", dateadd("d", variable_length, date_variable))
  If len(var_day) = 1 then var_day = "0" & var_day
  EMWriteScreen var_day, screen_row, screen_col + 3
  var_year = datepart("yyyy", dateadd("d", variable_length, date_variable))
  EMWriteScreen right(var_year, 2), screen_row, screen_col + 6
End function

'This function fixes the case for a phrase. For example, "ROBERT P. ROBERTSON" becomes "Robert P. Robertson". 
'	It capitalizes the first letter of each word.
Function fix_case(phrase_to_split, smallest_length_to_skip)									'Ex: fix_case(client_name, 3), where 3 means skip words that are 3 characters or shorter
	phrase_to_split = split(phrase_to_split)											'splits phrase into an array
	For each word in phrase_to_split												'processes each word independently
		If word <> "" then													'Skip blanks
			first_character = ucase(left(word, 1))									'grabbing the first character of the string, making uppercase and adding to variable
			remaining_characters = LCase(right(word, len(word) -1))						'grabbing the remaining characters of the string, making lowercase and adding to variable
			If len(word) > smallest_length_to_skip then								'skip any strings shorter than the smallest_length_to_skip variable
				output_phrase = output_phrase & first_character & remaining_characters & " "		'output_phrase is the output of the function, this combines the first_character and remaining_characters
			Else															
				output_phrase = output_phrase & word & " "							'just pops the whole word in if it's shorter than the smallest_length_to_skip variable
			End if
		End if
	Next
	phrase_to_split = output_phrase												'making the phrase_to_split equal to the output, so that it can be used by the rest of the script.
End function

Function ERRR_screen_check 'Checks for error prone cases
	EMReadScreen ERRR_check, 4, 2, 52
	If ERRR_check = "ERRR" then transmit
End Function

Function end_excel_and_script
  objExcel.Workbooks.Close
  objExcel.quit
  stopscript
End function

Function find_variable(x, y, z) 'x is string, y is variable, z is length of new variable
  row = 1
  col = 1
  EMSearch x, row, col
  If row <> 0 then EMReadScreen y, z, row, col + len(x)
End function

Function get_to_MMIS_session_begin
  Do 
    EMSendkey "<PF6>"
    EMWaitReady 0, 0
    EMReadScreen session_start, 18, 1, 7
  Loop until session_start = "SESSION TERMINATED"
End function

Function MAXIS_background_check
	Do
		EMReadScreen SELF_check, 4, 2, 50
		If SELF_check = "SELF" then
			PF3
			Pause 2
		End if
	Loop until SELF_check <> "SELF"
End function

Function MAXIS_case_number_finder(variable_for_MAXIS_case_number)
	row = 1
	col = 1
	EMSearch "Case Nbr:", row, col
	If row <> 0 then 
		EMReadScreen variable_for_MAXIS_case_number, 8, row, col + 10
		variable_for_MAXIS_case_number = replace(variable_for_MAXIS_case_number, "_", "")
		variable_for_MAXIS_case_number = trim(variable_for_MAXIS_case_number)
	End if
End function

'This function confirms that we're in MAXIS. If we aren't, it will stop.
Function maxis_check_function
	transmit
	EMReadScreen MAXIS_check, 5, 1, 39
	If MAXIS_check <> "MAXIS"  and MAXIS_check <> "AXIS " then script_end_procedure("You do not appear to be in MAXIS. You may be passworded out. Please check your MAXIS screen and try again.")
End function

Function HH_member_custom_dialog(HH_member_array)
  'THE FOLLOWING DIALOG WILL DYNAMICALLY CHANGE DEPENDING ON THE HH COMP. IT WILL ALLOW A WORKER TO SELECT WHICH HH MEMBERS NEED TO BE INCLUDED IN THE SCRIPT.
  EMReadScreen HH_member_01, 18, 5, 3                                       'THIS GATHERS THE HH MEMBERS DIRECTLY FROM A MAXIS SCREEN.
  EMReadScreen HH_member_02, 18, 6, 3
  EMReadScreen HH_member_03, 18, 7, 3
  EMReadScreen HH_member_04, 18, 8, 3
  EMReadScreen HH_member_05, 18, 9, 3
  EMReadScreen HH_member_06, 18, 10, 3
  EMReadScreen HH_member_07, 18, 11, 3
  EMReadScreen HH_member_08, 18, 12, 3
  EMReadScreen HH_member_09, 18, 13, 3
  EMReadScreen HH_member_10, 18, 14, 3
  EMReadScreen HH_member_11, 18, 15, 3
  EMReadScreen HH_member_12, 18, 16, 3
  EMReadScreen HH_member_13, 18, 17, 3
  EMReadScreen HH_member_14, 18, 18, 3
  EMReadScreen HH_member_15, 18, 19, 3
  dialog_size_variable = 50                                                 'DEFAULT IS 50, BUT IT CHANGES DEPENDING ON THE AMOUNT OF HH MEMBERS.
  If HH_member_03 <> "                  " then dialog_size_variable = 65     
  If HH_member_04 <> "                  " then dialog_size_variable = 80
  If HH_member_05 <> "                  " then dialog_size_variable = 95
  If HH_member_06 <> "                  " then dialog_size_variable = 110
  If HH_member_07 <> "                  " then dialog_size_variable = 125
  If HH_member_08 <> "                  " then dialog_size_variable = 140
  If HH_member_09 <> "                  " then dialog_size_variable = 155
  If HH_member_10 <> "                  " then dialog_size_variable = 170
  If HH_member_11 <> "                  " then dialog_size_variable = 185
  If HH_member_12 <> "                  " then dialog_size_variable = 200
  If HH_member_13 <> "                  " then dialog_size_variable = 215
  If HH_member_14 <> "                  " then dialog_size_variable = 230
  If HH_member_15 <> "                  " then dialog_size_variable = 245
  If HH_member_01 <> "                  " then client_01_check = 1          'ALL CHECKBOXES DEFAULT TO CHECKED, AS USUALLY WE NEED ALL HH MEMBER INFO.
  If HH_member_02 <> "                  " then client_02_check = 1
  If HH_member_03 <> "                  " then client_03_check = 1
  If HH_member_04 <> "                  " then client_04_check = 1
  If HH_member_05 <> "                  " then client_05_check = 1
  If HH_member_06 <> "                  " then client_06_check = 1
  If HH_member_07 <> "                  " then client_07_check = 1
  If HH_member_08 <> "                  " then client_08_check = 1
  If HH_member_09 <> "                  " then client_09_check = 1
  If HH_member_10 <> "                  " then client_10_check = 1
  If HH_member_11 <> "                  " then client_11_check = 1
  If HH_member_12 <> "                  " then client_12_check = 1
  If HH_member_13 <> "                  " then client_13_check = 1
  If HH_member_14 <> "                  " then client_14_check = 1
  If HH_member_15 <> "                  " then client_15_check = 1
  BeginDialog HH_memb_dialog, 0, 0, 191, dialog_size_variable, "HH member dialog"
    ButtonGroup ButtonPressed
      OkButton 135, 10, 50, 15
      CancelButton 135, 30, 50, 15
    Text 10, 5, 105, 10, "Household members to look at:"
    If HH_member_01 <> "                  " then CheckBox 10, 20, 120, 10, HH_member_01, client_01_check
    If HH_member_02 <> "                  " then CheckBox 10, 35, 120, 10, HH_member_02, client_02_check
    If HH_member_03 <> "                  " then CheckBox 10, 50, 120, 10, HH_member_03, client_03_check
    If HH_member_04 <> "                  " then CheckBox 10, 65, 120, 10, HH_member_04, client_04_check
    If HH_member_05 <> "                  " then CheckBox 10, 80, 120, 10, HH_member_05, client_05_check
    If HH_member_06 <> "                  " then CheckBox 10, 95, 120, 10, HH_member_06, client_06_check
    If HH_member_07 <> "                  " then CheckBox 10, 110, 120, 10, HH_member_07, client_07_check
    If HH_member_08 <> "                  " then CheckBox 10, 125, 120, 10, HH_member_08, client_08_check
    If HH_member_09 <> "                  " then CheckBox 10, 140, 120, 10, HH_member_09, client_09_check
    If HH_member_10 <> "                  " then CheckBox 10, 155, 120, 10, HH_member_10, client_10_check
    If HH_member_11 <> "                  " then CheckBox 10, 170, 120, 10, HH_member_11, client_11_check
    If HH_member_12 <> "                  " then CheckBox 10, 185, 120, 10, HH_member_12, client_12_check
    If HH_member_13 <> "                  " then CheckBox 10, 200, 120, 10, HH_member_13, client_13_check
    If HH_member_14 <> "                  " then CheckBox 10, 215, 120, 10, HH_member_14, client_14_check
    If HH_member_15 <> "                  " then CheckBox 10, 230, 120, 10, HH_member_15, client_15_check
  EndDialog
  'NOW IT SHOWS THE DIALOG FROM THE LAST SCREEN
  Do
    Dialog HH_memb_dialog
    If buttonpressed = 0 then stopscript
    transmit
    EMReadScreen MAXIS_check, 5, 1, 39
    IF MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You do not appear to be in MAXIS. You may have navigated away, or are passworded out. Clear up the issue, and try again."
  Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " 
  'DETERMINING WHICH HH MEMBERS TO LOOK AT
  If client_01_check = 1 then HH_member_array = HH_member_array & left(HH_member_01, 2) & " "
  If client_02_check = 1 then HH_member_array = HH_member_array & left(HH_member_02, 2) & " "
  If client_03_check = 1 then HH_member_array = HH_member_array & left(HH_member_03, 2) & " "
  If client_04_check = 1 then HH_member_array = HH_member_array & left(HH_member_04, 2) & " "
  If client_05_check = 1 then HH_member_array = HH_member_array & left(HH_member_05, 2) & " "
  If client_06_check = 1 then HH_member_array = HH_member_array & left(HH_member_06, 2) & " "
  If client_07_check = 1 then HH_member_array = HH_member_array & left(HH_member_07, 2) & " "
  If client_08_check = 1 then HH_member_array = HH_member_array & left(HH_member_08, 2) & " "
  If client_09_check = 1 then HH_member_array = HH_member_array & left(HH_member_09, 2) & " "
  If client_10_check = 1 then HH_member_array = HH_member_array & left(HH_member_10, 2) & " "
  If client_11_check = 1 then HH_member_array = HH_member_array & left(HH_member_11, 2) & " "
  If client_12_check = 1 then HH_member_array = HH_member_array & left(HH_member_12, 2) & " "
  If client_13_check = 1 then HH_member_array = HH_member_array & left(HH_member_13, 2) & " "
  If client_14_check = 1 then HH_member_array = HH_member_array & left(HH_member_14, 2) & " "
  If client_15_check = 1 then HH_member_array = HH_member_array & left(HH_member_15, 2) & " "
  HH_member_array = trim(HH_member_array)
  HH_member_array = split(HH_member_array, " ")
End function

Function memb_navigation_next
  HH_memb_row = HH_memb_row + 1
  EMReadScreen next_HH_memb, 2, HH_memb_row, 3
  If isnumeric(next_HH_memb) = False then
    HH_memb_row = HH_memb_row + 1
  Else
    EMWriteScreen next_HH_memb, 20, 76
    EMWriteScreen "01", 20, 79
  End if
End function

Function memb_navigation_prev
  HH_memb_row = HH_memb_row - 1
  EMReadScreen prev_HH_memb, 2, HH_memb_row, 3
  If isnumeric(prev_HH_memb) = False then
    HH_memb_row = HH_memb_row + 1
  Else
    EMWriteScreen prev_HH_memb, 20, 76
    EMWriteScreen "01", 20, 79
  End if
End function

Function MMIS_RKEY_finder
  'Now we use a Do Loop to get to the start screen for MMIS.
  Do 
    EMSendkey "<PF6>"
    EMWaitReady 0, 0
    EMReadScreen session_start, 18, 1, 7
  Loop until session_start = "SESSION TERMINATED"
  'Now we get back into MMIS. We have to skip past the intro screens.
  EMWriteScreen "mw00", 1, 2
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  'This section may not work for all OSAs, since some only have EK01. This will find EK01 and enter it.
  MMIS_row = 1
  MMIS_col = 1
  EMSearch "EK01", MMIS_row, MMIS_col
  If MMIS_row <> 0 then
    EMWriteScreen "x", MMIS_row, 4
    EMSendKey "<enter>"
    EMWaitReady 0, 0
  End if
  'This section starts from EK01. OSAs may need to skip the previous section.
  EMWriteScreen "x", 10, 3
  EMSendKey "<enter>"
  EMWaitReady 0, 0
End function


function navigate_to_screen(x, y)
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " then
    row = 1
    col = 1
    EMSearch "Function: ", row, col
    If row <> 0 then 
      EMReadScreen MAXIS_function, 4, row, col + 10
      EMReadScreen STAT_note_check, 4, 2, 45
      row = 1
      col = 1
      EMSearch "Case Nbr: ", row, col
      EMReadScreen current_case_number, 8, row, col + 10
      current_case_number = replace(current_case_number, "_", "")
      current_case_number = trim(current_case_number)
    End if
    If current_case_number = case_number and MAXIS_function = ucase(x) and STAT_note_check <> "NOTE" then 
      row = 1
      col = 1
      EMSearch "Command: ", row, col
      EMWriteScreen y, row, col + 9
      EMSendKey "<enter>"
      EMWaitReady 0, 0
    Else
      Do
        EMSendKey "<PF3>"
        EMWaitReady 0, 0
        EMReadScreen SELF_check, 4, 2, 50
      Loop until SELF_check = "SELF"
      EMWriteScreen x, 16, 43
      EMWriteScreen "________", 18, 43
      EMWriteScreen case_number, 18, 43
      EMWriteScreen footer_month, 20, 43
      EMWriteScreen footer_year, 20, 46
      EMWriteScreen y, 21, 70
      EMSendKey "<enter>"
      EMWaitReady 0, 0
      EMReadScreen abended_check, 7, 9, 27
      If abended_check = "abended" then
        EMSendKey "<enter>"
        EMWaitReady 0, 0
      End if
    End if
  End if
End function

Function navigate_to_PRISM_screen(x) 'x is the name of the screen
  EMWriteScreen x, 21, 18
  EMSendKey "<enter>"
  EMWaitReady 0, 0
End function

function navigation_buttons 'this works by calling the navigation_buttons function when the buttonpressed isn't -1
  If ButtonPressed = ABPS_button then call navigate_to_screen("stat", "ABPS")
  If ButtonPressed = ACCI_button then call navigate_to_screen("stat", "ACCI")
  If ButtonPressed = ACCT_button then call navigate_to_screen("stat", "ACCT")
  If ButtonPressed = ADDR_button then call navigate_to_screen("stat", "ADDR")
  If ButtonPressed = ALTP_button then call navigate_to_screen("stat", "ALTP")
  If ButtonPressed = AREP_button then call navigate_to_screen("stat", "AREP")
  If ButtonPressed = BILS_button then call navigate_to_screen("stat", "BILS")
  If ButtonPressed = BUSI_button then call navigate_to_screen("stat", "BUSI")
  If ButtonPressed = CARS_button then call navigate_to_screen("stat", "CARS")
  If ButtonPressed = CASH_button then call navigate_to_screen("stat", "CASH")
  If ButtonPressed = COEX_button then call navigate_to_screen("stat", "COEX")
  If ButtonPressed = DCEX_button then call navigate_to_screen("stat", "DCEX")
  If ButtonPressed = DIET_button then call navigate_to_screen("stat", "DIET")
  If ButtonPressed = DISA_button then call navigate_to_screen("stat", "DISA")
  If ButtonPressed = EATS_button then call navigate_to_screen("stat", "EATS")
  If ButtonPressed = ELIG_DWP_button then call navigate_to_screen("elig", "DWP_")
  If ButtonPressed = ELIG_FS_button then call navigate_to_screen("elig", "FS__")
  If ButtonPressed = ELIG_GA_button then call navigate_to_screen("elig", "GA__")
  If ButtonPressed = ELIG_HC_button then call navigate_to_screen("elig", "HC__")
  If ButtonPressed = ELIG_MFIP_button then call navigate_to_screen("elig", "MFIP")
  If ButtonPressed = ELIG_MSA_button then call navigate_to_screen("elig", "MSA_")
  If ButtonPressed = ELIG_WB_button then call navigate_to_screen("elig", "WB__")
  If ButtonPressed = FACI_button then call navigate_to_screen("stat", "FACI")
  If ButtonPressed = FMED_button then call navigate_to_screen("stat", "FMED")
  If ButtonPressed = HCRE_button then call navigate_to_screen("stat", "HCRE")
  If ButtonPressed = HEST_button then call navigate_to_screen("stat", "HEST")
  If ButtonPressed = IMIG_button then call navigate_to_screen("stat", "IMIG")
  If ButtonPressed = INSA_button then call navigate_to_screen("stat", "INSA")
  If ButtonPressed = JOBS_button then call navigate_to_screen("stat", "JOBS")
  If ButtonPressed = MEDI_button then call navigate_to_screen("stat", "MEDI")
  If ButtonPressed = MEMB_button then call navigate_to_screen("stat", "MEMB")
  If ButtonPressed = MEMI_button then call navigate_to_screen("stat", "MEMI")
  If ButtonPressed = MONT_button then call navigate_to_screen("stat", "MONT")
  If ButtonPressed = OTHR_button then call navigate_to_screen("stat", "OTHR")
  If ButtonPressed = PBEN_button then call navigate_to_screen("stat", "PBEN")
  If ButtonPressed = PDED_button then call navigate_to_screen("stat", "PDED")
  If ButtonPressed = PREG_button then call navigate_to_screen("stat", "PREG")
  If ButtonPressed = PROG_button then call navigate_to_screen("stat", "PROG")
  If ButtonPressed = RBIC_button then call navigate_to_screen("stat", "RBIC")
  If ButtonPressed = REST_button then call navigate_to_screen("stat", "REST")
  If ButtonPressed = REVW_button then call navigate_to_screen("stat", "REVW")
  If ButtonPressed = SCHL_button then call navigate_to_screen("stat", "SCHL")
  If ButtonPressed = SECU_button then call navigate_to_screen("stat", "SECU")
  If ButtonPressed = STIN_button then call navigate_to_screen("stat", "STIN")
  If ButtonPressed = STEC_button then call navigate_to_screen("stat", "STEC")
  If ButtonPressed = STWK_button then call navigate_to_screen("stat", "STWK")
  If ButtonPressed = SHEL_button then call navigate_to_screen("stat", "SHEL")
  If ButtonPressed = SWKR_button then call navigate_to_screen("stat", "SWKR")
  If ButtonPressed = TRAN_button then call navigate_to_screen("stat", "TRAN")
  If ButtonPressed = TYPE_button then call navigate_to_screen("stat", "TYPE")
  If ButtonPressed = UNEA_button then call navigate_to_screen("stat", "UNEA")
End function

function new_BS_BSI_heading
  EMGetCursor MAXIS_row, MAXIS_col
  If MAXIS_row = 4 then 
    EMSendKey "--------BURIAL SPACE/ITEMS---------------AMOUNT----------STATUS--------------" & "<newline>"
    MAXIS_row = 5
  end if
End function

function new_CAI_heading
  EMGetCursor MAXIS_row, MAXIS_col
  If MAXIS_row = 4 then 
    EMSendKey "--------CASH ADVANCE ITEMS---------------AMOUNT----------STATUS--------------" & "<newline>"
    MAXIS_row = 5
  end if
End function

function new_page_check
  EMGetCursor MAXIS_row, MAXIS_col
  If MAXIS_row = 17 then
    EMSendKey ">>>>MORE>>>>"
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
    MAXIS_row = 4
  End if
end function

function new_service_heading
  EMGetCursor MAXIS_service_row, MAXIS_service_col
  If MAXIS_service_row = 4 then 
    EMSendKey "--------------SERVICE--------------------AMOUNT----------STATUS--------------" & "<newline>"
    MAXIS_service_row = 5
  end if
End function

Function panel_navigation_next
  EMReadScreen current_panel, 1, 2, 73
  EMReadScreen amount_of_panels, 1, 2, 78
  If current_panel < amount_of_panels then new_panel = current_panel + 1
  If current_panel = amount_of_panels then new_panel = current_panel
  If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
End function

Function panel_navigation_prev
  EMReadScreen current_panel, 1, 2, 73
  EMReadScreen amount_of_panels, 1, 2, 78
  If current_panel = 1 then new_panel = current_panel
  If current_panel > 1 then new_panel = current_panel - 1
  If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
End function

Function PF1
  EMSendKey "<PF1>"
  EMWaitReady 0, 0
End function

Function PF2
  EMSendKey "<PF2>"
  EMWaitReady 0, 0
End function

function PF3
  EMSendKey "<PF3>"
  EMWaitReady 0, 0
end function

Function PF4
  EMSendKey "<PF4>"
  EMWaitReady 0, 0
End function

Function PF5
  EMSendKey "<PF5>"
  EMWaitReady 0, 0
End function

Function PF6
  EMSendKey "<PF6>"
  EMWaitReady 0, 0
End function

Function PF7
  EMSendKey "<PF7>"
  EMWaitReady 0, 0
End function

function PF8
  EMSendKey "<PF8>"
  EMWaitReady 0, 0
end function

function PF9
  EMSendKey "<PF9>"
  EMWaitReady 0, 0
end function

function PF10
  EMSendKey "<PF10>"
  EMWaitReady 0, 0
end function

Function PF11
  EMSendKey "<PF11>"
  EMWaitReady 0, 0
End function

Function PF12
  EMSendKey "<PF12>"
  EMWaitReady 0, 0
End function

Function PF19
  EMSendKey "<PF19>"
  EMWaitReady 0, 0
End function

function PF19
  EMSendKey "<PF19>"
  EMWaitReady 0, 0
end function

function PF20
  EMSendKey "<PF20>"
  EMWaitReady 0, 0
end function

function run_another_script(script_path)
  Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
  Set fso_command = run_another_script_fso.OpenTextFile(script_path)
  text_from_the_other_script = fso_command.ReadAll
  fso_command.Close
  Execute text_from_the_other_script
end function

function script_end_procedure(closing_message)
	If closing_message <> "" then MsgBox closing_message
	stop_time = timer
	script_run_time = stop_time - start_time
	If is_county_collecting_stats  = True then
		'Getting user name
		Set objNet = CreateObject("WScript.NetWork") 
		user_ID = objNet.UserName

		'Setting constants
		Const adOpenStatic = 3
		Const adLockOptimistic = 3

		'Creating objects for Access
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'Fixing a bug when the script_end_procedure has an apostrophe (this interferes with Access)
		closing_message = replace(closing_message, "'", "")

		'Opening DB
		objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = H:\VKC dev directory\Statistics\usage statistics.accdb"

		'Opening usage_log and adding a record
		objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX)" &  _
		"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & script_run_time & ", '" & closing_message & "')", objConnection, adOpenStatic, adLockOptimistic	
	End if
	stopscript
end function

function script_end_procedure_wsh(closing_message) 'For use when running a script outside of the BlueZone Script Host
	If closing_message <> "" then MsgBox closing_message
	stop_time = timer
	script_run_time = stop_time - start_time
	If is_county_collecting_stats = True then
		'Getting user name
		Set objNet = CreateObject("WScript.NetWork") 
		user_ID = objNet.UserName

		'Setting constants
		Const adOpenStatic = 3
		Const adLockOptimistic = 3

		'Creating objects for Access
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'Opening DB
		objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = H:\VKC dev directory\Statistics\usage statistics.accdb"

		'Opening usage_log and adding a record
		objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX)" &  _
		"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & script_run_time & ", '" & closing_message & "')", objConnection, adOpenStatic, adLockOptimistic
	End if
	Wscript.Quit
end function

function stat_navigation
  EMReadScreen STAT_check, 4, 20, 21
  If STAT_check = "STAT" then
    If ButtonPressed = prev_panel_button then 
      EMReadScreen current_panel, 1, 2, 73
      EMReadScreen amount_of_panels, 1, 2, 78
      If current_panel = 1 then new_panel = current_panel
      If current_panel > 1 then new_panel = current_panel - 1
      If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
    End if
    If ButtonPressed = next_panel_button then 
      EMReadScreen current_panel, 1, 2, 73
      EMReadScreen amount_of_panels, 1, 2, 78
      If current_panel < amount_of_panels then new_panel = current_panel + 1
      If current_panel = amount_of_panels then new_panel = current_panel
      If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
    End if
    If ButtonPressed = prev_memb_button then 
      HH_memb_row = HH_memb_row - 1
      EMReadScreen prev_HH_memb, 2, HH_memb_row, 3
      If isnumeric(prev_HH_memb) = False then
        HH_memb_row = HH_memb_row + 1
      Else
        EMWriteScreen prev_HH_memb, 20, 76
        EMWriteScreen "01", 20, 79
      End if
    End if
    If ButtonPressed = next_memb_button then 
      HH_memb_row = HH_memb_row + 1
      EMReadScreen next_HH_memb, 2, HH_memb_row, 3
      If isnumeric(next_HH_memb) = False then
        HH_memb_row = HH_memb_row + 1
      Else
        EMWriteScreen next_HH_memb, 20, 76
        EMWriteScreen "01", 20, 79
      End if
    End if
  End if
End function

Function step_through_handling 'This function will introduce "warning screens" before each transmit, which is very helpful for testing new scripts
	'To use this function, simply replace the "Execute text_from_the_other_script" line with:
	'Execute replace(text_from_the_other_script, "EMWaitReady 0, 0", "step_through_handling")
	step_through = MsgBox("Step " & step_number & chr(13) & chr(13) & "If you see something weird on your screen (like a MAXIS or PRISM error), PRESS CANCEL then email your script administrator about it. Make sure you include the step you're on.", 1)
	If step_number = "" then step_number = 1	'Declaring the variable
	If step_through = 2 then
		stopscript
	Else
		EMWaitReady 0, 0
		step_number = step_number + 1
	End if
End Function

function transmit
  EMSendKey "<enter>"
  EMWaitReady 0, 0
end function

Function write_editbox_in_case_note(x, y, z) 'x is the header, y is the variable for the edit box which will be put in the case note, z is the length of spaces for the indent.
  variable_array = split(y, " ")
  EMSendKey "* " & x & ": "
  For each x in variable_array 
    EMGetCursor row, col 
    If (row = 17 and col + (len(x)) >= 80) or (row = 4 and col = 3) then
      EMSendKey "<PF8>"
      EMWaitReady 0, 0
    End if
    EMReadScreen max_check, 51, 24, 2
    If max_check = "A MAXIMUM OF 4 PAGES ARE ALLOWED FOR EACH CASE NOTE" then exit for
    EMGetCursor row, col 
    If (row < 17 and col + (len(x)) >= 80) then EMSendKey "<newline>" & space(z)
    If (row = 4 and col = 3) then EMSendKey space(z)
    EMSendKey x & " "
    If right(x, 1) = ";" then 
      EMSendKey "<backspace>" & "<backspace>" 
      EMGetCursor row, col 
      If row = 17 then
        EMSendKey "<PF8>"
        EMWaitReady 0, 0
        EMSendKey space(z)
      Else
        EMSendKey "<newline>" & space(z)
      End if
    End if
  Next
  EMSendKey "<newline>"
  EMGetCursor row, col 
  If (row = 17 and col + (len(x)) >= 80) or (row = 4 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
End function

Function write_new_line_in_case_note(x)
  EMGetCursor row, col 
  If (row = 17 and col + (len(x)) >= 80 + 1 ) or (row = 4 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
  EMReadScreen max_check, 51, 24, 2
  EMSendKey x & "<newline>"
  EMGetCursor row, col 
  If (row = 17 and col + (len(x)) >= 80) or (row = 4 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
End function

Function write_three_columns_in_case_note(col_01_start_point, col_01_variable, col_02_start_point, col_02_variable, col_03_start_point, col_03_variable)
  EMGetCursor row, col 
  If (row = 17 and col + (len(x)) >= 80 + 1 ) or (row = 4 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
  EMReadScreen max_check, 51, 24, 2
  EMGetCursor row, col
  EMWriteScreen "                                                                              ", row, 3
  EMSetCursor row, col_01_start_point
  EMSendKey col_01_variable
  EMSetCursor row, col_02_start_point
  EMSendKey col_02_variable
  EMSetCursor row, col_03_start_point
  EMSendKey col_03_variable
  EMSendKey "<newline>"
  EMGetCursor row, col 
  If (row = 17 and col + (len(x)) >= 80) or (row = 4 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
End function



