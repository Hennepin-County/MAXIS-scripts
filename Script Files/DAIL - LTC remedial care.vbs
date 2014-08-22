'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "DAIL - LTC remedial care"
start_time = timer

'LOADING ROUTINE FUNCTIONS
'<<<GO THROUGH AND REMOVE REDUNDANT FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script


EMConnect ""

BeginDialog Dialog1, 0, 0, 191, 86, "Dialog"
  ButtonGroup ButtonPressed
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
  Text 10, 5, 115, 50, "This script will update your STAT/BILS panel's remedial care (27) entries, to the current deduction rate of $260. The script will only update the entries dated 07/01/2012 or later."
  Text 10, 65, 170, 20, "Press OK to start. Remember to case note when you are finished!"
EndDialog

Dialog dialog1
If ButtonPressed = 0 then stopscript

EMSendKey "s" & "<enter>"
EMWaitReady 0, 0

EMWriteScreen "bils", 20, 71
EMSendKey "<enter>"
EMWaitReady 0, 0

EMSendKey "<PF9>"
EMWaitReady 0, 0

Do
  EMReadScreen page_number, 2, 3, 72
  If page_number = " 1" then exit do
  EMSendKey "<PF19>" 'This is shift-PF7
  EMWaitReady 0, 0
Loop until page_number = " 1"

target_date = "06/30/2012" 'This sets the date range that should be changed, and will need to be updated in code at each COLA.
updates_made = 0 'Setting the variable for the following do...loop

Do

  EMReadScreen BILS_line_01, 54, 6, 26
  BILS_line_01 = replace(BILS_line_01, "$", " ")
  BILS_line_01 = split(BILS_line_01, "  ")
  BILS_line_01(1) = replace(BILS_line_01(1), " ", "/")
  If IsDate(BILS_line_01(1)) = True then 
    If datediff("d", target_date, BILS_line_01(1)) > 0 and BILS_line_01(2) = 27 and BILS_line_01(5) <> "260.00" then 
      EMWriteScreen "260.00", 6, 48
      EMWriteScreen "c", 6, 24
      updates_made = updates_made + 1
    End If
  End If

  EMReadScreen BILS_line_02, 54, 7, 26
  BILS_line_02 = replace(BILS_line_02, "$", " ")
  BILS_line_02 = split(BILS_line_02, "  ")
  BILS_line_02(1) = replace(BILS_line_02(1), " ", "/")
  If IsDate(BILS_line_02(1)) = True then 
    If datediff("d", target_date, BILS_line_02(1)) > 0 and BILS_line_02(2) = 27 and BILS_line_02(5) <> "260.00" then  
    EMWriteScreen "260.00", 7, 48
    EMWriteScreen "c", 7, 24
    updates_made = updates_made + 1
    End If
  End If

  EMReadScreen   BILS_line_03, 54, 8, 26
  BILS_line_03 = replace(BILS_line_03, "$", " ")
  BILS_line_03 = split(BILS_line_03, "  ")
  BILS_line_03(1) = replace(BILS_line_03(1), " ", "/")
  If IsDate(BILS_line_03(1)) = True then 
    If datediff("d", target_date, BILS_line_03(1)) > 0 and BILS_line_03(2) = 27 and BILS_line_03(5) <> "260.00" then  
    EMWriteScreen "260.00", 8, 48
    EMWriteScreen "c", 8, 24
    updates_made = updates_made + 1
    End If
  End If

  EMReadScreen   BILS_line_04, 54, 9, 26
  BILS_line_04 = replace(BILS_line_04, "$", " ")
  BILS_line_04 = split(BILS_line_04, "  ")
  BILS_line_04(1) = replace(BILS_line_04(1), " ", "/")
  If IsDate(BILS_line_04(1)) = True then 
    If datediff("d", target_date, BILS_line_04(1)) > 0 and BILS_line_04(2) = 27 and BILS_line_04(5) <> "260.00" then  
    EMWriteScreen "260.00", 9, 48
    EMWriteScreen "c", 9, 24
    updates_made = updates_made + 1
    End If
  End If

  EMReadScreen   BILS_line_05, 54, 10, 26
  BILS_line_05 = replace(BILS_line_05, "$", " ")
  BILS_line_05 = split(BILS_line_05, "  ")
  BILS_line_05(1) = replace(BILS_line_05(1), " ", "/")
  If IsDate(BILS_line_05(1)) = True then 
    If datediff("d", target_date, BILS_line_05(1)) > 0 and BILS_line_05(2) = 27 and BILS_line_05(5) <> "260.00" then  
    EMWriteScreen "260.00", 10, 48
    EMWriteScreen "c", 10, 24
    updates_made = updates_made + 1
    End If
  End If

  EMReadScreen   BILS_line_06, 54, 11, 26
  BILS_line_06 = replace(BILS_line_06, "$", " ")
  BILS_line_06 = split(BILS_line_06, "  ")
  BILS_line_06(1) = replace(BILS_line_06(1), " ", "/")
  If IsDate(BILS_line_06(1)) = True then 
    If datediff("d", target_date, BILS_line_06(1)) > 0 and BILS_line_06(2) = 27 and BILS_line_06(5) <> "260.00" then  
    EMWriteScreen "260.00", 11, 48
    EMWriteScreen "c", 11, 24
    updates_made = updates_made + 1
    End If
  End If

  EMReadScreen   BILS_line_07, 54, 12, 26
  BILS_line_07 = replace(BILS_line_07, "$", " ")
  BILS_line_07 = split(BILS_line_07, "  ")
  BILS_line_07(1) = replace(BILS_line_07(1), " ", "/")
  If IsDate(BILS_line_07(1)) = True then 
    If datediff("d", target_date, BILS_line_07(1)) > 0 and BILS_line_07(2) = 27 and BILS_line_07(5) <> "260.00" then  
    EMWriteScreen "260.00", 12, 48
    EMWriteScreen "c", 12, 24
    updates_made = updates_made + 1
    End If
  End If

  EMReadScreen   BILS_line_08, 54, 13, 26
  BILS_line_08 = replace(BILS_line_08, "$", " ")
  BILS_line_08 = split(BILS_line_08, "  ")
  BILS_line_08(1) = replace(BILS_line_08(1), " ", "/")
  If IsDate(BILS_line_08(1)) = True then 
    If datediff("d", target_date, BILS_line_08(1)) > 0 and BILS_line_08(2) = 27 and BILS_line_08(5) <> "260.00" then  
    EMWriteScreen "260.00", 13, 48
    EMWriteScreen "c", 13, 24
    updates_made = updates_made + 1
    End If
  End If

  EMReadScreen   BILS_line_09, 54, 14, 26
  BILS_line_09 = replace(BILS_line_09, "$", " ")
  BILS_line_09 = split(BILS_line_09, "  ")
  BILS_line_09(1) = replace(BILS_line_09(1), " ", "/")
  If IsDate(BILS_line_09(1)) = True then 
    If datediff("d", target_date, BILS_line_09(1)) > 0 and BILS_line_09(2) = 27 and BILS_line_09(5) <> "260.00" then  
    EMWriteScreen "260.00", 14, 48
    EMWriteScreen "c", 14, 24
    updates_made = updates_made + 1
    End If
  End If

  EMReadScreen   BILS_line_10, 54, 15, 26
  BILS_line_10 = replace(BILS_line_10, "$", " ")
  BILS_line_10 = split(BILS_line_10, "  ")
  BILS_line_10(1) = replace(BILS_line_10(1), " ", "/")
  If IsDate(BILS_line_10(1)) = True then 
    If datediff("d", target_date, BILS_line_10(1)) > 0 and BILS_line_10(2) = 27 and BILS_line_10(5) <> "260.00" then  
    EMWriteScreen "260.00", 15, 48
    EMWriteScreen "c", 15, 24
    updates_made = updates_made + 1
    End If
  End If

  EMReadScreen   BILS_line_11, 54, 16, 26
  BILS_line_11 = replace(BILS_line_11, "$", " ")
  BILS_line_11 = split(BILS_line_11, "  ")
  BILS_line_11(1) = replace(BILS_line_11(1), " ", "/")
  If IsDate(BILS_line_11(1)) = True then 
    If datediff("d", target_date, BILS_line_11(1)) > 0 and BILS_line_11(2) = 27 and BILS_line_11(5) <> "260.00" then  
    EMWriteScreen "260.00", 16, 48
    EMWriteScreen "c", 16, 24
    updates_made = updates_made + 1
    End If
  End If

  EMReadScreen   BILS_line_12, 54, 17, 26
  BILS_line_12 = replace(BILS_line_12, "$", " ")
  BILS_line_12 = split(BILS_line_12, "  ")
  BILS_line_12(1) = replace(BILS_line_12(1), " ", "/")
  If IsDate(BILS_line_12(1)) = True then 
    If datediff("d", target_date, BILS_line_12(1)) > 0 and BILS_line_12(2) = 27 and BILS_line_12(5) <> "260.00" then  
    EMWriteScreen "260.00", 17, 48
    EMWriteScreen "c", 17, 24
    updates_made = updates_made + 1
    End If
  End If

  EMReadScreen current_page, 1, 3, 73
  EMReadScreen total_pages, 1, 3, 78
  If cint(current_page) <> cint(total_pages) then
  EMSendKey "<PF20>"
  EMWaitReady 0, 0
  End If

Loop until cint(current_page) = cint(total_pages)

EMSendKey "<PF3>"
EMWaitReady 0, 0

EMSendKey "<PF3>"
EMWaitReady 0, 0

If updates_made <> 0 then MsgBox "Success! Updates made: " & updates_made & "."
If updates_made = 0 then MsgBox "Success! However, there were no remedial care entries found for after 07/01/2012. You may have already updated this case! Otherwise, this client may be at their renewal, or no remedial care deduction was made. If this appears to be an error, contact the script administrator."

script_end_procedure("")






