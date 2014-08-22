'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "DAIL - citizenship verified"
start_time = timer


'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connecting
EMConnect ""

'Setting variables
row = 1
col = 1

'Finding case number
EMSearch "CASE NBR:", row, col
If row <> 0 then 
  EMReadScreen case_number, 8, row, col + 10
  case_number = replace(case_number, "_", "")
  case_number = trim(case_number)
End if

'Error out in case it can't find the case number
If row = 0 then script_end_procedure("A case number could not be found on this DAIL message. Use the ''MAXIS notes'' version of the script at this time.")

'The following grabs all of the MEMI:CITIZENSHIP messages for this particular case
EMSearch "CASE NBR:", row, col
If row <> 0 then
  EMReadScreen new_case_number, 8, row, col + 10
  new_case_number = replace(new_case_number, "_", "")
  new_case_number = trim(new_case_number)
  If new_case_number = case_number then 
    EMReadScreen cit_ID_verified_check, 46, row + 1, 20
    If cit_ID_verified_check = "MEMI:CITIZENSHIP HAS BEEN VERIFIED THROUGH SSA" then
      EMReadScreen HH_memb_01, 56, row, 5
      HH_memb_01 = replace(HH_memb_01, "--", "")
      cut_position_01 = InStr(HH_memb_01, "(")
      HH_memb_01 = Right(HH_memb_01, len(HH_memb_01) - cut_position_01)
      HH_memb_01 = replace(HH_memb_01, ")", "")
      exempt_reason_01 = "n/a"
      cit_proof_01 = "Elect. verif."
      ID_proof_01 = "Elect. verif."
      row = row + 1
    End if
  End If
End if

EMSearch "CASE NBR:", row, col
If row <> 0 then
  EMReadScreen new_case_number, 8, row, col + 10
  new_case_number = replace(new_case_number, "_", "")
  new_case_number = trim(new_case_number)
  If new_case_number = case_number then 
    EMReadScreen cit_ID_verified_check, 46, row + 1, 20
    If cit_ID_verified_check = "MEMI:CITIZENSHIP HAS BEEN VERIFIED THROUGH SSA" then
      EMReadScreen HH_memb_02, 56, row, 5
      HH_memb_02 = replace(HH_memb_02, "--", "")
      cut_position_02 = InStr(HH_memb_02, "(")
      HH_memb_02 = Right(HH_memb_02, len(HH_memb_02) - cut_position_02)
      HH_memb_02 = replace(HH_memb_02, ")", "")
      exempt_reason_02 = "n/a"
      cit_proof_02 = "Elect. verif."
      ID_proof_02 = "Elect. verif."
      row = row + 1
    End if
  End If
End if

EMSearch "CASE NBR:", row, col
If row <> 0 then
  EMReadScreen new_case_number, 8, row, col + 10
  new_case_number = replace(new_case_number, "_", "")
  new_case_number = trim(new_case_number)
  If new_case_number = case_number then 
    EMReadScreen cit_ID_verified_check, 46, row + 1, 20
    If cit_ID_verified_check = "MEMI:CITIZENSHIP HAS BEEN VERIFIED THROUGH SSA" then
      EMReadScreen HH_memb_03, 56, row, 5
      HH_memb_03 = replace(HH_memb_03, "--", "")
      cut_position_03 = InStr(HH_memb_03, "(")
      HH_memb_03 = Right(HH_memb_03, len(HH_memb_03) - cut_position_03)
      HH_memb_03 = replace(HH_memb_03, ")", "")
      exempt_reason_03 = "n/a"
      cit_proof_03 = "Elect. verif."
      ID_proof_03 = "Elect. verif."
      row = row + 1
    End if
  End If
End if

EMSearch "CASE NBR:", row, col
If row <> 0 then
  EMReadScreen new_case_number, 8, row, col + 10
  new_case_number = replace(new_case_number, "_", "")
  new_case_number = trim(new_case_number)
  If new_case_number = case_number then 
    EMReadScreen cit_ID_verified_check, 46, row + 1, 20
    If cit_ID_verified_check = "MEMI:CITIZENSHIP HAS BEEN VERIFIED THROUGH SSA" then
      EMReadScreen HH_memb_04, 56, row, 5
      HH_memb_04 = replace(HH_memb_04, "--", "")
      cut_position_04 = InStr(HH_memb_04, "(")
      HH_memb_04 = Right(HH_memb_04, len(HH_memb_04) - cut_position_04)
      HH_memb_04 = replace(HH_memb_04, ")", "")
      exempt_reason_04 = "n/a"
      cit_proof_04 = "Elect. verif."
      ID_proof_04 = "Elect. verif."
      row = row + 1
    End if
  End If
End if

EMSearch "CASE NBR:", row, col
If row <> 0 then
  EMReadScreen new_case_number, 8, row, col + 10
  new_case_number = replace(new_case_number, "_", "")
  new_case_number = trim(new_case_number)
  If new_case_number = case_number then 
    EMReadScreen cit_ID_verified_check, 46, row + 1, 20
    If cit_ID_verified_check = "MEMI:CITIZENSHIP HAS BEEN VERIFIED THROUGH SSA" then
      EMReadScreen HH_memb_05, 56, row, 5
      HH_memb_05 = replace(HH_memb_05, "--", "")
      cut_position_05 = InStr(HH_memb_05, "(")
      HH_memb_05 = Right(HH_memb_05, len(HH_memb_05) - cut_position_05)
      HH_memb_05 = replace(HH_memb_05, ")", "")
      exempt_reason_05 = "n/a"
      cit_proof_05 = "Elect. verif."
      ID_proof_05 = "Elect. verif."
      row = row + 1
    End if
  End If
End if

EMSearch "CASE NBR:", row, col
If row <> 0 then
  EMReadScreen new_case_number, 8, row, col + 10
  new_case_number = replace(new_case_number, "_", "")
  new_case_number = trim(new_case_number)
  If new_case_number = case_number then 
    EMReadScreen cit_ID_verified_check, 46, row + 1, 20
    If cit_ID_verified_check = "MEMI:CITIZENSHIP HAS BEEN VERIFIED THROUGH SSA" then
      EMReadScreen HH_memb_06, 56, row, 5
      HH_memb_06 = replace(HH_memb_06, "--", "")
      cut_position_06 = InStr(HH_memb_06, "(")
      HH_memb_06 = Right(HH_memb_06, len(HH_memb_06) - cut_position_06)
      HH_memb_06 = replace(HH_memb_06, ")", "")
      exempt_reason_06 = "n/a"
      cit_proof_06 = "Elect. verif."
      ID_proof_06 = "Elect. verif."
      row = row + 1
    End if
  End If
End if

EMSearch "CASE NBR:", row, col
If row <> 0 then
  EMReadScreen new_case_number, 8, row, col + 10
  new_case_number = replace(new_case_number, "_", "")
  new_case_number = trim(new_case_number)
  If new_case_number = case_number then 
    EMReadScreen cit_ID_verified_check, 46, row + 1, 20
    If cit_ID_verified_check = "MEMI:CITIZENSHIP HAS BEEN VERIFIED THROUGH SSA" then
      EMReadScreen HH_memb_07, 56, row, 5
      HH_memb_07 = replace(HH_memb_07, "--", "")
      cut_position_07 = InStr(HH_memb_07, "(")
      HH_memb_07 = Right(HH_memb_07, len(HH_memb_07) - cut_position_07)
      HH_memb_07 = replace(HH_memb_07, ")", "")
      exempt_reason_07 = "n/a"
      cit_proof_07 = "Elect. verif."
      ID_proof_07 = "Elect. verif."
      row = row + 1
    End if
  End If
End if

EMSearch "CASE NBR:", row, col
If row <> 0 then
  EMReadScreen new_case_number, 8, row, col + 10
  new_case_number = replace(new_case_number, "_", "")
  new_case_number = trim(new_case_number)
  If new_case_number = case_number then 
    EMReadScreen cit_ID_verified_check, 46, row + 1, 20
    If cit_ID_verified_check = "MEMI:CITIZENSHIP HAS BEEN VERIFIED THROUGH SSA" then
      EMReadScreen HH_memb_08, 56, row, 5
      HH_memb_08 = replace(HH_memb_08, "--", "")
      cut_position_08 = InStr(HH_memb_08, "(")
      HH_memb_08 = Right(HH_memb_08, len(HH_memb_08) - cut_position_08)
      HH_memb_08 = replace(HH_memb_08, ")", "")
      exempt_reason_08 = "n/a"
      cit_proof_08 = "Elect. verif."
      ID_proof_08 = "Elect. verif."
      row = row + 1
    End if
  End If
End if

   
'The dialog
BeginDialog cit_ID_dialog, 0, 0, 346, 222, "CIT-ID dialog"
  Text 5, 10, 50, 10, "Case number:"
  EditBox 60, 5, 75, 15, case_number
  Text 20, 25, 45, 10, "HH member"
  Text 85, 25, 55, 10, "Exempt reason"
  Text 200, 25, 35, 10, "Cit proof"
  Text 290, 25, 35, 10, "ID proof"
  EditBox 5, 40, 65, 15, HH_memb_01
  ComboBox 80, 40, 85, 10, "(select or type here)"+chr(9)+"MEDI enrollee"+chr(9)+"SSI/RSDI recip."+chr(9)+"foster care"+chr(9)+"adoption assist."+chr(9)+"auto newborn", exempt_reason_01
  ComboBox 170, 40, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Birth Certificate"+chr(9)+"Nat. papers", cit_proof_01
  ComboBox 260, 40, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Drivers License"+chr(9)+"State ID"+chr(9)+"School ID"+chr(9)+"Parent Signature", ID_proof_01
  EditBox 5, 60, 65, 15, HH_memb_02
  ComboBox 80, 60, 85, 10, "(select or type here)"+chr(9)+"MEDI enrollee"+chr(9)+"SSI/RSDI recip."+chr(9)+"foster care"+chr(9)+"adoption assist."+chr(9)+"auto newborn", exempt_reason_02
  ComboBox 170, 60, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Birth Certificate"+chr(9)+"Nat. papers", cit_proof_02
  ComboBox 260, 60, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Drivers License"+chr(9)+"State ID"+chr(9)+"School ID"+chr(9)+"Parent Signature", ID_proof_02
  EditBox 5, 80, 65, 15, HH_memb_03
  ComboBox 80, 80, 85, 10, "(select or type here)"+chr(9)+"MEDI enrollee"+chr(9)+"SSI/RSDI recip."+chr(9)+"foster care"+chr(9)+"adoption assist."+chr(9)+"auto newborn", exempt_reason_03
  ComboBox 170, 80, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Birth Certificate"+chr(9)+"Nat. papers", cit_proof_03
  ComboBox 260, 80, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Drivers License"+chr(9)+"State ID"+chr(9)+"School ID"+chr(9)+"Parent Signature", ID_proof_03
  EditBox 5, 100, 65, 15, HH_memb_04
  ComboBox 80, 100, 85, 10, "(select or type here)"+chr(9)+"MEDI enrollee"+chr(9)+"SSI/RSDI recip."+chr(9)+"foster care"+chr(9)+"adoption assist."+chr(9)+"auto newborn", exempt_reason_04
  ComboBox 170, 100, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Birth Certificate"+chr(9)+"Nat. papers", cit_proof_04
  ComboBox 260, 100, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Drivers License"+chr(9)+"State ID"+chr(9)+"School ID"+chr(9)+"Parent Signature", ID_proof_04
  EditBox 5, 120, 65, 15, HH_memb_05
  ComboBox 80, 120, 85, 10, "(select or type here)"+chr(9)+"MEDI enrollee"+chr(9)+"SSI/RSDI recip."+chr(9)+"foster care"+chr(9)+"adoption assist."+chr(9)+"auto newborn", exempt_reason_05
  ComboBox 170, 120, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Birth Certificate"+chr(9)+"Nat. papers", cit_proof_05
  ComboBox 260, 120, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Drivers License"+chr(9)+"State ID"+chr(9)+"School ID"+chr(9)+"Parent Signature", ID_proof_05
  EditBox 5, 140, 65, 15, HH_memb_06
  ComboBox 80, 140, 85, 10, "(select or type here)"+chr(9)+"MEDI enrollee"+chr(9)+"SSI/RSDI recip."+chr(9)+"foster care"+chr(9)+"adoption assist."+chr(9)+"auto newborn", exempt_reason_06
  ComboBox 170, 140, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Birth Certificate"+chr(9)+"Nat. papers", cit_proof_06
  ComboBox 260, 140, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Drivers License"+chr(9)+"State ID"+chr(9)+"School ID"+chr(9)+"Parent Signature", ID_proof_06
  EditBox 5, 160, 65, 15, HH_memb_07
  ComboBox 80, 160, 85, 10, "(select or type here)"+chr(9)+"MEDI enrollee"+chr(9)+"SSI/RSDI recip."+chr(9)+"foster care"+chr(9)+"adoption assist."+chr(9)+"auto newborn", exempt_reason_07
  ComboBox 170, 160, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Birth Certificate"+chr(9)+"Nat. papers", cit_proof_07
  ComboBox 260, 160, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Drivers License"+chr(9)+"State ID"+chr(9)+"School ID"+chr(9)+"Parent Signature", ID_proof_07
  EditBox 5, 180, 65, 15, HH_memb_08
  ComboBox 80, 180, 85, 10, "(select or type here)"+chr(9)+"MEDI enrollee"+chr(9)+"SSI/RSDI recip."+chr(9)+"foster care"+chr(9)+"adoption assist."+chr(9)+"auto newborn", exempt_reason_08
  ComboBox 170, 180, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Birth Certificate"+chr(9)+"Nat. papers", cit_proof_08
  ComboBox 260, 180, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Drivers License"+chr(9)+"State ID"+chr(9)+"School ID"+chr(9)+"Parent Signature", ID_proof_08
  Text 5, 205, 65, 10, "Sign the case note:"
  EditBox 75, 200, 95, 15, worker_sig
  ButtonGroup ButtonPressed
    OkButton 195, 200, 50, 15
    CancelButton 250, 200, 50, 15
EndDialog



'Shows dialog and then seeks out case/note
Do
  Do
    Dialog cit_ID_dialog
    If buttonpressed = 0 then stopscript
    transmit
    EMReadScreen MAXIS_check, 5, 1, 39
    If MAXIS_check <> "MAXIS" then MsgBox "You are not in MAXIS. Navigate your ''S1'' screen to MAXIS and try again. You might be passworded out."
  Loop until MAXIS_check = "MAXIS"
  EMReadScreen mode_check, 7, 20, 3
  If mode_check <> "Mode: A" and mode_check <> "Mode: E" then
    call navigate_to_screen("case", "note")
    PF9
    EMReadScreen mode_check, 7, 20, 3
    If mode_check <> "Mode: A" and mode_check <> "Mode: E" then MsgBox "The script doesn't appear to be able to find your case note. Are you in inquiry? If so, navigate to production on the screen where you clicked the script button, and try again. Otherwise, you might have forgotten to type a valid case number."
  End if
Loop until mode_check = "Mode: A" or mode_check = "Mode: E"

'Sends the case note
EMSendKey "***CITIZENSHIP/IDENTITY***" & "<newline>"
EMSendKey string(77, "-") 
EMSendKey "    HH MEMB         EXEMPT REASON            CIT PROOF         ID PROOF" & "<newline>"
If HH_memb_01 <> "" then 
  EMWriteScreen string(76, " "), 7, 3
  EMWriteScreen HH_memb_01, 7, 5
  IF exempt_reason_01 <> "(select or type here)" then EMWriteScreen exempt_reason_01, 7, 22
  IF cit_proof_01 <> "(select or type here)" then EMWriteScreen cit_proof_01, 7, 45
  IF ID_proof_01 <> "(select or type here)" then EMWriteScreen ID_proof_01, 7, 63
End if
If HH_memb_02 <> "" then
  EMWriteScreen string(76, " "), 8, 3
  EMWriteScreen HH_memb_02, 8, 5
  IF exempt_reason_02 <> "(select or type here)" then EMWriteScreen exempt_reason_02, 8, 22
  IF cit_proof_02 <> "(select or type here)" then EMWriteScreen cit_proof_02, 8, 45
  IF ID_proof_02 <> "(select or type here)" then EMWriteScreen ID_proof_02, 8, 63
End if
If HH_memb_03 <> "" then
  EMWriteScreen string(76, " "), 9, 3
  EMWriteScreen HH_memb_03, 9, 5
  IF exempt_reason_03 <> "(select or type here)" then EMWriteScreen exempt_reason_03, 9, 22
  IF cit_proof_03 <> "(select or type here)" then EMWriteScreen cit_proof_03, 9, 45
  IF ID_proof_03 <> "(select or type here)" then EMWriteScreen ID_proof_03, 9, 63
End if
If HH_memb_04 <> "" then
  EMWriteScreen string(76, " "), 10, 3
  EMWriteScreen HH_memb_04, 10, 5
  IF exempt_reason_04 <> "(select or type here)" then EMWriteScreen exempt_reason_04, 10, 22
  IF cit_proof_04 <> "(select or type here)" then EMWriteScreen cit_proof_04, 10, 45
  IF ID_proof_04 <> "(select or type here)" then EMWriteScreen ID_proof_04, 10, 63
End if
If HH_memb_05 <> "" then
  EMWriteScreen string(76, " "), 11, 3
  EMWriteScreen HH_memb_05, 11, 5
  IF exempt_reason_05 <> "(select or type here)" then EMWriteScreen exempt_reason_05, 11, 22
  IF cit_proof_05 <> "(select or type here)" then EMWriteScreen cit_proof_05, 11, 45
  IF ID_proof_05 <> "(select or type here)" then EMWriteScreen ID_proof_05, 11, 63
End if
If HH_memb_06 <> "" then
  EMWriteScreen string(76, " "), 12, 3
  EMWriteScreen HH_memb_06, 12, 5
  IF exempt_reason_06 <> "(select or type here)" then EMWriteScreen exempt_reason_06, 12, 22
  IF cit_proof_06 <> "(select or type here)" then EMWriteScreen cit_proof_06, 12, 45
  IF ID_proof_06 <> "(select or type here)" then EMWriteScreen ID_proof_06, 12, 63
End if
If HH_memb_07 <> "" then
  EMWriteScreen string(76, " "), 13, 3
  EMWriteScreen HH_memb_07, 13, 5
  IF exempt_reason_07 <> "(select or type here)" then EMWriteScreen exempt_reason_07, 13, 22
  IF cit_proof_07 <> "(select or type here)" then EMWriteScreen cit_proof_07, 13, 45
  IF ID_proof_07 <> "(select or type here)" then EMWriteScreen ID_proof_07, 13, 63
End if
If HH_memb_08 <> "" then
  EMWriteScreen string(76, " "), 14, 3
  EMWriteScreen HH_memb_08, 14, 5
  IF exempt_reason_08 <> "(select or type here)" then EMWriteScreen exempt_reason_08, 14, 22
  IF cit_proof_08 <> "(select or type here)" then EMWriteScreen cit_proof_08, 14, 45
  IF ID_proof_08 <> "(select or type here)" then EMWriteScreen ID_proof_08, 14, 63
End if
EMSetCursor 15, 3
EMSendKey string(77, "-") & "<newline>"
Call write_new_line_in_case_note(worker_sig)

script_end_procedure("")






