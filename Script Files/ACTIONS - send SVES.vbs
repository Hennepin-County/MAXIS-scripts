'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - send SVES"
start_time = timer

''LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'THE SCRIPT----------------------------------------------------------------------------------------------------

EMConnect ""

BeginDialog SVES_dialog, 0, 0, 166, 132, "SVES"
  EditBox 95, 5, 60, 15, case_number
  EditBox 130, 25, 25, 15, member_number
  OptionGroup RadioGroup1
    RadioButton 40, 60, 60, 15, "SSN? (default)", SSN_radio
    RadioButton 40, 75, 75, 15, "Claim # on UNEA?", UNEA_radio
    RadioButton 40, 90, 75, 15, "Claim # on BNDX?", BNDX_radio
  ButtonGroup ButtonPressed
    OkButton 25, 110, 50, 15
    CancelButton 90, 110, 50, 15
  Text 10, 10, 80, 10, "Enter your case number:"
  Text 10, 30, 120, 10, "Enter your member number (ex: 01): "
  GroupBox 15, 50, 135, 55, "What number to use for SVES/QURY?"
EndDialog



call find_variable("Case Nbr: ", case_number, 8) 'x is string, y is variable, z is length of new variable
case_number = trim(replace(case_number, "_", ""))

Dialog SVES_dialog
If ButtonPressed = 0 then StopScript 'Cancels if the cancel button is pressed.


If member_number = "" then member_number = "01"

'It sends an enter to force the screen to refresh, in order to check for a password prompt.
transmit
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" then script_end_procedure("MAXIS is not found on this screen. Navigate to MAXIS and try again.")

call navigate_to_screen("stat", "memb")

EMWriteScreen "memb", 20, 71 'It does this to move past error prone cases
EMWriteScreen member_number, 20, 76 'It does this to make sure that it navigates to the right HH member.
transmit 'This transmits to STAT/MEMB for the client indicated.

EMReadScreen no_MEMB, 13, 8, 22 'If this member does not exist, this will stop the script from continuing.
If no_MEMB = "Arrival Date:" then script_end_procedure("This HH member does not exist.")

EMReadScreen SSN1, 3, 7, 42
EMReadScreen SSN2, 2, 7, 46
EMReadScreen SSN3, 4, 7, 49
EMReadScreen PMI, 8, 4, 46

If SSN_radio = 1 then
  call navigate_to_screen("infc", "sves")
  EMWriteScreen SSN1,  4, 68
  EMWriteScreen SSN2,  4, 71
  EMWriteScreen SSN3,  4, 73
  EMWriteScreen PMI,  5, 68
  EMWriteScreen "qury",  20, 70
  transmit 'Now we will enter the QURY screen to type the case number.
End if

If UNEA_radio = 1 then
  call navigate_to_screen("stat", "unea")

  EMWriteScreen "unea", 20, 71 'It does this to move past error prone cases
  EMWriteScreen member_number, 20, 76 'It does this to make sure that it navigates to the right HH member.
  transmit 'This transmits to STAT/UNEA for the client indicated.

  EMReadScreen PMI, 8, 4, 71 
  PMI = trim(PMI)
  
  'If there is no PMI, then there is no UNEA panel entered for the script to work off of.
  If PMI = "" then script_end_procedure("This HH member does not exist, or does not have a UNEA panel made.")
  
  EMReadScreen amt_of_unea_panels, 1, 2, 78
  If amt_of_unea_panels <> "1" then 
    dialog_size_variable = (15 * cint(amt_of_unea_panels)) + 20
    Do
      EMReadScreen UNEA_type, 2, 5, 37
      EMReadScreen UNEA_claim_number, 15, 6, 37
      UNEA_claim_number = replace(UNEA_claim_number, "_", "")
      If UNEA_type_01 = "" then
        UNEA_type_01 = UNEA_type
        UNEA_claim_number_01 = UNEA_claim_number
      ElseIf UNEA_type_02 = "" then
        UNEA_type_02 = UNEA_type
        UNEA_claim_number_02 = UNEA_claim_number
      ElseIf UNEA_type_03 = "" then
        UNEA_type_03 = UNEA_type
        UNEA_claim_number_03 = UNEA_claim_number
      ElseIf UNEA_type_04 = "" then
        UNEA_type_04 = UNEA_type
        UNEA_claim_number_04 = UNEA_claim_number
      ElseIf UNEA_type_05 = "" then
        UNEA_type_05 = UNEA_type
        UNEA_claim_number_05 = UNEA_claim_number
      End if
      EMReadScreen current_unea_panel, 1, 2, 73
      If current_unea_panel <> amt_of_unea_panels then transmit
    Loop until current_unea_panel = amt_of_unea_panels

    BeginDialog UNEA_claim_dialog, 0, 0, 240, dialog_size_variable, "UNEA claim dialog"
      ButtonGroup ButtonPressed
        OkButton 185, 10, 50, 15
        CancelButton 185, 30, 50, 15
      Text 10, 5, 105, 10, "UNEA types to look at:"
      OptionGroup RadioGroup1
      If UNEA_type_01 <> "" then RadioButton 10, 20, 160, 10, "Type " & UNEA_type_01 & ", claim number " & UNEA_claim_number_01, UNEA_type_01_check
      If UNEA_type_02 <> "" then RadioButton 10, 35, 160, 10, "Type " & UNEA_type_02 & ", claim number " & UNEA_claim_number_02, UNEA_type_02_check
      If UNEA_type_03 <> "" then RadioButton 10, 50, 160, 10, "Type " & UNEA_type_03 & ", claim number " & UNEA_claim_number_03, UNEA_type_03_check
      If UNEA_type_04 <> "" then RadioButton 10, 65, 160, 10, "Type " & UNEA_type_04 & ", claim number " & UNEA_claim_number_04, UNEA_type_04_check
      If UNEA_type_05 <> "" then RadioButton 10, 80, 160, 10, "Type " & UNEA_type_05 & ", claim number " & UNEA_claim_number_05, UNEA_type_05_check
    EndDialog
  
    Dialog UNEA_claim_dialog
    If ButtonPressed = 0 then stopscript
    If UNEA_type_01_check = 1 then
      claim_number = UNEA_claim_number_01
    ElseIf UNEA_type_02_check = 1 then
      claim_number = UNEA_claim_number_02
    ElseIf UNEA_type_03_check = 1 then
      claim_number = UNEA_claim_number_03
    ElseIf UNEA_type_04_check = 1 then
      claim_number = UNEA_claim_number_04
    ElseIf UNEA_type_05_check = 1 then
      claim_number = UNEA_claim_number_05
    End if
  Else  
    EMReadScreen claim_number, 15, 6, 37
    claim_number = replace(claim_number, "_", "")
  End if
  call navigate_to_screen("infc", "sves")
  EMWriteScreen PMI,  5, 68
  EMWriteScreen "qury",  20, 70
  transmit 'Now we will enter the QURY screen to type the claim number.

  EMWriteScreen claim_number, 7, 38
End if

If BNDX_radio = 1 then
  call navigate_to_screen ("infc", "____")
  EMWriteScreen SSN1,  4, 63
  EMWriteScreen SSN2,  4, 66
  EMWriteScreen SSN3,  4, 68
  EMWriteScreen "bndx", 20, 71
  transmit

  EMReadScreen BNDX_claim_number_01, 13, 5, 12
  EMReadScreen BNDX_claim_number_02, 13, 5, 38
  EMReadScreen BNDX_claim_number_03, 13, 5, 64

  If BNDX_claim_number_01 = "             " then script_end_procedure("BNDX claim number is not found. Was there a BNDX message for this client? Try sending using SSN or UNEA claim number.")

  If BNDX_claim_number_02 = "             " and BNDX_claim_number_03 = "             " then 
    claim_number = replace(BNDX_claim_number_01, " ", "")
  Else
    BeginDialog BNDX_claim_dialog, 0, 0, 240, 70, "BNDX claim dialog"
      ButtonGroup ButtonPressed
        OkButton 185, 10, 50, 15
        CancelButton 185, 30, 50, 15
      Text 10, 5, 105, 10, "BNDX claims to look at:"
      OptionGroup RadioGroup1
      If BNDX_claim_number_01 <> "" then RadioButton 10, 20, 160, 10, BNDX_claim_number_01, BNDX_claim_number_01_radio
      If BNDX_claim_number_02 <> "" then RadioButton 10, 35, 160, 10, BNDX_claim_number_02, BNDX_claim_number_02_radio
      If BNDX_claim_number_03 <> "" then RadioButton 10, 50, 160, 10, BNDX_claim_number_03, BNDX_claim_number_03_radio
    EndDialog
  
    Dialog BNDX_claim_dialog
    If ButtonPressed = 0 then stopscript
    If BNDX_claim_number_01_radio = 1 then
      claim_number = replace(BNDX_claim_number_01, " ", "")
    ElseIf BNDX_claim_number_02_radio = 1 then
      claim_number = replace(BNDX_claim_number_02, " ", "")
    ElseIf BNDX_claim_number_03_radio = 1 then
      claim_number = replace(BNDX_claim_number_03, " ", "")
    End if
  End if

  PF3
  EMWriteScreen "sves", 20, 71
  transmit
  EMWriteScreen "qury", 20, 70
  transmit
  EMWriteScreen "_________", 5, 38
  EMWriteScreen claim_number, 7, 38
  EMWriteScreen "________", 9, 38
  EMWriteScreen PMI, 9, 38
End if

EMWriteScreen case_number, 11, 38
EMWriteScreen "y", 14, 38

script_end_procedure("")






