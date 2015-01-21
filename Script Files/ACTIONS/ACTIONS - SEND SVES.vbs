'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - SEND SVES.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, FALSE									'Attempts to open the URL
req.send													'Sends request
IF req.Status = 200 THEN									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			script_end_procedure("Script ended due to error connecting to GitHub.")
END IF

'DIALOGS-----------------------------------------------------------------
BeginDialog send_SVES_dialog, 0, 0, 271, 85, "Send SVES Dialog"
  EditBox 90, 5, 60, 15, case_number
  EditBox 125, 25, 25, 15, member_number
  CheckBox 5, 50, 165, 10, "Check here to case note that a QURY was sent.", case_note_checkbox
  EditBox 80, 65, 70, 15, worker_signature
  OptionGroup RadioGroup1
    RadioButton 190, 15, 60, 15, "SSN? (default)", SSN_radiobutton
    RadioButton 190, 30, 75, 15, "Claim # on UNEA?", UNEA_radiobutton
    RadioButton 190, 45, 75, 15, "Claim # on BNDX?", BNDX_radiobutton
  ButtonGroup ButtonPressed
    OkButton 160, 65, 50, 15
    CancelButton 215, 65, 50, 15
  Text 5, 10, 80, 10, "Enter your case number:"
  Text 5, 30, 120, 10, "Enter your member number (ex: 01): "
  Text 5, 70, 70, 10, "Sign your case note:"
  GroupBox 185, 5, 80, 55, "Number to use?"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

'Grabs case number
call MAXIS_case_number_finder(case_number)

'Shows dialog
Dialog send_SVES_dialog
If ButtonPressed = cancel then StopScript 

'Defaults member number to 01
If member_number = "" then member_number = "01"

'Makes sure we're in MAXIS
call check_for_MAXIS(True)

'Goes to MEMB to get info
call navigate_to_screen("stat", "memb")

'Goes to the right HH member
EMWriteScreen member_number, 20, 76 'It does this to make sure that it navigates to the right HH member.
transmit 'This transmits to STAT/MEMB for the client indicated.

'If this member does not exist, this will stop the script from continuing.
EMReadScreen no_MEMB, 13, 8, 22 
If no_MEMB = "Arrival Date:" then script_end_procedure("Error! This HH member does not exist.")

'Reads the SSN pieces and the PMI
EMReadScreen SSN1, 3, 7, 42
EMReadScreen SSN2, 2, 7, 46
EMReadScreen SSN3, 4, 7, 49
EMReadScreen PMI, 8, 4, 46

If SSN_radiobutton = 1 then
  call navigate_to_screen("infc", "sves")
  EMWriteScreen SSN1,  4, 68
  EMWriteScreen SSN2,  4, 71
  EMWriteScreen SSN3,  4, 73
  EMWriteScreen PMI,  5, 68
  EMWriteScreen "qury",  20, 70
  transmit 'Now we will enter the QURY screen to type the case number.
ElseIf UNEA_radiobutton = 1 then
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
      If UNEA_type_01 <> "" then RadioButton 10, 20, 160, 10, "Type " & UNEA_type_01 & ", claim number " & UNEA_claim_number_01, UNEA_type_01_radiobutton
      If UNEA_type_02 <> "" then RadioButton 10, 35, 160, 10, "Type " & UNEA_type_02 & ", claim number " & UNEA_claim_number_02, UNEA_type_02_radiobutton
      If UNEA_type_03 <> "" then RadioButton 10, 50, 160, 10, "Type " & UNEA_type_03 & ", claim number " & UNEA_claim_number_03, UNEA_type_03_radiobutton
      If UNEA_type_04 <> "" then RadioButton 10, 65, 160, 10, "Type " & UNEA_type_04 & ", claim number " & UNEA_claim_number_04, UNEA_type_04_radiobutton
      If UNEA_type_05 <> "" then RadioButton 10, 80, 160, 10, "Type " & UNEA_type_05 & ", claim number " & UNEA_claim_number_05, UNEA_type_05_radiobutton
    EndDialog
  
    Dialog UNEA_claim_dialog
    If ButtonPressed = 0 then stopscript
    If UNEA_type_01_radiobutton = 1 then
      claim_number = UNEA_claim_number_01
    ElseIf UNEA_type_02_radiobutton = 1 then
      claim_number = UNEA_claim_number_02
    ElseIf UNEA_type_03_radiobutton = 1 then
      claim_number = UNEA_claim_number_03
    ElseIf UNEA_type_04_radiobutton = 1 then
      claim_number = UNEA_claim_number_04
    ElseIf UNEA_type_05_radiobutton = 1 then
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
ElseIf BNDX_radiobutton = 1 then
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
      If BNDX_claim_number_01 <> "" then RadioButton 10, 20, 160, 10, BNDX_claim_number_01, BNDX_claim_number_01_radiobutton
      If BNDX_claim_number_02 <> "" then RadioButton 10, 35, 160, 10, BNDX_claim_number_02, BNDX_claim_number_02_radiobutton
      If BNDX_claim_number_03 <> "" then RadioButton 10, 50, 160, 10, BNDX_claim_number_03, BNDX_claim_number_03_radiobutton
    EndDialog
  
    Dialog BNDX_claim_dialog
    If ButtonPressed = 0 then stopscript
    If BNDX_claim_number_01_radiobutton = 1 then
      claim_number = replace(BNDX_claim_number_01, " ", "")
    ElseIf BNDX_claim_number_02_radiobutton = 1 then
      claim_number = replace(BNDX_claim_number_02, " ", "")
    ElseIf BNDX_claim_number_03_radiobutton = 1 then
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

'Shuts down here if the user does not want to case note
If case_note_checkbox = unchecked then script_end_procedure("")

'Now it sends the SVES.
transmit



'Now it case notes
call navigate_to_screen("CASE", "NOTE")
PF9
call write_variable_in_case_note("~~~SVES/QURY sent for MEMB " & member_number & "~~~")
If SSN_radiobutton = 1 then
	call write_variable_in_case_note("* Used SSN for QURY.")
ElseIf UNEA_radiobutton = 1 then
	call write_variable_in_case_note("* Used claim number on UNEA for QURY.")
ElseIf BNDX_radiobutton = 1 then
	call write_variable_in_case_note("* Used claim number on BNDX for QURY.")
End If
call write_variable_in_case_note("* QURY sent using script.")
call write_variable_in_case_note("---")
call write_variable_in_case_note(worker_signature)

script_end_procedure("")






