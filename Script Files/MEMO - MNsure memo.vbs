'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MEMO - MNsure memo"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script



'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog MNsure_info_dialog, 0, 0, 196, 120, "MNsure Info Dialog"
  EditBox 60, 5, 70, 15, case_number
  DropListBox 110, 25, 75, 15, "denied"+chr(9)+"closed", how_case_ended
  EditBox 110, 45, 70, 15, denial_effective_date
  OptionGroup RadioGroup1
    RadioButton 20, 80, 35, 10, "WCOM", WCOM_check
    RadioButton 65, 80, 35, 10, "MEMO", MEMO_check
  EditBox 70, 100, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 140, 80, 50, 15
    CancelButton 140, 100, 50, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 30, 100, 10, "Was case closed or denied?:"
  Text 5, 50, 100, 10, "Denial/closure effective date:"
  GroupBox 10, 70, 100, 25, "How are you sending this?"
  Text 5, 105, 60, 10, "Worker signature:"
EndDialog



'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""
EMFocus

'Searches for a case number
call MAXIS_case_number_finder(case_number)

'Shows dialog, checks for MAXIS or WCOM status.
Do
  Do
    Dialog MNsure_info_dialog
    If ButtonPressed = 0 then stopscript
    If isdate(denial_effective_date) = False then MsgBox "You must put in a valid denial effective date (MM/DD/YYYY)."
  Loop until isdate(denial_effective_date) = True
  transmit 'sending refresh
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "MAXIS is not found. Check to make sure you're in MAXIS production and not passworded out."
Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "

'For the WCOM option it needs to go to SPEC/WCOM. Otherwise it goes to MEMO.
If radiogroup1 = 0 then
  'Navigating to SPEC/WCOM
  call navigate_to_screen("SPEC", "WCOM")  
  'This checks to make sure we've moved passed SELF.
  EMReadScreen SELF_check, 27, 2, 28
  If SELF_check = "Select Function Menu (SELF)" then script_end_procedure("Unable to get past SELF menu. Check for error messages and try again.")   
  'Updates to show HC only memos
  EMWriteScreen "Y", 3, 74
  transmit
  'Checks to make sure there's a waiting notice
  EMReadScreen waiting_check, 7, 7, 71
  If waiting_check <> "Waiting" then script_end_procedure("No waiting notice was found. You might be in the wrong footer month. If you still have this problem email your script administrator your footer month and case number. Also include a description of what's wrong.")
  'Creates a new WCOM. If it's unable the script will stop.
  EMWriteScreen "x", 7, 13
  transmit
  PF9
  EMReadScreen client_copy_check, 11, 1, 38
  If client_copy_check = "Client Copy" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
Else
  'Navigating to SPEC/MEMO
  call navigate_to_screen("SPEC", "MEMO")  
  'This checks to make sure we've moved passed SELF.
  EMReadScreen SELF_check, 27, 2, 28
  If SELF_check = "Select Function Menu (SELF)" then script_end_procedure("Unable to get past SELF menu. Check for error messages and try again.")   
  'Creates a new MEMO. If it's unable the script will stop.
  PF5
  EMReadScreen memo_display_check, 12, 2, 33
  If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
  EMWriteScreen "x", 5, 10
  transmit
End if

'Sends the home key to get to the top of the memo.
EMSendKey "<home>" 

'Enters different text for denials vs closures. This adds the different text to the first line
If how_case_ended = "denied" then EMSendKey "Your application was denied " 
If how_case_ended = "closed" then EMSendKey "Your case was closed " 

'Now it sends the rest of the memo, saves the memo and exits the memo screen
EMSendKey "effective " & denial_effective_date & "." & "<newline>" & "<newline>" & "You may be able to purchase medical insurance through MNsure. If your family is under an income limit you may get financial help to purchase insurance. You can apply online at www.mnsure.org. If you have questions or need help to apply you can call the MNsure Call Center at 1-855-366-7873."
PF4
PF3

'Navigates to case note
call navigate_to_screen("CASE", "NOTE")
PF9

'Enters case note
If radiogroup1 = 0 then EMSendKey "Added MNsure info to client notice via WCOM. -" & worker_signature
If radiogroup1 = 1 then EMSendKey "Sent client MNsure info via MEMO. -" & worker_signature

script_end_procedure("")






