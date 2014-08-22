'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - view INFC"
start_time = timer

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'CONNECTS TO MAXIS, SEEKS CASE NUMBER
EMConnect ""

row = 1
col = 1
EMSearch "Case Nbr: ", row, col
If row <> 0 then EMReadScreen case_number, 8, row, col + 10

'DIALOG FOR VIEWING AN INFC PANEL. DIALOG WILL ALLOW YOU TO SELECT BNDX, SDXS OR SVES/TPQY
BeginDialog view_INFC_dialog, 0, 0, 156, 102, "View INFC"
  EditBox 90, 5, 60, 15, case_number
  EditBox 125, 25, 25, 15, member_number
  DropListBox 65, 45, 75, 15, "BNDX"+chr(9)+"SDXS"+chr(9)+"SVES/TPQY", view_panel
  DropListBox 65, 65, 80, 10, "production"+chr(9)+"inquiry", results_screen
  ButtonGroup ButtonPressed
    OkButton 25, 85, 50, 15
    CancelButton 85, 85, 50, 15
  Text 5, 10, 80, 10, "Enter your case number:"
  Text 5, 30, 120, 10, "Enter your member number (ex: 01): "
  Text 10, 50, 55, 10, "Screen to view:"
  Text 10, 70, 50, 10, "View results in:"
EndDialog
view_panel = "SVES/TPQY" 'default setting

Dialog view_INFC_dialog
If ButtonPressed = 0 then StopScript 'Cancels if the cancel button is pressed.

'CHECKING FOR MAXIS
EMConnect "A"
attn
EMReadScreen MAI_check, 3, 1, 33
If MAI_check <> "MAI" then EMWaitReady 0, 0
EMReadScreen production_check, 7, 6, 15
EMReadScreen inquiry_check, 7, 7, 15
If inquiry_check = "RUNNING" and results_screen = "inquiry" then 
  EMWriteScreen "s", 7, 2    
  transmit
End if
If production_check = "RUNNING" and results_screen = "production" then 
  EMWriteScreen "s", 6, 2   
  transmit
End if
If inquiry_check <> "RUNNING" and results_screen = "inquiry" then 
  attn
  EMConnect "B"
  attn
  EMReadScreen MAI_check, 3, 1, 33
  If MAI_check <> "MAI" then EMWaitReady 0, 0
  EMReadScreen inquiry_B_check, 7, 7, 15
  If inquiry_B_check <> "RUNNING" then script_end_procedure("Inquiry could not be found. If inquiry is on, try running the script again. If the problem keeps happening, contact the script administrator.")
  If inquiry_B_check = "RUNNING" then
    EMWriteScreen "s", 7, 2    
    transmit
  End if
End if
If production_check <> "RUNNING" and results_screen = "production" then 
  attn
  EMConnect "B"
  attn
  EMReadScreen production_B_check, 7, 6, 15
  If production_B_check <> "RUNNING" then
    MsgBox "Production could not be found. If Production is on, try running the script again. If the problem keeps happening, contact the script administrator."
    stopscript
  End if
  If production_B_check = "RUNNING" then
    EMWriteScreen "s", 6, 2  
    transmit
  End if
End if

'TRANSMITS TO CHECK FOR PASSWORD
transmit
EMReadScreen password_prompt, 38, 2, 23
IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then StopScript

back_to_self

'GOES TO STAT/MEMB FOR THE SPECIFIC MEMBER NUMBER
EMWriteScreen "stat", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen "memb", 21, 70
EMWriteScreen member_number, 21, 75
transmit

'--------------ERROR PROOFING--------------
EMReadScreen still_self, 27, 2, 28 'This checks to make sure we've moved passed SELF.
If still_self = "ror Prone Edit Summary (ERR" then transmit
EMReadScreen no_MEMB, 13, 8, 22 'If this member does not exist, this will stop the script from continuing.
If no_MEMB = "Arrival Date:" then 
  MsgBox "This HH member does not exist."
  StopScript
End if
'--------------END ERROR PROOFING--------------
'READS THE PMI AND SSN
EMReadScreen PMI, 8, 4, 46
EMReadScreen SSN, 11, 7, 42

'NAVIGATES TO INFC
back_to_self
EMWriteScreen "infc", 16, 43
transmit

'FOR SVES/TPQY, IT HAS TO ENTER THE PMI.
If view_panel = "SVES/TPQY" then
  EMWriteScreen "sves", 20, 71
  transmit
  EMWriteScreen PMI, 5, 68
  EMWriteScreen "tpqy", 20, 70
  transmit
End if

'FOR BNDX, IT HAS TO ENTER THE SSN
If view_panel = "BNDX" then
  EMWriteScreen replace(SSN, " ", ""), 4, 63
  EMWriteScreen "bndx", 20, 71
  transmit
End if

'FOR SDXS, IT HAS TO ENTER THE SSN
If view_panel = "SDXS" then
  EMWriteScreen replace(SSN, " ", ""), 4, 63
  EMWriteScreen "SDXS", 20, 71
  transmit
End if

script_end_procedure("")






