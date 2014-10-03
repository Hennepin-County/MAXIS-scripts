'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MEMO - MFIP orientation"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog MFIP_orientation_dialog, 0, 0, 191, 140, "MFIP orientation letter"
  ButtonGroup ButtonPressed
    OkButton 70, 115, 50, 15
    CancelButton 125, 115, 50, 15
  EditBox 105, 5, 55, 15, case_number
  EditBox 105, 25, 55, 15, orientation_date
  EditBox 105, 45, 55, 15, orientation_time
  DropListBox 105, 65, 60, 45, county_office_list, interview_location
  EditBox 105, 90, 55, 15, worker_signature
  Text 25, 5, 65, 15, "Case Number"
  Text 25, 25, 70, 10, "Orientation Date"
  Text 25, 50, 65, 10, "Orientation Time"
  Text 25, 90, 60, 10, "Worker Signature"
  Text 25, 65, 60, 15, "Location"
EndDialog





'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

'Searches for a case number
row = 1
col = 1
EMSearch "Case Nbr: ", row, col
EMReadScreen case_number, 8, row, col + 10
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If isnumeric(case_number) = False then case_number = ""


'This Do...loop shows the appointment letter dialog, and contains logic to require most fields.
Do
 Do 
  Do 
   Do
     Do
	  Dialog MFIP_orientation_dialog
      If ButtonPressed = cancel then stopscript
      If isnumeric(case_number) = False or len(case_number) > 8 then MsgBox "You must fill in a valid case number. Please try again."
      Loop until isnumeric(case_number) = True and len(case_number) <= 8
     If isdate(orientation_date) = False then MsgBox "You did not enter a valid  date (MM/DD/YYYY format). Please try again."
     Loop until isdate(orientation_date) = True 
    If orientation_time = "" then MsgBox "You must type an interview time. Please try again."
    Loop until orientation_time <> ""
   If worker_signature = "" then MsgBox "You must provide a signature for your case note."
  Loop until worker_signature <> ""
  transmit
  EMReadScreen MAXIS_check, 5, 1, 39
  IF MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You need to be in MAXIS for this to work. Please try again."
Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "

'Using custom function to assign addresses to the selected office
call assign_county_address_variables(county_address_line_01, county_address_line_02)


'Navigating to SPEC/MEMO
call navigate_to_screen("SPEC", "MEMO")

'This checks to make sure we've moved passed SELF.
EMReadScreen SELF_check, 27, 2, 28
If SELF_check = "Select Function Menu (SELF)" then StopScript 

'Creates a new MEMO. If it's unable the script will stop.
PF5
EMReadScreen memo_display_check, 12, 2, 33
If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
EMWriteScreen "x", 5, 10
transmit

'Writes the MEMO.
EMSetCursor 3, 15
EMSendKey "************************************************************"
EMSendKey "You are required to attend a Financial Orientation for MFIP. "
EMSendKey "Your orientation is scheduled on " & orientation_date & " at " & orientation_time & "." & "<newline>"
EMSendKey "Your orientation is scheduled at the " & interview_location & " office located at: " & "<newline>"
EMSendKey county_address_line_01 & "<newline>"
EMSendKey county_address_line_02 & "<newline>"
EMSendKey "If you cannot attend this orientation, please contact the agency office to reschedule.  Failure to attend an orientation will result in a sanction (reduction) of your MFIP benefits." & "<newline>"
EMSendKey "************************************************************"

'Exits the MEMO
PF4


'Navigates to CASE/NOTE
call navigate_to_screen("case", "note")
PF9

'Writes the case note
EMSendKey "***MFIP orientation scheduled***" & "<newline>"
EMSendKey "Appt letter sent through spec / memo." & "<newline>"
EMSendKey "Orientation is scheduled on " & orientation_date & " at " & orientation_time & "<newline>"
EMSendKey "Location: " & interview_location & "<newline>"
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

script_end_procedure("")





