'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - CLIENT CONTACT (CALL CENTER VERSION).vbs"
start_time = timer

'FUNCTIONS----------------------------------------------------------------------------------------------------
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


'THE DIALOG----------------------------------------------------------------------------------------------------
BeginDialog contact_dialog, 0, 0, 386, 175, "Client contact"
  ComboBox 65, 5, 85, 15, "client"+chr(9)+"AREP"+chr(9)+"non-AREP"+chr(9)+"SWKR", who_contacted
  EditBox 180, 5, 200, 15, regarding
  EditBox 85, 35, 60, 15, phone_number
  EditBox 255, 35, 115, 15, when_contact_was_made
  EditBox 55, 60, 85, 15, case_number
  EditBox 55, 80, 325, 15, issue
  CheckBox 10, 100, 75, 10, "Answered question", answered_question_check
  CheckBox 100, 100, 80, 10, "Transferred question", transferred_question_check
  CheckBox 195, 100, 150, 10, "Reminded Client re: Importance of CAF I", caf_1_check
  EditBox 55, 115, 325, 15, other_action
  EditBox 310, 135, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 270, 155, 50, 15
    CancelButton 330, 155, 50, 15
  Text 5, 10, 55, 10, "Phone call from:"
  Text 160, 10, 15, 10, "Re:"
  GroupBox 5, 25, 370, 30, "Optional info:"
  Text 30, 40, 50, 10, "Phone number: "
  Text 165, 40, 85, 10, "When was contact made: "
  Text 5, 65, 50, 10, "Case number: "
  Text 5, 85, 50, 10, "Issue/subject: "
  Text 5, 120, 45, 10, "Other action: "
  Text 235, 140, 70, 10, "Sign your case note: "
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------


'Connects to BlueZone
EMConnect ""

'Finds the case number
row = 1
col = 1
EMSearch "Case Nbr:", row, col
If row <> 0 then 
  EMReadScreen case_number, 8, row, col + 10
  case_number = replace(case_number, "_", "")
  case_number = trim(case_number)
End if

'Updates the "when contact was made" variable to show the current time
when_contact_was_made = date & ", " & left(time, 5) & " " & right(time, 2)

'Shows the dialog
Do
  Do
    Do
      Dialog contact_dialog
      If buttonpressed = 0 then stopscript
      If isnumeric(case_number) = False then MsgBox "You must enter a valid MAXIS case number."
    Loop until (isnumeric(case_number) = True) 
    transmit
    If isnumeric(case_number) = True then
      EMReadScreen MAXIS_check, 5, 1, 39
      If MAXIS_check <> "MAXIS" then MsgBox "You are not in MAXIS. Navigate your screen to MAXIS and try again. You might be passworded out."
    End if
  Loop until MAXIS_check = "MAXIS"
  call navigate_to_screen("case", "note")
  PF9
  EMReadScreen mode_check, 7, 20, 3
  If mode_check <> "Mode: A" and mode_check <> "Mode: E" then MsgBox "The script doesn't appear to be able to find your case note. Are you in inquiry? If so, navigate to production on the screen where you clicked the script button, and try again. Otherwise, you might have forgotten to type a valid case number."
Loop until mode_check = "Mode: A" or mode_check = "Mode: E"

'Case noting
EMSendKey "Call center received phone call from " & who_contacted
If regarding <> "" then EMSendKey " re: " & regarding 
EMSendKey "<newline>"
If when_contact_was_made <> "" then Call write_editbox_in_case_note("Contact made", when_contact_was_made, 6)
If phone_number <> "" then Call write_editbox_in_case_note("Phone number", phone_number, 6)
If issue <> "" then Call write_editbox_in_case_note("Issue/subject", issue, 6)
If answered_question_check = 1 then call write_new_line_in_case_note("* Call center was able to answer client question.")
If transferred_question_check = 1 then call write_new_line_in_case_note("* Call center was unable to answer client question, transferred to worker.")
If caf_1_check = 1 then call write_new_line_in_case_note("* Reminded client about the importance of completing the CAF 1.")
If other_action <> "" then Call write_editbox_in_case_note("Other actions", other_action, 6)
Call write_new_line_in_case_note("---")
Call write_new_line_in_case_note(worker_signature)
script_end_procedure("")











