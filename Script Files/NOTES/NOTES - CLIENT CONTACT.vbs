'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - CLIENT CONTACT.vbs"
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

'THE DIALOG--------------------------------------------------------------------------------------------------
BeginDialog contact_dialog, 0, 0, 386, 280, "Client contact"
  ComboBox 50, 5, 60, 15, "Phone call"+chr(9)+"Voicemail"+chr(9)+"Email"+chr(9)+"Office visit"+chr(9)+"Letter", contact_type
  DropListBox 115, 5, 45, 10, "from"+chr(9)+"to", contact_direction
  ComboBox 165, 5, 85, 15, "client"+chr(9)+"AREP"+chr(9)+"SWKR", who_contacted
  EditBox 280, 5, 100, 15, regarding
  EditBox 95, 35, 60, 15, phone_number
  EditBox 285, 35, 85, 15, when_voicemail_was_left
  EditBox 55, 60, 85, 15, case_number
  EditBox 70, 80, 310, 15, contact_reason
  EditBox 55, 100, 325, 15, actions_taken
  EditBox 65, 135, 310, 15, verifs_needed
  EditBox 125, 155, 250, 15, cl_instructions
  EditBox 65, 175, 310, 15, case_status
  CheckBox 5, 200, 255, 10, "Check here if you want to TIKL out for this case after the case note is done.", TIKL_check
  CheckBox 5, 220, 255, 10, "Check here if you reminded client about the importance of the CAF 1.", caf_1_check
  EditBox 310, 240, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 270, 260, 50, 15
    CancelButton 330, 260, 50, 15
  Text 5, 10, 45, 10, "Contact type:"
  Text 260, 10, 15, 10, "Re:"
  GroupBox 5, 25, 370, 30, "Optional info:"
  Text 40, 40, 50, 10, "Phone number: "
  Text 195, 40, 85, 10, "When was contact made: "
  Text 5, 65, 50, 10, "Case number: "
  Text 5, 85, 65, 10, "Reason for contact:"
  Text 5, 105, 50, 10, "Actions taken: "
  GroupBox 5, 120, 375, 75, "Helpful info for call centers (or front desks) to pass on to clients"
  Text 15, 140, 50, 10, "Verifs needed: "
  Text 15, 160, 105, 10, "Instructions/message for client:"
  Text 15, 180, 45, 10, "Case status: "
  Text 240, 245, 70, 10, "Sign your case note: "
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------

EMConnect ""

row = 1
col = 1

EMSearch "Case Nbr:", row, col
If row <> 0 then 
	EMReadScreen case_number, 8, row, col + 10
	case_number = replace(case_number, "_", "")
	case_number = trim(case_number)
End if


DO
	Do
		Do
			Do
				Dialog contact_dialog
				If buttonpressed = 0 then stopscript
				IF contact_reason = "" or contact_type = "" Then MsgBox("You must enter a reason for contact, as well as a type (phone, etc.).")
			Loop until contact_reason <> "" and contact_type <> ""
			If (isnumeric(case_number) = False and len(case_number) <> 8) then MsgBox "You must enter either a valid MAXIS or MCRE case number."
		Loop until (isnumeric(case_number) = True) or (isnumeric(case_number) = False and len(case_number) = 8)
		transmit
		If isnumeric(case_number) = True then
			EMReadScreen MAXIS_check, 5, 1, 39
			If MAXIS_check <> "MAXIS" then MsgBox "You are not in MAXIS. Navigate your screen to MAXIS and try again. You might be passworded out."
		Else
			MMIS_row = 1
			MMIS_col = 1
			EMSearch "MMIS", MMIS_row, MMIS_col
			If MMIS_row <> 1 then
				EMReadScreen OSLT_check, 4, 1, 52 'Because cases that are on the "OSLT" screen in MMIS don't contain the characters "MMIS" in the top line.
				If OSLT_check = "OSLT" then MMIS_row = 1
			End if
			If MMIS_row <> 1 then MsgBox "You are not in MMIS. Navigate your screen to MMIS and try again. You might be passworded out."
		End if
	Loop until MAXIS_check = "MAXIS" or MMIS_row = 1
	If isnumeric(case_number) = True then
		call navigate_to_screen("case", "note")
		PF9
		EMReadScreen mode_check, 7, 20, 3
		If mode_check <> "Mode: A" and mode_check <> "Mode: E" then MsgBox "The script doesn't appear to be able to find your case note. Are you in inquiry? If so, navigate to production on the screen where you clicked the script button, and try again. Otherwise, you might have forgotten to type a valid case number."
	Else
		call MMIS_RKEY_finder
		EMWriteScreen "c", 2, 19
		EMWriteScreen case_number, 9, 19
		transmit
		EMReadScreen RKEY_check, 4, 1, 52 'CHECKING FOR RKEY, IF RENEWAL IS DUE A WARNING MESSAGE WILL NEED TO BE MOVED PAST.
		If RKEY_check = "RKEY" then transmit
		PF4
		PF11
		EMReadScreen MMIS_edit_mode_check, 5, 5, 2
		If MMIS_edit_mode_check <> "'''''" then script_end_procedure("MMIS edit mode not found. Are you in inquiry? Is MMIS not functioning? Shut down this script and try again. If it continues to not work, email your script administrator the case number, and a screenshot of MMIS.")
	End if
Loop until (mode_check = "Mode: A" or mode_check = "Mode: E") or (MMIS_edit_mode_check = "'''''") 

If isnumeric(case_number) = True then
	EMSendKey contact_type & " " & contact_direction & " " & who_contacted
	If regarding <> "" then EMSendKey " re: " & regarding 
	EMSendKey "<newline>"
	If when_voicemail_was_left <> "" then Call write_editbox_in_case_note("Contact made", when_voicemail_was_left, 6)
	If phone_number <> "" then Call write_editbox_in_case_note("Phone number", phone_number, 6)
	If contact_reason <> "" then Call write_editbox_in_case_note("Reason for Contact", contact_reason, 6)
	If actions_taken <> "" then Call write_editbox_in_case_note("Actions taken", actions_taken, 6)
	IF verifs_needed <> "" then Call write_editbox_in_case_note("Verifs Needed", verifs_needed, 6)
	If cl_instructions <> "" then Call write_editbox_in_case_note("Instructions/Message for CL", cl_instructions, 6)
	If case_status <> "" then Call write_editbox_in_case_note("Case status", case_status, 6)
      If caf_1_check = 1 then write_new_line_in_case_note ("* Reminded client about importance of submitting the CAF 1.")
	Call write_new_line_in_case_note("---")
	Call write_new_line_in_case_note(worker_signature)
      
	If TIKL_check = 0 then script_end_procedure("")

	'TIKLING
	MsgBox "The script will now navigate to a TIKL."
	call navigate_to_screen("dail", "writ")
Else
	EMSendKey contact_type & " " & contact_direction & " " & who_contacted
	If regarding <> "" then EMSendKey " re: " & regarding 
	PF11
	If when_voicemail_was_left <> "" then EMSendKey "* Contact made: " & when_voicemail_was_left
	PF11
	If phone_number <> "" then EMSendKey "* Phone number: " & phone_number
	PF11
	If issue <> "" then EMSendKey "* Reason for Contact: " & issue
	PF11
	If actions_taken <> "" then EMSendKey "* Actions taken: " & actions_taken
	PF11
	If cl_instructions <> "" then EMSendKey "* Instructions/Message for CL: " & cl_instructions
	PF11
	If verifs_needed <> "" then EMSendKey "* Verifs Needed: " & verifs_needed
	PF11
	EMSendKey "---"
	PF11
	EMSendKey worker_signature
	PF11
	EMSendKey "************************************************************************"
	If TIKL_check = 1 then script_end_procedure("Unable to TIKL for MCRE case. Find the MAXIS case and TIKL manually.")
End if

script_end_procedure("")
