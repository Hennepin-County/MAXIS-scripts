'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - DOCUMENTS RECEIVED.vbs"
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

'DIALOGS----------------------------------------------------------------------------------------------------

BeginDialog docs_received_dialog, 0, 0, 466, 140, "Docs received"
  EditBox 55, 5, 90, 15, case_number
  EditBox 215, 5, 65, 15, document_datestamp
  EditBox 60, 25, 215, 15, docs_received
  EditBox 75, 45, 390, 15, verif_notes
  EditBox 60, 65, 405, 15, actions_taken
  EditBox 70, 85, 110, 15, worker_signature
  'CheckBox 195, 85, 205, 10, "Check here to start the approved programs script after this.", approved_progs_check
  'CheckBox 195, 95, 200, 10, "Check here to start the closed programs script after this. ", closed_progs_check
  'CheckBox 195, 105, 195, 10, "Check here to start the denied programs script after this.", denied_progs_check
  EditBox 115, 120, 350, 15, docs_needed
  ButtonGroup ButtonPressed
    OkButton 355, 5, 50, 15
    CancelButton 410, 5, 50, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 160, 10, 55, 10, "Doc datestamp:"
  Text 5, 30, 55, 10, "Docs received:"
  Text 5, 50, 70, 10, "Notes on your docs:"
  Text 280, 30, 190, 10, "Note: just list the docs here. This is the title of your note."
  Text 5, 90, 65, 10, "Worker signature:"
  Text 5, 70, 50, 10, "Actions taken: "
  Text 5, 125, 110, 10, "Verifs still needed (if applicable):"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------

EMConnect ""

'Finds the case number
call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""

'Displays the dialog and navigates to case note
Do
	Do
		Do
			Dialog docs_received_dialog
			If buttonpressed = 0 then stopscript
			If case_number = "" then MsgBox "You must have a case number to continue!"
			Loop until case_number <> ""
		transmit
		EMReadScreen MAXIS_check, 5, 1, 39
		If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You appear to be locked out of MAXIS. Are you passworded out? Did you navigate away from MAXIS?"
	Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "
	call navigate_to_screen("case", "note")
	PF9
	EMReadScreen mode_check, 7, 20, 3
	If mode_check <> "Mode: A" and mode_check <> "Mode: E" then MsgBox "For some reason, the script can't get to a case note. Did you start the script in inquiry by mistake? Navigate to MAXIS production, or shut down the script and try again."
Loop until mode_check = "Mode: A" or mode_check = "Mode: E"

'Case notes
EMSendKey "Docs rec'd: "
call write_new_line_in_case_note(docs_received)
If document_datestamp <> "" then call write_editbox_in_case_note("Document datestamp", document_datestamp, 6)
If verif_notes <> "" then call write_editbox_in_case_note("Notes", verif_notes, 6) 
call write_editbox_in_case_note("Actions taken", actions_taken, 6) 
If docs_needed <> "" then call write_editbox_in_case_note("Verifs needed", docs_needed, 6) 
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

'Runs approved  progs if selected
If approved_progs_check = 1 then run_another_script("C:\DHS-MAXIS-Scripts\Script Files\NOTE - Approved Programs.vbs")

'Runs denied progs if selected
If closed_progs_check = 1 then run_another_script("C:\DHS-MAXIS-Scripts\Script Files\NOTE - closed progs.vbs")

'Runs denied progs if selected
If denied_progs_check = 1 then run_another_script("C:\DHS-MAXIS-Scripts\Script Files\NOTE - denied progs.vbs")

script_end_procedure("")


