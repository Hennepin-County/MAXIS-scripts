'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - MILEAGE REIMBURSEMENT REQUEST.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER FUNCTIONS LIBRARY.vbs"
Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, False									'Attempts to open the URL
req.send													'Sends request
IF req.Status = 200 Then									'200 means great success
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
BeginDialog mileage_dialog, 0, 0, 306, 125, "Mileage Reimbursement"
  EditBox 75, 5, 70, 15, case_number
  EditBox 230, 5, 70, 15, date_docs_recd
  EditBox 50, 25, 70, 15, total_reimbursement
  EditBox 230, 25, 70, 15, date_to_accounting
  EditBox 50, 45, 250, 15, docs_reqd
  EditBox 50, 65, 250, 15, other_notes
  EditBox 55, 85, 245, 15, actions_taken
  EditBox 70, 105, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 195, 105, 50, 15
    CancelButton 250, 105, 50, 15
  Text 5, 10, 70, 10, "MAXIS case number:"
  Text 170, 10, 55, 10, "Date Received:"
  Text 5, 30, 45, 10, "Total Amount:"
  Text 165, 30, 60, 10, "Date Sent to Acct:"
  Text 5, 50, 40, 10, "Doc's req'd:"
  Text 5, 70, 45, 10, "Other notes:"
  Text 5, 90, 50, 10, "Actions taken:"
  Text 5, 110, 60, 10, "Worker signature:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
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
			Dialog Mileage_dialog
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
EMSendKey ">>>>>MILEAGE REIMBURSEMENT REQUEST - ACTIONS TAKEN<<<<<" & "<newline>"
If date_docs_recd <> "" then call write_editbox_in_case_note("Date received", date_docs_recd, 6)
If total_reimbursement <> "" then call write_editbox_in_case_note("Total Amount", "$" & total_reimbursement, 6)
If date_to_accounting <> "" then call write_editbox_in_case_note("Date Sent to Accounting", date_to_accounting, 6)
If docs_reqd <> "" then call write_editbox_in_case_note("Docs requested", docs_reqd, 6)
If other_notes <> "" then call write_editbox_in_case_note("Other notes", other_notes, 6)
If actions_taken <> "" then call write_editbox_in_case_note("Actions taken", actions_taken, 6)
If worker_county_code = "x179" then call write_new_line_in_case_note("* Please note: DO NOT SCAN!! Accounting will scan into OnBase when processed.", 6)	'Should only do this for Wabasha County, unless other counties request it.
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

script_end_procedure("")