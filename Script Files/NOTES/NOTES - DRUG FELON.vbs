'Option Explicit

'DIM beta_agency
'DIM url, req, fso

'LOADING ROUTINE FUNCTIONS---------------------------------------------------------------
If beta_agency = "" or beta_agency = True then
	url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
Else
	url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
End if
SET req = CreateObject("Msxml2.XMLHttp.6.0")		'Creates an object to get a URL
req.open "GET", url, FALSE				'Attempts to open the URL
req.send					'Sends request
IF req.Status = 200 THEN				'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText				'Executes the script code
ELSE  'Error message tells user to try github.com, otherwise contact Veronica with details (and stops script).
	MsgBox "Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
	 vbCr & _
	 "Before contacting DHS, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
	 vbCr & _
	 "If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact " &_
	 "Veronica Cary and provide the following information:" & vbCr &_
	 vbTab & "- The name of the script you are running." & vbCr &_
	 vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
	 vbTab & "- The name and email for an employee from your IT department," & vbCr & _
	 vbTab & vbTab & "responsible for network issues." & vbCr &_
	 vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
	 vbCr & _
	 "Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
	 vbCr &_
	 "URL: " & url
	 stopscript
END IF


'DIMMING VARIABLES
'DIM case_number, conviction_date, probation_officer, authorization_on_file_check, complied_with_UA_check, UA_date, date_of_1st_offense, date_of_2nd_offense, worker_signature, ButtonPressed,drug_felon_dialog, UA_results, Maxis_drug_function, po_officer, Authorization_on_file, Complying_with_PO, actions_taken


'DIALOGS-------------------------------------------------------------------------------------------------------------------------------
BeginDialog drug_felon_dialog, 0, 0, 246, 235, "Drug Felon"
  EditBox 60, 5, 75, 15, case_number
  EditBox 65, 25, 60, 15, conviction_date
  EditBox 65, 45, 135, 15, probation_officer
  CheckBox 10, 65, 145, 10, "Check here if the authorization is on file:", authorization_on_file_check
  CheckBox 10, 80, 130, 10, "Check here if client complied with UA:", complied_with_UA_check
  EditBox 40, 95, 80, 15, UA_date
  DropListBox 50, 115, 65, 15, "select one..."+chr(9)+"Positive"+chr(9)+"Negative"+chr(9)+"Refused", UA_results
  EditBox 75, 135, 55, 15, date_of_1st_offense
  EditBox 75, 155, 70, 15, date_of_2nd_offense
  EditBox 60, 175, 180, 15, actions_taken
  EditBox 80, 195, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 135, 215, 50, 15
    CancelButton 190, 215, 50, 15
  Text 5, 160, 70, 10, "Date of 2nd Offense:"
  Text 5, 30, 55, 10, "Conviction Date:"
  Text 5, 180, 50, 10, "Actions Taken:"
  Text 5, 120, 40, 10, "UA Results:"
  Text 5, 200, 70, 15, "Sign your Case Note:"
  Text 5, 140, 65, 10, "Date of 1st Offense:"
  Text 5, 50, 60, 10, "Probation Officer:"
  Text 5, 100, 30, 10, "UA Date:"
  Text 5, 10, 50, 10, "Case Number:"
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect "" 

'Grabs the MAXIS case number
CALL MAXIS_case_number_finder(case_number)

'Show dialog
DO
	DO
		DO
			Dialog drug_felon_dialog
			IF Buttonpressed = 0 THEN StopScript
			IF worker_signature = "" THEN MsgBox "You must sign your case note"
		LOOP UNTIL worker_signature <> ""
		IF IsNumeric(case_number)= FALSE THEN MsgBox "You must type a valid numeric case number."
	LOOP UNTIL IsNumeric(case_number) = TRUE
	If UA_results = "select one..." THEN MsgBox "You must select 'UA results field'"
LOOP UNTIL UA_results <> "select one..."

'Checks MAXIS for password prompt

'Navigate to case note
CALL navigate_to_screen("CASE", "NOTE")



'Sends a PF9
PF9

'Writes the case note
CALL write_variable_in_case_note("***Drug Felon***")
CALL write_bullet_and_variable_in_case_note("Conviction date", conviction_date)
CALL write_bullet_and_variable_in_case_note("Probation Officer", po_officer)
IF authorization_on_file_check = checked THEN CALL write_variable_in_case_note("* Authorization on file.")
IF complied_with_UA_check = checked THEN CALL write_variable_in_case_note("* Complied with UA.")
CALL write_bullet_and_variable_in_case_note("UA Date", UA_date)
CALL write_bullet_and_variable_in_case_note("Date of 1st offence", date_of_1st_offense)
CALL write_bullet_and_variable_in_case_note("Date of 2nd offence", date_of_2nd_offense)
IF UA_results <> "select one..." THEN CALL write_bullet_and_variable_in_case_note("UA results", UA_results)
CALL write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

CALL script_end_procedure("")