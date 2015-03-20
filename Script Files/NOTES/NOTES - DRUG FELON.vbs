Option Explicit

DIM beta_agency
DIM url, req, fso

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
If beta_agency = "" or beta_agency = True then
	url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
Else
	url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
End if
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


'DIMMING VARIABLES
DIM case_number, conviction_date, probation_officer, authorization_on_file_check, complied_with_UA_check, UA_date, date_of_1st_offense, date_of_2nd_offense, worker_signature, ButtonPressed,drug_felon_dialog, UA_results, case_number_finder, Maxis_case_number, Maxis_drug_function, navigate_to_MAXIS_screen, po_officer, Authorization_on_file, Complying_with_PO, actions_taken


'DIALOGS-------------------------------------------------------------------------------------------------------------------------------

BeginDialog drug_felon_dialog, 0, 0, 246, 230, "Drug Felon"
  EditBox 60, 5, 75, 15, case_number
  EditBox 60, 25, 60, 15, conviction_date
  EditBox 65, 40, 135, 15, probation_officer
  CheckBox 10, 65, 145, 10, "Check here if the authorization is on file:", authorization_on_file_check
  CheckBox 10, 80, 130, 10, "Check here if client complied with UA:", complied_with_UA_check
  EditBox 35, 90, 80, 15, UA_date
  DropListBox 45, 110, 65, 15, "select one..."+chr(9)+"Positive"+chr(9)+"Negative"+chr(9)+"Refused", UA_results
  EditBox 70, 130, 55, 15, date_of_1st_offense
  EditBox 75, 150, 70, 15, date_of_2nd_offense
  EditBox 55, 170, 185, 15, actions_taken
  EditBox 75, 190, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 130, 210, 50, 15
    CancelButton 185, 210, 50, 15
  Text 5, 155, 70, 10, "Date of 2nd Offense:"
  Text 5, 25, 55, 15, "Conviction Date:"
  Text 5, 175, 50, 10, "Actions Taken:"
  Text 5, 115, 40, 10, "UA Results:"
  Text 5, 195, 70, 15, "Sign your Case Note:"
  Text 5, 135, 65, 10, "Date of 1st Offense:"
  Text 5, 45, 60, 10, "Probation Officer:"
  Text 5, 95, 45, 10, "UA Date:"
  Text 5, 10, 50, 10, "Case Number:"
EndDialog



'THE SCRIPT----------------------------------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect "" 

'Grabs MAXIS case number

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

'MsgBox checked

'Navigate to case note
CALL navigate_to_screen("CASE", "NOTE")



'Sends a PF9
PF9

'Writes the case note
CALL write_new_line_in_case_note("***Drug Felon***")
IF conviction_date <> "" THEN CALL write_bullet_and_variable_in_case_note("conviction date", conviction_date)
IF probation_officer <> "" THEN CALL write_bullet_and_variable_in_case_note("probation officer", po_officer)
IF authorization_on_file_check = 1 THEN CALL write_new_line_in_case_note("* authorization on file")
IF complied_with_UA_check = 1 THEN CALL write_new_line_in_case_note("* complied with UA")
IF UA_date <> "" THEN CALL write_bullet_and_variable_in_case_note("UA Date", UA_date)
IF date_of_1st_offense <> "" THEN CALL write_bullet_and_variable_in_case_note("date of 1st offense", date_of_1st_offense)
IF date_of_2nd_offense <> "" THEN CALL write_bullet_and_variable_in_case_note("date of 2nd offense", date_of_2nd_offense)
IF UA_results <> "select one..." THEN CALL write_bullet_and_variable_in_case_note("UA results", UA_results)
IF actions_taken <> "" THEN CALL write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
IF authorization_on_file_check = 1 THEN CALL write_bullet_and_variable_in_case_note("Authorization on file", Authorization_on_file)
IF complied_with_UA_check = 1 THEN CALL write_bullet_and_variable_in_case_note("Complying with PO", complying_with_PO)

CALL write_new_line_in_case_note("---")
CALL write_new_line_in_case_note(worker_signature)
