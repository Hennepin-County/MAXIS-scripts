'STATS GATHERING-----------------------------------------------------------------------------------
name_of_script = "NOTES - DWP BUDGET.vbs"
start_time = timer

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

'This is the dialog box information/code
BeginDialog DWP_budget_dialog, 0, 0, 426, 165, "DWP Budget Dialog"
  EditBox 60, 5, 45, 15, case_number
  EditBox 195, 5, 45, 15, ES_appointment_date
  EditBox 370, 5, 45, 15, ES_deadline_date
  EditBox 55, 25, 365, 15, income_info
  EditBox 55, 45, 365, 15, shelter_info
  EditBox 165, 65, 15, 15, personal_needs
  EditBox 75, 85, 160, 15, vendor_information
  EditBox 50, 105, 230, 15, other_notes
  EditBox 60, 125, 120, 15, months_eligible
  EditBox 255, 125, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 315, 145, 50, 15
    CancelButton 370, 145, 50, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 120, 10, 75, 10, "ES Appointment Date:  "
  Text 255, 10, 115, 10, "ES Deadline (10 Business Days):"
  Text 5, 30, 45, 10, "Income Info:"
  Text 5, 50, 45, 10, "Shelter Info: "
  Text 5, 70, 165, 10, "Personal Needs (Number of DWP HH Members):"
  Text 190, 65, 230, 20, "(This will multiply the number of eligible DWP household members by $70.00/person.)"
  Text 5, 90, 70, 10, "Vendor Information: "
  Text 5, 110, 45, 10, "Other Notes: "
  Text 5, 130, 60, 10, "Months Eligible: "
  Text 195, 130, 65, 10, "Worker Signature:"
EndDialog


'Connecting to BlueZone

'THE SCRIPT----------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""

'Finds the case number
CALL MAXIS_case_number_finder(case_number)

'Displays the dialog
DO
	DO
		Dialog DWP_budget_dialog
		IF buttonpressed = cancel THEN stopscript
		IF case_number = "" THEN MsgBox "You must have a case number to continue!"
	LOOP UNTIL case_number <> ""
	IF worker_signature = "" THEN MsgBox "You must sign your case note!"
LOOP UNTIL worker_signature <> ""

'Calculates personal needs info
personal_needs = "$" & personal_needs * 70


'Checks to make sure worker is not passworded out
CALL check_for_MAXIS(False)

'Navigates to CASE/NOTE
CALL navigate_to_screen("CASE","NOTE")

'Adding new blank case note
PF9

'Writing to CASE/NOTE
CALL write_variable_in_case_note("***DWP ES Referral and Budget Info***")
IF ES_appointment_date <> "" THEN CALL write_bullet_and_variable_in_case_note("ES Appointment Date", ES_appointment_date)
IF ES_deadline_date <> "" THEN CALL write_bullet_and_variable_in_case_note("ES Deadline Date", ES_deadline_date)
IF income_info <> "" THEN CALL write_bullet_and_variable_in_case_note("Income Info", income_info)
IF shelter_info <> "" THEN CALL write_bullet_and_variable_in_case_note("Shelter Info", shelter_info)
IF personal_needs <> "" THEN CALL write_bullet_and_variable_in_case_note("Personal Needs", personal_needs)
IF vendor_information <> "" THEN CALL write_bullet_and_variable_in_case_note("Vendor Information", vendor_information)
IF other_notes <> "" THEN CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
IF months_eligible <> "" THEN CALL write_bullet_and_variable_in_case_note("Months Eligible", months_eligible)
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

'End script
CALL script_end_procedure("")














