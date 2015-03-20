'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - MEDICAL OPINION FORM RECEIVED.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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


'Dialog---------------------------------------------------------------------------------------------------------------------------

BeginDialog MOF_recd, 0, 0, 186, 220, "Medical Opinion Form Received"
  EditBox 55, 5, 100, 15, case_number
  EditBox 55, 25, 95, 15, date_recd
  CheckBox 5, 45, 90, 10, "Client signed release?", client_release
  EditBox 90, 60, 85, 15, doctor_date
  EditBox 45, 80, 130, 15, diagnosis
  EditBox 70, 100, 105, 15, condition_will_last
  EditBox 85, 120, 90, 15, ability_to_work
  EditBox 50, 140, 125, 15, other_notes
  EditBox 50, 160, 125, 15, action_taken
  EditBox 70, 180, 105, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 70, 200, 50, 15
    CancelButton 125, 200, 50, 15
  Text 5, 10, 50, 10, "Case Number: "
  Text 5, 30, 50, 10, "Date received: "
  Text 5, 65, 80, 10, "Date doctor signed form: "
  Text 5, 85, 40, 10, "Diagnosis"
  Text 5, 105, 65, 10, "Condition will last:"
  Text 5, 125, 75, 10, "Client's ability to work: "
  Text 5, 165, 45, 10, "Action taken: "
  Text 5, 185, 60, 10, "Worker Signature: "
  Text 5, 145, 40, 10, "Other notes: "
EndDialog



'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------

EMConnect ""

'Finds the case number
CALL find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
IF IsNumeric(case_number) = False then case_number = ""

CALL check_for_MAXIS(True)


'calling the dialog---------------------------------------------------------------------------------------------------------------
DO
	DO
		Dialog MOF_recd
		IF buttonpressed = 0 THEN stopscript
		IF case_number = "" THEN MsgBox "You must have a case number to continue!"
		IF worker_signature = "" THEN MsgBox "You must enter a worker signature."
		IF overpayment_yn = "Select One..." THEN Msgbox "You must select an option for overpayment."
	LOOP until case_number <> "" and worker_signature <> ""
	CALL check_for_MAXIS(TRUE)
	CALL navigate_to_screen("case", "note")
	PF9
	EMReadscreen mode_check, 7, 20, 3
	IF mode_check <> "Mode: A" AND mode_check <> "Mode: E" THEN MsgBox "For some reason, the script can't get to a case note. Did you start the script in inquiry by mistake? Navigate to MAXIS production, or shut down the script and try again."
LOOP until mode_check = "Mode: A" OR mode_check = "Mode: E"



'The case note---------------------------------------------------------------------------------------------------------------------
CALL write_variable_in_CASE_NOTE("***Medical Opinion Form Rec'd " & date_recd & "***")
IF client_release = checked THEN CALL write_variable_in_CASE_NOTE ("* Client signed release on MOF.")
CALL write_bullet_and_variable_in_CASE_NOTE("Diagnosis", diagnosis)
CALL write_bullet_and_variable_in_CASE_NOTE("Condition will last", condition_will_last)
CALL write_bullet_and_variable_in_CASE_NOTE("Ability to work", ability_to_work)
CALL write_bullet_and_variable_in_CASE_NOTE("Doctor signed form", doctor_date)
CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
CALL write_bullet_and_variable_in_CASE_NOTE("Action taken", action_taken)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

Script_end_procedure("")

	





