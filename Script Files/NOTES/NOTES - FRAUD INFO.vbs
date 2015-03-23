'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - FRAUD INFO.vbs"
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

'Option Explicit
'DIM case_number
'DIM referral_date
'DIM referral_reason
'DIM fraud_findings
'DIM actions_taken
'DIM yes_overpayment
'DIM no_overpayment
'DIM worker_signature

'Dialog---------------------------------------------------------------------------------------------------------------------------
BeginDialog Fraud_Dialog, 0, 0, 211, 245, "Fraud Info"
  EditBox 65, 10, 90, 15, case_number
  EditBox 75, 30, 115, 15, referral_date
  EditBox 10, 65, 195, 15, referral_reason
  EditBox 10, 100, 195, 15, fraud_findings
  EditBox 10, 135, 195, 15, actions_taken
  DropListBox 10, 170, 55, 15, "Select One..."+chr(9)+"Yes"+chr(9)+"No", overpayment_yn
  EditBox 95, 200, 95, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 95, 220, 50, 15
    CancelButton 150, 220, 50, 15
  Text 10, 15, 55, 10, "Case Number: "
  Text 10, 35, 65, 10, "Date referral made:"
  Text 10, 50, 110, 10, "Reason for referral (be specific):"
  Text 10, 85, 55, 10, "Fraud findings:"
  Text 10, 120, 50, 10, "Actions taken:"
  Text 10, 155, 50, 10, "Overpayment?"
  Text 90, 155, 90, 35, "If yes for overpayment please use overpayment script to case note the specific details regarding it. "
  Text 30, 205, 60, 10, "Worker Signature: "
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------

EMConnect ""

'Finds the case number
CALL MAXIS_case_number_finder(case_number)

CALL check_for_MAXIS(True)


'calling the dialog---------------------------------------------------------------------------------------------------------------
DO
	DO
		Dialog fraud_dialog
		IF buttonpressed = 0 THEN stopscript
		IF case_number = "" THEN MsgBox "You must have a case number to continue!"
		IF worker_signature = "" THEN MsgBox "You must enter a worker signature."
		IF overpayment_yn = "Select One..." THEN Msgbox "You must select an option for overpayment."
	LOOP until case_number <> "" and worker_signature <> "" and (overpayment_yn = "Yes" or overpayment_yn ="No")
	CALL check_for_MAXIS(TRUE)
	CALL navigate_to_screen("case", "note")
	PF9
	EMReadscreen mode_check, 7, 20, 3
	IF mode_check <> "Mode: A" AND mode_check <> "Mode: E" THEN MsgBox "For some reason, the script can't get to a case note. Did you start the script in inquiry by mistake? Navigate to MAXIS production, or shut down the script and try again."
LOOP until mode_check = "Mode: A" OR mode_check = "Mode: E"

'debatable to include?
IF overpayment_yn = "Yes" THEN overpayment_yn = " Yes. See overpayment case note for more details."


'The case note---------------------------------------------------------------------------------------------------------------------
CALL write_variable_in_CASE_NOTE("***Fraud Referral Info***")
CALL write_bullet_and_variable_in_CASE_NOTE("Referral Date", referral_date)
CALL write_bullet_and_variable_in_CASE_NOTE("Referral Reason", referral_reason)
CALL write_bullet_and_variable_in_CASE_NOTE("Findings", fraud_findings)
CALL write_bullet_and_variable_in_CASE_NOTE("Actions Taken", actions_taken)
CALL write_bullet_and_variable_in_CASE_NOTE("Overpayment?", overpayment_yn)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

Script_end_procedure("")

	





