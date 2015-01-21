'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - MFIP SANCTION AND DWP DISQUALIFICATION.vbs"
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

'DIALOGS-------------------------------
'MFIP Sanction/DWP Disqualification Dialog Box
BeginDialog MFIP_Sanction_DWP_Disq_Dialog, 0, 0, 296, 285, "MFIP Sanction - DWP Disqualification"
  Text 5, 10, 70, 10, "MAXIS Case Number:"
  EditBox 75, 5, 90, 15, case_number
  Text 170, 10, 80, 10, "HH Member's Number:"
  EditBox 255, 5, 35, 15, HH_Member_Number
  Text 5, 30, 95, 10, "Type of Sanction (ES or CS):"
  EditBox 110, 25, 85, 15, Type_Sanction
  Text 5, 50, 140, 10, "Effective Date of Sanction/Disqualification:"
  EditBox 150, 45, 90, 15, Date_Sanction
  Text 5, 70, 75, 10, "Number of occurences:"
  EditBox 90, 65, 30, 15, Number_Occurrences
  Text 155, 70, 70, 10, "Sanction Percentage:"
  DropListBox 230, 65, 60, 15, "Select One:"+chr(9)+"10%"+chr(9)+"30%"+chr(9)+"100%", Sanction_Percentage
  Text 5, 90, 145, 10, "Sanction information received from and how:"
  EditBox 150, 85, 140, 15, Sanction_Notification_Received
  Text 5, 110, 170, 10, "Last day to cure (10 day cutoff or last day of month):"
  EditBox 180, 105, 70, 15, Last_Day_Cure
  Text 5, 130, 80, 10, "Reason for the sanction:"
  EditBox 90, 125, 200, 15, Reason_for_Sanction
  Text 5, 150, 85, 10, "Impact to other programs:"
  EditBox 95, 145, 195, 15, Impact_Other_Programs
  Text 5, 165, 145, 20, "Communicated with client to cure santion by sending (i.e., SPEC/MEMO):"
  EditBox 155, 165, 135, 15, Memo_to_Client
  Text 5, 190, 105, 25, "Vendor information (if vendoring due to the sanction, vendor #, etc.):"
  EditBox 115, 190, 175, 15, Vendor_Information
  CheckBox 5, 220, 225, 10, "Check here if you sent a status update to Employment Services.", Update_Sent_ES_Checkbox
  CheckBox 5, 235, 220, 10, "Check here if you sent a status update to Child Care Assistance.", Update_Sent_CCA_Checkbox
  CheckBox 5, 250, 125, 10, "Check here if you FIATed this case.", Fiat_Checkbox
  Text 5, 270, 70, 10, "Sign your case note:"
  EditBox 80, 265, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 185, 265, 50, 15
    CancelButton 240, 265, 50, 15
EndDialog




'THE SCRIPT--------------------------------------------

'Connects to BlueZone
EMConnect ""

'Asks for Case Number
CALL MAXIS_case_number_finder(case_number)

'Shows dialog
DO
	DO
		DO
			DO
				DO
					DO
						Dialog MFIP_Sanction_DWP_Disq_Dialog
						IF ButtonPressed = 0 THEN StopScript
						IF worker_signature = "" THEN MsgBox "You must sign your case note!"
					LOOP UNTIL worker_signature <> ""
					IF IsNumeric(case_number) = FALSE THEN MsgBox "You must type a valid numeric case number!"
				LOOP UNTIL IsNumeric(case_number) = TRUE
				IF IsDate(Date_Sanction) = FALSE THEN MsgBox "You must type a valid date of sanction!"
			LOOP UNTIL IsDate(Date_Sanction) = TRUE
			IF Sanction_Percentage = "Select One:" THEN MsgBox "You must select a sanction percentage!"
		LOOP UNTIL Sanction_Percentage <> "Select One:"
		IF HH_Member_Number = "" THEN MsgBox "You must enter a HH member number!"
	LOOP UNTIL HH_Member_Number <> ""
	IF Type_Sanction = "" THEN MsgBox "You must enter a sanction type!"
LOOP UNTIL Type_Sanction <> ""
'Checks MAXIS for password prompt
MAXIS_check_function

'Navigates to case note
CALL navigate_to_screen("CASE", "NOTE")

'Send PF9 to case note
PF9


'Writes case note
EMSENDKEY "***" & Sanction_Percentage & " " & ucase(Type_Sanction) & " SANCTION " & "MEMBER " & HH_Member_Number & " EFF " & Date_Sanction & "***" & "<NEWLINE>"

IF HH_Member_Number <> "" THEN CALL write_bullet_and_variable_in_case_note("HH Member's Number", HH_Member_Number)
IF Type_Sanction <> "" THEN CALL write_bullet_and_variable_in_case_note("Type of Sanction", Type_Sanction)
IF Date_Sanction <> "" THEN CALL write_bullet_and_variable_in_case_note("Effective date of sanction/disqualification", Date_Sanction)
IF Number_Occurrences <> "" THEN CALL write_bullet_and_variable_in_case_note("Number of occurences", Number_Occurrences)
IF Sanction_Percentage <> "" THEN CALL write_bullet_and_variable_in_case_note("Sanction Percent is", Sanction_Percentage)
IF Sanction_Notification_Received<> "" THEN CALL write_bullet_and_variable_in_case_note("Sanction information received from", Sanction_Notification_Received)
IF Last_Day_Cure <> "" THEN CALL write_bullet_and_variable_in_case_note("Last day to cure", Last_Day_Cure)
IF Reason_for_Sanction <> "" THEN CALL write_bullet_and_variable_in_case_note ("Reason for the sanction", Reason_for_Sanction)
IF Impact_Other_Programs <> "" THEN CALL write_bullet_and_variable_in_case_note ("Impact to other programs", Impact_Other_Programs)
IF Memo_to_Client <> "" THEN CALL write_bullet_and_variable_in_case_note ("Communicated with client to cure sanction by sending", Memo_to_Client)
IF Vendor_Information <> "" THEN CALL write_bullet_and_variable_in_case_note("Vendoring information", Vendor_Information)
IF Update_Sent_ES_Checkbox = 1 THEN CALL write_variable_in_case_note("* Status update information was sent to Employment Services.")
IF Update_Sent_CCA_Checkbox = 1 THEN CALL write_variable_in_case_note("* Status update information was sent to Child Care Assistance.")
IF FIAT_Checkbox = 1 THEN CALL write_variable_in_case_note("* Case has been FIATed.")

'case note worker signature
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

script_end_procedure("")


