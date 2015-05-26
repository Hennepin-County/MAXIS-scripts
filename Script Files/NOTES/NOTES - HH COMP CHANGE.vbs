'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - HH COMP CHANGE.vbs"
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




'DIM variables from dialog; you can include a space and underscore return to start a new line and DIM will read it otherwise it has to be all on one line
DIM HH_Comp_Change_Dialog, Case_Number, HH_Member, Date_Reported, Effective_Date, Temporary_Change_Checkbox, Action_Taken, Additional_Notes, Worker_signature, Worker_name, Baby_Born_Checkbox, ButtonPressed
  

'dialog box for HH comp change
BeginDialog HH_Comp_Change_Dialog, 0, 0, 301, 205, "Household Comp Change"
  Text 5, 10, 50, 15, "Case Number"
  EditBox 60, 10, 100, 15, Case_Number
  Text 5, 30, 80, 15, "Unit Member HH Change"
  EditBox 90, 30, 45, 15, HH_Member
  Text 5, 55, 85, 15, "Date Reported/Addendum"
  EditBox 95, 50, 60, 15, Date_Reported
  Text 165, 55, 55, 15, "Effective Date"
  EditBox 215, 50, 70, 15, Effective_Date
  CheckBox 75, 75, 100, 15, "Is the change temporary?", Temporary_Change_Checkbox
  Text 5, 105, 50, 15, "Action Taken"
  EditBox 55, 100, 230, 15, Action_Taken
  Text 5, 130, 60, 15, "Additional Notes"
  EditBox 65, 125, 220, 15, Additional_Notes
  CheckBox 70, 150, 130, 15, "Is client reporting the birth of a baby?", Baby_Born_Checkbox
  Text 5, 180, 50, 15, "Worker Name"
  EditBox 60, 180, 100, 15, Worker_signature
  ButtonGroup ButtonPressed
    OkButton 190, 185, 50, 15
    CancelButton 245, 185, 50, 15
EndDialog


'Baby born dialog
DIM Baby_Born_Dialog, Baby_Name, Date_Birth, Baby_SSN, Father_Name, Father_HH_Checkbox, Father_Employer, Mother_Employer, OHI_Checkbox, OHI_Source, MAXIS_case_number_finder


'dialog box for baby born
BeginDialog Baby_Born_Dialog, 0, 0, 311, 225, "Client Reports Birth of Baby"
  Text 5, 10, 50, 15, "Case Number"
  EditBox 60, 10, 100, 15, Case_Number
  Text 5, 30, 55, 15, "Name of baby"
  EditBox 60, 30, 150, 15, Baby_Name
  Text 5, 55, 50, 15, "Date of Birth"
  EditBox 55, 55, 60, 15, Date_Birth
  Text 130, 55, 55, 15, "Baby's SSN"
  EditBox 180, 55, 85, 15, Baby_SSN
  CheckBox 5, 75, 100, 15, "Father is in the household.", Father_HH_Checkbox
  Text 5, 105, 50, 15, "Father's Name"
  EditBox 60, 100, 150, 15, Father_Name
  Text 5, 130, 70, 15, "Father's Employer"
  EditBox 75, 125, 225, 15, Father_Employer
  Text 5, 155, 70, 15, "Mother's Employer"
  EditBox 75, 150, 225, 15, Mother_Employer
  CheckBox 5, 170, 80, 15, "The baby has OHI", OHI_Checkbox
  Text 130, 175, 55, 15, "Source of OHI"
  EditBox 185, 170, 115, 15, OHI_Source
  
  ButtonGroup ButtonPressed
    OkButton 190, 205, 50, 15
    CancelButton 245, 205, 50, 15
EndDialog





'Connecting to BlueZone
EMConnect ""

'autofill case number function 
'CALL MAXIS_case_number_finder(case_number)



'run dialog - need to do the DO loops right after running dialog
DO
	DO
		DO
			DO
				DO
					DIALOG HH_Comp_Change_Dialog
					IF ButtonPressed = 0 THEN StopScript

					IF Case_Number = "" THEN MsgBox "You must enter case number!"
				LOOP UNTIL Case_Number <> ""
					IF HH_Member = "" THEN MsgBox "You must enter a HH Member"
			LOOP UNTIL HH_Member <> ""
			IF Date_Reported = "" THEN MsgBox "You must enter date reported"
		LOOP UNTIL Date_Reported <> ""
		IF Effective_Date = "" THEN MsgBox "You must enter effective date"
	LOOP UNTIL Effective_Date <> ""
	IF worker_signature = "" THEN MsgBox "Please sign your note"
LOOP UNTIL worker_signature <> ""


IF Baby_Born_Checkbox = 1 THEN 

'Do loop for Baby Born Dialogbox
DO
	DO
		DO
			DO
				
					DIALOG Baby_Born_Dialog
					IF ButtonPressed = 0 THEN StopScript

					IF Case_Number = "" THEN MsgBox "You must enter case number!"
				LOOP UNTIL Case_Number <> ""
					IF Date_Birth = "" THEN MsgBox "You must enter a birth date"
			LOOP UNTIL Date_Birth <> ""
			IF Baby_SSN = "" THEN MsgBox "You must enter baby's Social Security Number"
		LOOP UNTIL Baby_SSN <> ""
		IF Father_Name = "" THEN MsgBox "You must enter Father's name"
	LOOP UNTIL Father_Name <> ""
	

END IF



'Checks MAXIS for password prompt
CALL check_for_MAXIS(true)

'Navigates to case note
CALL navigate_to_MAXIS_screen("CASE", "NOTE")

'Send PF9 to case note
PF9




'Writes case note
CALL write_variable_in_case_note("HH Comp Change Reported")
CALL write_bullet_and_variable_in_Case_Note("Unit member HH Member: ", HH_Member)
CALL write_bullet_and_variable_in_Case_Note("Date Reported/Addendum: ", Date_Reported)
CALL write_bullet_and_variable_in_Case_Note("Date effective: ", Effective_Date)
CALL write_bullet_and_variable_in_Case_Note("Action Taken: ", Action_Taken)
CALL write_bullet_and_variable_in_Case_Note("Additional Notes: ", Additional_Notes)

'checkboxes
IF Temporary_Change_Checkbox = 1 THEN CALL write_variable_in_Case_Note("* Change is temporary")

'writes case note for baby born
IF Baby_Born_Checkbox = 1 THEN

	CALL write_variable_in_Case_Note("**Client reports birth of baby**")
	CALL write_bullet_and_variable_in_Case_Note("Baby's name: ", Baby_Name)	
	CALL write_bullet_and_variable_in_Case_Note("Date of birth: ", Date_birth)
	CALL write_bullet_and_variable_in_Case_Note("Baby's SSN: ", Baby_SSN)
	CALL write_bullet_and_variable_in_Case_Note("Father's name: ", Father_Name)
	CALL write_bullet_and_variable_in_Case_Note("Father's employer: ", Father_Employer)
	CALL write_bullet_and_variable_in_Case_Note("Mother's employer: ", Mother_Employer)
	IF OHI_Checkbox = 1 THEN CALL write_bullet_and_variable_in_Case_Note("* OHI: ", OHI_Source)
END IF

'signs case note
CALL write_variable_in_Case_Note("----" & worker_signature)

Script_End_Procedure ""





