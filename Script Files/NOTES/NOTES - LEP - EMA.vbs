'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - LEP - EMA.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
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
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'Required for statistical purposes==========================================================================================
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 270          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog EMA_dialog, 0, 0, 311, 305, "EMA "
  EditBox 85, 5, 75, 15, case_number
  EditBox 110, 30, 95, 15, date_received
  EditBox 55, 50, 45, 15, HH_COMP
  EditBox 55, 70, 115, 15, CIT_ID
  EditBox 75, 95, 80, 15, EMMA_Begin_date
  EditBox 75, 120, 80, 15, EMMA_End_Date
  DropListBox 75, 155, 125, 15, "SELECT ONE..."+chr(9)+"Healthy Jeopardy"+chr(9)+"Serious Impairment"+chr(9)+"Serious Dysfunction", CONSEQUENCE
  EditBox 80, 185, 195, 15, NOTES_ON_INCOME
  DropListBox 80, 220, 125, 15, "SELECT ONE..."+chr(9)+"APPROVED"+chr(9)+"DENIED"+chr(9)+"INCOMPLETE", ACTION_TAKEN
  EditBox 85, 250, 135, 15, Worker_signature
  ButtonGroup ButtonPressed
    OkButton 160, 285, 50, 15
    CancelButton 230, 285, 50, 15
  Text 5, 100, 65, 10, "EMMA Begin Date: "
  Text 5, 185, 70, 10, "NOTES ON INCOME:"
  Text 5, 55, 45, 10, "HH COMP: "
  Text 5, 155, 60, 10, "CONSEQUENCE:"
  Text 5, 220, 65, 10, "ACTION TAKEN:"
  Text 5, 10, 75, 10, "Maxis Case Number:"
  Text 5, 250, 75, 10, "Sign Your Case Note:"
  Text 5, 75, 40, 10, "CIT/ID: "
  Text 5, 125, 60, 15, "EMMA End Date: "
  Text 5, 35, 100, 10, "EMA MNSURE App Received: "
EndDialog



'script code-----------------------------------------------------------------------------------------------

'Connect to Bluezone
EMConnect ""

'Grabs Maxis Case number
CALL MAXIS_case_number_finder(case_number)

'Shows dialog
DO
	DO
	
		Dialog EMA_DIALOG
		IF ButtonPressed = 0 THEN StopScript
		IF worker_signature = "" THEN MsgBox "You must sign your case note!"
		LOOP UNTIL worker_signature <> ""
	IF IsNumeric(case_number) = FALSE THEN MsgBox "You must type a valid numeric case number."
LOOP UNTIL IsNumeric(case_number) = TRUE
	

'Checks Maxis for password prompt
CALL check_for_MAXIS(True)


'Navigates to case note
CALL navigate_to_screen("CASE", "NOTE")

'Sends a PF9
PF9

'Writes the case note
CALL write_variable_in_case_note ("***EMA***")													'writes title in case note
CALL write_bullet_and_variable_in_case_note("Ema mnsure app received", date_receive)						' writes the date application was received
CALL write_bullet_and_variable_in_case_note("hh comp", HH_comp)										' writes the number of people in HH
CALL write_bullet_and_variable_in_case_note("cit/id", cit_id)										' writes whether or no client is a citizen
CALL write_bullet_and_variable_in_case_note("emma begin date", emma_begin_date)							' writes the date the EMA began
CALL write_bullet_and_variable_in_case_note("emma end date", emma_end_date)
IF CONSEQUENCE <> "Select One..." THEN CALL write_bullet_and_variable_in_case_note("consequence", CONSEQUENCE)		' writes how EMA affects clients health
CALL write_bullet_and_variable_in_case_note("notes on income", notes_on_income)							' writes what type of income client has
IF ACTION_TAKEN <> "SELECT ONE..." THEN CALL write_bullet_and_variable_in_case_note("action taken", ACTION_TAKEN)		' writes outcome of application
CALL write_variable_in_case_note ("---")
CALL write_variable_in_case_note (worker_signature)




CALL script_end_procedure("")


