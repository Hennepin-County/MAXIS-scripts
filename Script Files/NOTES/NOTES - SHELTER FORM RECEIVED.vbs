'Option Explicit
name_of_script = "NOTES - SHELTER FORM RECEIVED.vbs"
start_time = timer

'DIM beta_agency

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
STATS_manualtime = 90           'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'DIM shelter_form_received_dialog		'declaring variables that are being used in the rest of the script
'DIM date_moved_in
'DIM ButtonPressed
'DIM new_address
'DIM cost_per_person
'DIM how_many_residents
'DIM total_cost
'DIM other_notes
'DIM worker_signature
'DIM utilities_paid_by_resident_listbox
'DIM phonenumber
'DIM subsidized_amount
'DIM garage_amount_check
'DIM garage_amount
'DIM signed_by_LLMgr_check
'DIM signed_by_client_check
'DIM case_number
'DIM month
'DIM year
'DIM case_number_dialogbox
'DIM subsidized_amount_check

'THE DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 146, 70, "Case number dialog"
  EditBox 80, 5, 60, 15, case_number
  EditBox 80, 25, 25, 15, MAXIS_footer_month
  EditBox 115, 25, 25, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 35, 45, 50, 15
    CancelButton 90, 45, 50, 15
  Text 10, 30, 65, 10, "Footer month/year:"
  Text 10, 10, 45, 10, "Case number: "
EndDialog


BeginDialog Shelter_form_received_dialog, 0, 0, 206, 225, "Dialog"							'Dialogue box completed by worker with information provided by the client regarding the shelter form that was received
  EditBox 50, 5, 55, 15, date_moved_in
  EditBox 155, 5, 55, 15, how_many_residents
  EditBox 50, 25, 150, 15, new_address
  EditBox 50, 45, 100, 15, phone_number
  EditBox 50, 65, 45, 15, total_cost
  EditBox 150, 65, 50, 15, cost_per_person
  checkbox 10, 90, 75, 10, "Subsidized amount", subsidized_amount_check
  EditBox 95, 90, 40, 15, subsidized_amount
  checkbox 10, 105, 75, 10, "Garage amount", garage_amount_check
  EditBox 95, 105, 40, 15, garage_amount
  DropListBox 55, 125, 90, 15, "(Select one...)"+chr(9)+"Heat/AC"+chr(9)+"Phone/Electric"+chr(9)+"Phone only"+chr(9)+"Electric only"+chr(9)+"All utilities included in rent"+chr(9)+"None", utilities_paid_by_resident_listbox
  EditBox 50, 150, 120, 15, other_notes
  checkbox 20, 170, 75, 10, "Signed by LL/Mgr?", signed_by_LLMgr_check
  checkbox 110, 170, 75, 10, "Signed by client?", signed_by_client_check
  EditBox 65, 190, 80, 15, worker_signature
  ButtonGroup ButtonPressed
	OkButton 100, 210, 50, 15
	CancelButton 155, 210, 50, 15
  Text 5, 45, 45, 15, "Phone #:"
  Text 5, 65, 35, 15, "Total cost:"
  Text 95, 65, 55, 10, "Cost per person:"
  Text 5, 125, 50, 15, "Utilities paid by resident:"
  Text 5, 150, 45, 15, "Other notes:"
  Text 0, 190, 60, 15, "Worker signature:"
  Text 5, 5, 40, 15, "Date moved in:"
  Text 5, 25, 45, 15, "New address:"
  Text 110, 5, 40, 15, "How many residents?"
  Text 100, 80, 30, 10, "Amount"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to Bluezone & grabbing case number and footer year/month
EMConnect ""
call MAXIS_case_number_finder(case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

DO
	Dialog case_number_dialog																'calls up dialog for worker to enter case number and applicable month and year.	 Script will 'loop' 
	IF buttonpressed = 0 THEN StopScript						   'and verbally request the worker to enter a case number until the worker enters a case number.
	IF case_number = "" THEN MsgBox "You must enter a case number"
LOOP UNTIL case_number <> ""


DO
	DO
		DO
			DO
				DO
					DO
						Dialog Shelter_form_received_dialog									   'calls up dialog for worker to enter information provided by client on the shelter form
						cancel_confirmation
						IF worker_signature = "" THEN MsgBox "You must sign your case note"
					LOOP UNTIL worker_signature <> ""
					IF date_moved_in = "" THEN MsgBox "You must enter the date the client moved in"
				LOOP UNTIL date_moved_in <> ""
				IF how_many_residents = "" THEN MsgBox "You must enter the amount of residents"
			LOOP UNTIL how_many_residents <> ""
			IF new_address = "" THEN MsgBox "You must enter the new address"
		LOOP UNTIL new_address <> ""
		IF total_cost = "" THEN MsgBox "You must enter the total shelter cost"
	LOOP UNTIL total_cost <> ""
	IF utilities_paid_by_resident_listbox =	 "(Select one...)" THEN MsgBox "You must make a valid selection of utilities paid by the resident"
LOOP UNTIL utilities_paid_by_resident_listbox <> "(Select one...)"


'checking for an active MAXIS session
Call check_for_MAXIS(False)


'Dollar bill symbol will be added to numeric variables
IF total_cost <> "" THEN total_cost = "$" & total_cost
IF cost_per_person <> "" THEN cost_per_person = "$" & cost_per_person
IF subsidized_amount <> "" THEN subsidized_amount = "$" & subsidized_amount
IF garage_amount <> "" THEN garage_amount = "$" & garage_amount

'THE CASE NOTE----------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
Call write_variable_in_case_note ("~~~Shelter form rec'd~~~")
Call write_bullet_and_variable_in_case_note ("Date client moved in", date_moved_in )
Call write_bullet_and_variable_in_case_note ("Number of residents", how_many_residents)
Call write_bullet_and_variable_in_case_note ("New address", new_address)
Call write_bullet_and_variable_in_case_note ("Phone number", phone_number)
Call write_bullet_and_variable_in_case_note ("Total cost", total_cost)
Call write_bullet_and_variable_in_case_note ("Cost per person", cost_per_person)
IF subsidized_amount_check = 1 THEN Call write_bullet_and_variable_in_case_note ("Subsidized amount", subsidized_amount)
IF garage_amount_check = 1 THEN Call write_bullet_and_variable_in_case_note ("Garage amount", garage_amount)
Call write_bullet_and_variable_in_case_note ("Utilities paid by resident", utilities_paid_by_resident_listbox)
Call write_bullet_and_variable_in_case_note ("Other notes", other_notes)
IF signed_by_LLMgr_check = 1 THEN Call write_variable_in_case_note ("* Signed by LL/Mgr.")
IF signed_by_client_check = 1 THEN Call write_variable_in_case_note ("* Signed by client.")
Call write_variable_in_case_note ("---")
call write_variable_in_case_note (worker_signature)

script_end_procedure ("")																										   'closing script and writing stats
