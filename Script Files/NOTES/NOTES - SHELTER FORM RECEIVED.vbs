'Option Explicit

'DIM beta_agency

'LOADING ROUTINE FUNCTIONS---------------------------------------------------------------
'DIM url, req, fso
If beta_agency = "" or beta_agency = True then
	url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
Else
	url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
End if
SET req = CreateObject("Msxml2.XMLHttp.6.0")															   'Creates an object to get a URL
req.open "GET", url, FALSE																																			'Attempts to open the URL
req.send																																																			  'Sends request
IF req.Status = 200 THEN																																			   '200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText																														  'Executes the script code
ELSE																																																					   'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox "Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
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


'DIM shelter_form_received_dialog		'declaring variables that are being used in the rest of the script
'DIM date_moved_in_editbox
'DIM ButtonPressed
'DIM new_address_editbox
'DIM cost_per_person_editbox
'DIM how_many_residents_editbox
'DIM total_cost_editbox
'DIM other_notes_editbox
'DIM worker_signature
'DIM utilities_paid_by_resident_listbox
'DIM phonenumber_editbox
'DIM subsidized_amount_editbox
'DIM garage_amount_checkbox
'DIM garage_amount_editbox
'DIM signed_by_LLMgr_checkbox
'DIM signed_by_client_checkbox
'DIM case_number
'DIM month_editbox
'DIM year_editbox
'DIM case_number_dialogbox
'DIM subsidized_amount_checkbox


BeginDialog case_number_dialogbox, 0, 0, 191, 80, "Dialog"						   'dialog box where worker enters the case number (and at some point applicable month & year)
  EditBox 75, 15, 80, 15, case_number																	   'once worker selects ok, it will move to the next dialog box.  If worker selects cancel, then
  'EditBox 75, 35, 40, 15, month_editbox																 'script will end
  'EditBox 120, 35, 35, 15, year_editbox
  ButtonGroup ButtonPressed
	OkButton 80, 60, 50, 15
	CancelButton 135, 60, 50, 15
  'Text 10, 35, 60, 15, "Footer month:"
  Text 10, 15, 60, 15, "Case number: "
EndDialog

BeginDialog Shelter_form_received_dialog, 0, 0, 206, 225, "Dialog"							'Dialogue box completed by worker with information provided by the client regarding the shelter form that was received 
  EditBox 50, 5, 55, 15, date_moved_in_editbox
  EditBox 155, 5, 55, 15, how_many_residents_editbox
  EditBox 50, 25, 150, 15, new_address_editbox
  EditBox 50, 45, 100, 15, phonenumber_editbox
  EditBox 50, 65, 45, 15, total_cost_editbox
  EditBox 150, 65, 50, 15, cost_per_person_editbox
  CheckBox 10, 90, 75, 10, "Subsidized amount", subsidized_amount_checkbox
  EditBox 95, 90, 40, 15, subsidized_amount_editbox
  CheckBox 10, 105, 75, 10, "Garage amount", garage_amount_checkbox
  EditBox 95, 105, 40, 15, garage_amount_editbox
  DropListBox 55, 125, 90, 15, "(Select one...)"+chr(9)+"Heat/AC"+chr(9)+"Phone/Electric"+chr(9)+"Phone only"+chr(9)+"Electric only"+chr(9)+"All utilities included in rent"+chr(9)+"None", utilities_paid_by_resident_listbox
  EditBox 50, 150, 120, 15, other_notes_editbox
  CheckBox 20, 170, 75, 10, "Signed by LL/Mgr?", signed_by_LLMgr_checkbox
  CheckBox 110, 170, 75, 10, "Signed by client?", signed_by_client_checkbox
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



EMConnect ""				   'Connecting to Bluezone

call MAXIS_case_number_finder(case_number)								'function autofills case number that worker already has on MAXIS screen

DO
	Dialog case_number_dialogbox																'calls up dialog for worker to enter case number and applicable month and year.	 Script will 'loop' 
	IF buttonpressed = cancel THEN StopScript						   'and verbally request the worker to enter a case number until the worker enters a case number.
	IF case_number = "" THEN MsgBox "You must enter a case number"
LOOP UNTIL case_number <> ""

Call check_for_MAXIS(true)																		 'ensures that worker has not "passworded" out of MAXIS

DO
	DO
		DO
			DO
				DO
					DO
						Dialog Shelter_form_received_dialog									   'calls up dialog for worker to enter information provided by client on the shelter form
						IF buttonpressed = cancel THEN StopScript
						IF worker_signature = "" THEN MsgBox "You must sign your case note"
					LOOP UNTIL worker_signature <> ""
					IF date_moved_in_editbox = "" THEN MsgBox "You must enter the date the client moved in"
				LOOP UNTIL date_moved_in_editbox <> ""
				IF how_many_residents_editbox = "" THEN MsgBox "You must enter the amount of residents"
			LOOP UNTIL how_many_residents_editbox <> ""
			IF new_address_editbox = "" THEN MsgBox "You must enter the new address"
		LOOP UNTIL new_address_editbox <> ""
		IF total_cost_editbox = "" THEN MsgBox "You must enter the total shelter cost"
	LOOP UNTIL total_cost_editbox <> ""
	IF utilities_paid_by_resident_listbox =	 "(Select one...)" THEN MsgBox "You must make a valid selection of utilities paid by the resident"
LOOP UNTIL utilities_paid_by_resident_listbox <> "(Select one...)"


Call check_for_MAXIS(true)																										 'ensures that worker has not "passworded" out of MAXIS

Call navigate_to_screen ("case", "note")								'function to navigate user to case note
PF9																																																			'brings case note into edit mode

'Dollar bill symbol will be added to numeric variables 
IF total_cost_editbox <> "" THEN total_cost_editbox = "$" & total_cost_editbox
IF cost_per_person_editbox <> "" THEN cost_per_person_editbox = "$" & cost_per_person_editbox
IF subsidized_amount_editbox <> "" THEN subsidized_amount_editbox = "$" & subsidized_amount_editbox
IF garage_amount_editbox <> "" THEN garage_amount_editbox = "$" & garage_amount_editbox


Call write_variable_in_case_note ("~~~Shelter form rec'd~~~")												'adding information to case note
Call write_bullet_and_variable_in_case_note ("Date client moved in", date_moved_in_editbox )		
Call write_bullet_and_variable_in_case_note ("Number of residents", how_many_residents_editbox)			
Call write_bullet_and_variable_in_case_note ("New address", new_address_editbox)		  
Call write_bullet_and_variable_in_case_note ("Phone number", phonenumber_editbox)				 
Call write_bullet_and_variable_in_case_note ("Total cost", total_cost_editbox)					 
Call write_bullet_and_variable_in_case_note ("Cost per person", cost_per_person_editbox)			 
IF subsidized_amount_checkbox = 1 THEN Call write_bullet_and_variable_in_case_note ("Subsidized amount", subsidized_amount_editbox)			   
IF garage_amount_checkbox = 1 THEN Call write_bullet_and_variable_in_case_note ("Garage amount", garage_amount_editbox)				  
Call write_bullet_and_variable_in_case_note ("Utilities paid by resident", utilities_paid_by_resident_listbox) 
Call write_bullet_and_variable_in_case_note ("Other notes", other_notes_editbox)				
IF signed_by_LLMgr_checkbox = 1 THEN Call write_variable_in_case_note ("* Signed by LL/Mgr.")			  
IF signed_by_client_checkbox = 1 THEN Call write_variable_in_case_note ("* Signed by client.")
Call write_variable_in_case_note ("---")						 
call write_variable_in_case_note (worker_signature)

script_end_procedure ("")																										   'closing script and writing stats
