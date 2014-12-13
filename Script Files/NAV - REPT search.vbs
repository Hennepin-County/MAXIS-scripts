'This script can be used to look up clients in the REPT/ACTV. This could be useful when checking voicemails if the client garbles their name or has a difficult name to look up or if you just want an easier way of checking your REPT screens.

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NAV - REPT Search"
start_time = timer


'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

BeginDialog search_dialog, 0, 0, 186, 135, "Client Look Up"
  EditBox 120, 10, 55, 15, person_look_up
  EditBox 120, 65, 55, 15, case_load_look_up
  DropListBox 115, 90, 60, 15, "REPT/ACTV"+chr(9)+"REPT/INAC"+chr(9)+"REPT/PND1"+chr(9)+"REPT/PND2"+chr(9)+"REPT/REVW", search_where
  ButtonGroup ButtonPressed
    OkButton 45, 115, 50, 15
    CancelButton 95, 115, 50, 15
  Text 5, 10, 105, 25, "Client last name (need not be full last name. Ex. H to search for Henderson)"
  Text 5, 70, 115, 10, "Worker/Team Number (x######):"
  Text 5, 45, 160, 15, "NOTE: Leave the following edit box BLANK to search your own case load."
  Text 5, 90, 100, 10, "Where to search:"
EndDialog

EMConnect ""

maxis_check_function

'Shows dialog, requires length to be 7, or if it's 3, will add the worker_county_code in.
DO
	DIALOG search_dialog
	IF ButtonPressed = 0 THEN stopscript
	IF len(case_load_look_up) = 3 THEN case_load_look_up = worker_county_code & case_load_look_up
	If len(case_load_look_up) <> 7 then MsgBox("You must enter in a seven digit worker number. Please try again.")
LOOP UNTIL case_load_look_up = "" OR (len(case_load_look_up) = 7 AND (ucase(LEFT(case_load_look_up, 1) = "X") or lcase(LEFT(case_load_look_up, 1) = "x")))

person_look_up = UCASE(person_look_up)
search_length = len(person_look_up)

'========== Checks REPT/ACTV ==========
IF search_where = "REPT/ACTV" THEN 
	Call navigate_to_screen("rept", "actv")
	IF case_load_look_up <> "" THEN
		EMWriteScreen case_load_look_up, 21, 13
		transmit
	END IF
	
	IF len(person_look_up) > 8 THEN person_look_up = left(person_look_up, 8)
	EMWriteScreen person_look_up, 21, 33
	transmit
	MsgBox("REPT/ACTV Search Complete. The person matching your search may be on this page. If not, consider checking the next page(s) or revising your search.")
END IF

'========== Checks REPT/INAC ==========
IF search_where = "REPT/INAC" THEN
	Call navigate_to_screen("rept", "inac")
	IF case_load_look_up <> "" THEN
		EMWriteScreen case_load_look_up, 21, 16
		transmit
	END IF

	DO
		MAXIS_row = 7
		EMReadScreen last_page_check, 21, 24, 2
		DO
			EMReadScreen client_name, 24, MAXIS_row, 14
			client_name = left(client_name, search_length)
			IF person_look_up = client_name THEN
				EMReadScreen full_name, 24, MAXIS_row, 14
				EMReadScreen case_number, 8, MAXIS_row, 3
				MsgBox("REPT/INAC Search Complete. The person matching your search may be on this page. If not, consider checking the next page(s) or revising your search.")
			ELSE
				MAXIS_row = MAXIS_row + 1
			END IF
		LOOP UNTIL MAXIS_row = 19
		PF8
		IF last_page_check = "THIS IS THE LAST PAGE" & client_name <> person_look_up THEN MsgBox("Your search yielded no successful hits in " & search_where & ". You may want to check another REPT or the spelling of the client's name.")
	LOOP UNTIL last_page_check = "THIS IS THE LAST PAGE" or client_name = person_look_up
END IF

'========== Checks REPT/PND1 ==========
IF search_where = "REPT/PND1" THEN
	Call navigate_to_screen("rept", "pnd1")
	IF case_load_look_up <> "" THEN
		EMWriteScreen case_load_look_up, 21, 13
		transmit
	END IF

	DO
		MAXIS_row = 7
		EMReadScreen last_page_check, 21, 24, 2
		DO
			EMReadScreen client_name, 24, MAXIS_row, 14
			client_name = left(client_name, search_length)
			IF person_look_up = client_name THEN
				EMReadScreen full_name, 24, MAXIS_row, 14
				EMReadScreen case_number, 8, MAXIS_row, 3
				MsgBox("REPT/PND1 Search Complete. The person matching your search may be on this page. If not, consider checking the next page(s) or revising your search.")
			ELSE
				MAXIS_row = MAXIS_row + 1
			END IF
		LOOP UNTIL MAXIS_row = 19
		PF8
		IF last_page_check = "THIS IS THE LAST PAGE" & client_name <> person_look_up THEN MsgBox("Your search yielded no successful hits in " & search_where & ". You may want to check another REPT or the spelling of the client's name.")
	LOOP UNTIL last_page_check = "THIS IS THE LAST PAGE" or client_name = person_look_up
END IF

'========== Checks REPT/PND2 ==========
IF search_where = "REPT/PND2" THEN
	Call navigate_to_screen("rept", "pnd2")
		IF case_load_look_up <> "" THEN
			EMWriteScreen case_load_look_up, 21, 13
			transmit
		END IF

	DO
		MAXIS_row = 7
		EMReadScreen last_page_check, 21, 24, 2
		DO
			EMReadScreen client_name, 24, MAXIS_row, 16
			client_name = left(client_name, search_length)
			IF person_look_up = client_name THEN
				EMReadScreen full_name, 24, MAXIS_row, 16
				EMReadScreen case_number, 8, MAXIS_row, 5
				MsgBox("REPT/PND2 Search Complete. The person matching your search may be on this page. If not, consider checking the next page(s) or revising your search.")
			ELSE
				MAXIS_row = MAXIS_row + 1
			END IF
		LOOP UNTIL MAXIS_row = 19
		PF8
		IF last_page_check = "THIS IS THE LAST PAGE" AND client_name <> person_look_up THEN MsgBox("Your search yielded no successful hits in " & search_where & ". You may want to check another REPT or the spelling of the client's name.")
	LOOP UNTIL last_page_check = "THIS IS THE LAST PAGE" or client_name = person_look_up
END IF

'========== Checks REPT/REVW ==========
IF search_where = "REPT/REVW" THEN
	Call navigate_to_screen("rept", "REVW")
	IF case_load_look_up <> "" THEN
		EMWriteScreen case_load_look_up, 21, 6
		transmit
	END IF

	DO
		MAXIS_row = 7
		EMReadScreen last_page_check, 21, 24, 2
		DO
			EMReadScreen client_name, 24, MAXIS_row, 16
			client_name = left(client_name, search_length)
			IF person_look_up = client_name THEN
				EMReadScreen full_name, 24, MAXIS_row, 16
				EMReadScreen case_number, 8, MAXIS_row, 6
				MsgBox("REPT/REVW Search Complete. The person matching your search may be on this page. If not, consider checking the next page(s) or revising your search.")
			ELSE
				MAXIS_row = MAXIS_row + 1
			END IF
		LOOP UNTIL MAXIS_row = 19
		PF8
		IF last_page_check = "THIS IS THE LAST PAGE" & client_name <> person_look_up THEN MsgBox("Your search yielded no successful hits in " & search_where & ". You may want to check another REPT or the spelling of the client's name.")
	LOOP UNTIL last_page_check = "THIS IS THE LAST PAGE" or client_name = person_look_up
END IF

'Ends script (and writes stats if stats_collection = True on FUNCTIONS FILE
script_end_procedure("")