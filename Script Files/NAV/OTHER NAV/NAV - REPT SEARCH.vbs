'This script can be used to look up clients in the REPT/ACTV. This could be useful when checking voicemails if the client garbles their name or has a difficult name to look up or if you just want an easier way of checking your REPT screens.

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NAV - REPT-SEARCH.vbs"
start_time = timer


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
