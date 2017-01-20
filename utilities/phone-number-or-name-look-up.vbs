'This script combines the name search and phone number search OTHER NAV scripts. By Tim DeLong.
'Required for statistical purposes===============================================================================
name_of_script = "UTILITIES - PHONE NUMBER OR NAME LOOK UP.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 10                      'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
'END OF stats block==============================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Checks for county info from global variables, or asks if it is not already defined.
get_county_code

BeginDialog Dialog1, 0, 0, 186, 180, "Client Look Up"
  EditBox 120, 25, 55, 15, person_look_up
  EditBox 120, 60, 55, 15, phone_look_up
  EditBox 120, 115, 55, 15, case_load_look_up
  DropListBox 95, 140, 60, 15, "REPT/ACTV"+chr(9)+"REPT/INAC"+chr(9)+"REPT/PND1"+chr(9)+"REPT/PND2"+chr(9)+"REPT/REVW", search_where
  ButtonGroup ButtonPressed
    OkButton 20, 160, 50, 15
    CancelButton 120, 160, 50, 15
  Text 5, 20, 105, 25, "Client last name (need not be full last name. Ex. H to search for Henderson)"
  Text 5, 50, 105, 35, "Phone number to search. 10 Digit format (including area code). Do not include spaces or dashes"
  Text 5, 115, 110, 15, "Worker/Team Number (x######):"
  Text 20, 90, 150, 20, "NOTE: Leave the following edit box BLANK to search your own case load."
  Text 30, 140, 60, 10, "Where to search:"
  Text 10, 5, 170, 10, "**SEARCH BY PHONE NUMBER OR LAST NAME.**"
EndDialog




EMConnect ""

CALL check_for_MAXIS(TRUE)

call find_variable("User: ", user_number, 7)

DO
	DIALOG
	IF ButtonPressed = 0 THEN stopscript
	IF len(case_load_look_up) = 3 THEN case_load_look_up = worker_county_code & case_load_look_up
LOOP UNTIL case_load_look_up = "" OR (len(case_load_look_up) = 7 AND (ucase(LEFT(case_load_look_up, 1) = "X") or lcase(LEFT(case_load_look_up, 1) = "x")))

phone = 0
person = 0
last_page_check = 0

IF phone_look_up > "" THEN phone = 1
IF person_look_up > "" THEN person = 1

phone_look_up = replace(phone_look_up, " ", "")
phone_look_up = replace(phone_look_up, "-", "")

person_look_up = UCASE(person_look_up)
search_length = len(person_look_up)


'========== Checks REPT/ACTV ==========
IF search_where = "REPT/ACTV" THEN
	Call navigate_to_MAXIS_screen("rept", "actv")
	IF case_load_look_up <> "" and ucase(user_number) <> ucase(case_load_look_up) THEN
		EMWriteScreen case_load_look_up, 21, 13
		transmit
  	END IF
	IF person = 1 THEN
		IF len(person_look_up) > 8 THEN person_look_up = left(person_look_up, 8)
		EMWriteScreen person_look_up, 21, 33
		transmit
		MsgBox("REPT/ACTV Search Complete. The person matching your search may be on this page. If not, consider checking the next page(s) or revising your search.")
	END IF

	IF phone = 1 THEN
  		Do
			MAXIS_row = 7
			EMReadScreen last_page_check, 21, 24, 2
			Do
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 12
				If MAXIS_case_number = "        " then exit do
				MAXIS_case_number = replace(MAXIS_case_number, " ", "")
				case_number_array = case_number_array & " " & MAXIS_case_number
				MAXIS_row = MAXIS_row + 1
			Loop until MAXIS_row = 19
			PF8	'No need for STATS counter on this because there's direct navigation for last name
			EMReadScreen last_page_check, 21, 24, 2
  		Loop until last_page_check = "THIS IS THE LAST PAGE"
	END IF
END IF

'========== Checks REPT/INAC ==========
IF search_where = "REPT/INAC" THEN
	Call navigate_to_MAXIS_screen("rept", "inac")
	IF case_load_look_up <> "" and ucase(user_number) <> ucase(case_load_look_up) THEN
		EMWriteScreen case_load_look_up, 21, 16
		transmit
  	END IF

	IF person = 1 THEN
		DO
			MAXIS_row = 7
			EMReadScreen last_page_check, 21, 24, 2
			DO
				EMReadScreen client_name, 24, MAXIS_row, 14
				client_name = left(client_name, search_length)
				IF person_look_up = client_name THEN
					EMReadScreen full_name, 24, MAXIS_row, 14
					EMReadScreen MAXIS_case_number, 8, MAXIS_row, 3
					MsgBox("REPT/INAC Search Complete. The person matching your search may be. If not, consider checking the next page(s) or revising your search.")
				ELSE
					MAXIS_row = MAXIS_row + 1
				END IF
			LOOP UNTIL MAXIS_row = 19 or person_look_up = client_name or last_page_check = "THIS IS THE LAST PAGE"
			IF MAXIS_row = 19 THEN PF8
		LOOP UNTIL person_look_up = client_name or last_page_check = "THIS IS THE LAST PAGE"
		IF last_page_check = "THIS IS THE LAST PAGE" THEN
			IF client_name <> person_look_up THEN MsgBox ("Your search yielded no successful hits in " & search_where & ". You may want to check another REPT or the spelling of the client's name.")
		END IF
	END IF

	IF phone = 1 THEN
  		Do
			MAXIS_row = 7
			EMReadScreen last_page_check, 21, 24, 2
			Do
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 3
				If MAXIS_case_number = "        " then exit do
				MAXIS_case_number = replace(MAXIS_case_number, " ", "")
				case_number_array = case_number_array & " " & MAXIS_case_number
				MAXIS_row = MAXIS_row + 1
			Loop until MAXIS_row = 19
			PF8
			STATS_counter = STATS_counter + 1
			EMReadScreen last_page_check, 21, 24, 2
  		Loop until last_page_check = "THIS IS THE LAST PAGE"
	END IF
END IF

'========== Checks REPT/PND1 ==========
IF search_where = "REPT/PND1" THEN
  	Call navigate_to_MAXIS_screen("rept", "pnd1")
  	IF case_load_look_up <> "" and ucase(user_number) <> ucase(case_load_look_up) THEN
    	EMWriteScreen case_load_look_up, 21, 13
    	transmit
  	END IF

	IF person = 1 THEN
		DO
			MAXIS_row = 7
			EMReadScreen last_page_check, 21, 24, 2
			DO
				EMReadScreen client_name, 24, MAXIS_row, 14
				client_name = left(client_name, search_length)
				IF person_look_up = client_name THEN
					EMReadScreen full_name, 24, MAXIS_row, 14
					EMReadScreen MAXIS_case_number, 8, MAXIS_row, 3
					MsgBox("REPT/PND1 Search Complete. The person matching your search may be on this page. If not, consider checking the next page(s) or revising your search.")
				ELSE
					MAXIS_row = MAXIS_row + 1
				END IF
			LOOP UNTIL MAXIS_row = 19 or person_look_up = client_name or last_page_check = "THIS IS THE LAST PAGE"
			IF MAXIS_row = 19 THEN
				PF8
				STATS_counter = STATS_counter + 1
			End if
		LOOP UNTIL person_look_up = client_name or last_page_check = "THIS IS THE LAST PAGE"
		IF last_page_check = "THIS IS THE LAST PAGE" THEN
			IF client_name <> person_look_up THEN MsgBox ("Your search yielded no successful hits in " & search_where & ". You may want to check another REPT or the spelling of the client's name.")
		END IF
	END IF


	IF phone = 1 THEN
  		Do
			MAXIS_row = 7
			EMReadScreen last_page_check, 21, 24, 2
			Do
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 3
				If MAXIS_case_number = "        " then exit do
				MAXIS_case_number = replace(MAXIS_case_number, " ", "")
				case_number_array = case_number_array & " " & MAXIS_case_number
				MAXIS_row = MAXIS_row + 1
			Loop until MAXIS_row = 19
			PF8
			STATS_counter = STATS_counter + 1
			EMReadScreen last_page_check, 21, 24, 2
  		Loop until last_page_check = "THIS IS THE LAST PAGE"
	END IF

END IF

'========== Checks REPT/PND2 ==========
IF search_where = "REPT/PND2" THEN
  	Call navigate_to_MAXIS_screen("rept", "pnd2")
  	IF case_load_look_up <> "" and ucase(user_number) <> ucase(case_load_look_up) THEN
    	EMWriteScreen case_load_look_up, 21, 13
    	transmit
  	END IF

	IF person = 1 THEN
		DO
			MAXIS_row = 7
			EMReadScreen last_page_check, 21, 24, 2
			DO
				EMReadScreen client_name, 24, MAXIS_row, 16
				client_name = left(client_name, search_length)
				IF person_look_up = client_name THEN
					EMReadScreen full_name, 24, MAXIS_row, 16
					EMReadScreen MAXIS_case_number, 8, MAXIS_row, 5
					MsgBox("REPT/PND2 Search Complete. The person matching your search may be on this page. If not, consider checking the next page(s) or revising your search.")
				ELSE
					MAXIS_row = MAXIS_row + 1
				END IF
			LOOP UNTIL MAXIS_row = 19 or person_look_up = client_name or last_page_check = "THIS IS THE LAST PAGE"
			IF MAXIS_row = 19 THEN
				PF8
				STATS_counter = STATS_counter + 1
			End if
		LOOP UNTIL person_look_up = client_name or last_page_check = "THIS IS THE LAST PAGE"
		IF last_page_check = "THIS IS THE LAST PAGE" THEN
			IF client_name <> person_look_up THEN MsgBox ("Your search yielded no successful hits in " & search_where & ". You may want to check another REPT or the spelling of the client's name.")
		END IF
	END IF

	IF phone = 1 THEN
		Do
			MAXIS_row = 7
			EMReadScreen last_page_check, 21, 24, 2
			Do
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 5
				If MAXIS_case_number = "        " then exit do
				MAXIS_case_number = replace(MAXIS_case_number, " ", "")
				case_number_array = case_number_array & " " & MAXIS_case_number
				MAXIS_row = MAXIS_row + 1
			Loop until MAXIS_row = 19
			PF8
			STATS_counter = STATS_counter + 1
			EMReadScreen last_page_check, 21, 24, 2
 		Loop until last_page_check = "THIS IS THE LAST PAGE"
	END IF
END IF

'========== Checks REPT/REVW ==========
IF search_where = "REPT/REVW" THEN
  	Call navigate_to_MAXIS_screen("rept", "REVW")
  	IF case_load_look_up <> "" and ucase(user_number) <> ucase(case_load_look_up) THEN
    	EMWriteScreen case_load_look_up, 21, 6
    	transmit
  	END IF

	IF person = 1 THEN
		DO
			MAXIS_row = 7
			EMReadScreen last_page_check, 21, 24, 2
			DO
				EMReadScreen client_name, 24, MAXIS_row, 16
				client_name = left(client_name, search_length)
				IF person_look_up = client_name THEN
					EMReadScreen full_name, 24, MAXIS_row, 16
					EMReadScreen MAXIS_case_number, 8, MAXIS_row, 6
					MsgBox("REPT/REVW Search Complete. The person matching your search may be on this page. If not, consider checking the next page(s) or revising your search.")
				ELSE
					MAXIS_row = MAXIS_row + 1
				END IF
			LOOP UNTIL MAXIS_row = 19 or person_look_up = client_name or last_page_check = "THIS IS THE LAST PAGE"
			IF MAXIS_row = 19 THEN
				PF8
				STATS_counter = STATS_counter + 1
			End if
		LOOP UNTIL person_look_up = client_name or last_page_check = "THIS IS THE LAST PAGE"
		IF last_page_check = "THIS IS THE LAST PAGE" THEN
			IF client_name <> person_look_up THEN MsgBox ("Your search yielded no successful hits in " & search_where & ". You may want to check another REPT or the spelling of the client's name.")
		END IF
	END IF


	IF phone = 1 THEN
  		Do
			MAXIS_row = 7
			EMReadScreen last_page_check, 21, 24, 2
			Do
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 6
				If MAXIS_case_number = "        " then exit do
				MAXIS_case_number = replace(MAXIS_case_number, " ", "")
				case_number_array = case_number_array & " " & MAXIS_case_number
				MAXIS_row = MAXIS_row + 1
			Loop until MAXIS_row = 19
			PF8
			STATS_counter = STATS_counter + 1
			EMReadScreen last_page_check, 21, 24, 2
  		Loop until last_page_check = "THIS IS THE LAST PAGE"
	END IF
END IF

'========== Checking ADDR against reported phone number =====================
'cleaning up array
IF phone = 1 THEN

'Required for statistical purposes==========================================================================================
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 15                      'manual run time in seconds
STATS_denomination = "I"                   'I is for each ITEM
'END OF stats block=========================================================================================================

case_number_array = TRIM(case_number_array)
case_number_array = SPLIT(case_number_array)

FOR EACH MAXIS_case_number in case_number_array
	back_to_self
	EMwritescreen "          ", 18, 43
	EMwritescreen MAXIS_case_number, 18, 43
	CALL navigate_to_MAXIS_screen("STAT", "ADDR")
	row = 1
	col = 1
	EMSearch "PRIVILEGED", row, col
	IF row <> 0 THEN msgbox MAXIS_case_number
	IF row = 0 THEN
		EMReadscreen area_code_1, 3, 17, 45
		EMReadscreen addr_phone_number_1, 8, 17, 51
		EMReadscreen area_code_2, 3, 18, 45
		EMReadscreen addr_phone_number_2, 8, 18, 51
		EMReadscreen area_code_3, 3, 19, 45
		EMReadscreen addr_phone_number_3, 8, 19, 51
		complete_phone_1 = area_code_1 & replace(addr_phone_number_1, " ", "")
		complete_phone_2 = area_code_2 & replace(addr_phone_number_2, " ", "")
		complete_phone_3 = area_code_3 & replace(addr_phone_number_3, " ", "")
		IF complete_phone_1 = phone_look_up OR complete_phone_2 = phone_look_up OR complete_phone_3 = phone_look_up then script_end_procedure(MAXIS_case_number & " contains requested phone number " & phone_look_up & ".")
		CALL navigate_to_MAXIS_screen("STAT", "AREP")
		EMReadscreen arep_area_code_1, 3, 8, 34
		EMReadscreen arep_phone_number_1, 8, 8, 40
		EMReadscreen arep_area_code_2, 3, 9, 34
		EMReadscreen arep_phone_number_2, 8, 9, 40
		arep_complete_phone_1 = arep_area_code_1 & replace(arep_phone_number_1, " ", "")
		arep_complete_phone_2 = arep_area_code_2 & replace(arep_phone_number_2, " ", "")
		IF arep_complete_phone_1 = phone_look_up OR arep_complete_phone_2 = phone_look_up then script_end_procedure("AREP on case " & MAXIS_case_number & " contains requested phone number " & phone_look_up & ".")
	END IF
	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
NEXT

script_end_procedure(phone_look_up & " was not found in selected REPT. Feel free to try another REPT list, change your footer month, or verify that the number you entered is correct")

END IF

script_end_procedure("")
