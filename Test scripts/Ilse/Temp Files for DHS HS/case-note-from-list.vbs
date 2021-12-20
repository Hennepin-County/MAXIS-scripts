'This is used by almost every script which calls a specific agency worker number (like the REPT/ACTV nav and list gen scripts).
worker_county_code = "x191"

'This is an "updated date" variable, which is updated dynamically by the intaller.
scripts_updated_date = "01/01/2099"

'This is a setting to determine if changes to scripts will be displayed in messageboxes in real time to end users
changelog_enabled = true

'COLLECTING STATISTICS=========================

'This is used for determining whether script_end_procedure will also log usage info in an Access table.
collecting_statistics = true

'This is a variable used to determine if the agency is using a SQL database or not. Set to true if you're using SQL. Otherwise, set to false.
using_SQL_database = true

'This is the file path for the statistics Access database.
stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"

'If the "enhanced database" is used (with new features added in January 2016), this variable should be set to true
STATS_enhanced_db = true

'If set to true, the case number will be collected and input into the database
collect_MAXIS_case_number = true

'Required for statistical purposes===============================================================================
name_of_script = "BULK - CASE NOTE FROM LIST.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 180                               'manual run time in seconds
STATS_denomination = "C"       'C is for each Case
'END OF stats block==============================================================================================

'============THAT MEANS THAT IF YOU BREAK THIS SCRIPT, ALL OTHER SCRIPTS ****STATEWIDE**** WILL NOT WORK! MODIFY WITH CARE!!!!!============
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'GLOBAL CONSTANTS----------------------------------------------------------------------------------------------------
Dim checked, unchecked, cancel, OK, blank, t_drive, STATS_counter, STATS_manualtime, STATS_denomination, script_run_lowdown, testing_run, MAXIS_case_number		'Declares this for Option Explicit users

checked = 1			'Value for checked boxes
unchecked = 0		'Value for unchecked boxes
cancel = 0			'Value for cancel button in dialogs
OK = -1			'Value for OK button in dialogs
blank = ""

'Determines CM and CM+1 month and year using the two rightmost chars of both the month and year. Adds a "0" to all months, which will only pull over if it's a single-digit-month
Dim CM_mo, CM_yr, CM_plus_1_mo, CM_plus_1_yr, CM_plus_2_mo, CM_plus_2_yr
'var equals...  the right part of...    the specific part...    of either today or next month... just the right 2 chars!
CM_mo =         right("0" &             DatePart("m",           date                             ), 2)
CM_yr =         right(                  DatePart("yyyy",        date                             ), 2)

CM_plus_1_mo =  right("0" &             DatePart("m",           DateAdd("m", 1, date)            ), 2)
CM_plus_1_yr =  right(                  DatePart("yyyy",        DateAdd("m", 1, date)            ), 2)

CM_plus_2_mo =  right("0" &             DatePart("m",           DateAdd("m", 2, date)            ), 2)
CM_plus_2_yr =  right(                  DatePart("yyyy",        DateAdd("m", 2, date)            ), 2)

If worker_county_code   = "" then worker_county_code = "MULTICOUNTY"
IF PRISM_script <> true then county_name = ""		'VKC NOTE 08/12/2016: ADDED IF...THEN CONDITION BECAUSE PRISM IS STILL USING THIS VARIABLE IN ALL SCRIPTS.vbs. IT WILL BE REMOVED AND THIS CAN BE RESTORED.

If ButtonPressed <> "" then ButtonPressed = ""		'Defines ButtonPressed if not previously defined, allowing scripts the benefit of not having to declare ButtonPressed all the time

'>>>>> Function to build dlg for manual entry <<<<<
FUNCTION build_manual_entry_dlg(case_number_array, case_note_header, case_note_body, worker_signature)
	'Array for all case numbers
	'This was chosen over building a dlg with 50 variables
	REDim all_cases_array(50, 0)

	'case_note_header = "***Recertification Accuracy Update***"
    'case_note_body = "This client receives a special diet allotment. The Special Diet form was mailed to the client to allow time for a physician to complete the form before the 06/20 recertification is due. If the special diet form is not returned, the MSA will be approved without the special diet allotment. ---CM 23.12 Special Diets need to be verified at recertification even if the special diet form says lifelong or ongoing.--- "
    Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 331, 330, "Enter MAXIS case numbers"
		Text 10, 15, 140, 10, "Enter MAXIS case numbers below..."
		dlg_row = 30
		dlg_col = 10
		FOR i = 1 TO 50
			EditBox dlg_col, dlg_row, 55, 15, all_cases_array(i, 0)
			dlg_row = dlg_row + 20
			IF dlg_row = 230 THEN
				dlg_row = 30
				dlg_col = dlg_col + 65
			END IF
		NEXT
		text 10, 235, 120, 10, "Enter case note below"
		Text 10, 255, 25, 10, "Header:"
		Text 10, 275, 20, 10, "Body:"
		Text 10, 295, 60, 10, "Worker Signature:"
		EditBox 45, 250, 280, 15, case_note_header
		EditBox 35, 270, 290, 15, case_note_body
		EditBox 75, 290, 150, 15, worker_signature
		ButtonGroup ButtonPressed
			OkButton 220, 310, 50, 15
			CancelButton 270, 310, 50, 15
	EndDialog
 'TODO add exclamtion point and explain
	'Calling the dlg within the function
	DO
		'err_msg handling
		err_msg = ""
		DIALOG Dialog1
			cancel_without_confirmation
			FOR i = 1 TO 50
				all_cases_array(i, 0) = replace(all_cases_array(i, 0), " ", "")
				IF all_cases_array(i, 0) <> "" THEN
					IF len(all_cases_array(i, 0)) > 8 THEN err_msg = err_msg & vbCr & "* Case number " & all_cases_array(i, 0) & " is too long to be a valid MAXIS case number."
					IF isnumeric(all_cases_array(i, 0)) = FALSE THEN err_msg = err_msg & vbCr & "* Case number " & all_cases_array(i, 0) & " contains alphabetic characters. These are not valid."
				END IF
			NEXT
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""

	'building the array
	case_number_array = ""
	FOR i = 1 TO 50
		IF all_cases_array(i, 0) <> "" THEN case_number_array = case_number_array & all_cases_array(i, 0) & "~~~"
	NEXT
END FUNCTION

'>>>>> This function converts the letter for a number so the script can work with it <<<<<
FUNCTION convert_excel_letter_to_excel_number(excel_col)
	IF isnumeric(excel_col) = FALSE THEN
		alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		excel_col = ucase(excel_col)
		IF len(excel_col) = 1 THEN
			excel_col = InStr(alphabet, excel_col)
		ELSEIF len(excel_col) = 2 THEN
			excel_col = (26 * InStr(alphabet, left(excel_col, 1))) + (InStr(alphabet, right(excel_col, 1)))
		END IF
	ELSE
		excel_col = CInt(excel_col)
	END IF
END FUNCTION

function check_for_password(are_we_passworded_out)
'--- This function checks to make sure a user is not passworded out. If they are, it allows the user to password back in. NEEDS TO BE ADDED INTO dialog DO...lOOPS
'~~~~~ are_we_passworded_out: When adding to dialog enter "Call check_for_password(are_we_passworded_out)", then Loop until are_we_passworded_out = false. Parameter will remain true if the user still needs to input password.
'===== Keywords: MAXIS, PRISM, password
	Transmit 'transmitting to see if the password screen appears
	Emreadscreen password_check, 8, 2, 33 'checking for the word password which will indicate you are passworded out
	If password_check = "PASSWORD" then 'If the word password is found then it will tell the worker and set the parameter to be true, otherwise it will be set to false.
		Msgbox "Are you passworded out? Press OK and the dialog will reappear. Once it does, you can enter your password."
		are_we_passworded_out = true
	Else
		are_we_passworded_out = false
	End If
end function

function navigate_to_MAXIS_screen(function_to_go_to, command_to_go_to)
'--- This function is to be used to navigate to a specific MAXIS screen
'~~~~~ function_to_go_to: needs to be MAXIS function like "STAT" or "REPT"
'~~~~~ command_to_go_to: needs to be MAXIS function like "WREG" or "ACTV"
'===== Keywords: MAXIS, navigate
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " then
    EMReadScreen locked_panel, 23, 2, 30
    IF locked_panel = "Program History Display" then
	PF3 'Checks to see if on Program History panel - which does not allow the Command line to be updated
    END IF
    row = 1
    col = 1
    EMSearch "Function: ", row, col
    If row <> 0 then
      EMReadScreen MAXIS_function, 4, row, col + 10
      EMReadScreen STAT_note_check, 4, 2, 45
      row = 1
      col = 1
      EMSearch "Case Nbr: ", row, col
      EMReadScreen current_case_number, 8, row, col + 10
      current_case_number = replace(current_case_number, "_", "")
      current_case_number = trim(current_case_number)
    End if
    If current_case_number = MAXIS_case_number and MAXIS_function = ucase(function_to_go_to) and STAT_note_check <> "NOTE" then
      row = 1
      col = 1
      EMSearch "Command: ", row, col
      EMWriteScreen command_to_go_to, row, col + 9
      EMSendKey "<enter>"
      EMWaitReady 0, 0
    Else
      Do
        EMSendKey "<PF3>"
        EMWaitReady 0, 0
        EMReadScreen SELF_check, 4, 2, 50
      Loop until SELF_check = "SELF"
      EMWriteScreen function_to_go_to, 16, 43
      EMWriteScreen "________", 18, 43
      EMWriteScreen MAXIS_case_number, 18, 43
      EMWriteScreen MAXIS_footer_month, 20, 43
      EMWriteScreen MAXIS_footer_year, 20, 46
      EMWriteScreen command_to_go_to, 21, 70
      EMSendKey "<enter>"
      EMWaitReady 0, 0
      EMReadScreen abended_check, 7, 9, 27
      If abended_check = "abended" then
        EMSendKey "<enter>"
        EMWaitReady 0, 0
      End if
	  EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	  If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
    End if
  End if
end function

function write_value_and_transmit(input_value, row, col)
'--- This function writes a specific value and transmits.
'~~~~~ input_value: information to be entered
'~~~~~ row: row to write the input_value
'~~~~~ col: column to write the input_value
'===== Keywords: MAXIS, PRISM, case note, three columns, format
	EMWriteScreen input_value, row, col
	transmit
end function

function script_end_procedure(closing_message)
'--- This function is how all user stats are collected when a script ends.
'~~~~~ closing_message: message to user in a MsgBox that appears once the script is complete. Example: "Success! Your actions are complete."
'===== Keywords: MAXIS, MMIS, PRISM, end, script, statistics, stopscript
	stop_time = timer
	If closing_message <> "" AND left(closing_message, 3) <> "~PT" then MsgBox closing_message '"~PT" forces the message to "pass through", i.e. not create a pop-up, but to continue without further diversion to the database, where it will write a record with the message
	script_run_time = stop_time - start_time
	If is_county_collecting_stats  = True then
		'Getting user name
		Set objNet = CreateObject("WScript.NetWork")
		user_ID = objNet.UserName

		'Setting constants
		Const adOpenStatic = 3
		Const adLockOptimistic = 3

        'Determining if the script was successful
        If closing_message = "" or left(ucase(closing_message), 7) = "SUCCESS" THEN
            SCRIPT_success = -1
        else
            SCRIPT_success = 0
        end if

		'Determines if the value of the MAXIS case number - BULK scripts will not have case number informaiton input into the database
		IF left(name_of_script, 4) = "BULK" then MAXIS_CASE_NUMBER = ""

		'Creating objects for Access
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'Fixing a bug when the script_end_procedure has an apostrophe (this interferes with Access)
		closing_message = replace(closing_message, "'", "")

		'Opening DB
		IF using_SQL_database = TRUE then
    		objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" & stats_database_path & ""
		ELSE
			objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & "" & stats_database_path & ""
		END IF

        'Adds some data for users of the old database, but adds lots more data for users of the new.
        If STATS_enhanced_db = false or STATS_enhanced_db = "" then     'For users of the old db
    		'Opening usage_log and adding a record
    		objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX)" &  _
    		"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & script_run_time & ", '" & closing_message & "')", objConnection, adOpenStatic, adLockOptimistic
		'collecting case numbers counties
		Elseif collect_MAXIS_case_number = true then
			objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX, STATS_COUNTER, STATS_MANUALTIME, STATS_DENOMINATION, WORKER_COUNTY_CODE, SCRIPT_SUCCESS, CASE_NUMBER)" &  _
			"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & abs(script_run_time) & ", '" & closing_message & "', " & abs(STATS_counter) & ", " & abs(STATS_manualtime) & ", '" & STATS_denomination & "', '" & worker_county_code & "', " & SCRIPT_success & ", '" & MAXIS_CASE_NUMBER & "')", objConnection, adOpenStatic, adLockOptimistic
		 'for users of the new db
		Else
            objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX, STATS_COUNTER, STATS_MANUALTIME, STATS_DENOMINATION, WORKER_COUNTY_CODE, SCRIPT_SUCCESS)" &  _
            "VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & abs(script_run_time) & ", '" & closing_message & "', " & abs(STATS_counter) & ", " & abs(STATS_manualtime) & ", '" & STATS_denomination & "', '" & worker_county_code & "', " & SCRIPT_success & ")", objConnection, adOpenStatic, adLockOptimistic
        End if

		'Closing the connection
		objConnection.Close
	End if
	If disable_StopScript = FALSE or disable_StopScript = "" then stopscript
end function

function transmit()
'--- This function sends or hits the transmit key.
 '===== Keywords: MAXIS, MMIS, PRISM, transmit
  EMSendKey "<enter>"
  EMWaitReady 0, 0
end function

function cancel_without_confirmation()
'--- This function ends a script after a user presses cancel. There is no confirmation message box but the end message for statistical information that cancel was pressed.
'===== Keywords: MAXIS, PRISM, MMIS, cancel, script_end_procedure
	If ButtonPressed = 0 then
        script_end_procedure("~PT: user pressed cancel")
        'script_end_procedure text added for statistical purposes. If script was canceled prior to completion, the statistics will reflect this.
        'Left the If...End If in the tier in case we want more stats or error handling, or if we need specialty processing for workflows
    End if
end function

function PF3()
'--- This function sends or hits the PF3 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF3
  EMSendKey "<PF3>"
  EMWaitReady 0, 0
end function

function check_for_MAXIS(end_script)
'--- This function checks to ensure the user is in a MAXIS panel
'~~~~~ end_script: If end_script = TRUE the script will end. If end_script = FALSE, the user will be given the option to cancel the script, or manually navigate to a MAXIS screen.
'===== Keywords: MAXIS, production, script_end_procedure
	Do
		transmit
		EMReadScreen MAXIS_check, 5, 1, 39
		If MAXIS_check <> "MAXIS"  and MAXIS_check <> "AXIS " then
			If end_script = True then
				script_end_procedure("You do not appear to be in MAXIS. You may be passworded out. Please check your MAXIS screen and try again.")
			Else
                BeginDialog Password_dialog, 0, 0, 156, 55, "Password Dialog"
                ButtonGroup ButtonPressed
                OkButton 45, 35, 50, 15
                CancelButton 100, 35, 50, 15
                Text 5, 5, 150, 25, "You have passworded out. Please enter your password, then press OK to continue. Press CANCEL to stop the script. "
                EndDialog
                Do
                    Do
                        dialog Password_dialog
                        cancel_confirmation
                    Loop until ButtonPressed = -1
                    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
                Loop until are_we_passworded_out = false					'loops until user passwords back in
			End if
		End if
	Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "
end function

function find_variable(opening_string, variable_name, length_of_variable)
'--- This function finds a string on a page in BlueZone
'~~~~~ opening_string: string to search for
'~~~~~ variable_name: variable name of the string
'~~~~~ length_of_variable: length of the string
'===== Keywords: MAXIS, MMIS, PRISM, find
  row = 1
  col = 1
  EMSearch opening_string, row, col
  If row <> 0 then EMReadScreen variable_name, length_of_variable, row, col + len(opening_string)
end function

function file_selection_system_dialog(file_selected, file_extension_restriction)
'--- This function allows a user to select a file to be opened in a script
'~~~~~ file_selected: variable for the name of the file
'~~~~~ file_extension_restriction: restricts all other file type besides allowed file type. Example: ".csv" only allows a CSV file to be accessed.
'===== Keywords: MAXIS, MMIS, PRISM, file
	'Creates a Windows Script Host object
	Set wShell=CreateObject("WScript.Shell")

	'This loops until the right file extension is selected. If it isn't specified (= ""), it'll always exit here.
	Do
		'Creates an object which executes the "select a file" dialog, using a Microsoft HTML application (MSHTA.exe), and some handy-dandy HTML.
		Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE ><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")

		'Creates the file_selected variable from the exit
		file_selected = oExec.StdOut.ReadLine

		'If no file is selected the script will stop
		If file_selected = "" then stopscript

		'If the rightmost characters of the file selected don't match what was in the file_extension_restriction argument, it'll tell the user. Otherwise the loop (and function) ends.
		If right(file_selected, len(file_extension_restriction)) <> file_extension_restriction then MsgBox "You've entered an incorrect file type. The allowable file type is: " & file_extension_restriction & "."
	Loop until right(file_selected, len(file_extension_restriction)) = file_extension_restriction
end function

function write_variable_in_CASE_NOTE(variable)
'--- This function writes a variable in CASE note
'~~~~~ variable: information to be entered into CASE note from script/edit box
'===== Keywords: MAXIS, CASE note
	If trim(variable) <> "" THEN
		EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
		noting_col = 3											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
		'The following figures out if we need a new page, or if we need a new case note entirely as well.
		Do
			EMReadScreen character_test, 40, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
            character_test = trim(character_test)
            If character_test <> "" or noting_row >= 18 then

				'If we get to row 18 (which can't be read here), it will go to the next panel (PF8).
				If noting_row >= 18 then
					EMSendKey "<PF8>"
					EMWaitReady 0, 0

                    EMReadScreen check_we_went_to_next_page, 75, 24, 2
                    check_we_went_to_next_page = trim(check_we_went_to_next_page)

					'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
					EMReadScreen end_of_case_note_check, 1, 24, 2
					If end_of_case_note_check = "A" then
						EMSendKey "<PF3>"												'PF3s
						EMWaitReady 0, 0
						EMSendKey "<PF9>"												'PF9s (opens new note)
						EMWaitReady 0, 0
						EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
						EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
						noting_row = 5													'Resets this variable to work in the new locale
                    ElseIf check_we_went_to_next_page = "PLEASE PRESS PF3 TO EXIT OR FILL PAGE BEFORE SCROLLING TO NEXT PAGE" Then
                        noting_row = 4
                        Do
                            EMReadScreen character_test, 40, noting_row, 3 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
                            character_test = trim(character_test)
                            If character_test <> "" then noting_row = noting_row + 1
                        Loop until character_test = ""
                    Else
						noting_row = 4													'Resets this variable to 4 if we did not need a brand new note.
                    End If
                Else
                    noting_row = noting_row + 1
				End if
			End if
		Loop until character_test = ""

		'Splits the contents of the variable into an array of words
		variable_array = split(variable, " ")

		For each word in variable_array

			'If the length of the word would go past col 80 (you can't write to col 80), it will kick it to the next line and indent the length of the bullet
			If len(word) + noting_col > 80 then
				noting_row = noting_row + 1
				noting_col = 3
			End if

			'If the next line is row 18 (you can't write to row 18), it will PF8 to get to the next page
			If noting_row >= 18 then
				EMSendKey "<PF8>"
				EMWaitReady 0, 0

                EMReadScreen check_we_went_to_next_page, 75, 24, 2
                check_we_went_to_next_page = trim(check_we_went_to_next_page)

				'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
				EMReadScreen end_of_case_note_check, 1, 24, 2
				If end_of_case_note_check = "A" then
					EMSendKey "<PF3>"												'PF3s
					EMWaitReady 0, 0
					EMSendKey "<PF9>"												'PF9s (opens new note)
					EMWaitReady 0, 0
					EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
					EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
					noting_row = 5													'Resets this variable to work in the new locale
                ElseIf check_we_went_to_next_page = "PLEASE PRESS PF3 TO EXIT OR FILL PAGE BEFORE SCROLLING TO NEXT PAGE" Then
                    noting_row = 4
                    Do
                        EMReadScreen character_test, 40, noting_row, 3 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
                        character_test = trim(character_test)
                        If character_test <> "" then noting_row = noting_row + 1
                    Loop until character_test = ""
				Else
					noting_row = 4													'Resets this variable to 4 if we did not need a brand new note.
				End if
			End if

			'Writes the word and a space using EMWriteScreen
			EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col

			'Increases noting_col the length of the word + 1 (for the space)
			noting_col = noting_col + (len(word) + 1)
		Next

		'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
		EMSetCursor noting_row + 1, 3
	End if
end function

function cancel_confirmation()
'--- This function asks if the user if they want to cancel. If you say yes, the script will end. If no, the dialog will appear for the user again.
'===== Keywords: MAXIS, PRISM, MMIS, cancel, script_end_procedure
	If ButtonPressed = 0 then
		cancel_confirm = MsgBox("Are you sure you want to cancel the script? Press YES to cancel. Press NO to return to the script.", vbYesNo)
		If cancel_confirm = vbYes then script_end_procedure("~PT: user pressed cancel")
        'script_end_procedure text added for statistical purposes. If script was canceled prior to completion, the statistics will reflect this.
	End if
end function

'----------------------------------------------------------------------------------------------------The script
EMConnect ""

CALL check_for_MAXIS(true)
copy_case_note = FALSE

'Checking to see if script is being started on an already created case note
EMReadScreen case_note_check, 10, 2, 33
EMReadScreen case_note_list, 10, 2, 2
EMReadScreen mode_check, 1, 20, 9

'If the script is started from a case note the script will ask if this is the note the worker wants to copy
If case_note_check = "Case Notes" AND case_note_list = "          " Then
	If mode_check = "D" or mode_check = "E" Then
		use_existing_note = MsgBox("It appears that you are currently in a case note that has already been written." & vbNewLine & "Would you like to copy this case note into other cases?", vbYesNo + vbQuestion, "Is this the case note?")
	End If
End If

'If it is the note the worker wants to copy, the script will create the message array from reading the case note lines'\
If use_existing_note = vbYes Then
	copy_case_note = TRUE 	'Creating a boolean variable for future use if needed
	note_row = 4			'Beginning of the case notes
	Do 						'Read each line
		EMReadScreen note_line, 77, note_row, 3
		If trim(note_line) = "" Then Exit Do		'Any blank line indicates the end of the case note because there can be no blank lines in a note
		message_array = message_array & note_line & "~%~"		'putting the lines together
		note_row = note_row + 1
		If note_row = 18 then 									'End of a single page of the case note
			EMReadScreen next_page, 7, note_row, 3
			If next_page = "More: +" Then 						'This indicates there is another page of the case note
				PF8												'goes to the next line and resets the row to read'\
				note_row = 4
			End If
		End If
	Loop until next_page = "More:  " OR next_page = "       "	'No more pages
	message_array = message_array & "**Processed in bulk script**"	'Adding the last line of the case note, indicating the note was bulk entered
	message_array = split(message_array, "~%~")					'Creates the array
	case_note_header = message_array (0)						'This defines the variables for the dialog boxes to come
	For message_line = 1 to (UBound(message_array) - 2)
		case_note_body = case_note_body & ", " & trim(message_array(message_line))
	Next
	case_note_body = right(case_note_body, (len(case_note_body) - 2))
	worker_signature = message_array (UBound(message_array) - 1)
End If

'>>>>> loading the main dialog <<<<<
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 201, 65, "Case Note from List"
  DropListBox 5, 40, 80, 10, "Manual Entry"+chr(9)+"REPT/ACTV"+chr(9)+"Excel File", run_mode
  ButtonGroup ButtonPressed
    OkButton 90, 40, 50, 15
    CancelButton 140, 40, 50, 15
  Text 10, 10, 185, 25, "Please select a run mode for the script. You can either enter the case numbers manually, from REPT/ACTV, or from an Excel file..."
EndDialog

DIALOG Dialog1
	cancel_without_confirmation
	'>>>>> the script has different ways of building case_number_array
	IF run_mode = "Manual Entry" THEN
		CALL build_manual_entry_dlg(case_number_array, case_note_header, case_note_body, worker_signature)

	ELSEIF run_mode = "REPT/ACTV" THEN
		'script_end_procedure("This mode is not yet supported.")
		CALL find_variable("User: ", worker_number, 7)
        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 231, 130, "Enter worker number and Case Note text..."
          EditBox 145, 10, 65, 15, worker_number
          EditBox 45, 50, 180, 15, case_note_header
          EditBox 30, 70, 190, 15, case_note_body
          EditBox 75, 90, 150, 15, worker_signature
          ButtonGroup ButtonPressed
            OkButton 60, 110, 50, 15
            CancelButton 115, 110, 50, 15
          Text 10, 15, 130, 10, "Please enter the 7-digit worker number:"
          Text 10, 35, 95, 10, "Enter your Case Note text..."
          Text 10, 55, 25, 10, "Header:"
          Text 10, 95, 60, 10, "Worker Signature:"
          Text 10, 75, 20, 10, "Body:"
        EndDialog
		DO
			err_msg = ""
			DIALOG Dialog1
				cancel_without_confirmation
				worker_number = trim(worker_number)
				IF worker_number = "" THEN err_msg = err_msg & vbCr & "* You must enter a worker number."
				IF len(worker_number) <> 7 THEN err_msg = err_msg & vbCr & "* Your worker number must be 7 characters long."
				IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		LOOP UNTIL err_msg = ""

		CALL check_for_MAXIS(false)

		'Checking that case number is blank so as to get a full REPT/ACTV
		CALL find_variable("Case Nbr: ", MAXIS_case_number, 8)
		MAXIS_case_number = replace(MAXIS_case_number, "_", " ")
		MAXIS_case_number = trim(MAXIS_case_number)
		IF MAXIS_case_number <> "" THEN
			back_to_SELF
			EMWriteScreen "________", 18, 43
		END IF
		'Checking that MAXIS is not already in REPT/ACTV so as to get a full REPT/ACTV
		EMReadScreen at_REPT_ACTV, 4, 2, 48
		IF at_REPT_ACTV = "ACTV" THEN back_to_SELF

		CALL navigate_to_MAXIS_screen("REPT", "ACTV")
		CALL write_value_and_transmit(worker_number, 21, 13)
		'Making sure we are at the beginning of REPT/ACTV
		DO
			PF7
			EMReadScreen page_one, 2, 3, 78
			IF isnumeric(page_one) = false then page_one = page_one * 1  'this is future proofing since reading variables keep switching back from numeric and non numeric.
		LOOP UNTIL page_one = 1

		rept_row = 7
		DO
			EMReadScreen MAXIS_case_number, 8, rept_row, 12
			MAXIS_case_number = trim(MAXIS_case_number)
			IF MAXIS_case_number <> "" THEN
				case_number_array = case_number_array & MAXIS_case_number & "~~~"
				rept_row = rept_row + 1
				IF rept_row = 19 THEN
					EMReadScreen next_page_check, 7, 19, 3			'this prevents the script from erroring out if the worker only has one completely full page of cases.
					If next_page_check = "More: +" Then
						rept_row = 7
						PF8
					Else
						Exit Do
					End If
				END IF
			END IF
		LOOP until MAXIS_case_number = ""

	ELSEIF run_mode = "Excel File" THEN
		'Opening the Excel file

		DO
			call file_selection_system_dialog(excel_file_path, ".xlsx")

			Set objExcel = CreateObject("Excel.Application")
			Set objWorkbook = objExcel.Workbooks.Open(excel_file_path)
			objExcel.Visible = True
			objExcel.DisplayAlerts = True

			confirm_file = MsgBox("Is this the correct file? Press YES to continue. Press NO to try again. Press CANCEL to stop the script.", vbYesNoCancel)
			IF confirm_file = vbCancel THEN
				objWorkbook.Close
				objExcel.Quit
				stopscript
			ELSEIF confirm_file = vbNo THEN
				objWorkbook.Close
				objExcel.Quit
			END IF
		LOOP UNTIL confirm_file = vbYes


        '>>>>>DLG for Excel mode<<<<<
        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 256, 135, "Case Note Information"
          EditBox 220, 10, 25, 15, excel_col
          EditBox 65, 30, 40, 15, excel_row
          EditBox 190, 30, 40, 15, end_row
          EditBox 45, 50, 205, 15, case_note_header
          EditBox 35, 70, 215, 15, case_note_body
          EditBox 75, 90, 150, 15, worker_signature
          ButtonGroup ButtonPressed
            OkButton 130, 115, 55, 15
            CancelButton 190, 115, 60, 15
          Text 10, 15, 205, 10, "Please enter the column containing the MAXIS case numbers..."
          Text 10, 35, 50, 10, "Row to start..."
          Text 135, 35, 50, 10, "Row to end..."
          Text 10, 55, 25, 10, "Header:"
          Text 10, 95, 60, 10, "Worker Signature:"
          Text 10, 75, 20, 10, "Body:"
        EndDialog

		'Gathering the information from the user about the fields in Excel to look for.
		DO
			err_msg = ""

			DIALOG Dialog1
				cancel_confirmation
				IF isnumeric(excel_col) = FALSE AND len(excel_col) > 2 THEN
					err_msg = err_msg & vbCr & "* Please do not use such a large column. The script cannot handle it."
				ELSE
					IF (isnumeric(right(excel_col, 1)) = TRUE AND isnumeric(left(excel_col, 1)) = FALSE) OR (isnumeric(right(excel_col, 1)) = FALSE AND isnumeric(left(excel_col, 1)) = TRUE) THEN
						err_msg = err_msg & vbCr & "* Please use a valid Column indicator. " & excel_col & " contains BOTH a letter and a number."
					ELSE
						call convert_excel_letter_to_excel_number(excel_col)
						IF isnumeric(excel_row) = false or isnumeric(end_row) = false THEN err_msg = err_msg & vbCr & "* Please enter the Excel rows as numeric characters."
						IF end_row = "" THEN err_msg = err_msg & vbCr & "* Please enter an end to the search. The script needs to know when to stop searching."
					END IF
				END IF
				IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		LOOP UNTIL err_msg = ""

		CALL check_for_MAXIS(false)
		'Generating a CASE NOTE for each case.
		FOR i = excel_row TO end_row
			IF objExcel.Cells(i, excel_col).Value <> "" THEN
				case_number_array = case_number_array & objExcel.Cells(i, excel_col).Value & "~~~"
			END IF
		NEXT
	END IF

CALL check_for_MAXIS(false)

'The business of sending Case notes
case_number_array = trim(case_number_array)
case_number_array = split(case_number_array, "~~~")

'Formatting case note
If copy_case_note = FALSE Then
	message_array = case_note_header & "~%~" & case_note_body & "~%~" & "---" & "~%~" & worker_signature & "~%~" & "---" & "~%~" & "**Processed in bulk script**"
	message_array = split(message_array, "~%~")
End If

privileged_array = ""

FOR EACH MAXIS_case_number IN case_number_array
	IF MAXIS_case_number <> "" THEN
		CALL navigate_to_MAXIS_screen("CASE", "NOTE")
		'Checking for privileged
		EMReadScreen privileged_case, 40, 24, 2
		IF InStr(privileged_case, "PRIVILEGED") <> 0 THEN
			privileged_array = privileged_array & MAXIS_case_number & "~~~"
		ELSE
			PF9
			'-----Added because the script was only case noting the header, footer and worker_signature on the first case.
			FOR EACH message_part IN message_array
				CALL write_variable_in_CASE_NOTE(message_part)
				STATS_counter = STATS_counter + 1    'adds one instance to the stats counter
			NEXT
		END IF
	END IF
NEXT

IF privileged_array <> "" THEN
	privileged_array = replace(privileged_array, "~~~", vbCr)
	MsgBox "The script could not generate a CASE NOTE for the following cases..." & vbCr & privileged_array
END IF

STATS_counter = STATS_counter - 1  'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success!!")
