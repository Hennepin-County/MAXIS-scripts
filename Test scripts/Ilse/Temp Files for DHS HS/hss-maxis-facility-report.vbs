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
stats_database_path = "hssqlpw017;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"

'If the "enhanced database" is used (with new features added in January 2016), this variable should be set to true
STATS_enhanced_db = true

'If set to true, the case number will be collected and input into the database
collect_MAXIS_case_number = true

'Required for statistical purposes===============================================================================
name_of_script = "MISC - HSS MAXIS FACILITY REPORT.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 80                      'manual run time in seconds
STATS_denomination = "C"       				'C is for each CASE
'END OF stats block==============================================================================================

'---------------------------------------------------------------------------------------------------
'HOW THIS SCRIPT WORKS:
'
'This script "library" contains functions and variables that the other BlueZone scripts use very commonly. The other BlueZone scripts contain a few lines of code that run
'this script and get the functions. This saves time in writing and copy/pasting the same functions in many different places. Only add functions to this script if they've
'been tested in other scripts first. This document is actively used by live scripts, so it needs to be functionally complete at all times.
'
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

'----------------------------------------------------------------------------------------------------Custom Functions 
function back_to_SELF()
'--- This function will return back to the 'SELF' menu or the MAXIS home menu
'===== Keywords: MAXIS, SELF, navigate
  Do
    EMSendKey "<PF3>"
    EMWaitReady 0, 0
    EMReadScreen SELF_check, 4, 2, 50
  Loop until SELF_check = "SELF"
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

function transmit()
'--- This function sends or hits the transmit key.
 '===== Keywords: MAXIS, MMIS, PRISM, transmit
  EMSendKey "<enter>"
  EMWaitReady 0, 0
end function

function attn()
 '--- This function sends or hits the ESC (escape) key.
  '===== Keywords: MAXIS, MMIS, PRISM, ESC
  EMSendKey "<attn>"
  EMWaitReady -1, 0
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

Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
'--- This function creates a an outlook appointment
'~~~~~ (email_recip): email address for recipeint - seperated by semicolon
'~~~~~ (email_recip_CC): email address for recipeints to cc - seperated by semicolon
'~~~~~ (email_subject): subject of email in quotations or a variable
'~~~~~ (email_body): body of email in quotations or a variable
'~~~~~ (email_attachment): set as "" if no email or file location
'~~~~~ (send_email): set as TRUE or FALSE
'===== Keywords: MAXIS, PRISM, create, outlook, email

	'Setting up the Outlook application
    Set objOutlook = CreateObject("Outlook.Application")
    Set objMail = objOutlook.CreateItem(0)
    If send_email = False then objMail.Display      'To display message only if the script is NOT sending the email for the user.

    'Adds the information to the email
    objMail.to = email_recip                        'email recipient
    objMail.cc = email_recip_CC                     'cc recipient
    objMail.Subject = email_subject                 'email subject
    objMail.Body = email_body                       'email body
    If email_attachment <> "" then objMail.Attachments.Add(email_attachment)       'email attachement (can only support one for now)
    'Sends email
    If send_email = true then objMail.Send	                   'Sends the email
    Set objMail =   Nothing
    Set objOutlook = Nothing
End Function

function excel_open(file_url, visible_status, alerts_status, ObjExcel, objWorkbook)
'--- This function opens a specific excel file.
'~~~~~ file_url: name of the file
'~~~~~ visable_status: set to either TRUE (visible) or FALSE (not-visible)
'~~~~~ alerts_status: set to either TRUE (show alerts) or FALSE (suppress alerts)
'~~~~~ ObjExcel: leave as 'objExcel'
'~~~~~ objWorkbook: leave as 'objWorkbook'
'===== Keywords: MAXIS, PRISM, MMIS, Excel
	Set objExcel = CreateObject("Excel.Application") 'Allows a user to perform functions within Microsoft Excel
	objExcel.Visible = visible_status
	Set objWorkbook = objExcel.Workbooks.Open(file_url) 'Opens an excel file from a specific URL
	objExcel.DisplayAlerts = alerts_status
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

function MAXIS_footer_month_confirmation()
'--- This function is for checking and changing the footer month to the MAXIS_footer_month & MAXIS_footer_year selected by the user in the inital dialog if necessary
'===== Keywords: MAXIS, footer, month, year
	EMReadScreen SELF_check, 4, 2, 50			'Does this to check to see if we're on SELF screen
	IF SELF_check = "SELF" THEN
		EMReadScreen panel_footer_month, 2, 20, 43
		EMReadScreen panel_footer_year, 2, 20, 46
	ELSE
		Call find_variable("Month: ", MAXIS_footer, 5)	'finding footer month and year if not on the SELF screen
		panel_footer_month = left(MAXIS_footer, 2)
		panel_footer_year = right(MAXIS_footer, 2)
		If row <> 0 then
  			panel_footer_month = panel_footer_month		'Establishing variables
			panel_footer_year =panel_footer_year
		END IF
	END IF
	panel_date = panel_footer_month & panel_footer_year		'creating new variable combining month and year for the date listed on the MAXIS panel
	dialog_date = MAXIS_footer_month & MAXIS_footer_year	'creating new variable combining the MAXIS_footer_month & MAXIS_footer_year to measure against the panel date
	IF panel_date <> dialog_date then 						'if dates are not equal
		back_to_SELF
		EMWriteScreen MAXIS_footer_month, 20, 43			'goes back to self and enters the date that the user selcted'
		EMWriteScreen MAXIS_footer_year, 20, 46
	END IF
end function

function navigate_to_MAXIS_screen_review_PRIV(function_to_go_to, command_to_go_to, is_this_priv)
'--- This function is to be used to navigate to a specific MAXIS screen and will check for privileged status
'~~~~~ function_to_go_to: needs to be MAXIS function like "STAT" or "REPT"
'~~~~~ command_to_go_to: needs to be MAXIS function like "WREG" or "ACTV"
'~~~~~ is_this_priv: This returns a true or false based on if the case appears to be privileged in MAXIS
'===== Keywords: MAXIS, navigate
  is_this_priv = FALSE
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " then
    EMReadScreen locked_panel, 23, 2, 30
    IF locked_panel = "Program History Display" then PF3 'Checks to see if on Program History panel - which does not allow the Command line to be updated
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

	  EMReadScreen priv_check, 6, 24, 14              'If it can't get into the case then it will return true as privileged response.
	  If priv_check = "PRIVIL" THEN is_this_priv = TRUE
    End if
  End if
end function

function navigate_to_MAXIS(maxis_mode)
'--- This function is to be used when navigating back to MAXIS from another function in BlueZone (MMIS, PRISM, INFOPAC, etc.)
'~~~~~ maxis_mode: This parameter needs to be "maxis_mode"
'===== Keywords: MAXIS, navigate
    attn
    Do
        EMReadScreen MAI_check, 3, 1, 33
        If MAI_check <> "MAI" then EMWaitReady 1, 1
    Loop until MAI_check = "MAI"

    EMReadScreen prod_check, 7, 6, 15
    IF prod_check = "RUNNING" THEN
        Call write_value_and_transmit("1", 2, 15)
    ELSE
        EMConnect"A"
        attn
        EMReadScreen prod_check, 7, 6, 15
        IF prod_check = "RUNNING" THEN
            Call write_value_and_transmit("1", 2, 15)
        ELSE
            EMConnect"B"
            attn
            EMReadScreen prod_check, 7, 6, 15
            IF prod_check = "RUNNING" THEN
                Call write_value_and_transmit("1", 2, 15)
            Else
                script_end_procedure("You do not appear to have Production mode running. This script will now stop. Please make sure you have production and MMIS open in the same session, and re-run the script.")
            END IF
        END IF
    END IF
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

function PF3()
'--- This function sends or hits the PF3 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF3
  EMSendKey "<PF3>"
  EMWaitReady 0, 0
end function

function PF6()
'--- This function sends or hits the PF6 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF6
  EMSendKey "<PF6>"
  EMWaitReady 0, 0
end function

function PF8()
'--- This function sends or hits the PF8 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF8
  EMSendKey "<PF8>"
  EMWaitReady 0, 0
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

function script_end_procedure_with_error_report(closing_message)
'--- This function is how all user stats are collected when a script ends.
'~~~~~ closing_message: message to user in a MsgBox that appears once the script is complete. Example: "Success! Your actions are complete."
'===== Keywords: MAXIS, MMIS, PRISM, end, script, statistics, stopscript
	stop_time = timer
    send_error_message = ""
	If closing_message <> "" AND left(closing_message, 3) <> "~PT" then        '"~PT" forces the message to "pass through", i.e. not create a pop-up, but to continue without further diversion to the database, where it will write a record with the message
        If testing_run = TRUE Then
            MsgBox(closing_message & vbNewLine & vbNewLine & "Since this script is in testing, please provide feedback")
            send_error_message = vbYes
        Else
            send_error_message = MsgBox(closing_message & vbNewLine & vbNewLine & "Do you need to send an error report about this script run?", vbSystemModal + vbQuestion + vbDefaultButton2 + vbYesNo, "Script Run Completed")
        End If
    End If
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

    If send_error_message = vbYes Then
        'dialog here to gather more detail
        error_type = ""
        If testing_run = TRUE Then error_type = "TESTING RESPONSE"

        If trim(MAXIS_case_number) = "" Then
            If trim(MMIS_case_number) <> "" Then MAXIS_case_number = MMIS_case_number
        End If

        Do
            Do
                confirm_err = ""

                case_note_checkbox = unchecked
                stat_update_checkbox = unchecked
                date_checkbox = unchecked
                math_checkbox = unchecked
                tikl_checkbox = unchecked
                memo_wcom_checkbox = unchecked
                document_checkbox = unchecked
                missing_spot_checkbox = unchecked

                Dialog1 = ""
                BeginDialog Dialog1, 0, 0, 401, 175, "Report Error Detail"
                  Text 60, 35, 55, 10, MAXIS_case_number
                  ComboBox 220, 30, 175, 45, error_type+chr(9)+"BUG - something happened that was wrong"+chr(9)+"ENHANCEMENT - something could be done better"+chr(9)+"TYPO - grammatical/spelling type errors"+chr(9)+"DAIL - add support for this DAIL message.", error_type
                  EditBox 65, 50, 330, 15, error_detail
                  CheckBox 20, 100, 65, 10, "CASE/NOTE", case_note_checkbox
                  CheckBox 95, 100, 65, 10, "Update in STAT", stat_update_checkbox
                  CheckBox 170, 100, 75, 10, "Problems with Dates", date_checkbox
                  CheckBox 265, 100, 65, 10, "Math is incorrect", math_checkbox
                  CheckBox 20, 115, 65, 10, "TIKL is incorrect", tikl_checkbox
                  CheckBox 95, 115, 70, 10, "MEMO or WCOM", memo_wcom_checkbox
                  CheckBox 170, 115, 75, 10, "Created Document", document_checkbox
                  CheckBox 265, 115, 115, 10, "Missing a place for Information", missing_spot_checkbox
                  EditBox 60, 140, 165, 15, worker_signature
                  ButtonGroup ButtonPressed
                    OkButton 290, 140, 50, 15
                    CancelButton 345, 140, 50, 15
                  Text 10, 10, 300, 10, "Information is needed about the error for our scriptwriters to review and resolve the issue. "
                  Text 5, 35, 50, 10, "Case Number:"
                  Text 125, 35, 95, 10, "What type of error occured?"
                  Text 5, 55, 60, 10, "Explain in detail:"
                  GroupBox 10, 75, 380, 60, "Common areas of issue"
                  Text 20, 85, 200, 10, "Check any that were impacted by the error you are reporting."
                  Text 10, 145, 50, 10, "Worker Name:"
                  Text 25, 160, 335, 10, "*** Remember to leave the case as is if possible. We can resolve error better when in a live case. ***"
                EndDialog

                Dialog Dialog1

                If ButtonPressed = 0 Then
                    cancel_confirm_msg = MsgBox("An Error Report will NOT be sent as you pressed 'Cancel'." & vbNewLine & vbNewLine & "Is this what you would like to do?", vbQuestion + vbYesNo, "Confirm Cancel")
                    If cancel_confirm_msg = vbYes Then confirm_err = ""
                    If cancel_confirm_msg = vbNo Then confirm_err = "LOOP" & vbNewLine & confirm_err
                End If

                If ButtonPressed = -1 Then
                    full_text = "Error occurred on " & date & " at " & time
                    full_text = full_text & vbCr & "Error type - " & error_type
                    full_text = full_text & vbCr & "Script name - " & name_of_script & " was run on Case #" & MAXIS_case_number & " with a runtime of " & script_run_time & " seconds."
                    full_text = full_text & vbCr & "Information: " & error_detail
                    If case_note_checkbox = checked OR stat_update_checkbox = checked OR date_checkbox = checked OR math_checkbox = checked OR tikl_checkbox = checked OR memo_wcom_checkbox = checked OR document_checkbox = checked OR missing_spot_checkbox = checked Then full_text = full_text & vbCr & vbCr & "Script has issues/concerns in the following areas:"

                    If case_note_checkbox = checked Then full_text = full_text & vbCr & " - CASE/NOTE"
                    If stat_update_checkbox = checked Then full_text = full_text & vbCr & " - Update in STAT"
                    If date_checkbox = checked Then full_text = full_text & vbCr & " - Dates are incorrect"
                    If math_checkbox = checked Then full_text = full_text & vbCr & " - Math is incorrect"
                    If tikl_checkbox = checked Then full_text = full_text & vbCr & " - TIKL"
                    If memo_wcom_checkbox = checked Then full_text = full_text & vbCr & " - NOTICES (WCOM/MEMO)"
                    If document_checkbox = checked Then full_text = full_text & vbCr & " - The Excel or Word Document"
                    If missing_spot_checkbox = checked Then full_text = full_text & vbCr & " - There is no space to enter particular information"

                    full_text = full_text & vbCr & "Closing message: " & closing_message
                    full_text = full_text & vbCr & vbCr & "Sent by: " & worker_signature

                    send_confirm_msg = MsgBox("** This is what will be sent as an email to the BlueZone Script team:" & vbNewLine & vbNewLine & full_text & vbNewLine & vbNewLine & "*** Is this what you want to send? ***", vbQuestion + vbYesNo, "Confirm Error Report")

                    If send_confirm_msg = vbYes Then confirm_err = ""
                    If send_confirm_msg = vbNo Then confirm_err = "LOOP" & vbNewLine & confirm_err
                End If
            Loop until confirm_err = ""
            call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
        LOOP UNTIL are_we_passworded_out = false
        'sent email here
        full_text = ""
        If ButtonPressed = -1 Then
            bzt_email = "HSPH.EWS.BlueZoneScripts@hennepin.us"
            subject_of_email = "Script Error -- " & name_of_script & " (Automated Report)"

            full_text = "Error occurred on " & date & " at " & time
            full_text = full_text & vbCr & "Error type - " & error_type
            full_text = full_text & vbCr & "Script name - " & name_of_script & " was run on Case #" & MAXIS_case_number & " with a runtime of " & script_run_time & " seconds."
            full_text = full_text & vbCr & "Information: " & error_detail
            If case_note_checkbox = checked OR stat_update_checkbox = checked OR date_checkbox = checked OR math_checkbox = checked OR tikl_checkbox = checked OR memo_wcom_checkbox = checked OR document_checkbox = checked OR missing_spot_checkbox = checked Then full_text = full_text & vbCr & vbCr & "Script has issues/concerns in the following areas:"

            If case_note_checkbox = checked Then full_text = full_text & vbCr & " - CASE/NOTE"
            If stat_update_checkbox = checked Then full_text = full_text & vbCr & " - Update in STAT"
            If date_checkbox = checked Then full_text = full_text & vbCr & " - Dates are incorrect"
            If math_checkbox = checked Then full_text = full_text & vbCr & " - Math is incorrect"
            If tikl_checkbox = checked Then full_text = full_text & vbCr & " - TIKL"
            If memo_wcom_checkbox = checked Then full_text = full_text & vbCr & " - NOTICES (WCOM/MEMO)"
            If document_checkbox = checked Then full_text = full_text & vbCr & " - The Excel or Word Document"
            If missing_spot_checkbox = checked Then full_text = full_text & vbCr & " - There is no space to enter particular information"

            full_text = full_text & vbCr & "Closing message: " & closing_message
            full_text = full_text & vbCr & vbCr & "Sent by: " & worker_signature

            If script_run_lowdown <> "" Then full_text = full_text & vbCr & vbCr & "All Script Run Details:" & vbCr & script_run_lowdown

            Call create_outlook_email(bzt_email, "", subject_of_email, full_text, "", true)

            MsgBox "Error Report completed!" & vbNewLine & vbNewLine & "Thank you for working with us for Continuous Improvement."
        Else
            MsgBox "Your error report has been cancelled and has NOT been sent to the BlueZone Script Team"
        End If
    End If
	If disable_StopScript = FALSE or disable_StopScript = "" then stopscript
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

'END OF FUNCTIONS LIBRARY========================================================================================================================================================================================

function sort_dates(dates_array)
'--- Takes an array of dates and reorders them to be  .
'~~~~~ dates_array: an array of dates only
'===== Keywords: MAXIS, date, order, list, array
    dim ordered_dates ()
    redim ordered_dates(0)
    original_array_items_used = "~"
    days =  0
    do
        prev_date = ""
        original_array_index = 0
        for each thing in dates_array
            check_this_date = TRUE
            new_array_index = 0
            For each known_date in ordered_dates
                if known_date = thing Then check_this_date = FALSE
                new_array_index = new_array_index + 1
                ' MsgBox "known dates is " & known_date & vbNewLine & "thing is " & thing & vbNewLine & "match - " & check_this_date
            next
            ' MsgBox "known dates is " & known_date & vbNewLine & "thing is " & thing & vbNewLine & "check this date - " & check_this_date
            if check_this_date = TRUE Then
                if prev_date = "" Then
                    prev_date = thing
                    index_used = original_array_index
                Else
                    if DateDiff("d", prev_date, thing) < 0 then
                        prev_date = thing
                        index_used = original_array_index
                    end if
                end if
            end if
            original_array_index = original_array_index + 1
        next
        if prev_date <> "" Then
            redim preserve ordered_dates(days)
            ordered_dates(days) = prev_date
            original_array_items_used = original_array_items_used & index_used & "~"
            days = days + 1
        end if
        counter = 0
        For each thing in dates_array
            If InStr(original_array_items_used, "~" & counter & "~") = 0 Then
                For each new_date_thing in ordered_dates
                    If thing = new_date_thing Then
                        original_array_items_used = original_array_items_used & counter & "~"
                        days = days + 1
                    End If
                Next
            End If
            counter = counter + 1
        Next
        ' MsgBox "Ordered Dates array - " & join(ordered_dates, ", ") & vbCR & "days - " & days & vbCR & "Ubound - " & UBOUND(dates_array) & vbCR & "used list - " & original_array_items_used
    loop until days > UBOUND(dates_array)

    dates_array = ordered_dates
end function

'----------------------------------------------------------------------------------------------------The Script 
'CONNECTS TO BlueZone
EMConnect ""
Call check_for_MAXIS(false)

'----------------------------------Set up code 
MAXIS_footer_month = CM_mo 
MAXIS_footer_year = CM_yr 

'Excel columns
const HS_status_col     = 16
const vendor_num_col    = 17
const faci_name_col     = 18
const faci_in_col       = 19
const faci_out_col      = 20
const impact_vnd_col    = 21
const exempt_code_col   = 22
const HDL_one_col       = 23
const HDL_two_col       = 24
const HDL_three_col     = 25
const case_status_col   = 26

'User interface dialog - There's just one in this script. 
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 481, 90, "HSS MAXIS Facility Report"
  ButtonGroup ButtonPressed
    PushButton 420, 45, 50, 15, "Browse...", select_a_file_button
    OkButton 365, 65, 50, 15
    CancelButton 420, 65, 50, 15
  EditBox 15, 45, 400, 15, file_selection_path
  Text 15, 20, 455, 20, "This script should be used when adding MAXIS Facility information to an exisiting spreadsheet with an initial data set provided by DHS for the purposes of possible Supplemental Service Rate reductions due to overlapping Housing Stabilization Services (HSS)."
  Text 30, 70, 335, 10, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 465, 80, "Using this script:"
EndDialog

'Display dialog and dialog DO...Loop for mandatory fields and password prompting  
Do 
    Do
        err_msg = ""
        dialog Dialog1
        cancel_without_confirmation 
        If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
        If trim(file_selection_path) = "" then err_msg = err_msg & vbcr & "* Select a file to continue." 
        If err_msg <> "" Then MsgBox err_msg
    Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file

'Setting up the Excel spreadsheet
ObjExcel.Cells(1, HS_status_col).Value   = date & " MAXIS HS Status"   'col 16
ObjExcel.Cells(1, vendor_num_col).Value  = "Vendor #"                  'col 17
ObjExcel.Cells(1, faci_name_col).Value   = "Facility Name"             'col 18
ObjExcel.Cells(1, faci_in_col).Value     = "Faci In Date"              'col 19
ObjExcel.Cells(1, faci_out_col).Value    = "Faci Out Date"             'col 20
ObjExcel.Cells(1, impact_vnd_col).Value  = "Impacted Vendor?"          'col 21
ObjExcel.Cells(1, exempt_code_col).Value = "VND2 Exemption Code"       'col 22
ObjExcel.Cells(1, HDL_one_col).Value     = "VND2 HDL 1 Code"           'col 23
ObjExcel.Cells(1, HDL_two_col).Value     = "VND2 HDL 2 Code"           'col 24
ObjExcel.Cells(1, HDL_three_col).Value   = "VND2 HDL 3 Code"           'col 25
ObjExcel.Cells(1, case_status_col).Value = "Case Status"               'col 26

FOR i = 16 to 26		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'----------------------------------------------------------------------------------------------------MAXIS DATA GATHER
Call check_for_MAXIS(False)             'Ensuring we're actually in MAXIS 
Call MAXIS_footer_month_confirmation    'Ensuring we're in the right footer month/year: current footer month/year for this process. 

Dim faci_array()                        'Delcaring array
ReDim faci_array(faci_out_const, 0)     'Resizing the array to size of last const 

const vendor_number_const   = 0         'creating array constants
const faci_name_const       = 1
const faci_in_const         = 2
const faci_out_const        = 3

excel_row = 2
Do
    client_PMI = trim(objExcel.cells(excel_row, 1).Value)
    If client_PMI = "" then exit do
    'removing preceeding 0's from the client PMI. This is needed to measure the PMI's on CASE/PERS. 
    Do 
		if left(client_PMI, 1) = "0" then client_PMI = right(client_PMI, len(client_PMI) -1)   'trimming off left-most 0 from client_PMI
	Loop until left(client_PMI, 1) <> "0"                                                      'Looping until 0's are all removed
    client_PMI = trim(client_PMI)
    
	MAXIS_case_number = trim(objExcel.cells(excel_row, 2).Value)
    case_status = ""            'defaulting case_status to "" to increment later in certain circumsatnces
    
    faci_count = 0                          'setting increment for array
    
    '----------------------------------------------------------------------------------------------------CASE/PERS & PERS Search 
    Call navigate_to_MAXIS_screen_review_PRIV("CASE", "PERS", is_this_priv) 
    If is_this_priv = True then 
        case_status = "Privileged Case. Unable to access."
    Else 
        member_found = False 
        Call navigate_to_MAXIS_screen("CASE", "PERS")
        row = 10    'staring row for 1st member 
        Do
            EMReadScreen person_PMI, 8, row, 34
            person_PMI = trim(person_PMI)
            If person_PMI = "" then exit do 
            If trim(person_PMI) = client_PMI then 
                EmReadscreen HS_status, 1, row, 66
                If trim(HS_status) <> "" then
                    EmReadscreen member_number, 2, row, 3 
                    member_found = True 
                    exit do
                End if 
            Else 
                row = row + 3			'information is 3 rows apart. Will read for the next member. 
                If row = 19 then
                    PF8
                    row = 10					'changes MAXIS row if more than one page exists
                END if
            END if
        LOOP 
        If trim(member_number) = "" then case_status = "Unable to locate case for member."
    End if 
    
    If trim(case_status) = "" then    
    '----------------------------------------------------------------------------------------------------FACI panel determination 
	   call navigate_to_MAXIS_screen("STAT", "FACI")
       EmWriteScreen member_number, 20, 76
       Call write_value_and_transmit("01", 20, 79)  'making sure we're on the 1st instance for member 
       'Based on how many FACI panels exist will determine if/how the information is read. 
	    EMReadScreen FACI_total_check, 1, 2, 78
	    If FACI_total_check = "0" then 
	    	case_status = "No FACI panel on this case for member #" & member_number & "."
	    Elseif FACI_total_check = "1" then 
            'just looking through a singular faci panel 
            EmReadscreen faci_name, 30, 6, 43
            faci_name = trim(replace(faci_name, "_", ""))   'faci name trimmed and replaced underscores 
            EmReadscreen vendor_number, 8, 5, 43
            vendor_number = trim(replace(vendor_number, "_", ""))   'vendor # trimmed and replaced underscores 
         
        	row = 18
	    	Do
                EMReadScreen faci_out, 10, row, 71      'faci out date
                If faci_out = "__ __ ____" then 
                    faci_out = ""                       'blanking out faci out if not a date 
                Else 
                    faci_out = replace(faci_out, " ", "/")  'reformatting to output with /, like dates do. 
                End if 
                EMReadScreen faci_in, 10, row, 47       'faci in date
                If faci_in = "__ __ ____" then 
                    faci_in = ""                        'blanking out faci in if not a date 
                Else 
                    faci_in = replace(faci_in, " ", "/")  'reformatting to output with /, like dates do. 
                End if 
	    		If faci_out = "" then 
					If faci_in = "" then  
                        row = row - 1   'no faci info on this row 
                    else 
                        If faci_in <> "" then exit do    'open ended faci found 
                    End if 
	    		Elseif faci_out <> "" then
                    If faci_in <> "" then exit do    'most recent faci span identified 
	    		End if 	
            Loop 
        Else    
            'Evaluate multiple faci panels 
            faci_out_dates_string = ""                  'setting up blank string to increment
            current_faci_found = False                  'defaulting to false - this boolean will determine if evaluation of the last date is needed. Will become true statement if open-ended faci panel is detected.
            For item = 1 to FACI_total_check        
                
                Call write_value_and_transmit("0" & item, 20, 79)   'Entering the item's faci panel via direct navigation field on FACI panel. 
                row = 18
                Do
                    EMReadScreen faci_out, 10, row, 71      'faci out date
                    If faci_out = "__ __ ____" then 
                        faci_out = ""                       'blanking out faci out if not a date 
                    Else 
                        faci_out = replace(faci_out, " ", "/")  'reformatting to output with /, like dates do. 
                    End if 
                    EMReadScreen faci_in, 10, row, 47       'faci in date
                    If faci_in = "__ __ ____" then 
                        faci_in = ""                        'blanking out faci in if not a date 
                    Else 
                        faci_in = replace(faci_in, " ", "/")  'reformatting to output with /, like dates do. 
                    End if 
                    
                    EmReadscreen faci_name, 30, 6, 43
                    faci_name = trim(replace(faci_name, "_", ""))   'faci name trimmed and replaced underscores 
                    EmReadscreen vendor_number, 8, 5, 43
                    vendor_number = trim(replace(vendor_number, "_", ""))   'vendor # trimmed and replaced underscores 
                    'Reading the faci in and out dates 
                    If faci_out = "" then 
                        If faci_in = "" then  
                            row = row - 1   'no faci info on this row - this is blank 
                        else 
                            If faci_in <> "" then 
                                current_faci_found = True   'Condition is met so date evaluation via FACI_array is not needed. 
                                exit do    'open ended faci found 
                            End if 
                        End if 
                    Elseif faci_out <> "" then
                        If faci_in <> "" then 
                            faci_out_dates_string = faci_out_dates_string & faci_out & "|"
                            
                            Redim Preserve faci_array(faci_out_const, faci_count)
                            faci_array(vendor_number_const, faci_count) = vendor_number
                            faci_array(faci_name_const,     faci_count) = faci_name 
                            faci_array(faci_in_const,       faci_count) = faci_in
                            faci_array(faci_out_const,      faci_count) = faci_out 
                            faci_count = faci_count + 1
                            exit do    'most recent faci span identified 
                        End if 
                    End if 	
                Loop 
                If current_faci_found = True then exit for  'exiting the for since most current FACI has been found 
            Next 
            
            'If an open-ended faci is NOT found, then futher evaluation is needed to determine the most recent date. 
            If current_faci_found = False then
                faci_out_dates_string = left(faci_out_dates_string, len(faci_out_dates_string) - 1)
                faci_out_dates = split(faci_out_dates_string, "|")
                call sort_dates(faci_out_dates)
                first_date = faci_out_dates(0)                              'setting the first and last check dates
                last_date = faci_out_dates(UBOUND(faci_out_dates))
                
                'finding the most recent date if none of the dates are open-ended 
                For item = 0 to Ubound(faci_array, 2)
                    If faci_array(faci_out_const, item) = last_date then 
                        vendor_number   = faci_array(vendor_number_const, item)
                        faci_name       = faci_array(faci_name_const, item)
                        faci_in         = faci_array(faci_in_const, item)
                        faci_out        = faci_array(faci_out_const, item)
                    End if 
                Next  
            End if 
            ReDim faci_array(faci_out_const, 0)     'Resizing the array back to original size
            Erase faci_array                        'then once resized it gets erased. 
	    End if 
        
        '----------------------------------------------------------------------------------------------------VNDS/VND2
        Call Navigate_to_MAXIS_screen("MONY", "VNDS")
        Call write_value_and_transmit(vendor_number, 4, 59)
        Call write_value_and_transmit("VND2", 20, 70)
        EMReadScreen VND2_check, 4, 2, 54
        If VND2_check <> "VND2" then 
            case_status = "Unable to find MONY/VND2 panel"
        Else 
            health_depart_reason = False    'defalthing to false 
            exemption_reason = False
            
            EmReadscreen exemption_code, 2, 9, 69
            If exemption_code = "__" then exemption_code = ""
            EmReadscreen HDL_one, 2, 10, 69
            EmReadscreen HDL_two, 2, 10, 72
            EmReadscreen HDL_three, 2, 10, 75
            If HDL_one = "__" then HDL_one = ""
            If HDL_two = "__" then HDL_two = ""
            If HDL_three = "__" then HDL_three = ""
            HDL_string = HDL_one & "|" & HDL_two & "|" & HDL_three
            
            HDL_applicable_codes = "08,09,10"
            If HDL_one <> "" then 
                If instr(HDL_applicable_codes, HDL_one) then health_depart_reason = True 
            End if
            
            If HDL_two <> "" then 
                If instr(HDL_applicable_codes, HDL_two) then health_depart_reason = True 
            End if
            
            If HDL_three <> "" then 
                If instr(HDL_applicable_codes, HDL_three) then health_depart_reason = True 
            End if
            
            If exemption_code = "15" or exemption_code = "26" or exemption_code = "28" then 
                exemption_reason = True 
            Else 
                exmption_reason = False 
            End if 
            
            If exemption_code = "28" and instr(HDL_string, "10") then 
                impacted_vendor = "No"
            Else 
                If (exemption_reason = True and health_depart_reason = True) then 
                    impacted_vendor = "Yes" 
                Else 
                    impacted_vendor = "No"
                End if 
            End if 
        End if
    End if 
    
    'outputting to Excel 
    ObjExcel.Cells(excel_row, HS_status_col).Value   = HS_status
    ObjExcel.Cells(excel_row, vendor_num_col).Value  = vendor_number
    ObjExcel.Cells(excel_row, faci_name_col).Value   = faci_name
    ObjExcel.Cells(excel_row, faci_in_col).Value     = faci_in
    ObjExcel.Cells(excel_row, faci_out_col).Value    = faci_out
    ObjExcel.Cells(excel_row, impact_vnd_col).Value  = impacted_vendor
    ObjExcel.Cells(excel_row, exempt_code_col).Value = exemption_code
    ObjExcel.Cells(excel_row, HDL_one_col).Value     = HDL_one
    ObjExcel.Cells(excel_row, HDL_two_col).Value     = HDL_two
    ObjExcel.Cells(excel_row, HDL_three_col).Value   = HDL_three
    ObjExcel.Cells(excel_row, case_status_col).Value = case_status

    'Blanking out variables at the end of the loop 
    HS_status = ""
    vendor_number = ""
    faci_name = ""
    faci_in = ""
    faci_out = ""
    impacted_vendor = ""
    exemption_code = ""
    HDL_one = ""
    HDL_two = ""
    HDL_three = ""
    case_status = ""
    excel_row = excel_row + 1 'setting up the script to check the next row.
    stats_counter = stats_counter + 1
LOOP UNTIL objExcel.Cells(excel_row, 2).Value = ""	'Loops until there are no more cases in the Excel list

'formatting the cells
FOR i = 1 to 26
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT

MAXIS_case_number = ""  'blanking out for statistical purposes. Cannot collect more than one case number. 

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure_with_error_report("Success! Your facility data has been created.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation 
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------08/13/2021
'--Tab orders reviewed & confirmed----------------------------------------------08/13/2021  
'--Mandatory fields all present & Reviewed--------------------------------------08/13/2021
'--All variables in dialog match mandatory fields-------------------------------08/13/2021
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------08/13/2021-----------------No CASE:NOTE, data only 
'--CASE:NOTE Header doesn't look funky------------------------------------------08/13/2021
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------08/13/2021----------------N/A: Bulk Process
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------08/13/2021
'--MAXIS_background_check reviewed (if applicable)------------------------------08/13/2021----------------N/A: Not updating in MAXIS
'--PRIV Case handling reviewed -------------------------------------------------08/13/2021
'--Out-of-County handling reviewed----------------------------------------------08/13/2021----------------N/A: DHS script 
'--script_end_procedures (w/ or w/o error messaging)----------------------------08/13/2021
'--BULK - review output of statistics and run time/count (if applicable)--------08/13/2021
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------08/13/2021
'--Incrementors reviewed (if necessary)-----------------------------------------08/13/2021
'--Denomination reviewed -------------------------------------------------------08/13/2021
'--Script name reviewed---------------------------------------------------------08/13/2021
'--BULK - remove 1 incrementor at end of script reviewed------------------------08/13/2021

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub taks are complete-----------------------------------------08/13/2021
'--comment Code-----------------------------------------------------------------08/13/2021
'--Update Changelog for release/update------------------------------------------08/13/2021
'--Remove testing message boxes-------------------------------------------------08/13/2021
'--Remove testing code/unnecessary code-----------------------------------------08/13/2021
'--Review/update SharePoint instructions----------------------------------------08/13/2021-------------------N/A: Logic Map provided to DHS
'--Review Best Practices using BZS page ----------------------------------------08/13/2021-------------------N/A: DHS script 
'--Review script information on SharePoint BZ Script List-----------------------08/13/2021-------------------N/A: DHS script
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------08/13/2021-------------------N/A: DHS script
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------08/13/2021-------------------N/A: DHS script
'--Complete misc. documentation (if applicable)---------------------------------08/13/2021
'--Update project team/issue contact (if applicable)----------------------------08/13/2021