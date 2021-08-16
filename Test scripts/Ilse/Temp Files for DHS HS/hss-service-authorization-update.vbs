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
name_of_script = "MISC - HSS SERVICE AUTHORIZATION UPDATE.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 800                      'manual run time in seconds
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

function cancel_confirmation()
'--- This function asks if the user if they want to cancel. If you say yes, the script will end. If no, the dialog will appear for the user again.
'===== Keywords: MAXIS, PRISM, MMIS, cancel, script_end_procedure
	If ButtonPressed = 0 then
		cancel_confirm = MsgBox("Are you sure you want to cancel the script? Press YES to cancel. Press NO to return to the script.", vbYesNo)
		If cancel_confirm = vbYes then script_end_procedure("~PT: user pressed cancel")
        'script_end_procedure text added for statistical purposes. If script was canceled prior to completion, the statistics will reflect this.
	End if
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

function check_for_MMIS(end_script)
'--- This function checks to ensure the user is in a MMIS panel
'~~~~~ end_script: If end_script = TRUE the script will end. If end_script = FALSE, the user will be given the option to cancel the script, or manually navigate to a MMIS screen.
'===== Keywords: MMIS, production, script_end_procedure
	Do
		transmit
		row = 1
		col = 1
		EMSearch "MMIS", row, col
		IF row <> 1 then
			If end_script = True then
				script_end_procedure("You do not appear to be in MMIS. You may be passworded out. Please check your MMIS screen and try again.")
			Else
                Dialog1 = ""
                BeginDialog Dialog1, 0, 0, 216, 55, "MMIS Dialog"
                ButtonGroup ButtonPressed
                OkButton 125, 35, 40, 15
                CancelButton 170, 35, 40, 15
                Text 5, 5, 210, 25, "You do not appear to be in MMIS. You may be passworded out. Please check your MMIS screen and try again, or press CANCEL to exit the script."
                EndDialog
                Do
                    Dialog Dialog1
                    cancel_without_confirmation
                Loop until ButtonPressed = -1
			End if
		End if
	Loop until row = 1
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

function clear_line_of_text(row, start_column)
'--- This function clears out a single line of text
'~~~~~ row: coordinate of row to clear
'~~~~~ start_column: coordinate of column to start clearing
'===== Keywords: MAXIS, PRISM, production, clear
  EMSetCursor row, start_column
  EMSendKey "<EraseEof>"
  EMWaitReady 0, 0
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

Function MMIS_panel_confirmation(panel_name, col)
'--- This function confirms that a user in on the correct MMIS panel.
'~~~~~ panel_name: name of the panel you are confirming as a string in ""
'~~~~~ col: The column which to start reading the panel name. For instance, this is usually 51 or 52 in MMIS.
'===== Keywords: MMIS, navigate, confirm
	Do
		EMReadScreen panel_check, 4, 1, col
		If panel_check <> panel_name then Call write_value_and_transmit(panel_name, 1, 8)
	Loop until panel_check = panel_name
End function

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

function navigate_to_MMIS_region(group_security_selection)
'--- This function is to be used when navigating to MMIS from another function in BlueZone (MAXIS, PRISM, INFOPAC, etc.)
'~~~~~ group_security_selection: region of MMIS to access - programed options are "CTY ELIG STAFF/UPDATE", "GRH UPDATE", "GRH INQUIRY", "MMIS MCRE"
'===== Keywords: MMIS, navigate
	attn
	Do
		EMReadScreen MAI_check, 3, 1, 33
		If MAI_check <> "MAI" then EMWaitReady 1, 1
	Loop until MAI_check = "MAI"

	EMReadScreen mmis_check, 7, 15, 15
	IF mmis_check = "RUNNING" THEN
		EMWriteScreen "10", 2, 15
		transmit
	ELSE
		EMConnect"A"
		attn
		EMReadScreen mmis_check, 7, 15, 15
		IF mmis_check = "RUNNING" THEN
			EMWriteScreen "10", 2, 15
			transmit
		ELSE
			EMConnect"B"
			attn
			EMReadScreen mmis_b_check, 7, 15, 15
			IF mmis_b_check <> "RUNNING" THEN
				script_end_procedure("You do not appear to have MMIS running. This script will now stop. Please make sure you have an active version of MMIS and re-run the script.")
			ELSE
				EMWriteScreen "10", 2, 15
				transmit
			END IF
		END IF
	END IF

	DO
		PF6
		EMReadScreen password_prompt, 38, 2, 23
		IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then
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
		EMReadScreen session_start, 18, 1, 7
	LOOP UNTIL session_start = "SESSION TERMINATED"

	'Getting back in to MMIS and trasmitting past the warning screen (workers should already have accepted the warning when they logged themselves into MMIS the first time, yo.
	EMWriteScreen "MW00", 1, 2
	transmit
	transmit

	group_security_selection = UCASE(group_security_selection)

	EMReadScreen MMIS_menu, 24, 3, 30
	If MMIS_menu <> "GROUP SECURITY SELECTION" Then
		EMReadScreen mmis_group_selection, 4, 1, 65
		EMReadScreen mmis_group_type, 4, 1, 57

		correct_group = FALSE

		Select Case group_security_selection

		Case "CTY ELIG STAFF/UPDATE"
			mmis_group_selection_part = left(mmis_group_selection, 2)

			If mmis_group_selection_part = "C3" Then correct_group = TRUE
			If mmis_group_selection_part = "C4" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the County Eligibility region. The script will now stop.")

            menu_to_enter = "RECIPIENT FILE APPLICATION"

		Case "GRH UPDATE"
			If mmis_group_selection  = "GRHU" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the GRH Update region. The script will now stop.")

            menu_to_enter = "PRIOR AUTHORIZATION   "

		Case "GRH INQUIRY"
			If mmis_group_selection  = "GRHI" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the GRH Inquiry region. The script will now stop.")

            menu_to_enter = "PRIOR AUTHORIZATION   "

		Case "MMIS MCRE"
			If mmis_group_selection  = "EK01" Then correct_group = TRUE
			If mmis_group_selection  = "EKIQ" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the MCRE region. The script will now stop.")

            menu_to_enter = "RECIPIENT FILE APPLICATION"

		End Select

        'Now it finds the recipient file application feature and selects it.
        row = 1
        col = 1
        EMSearch menu_to_enter, row, col
        EMWriteScreen "x", row, col - 3
        transmit

	Else
		Select Case group_security_selection

		Case "CTY ELIG STAFF/UPDATE"
			row = 1
			col = 1
			EMSearch " C3", row, col
			If row <> 0 Then
				EMWriteScreen "x", row, 4
				transmit
			Else
				row = 1
				col = 1
				EMSearch " C4", row, col
				If row <> 0 Then
					EMWriteScreen "x", row, 4
					transmit
				Else
					script_end_procedure("You do not appear to have access to the County Eligibility area of MMIS, this script requires access to this region. The script will now stop.")
				End If
			End If

			'Now it finds the recipient file application feature and selects it.
			row = 1
			col = 1
			EMSearch "RECIPIENT FILE APPLICATION", row, col
			EMWriteScreen "x", row, col - 3
			transmit

		Case "GRH UPDATE"
			row = 1
			col = 1
			EMSearch "GRHU", row, col
			If row <> 0 Then
				EMWriteScreen "x", row, 4
				transmit
			Else
				script_end_procedure("You do not appear to have access to the GRH area of MMIS, this script requires access to this region. The script will now stop.")
			End If

			'Now it finds the pror authorization application feature and selects it.
			row = 1
			col = 1
			EMSearch "PRIOR AUTHORIZATION   ", row, col
			EMWriteScreen "x", row, col - 3
			transmit

		Case "GRH INQUIRY"
			row = 1
			col = 1
			EMSearch "GRHI", row, col
			If row <> 0 Then
				EMWriteScreen "x", row, 4
				transmit
			Else
				script_end_procedure("You do not appear to have access to the GRH Inquiry area of MMIS, this script requires access to this region. The script will now stop.")
			End If

			'Now it finds the pror authorization application feature and selects it.
			row = 1
			col = 1
			EMSearch "PRIOR AUTHORIZATION   ", row, col
			EMWriteScreen "x", row, col - 3
			transmit

		Case "MMIS MCRE"
			row = 1
			col = 1
			EMSearch "EK01", row, col
			If row <> 0 Then
				EMWriteScreen "x", row, 4
				transmit
			Else
				row = 1
				col = 1
				EMSearch "EKIQ", row, col
				If row <> 0 Then
					EMWriteScreen "x", row, 4
					transmit
				Else
					script_end_procedure("You do not appear to have access to the MCRE area of MMIS, this script requires access to this region. The script will now stop.")
				End If
			End If

			'Now it finds the recipient file application feature and selects it.
			row = 1
			col = 1
			EMSearch "RECIPIENT FILE APPLICATION", row, col
			EMWriteScreen "x", row, col - 3
			transmit
		End Select
	End If
end function

function ONLY_create_MAXIS_friendly_date(date_variable)
'--- This function creates a MM DD YY date.
'~~~~~ date_variable: the name of the variable to output
    date_variable = dateadd("d", 0, date_variable)    'janky way to convert to a date, but hey it works.
    var_month     = right("0" & DatePart("m",    date_variable), 2)
    var_day       = right("0" & DatePart("d",    date_variable), 2)
    var_year      = right("0" & DatePart("yyyy", date_variable), 2)
	date_variable = var_month &"/" & var_day & "/" & var_year
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

function PF9()
'--- This function sends or hits the PF9 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF9
  EMSendKey "<PF9>"
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
'END OF FUNCTIONS LIBRARY========================================================================================================================================================================================

'TODO: Once new code is updated in Funclib, remove function and test variable
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
            next
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
    loop until days > UBOUND(dates_array)

    dates_array = ordered_dates
end function

'----------------------------------------------------------------------------------------------------The Script
'CONNECTS TO BlueZone
EMConnect ""
Check_for_MMIS(false)   'checking for, and allowing user to navigate into MMIS.

'----------------------------------Set up code
'Excel columns
const recip_PMI_col         = 1     'Col A
const case_number_col       = 2     'Col B
const HSS_start_col         = 7     'Col G
const HSS_end_col           = 8     'Col H
const SA_number_col         = 9     'Col I
const agreement_start_col   = 10    'Col J
const agreement_end_col     = 11    'Col K
const rate_amt_col          = 13    'Col M
const NPI_number_col        = 15    'Col O
const HS_status_col         = 16    'Col P
const faci_in_col           = 19    'Col Q
const faci_out_col          = 20    'Col R
const impacted_vendor_col   = 21    'Col S
const case_status_col       = 26    'Col Z
const rate_reduction_col    = 27    'Col AA

'User interface dialog - There's just one in this script.
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 481, 90, "HSS SERVICE AUTHORIZATION UPDATE"
  ButtonGroup ButtonPressed
    PushButton 420, 45, 50, 15, "Browse...", select_a_file_button
    OkButton 365, 65, 50, 15
    CancelButton 420, 65, 50, 15
  EditBox 15, 45, 400, 15, file_selection_path
  Text 15, 20, 455, 20, "This script should be used when a list of recipients who have Supplemental Service Rate adjustments in MMIS due to overlapping Housing Stabilization Services (HSS)."
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
ObjExcel.Cells(1, rate_reduction_col).Value = "Rate Reduction Status"   'col 27

'formatting the cells
objExcel.Cells(1, 27).Font.Bold = True		'bold font'
objExcel.Columns(27).ColumnWidth = 120		'sizing the last column

Dim adjustment_array()                        'Delcaring array
ReDim adjustment_array(rr_status_const, 0)     'Resizing the array to size of last const

const recip_PMI_const               = 0         'creating array constants
const case_number_const             = 1
const HSS_start_const               = 2
const HSS_end_const                 = 3
const SA_number_const               = 4
const agreement_start_const         = 5
const agreement_end_const           = 6
const npi_number_const              = 7
const HS_status_const               = 8
const faci_in_const                 = 9
const faci_out_const                = 10
const impacted_vendor_const         = 11
const case_status_const             = 12
const prev_start_const              = 13
const prev_end_const                = 14
const new_start_const               = 15
const new_end_const                 = 16
const excel_row_const               = 17
const MAXIS_note_conf_const         = 18
const MMIS_note_conf_const          = 19
const reduce_rate_const             = 20
const adjustment_start_date_const   = 21
const passed_case_tests_const       = 22
const pmi_count_const               = 23
const rate_amt_const                = 24
const rate_reduction_notes_const    = 25
const rr_status_const               = 26

excel_row = 2
entry_record = 0 'incrementor for the array

Do
    recip_PMI = trim(objExcel.cells(excel_row, recip_PMI_col).Value)
    If recip_PMI = "" then exit do

    SA_number       = trim(objExcel.cells(excel_row, SA_number_col).Value)
    SA_number = right("00000000" & SA_number, 11) 'ensures the variable is 11 digits. Inhibiting erorr

    'Adding recipient information to the array
    ReDim Preserve adjustment_array(rr_status_const, entry_record)	'This resizes the array based on the number of rows in the Excel File'
    adjustment_array(recip_PMI_const            , entry_record) = recip_PMI
    adjustment_array(case_number_const          , entry_record) = trim(objExcel.cells(excel_row, case_number_col).Value)
    adjustment_array(HSS_start_const            , entry_record) = trim(objExcel.cells(excel_row, HSS_start_col).Value)
    adjustment_array(HSS_end_const              , entry_record) = trim(objExcel.cells(excel_row, HSS_end_col).Value)
    adjustment_array(SA_number_const            , entry_record) = SA_number
    adjustment_array(agreement_start_const      , entry_record) = trim(objExcel.cells(excel_row, agreement_start_col).Value)
    adjustment_array(agreement_end_const        , entry_record) = trim(objExcel.cells(excel_row, agreement_end_col).Value)
    adjustment_array(npi_number_const           , entry_record) = trim(objExcel.cells(excel_row, NPI_number_col).Value)
    adjustment_array(HS_status_const            , entry_record) = trim(objExcel.cells(excel_row, HS_status_col).Value)
    adjustment_array(faci_in_const              , entry_record) = trim(objExcel.cells(excel_row, faci_in_col).Value)
    adjustment_array(faci_out_const             , entry_record) = trim(objExcel.cells(excel_row, faci_out_col).Value)
    adjustment_array(impacted_vendor_const      , entry_record) = trim(objExcel.cells(excel_row, impacted_vendor_col).Value)
    adjustment_array(case_status_const          , entry_record) = trim(objExcel.cells(excel_row, case_status_col).Value)
    adjustment_array(rate_amt_const             , entry_record) = trim(objExcel.cells(excel_row, rate_amt_col).Value)
    adjustment_array(excel_row_const            , entry_record) = excel_row
    adjustment_array(passed_case_tests_const    , entry_record) = False 'defaulting to false
    adjustment_array(MAXIS_note_conf_const      , entry_record) = False 'defaulting to false
    adjustment_array(MMIS_note_conf_const       , entry_record) = False 'defaulting to false
    adjustment_array(reduce_rate_const          , entry_record) = False 'defaulting to false
    adjustment_array(rate_reduction_notes_const , entry_record) = trim(objExcel.cells(excel_row, rate_reduction_col).Value)
    entry_record = entry_record + 1			'This increments to the next entry in the array'
    stats_counter = stats_counter + 1
    excel_row = excel_row + 1
    recip_PMI = ""  'Blanking out variables for next loop
    SA_number = ""  'Blanking out variables for next loop
Loop

'----------------------------------------------------------------------------------------------------determine which rows of information are going to have a rate reduction or not.
For item = 0 to Ubound(adjustment_array, 2)
    'Determining which date to use to end/start the agreements. Initial conversion date is 07/01/21. We cannot use a date earlier than this. If a date is earlier than this, the date is 07/01/21.
    'This supports both the initial conversion and ongoing cases.
    If DateDiff("d", #07/01/21#, adjustment_array(HSS_start_const, item)) <= 0 then
        'if HSS start date is a negative/a date before 07/01/21 (past date), then use 07/01/21.
        new_agreement_start_date = #07/01/21#
        Call ONLY_create_MAXIS_friendly_date(new_agreement_start_date)
        adjustment_array(adjustment_start_date_const, item) = new_agreement_start_date
    End if 

    'if this date is a negative then the agreement start date is after the HSS start date. Use the agreement start date instead of HSS start date.
    If DateDiff("d", adjustment_array(agreement_start_const, item), adjustment_array(HSS_start_const, item)) <= 0 then 
        Call ONLY_create_MAXIS_friendly_date(adjustment_array(agreement_start_const, item))
        adjustment_array(adjustment_start_date_const, item) = adjustment_array(agreement_start_const, item)
    Else 
        'Using the HSS start date, changing to friendly of MM/DD/YY date 
        Call ONLY_create_MAXIS_friendly_date(adjustment_array(HSS_start_const, item))
        adjustment_array(adjustment_start_date_const, item) = adjustment_array(HSS_start_const, item)
    End if 
    
    'Finding facility panels that may have ended before the HSS start date
    active_facility = False     'default value
    If (adjustment_array(faci_in_const, item) <> "" and adjustment_array(faci_out_const, item) = "") then
        active_facility = True
    ElseIf adjustment_array(faci_out_const, item) <> "" then
        If DateDiff("d", adjustment_array(faci_out_const, item), adjustment_array(adjustment_start_date_const, item)) <= 0 then
            'Facility end date is NOT before the agreement start date.
            active_facility = True
        End if
    End if

    rate_reduction_status = "Failed Case Test(s): "
    'These are the initial case tests that will fail:
    'Rows with Case Status of “Unable to find MONY/VND2 panel”
    'Rows with Case Status of “Privileged Case. Unable to access.”
    'Row’s that have more than one MAXIS case identified, and HS is not active for the recipient on that case.
    'Row’s that are not identified as an Impacted Vendor (“Yes”)
    'Open-ended facility spans or recipients that have faci panels that close after the HSS start date.
    'Rows that may already be done.
    'Rate costs that are not 15.87
    If (adjustment_array(case_status_const, item) = "" and _
        adjustment_array(rate_reduction_notes_const, item) = "" and _
        adjustment_array(HS_status_const, item) <> "" and _
        adjustment_array(impacted_vendor_const, item) = "Yes" and _
        adjustment_array(rate_amt_const, item) = "15.87" and _
        active_facility = True) then
        adjustment_array(passed_case_tests_const, item) = True
    Else
    'Failure Reasons
        If adjustment_array(HS_status_const, item) = "" then rate_reduction_status = rate_reduction_status & "No HS Status in MAXIS Case. "
        If adjustment_array(impacted_vendor_const, item) = "Yes" and adjustment_array(rate_amt_const, item) <> "15.87" then rate_reduction_status = rate_reduction_status & "Rate is not 15.87, review manually. "
        If adjustment_array(impacted_vendor_const, item) <> "Yes" then rate_reduction_status = rate_reduction_status & "Not an impacted vendor. "
        If active_facility = False then rate_reduction_status = rate_reduction_status & "Not an active facility. "
        If adjustment_array(case_status_const, item) <> "" then rate_reduction_status = rate_reduction_status & adjustment_array(case_status_const, item)
        If adjustment_array(rate_reduction_notes_const, item) <> "" then rate_reduction_status = adjustment_array(rate_reduction_notes_const, item) 'not incrementing this failure reason. Just inputting exiting notes.
    End if
    If rate_reduction_status <> "Failed Case Test(s): " then adjustment_array(rr_status_const, item) = rate_reduction_status
Next

'If duplicates still exist after the intital case tests, then these need to be figured out manually at this point.
For item = 0 to Ubound(adjustment_array, 2)
    recip_PMI = adjustment_array(recip_PMI_const, item)
    PMI_count = 0
    For i = 0 to Ubound(adjustment_array, 2)
        If recip_PMI = adjustment_array(recip_PMI_const, i) then
            If adjustment_array(passed_case_tests_const, i) = True then PMI_count = PMI_count + 1
        End if
    Next

    adjustment_array(pmi_count_const, item) = PMI_count
Next

For item = 0 to Ubound(adjustment_array, 2)
    If adjustment_array(pmi_count_const, item) > 1 then
        If adjustment_array(passed_case_tests_const, item) = True then
            adjustment_array(passed_case_tests_const, item) = False
            adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & "Duplicate agreements found. Review manually."
        End if
    End if

    If adjustment_array(passed_case_tests_const, item) = True then adjustment_array(reduce_rate_const, item) = True    'cases that have passed the cases tests will be initially set to reduce.
    rate_reduction_status = ""  'blanking out variable.
Next

'----------------------------------------------------------------------------------------------------MMIS STEPS
Call check_for_MMIS(False)

For item = 0 to Ubound(adjustment_array, 2)
    If adjustment_array(reduce_rate_const, item) = True then
        'start the rate reductions in MMIS
        Call navigate_to_MMIS_region("GRH UPDATE")	'function to navigate into MMIS, select the GRH update realm, and enter the prior authorization area
        Call MMIS_panel_confirmation("AKEY", 51)				'ensuring we are on the right MMIS screen
        EmWriteScreen "C", 3, 22
        Call write_value_and_transmit(adjustment_array(SA_number_const, item), 9, 36) 'Entering Service Authorization Number and transmit to ASA1
        EmReadscreen current_panel, 4, 1, 51
        If current_panel = "AKEY" then
            error_message = ""
            EmReadscreen error_message, 50, 24, 2
            adjustment_array(reduce_rate_const, item) = False
            adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & "Authorization Number is not valid."
        Else
            EMReadScreen AGMT_STAT, 1, 3, 17
            If AGMT_STAT <> "A" then
                adjustment_array(reduce_rate_const, item) = False
                adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & "Authorization Status is coded as: " & AGMT_STAT & "."
            Else
                'EmWriteScreen "S", 3, 17
                'PF3     'to AKEY screen
                'EmReadscreen current_panel, 4, 1, 51
                ''msgbox current_panel
                'If current_panel <> "AKEY" then
                '    adjustment_array(reduce_rate_const, item) = False
                '    adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & "Unknown issue occured after changeing AGMT STAT on ASA1."
                '    msgbox "Failed! Unknown issue occured after changeing AGMT STAT on ASA1."
                'Else
                    transmit 'to ASA1
                    Call write_value_and_transmit("ASA3", 1, 8)             'Direct navigate to ASA3
                    Call MMIS_panel_confirmation("ASA3", 51)				'ensuring we are on the right MMIS screen

                    'Checking Line 2 to ensure it's blank
                    EmReadscreen line_2_check, 6, 14, 60
                    If trim(line_2_check) <> "" then
                        PF6 'cancel
                        transmit 'to re-enter ASA1
                        EmWriteScreen "A", 3, 17   'Restoring the agreement on ASA1 in AGMT/TYPE STAT field
                        PF3
                        adjustment_array(reduce_rate_const, item) = False
                        adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & "Agreement already exists in Line 2. Review Manually."
                    Else
                        'Reading and converting start and end dates
                        'agreement start date
                        EMReadScreen start_month, 2, 8, 60
                        EMReadScreen start_day, 2, 8, 62
                        EMReadScreen start_year, 2, 8, 64
                        Line_1_start_date = start_month & "/" & start_day & "/" & start_year
                        Call ONLY_create_MAXIS_friendly_date(Line_1_start_date)

                        'For cases that Line 1 agreements are the same day or before the HSS start date.
                        If DateDiff("d", Line_1_start_date, adjustment_array(adjustment_start_date_const, item)) < 0 then
                            'if this date is a negative or a date before 07/01/21 (past date), then use 07/01/21.
                            PF6 'cancel
                            transmit 'to re-enter ASA1
                            EmWriteScreen "A", 3, 17   'Restoring the agreement on ASA1 in AGMT/TYPE STAT field
                            PF3
                            adjustment_array(reduce_rate_const, item) = False
                            adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & "Agreement start date (" & Line_1_start_date & ") is <= HSS start date (" & adjustment_array(adjustment_start_date_const, item) & ")."
                        Else
                            'agreement end date - original end date from line 1
                            EMReadScreen end_month, 2, 8, 67
                            EMReadScreen end_day, 2, 8, 69
                            EMReadScreen end_year, 2, 8, 71
                            original_end_date = end_month & "/" & end_day & "/" & end_year
                            Call ONLY_create_MAXIS_friendly_date(original_end_date)
                            write_original_end_date = replace(original_end_date, "/", "")  'for line 2

                            'Failing cases that the end date is less than the new agreement start date
                            If DateDiff("d", adjustment_array(adjustment_start_date_const, item), original_end_date) <= 0 then
                                'if this date is a positive then its a date before the HSS start date and needs to fail.
                                PF6 'cancel
                                transmit 'to re-enter ASA1
                                EmWriteScreen "A", 3, 17   'Restoring the agreement on ASA1 in AGMT/TYPE STAT field
                                PF3
                                adjustment_array(reduce_rate_const, item) = False
                                adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & "Agreement end date (" & original_end_date & ") is < HSS start date (" & adjustment_array(adjustment_start_date_const, item) & ")."
                            Else
                                'Creating a date that is the day before the HSS start date/conversion date - for LINE 1
                                new_line_1_end_date = dateadd("d", -1, adjustment_array(adjustment_start_date_const, item))
                                'using the HSS start date as this is after 07/01/21 (future date from initial coversion date of 07/01/21)
                                Call ONLY_create_MAXIS_friendly_date(new_line_1_end_date)

                                'removing date formatting for ASA3 input
                                write_new_line_1_end_date = replace(new_line_1_end_date, "/", "")

                                line_1_total_units = datediff("d", Line_1_start_date, new_line_1_end_date) + 1

                                'Unable to close agreements that have been overbilled by the facility.
                                over_billed = True      'Defaulting to True
                                EmReadscreen billed_units, 6, 11, 60
                                billed_units = trim(billed_units)
                                If trim(billed_units) = "" then
                                    over_billed = False   'no billing exists - blank
                                ElseIf cint(billed_units) = cint(line_1_total_units) then
                                    over_billed = False 'facility only billed up to the amount of the date we are closing this agreement date.
                                Elseif cint(billed_units) < cint(line_1_total_units) then
                                    over_billed = False  'facility billed less than the amount of the date we are closing this agreement date.
                                End if

                                If over_billed = True then
                                    PF6 'cancel
                                    transmit 'to re-enter ASA1
                                    EmWriteScreen "A", 3, 17   'Restoring the agreement on ASA1 in AGMT/TYPE STAT field
                                    PF3
                                    adjustment_array(reduce_rate_const, item) = False
                                    adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & " Unable to reduce Line 1 agreement due to overbilling. Billed units: & " & billed_units & " vs. " & line_1_total_units & "."
                                Else
                                    'Deleting the orginal agreement if the start dates are the same date
                                    If DateDiff("d", Line_1_start_date, adjustment_array(adjustment_start_date_const, item)) = 0 then
                                        EmWriteScreen "D", 12, 19 'Deny orginal agreement
                                    Else
                                        '----------------------------------------------------------------------------------------------------Updating LINE 1 agreement
                                        EmWriteScreen write_new_line_1_end_date, 8, 67
                                        Call clear_line_of_text(9, 60)
                                        EmWriteScreen line_1_total_units, 9, 60
                                    End if
                                    '----------------------------------------------------------------------------------------------------Entering LINE 2 Information
                                    EmWriteScreen "H0043", 13, 36
                                    EmWriteScreen "U5", 13, 44

                                    write_new_agrement_start_date = replace(adjustment_array(adjustment_start_date_const, item), "/", "")

                                    EmWriteScreen write_new_agrement_start_date, 14, 60
                                    EmWriteScreen write_original_end_date, 14, 67

                                    EmReadscreen old_rate, 5, 9, 24
                                    new_rate = old_rate / 2 'divide total by two, and round to integer
                                    new_rate = Round(new_rate, 2) 'round to two decimal places
                                    EmWriteScreen new_rate, 15, 20

                                    line_2_total_units = datediff("d", adjustment_array(adjustment_start_date_const, item), original_end_date) + 1
                                    EmWriteScreen line_2_total_units, 15, 60

                                    EMReadscreen agreement_NPI_number, 10, 10, 20   'Reading line 1 NPI Number
                                    EmReadscreen facility_name, 35, 10, 31
                                    EmWriteScreen agreement_NPI_number, 16, 20      'Enetering NPI in Line 2 agreement

                                    EmWriteScreen new_rate, 17, 20
                                    EmWriteScreen "MM", 17, 35

                                    EmWriteScreen "A", 18, 19   'Approving the agreement on ASA3 in STAT CD/DATE field
                                    EmWriteScreen "A", 3, 20   'Approving the agreement on ASA3 in AGMT/TYPE STAT field
                                    transmit

                                    'PF3 ' to save
                                    EMReadScreen PPOP_check, 4, 1, 52
                                    If PPOP_check = "PPOP" then
                                        faci_found = False
                                        'Setting default rows to start
                                        faci_name_row = 5
                                        active_status_row = 8

                                        Do
                                            EmReadscreen faci_name, 35, faci_name_row, 5
                                            If trim(facility_name) = trim(faci_name) then
                                                EmReadscreen provider_type, 18, faci_name_row, 52
                                                EmReadscreen facility_status, 10, active_status_row, 49
                                                If trim(provider_type) = "18 H/COMM PRV" and trim(facility_status) = "ACTIVE" then
                                                    faci_found = True
                                                    Call write_value_and_transmit("X", faci_name_row, 2)    'selecting the found file. Will only select the 1st instance it can find.
                                                    exit do
                                                Else
                                                    faci_name_row = faci_name_row + 4               'incrementing to next facility information section
                                                    active_status_row = active_status_row + 4
                                                    If faci_name_row = 21 then
                                                        PF8                     'Accounting for more than one page of facilities
                                                        faci_name_row = 5       'resetting the rows to the 1st facility set
                                                        active_status_row = 8
                                                        EmReadscreen last_page, 60, 24, 20
                                                    End if
                                                End if

                                            Else
                                                faci_name_row = faci_name_row + 4               'incrementing to next facility information section
                                                active_status_row = active_status_row + 4
                                                If faci_name_row = 21 then
                                                    PF8                     'Accounting for more than one page of facilities
                                                    faci_name_row = 5       'resetting the rows to the 1st facility set
                                                    active_status_row = 8
                                                    EmReadscreen last_page, 60, 24, 20
                                                End if
                                            End if
                                        Loop until trim(last_page) = "CANNOT SCROLL FORWARD - NO MORE DATA TO DISPLAY."

                                        If faci_found = False then
                                            Dialog1 = ""
                                                BeginDialog Dialog1, 0, 0, 181, 130, "PPOP screen - Choose Facility"
                                                ButtonGroup ButtonPressed
                                                  OkButton 65, 105, 50, 15
                                                  CancelButton 120, 105, 50, 15
                                                Text 5, 5, 170, 35, "Please select the correct facility name/address from the list in PPOP by putting a 'X' next to the name. DO NOT TRANSMIT. Press OK when ready. Press CANCEL to stop the script."
                                                Text 5, 45, 175, 20, "* Provider types for GRH must be '18/H COMM PRV' and the status must be '1 ACTIVE.'"
                                                Text 5, 75, 175, 20, "Line 1 Provider Name: " & trim(facility_name)
                                            EndDialog
                                            Do
                                                dialog Dialog1
                                                cancel_confirmation
                                            Loop until ButtonPressed = -1
                                		    EMReadScreen PPOP_check, 4, 1, 52
                                            If PPOP_check = "PPOP" then transmit     'to exit PPOP
                                            If PPOP_check = "SA3 " then transmit    'to navigate to ACF1 - this is the partial screen check for ASA3
                                            transmit ' to next available screen (does not need to be updated)
                                            Call write_value_and_transmit("ACF3", 1, 51)
                                        End if
                                    End if
                                    'saving the agreements
                                    PF3
                                    EmReadscreen current_panel, 4, 1, 51
                                    
                                    If current_panel = "AKEY" then
                                        error_message = ""
                                        EmReadscreen error_message, 50, 24, 2
                                        If trim(error_message) = "ACTION COMPLETED" then
                                            adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & "Agreement successfully reduced to " & new_rate & "."
                                        Else
                                            adjustment_array(reduce_rate_const, item) = False
                                            adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & "Not reduced. MMIS Error: " & trim(error_message)
                                        End if
                                    Else
                                        error_message = ""
                                        EmReadscreen error_message, 80, 21, 2       'reading error message on any other screen.
                                        adjustment_array(reduce_rate_const, item) = False
                                        adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & "Not reduced. MMIS Error: " & trim(error_message)
                                        PF6 'cancel
                                        transmit 'to re-enter ASA1
                                        EmWriteScreen "A", 3, 17   'Restoring the agreement on ASA1 in AGMT/TYPE STAT field
                                        PF3
                                    End if
                                End if
                            End if
                        End if
                    End if
                End if
            'End if
        End if
    End if
Next

write_this_thing = "DHS SUPPLEMENTAL SERVICE RATE ADJUSTMENT" & "~" & "THERE IS AN ACTIVE HOUSING SUPPORT SUPPLEMENTAL SERVICE RATE (SSR)" & "~" & "SERVICE AUTHORIZATION IN MMIS FOR THIS MAXIS CASE. DHS ADJUSTED THE" & "~" &_
				   "MMIS SERVICE AUTHORIZATION(S) FOR HOUSING SUPPORT SSR THROUGH THE" & "~" & "EXISITING END DATE OF THE SERVICE AUTHORIZATION." & "~" & "REVISIONS ARE BASED ON A DETERMINATION OF THE RECIPIENT'S CONCURRENT" & "~" &_
				   "ELIGBILITY HOUSING STABILIZATION SERVICES. MMIS ISSUED A REVISED" & "~" & "SERVICE AUTORIZATION WITH THE CORRECT SSR PER DIEM TO THE HOUSING" & "~" & "SUPPORT PROVIDER ASSOCIATED WITH THE MMIS SERVICE AUTHORIZATION." & "~" &_
				   "ELIGIBILITY WORKERS DO NOT NEED TO TAKE ANY ACTION IN MAXIS." & "~" & "**********************************************************************"
AN_ARRAY_OF_THE_THING_TO_WRITE = split(write_this_thing, "~")
'----------------------------------------------------------------------------------------------------DHS NOTES on ADHS screen in GRHU realm
For item = 0 to Ubound(adjustment_array, 2)
    If adjustment_array(reduce_rate_const, item) = True then
        'start the rate reductions in MMIS
        Call navigate_to_MMIS_region("GRH UPDATE")	'function to navigate into MMIS, select the GRH update realm, and enter the prior authorization area
        Call MMIS_panel_confirmation("AKEY", 51)				'ensuring we are on the right MMIS screen
        EmWriteScreen "C", 3, 22
        Call write_value_and_transmit(adjustment_array(SA_number_const, item), 9, 36) 'Entering Service Authorization Number and transmit to ASA1
        Call MMIS_panel_confirmation("ASA1", 51)				'ensuring we are on the right MMIS screen
        Call write_value_and_transmit("ADHS", 1, 8)
        Call MMIS_panel_confirmation("ADHS", 51)				'ensuring we are on the right MMIS screen
        row = 6
        Do
            EmReadscreen blank_row_check, 6, row, 3
            If trim(blank_row_check) = "" then
                exit do
            Else
                row = row + 1
            End if
        Loop
        
        'Writing in the ADHS - DHS Comments Notes 
		for each comment_line in AN_ARRAY_OF_THE_THING_TO_WRITE
			EmWriteScreen comment_line, row, 3
			row = row + 1
			If row = 14 Then Exit For
		Next
        
        PF3
        error_message = ""
        EmReadscreen error_message, 40, 24, 2
        If trim(error_message) =  "ACTION COMPLETED" then
            adjustment_array(MMIS_note_conf_const, item) = True
        Else
            adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & " Unable to enter note on ADHS - " & trim(error_message)
        End if
    End if
Next

'----------------------------------------------------------------------------------------------------CASE:NOTE - MAXIS
Call navigate_to_MAXIS(maxis_mode)  'Function to navigate back to MAXIS
Call check_for_MAXIS(False)         'Checking to see if we're in MAXIS and/or passworded out.

For item = 0 to Ubound(adjustment_array, 2)
    If adjustment_array(reduce_rate_const, item) = True then
        MAXIS_case_number = adjustment_array(case_number_const, item)
        Call navigate_to_MAXIS_screen_review_PRIV(function_to_go_to, command_to_go_to, is_this_priv)    'Checking for PRIV case note status
        If is_this_priv = False then
            'case note
            Call navigate_to_MAXIS_screen("CASE", "NOTE")
            PF9
            error_message = ""
            EmReadscreen case_note_edit_errors, 70, 3, 3
            EmReadscreen error_message, 50, 24, 2
            If trim(error_message) <> ""  then
                adjustment_array(MAXIS_note_conf_const, item) = False
                adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & " Unable to enter MAXIS CASE:NOTE - " & trim(error_message)
            Elseif trim(case_note_edit_errors) <> "Please enter your note on the lines below:" then
                adjustment_array(MAXIS_note_conf_const, item) = False
                adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & " Unable to edit MAXIS CASE:NOTE - " & trim(error_message)
            Else
                Call write_variable_in_CASE_NOTE("DHS Supplemental Service Rate Adjustment")
                Call write_variable_in_CASE_NOTE("---")
                Call write_variable_in_CASE_NOTE("There is an active Housing Support supplemental service rate (SSR) service authorization in MMIS for this MAXIS case. DHS adjusted the MMIS service authorization(s) for Housing Support SSR through the existing end date of the service authorization.")
                Call write_variable_in_CASE_NOTE("")
                Call write_variable_in_CASE_NOTE("Revisions are based on a determination of the recipient's concurrent eligibility for Housing Stabilization Services. MMIS issued a revised service authorization with the correct SSR per diem to the Housing Support provider associated with the MMIS service authorization.")
                Call write_variable_in_CASE_NOTE("")
                Call write_variable_in_CASE_NOTE("Eligibility workers do not need to take any action in MAXIS.")
                PF3 'to save
                adjustment_array(MAXIS_note_conf_const, item) = True
            End if
        Else
            adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & " Unable to enter MAXIS CASE:NOTE - PRIV Case."
        End if
    End if
Next

'Excel output of rate reduction statuses
For item = 0 to Ubound(adjustment_array, 2)
    objExcel.Cells(adjustment_array(excel_row_const, item), rate_reduction_col).Value = adjustment_array(rr_status_const, item)
Next


'formatting the cells
FOR i = 1 to 27
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT

MAXIS_case_number = ""  'blanking out for statistical purposes. Cannot collect more than one case number.
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure_with_error_report("Success! The script run is complete. Please review the worksheet for reduction statuses and manual updates.")

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
'--All variables are CASE:NOTEing (if required)---------------------------------08/13/2021-----------------No variables, just singular message
'--CASE:NOTE Header doesn't look funky------------------------------------------08/13/2021
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------08/13/2021----------------N/A: Bulk Process
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------08/13/2021
'--MAXIS_background_check reviewed (if applicable)------------------------------08/13/2021----------------N/A: Not updating in MAXIS
'--PRIV Case handling reviewed -------------------------------------------------08/13/2021
'--Out-of-County handling reviewed----------------------------------------------08/13/2021----------------Can make updates in MMIS, MAXIS CASE:NOTES has OOC handling
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