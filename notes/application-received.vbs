'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - APPLICATION RECEIVED.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 145                     'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
'END OF stats block=========================================================================================================

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

'DIALOGS-------------------------------------------------------------
BeginDialog case_appld_dialog, 0, 0, 161, 65, "Application Received"
  EditBox 95, 5, 60, 15, MAXIS_case_number
  EditBox 95, 25, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 45, 45, 50, 15
    CancelButton 105, 45, 50, 15
  Text 5, 10, 85, 10, "Enter your case number:"
  Text 5, 30, 85, 10, "Worker Signature"
EndDialog

BeginDialog app_detail_dialog, 0, 0, 221, 280, "Detail of application"
  DropListBox 80, 5, 135, 45, "Select One"+chr(9)+"In Person"+chr(9)+"Dropped Off"+chr(9)+"Mail"+chr(9)+"Online"+chr(9)+"Fax"+chr(9)+"Email", how_app_recvd
  DropListBox 80, 25, 135, 20, "Select One"+chr(9)+"CAF"+chr(9)+"ApplyMN"+chr(9)+"HC - Certain Populations"+chr(9)+"HCAPP"+chr(9)+"Addendum", app_type
  EditBox 80, 45, 135, 15, confirmation_number
  EditBox 80, 65, 135, 15, date_of_app
  CheckBox 5, 105, 30, 10, "Cash", cash_pend
  CheckBox 45, 105, 30, 10, "SNAP", fs_pend
  CheckBox 90, 105, 50, 10, "Emergency", emer_pend
  CheckBox 150, 105, 20, 10, "HC", hc_pend
  CheckBox 185, 105, 30, 10, "GRH", grh_pend
  EditBox 60, 120, 75, 15, time_of_app
  DropListBox 145, 120, 70, 15, "AM"+chr(9)+"PM", AM_PM
  EditBox 50, 140, 165, 15, worker_name
  EditBox 120, 160, 95, 15, worker_number
  EditBox 150, 180, 65, 15, pended_date
  EditBox 5, 200, 210, 15, entered_notes
  CheckBox 5, 220, 205, 15, "Check here to have script transfer case to assigned worker", transfer_case
  EditBox 145, 240, 70, 15, app_in_intake_date
  ButtonGroup ButtonPressed
    OkButton 110, 260, 50, 15
    CancelButton 165, 260, 50, 15
  Text 5, 10, 70, 10, "Application received"
  Text 5, 30, 65, 10, "Type of application"
  Text 5, 50, 60, 10, "Confirmation #"
  Text 5, 70, 65, 10, "Date of Application"
  Text 5, 90, 70, 10, "Programs Applied for:"
  Text 5, 125, 50, 10, "Time received"
  Text 5, 145, 40, 10, "Assigned to:"
  Text 5, 165, 110, 10, "Worker Number (X###### format)"
  Text 5, 185, 25, 10, "Notes:"
  Text 110, 185, 40, 10, "Pended on:"
  Text 25, 245, 115, 10, "Date application received in intake:"
EndDialog

'Grabs the case number
EMConnect ""

'Checks for county info from global variables, or asks if it is not already defined.
get_county_code

CALL MAXIS_case_number_finder (MAXIS_case_number)

'Runs the first dialog - which confirms the case number and gathers worker signature
Do
	Dialog case_appld_dialog
	If buttonpressed = cancel then stopscript
	If MAXIS_case_number = "" then MsgBox "You must have a case number to continue!"
	If worker_signature = "" then Msgbox "Please sign your case note"
Loop until MAXIS_case_number <> "" AND worker_signature <> ""

call check_for_MAXIS(true)

'Gathers Date of application and creates MAXIS friendly dates to be sure to navigate to the correct time frame
'This only functions if case is in PND2 status
call navigate_to_MAXIS_screen("REPT","PND2")
dateofapp_row = 1
dateofapp_col = 1
EMSearch MAXIS_case_number, dateofapp_row, dateofapp_col
EMReadScreen MAXIS_footer_month, 2, dateofapp_row, 38
EMReadScreen app_day, 2, dateofapp_row, 41
EMReadScreen MAXIS_footer_year, 2, dateofapp_row, 44
date_of_app = MAXIS_footer_month & "/" & app_day & "/" & MAXIS_footer_year

'If case is not in PND2 status this defaults the date information to current date to allow correct navigation
If date_of_app = "  /  /  " then
	date_of_app = date
	Call convert_date_into_MAXIS_footer_month (date, MAXIS_footer_month, MAXIS_footer_year)
End If

'Determines which programs are currently pending in the month of application
call navigate_to_MAXIS_screen("STAT","PROG")
EMReadScreen cash1_pend, 4, 6, 74
EMReadScreen cash2_pend, 4, 7, 74
EMReadScreen emer_pend, 4, 8, 74
EMReadScreen grh_pend, 4, 9, 74
EMReadScreen fs_pend, 4, 10, 74
EMReadScreen ive_pend, 4, 11, 74
EMReadScreen hc_pend, 4, 12, 74

'Assigns a value so the programs pending will show up in check boxes
IF cash1_pend = "PEND" THEN
	cash1_pend = 1
	Else
	cash1_pend = 0
End If

If cash2_pend = "PEND" THEN
	cash2_pend = 1
	Else
	cash2_pend = 0
End if

If cash1_pend = 1 OR cash2_pend = 1 then cash_pend = 1

If emer_pend = "PEND" THEN
	emer_pend = 1
	Else
	emer_pend = 0
End if

If grh_pend = "PEND" THEN
	grh_pend = 1
	Else
	grh_pend = 0
End if

If fs_pend = "PEND" THEN
	fs_pend = 1
	Else
	fs_pend = 0
End if

If ive_pend = "PEND" THEN
	ive_pend = 1
	Else
	ive_pend = 0
End if

If hc_pend = "PEND" THEN
	hc_pend = 1
	Else
	hc_pend = 0
End if

'Defaults the date pended to today
pended_date = date & ""

'Runs the second dialog - which gathers information about the application
Do
	Do
		err_msg = ""
		Dialog app_detail_dialog
		cancel_confirmation
		
		If date_of_app = "" Then err_msg = err_msg & vbNewLine & "* Enter the date of application."
		If app_type = "Select One" then err_msg = err_msg & vbNewLine &  "* Please enter the type of application received."
		If how_app_recvd = "Select One" then err_msg = err_msg & vbNewLine &  "* Please enter how the application was received to the agency."
		If worker_name = "" then err_msg = err_msg & vbNewLine &  "* Please enter who this case was assigned to."
		If transfer_case = 1 AND (worker_number = "" OR len(worker_number) <> 7) then err_msg = err_msg & vbNewLine &  "* You must enter the MAXIS number of the worker if you would like the case to be transfered by the script, be sure that it is in X###### format."
		If app_type = "ApplyMN" AND isnumeric(confirmation_number) = false AND time_of_app = "" then err_msg = err_msg & vbNewLine &  "* If an ApplyMN was received, you must enter the confirmation number and time received"
		if err_msg <> "" Then MsgBox "Please resolve before continuing:" & vbNewLine & err_msg
	Loop until err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
Loop until are_we_passworded_out = false

'Creates a variable that lists all the programs pending.
If cash_pend = 1 THEN programs_applied_for = programs_applied_for & "Cash, "
If emer_pend = 1 THEN programs_applied_for = programs_applied_for & "Emergency, "
If grh_pend = 1 THEN programs_applied_for = programs_applied_for & "GRH, "
If fs_pend = 1 THEN programs_applied_for = programs_applied_for & "SNAP, "
If ive_pend = 1 THEN programs_applied_for = programs_applied_for & "IV-E, "
If hc_pend = 1 THEN programs_applied_for = programs_applied_for & "HC"

'Transfers the case to the assigned worker if this was selected in the second dialog box
IF transfer_case = 1 THEN
	CALL navigate_to_MAXIS_screen ("SPEC", "XFER")
	EMWriteScreen "x", 7, 16
	transmit
	PF9
	EMWriteScreen worker_number, 18, 61
	transmit
	EMReadScreen worker_check, 9, 24, 2
	IF worker_check = "SERVICING" THEN
		MsgBox "The correct worker number was not entered, this X-Number is not a valid worker in MAXIS. You will need to transfer the case manually"
		PF10
		transfer_case = unchecked
	Else 
		transfer_system = "MAXIS."
	End If
End If 
	
IF transfer_case = 1 THEN
	If left(county_name, 6) = "Ramsey" Then 
		transfer_in_lincDoc_msg = MsgBox("Would you like to transfer this case in LincDoc?" & vbNewLine & vbNewLine & "Chose 'No' if you are transferring to an assignment queue.", vbYesNo + vbQuestion, "LincDoc Transfer")
		
		If transfer_in_lincDoc_msg = vbYes Then 
			STATS_manualtime = STATS_manualtime + 45

			'Creating the IE window and setting up the view to better navigate
			Set x = CreateObject("WScript.Shell")
			Set IE = CreateObject("InternetExplorer.Application")
			x.AppActivate("InternetExplorer")
			Leave_IE_Open = FALSE
			
			Function IEToFrontV
			'must not use WScript when running within IE 
				Do While Not x.AppActivate("Active Agents - Internet Explorer")
					'wscript.Sleep(1)
					
				Loop
			End Function
			
			'Going to the website for the LincDoc form
			IE.Navigate "https://isveforms2.co.ramsey.mn.us/lincdoc/?ramsey.CaseCreation"
''			IE.Navigate "https://isveformsdev.co.ramsey.mn.us/lincdoc/login/default/ramsey/CaseCreation"
			
			'Function to keep waiting while IE loads - this only works for initial load
			Function WaitForLoad
			Do While (IE.Busy)
				EMWaitReady 0, 10000
			Loop
			End Function
			
			Call WaitForLoad
			unique_hwnd = IE.hwnd
			
			'view settings - Dimensions are not vital but the others ARE
			IE.Toolbar = 0
			IE.StatusBar = 0
			IE.Height = 750
			IE.Width = 900
			IE.Top = 50
			IE.Left = 50
			IE.Visible = True
			
			tab_end = 4
			
			Do
				'Need to read the webpage title to determine if worker is logged in
				where_am_i = IE.LocationName
				
				'Script will pause and allow worker to log in if on the login screen
				IF right(where_am_i, 5) = "Login" then 
					MsgBox "It appears you are not logged in to LincDoc. Please enter your password to continue."
					tab_end = 4
					Leave_IE_Open = TRUE 
				End If 	
			Loop until right(where_am_i, 5) <> "Login"

			Set oShell = CreateObject("Shell.Application")
			For Each Wnd in oShell.Windows
				If Wnd.hwnd = unique_hwnd Then Set IE = Wnd
			Next
			
			HEADER = IE.LocationName

			x.AppActivate(HEADER)
			
			EMWaitReady 0, 250	'Wait
			x.SendKeys "{Tab}"	'Establish tabbing
			
			EMWaitReady 0, 250	'Wait 
			x.SendKeys "{F3}"	'Open search window
			
			EMWaitReady 0, 250	'Wait 
			x.SendKeys chr(35) 	'type '#'
			
			EMWaitReady 0, 250	'Wait 
			x.SendKeys "{ESC}"	'Press escape to close search window '#' will be highlighted
			
			EMWaitReady 0, 250	'Wait 
			x.SendKeys "{Tab}"	'Tab one time

			EMWaitReady 0, 250	'Wait 
			'entering the case number
			MAXIS_case_number = right (("00000000" & MAXIS_case_number), 8)
		''	x.SendKeys MAXIS_case_number
			IE.Document.getElementsByTagName("input").item(7).value = MAXIS_case_number
			EMWaitReady 0, 250	'Wait
			x.SendKeys "{Tab}"	'Tab one time - now 'search' should be highlighted

			EMWaitReady 0, 250	'Wait 
			x.SendKeys "{ENTER}"'Enter to 'search'
			
			EMWaitReady 0, 3500	'Wait for 3 seconds
			x.AppActivate(HEADER)
			
			'If the case number has missing CCIs or the worker is not known an error message pops up that the script can read.
			'This indicates the transfer in lincdoc has failed
			On Error Resume Next 
			i = 0
			Do 	
				x.AppActivate(HEADER)	'Focus on Internet Explorer
				if instr(UCase(IE.Document.getElementsByTagName("button").item(i).innerhtml),"OK")<>0 then
					If Err.Number = 0 Then 
''						EMWaitReady 0, 1000	'wait for 1 second
						IE.Document.getElementsByTagName("button").item(i).click
						end_msg = "The script was unable to transfer the case in LincDoc for case number " & MAXIS_case_number & vbNewLine & "Please review LincDoc form and process the LF transfer manually."
						linc_doc_transfer = FALSE
						Exit Do
					End If 
				end if
				i = i + 1
			Loop until i = 10
			If linc_doc_transfer = "" Then linc_doc_transfer = TRUE 
			'MsgBox "Transfer Boolean - " & linc_doc_transfer 
			If linc_doc_transfer = TRUE Then 
				'finds and pushes the submit button
				i = 0
				Do 	
					EMWaitReady 0, 150
					x.AppActivate(HEADER)	'Focus on Internet Explorer
					if instr(UCase(IE.Document.getElementsByTagName("button").item(i).innerhtml),"SUBMIT")<>0 then
						IE.Document.getElementsByTagName("button").item(i).click
						submit_clicked = true
						exit do
					end if
					i = i + 1
				Loop until i = 50
				
				EMWaitReady 0, 3000	'Wait for 6 seconds
				
				If submit_clicked = true Then 
					linc_doc_transfer = TRUE 
					transfer_system = "MAXIS & LincDoc."
					end_msg = "Case " & MAXIS_case_number & " transferred in MAXIS and LincDoc."
				Else 
					linc_doc_transfer - False 
				End If 
				
			End If 
		End If 
	End If 
End If

IF time_of_app <> "" Then
	If AM_PM <> "PM" Then 
		colon_place = InStr(time_of_app, ":")
		If colon_place <> 0 Then 
			time_stamp_hour = left(time_of_app, colon_place - 1)
			time_stamp_hour = time_stamp_hour * 1
			If time_stamp_hour > 12 Then 
				time_stamp_hour = time_stamp_hour - 12
				AM_PM = "PM"
			Else 
				AM_PM = "AM"
			End If
			time_stamp_min = right(time_of_app, len(time_of_app) - colon_place)
			time_of_app = time_stamp_hour & ":" & time_stamp_min
		End If 
	End If 
	time_stamp = " at " & time_of_app & " " & AM_PM
ELSE
	time_stamp = " "
End If

app_day_30 = DateAdd("d", 30, date_of_app)

'Writes the case note
CALL start_a_blank_case_note
CALL write_variable_in_CASE_NOTE ("APP PENDED - " & app_type & " rec'vd via " & how_app_recvd & " on " & date_of_app & time_stamp)
IF isnumeric(confirmation_number) = true THEN CALL write_variable_in_CASE_NOTE ("* Confirmation # " & confirmation_number)
CALL write_bullet_and_variable_in_CASE_NOTE ("Requesting", programs_applied_for)
CALL write_bullet_and_variable_in_CASE_NOTE ("Pended on", pended_date)
CALL write_bullet_and_variable_in_CASE_NOTE ("Application assigned to", worker_name)
IF transfer_case = checked THEN CALL write_variable_in_CASE_NOTE ("* Case transfered to " & worker_name & " in " & transfer_system)
IF entered_notes <> "" THEN CALL write_bullet_and_variable_in_CASE_NOTE ("Notes", entered_notes)
CALL write_bullet_and_variable_in_CASE_NOTE ("Day 30", app_day_30)
CALL write_variable_in_CASE_NOTE ("---")
If app_in_intake_date <> "" Then 
	CALL write_bullet_and_variable_in_CASE_NOTE ("Application Received in Intake on", app_in_intake_date)
	CALL write_variable_in_CASE_NOTE ("* (Used for internal tracking only)")
	CALL write_variable_in_CASE_NOTE ("---")
End If 
CALL write_variable_in_CASE_NOTE (worker_signature)

'Reminder to screen for XFS if SNAP is pending.
IF fs_pend = 1 THEN end_msg = "SNAP is pending, be sure to run the NOTES-Expedited Screening script as well to note potential XFS eligibility" & vbNewLine & vbNewLine & end_msg

script_end_procedure (end_msg)