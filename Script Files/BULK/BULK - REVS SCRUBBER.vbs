'OPTION EXPLICIT

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - REVS SCRUBBER.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" OR default_directory = "" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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

'Declaring variables
'DIM month, use_date, num_of_days, next_month, month_array, month_to_use
'DIM first_appointment_listbox, olAppointmentItem, olRecursDaily
'DIM appointment_length_listbox, last_appointment_listbox, time_array_30_min
'DIM alt_appointment_length_listbox, alt_first_appointment_listbox
'DIM last_day_of_recert, cm_plus_1, cm_plus_2, alt_duplicate_appt_times
'DIM worker_numberm, calendar_dlgm, ButtonPressed, duplicate_appt_times 
'DIM alt_appointments_per_time_slot, contact_phone_number, worker_signature
'DIM err_msg, day_of_month, time_array, outlook_calendar_check, worker_number
'DIM REVS_scrubber_time_dialog, REVS_scrubber_initial_dialog
'DIM objWorkbook, objOutlook, objRange, appointments_per_time_slot
'DIM calendar_month, worker_county_code, objAppointment, objExcel
'DIM appt_date, appt_category, appt_body, appt_month, developer_mode
'DIM appt_minute_place_holder_because_reasons, appt_location, appt_end_time
'DIM appt_time_list, appt_subject, appt_start_time, appt_reminder
'DIM current_year, current_worker, current_month, y, i, x, appt_year
'DIM MAXIS_row, case_number, excel_row, excel, revs_year, revs_month
'DIM alt_last_appointment_listbox, calendar_dlg, SNAP_popup_check, SNAP_status
'DIM cash_status, last_page_check, HC_status, add_case_info_to_Excel, CSR_mo
'DIM CSR_yr, recert_mo, recert_status, recert_yr, appointment_time
'DIM appointment_time_for_comparison, appointment_time_for_viewing, appointments, j
'DIM last_appointment_listbox_for_comparison, last_appointment_time, interview_time
'DIM am_pm, recert_status, forms_to_swkr


'Required variables/arrays
appt_time_list = "15 mins"+chr(9)+"30 mins"+chr(9)+"45 mins"+chr(9)+"60 mins"
call convert_array_to_droplist_items(time_array_30_min, time_array)

'Custom functions (should merge with FuncLib when tested/confirmed to work)---------------------------------------------------

'Function to create dynamic calendar out of checkboxes
FUNCTION create_calendar(month_to_use, month_array)
	'Generating a calendar
	'Determining the number of days in the calendar month.
	next_month = DateAdd("M", 1, month_to_use)
	next_month = DatePart("M", next_month) & "/01/" & DatePart("YYYY", next_month)
	num_of_days = DatePart("D", (DateAdd("D", -1, next_month)))

	ReDim month_array(num_of_days, 0)
	
	'=====Another dialog=====
	BeginDialog calendar_dlg, 0, 0, 280, 190, month_to_use
		Text 5, 10, 270, 25, "Please check the days to schedule appointments. You cannot schedule appointments prior to the 8th."
		Text 5, 35, 270, 25, "Please note that Auto-Close Notices are sent on the 16th. To reduce confusion, you may want to schedule before the 16th."
		Text 5, 65, 265, 10, ("Check appointment dates in " & MonthName(DatePart("M", month_to_use)) & ", " & DatePart("YYYY", month_to_use) & ", for " & MonthName(DatePart("M", next_month)) & ", " & DatePart("YYYY", next_month) & ", recertifications.")
		y = 85
		FOR i = 1 TO num_of_days
			use_date = (DatePart("M", month_to_use) & "/" & i & "/" & DatePart("YYYY", month_to_use))
			x = 15 + (40 * (WeekDay(use_date) - 1))
			IF WeekDay(use_date) = 1 AND i <> 1 THEN y = y + 15
			IF WeekDay(use_date) = 1 OR WeekDay(use_date) = 7 THEN
				month_array(i, 0) = 0 
			Else
				month_array(i, 0) = 1
			End If
			IF i < 8 THEN 
				Text x, y, 30, 10, " x " & i
				month_array(i, 0) = o 'unchecking so selections cannot be made for DD 01-07
			ELSE
				CheckBox x, y, 35, 10, i, month_array(i, 0)
			END IF
		NEXT
		ButtonGroup ButtonPressed
		OkButton 175, 170, 50, 15
		CancelButton 225, 170, 50, 15
	EndDialog
	
	Dialog calendar_dlg
		IF ButtonPressed = 0 THEN stopscript
END FUNCTION

FUNCTION create_outlook_appointment(appt_date, appt_start_time, appt_end_time, appt_subject, appt_body, appt_location, appt_reminder, appt_category)
	'Assigning needed numbers as variables for readability
	olAppointmentItem = 1
	olRecursDaily = 0
	
	'Creating an Outlook object item
	Set objOutlook = CreateObject("Outlook.Application")
	Set objAppointment = objOutlook.CreateItem(olAppointmentItem)
	
	'Assigning individual appointment options
	objAppointment.Start = appt_date & " " & appt_start_time		'Start date and time are carried over from parameters
	objAppointment.End = appt_date & " " & appt_end_time			'End date and time are carried over from parameters
	objAppointment.AllDayEvent = False 								'Defaulting to false for this. Perhaps someday this can be true. Who knows.
	objAppointment.Subject = appt_subject							'Defining the subject from parameters
	objAppointment.Body = appt_body									'Defining the body from parameters
	objAppointment.Location = appt_location							'Defining the location from parameters
	If appt_reminder = FALSE then									'If the reminder parameter is false, it skips the reminder, otherwise it sets it to match the number here.
		objAppointment.ReminderSet = False
	Else
		objAppointment.ReminderMinutesBeforeStart = appt_reminder
		objAppointment.ReminderSet = True
	End if
	objAppointment.Categories = appt_category						'Defines a category
	objAppointment.Save												'Saves the appointment

END FUNCTION

'DIALOGS -----------------------------------------------------------------------------------------------

BeginDialog REVS_scrubber_initial_dialog, 0, 0, 136, 130, "REVS scrubber initial dialog"
  EditBox 65, 5, 65, 15, worker_number
  EditBox 65, 25, 65, 15, worker_signature
  EditBox 70, 45, 60, 15, contact_phone_number
  ButtonGroup ButtonPressed
    OkButton 25, 110, 50, 15
    CancelButton 80, 110, 50, 15
  Text 5, 10, 55, 10, "Worker number:"
  Text 5, 30, 60, 10, "Worker signature:"
  Text 5, 45, 60, 60, "Please enter a phone number client can call to report a change in phone number (Include area code)"
EndDialog

BeginDialog REVS_scrubber_time_dialog, 0, 0, 286, 280, "REVS Scrubber Time Dialog"
  DropListBox 75, 15, 60, 15, "Select one..."+chr(9)+time_array, first_appointment_listbox
  DropListBox 210, 15, 60, 15, "Select one..."+chr(9)+time_array, last_appointment_listbox
  DropListBox 115, 35, 50, 15, "Select one..."+chr(9)+appt_time_list, appointment_length_listbox
  CheckBox 10, 55, 135, 10, "Duplicate appointments per time slot?", duplicate_appt_times
  EditBox 110, 70, 35, 15, appointments_per_time_slot
  DropListBox 75, 135, 60, 15, "Select one..."+chr(9)+time_array, alt_first_appointment_listbox
  DropListBox 210, 135, 60, 15, "Select one..."+chr(9)+time_array, alt_last_appointment_listbox
  DropListBox 115, 155, 50, 15, "Select one..."+chr(9)+appt_time_list, alt_appointment_length_listbox
  CheckBox 10, 175, 135, 10, "Duplicate appointments per time slot?", alt_duplicate_appt_times
  EditBox 110, 190, 35, 15, alt_appointments_per_time_slot
  CheckBox 10, 235, 200, 10, "Check here to add appointments to your Outlook calendar.", outlook_calendar_check
  ButtonGroup ButtonPressed
    OkButton 180, 260, 50, 15
    CancelButton 230, 260, 50, 15
  Text 10, 20, 60, 10, "First appointment:"
  Text 145, 20, 60, 10, "Last appointment:"
  Text 10, 35, 95, 10, "Time between Appointments:"
  Text 10, 140, 60, 10, "First appointment:"
  Text 145, 140, 60, 10, "Last appointment:"
  Text 10, 155, 95, 10, "Time between Appointments:"
  Text 15, 75, 90, 10, "Appointments per time slot:"
  GroupBox 5, 5, 275, 85, "Main Appointment Block"
  GroupBox 5, 105, 275, 110, "Additional Appointment Block"
  Text 10, 120, 260, 10, "*NOTE: Use this block for scheduling appointments around your lunch break."
  Text 15, 195, 90, 10, "Appointments per time slot:"
EndDialog

'-----THE SCRIPT, dawg
'Connects to BlueZone
EMConnect ""

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add() 
objExcel.DisplayAlerts = True

'formatting excel file with columns for case number and interview date/time
objExcel.cells(1, 1).Value = "CASE NUMBER"
objExcel.Cells(1, 1).Font.Bold = TRUE
objExcel.Cells(1, 2).Value = "Interview Date & Time"
objExcel.cells(1, 2).Font.Bold = TRUE
objExcel.cells(1, 5).Value = "Privileged Cases"
objExcel.cells(1, 5).Font.Bold = TRUE

'creating month plus 1 and plus 2
cm_plus_1 = dateadd("M", 1, date)
cm_plus_2 = dateadd("M", 2, date)
'creating a last day of recert variable
last_day_of_recert = DatePart("M", cm_plus_2) & "/01/" & DatePart("YYYY", cm_plus_2)
last_day_of_recert = dateadd("D", -1, last_day_of_recert)

'Grabbing the worker's X number.
CALL find_variable("User: ", worker_number, 7)

'if user is Hennepin Co. user, then contact phone number auto-fills with EZ info number
if worker_county_code = "x127" Then contact_phone_number = "612-596-1300"

'Display REVS scrubber initial dialog
DIALOG REVS_scrubber_initial_dialog
IF ButtonPressed = 0 THEN stopscript

'Entering developer mode
If contact_phone_number = "UUDDLRLRBA" then 
	developer_mode = true
	MsgBox "You have enabled Developer Mode." & vbCr & vbCr & "The script will not enter information into MAXIS, but it will navigate, showing you where the script would otherwise have been."
END IF

'Stopping the script is the user is running it before the 16th of the month.
day_of_month = DatePart("D", date)
If developer_mode <> true then
	IF day_of_month < 16 THEN script_end_procedure("You cannot run this script before the 16th of the month.") 'to boot the user before the script tries to access a blank REPT/REVS.
End IF

'Formatting the dates
calendar_month = DateAdd("M", 1, date)
appt_month = DatePart("M", calendar_month)
appt_year = DatePart("YYYY", calendar_month)
next_month = DateAdd("M", 1, calendar_month)
next_month = DatePart("M", next_month) & "/01/" & DatePart("YYYY", next_month)
num_of_days = DatePart("D", (DateAdd("D", -1, next_month)))

'Generating the calendar
ReDim month_array(num_of_days, 0)
CALL create_calendar(calendar_month, month_array)

'Determining the appropriate times to set appointments.
DO		
	err_msg = ""
	DIALOG REVS_scrubber_time_dialog
	IF ButtonPressed = 0 THEN stopscript
	IF first_appointment_listbox = "Select one..." THEN err_msg = err_msg & VbCr & "* You must choose an initial appointment time."				
	IF first_appointment_listbox <> "Select one..." AND last_appointment_listbox <> "Select one..." THEN 
		'Converting the appointment times for comparison. VBS runs in military time.
		IF DatePart("H", last_appointment_listbox) < 7 THEN 
			last_appointment_listbox = DateAdd("H", 12, last_appointment_listbox)
			first_appointment_listbox = DateAdd("H", 0, first_appointment_listbox)
		END IF
		IF DatePart("H", first_appointment_listbox) < 7 THEN 
			first_appointment_listbox = DateAdd("H", 12, first_appointment_listbox)
			last_appointment_listbox = DateAdd("H", 0, last_appointment_listbox)
		END IF
		
		IF DateDiff("N", first_appointment_listbox, last_appointment_listbox) < 0 THEN err_msg = err_msg & VbCr & "* The last appointment may not be earlier than the first appointment."
		
		'Converting the appointment times back from military time.
		IF DatePart("H", last_appointment_listbox) > 12 THEN 
			last_appointment_listbox = DateAdd("H", -12, last_appointment_listbox)
			first_appointment_listbox = DateAdd("H", 0, first_appointment_listbox)
		END IF
		IF DatePart("H", first_appointment_listbox) > 12 THEN 
			first_appointment_listbox = DateAdd("H", -12, first_appointment_listbox)
			last_appointment_listbox = DateAdd("H", 0, last_appointment_listbox)
		END IF
	END IF
	IF alt_first_appointment_listbox <> "Select one..." AND alt_last_appointment_listbox <> "Select one..." THEN
		'Converting the appointment times for comparison. VBS runs in military time.
		IF DatePart("H", alt_last_appointment_listbox) < 7 THEN 
			alt_last_appointment_listbox = DateAdd("H", 12, alt_last_appointment_listbox)
			alt_first_appointment_listbox = DateAdd("H", 0, alt_first_appointment_listbox)
		END IF
		IF DatePart("H", alt_first_appointment_listbox) < 7 THEN 
			alt_first_appointment_listbox = DateAdd("H", 12, alt_first_appointment_listbox)
			alt_last_appointment_listbox = DateAdd("H", 0, alt_last_appointment_listbox)
		END IF
		
		IF DateDiff("N", alt_first_appointment_listbox, alt_last_appointment_listbox) < 0 THEN err_msg = err_msg & VbCr & "* The additional appointment block has an ending earlier than it begins."
		
		'Converting the appointment times back from military time.
		IF DatePart("H", alt_last_appointment_listbox) > 12 THEN 
			alt_last_appointment_listbox = DateAdd("H", -12, alt_last_appointment_listbox)
			alt_first_appointment_listbox = DateAdd("H", 0, alt_first_appointment_listbox)
		END IF
		IF DatePart("H", alt_first_appointment_listbox) > 12 THEN 
			alt_first_appointment_listbox = DateAdd("H", -12, alt_first_appointment_listbox)
			alt_last_appointment_listbox = DateAdd("H", 0, alt_last_appointment_listbox)
		END IF
	END IF
	IF last_appointment_listbox <> "Select one..." AND alt_first_appointment_listbox <> "Select one..." THEN
		'Converting the appointment times for comparison. VBS runs in military time.
		IF DatePart("H", last_appointment_listbox) < 7 THEN 
			last_appointment_listbox = DateAdd("H", 12, last_appointment_listbox)
			alt_first_appointment_listbox = DateAdd("H", 0, alt_first_appointment_listbox)
		END IF
		IF DatePart("H", alt_first_appointment_listbox) < 7 THEN 
			alt_first_appointment_listbox = DateAdd("H", 12, alt_first_appointment_listbox)
			last_appointment_listbox = DateAdd("H", 0, last_appointment_listbox)
		END IF
		'IF DateDiff("N", alt_appointment_length_listbox, last_appointment_listbox) <= 0 THEN err_msg = err_msg & VbCr & "* The additional appointment block may not begin prior or equal to the first appointment block ending."
		
		'Converting the appointment times back from military time.
		IF DatePart("H", last_appointment_listbox) > 12 THEN 
			last_appointment_listbox = DateAdd("H", -12, last_appointment_listbox)
			alt_first_appointment_listbox = DateAdd("H", 0, alt_first_appointment_listbox)
		END IF
		IF DatePart("H", alt_first_appointment_listbox) > 12 THEN 
			alt_first_appointment_listbox = DateAdd("H", -12, alt_first_appointment_listbox)
			last_appointment_listbox = DateAdd("H", 0, last_appointment_listbox)
		END IF
	END IF
	IF last_appointment_listbox = "Select one..." THEN err_msg = err_msg & VbCr & "* You must choose a final appointment time."
	IF alt_first_appointment_listbox <> "Select one..." and alt_last_appointment_listbox = "Select one..." THEN err_msg = err_msg & VbCr & "* You have selected an initial appointment time for the additional appointment block, you must select a final appointment time."
	IF alt_last_appointment_listbox <> "Select one..." and alt_first_appointment_listbox = "Select one.." THEN err_msg = err_msg & VbCr & "* You have selected a final appointment time for the additional appointment block, you must select an initial appointment time."
	IF appointment_length_listbox = "Select one..." THEN err_msg = err_msg & VbCr & "* You must select an appointment length."
	IF alt_first_appointment_listbox <> "Select one..." and alt_appointment_length_listbox = "Select one..." THEN err_msg = err_msg & VbCr & "* Please choose an appointment length for the additional appointment block."
	IF err_msg <> "" THEN msgbox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

IF appointments_per_time_slot = "" THEN appointments_per_time_slot = 1
IF alt_appointments_per_time_slot = "" THEN alt_appointments_per_time_slot = 1

'Navigating to MAXIS
CALL check_for_MAXIS(false)
back_to_SELF
current_month = DatePart("M", date)
IF len(current_month) = 1 THEN current_month = "0" & current_month
current_year = DatePart("YYYY", date)
current_year = right(current_year, 2)

'Determining the month that the script will access REPT/REVS.
revs_month = DateAdd("M", 2, date)
IF developer_mode = True THEN revs_month = DateAdd("M", -1, revs_month)
revs_year = DatePart("YYYY", revs_month)
revs_year = right(revs_year, 2)
revs_month = DatePart("M", revs_month)
IF len(revs_month) = 1 THEN revs_month = "0" & revs_month

'writing current month
EMWriteScreen current_month, 20, 43
EMWriteScreen current_year, 20, 46
transmit

'navigating to REVS and entering REVS Month and year
CALL navigate_to_MAXIS_screen("REPT", "REVS")
EMWriteScreen revs_month, 20, 55
EMWriteScreen revs_year, 20, 58
transmit

'Checking to see if the worker running the script is the the worker selected, if not it will enter the selected worker's number
EMReadScreen current_worker, 7, 21, 6
IF UCASE(current_worker) <> UCASE(worker_number) THEN
	EMWriteScreen UCASE(worker_number), 21, 6
	transmit
END IF

'Grabbing case numbers from REVS for requested worker
Excel_row = 2
DO
	MAXIS_row = 7
	DO
		EMReadScreen case_number, 8, MAXIS_row, 6
		EMReadScreen SNAP_status, 1, MAXIS_row, 45
		EMReadScreen cash_status, 1, MAXIS_row, 34
		
		
		IF case_number = "        " then exit do     'navigates though until it runs out of case numbers to read
		
		'For some goofy reason the dash key shows up instead of the space key. No clue why. This will turn them into null variables.
		If cash_status = "-" then cash_status = ""
		If SNAP_status = "-" then SNAP_status = ""
		If HC_status = "-" then HC_status = ""
		
				'Using if...thens to decide if a case should be added (status isn't blank and respective box is checked)
		If trim(SNAP_status) = "N" or trim(SNAP_status) = "I" or trim(SNAP_status) = "U" then add_case_info_to_Excel = True
		If trim(cash_status) = "N" or trim(cash_status) = "I" or trim(cash_status) = "U" then add_case_info_to_Excel = True 
		'Adding the case to Excel
		If add_case_info_to_Excel = True then 
			ObjExcel.Cells(excel_row, 1).Value = case_number
			excel_row = excel_row + 1
		End if
		MAXIS_row = MAXIS_row + 1
		add_case_info_to_Excel = ""	'Blanking out variable
		case_number = ""			'Blanking out variable
	Loop until MAXIS_row = 19
	PF8
	EMReadScreen last_page_check, 21, 24, 2	'checking to see if we're at the end
Loop until last_page_check = "THIS IS THE LAST PAGE"

'Now the script will go through STAT/REVW for each case and check that the case is at CSR or ER and remove the cases that are at CSR from the list.
excel_row = 2
DO
	case_number = objExcel.cells(excel_row, 1).Value
	CALL navigate_to_MAXIS_screen("STAT", "REVW")
	'Checking for PRIV cases.
	EMReadScreen priv_check, 6, 24, 14 'If it can't get into the case needs to skip
	IF priv_check = "PRIVIL" THEN 'Delete priv cases from excel sheet, save to a list for later
		priv_case_list = priv_case_list & "|" & case_number
		SET objRange = objExcel.Cells(excel_row, 1).EntireRow			
		objRange.Delete
		excel_row = excel_row - 1
		msgbox priv_case_list
	ELSE
		EMwritescreen "x", 5, 58
		Transmit
		DO											'looping to check if the SNAP REVW popup is on the screen
			EMReadScreen SNAP_popup_check, 7, 5, 43
		LOOP until SNAP_popup_check = "Reports"
		
		'The script will now read the CSR MO/YR and the Recert MO/YR
		EMReadScreen CSR_mo, 2, 9, 26
		EMReadScreen CSR_yr, 2, 9, 32
		EMReadScreen recert_mo, 2, 9, 64
		EMReadScreen recert_yr, 2, 9, 70
		
		'It then compares what it read to the previously established current month plus 2 and determines if it is a recert or not. If it is a recert we need an interview
		IF CSR_mo = left(cm_plus_2, 2) and CSR_yr = right(cm_plus_2, 2) THEN RECERT_STATUS = "NO"
		IF recert_mo = left(cm_plus_2, 2) and recert_yr = right(cm_plus_2, 2) THEN RECERT_STATUS = "YES"
	
		'If it's not a recert, delete it from the excel list and move on with our lives
		IF RECERT_STATUS = "NO" THEN		
			Call navigate_to_MAXIS_screen("STAT", "PROG")
			EMReadScreen MFIP_prog_check, 2, 6, 67		'checking for an active MFIP case
			EMReadScreen MFIP_status_check, 4, 6, 74
			If MFIP_prog_check <> "MF" AND MFIP_status_check <> "ACTV" THEN 	'if MFIP is active, then case will not be deleted.
				SET objRange = objExcel.Cells(excel_row, 1).EntireRow			
				objRange.Delete				'all other cases that are not due for a recert will be deleted
				excel_row = excel_row - 1
			END If
		END IF
	END IF
	excel_row = excel_row + 1
LOOP UNTIL objExcel.Cells(excel_row, 1).Value = ""	'looping until the list of cases to check for recert is complete


'Now the script needs to go back to the start of the Excel file and start assigning appointments.
'FOR EACH day that is not checked, start assigning appointments according to DatePart("N", appointment) because DatePart"N" is minutes. Once datepart("N") = last_appointment_time THEN the script needs to jump to the next day.

'Going back to the top of the Excel to insert the appointment date and time in the list, yo
appointment_length_listbox = left(appointment_length_listbox, 2)	'Hacking the "mins" off the end of the appointment_length_listbox variable
alt_appointment_length_listbox = left(alt_appointment_length_listbox, 2)
excel_row = 2	'Declaring variable prior to the next for...next loop

FOR i = 1 to num_of_days
	IF month_array(i, 0) = 1 THEN		'These are the dates that the user has determined the agency/unit/worker
		appointment_time = appt_month & "/" & i & "/" & appt_year & " " & first_appointment_listbox		'putting together the date and time values.
		DO
			appointment_time = DateAdd("N", 0, appointment_time)	'Putting the date in a MM/DD/YYYY HH:MM format. It just looks nicer.
			appointment_time_for_viewing = appointment_time			'creating a new variable to handle the display of time to get it out of military time.
			IF DatePart("H", appointment_time_for_viewing) >= 13 THEN appointment_time_for_viewing = DateAdd("H", -12, appointment_time_for_viewing)
			FOR j = 1 TO appointments_per_time_slot					'Having the script create appointments_per_time_slot for each day and time.
				objExcel.Cells(excel_row, 2).Value = appointment_time_for_viewing
				excel_row = excel_row + 1
				IF objExcel.Cells(excel_row, 1).Value = "" THEN EXIT FOR
			NEXT
			IF objExcel.Cells(excel_row, 1).Value = "" THEN EXIT DO
			
			'This is where the script adds minutes for the next appointment.
			appointment_time = DateAdd("N", appointment_length_listbox, appointment_time)
			appointment_time = DateAdd("N", 0, appointment_time) 'Putting the date in a MM/DD/YYYY HH:MM format. Otherwise, the format is M/D/YYYY. It just looks nicer.
			
			'The variables "last_appointment_listbox_for_comparison" and "appointment_time_for_comparison" are used for the DO-LOOP. Because the script
			'handles time in military time, but clients do not, we need a way of handling the display of the date/time and the comparison of appointment times
			'against last appointment time variable.
			IF DatePart("H", last_appointment_listbox) < 7 THEN 
				last_appointment_listbox_for_comparison = DateAdd("H", 12, last_appointment_listbox)
			ELSE
				last_appointment_listbox_for_comparison = last_appointment_listbox
			END IF
			
			IF DatePart("H", appointment_time) < 7 THEN 
				appointment_time_for_comparison = DateAdd("H", 12, appointment_time)
			ELSE
				appointment_time_for_comparison = appointment_time
			END IF
		LOOP UNTIL (DatePart("H", appointment_time_for_comparison) > DatePart("H", last_appointment_listbox_for_comparison)) OR ((DatePart("H", appointment_time_for_comparison) >= DatePart("H", last_appointment_listbox_for_comparison)) AND DatePart("N", appointment_time_for_comparison) > DatePart("N", last_appointment_listbox_for_comparison))

		IF objExcel.Cells(excel_row, 1).Value = "" THEN EXIT FOR	'If there's nothing in the row, it means it's over.
		
		IF alt_first_appointment_listbox <> "Select one..." THEN 	'Same as above but for the second block 'o time
			appointment_time = appt_month & "/" & i & "/" & appt_year & " " & alt_first_appointment_listbox
			DO
				appointment_time = DateAdd("N", 0, appointment_time)	'Putting the date in a MM/DD/YYYY HH:MM format. It just looks nicer.
				appointment_time_for_viewing = appointment_time			'creating a new variable to handle the display of time to get it out of military time.
				IF DatePart("H", appointment_time_for_viewing) >= 13 THEN appointment_time_for_viewing = DateAdd("H", -12, appointment_time_for_viewing)
				FOR k = 1 TO alt_appointments_per_time_slot					'Having the script create appointments_per_time_slot for each day and time.
					objExcel.Cells(excel_row, 2).Value = appointment_time_for_viewing
					excel_row = excel_row + 1
					IF objExcel.Cells(excel_row, 1).Value = "" THEN EXIT FOR
				NEXT
				IF objExcel.Cells(excel_row, 1).Value = "" THEN EXIT DO
				
				'This is where the script adds minutes for the next appointment.
				appointment_time = DateAdd("N", alt_appointment_length_listbox, appointment_time)
				appointment_time = DateAdd("N", 0, appointment_time) 'Putting the date in a MM/DD/YYYY HH:MM format. Otherwise, the format is M/D/YYYY. It just looks nicer.
				
				'The variables "last_appointment_listbox_for_comparison" and "appointment_time_for_comparison" are used for the DO-LOOP. Because the script
				'handles time in military time, but clients do not, we need a way of handling the display of the date/time and the comparison of appointment times
				'against last appointment time variable.
				IF DatePart("H", alt_last_appointment_listbox) < 7 THEN 
					last_appointment_listbox_for_comparison = DateAdd("H", 12, alt_last_appointment_listbox)
				ELSE
					last_appointment_listbox_for_comparison = alt_last_appointment_listbox
				END IF
				
				IF DatePart("H", appointment_time) < 7 THEN 
					appointment_time_for_comparison = DateAdd("H", 12, appointment_time)
				ELSE
					appointment_time_for_comparison = appointment_time
				END IF
			LOOP UNTIL (DatePart("H", appointment_time_for_comparison) > DatePart("H", last_appointment_listbox_for_comparison)) OR ((DatePart("H", appointment_time_for_comparison) >= DatePart("H", last_appointment_listbox_for_comparison)) AND DatePart("N", appointment_time_for_comparison) > DatePart("N", last_appointment_listbox_for_comparison))		
		
		END IF	
		IF objExcel.Cells(excel_row, 1) = "" THEN EXIT FOR		'Because we're all out of cases at this point
	END IF
NEXT


If developer_mode = true Then
	excel_row = 2					'resetting excel row to start reading at the top 
	DO 								'looping until it meets a blank excel cell without a case number
		recert_status = ""			'resetting recert status for each run through the loop/case number
		forms_to_arep = ""
		forms_to_swkr = ""
		case_number = objExcel.cells(excel_row, 1).Value
		interview_time = objExcel.Cells(excel_row, 2).Value
		IF DatePart("H", interview_time) < 7 OR DatePart("H", interview_time) = 12 THEN    'converting from military time
			am_pm = "PM"
		ELSE	
			am_pm = "AM"
		END IF
		appt_minute_place_holder_because_reasons = DatePart("N", interview_time)
		IF appt_minute_place_holder_because_reasons = "0" THEN appt_minute_place_holder_because_reasons = "00"	'This is needed because DatePart("N", 10:00) = 0 and not 00. Times were being displayed 10:0
		interview_time = DatePart("M", interview_time) & "/" & DatePart("D", interview_time) & "/" & DatePart("YYYY", interview_time) & " " & DatePart("H", interview_time) & ":" & appt_minute_place_holder_because_reasons & " " & am_pm
		IF case_number = "" THEN EXIT DO      'exiting do if it finds a blank cell on the case number column
		
		back_to_self
		IF len(datepart("m", date)) = 1 THEN EMwritescreen "0" & datepart("m", date), 20, 43			'writing current month
		EMwritescreen right(datepart("YYYY", date), 2), 20, 46		'writing current year
		transmit
		
		If county_code <> "x127" then
			'Grabbing the phone number from ADDR
			CALL navigate_to_screen("STAT", "ADDR")
			EMReadScreen area_code, 3, 17, 45
			EMReadScreen remaining_digits, 9, 17, 50
			IF area_code = "___" THEN 'Reading phone 2 in case it is the only entered number
				EMReadScreen area_code, 3, 18, 45
				EMReadScreen remaining_digits, 9, 18, 50
			END IF
			IF area_code = "___" THEN 
				EMReadScreen area_code, 3, 19, 45 ' reading phone 3 
				EMReadScreen remaining_digits, 9, 19, 50
			END IF
			phone_number = area_code & remaining_digits
			phone_number = replace(phone_number, "_", " ") 'replaces _ to blank space so it can work with if statements looking for no phone numbers which looks for 12 spaces
		End If
		
		back_to_self
		CALL navigate_to_screen("SPEC", "MEMO")
		'Checking for AREP if found sending memo to them as well 
		'AREP/ALTP DISPLAY NOT CURRENTLY WORKING IN DEV MODE, THIS IS SOMETHING THAT SHOULD BE ADDED

		Memo_to_display = "The State DHS sent you a packet of paperwork. This is renewal paperwork for your SNAP case Your SNAP case is set to close on " &  last_day_of_recert & ". Please sign, date and return your renewal paperwork by " & left(cm_plus_1, 2) & "/08/" & right(cm_plus_1, 2) & vbNewLine &_
			"You must also do an interview for your SNAP case to continue. Your phone interview is scheduled for " & interview_time & "." & vbNewLine
		IF county_code = "x127" then    'allows for county 27 to have clients call them.
			Memo_to_display = Memo_to_display & "The phone number for you to call is " & contact_phone_number & "."
		ELSE	
			IF phone_number <> "            " THEN
				Memo_to_display = Memo_to_display & "The phone number we have for you is " & phone_number & ". This is the number we will call." & vbNewLine 
			ELSE
				Memo_to_display = Memo_to_display & "We currently do not have a phone number on file for you." & vbNewLine &_
					"Please call us at " & contact_phone_number & " to update your phone number, or if you would prefer an in-person interview." & vbNewLine
			END IF
		END IF
		
		Memo_to_display = Memo_to_display & "We must have your renewal paperwork to do your interview. Please send proofs with your renewal paperwork." & vbNewline &_
							" * Examples of income proofs: paystubs, income reports, business ledgers, income tax forms, etc." & vbNewLine &_
							" * Examples of housing cost proofs(if changed): rent/house payment receipts, mortgage, lease, etc." & vbNewline & vbNewLine &_
							" * Examples of medical cost proofs(if changed): prescriptions and medical bills, etc." & vbNewLine & vbNewLine &_
							"Please call us at ###-###-#### if you need to:" & vbNewLine &_ 
							" * Reschedule your appointment." & vbNewLine &_
							" * Report a new phone number, or other changes" & vbNewLine & vbNewLine &_
							" * Request an in-person interview."
		MsgBox Memo_to_display
		
		Case_note_to_display = "Case Note: " & "***SNAP Recertification Interview Scheduled***" & vbNewLine
		Case_note_to_display = Case_note_to_display & "* A phone interview has been scheduled for " & interview_time & "." & vbNewLine
		IF phone_number = "            " THEN 
				Case_note_to_display = Case_note_to_display & "No phone number in MAXIS as of " & date & "." & vbNewLine
			ELSE
				Case_note_to_display = Case_note_to_display & "* Client phone: " & phone_number & vbNewLine
		END IF
		If forms_to_arep = "Y" then Case_note_to_display = Case_note_to_display & "* Copy of notice sent to AREP." & vbNewLine
		If forms_to_swkr = "Y" then Case_note_to_display = Case_note_to_display & "* Copy of notice sent to Social Worker." & vbNewLine
		Case_note_to_display = Case_note_to_display & "---" & vbNewLine & worker_signature
		
		msgbox Case_note_to_display
		
		tikl_date = DatePart("M", interview_time) & "/" & DatePart("D", interview_time) & "/" & DatePart("YYYY", interview_time)
		
		MsgBox "Dail: ~*~*~CLIENT HAD RECERT INTERVIEW APPT AT " & interview_time & ". IF MISSED SEND NOMI." & vbNewLine &_
				"tikl date: " & tikl_date
		
		excel_row = excel_row + 1
			
	LOOP until objExcel.cells(excel_row, 1).Value = ""
	
		
	'Formatting the columns to autofit after they are all finished being created. 
	objExcel.Columns(1).autofit()
	objExcel.Columns(2).autofit()
	objExcel.Columns(3).autofit()
	objExcel.Columns(4).autofit()
	
Else    'if worker is actually running the script it will do this
	excel_row = 2					'resetting excel row to start reading at the top 
	DO 								'looping until it meets a blank excel cell without a case number
		recert_status = ""			'resetting recert status for each run through the loop/case number
		forms_to_arep = ""
		forms_to_swkr = ""
		case_number = objExcel.cells(excel_row, 1).Value
		interview_time = objExcel.Cells(excel_row, 2).Value
		IF DatePart("H", interview_time) < 7 OR DatePart("H", interview_time) = 12 THEN    'converting from military time
			am_pm = "PM"
		ELSE	
			am_pm = "AM"
		END IF
		appt_minute_place_holder_because_reasons = DatePart("N", interview_time)
		IF appt_minute_place_holder_because_reasons = 0 THEN appt_minute_place_holder_because_reasons = "00"	'This is needed because DatePart("N", 10:00) = 0 and not 00. Times were being displayed 10:0
		interview_time = DatePart("M", interview_time) & "/" & DatePart("D", interview_time) & "/" & DatePart("YYYY", interview_time) & " " & DatePart("H", interview_time) & ":" & appt_minute_place_holder_because_reasons & " " & am_pm
		IF case_number = "" THEN EXIT DO      'exiting do if it finds a blank cell on the case number column
		
		back_to_self
		IF len(datepart("m", date)) = 1 THEN EMwritescreen "0" & datepart("m", date), 20, 43			'writing current month
		EMwritescreen right(datepart("YYYY", date), 2), 20, 46		'writing current year
		transmit
		
		If county_code <> "x127" then   'county 27 has clients call in so no need to navigate to stat to grab phone number
			'Grabbing the phone number from ADDR
			CALL navigate_to_screen("STAT", "ADDR")
			EMReadScreen area_code, 3, 17, 45
			EMReadScreen remaining_digits, 9, 17, 50
			IF area_code = "___" THEN 'Reading phone 2 in case it is the only entered number
				EMReadScreen area_code, 3, 18, 45
				EMReadScreen remaining_digits, 9, 18, 50
			END IF
			IF area_code = "___" THEN 
				EMReadScreen area_code, 3, 19, 45 ' reading phone 3 
				EMReadScreen remaining_digits, 9, 19, 50
			END IF
			phone_number = area_code & remaining_digits
			phone_number = replace(phone_number, "_", " ")  'replaces _ to blank space so it can work with if statements looking for no phone numbers
		end If
		
		back_to_self
		CALL navigate_to_screen("SPEC", "MEMO")
		PF5
		EMReadScreen memo_display_check, 12, 2, 33
		If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
		'Checking for AREP if found sending memo to them as well
		row = 4
		col = 1
		EMSearch "ALTREP", row, col
		IF row > 4 THEN
			arep_row = row
			CALL navigate_to_screen("STAT", "AREP")
			EMReadscreen forms_to_arep, 1, 10, 45
			call navigate_to_screen("SPEC", "MEMO")
			PF5
		END IF
		'Checking for SWKR if found sending MEMO to them as well
		row = 4
		col = 1
		EMSearch "SOCWKR", row, col
		IF row > 4 THEN
			swkr_row = row
			call navigate_to_screen("STAT", "SWKR")
			EMReadscreen forms_to_swkr, 1, 15, 63
			call navigate_to_screen("SPEC", "MEMO")
			PF5
		END IF
		EMWriteScreen "x", 5, 10
		IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 10
		IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 10
		transmit
		'Writing the appointment and letter into a memo
		EMSendKey("************************************************************")
		CALL write_variable_in_SPEC_MEMO("The State DHS sent you a packet of paperwork. This is renewal paperwork for your SNAP case Your SNAP case is set to close on " &  last_day_of_recert & ". Please sign, date and return your renewal paperwork by " & left(cm_plus_1, 2) & "/08/" & right(cm_plus_1, 2) & ".")
		CALL write_variable_in_SPEC_MEMO("")
		CALL write_variable_in_SPEC_MEMO("You must also do an interview for your SNAP case to continue.")
		CALL write_variable_in_SPEC_MEMO("")
		IF county_code = "x127" THEN    'allows for county 27 to have clients call them.
			CALL write_variable_in_SPEC_MEMO("Your phone interview is scheduled for " & interview_time & ". The phone number for you to call is " & contact_phone_number & ".")
			ELSE
			IF phone_number <> "            " THEN
				CALL write_variable_in_SPEC_MEMO("Your phone interview is scheduled for " & interview_time & ". The phone number we have for you is " & phone_number & ". This is the number we will call.")
			else
				CALL write_variable_in_SPEC_MEMO("Your phone interview is scheduled for " & interview_time & ". We currently do not have a phone number on file for you.")
				CALL write_variable_in_SPEC_MEMO("Please call us at " & contact_phone_number & " to update your phone number, or if you would prefer an in-person interview.")
			end if
		End If
		CALL write_variable_in_SPEC_MEMO("")
		CALL write_variable_in_SPEC_MEMO("We must have your renewal paperwork to do your interview.")
		CALL write_variable_in_SPEC_MEMO("")
		CALL write_variable_in_SPEC_MEMO("Please send proofs with your renewal paperwork.")
		CALL write_variable_in_SPEC_MEMO(" * Examples of income proofs: paystubs, income reports, business ledgers, income tax forms, etc.")
		CALL write_variable_in_SPEC_MEMO("")
		CALL write_variable_in_SPEC_MEMO(" * Examples of housing cost proofs(if changed): rent/house payment receipt, mortgage, lease, etc.")
		CALL write_variable_in_SPEC_MEMO("")
		CALL write_variable_in_SPEC_MEMO(" * Examples of medical cost proofs(if changed): Prescription and medical bills, etc.")
		CALL write_variable_in_SPEC_MEMO("")
		CALL write_variable_in_SPEC_MEMO("Please call us at " & contact_phone_number & " if you need to:")
		CALL write_variable_in_SPEC_MEMO(" * Reschedule your appointment.")
		CALL write_variable_in_SPEC_MEMO(" * Report a new phone number, or other changes.")
		CALL write_variable_in_SPEC_MEMO(" * Request an in-person interview.")
		PF4
		back_to_self
		
		'case noting appointment time and date
		CALL navigate_to_screen("CASE", "NOTE")
		PF9
		
		EMSendKey "***SNAP Recertification Interview Scheduled***"
		CALL write_variable_in_case_note("* A phone interview has been scheduled for " & interview_time & ".")
		IF phone_number = "            " THEN 
				CALL write_variable_in_case_note("No phone number in MAXIS as of " & date & ".")
			ELSE
				CALL write_variable_in_case_note("* Client phone: " & phone_number)
		END IF
		If forms_to_arep = "Y" then call write_variable_in_case_note("* Copy of notice sent to AREP.")
		If forms_to_swkr = "Y" then call write_variable_in_case_note("* Copy of notice sent to Social Worker.")
		call write_variable_in_case_note("---")
		call write_variable_in_case_note(worker_signature)
		
		'adding appointment time and date to outlook calendar if requested by worker
		IF outlook_calendar_check = 1 THEN 
			appt_date_for_outlook = DatePart("M", interview_time) & "/" & DatePart("D", interview_time) & "/" & DatePart("YYYY", interview_time)
			appt_time_for_outlook = DatePart("H", interview_time) & ":" & DatePart("N", interview_time)
			IF DatePart("N", interview_time) = 0 THEN appt_time_for_outlook = DatePart("H", interview_time) & ":00"
			appt_end_time_for_outlook = DateAdd("N", appointment_length_listbox, interview_time)
			appt_end_time_for_outlook = DatePart("H", appt_end_time_for_outlook) & ":" & DatePart("N", appt_end_time_for_outlook)
			IF DatePart("N", interview_time) = 0 THEN appt_end_time_for_outlook = DatePart("H", appt_end_time_for_outlook) & ":00"
			appointment_subject = "SNAP RECERT"
			appointment_body = "Case Number: " & case_number
			IF phone_number = "            " THEN 
				appointment_location = "No phone number in MAXIS as of " & date & "."
			ELSE
				appointment_location = "Phone: " & phone_number
			END IF
			appointment_reminder = True
			appointment_category = "Recertification Interview"
			'using the variables created above to generate the Outlook Appointment from the custom function.
			CALL create_outlook_appointment(appt_date_for_outlook, appt_time_for_outlook, appt_end_time_for_outlook, appointment_subject, appointment_body, appointment_location, appointment_reminder, appointment_category)
		END IF
		
		'TIKLing to remind the worker to send NOMI if appointment is missed.
		CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
		tikl_date = DatePart("M", interview_time) & "/" & DatePart("D", interview_time) & "/" & DatePart("YYYY", interview_time)
		CALL create_MAXIS_friendly_date(tikl_date, 0, 5, 18)
		EMWriteScreen "~*~*~CLIENT HAD RECERT INTERVIEW APPT AT " & interview_time, 9, 3
		EMWriteScreen "IF MISSED SEND NOMI", 10, 3
		transmit
		PF3
		'END IF
		
		excel_row = excel_row + 1
			
	LOOP until objExcel.cells(excel_row, 1).Value = ""
	
	'Formatting the columns to autofit after they are all finished being created. 
	objExcel.Columns(1).autofit()
	objExcel.Columns(2).autofit()
	objExcel.Columns(3).autofit()
	objExcel.Columns(4).autofit()
End IF
'Creating the list of privileged cases and adding to the spreadsheet
	prived_case_array = split(priv_case_list, "|")
	excel_row = 2
FOR EACH case_number in prived_case_array
	objExcel.cells(excel_row, 5).value = case_number
	excel_row = excel_row + 1
NEXT
		
script_end_procedure("Success, the excel file now has all of the cases that have had interviews scheduled.  Please manually review the list of priveleged cases.")
