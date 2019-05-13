'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - REVS SCRUBBER.vbs"
start_time = Timer
STATS_counter = 1			 'sets the stats counter at one
STATS_manualtime = 304			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
'END OF stats block==============================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
If IsEmpty(FuncLib_URL) = True Then	'Shouldn't load FuncLib if it already loaded once
	If run_locally = False Or run_locally = "" Then	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		If use_master_branch = True Then			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End If
		Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, False							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		If req.Status = 200 Then									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		Else														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		End If
	Else
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	End If
End If
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = Array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
Call changelog_update("1/2/2018", "Fixing bug that prevented the script from writing SPEC/MEMO due to MAXIS updates.", "Casey Love, Ramsey County")
Call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Declaring variables----------------------------------------------------------------------------------------------------
'DIM month, full_date_to_display, num_of_days, next_month, available_dates_array, month_to_use
'DIM first_appointment_listbox, olAppointmentItem, olRecursDaily
'DIM appointment_length_listbox, last_appointment_listbox, time_array_30_min
'DIM alt_appointment_length_listbox, alt_first_appointment_listbox
'DIM last_day_of_recert, cm_plus_1, cm_plus_2, alt_duplicate_appt_times
'DIM worker_numberm, calendar_dlgm, ButtonPressed, duplicate_appt_times
'DIM alt_appointments_per_time_slot, contact_phone_number, worker_signature
'DIM err_msg, day_of_month, time_array, worker_number
'DIM REVS_scrubber_time_dialog, REVS_scrubber_initial_dialog
'DIM objWorkbook, objRange, appointments_per_time_slot
'DIM calendar_month, worker_county_code, objAppointment, objExcel
'DIM appt_date, appt_category, appt_body, appt_month, developer_mode
'DIM appt_minute_place_holder_because_reasons, appt_location, appt_end_time
'DIM appt_time_list, appt_subject, appt_start_time, appt_reminder
'DIM current_year, current_worker, current_month, y, i, x, appt_year
'DIM MAXIS_row, MAXIS_case_number, excel_row, excel, REPT_year, REPT_month
'DIM alt_last_appointment_listbox, calendar_dialog, SNAP_popup_check, SNAP_status
'DIM cash_status, last_page_check, HC_status, add_case_info_to_Excel, CSR_mo
'DIM CSR_yr, recert_mo, recert_status, recert_yr, appointment_time
'DIM appointment_time_for_comparison, appointment_time_for_viewing, appointments, j
'DIM last_appointment_listbox_for_comparison, last_appointment_time, interview_time
'DIM am_pm, recert_status, forms_to_swkr

'Variables needed for the script-----------------------------------------------------------------------------------------------------
appt_time_list = "15 mins"+Chr(9)+"30 mins"+Chr(9)+"45 mins"+Chr(9)+"60 mins"		'Used by the dialog to display possible times/lengths of appointments
Call convert_array_to_droplist_items(time_array_30_min, time_array)					'Schedules time blocks in 30 minute increments

'CONSTANTS WE LIKE--------------------------------------------------------------
Const worker_number_col 	= 1
Const case_number_col 		= 2
Const interview_time_col 	= 3
Const privileged_case_col	= 5

'Custom functions (should merge with FuncLib when tested/confirmed to work)---------------------------------------------------
'Function to create dynamic calendar out of checkboxes
Function create_dynamic_calendar_dialog(month_to_use, available_dates_array)
	'Generating a calendar
	'Determining the number of days in the calendar month.
	next_month = DateAdd("M", 1, month_to_use)												'This is the next calendar month (1 month ahead of the date in month_to_use)
	next_month = DatePart("M", next_month) & "/01/" & DatePart("YYYY", next_month)			'Converts whatever the next_month variable is to a MM/01/YYYY format
	num_of_days = DatePart("D", (DateAdd("D", -1, next_month)))								'Determines the number of days in a month by using DatePart to get the day of the last day of the month, and just using the day variable gives us a total

	'Redeclares the available dates array to be sized appropriately (with the right amount of dates) and another dimension for whether-or-not it was selected
	ReDim available_dates_array(num_of_days, 0)

	'Actually displays the dialog
	BeginDialog calendar_dialog, 0, 0, 280, 190, month_to_use
		Text 5, 10, 270, 25, "Please check the days to schedule appointments. You cannot schedule appointments prior to the 8th."
		Text 5, 35, 270, 25, "Please note that Auto-Close Notices are sent on the 16th. To reduce confusion, you may want to schedule before the 16th."
		'This next part`creates a line similar to "Check appointment dates in February 2016 for March 2016 recertifications."
		'														The month name of month_to_use					The year of month_to_use					Next month name									Next month year
		Text 5, 65, 265, 10, ("Check appointment dates in " & MonthName(DatePart("M", month_to_use)) & " " & DatePart("YYYY", month_to_use) & " for " & MonthName(DatePart("M", next_month)) & " " & DatePart("YYYY", next_month) & " recertifications.")

		'Defining the vertical position starting point for the for...next which displays dates in the dialog
		vertical_position = 85

		'This for...next displays dates in the dialog, and has checkboxes for available dates (defined in-code as dates before the 8th)
		For day_to_display = 1 To num_of_days																						'From first day of month to last day of month...

			full_date_to_display = (DatePart("M", month_to_use) & "/" & day_to_display & "/" & DatePart("YYYY", month_to_use))		'Determines the full date to display in the dialog. It needs the full date to determine the day-of-week (we obviously don't want weekends)
			horizontal_position = 15 + (40 * (WeekDay(full_date_to_display) - 1))													'horizontal position of this is the weekday numeric value (1-7) * 40, minus 1, and plus 15 pixels
			If WeekDay(full_date_to_display) = vbSunday And day_to_display <> 1 Then vertical_position = vertical_position + 15		'If the day of the week isn't Sunday and the day isn't first of the month, kick the vertical position up another 15 pixels

			'If the weekday is a Sunday or Saturday, then we'll uncheck it. Otherwise it's checked.
			If WeekDay(full_date_to_display) = vbSunday Or WeekDay(full_date_to_display) = vbSaturday Then
				available_dates_array(day_to_display, 0) = unchecked
			Else
				available_dates_array(day_to_display, 0) = checked
			End If

			'This blocks out anything that's an unavailable date, currently defined as any date before the 8th. Other dates display as a checkbox.
			If day_to_display < 8 Then
				Text horizontal_position, vertical_position, 30, 10, " x " & day_to_display
				available_dates_array(day_to_display, 0) = o 'unchecking so selections cannot be made for DD 01-07
			Else
				CheckBox horizontal_position, vertical_position, 35, 10, day_to_display, available_dates_array(day_to_display, 0)
			End If
		Next
		ButtonGroup ButtonPressed
		OkButton 175, 170, 50, 15
		CancelButton 225, 170, 50, 15
	EndDialog

	Dialog calendar_dialog
	If ButtonPressed = cancel Then stopscript
End Function


'DIALOGS -----------------------------------------------------------------------------------------------
BeginDialog REVS_scrubber_initial_dialog, 0, 0, 501, 130, "REVS scrubber initial dialog"
  EditBox 165, 5, 195, 15, worker_number_editbox
  EditBox 165, 25, 55, 15, worker_signature
  EditBox 290, 45, 70, 15, contact_phone_number
  ButtonGroup ButtonPressed
    OkButton 380, 110, 50, 15
    CancelButton 435, 110, 50, 15
    PushButton 250, 115, 105, 10, "SIR instructions for this script", SIR_instructions_button
  Text 5, 10, 155, 10, "Worker numbers to run, separated by commas:"
  Text 5, 30, 155, 10, "Worker signature (for appointment case notes):"
  Text 5, 50, 280, 10, "Contact phone number with area code (so client can report changes in phone number):"
  GroupBox 365, 5, 130, 55, "Description"
  Text 375, 15, 115, 40, "This script will schedule appointments (in advance) for cases that require an interview for recertification, and will do so for an entire caseload."
  GroupBox 365, 65, 130, 40, "What you need before you start"
  Text 375, 75, 115, 25, "Individual or case-banked caseloads which have cases that require an interview. "
  GroupBox 5, 65, 350, 40, "PLEASE NOTE"
  Text 10, 75, 340, 25, "The script will not be available for use until the 16th of each month, as the script goes into current month plus two to schedule the appointments with proper advance notice (example: REVS scrubber will be available 02/16/YYYY to schedule recertification interviews for 04/YYYY reviews)."
EndDialog

BeginDialog REVS_scrubber_time_dialog, 0, 0, 291, 270, "REVS Scrubber Time Dialog"
  DropListBox 75, 25, 60, 15, "Select one..."+Chr(9)+ time_array, first_appointment_listbox
  DropListBox 215, 25, 60, 15, "Select one..."+Chr(9)+ time_array, last_appointment_listbox
  DropListBox 110, 50, 65, 15, "Select one..."+Chr(9)+ appt_time_list, appointment_length_listbox
  CheckBox 10, 75, 140, 10, "Duplicate appointments per time slot?", duplicate_appt_times
  EditBox 240, 70, 35, 15, appointments_per_time_slot
  DropListBox 75, 135, 60, 15, "Select one..."+Chr(9)+ time_array, alt_first_appointment_listbox
  DropListBox 215, 135, 60, 15, "Select one..."+Chr(9)+ time_array, alt_last_appointment_listbox
  DropListBox 115, 160, 65, 15, "Select one..."+Chr(9)+ appt_time_list, alt_appointment_length_listbox
  CheckBox 10, 185, 135, 10, "Duplicate appointments per time slot?", alt_duplicate_appt_times
  EditBox 240, 180, 35, 15, alt_appointments_per_time_slot
  	'County Add in below
    CheckBox 10, 215, 200, 10, "Check here to add appointments to your Outlook calendar.", outlook_calendar_check

  EditBox 230, 210, 45, 15, max_reviews_per_worker
  ButtonGroup ButtonPressed
    OkButton 170, 230, 50, 15
    CancelButton 225, 230, 50, 15
    PushButton 10, 235, 105, 10, "SIR instructions for this script", SIR_instructions_button
  Text 10, 30, 60, 10, "First appointment:"
  Text 150, 30, 60, 10, "Last appointment:"
  Text 10, 55, 95, 10, "Time between Appointments:"
  Text 160, 75, 80, 10, "How many per time slot:"
  Text 20, 115, 260, 10, "*NOTE: Use this block for scheduling appointments around your lunch break."
  Text 10, 140, 60, 10, "First appointment:"
  Text 150, 140, 60, 10, "Last appointment:"
  Text 10, 165, 95, 10, "Time between Appointments:"
  Text 160, 185, 80, 10, "How many per time slot:"
  Text 10, 215, 215, 10, "Maximum reviews to schedule per worker (leave blank for no cap):"
  GroupBox 5, 10, 275, 85, "Main Appointment Block"
  GroupBox 5, 100, 275, 105, "Additional Appointment Block"
EndDialog

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""

'creating a last day of recert variable
last_day_of_recert = CM_plus_2_mo & "/01/" & CM_plus_2_yr
last_day_of_recert = DateAdd("D", -1, last_day_of_recert)

'Grabbing the worker's X number.
Call find_variable("User: ", worker_number, 7)
get_county_code

'if user is Hennepin Co. user, then contact phone number auto-fills with EZ info number
If worker_county_code = "x127" Then contact_phone_number = "612-596-1300"

'Display REVS scrubber initial dialog. If contact_phone_number is UUDDLRLRBA then it'll enable developer mode.
Do
	Do
		Do
			err_msg = ""
			Dialog REVS_scrubber_initial_dialog
			If ButtonPressed = 0 Then stopscript
			If ButtonPressed = SIR_instructions_button Then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/Script%20Instructions%20Wiki/REVS%20Scrubber.aspx")
		Loop until ButtonPressed = -1
		If worker_number_editbox = "" 		Then err_msg = err_msg & vbNewLine & "* You must enter at least one worker number in the worker number editbox."
		If worker_signature = "" 			Then err_msg = err_msg & vbNewLine & "* You must sign the case notes regarding the appointments."
		If contact_phone_number = ""		Then err_msg = err_msg & vbNewLine & "* You must provide a contact phone number in case the client needs to report a new phone number to you."
		'Display the error message
		If err_msg <> "" Then MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue."
	Loop UNTIL err_msg = ""
	Call check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = False					'loops until user passwords back in

'Entering developer mode if Konami code entered as contact_phone_number.
If contact_phone_number = "UUDDLRLRBA" Then
	developer_mode = True
	MsgBox "You have enabled Developer Mode." & vbCr & vbCr & "The script will not enter information into MAXIS, but it will navigate, showing you where the script would otherwise have been."
End If

'Determining the month that the script will access REPT/REVS which is CM + 2
REPT_month = CM_plus_2_mo
REPT_year = CM_plus_2_yr

'resetting variables for developer mode to allow testing in current month before the 16th.
day_of_month = DatePart("D", Date)
'Stopping the script if the user is running it before the 16th of the month.
If developer_mode <> True Then
	If day_of_month < 16 Then script_end_procedure("You cannot run this script before the 16th of the month.") 'to boot the user before the script tries to access a blank REPT/REVS.
ElseIf developer_mode = True Then
    If day_of_month < 16 Then
		MsgBox "Current month plus 2 is not available currently. Script will continue in current month plus 1."
        REPT_month = CM_plus_1_mo
        REPT_year = CM_plus_1_yr
	End If
End If

'Formatting the dates
calendar_month = DateAdd("M", 1, Date)
appt_month = DatePart("M", calendar_month)
appt_year = DatePart("YYYY", calendar_month)
next_month = DateAdd("M", 1, calendar_month)
next_month = DatePart("M", next_month) & "/01/" & DatePart("YYYY", next_month)
num_of_days = DatePart("D", (DateAdd("D", -1, next_month)))

'Generating the calendar
ReDim available_dates_array(num_of_days, 0)
Call create_dynamic_calendar_dialog(calendar_month, available_dates_array)

'Determining the appropriate times to set appointments.
Do
	Do
		err_msg = ""
		Do
			dialog REVS_scrubber_time_dialog
			cancel_confirmation
			If ButtonPressed = SIR_instructions_button Then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/Script%20Instructions%20Wiki/REVS%20Scrubber.aspx")
		Loop until ButtonPressed = -1

		If first_appointment_listbox = "Select one..." Then err_msg = err_msg & VbCr & "* You must choose an initial appointment time."
		If first_appointment_listbox <> "Select one..." And last_appointment_listbox <> "Select one..." Then
			'Converting the appointment times for comparison. VBS runs in military time.
			If DatePart("H", last_appointment_listbox) < 7 Then
				last_appointment_listbox = DateAdd("H", 12, last_appointment_listbox)
				first_appointment_listbox = DateAdd("H", 0, first_appointment_listbox)
			End If
			If DatePart("H", first_appointment_listbox) < 7 Then
				first_appointment_listbox = DateAdd("H", 12, first_appointment_listbox)
				last_appointment_listbox = DateAdd("H", 0, last_appointment_listbox)
			End If

			If DateDiff("N", first_appointment_listbox, last_appointment_listbox) < 0 Then err_msg = err_msg & VbCr & "* The last appointment may not be earlier than the first appointment."

			'Converting the appointment times back from military time.
			If DatePart("H", last_appointment_listbox) > 12 Then
				last_appointment_listbox = DateAdd("H", -12, last_appointment_listbox)
				first_appointment_listbox = DateAdd("H", 0, first_appointment_listbox)
			End If
			If DatePart("H", first_appointment_listbox) > 12 Then
				first_appointment_listbox = DateAdd("H", -12, first_appointment_listbox)
				last_appointment_listbox = DateAdd("H", 0, last_appointment_listbox)
			End If
		End If
		If alt_first_appointment_listbox <> "Select one..." And alt_last_appointment_listbox <> "Select one..." Then
			'Converting the appointment times for comparison. VBS runs in military time.
			If DatePart("H", alt_last_appointment_listbox) < 7 Then
				alt_last_appointment_listbox = DateAdd("H", 12, alt_last_appointment_listbox)
				alt_first_appointment_listbox = DateAdd("H", 0, alt_first_appointment_listbox)
			End If
			If DatePart("H", alt_first_appointment_listbox) < 7 Then
				alt_first_appointment_listbox = DateAdd("H", 12, alt_first_appointment_listbox)
				alt_last_appointment_listbox = DateAdd("H", 0, alt_last_appointment_listbox)
			End If

			If DateDiff("N", alt_first_appointment_listbox, alt_last_appointment_listbox) < 0 Then err_msg = err_msg & VbCr & "* The additional appointment block has an ending earlier than it begins."

			'Converting the appointment times back from military time.
			If DatePart("H", alt_last_appointment_listbox) > 12 Then
				alt_last_appointment_listbox = DateAdd("H", -12, alt_last_appointment_listbox)
				alt_first_appointment_listbox = DateAdd("H", 0, alt_first_appointment_listbox)
			End If
			If DatePart("H", alt_first_appointment_listbox) > 12 Then
				alt_first_appointment_listbox = DateAdd("H", -12, alt_first_appointment_listbox)
				alt_last_appointment_listbox = DateAdd("H", 0, alt_last_appointment_listbox)
			End If
		End If
		If last_appointment_listbox <> "Select one..." And alt_first_appointment_listbox <> "Select one..." Then
			'Converting the appointment times for comparison. VBS runs in military time.
			If DatePart("H", last_appointment_listbox) < 7 Then
				last_appointment_listbox = DateAdd("H", 12, last_appointment_listbox)
				alt_first_appointment_listbox = DateAdd("H", 0, alt_first_appointment_listbox)
			End If
			If DatePart("H", alt_first_appointment_listbox) < 7 Then
				alt_first_appointment_listbox = DateAdd("H", 12, alt_first_appointment_listbox)
				last_appointment_listbox = DateAdd("H", 0, last_appointment_listbox)
			End If

			'Converting the appointment times back from military time.
			If DatePart("H", last_appointment_listbox) > 12 Then
				last_appointment_listbox = DateAdd("H", -12, last_appointment_listbox)
				alt_first_appointment_listbox = DateAdd("H", 0, alt_first_appointment_listbox)
			End If
			If DatePart("H", alt_first_appointment_listbox) > 12 Then
				alt_first_appointment_listbox = DateAdd("H", -12, alt_first_appointment_listbox)
				last_appointment_listbox = DateAdd("H", 0, last_appointment_listbox)
			End If
		End If
		'Error message handling
		If last_appointment_listbox = "Select one..." 																Then err_msg = err_msg & VbCr & "* You must choose a final appointment time."
		If alt_first_appointment_listbox <> "Select one..." And alt_last_appointment_listbox = "Select one..." 		Then err_msg = err_msg & VbCr & "* You have selected an initial appointment time for the additional appointment block, you must select a final appointment time."
		If alt_last_appointment_listbox <> "Select one..." And alt_first_appointment_listbox = "Select one.." 		Then err_msg = err_msg & VbCr & "* You have selected a final appointment time for the additional appointment block, you must select an initial appointment time."
		If appointment_length_listbox = "Select one..." 															Then err_msg = err_msg & VbCr & "* You must select an appointment length."
		If alt_first_appointment_listbox <> "Select one..." And alt_appointment_length_listbox = "Select one..." 	Then err_msg = err_msg & VbCr & "* Please choose an appointment length for the additional appointment block."

		'Display the error message
		If err_msg <> "" Then MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	Loop UNTIL err_msg = ""
	Call check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = False					'loops until user passwords back in

'Converts to integer
If max_reviews_per_worker <> "" Then max_reviews_per_worker = Abs(max_reviews_per_worker)

'Opening the Excel file, (now that the dialog is done)
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'formatting excel file with columns for case number and interview date/time
objExcel.cells(1, worker_number_col).value 			= "x1 number"
objExcel.Cells(1, worker_number_col).Font.Bold 		= True
objExcel.cells(1, case_number_col).value 			= "CASE NUMBER"
objExcel.Cells(1, case_number_col).Font.Bold 		= True
objExcel.Cells(1, interview_time_col).value 		= "Interview Date & Time"
objExcel.cells(1, interview_time_col).Font.Bold 	= True
objExcel.cells(1, privileged_case_col).value 		= "Privileged Cases"
objExcel.cells(1, privileged_case_col).Font.Bold 	= True

'If the appointments_per_time_slot variable isn't declared, it defaults to 1
If appointments_per_time_slot = "" Then appointments_per_time_slot = 1
If alt_appointments_per_time_slot = "" Then alt_appointments_per_time_slot = 1

'We need to get back to SELF and manually update the footer month
back_to_SELF

'writing current month and transmitting
EMWriteScreen CM_mo, 20, 43
EMWriteScreen CM_yr, 20, 46
transmit

'navigating to REVS and entering REVS Month and year as determined above
Call navigate_to_MAXIS_screen("REPT", "REVS")
EMWriteScreen REPT_month, 20, 55
EMWriteScreen REPT_year, 20, 58
transmit

worker_number_array = Split(worker_number_editbox, ",")

excel_row = 2	'Declaring variable prior to do...loops

For Each worker_number in worker_number_array

	'Trims any spaces off the worker_number
	worker_number = Trim(worker_number)

	'Checking to see if the worker running the script is the the worker selected, if not it will enter the selected worker's number
	EMReadScreen current_worker, 7, 21, 6
	If UCase(current_worker) <> UCase(worker_number) Then
		EMWriteScreen UCase(worker_number), 21, 6
		transmit
	End If

	'Grabbing case numbers from REVS for requested worker
	reviews_total = 0	'Sets this to 0 for the following do...loop. It'll exit once it's hit the reviews cap

	Do	'All of this loops until last_page_check = "THIS IS THE LAST PAGE"
		MAXIS_row = 7	'Setting or resetting this to look at the top of the list
		Do		'All of this loops until MAXIS_row = 19
			'Reading case information (case number, SNAP status, and cash status)
			EMReadScreen MAXIS_case_number, 8, MAXIS_row, 6
			EMReadScreen SNAP_status, 1, MAXIS_row, 45
			EMReadScreen cash_status, 1, MAXIS_row, 39

			'Navigates though until it runs out of case numbers to read
			If MAXIS_case_number = "        " Then Exit Do

			'For some goofy reason the dash key shows up instead of the space key. No clue why. This will turn them into null variables.
			If cash_status = "-" 	Then cash_status = ""
			If SNAP_status = "-" 	Then SNAP_status = ""
			If HC_status = "-" 		Then HC_status = ""

			'Using if...thens to decide if a case should be added (status isn't blank and respective box is checked)
			If ( ( Trim(SNAP_status) = "N" Or Trim(SNAP_status) = "I" Or Trim(SNAP_status) = "U" ) Or ( Trim(cash_status) = "N" Or Trim(cash_status) = "I" Or Trim(cash_status) = "U" ) ) And reviews_total <= max_reviews_per_worker Then
				add_case_info_to_Excel = True
				reviews_total = reviews_total + 1
			End If

			'Adding the case to Excel
			If add_case_info_to_Excel = True Then
				ObjExcel.Cells(excel_row, worker_number_col).value = worker_number
				ObjExcel.Cells(excel_row, case_number_col).value = MAXIS_case_number
				excel_row = excel_row + 1
			End If

			'On the next loop it must look to the next row
			MAXIS_row = MAXIS_row + 1

			'Clearing variables before next loop
			add_case_info_to_Excel = ""
			MAXIS_case_number = ""

			If reviews_total >= max_reviews_per_worker Then Exit Do

		Loop until MAXIS_row = 19		'Last row in REPT/REVS

		'Because we were on the last row, or exited the do...loop because the case number is blank, it PF8s, then reads for the "THIS IS THE LAST PAGE" message (if found, it exits the larger loop)
		PF8
		EMReadScreen last_page_check, 21, 24, 2	'checking to see if we're at the end
        'if max reviews are reached, the goes to next worker is applicable
		If reviews_total >= max_reviews_per_worker Then Exit Do
	Loop until last_page_check = "THIS IS THE LAST PAGE"
Next

'Now the script will go through STAT/REVW for each case and check that the case is at CSR or ER and remove the cases that are at CSR from the list.
excel_row = 2		'Resets the variable to 2, as it needs to look through all of the cases on the Excel sheet!

Do 'Loops until there are no more cases in the Excel list
    recert_status = "NO"	'Defaulting this to no because if SNAP or MFIP are not active - no recert will be scheduled
	'Grabs the case number
	MAXIS_case_number = objExcel.cells(excel_row, case_number_col).value
	'Goes to STAT/PROG
	Call navigate_to_MAXIS_screen("STAT", "PROG")

	'Checking for PRIV cases.
	EMReadScreen priv_check, 6, 24, 14 'If it can't get into the case needs to skip
	If priv_check = "PRIVIL" Then 'Delete priv cases from excel sheet, save to a list for later
		priv_case_list = priv_case_list & "|" & MAXIS_case_number
		Set objRange = objExcel.Cells(excel_row, 1).EntireRow
		objRange.Delete
		excel_row = excel_row - 1
	Else		'For all of the cases that aren't privileged...
		MFIP_ACTIVE = False		'Setting some variables for the loop
		SNAP_ACTIVE = False

		SNAP_status_check = ""
		MFIP_prog_1_check = ""
		MFIP_status_1_check = ""
		MFIP_prog_2_check = ""
		MFIP_status_2_check = ""

		'Reading the status and program
		EMReadScreen SNAP_status_check, 4, 10, 74

		EMReadScreen MFIP_prog_1_check, 2, 6, 67		'checking for an active MFIP case
		EMReadScreen MFIP_status_1_check, 4, 6, 74
		EMReadScreen MFIP_prog_2_check, 2, 6, 67		'checking for an active MFIP case
		EMReadScreen MFIP_status_2_check, 4, 6, 74

		'Logic to determine if MFIP is active
		If MFIP_prog_1_check = "MF" Then
			If MFIP_status_1_check = "ACTV" Then MFIP_ACTIVE = True
		ElseIf MFIP_prog_2_check = "MF" Then
			If MFIP_status_2_check = "ACTV" Then MFIP_ACTIVE = True
		End If

		'Only looks for SNAP if MFIP is not active
		If MFIP_ACTIVE = False Then
			If SNAP_status_check = "ACTV" Then SNAP_ACTIVE = True
		End If

		'Going to STAT/REVW to to check for ER vs CSR for SNAP cases
		Call navigate_to_MAXIS_screen("STAT", "REVW")
		If MFIP_ACTIVE = True Then recert_status = "YES"	'MFIP will only have an ER - so if listed on REVS - will be an ER - don't need to check dates
		If SNAP_ACTIVE = True Then
			EMReadScreen SNAP_review_check, 8, 9, 57
			If SNAP_review_check = "__ 01 __" Then 		'If this is blank there are big issues
				recert_status = "NO"
			Else
				EMwritescreen "x", 5, 58		'Opening the SNAP popup
				Transmit
				Do
				    EMReadScreen SNAP_popup_check, 7, 5, 43
				Loop until SNAP_popup_check = "Reports"

				'The script will now read the CSR MO/YR and the Recert MO/YR
				EMReadScreen CSR_mo, 2, 9, 26
				EMReadScreen CSR_yr, 2, 9, 32
				EMReadScreen recert_mo, 2, 9, 64
				EMReadScreen recert_yr, 2, 9, 70

				'Comparing CSR and ER daates to the month of REVS review
				If CSR_mo = Left(REPT_month, 2) And CSR_yr = Right(REPT_year, 2) Then recert_status = "NO"
				If recert_mo = Left(REPT_month, 2) And recert_yr <> Right(REPT_year, 2) Then recert_status = "NO"
				If recert_mo = Left(REPT_month, 2) And recert_yr = Right(REPT_year, 2) Then recert_status = "YES"
			End If
		End If

		'Removing the case from the spreadsheet if not a recert
		If recert_status = "NO" Then
			Set objRange = objExcel.Cells(excel_row, 1).EntireRow
			objRange.Delete				'all other cases that are not due for a recert will be deleted
			excel_row = excel_row - 1
		End If
	End If
    STATS_counter = STATS_counter + 1						'adds one instance to the stats counter
    excel_row = excel_row + 1
Loop UNTIL objExcel.Cells(excel_row, 2).value = ""	'looping until the list of cases to check for recert is complete

'Now the script needs to go back to the start of the Excel file and start assigning appointments.
'FOR EACH day that is not checked, start assigning appointments according to DatePart("N", appointment) because DatePart"N" is minutes. Once datepart("N") = last_appointment_time THEN the script needs to jump to the next day.

'Going back to the top of the Excel to insert the appointment date and time in the list
appointment_length_listbox = Left(appointment_length_listbox, 2)	'Hacking the "mins" off the end of the appointment_length_listbox variable
alt_appointment_length_listbox = Left(alt_appointment_length_listbox, 2)
excel_row = 2	'Declaring variable prior to the next for...next loop

For i = 1 To num_of_days

	If available_dates_array(i, 0) = 1 Then		'These are the dates that the user has determined the agency/unit/worker
		appointment_time = appt_month & "/" & i & "/" & appt_year & " " & first_appointment_listbox		'putting together the date and time values.
		Do
			appointment_time = DateAdd("N", 0, appointment_time)	'Putting the date in a MM/DD/YYYY HH:MM format. It just looks nicer.
			appointment_time_for_viewing = appointment_time			'creating a new variable to handle the display of time to get it out of military time.
			If DatePart("H", appointment_time_for_viewing) >= 13 Then appointment_time_for_viewing = DateAdd("H", -12, appointment_time_for_viewing)
			For j = 1 To appointments_per_time_slot					'Having the script create appointments_per_time_slot for each day and time.
				objExcel.Cells(excel_row, interview_time_col).value = appointment_time_for_viewing
				excel_row = excel_row + 1
				reviews_total = reviews_total + 1	'Adds one to the reviews total (will end when the loop starts if we're at the reviews total)
				If objExcel.Cells(excel_row, case_number_col).value = "" Then Exit For
			Next
			If objExcel.Cells(excel_row, case_number_col).value = "" Then Exit Do

			'This is where the script adds minutes for the next appointment.
			appointment_time = DateAdd("N", appointment_length_listbox, appointment_time)
			appointment_time = DateAdd("N", 0, appointment_time) 'Putting the date in a MM/DD/YYYY HH:MM format. Otherwise, the format is M/D/YYYY. It just looks nicer.

			'The variables "last_appointment_listbox_for_comparison" and "appointment_time_for_comparison" are used for the DO-LOOP. Because the script
			'handles time in military time, but clients do not, we need a way of handling the display of the date/time and the comparison of appointment times
			'against last appointment time variable.
			If DatePart("H", last_appointment_listbox) < 7 Then
				last_appointment_listbox_for_comparison = DateAdd("H", 12, last_appointment_listbox)
			Else
				last_appointment_listbox_for_comparison = last_appointment_listbox
			End If

			If DatePart("H", appointment_time) < 7 Then
				appointment_time_for_comparison = DateAdd("H", 12, appointment_time)
			Else
				appointment_time_for_comparison = appointment_time
			End If
		Loop UNTIL (DatePart("H", appointment_time_for_comparison) > DatePart("H", last_appointment_listbox_for_comparison)) Or ((DatePart("H", appointment_time_for_comparison) >= DatePart("H", last_appointment_listbox_for_comparison)) And DatePart("N", appointment_time_for_comparison) > DatePart("N", last_appointment_listbox_for_comparison))

		If objExcel.Cells(excel_row, case_number_col).value = "" Then Exit For	'If there's nothing in the row, it means it's over.

		If alt_first_appointment_listbox <> "Select one..." Then 	'Same as above but for the second block 'o time
			appointment_time = appt_month & "/" & i & "/" & appt_year & " " & alt_first_appointment_listbox
			Do
				appointment_time = DateAdd("N", 0, appointment_time)	'Putting the date in a MM/DD/YYYY HH:MM format. It just looks nicer.
				appointment_time_for_viewing = appointment_time			'creating a new variable to handle the display of time to get it out of military time.
				If DatePart("H", appointment_time_for_viewing) >= 13 Then appointment_time_for_viewing = DateAdd("H", -12, appointment_time_for_viewing)
				For k = 1 To alt_appointments_per_time_slot					'Having the script create appointments_per_time_slot for each day and time.
					objExcel.Cells(excel_row, interview_time_col).value = appointment_time_for_viewing
					excel_row = excel_row + 1
					If objExcel.Cells(excel_row, case_number_col).value = "" Then Exit For
				Next
				If objExcel.Cells(excel_row, case_number_col).value = "" Then Exit Do

				'This is where the script adds minutes for the next appointment.
				appointment_time = DateAdd("N", alt_appointment_length_listbox, appointment_time)
				appointment_time = DateAdd("N", 0, appointment_time) 'Putting the date in a MM/DD/YYYY HH:MM format. Otherwise, the format is M/D/YYYY. It just looks nicer.

				'The variables "last_appointment_listbox_for_comparison" and "appointment_time_for_comparison" are used for the DO-LOOP. Because the script
				'handles time in military time, but clients do not, we need a way of handling the display of the date/time and the comparison of appointment times
				'against last appointment time variable.
				If DatePart("H", alt_last_appointment_listbox) < 7 Then
					last_appointment_listbox_for_comparison = DateAdd("H", 12, alt_last_appointment_listbox)
				Else
					last_appointment_listbox_for_comparison = alt_last_appointment_listbox
				End If

				If DatePart("H", appointment_time) < 7 Then
					appointment_time_for_comparison = DateAdd("H", 12, appointment_time)
				Else
					appointment_time_for_comparison = appointment_time
				End If
			Loop UNTIL (DatePart("H", appointment_time_for_comparison) > DatePart("H", last_appointment_listbox_for_comparison)) Or ((DatePart("H", appointment_time_for_comparison) >= DatePart("H", last_appointment_listbox_for_comparison)) And DatePart("N", appointment_time_for_comparison) > DatePart("N", last_appointment_listbox_for_comparison))
		End If
		If objExcel.Cells(excel_row, case_number_col).Value = "" Then Exit For		'Because we're all out of cases at this point
	End If
Next

'Formatting the columns to autofit after they are all finished being created.
objExcel.Columns(worker_number_col).autofit()
objExcel.Columns(case_number_col).autofit()
objExcel.Columns(interview_time_col).autofit()

'VbOKCancel allows the user to review and confirm the scheduled appointments prior to case notes, TIKL's and MEMOs being sent
recertificaiton_date_confirmation = MsgBox("Please review your Excel spreadsheet carefully to ensure that the dates and times are accurate." & vbNewLine & vbNewLine & "Press OK to confirm the dates and times of your scheduled appointments. Press cancel to stop the script.", vbOKCancel + vbExclamation, "Please review and confirm")
If recertificaiton_date_confirmation = vbCancel Then script_end_procedure("The script has ended. No appointments have been scheduled, or appointment letters have been issued.")

excel_row = 2					'resetting excel row to start reading at the top
Do 								'looping until it meets a blank excel cell without a case number
	recert_status = ""			'resetting recert status for each run through the loop/case number
	forms_to_arep = ""
	forms_to_swkr = ""
	MAXIS_case_number = objExcel.cells(excel_row, case_number_col).value
	interview_time = objExcel.Cells(excel_row, interview_time_col).value
	If DatePart("H", interview_time) < 7 Or DatePart("H", interview_time) = 12 Then    'converting from military time
		am_pm = "PM"
	Else
		am_pm = "AM"
	End If
	appt_minute_place_holder_because_reasons = DatePart("N", interview_time)
	If appt_minute_place_holder_because_reasons = "0" Then appt_minute_place_holder_because_reasons = "00"	'This is needed because DatePart("N", 10:00) = 0 and not 00. Times were being displayed 10:0
	interview_time = DatePart("M", interview_time) & "/" & DatePart("D", interview_time) & "/" & DatePart("YYYY", interview_time) & " " & DatePart("H", interview_time) & ":" & appt_minute_place_holder_because_reasons & " " & am_pm
	If MAXIS_case_number = "" Then Exit Do      'exiting do if it finds a blank cell on the case number column

	back_to_self
	EMwritescreen CM_mo, 20, 43			'writing current month
	EMwritescreen CM_yr, 20, 46		'writing current year
	transmit

	If county_code <> "x127" Then
		'Grabbing the phone number from ADDR
		Call navigate_to_MAXIS_screen("STAT", "ADDR")
		EMReadScreen area_code, 3, 17, 45
		EMReadScreen remaining_digits, 9, 17, 50
		If area_code = "___" Then 'Reading phone 2 in case it is the only entered number
			EMReadScreen area_code, 3, 18, 45
			EMReadScreen remaining_digits, 9, 18, 50
		End If
		If area_code = "___" Then
			EMReadScreen area_code, 3, 19, 45 ' reading phone 3
			EMReadScreen remaining_digits, 9, 19, 50
		End If
		phone_number = area_code & remaining_digits
		phone_number = Replace(phone_number, "_", " ") 'replaces _ to blank space so it can work with if statements looking for no phone numbers which looks for 12 spaces
	End If

	back_to_self
	If developer_mode = True Then
		Call navigate_to_MAXIS_screen("SPEC", "MEMO")
		'Checking for AREP if found sending memo to them as well
		'AREP/ALTP DISPLAY NOT CURRENTLY WORKING IN DEV MODE, THIS IS SOMETHING THAT SHOULD BE ADDED

		Memo_to_display = "The State DHS sent you a packet of paperwork. This is renewal paperwork for your SNAP case. Your SNAP case is set to close on " &  last_day_of_recert & ". Please sign, date and return your renewal paperwork by: " & Left(CM_plus_1_mo, 2) & "/08/" & Right(CM_plus_1_yr, 2) & "." & vbNewLine &_
		"You must also do an interview for your SNAP case to continue. Your phone interview is scheduled for: " & interview_time & "." & vbNewLine
		If county_code = "x127" Then    'allows for county 27 to have clients call them.
			Memo_to_display = Memo_to_display & "The phone number for you to call is " & contact_phone_number & "."
		Else
			If phone_number <> "            " Then
				Memo_to_display = Memo_to_display & "The phone number we have for you is: " & phone_number & ". This is the number we will call." & vbNewLine
			Else
				Memo_to_display = Memo_to_display & "We currently do not have a phone number on file for you." & vbNewLine &_
				"Please call us at " & contact_phone_number & " to update your phone number, or if you would prefer an in-person interview." & vbNewLine
			End If
		End If

		Memo_to_display = Memo_to_display & "Important: We must have your renewal paperwork to do your interview. Please send proofs with your renewal paperwork." & vbNewline &_
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
		If phone_number = "            " Then
			Case_note_to_display = Case_note_to_display & "No phone number in MAXIS as of " & Date & "." & vbNewLine
		Else
			Case_note_to_display = Case_note_to_display & "* Client phone: " & phone_number & vbNewLine
		End If
		If forms_to_arep = "Y" Then Case_note_to_display = Case_note_to_display & "* Copy of notice sent to AREP." & vbNewLine
		If forms_to_swkr = "Y" Then Case_note_to_display = Case_note_to_display & "* Copy of notice sent to Social Worker." & vbNewLine
		Case_note_to_display = Case_note_to_display & "---" & vbNewLine & worker_signature

		MsgBox Case_note_to_display

        tikl_date = DatePart("M", interview_time) & "/" & DatePart("D", interview_time) & "/" & Right(DatePart("YYYY", interview_time), 2)
		MsgBox MAXIS_case_number & vbnewLine & "Dail: ~*~*~CLIENT HAD RECERT INTERVIEW APPT AT " & interview_time & ". IF MISSED SEND NOMI." & vbNewLine & _
			"tikl date: " & tikl_date

	Else			'ELSE in this case is LIVE cases, not testing in developer mode

		'County Code - Added - 748 to 779
		Call navigate_to_MAXIS_screen("SPEC", "MEMO")
		PF5
		EMReadScreen memo_display_check, 12, 2, 33
		If memo_display_check = "Memo Display" Then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
		'Checking for AREP if found sending memo to them as well
		row = 4
		col = 1
		EMSearch "ALTREP", row, col
		If row > 4 Then
			arep_row = row
			Call navigate_to_MAXIS_screen("STAT", "AREP")
			EMReadscreen forms_to_arep, 1, 10, 45
			Call navigate_to_MAXIS_screen("SPEC", "MEMO")
			PF5
		End If

		'Checking for SWKR if found sending MEMO to them as well
		row = 4
		col = 1
		EMSearch "SOCWKR", row, col
		If row > 4 Then
			swkr_row = row
			Call navigate_to_MAXIS_screen("STAT", "SWKR")
			EMReadscreen forms_to_swkr, 1, 15, 63
			Call navigate_to_MAXIS_screen("SPEC", "MEMO")
			PF5
		End If

		EMWriteScreen "x", 5, 10
		If forms_to_arep = "Y" Then EMWriteScreen "x", arep_row, 10
		If forms_to_swkr = "Y" Then EMWriteScreen "x", swkr_row, 10
		transmit

		'Navigating to SPEC/MEMO and starting a new memo
		start_a_new_spec_memo

		'Writing the appointment and letter into a memo
		EMSendKey("************************************************************")
		Call write_variable_in_SPEC_MEMO("The State DHS sent you a packet of paperwork. This is renewal paperwork for your SNAP case. Your SNAP case is set to close on " &  last_day_of_recert & ". Please sign, date and return your renewal paperwork by " & Left(CM_plus_1_mo, 2) & "/08/" & Right(CM_plus_1_yr, 2) & ".")
		Call write_variable_in_SPEC_MEMO("")
		Call write_variable_in_SPEC_MEMO("You must also do an interview for your SNAP case to continue.")
		Call write_variable_in_SPEC_MEMO("")
		If county_code = "x127" Then    'allows for county 27 to have clients call them.
			Call write_variable_in_SPEC_MEMO("Your phone interview is scheduled for: " & interview_time & ". The phone number for you to call is: " & contact_phone_number & ".")
		Else
			If phone_number <> "            " Then
				Call write_variable_in_SPEC_MEMO("Your phone interview is scheduled for: " & interview_time & ". The phone number we have for you is: " & phone_number & ". This is the number we will call.")
			Else
				Call write_variable_in_SPEC_MEMO("Your phone interview is scheduled for: " & interview_time & ". We currently do not have a phone number on file for you.")
				Call write_variable_in_SPEC_MEMO("Please call us at: " & contact_phone_number & " to update your phone number, or if you would prefer an in-person interview.")
			End If
		End If
		Call write_variable_in_SPEC_MEMO("")
		Call write_variable_in_SPEC_MEMO("IMPORTANT: We must have your renewal paperwork to do your interview.")
		Call write_variable_in_SPEC_MEMO("")
		Call write_variable_in_SPEC_MEMO("Please send proofs with your renewal paperwork.")
		Call write_variable_in_SPEC_MEMO(" * Examples of income proofs: paystubs, income reports, business ledgers, income tax forms, etc.")
		Call write_variable_in_SPEC_MEMO("")
		Call write_variable_in_SPEC_MEMO(" * Examples of housing cost proofs(if changed): rent/house payment receipt, mortgage, lease, etc.")
		Call write_variable_in_SPEC_MEMO("")
		Call write_variable_in_SPEC_MEMO(" * Examples of medical cost proofs(if changed): Prescription and medical bills, etc.")
		Call write_variable_in_SPEC_MEMO("")
		Call write_variable_in_SPEC_MEMO("Please call us at " & contact_phone_number & " if you need to:")
		Call write_variable_in_SPEC_MEMO(" * Reschedule your appointment.")
		Call write_variable_in_SPEC_MEMO(" * Report a new phone number, or other changes.")
		Call write_variable_in_SPEC_MEMO(" * Request an in-person interview.")
		PF4
		back_to_self

		'case noting appointment time and date
		start_a_blank_case_note
		EMSendKey "***SNAP Recertification Interview Scheduled***"
		Call write_variable_in_case_note("* A phone interview has been scheduled for " & interview_time & ".")
		If phone_number = "            " Then
			Call write_variable_in_case_note("No phone number in MAXIS as of " & Date & ".")
		Else
			Call write_variable_in_case_note("* Client phone: " & phone_number)
		End If
		If forms_to_arep = "Y" Then Call write_variable_in_case_note("* Copy of notice sent to AREP.")
		If forms_to_swkr = "Y" Then Call write_variable_in_case_note("* Copy of notice sent to Social Worker.")
		Call write_variable_in_case_note("---")
		Call write_variable_in_case_note(worker_signature)

		'TIKLing to remind the worker to send NOMI if appointment is missed.
		Call navigate_to_MAXIS_screen("DAIL", "WRIT")
        tikl_date = DatePart("M", interview_time) & "/" & DatePart("D", interview_time) & "/" & DatePart("YYYY", interview_time)
		Call create_MAXIS_friendly_date(tikl_date, 0, 5, 18)
		EMWriteScreen "~*~*~CLIENT HAD RECERT INTERVIEW APPT AT: " & interview_time, 9, 3
		EMWriteScreen "IF MISSED SEND NOMI", 10, 3
		transmit
		PF3
	End If

'adding appointment time and date to outlook calendar if requested by worker County Code 841 to 860
	If outlook_calendar_check = 1 Then
		appt_date_for_outlook = DatePart("M", interview_time) & "/" & DatePart("D", interview_time) & "/" & DatePart("YYYY", interview_time)
		appt_time_for_outlook = DatePart("H", interview_time) & ":" & DatePart("N", interview_time)
		If DatePart("N", interview_time) = 0 Then appt_time_for_outlook = DatePart("H", interview_time) & ":00"
		appt_end_time_for_outlook = DateAdd("N", appointment_length_listbox, interview_time)
		appt_end_time_for_outlook = DatePart("H", appt_end_time_for_outlook) & ":" & DatePart("N", appt_end_time_for_outlook)
		If DatePart("N", interview_time) = 0 Then appt_end_time_for_outlook = DatePart("H", appt_end_time_for_outlook) & ":00"
		appointment_subject = "SNAP RECERT"
		appointment_body = "Case Number: " & MAXIS_case_number
		If phone_number = "            " Then
			appointment_location = "No phone number in MAXIS as of " & Date & "."
		Else
			appointment_location = "Phone: " & phone_number
		End If
		appointment_reminder = 5	'5 minutes
		appointment_category = "Recertification Interview"
		'using the variables created above to generate the Outlook Appointment from the custom function.
		Call create_outlook_appointment(appt_date_for_outlook, appt_time_for_outlook, appt_end_time_for_outlook, appointment_subject, appointment_body, appointment_location, TRUE, appointment_reminder, appointment_category)
	End If


	excel_row = excel_row + 1
Loop until objExcel.cells(excel_row, case_number_col).value = "" Or objExcel.cells(excel_row, interview_time_col).value = "SKIPPED" 'If this was skipped it needs to stop here

'Formatting the columns to autofit after they are all finished being created.
objExcel.Columns(1).autofit()
objExcel.Columns(2).autofit()
objExcel.Columns(3).autofit()
objExcel.Columns(4).autofit()

'Creating the list of privileged cases and adding to the spreadsheet
If priv_case_list <> "" Then
	priv_case_list = Right(priv_case_list, (Len(priv_case_list)-1))
	prived_case_array = Split(priv_case_list, "|")
	excel_row = 2


	For Each MAXIS_case_number in prived_case_array
		objExcel.cells(excel_row, privileged_case_col).value = MAXIS_case_number
		excel_row = excel_row + 1
	Next
End If
script_end_procedure("Success! The Excel file now has all of the cases that have had interviews scheduled.  Please manually review the list of privileged cases.")
