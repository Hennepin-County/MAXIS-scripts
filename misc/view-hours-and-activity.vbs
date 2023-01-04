'Required for statistical purposes==========================================================================================
name_of_script = "TASKS - DASH.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================
run_locally = true		'this script does NOT open Global Variables. Setting the runLocally here.
'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
		FuncLib_URL = "C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'FUNCTIONS==================================================================================================================
function format_time_variable(time_variable, is_this_from_excel)
	If is_this_from_excel = True Then time_variable = time_variable * 24
	time_hour = Int(time_variable)
	time_minute = time_variable - time_hour
	' MsgBox time_mi6ute
	time_minute = time_minute * 60
	time_variable = TimeSerial(time_hour, time_minute, 0)
end function

function make_time_string(time_variable)
	the_hour = Int(time_variable)
	the_min = time_variable - the_hour
	the_min = the_min * 60
	the_min = Round(the_min)
	time_variable = the_hour & " hr, " & the_min & " min"
end function

function create_time_spent_totals(start_date, end_date, sort_type, total_hours, hours_in_meetings, TYPE_ARRAY)
	ReDim TYPE_ARRAY(type_last_const, 0)
	TYPE_ARRAY(type_detail_const, 0) = ""
	TYPE_ARRAY(total_hours_const, 0) = 0
	TYPE_ARRAY(total_hours_string_const, 0) = ""
	TYPE_ARRAY(type_url_const, 0) = ""
	TYPE_ARRAY(type_last_const, 0) = ""

	type_counter = 0

	total_hours = 0
	hours_in_meetings = 0

	button_counter = 100

	For logged_activity = 0 to UBound(TIME_TRACKING_ARRAY, 2)
		If DateDiff("d", start_date, TIME_TRACKING_ARRAY(activity_date_const, logged_activity)) >= 0 AND DateDiff("d",  TIME_TRACKING_ARRAY(activity_date_const, logged_activity),end_date) >= 0 AND TIME_TRACKING_ARRAY(activity_paid_yn, logged_activity) = "Y" Then
			total_hours = total_hours + TIME_TRACKING_ARRAY(activity_time_spent_val, logged_activity)
			If TIME_TRACKING_ARRAY(activity_meeting, logged_activity) = "Yes" Then hours_in_meetings = hours_in_meetings + TIME_TRACKING_ARRAY(activity_time_spent_val, logged_activity)
			type_found = false
			For each_type = 0 to UBound(TYPE_ARRAY, 2)
				If sort_type = "CATEGORY" Then
					If TYPE_ARRAY(type_detail_const, each_type) = TIME_TRACKING_ARRAY(activity_category, logged_activity) Then type_found = True
					If TYPE_ARRAY(type_detail_const, each_type) = "BLANK" AND TIME_TRACKING_ARRAY(activity_category, logged_activity) = "" Then type_found = True
				End If
				If sort_type = "PROJECT" Then
					If trim(TIME_TRACKING_ARRAY(activity_project, logged_activity)) = "" Then
						If TYPE_ARRAY(type_detail_const, each_type) = "BLANK" Then type_found = True
					Else
						If TYPE_ARRAY(type_detail_const, each_type) = TIME_TRACKING_ARRAY(activity_project, logged_activity) Then type_found = True
					End If
				End If
				If sort_type = "GITHUB ISSUE" Then
					If trim(TIME_TRACKING_ARRAY(activity_gh_issue_numb, logged_activity)) = "" Then
						If TYPE_ARRAY(type_detail_const, each_type) = "BLANK" Then type_found = True
					Else
						If TYPE_ARRAY(type_detail_const, each_type) = TIME_TRACKING_ARRAY(activity_gh_issue_numb, logged_activity) Then type_found = True
					End If
				End If
				If sort_type = "DAY" Then
					If TYPE_ARRAY(type_detail_const, each_type) = TIME_TRACKING_ARRAY(activity_date_const, logged_activity) Then type_found = True
				End If
				If type_found = True Then
					TYPE_ARRAY(total_hours_const, each_type) = TYPE_ARRAY(total_hours_const, each_type) + TIME_TRACKING_ARRAY(activity_time_spent_val, logged_activity)
					this_one = each_type
					Exit For
				End If
			Next
			If type_found = False Then
				ReDim Preserve TYPE_ARRAY(type_last_const, type_counter)
				If sort_type = "CATEGORY" Then TYPE_ARRAY(type_detail_const, type_counter) = TIME_TRACKING_ARRAY(activity_category, logged_activity)
				If sort_type = "PROJECT" Then TYPE_ARRAY(type_detail_const, type_counter) = TIME_TRACKING_ARRAY(activity_project, logged_activity)
				If sort_type = "GITHUB ISSUE" Then TYPE_ARRAY(type_detail_const, type_counter) = TIME_TRACKING_ARRAY(activity_gh_issue_numb, logged_activity)
				If sort_type = "DAY" Then TYPE_ARRAY(type_detail_const, type_counter) = TIME_TRACKING_ARRAY(activity_date_const, logged_activity)

				If TYPE_ARRAY(type_detail_const, type_counter) = "" Then TYPE_ARRAY(type_detail_const, type_counter) = "BLANK"

				TYPE_ARRAY(total_hours_const, type_counter) = TIME_TRACKING_ARRAY(activity_time_spent_val, logged_activity)
				If sort_type = "GITHUB ISSUE" Then
					TYPE_ARRAY(type_url_const, type_counter) = TIME_TRACKING_ARRAY(activity_gh_issue_url, logged_activity)
					TYPE_ARRAY(type_btn_const, type_counter) = button_counter + type_counter
				End If
				this_one = type_counter

				type_counter = type_counter + 1
			End If
		End If
	Next
	For each_type = 0 to UBound(TYPE_ARRAY, 2)
		If TYPE_ARRAY(total_hours_const, each_type) = "" Then TYPE_ARRAY(total_hours_const, each_type) = 0
		TYPE_ARRAY(total_hours_string_const, each_type) = TYPE_ARRAY(total_hours_const, each_type)
		Call make_time_string(TYPE_ARRAY(total_hours_string_const, each_type))
		If sort_type = "PROJECT" AND TYPE_ARRAY(type_detail_const, each_type) = "BLANK" Then TYPE_ARRAY(type_detail_const, each_type) = "No Specified Project"
		If sort_type = "GITHUB ISSUE" AND TYPE_ARRAY(type_detail_const, each_type) = "BLANK" Then TYPE_ARRAY(type_detail_const, each_type) = "No Specified Issue"
		If sort_type = "DAY" Then
			If IsDate(TYPE_ARRAY(type_detail_const, each_type)) = False Then TYPE_ARRAY(type_detail_const, each_type) = start_date
			TYPE_ARRAY(type_detail_const, each_type) = TYPE_ARRAY(type_detail_const, each_type) & " - " & WeekdayName(WeekDay(TYPE_ARRAY(type_detail_const, each_type)))
		End If
	Next
end function
'===========================================================================================================================
the_year = DatePart("yyyy", date)

pay_period_begin_date = #9/11/2022#		'hard coded as a known start date to a pay period. The script will loop forward to find the current ppay period.
current_pay_period = pay_period_begin_date
Do while DateDiff("d", current_pay_period, date) > 13
	current_pay_period = DateAdd("d", 14, current_pay_period)
Loop

current_pay_period_start = current_pay_period
current_pay_period_end = DateAdd("d", 13, current_pay_period)

active_period_start = DateAdd("ww", -6, current_pay_period_start)
active_period_end = current_pay_period_end
first_date = active_period_start & ""

'DECLARATIONS===============================================================================================================
'Lists for Dialogs
month_list = "Select"
week_list = "Select"
sunday_of_current_week = date
Do while WeekdayName(Weekday(sunday_of_current_week)) <> "Sunday"
	sunday_of_current_week = DateAdd("d", -1, sunday_of_current_week)
Loop
sunday_date = sunday_of_current_week
saturday_date = DateAdd("d", 6, sunday_date)

Do																				'finding each week
	sunday_month = MonthName(DatePart("m", sunday_date))
	saturday_month = MonthName(DatePart("m", saturday_date))
	If InStr(month_list, sunday_month) = 0 Then month_list = month_list+chr(9)+sunday_month
	If InStr(month_list, saturday_month) = 0 Then month_list = month_list+chr(9)+saturday_month
	first_month = sunday_month			'this is going to keep being redefined until we get to the first one'

	week_string = sunday_date & " - " & saturday_date
	week_list = week_list+chr(9)+week_string
	first_week = week_string & ""			'this is going to keep being redefined until we get to the first one'

	sunday_date = DateAdd("d", -7, sunday_date)
	saturday_date = DateAdd("d", 6, sunday_date)
Loop Until DateDiff("d", sunday_date, active_period_start) > 0
week_array = split(week_list, chr(9))

biweek_list = "Select"
sunday_date = current_pay_period_start
saturday_date = current_pay_period_end
Do																				'finding each pay period'
	biweek_string = sunday_date & " - " & saturday_date
	biweek_list = biweek_list+chr(9)+biweek_string
	first_pay_pd = biweek_string & ""			'this is going to keep being redefined until we get to the first one'

	sunday_date = DateAdd("d", -14, sunday_date)
	saturday_date = DateAdd("d", 13, sunday_date)
Loop Until DateDiff("d", sunday_date, active_period_start) > 0
biweek_array = split(biweek_list, chr(9))

'Constants and arrays
const activity_date_const 		= 00
const activity_start_time		= 01
const activity_end_time			= 02
const activity_time_spent		= 03
const activity_category			= 04
const activity_meeting			= 05
const activity_detail			= 06
const activity_gh_issue_numb	= 07
const activity_gh_issue_url		= 08
const activity_project			= 09
const activity_paid_yn			= 10
const activity_time_spent_val	= 11
const moved_item				= 12
const item_xlrow				= 13
const last_const 				= 14

Dim TIME_TRACKING_ARRAY()
ReDim TIME_TRACKING_ARRAY(last_const, 0)

const type_detail_const	= 0
const total_hours_const	= 1
const type_url_const 	= 2
const type_btn_const 	= 3
const total_hours_string_const = 4
const type_last_const 	= 5

Dim CATEGORY_ARRAY()
ReDim CATEGORY_ARRAY(type_last_const, 0)

Dim PROJECT_ARRAY()
ReDim PROJECT_ARRAY(type_last_const, 0)

Dim GITHUB_ISSUE_ARRAY()
ReDim GITHUB_ISSUE_ARRAY(type_last_const, 0)

Dim DAY_SORT_ARRAY()
ReDim DAY_SORT_ARRAY(type_last_const, 0)

'constants for the view changes in the dialog
const day_view = 1
const week_view = 2
const biweek_view = 3
const month_view = 4
const custom_view = 5

'buttons
day_button = 1001
week_button = 1002
pay_period_button = 1003
month_button = 1004
custom_time_button = 1005

category_button = 2001
project_button = 2002
git_hub_issue_button = 2003
day_sort_button = 2004

show_excel_button = 5000
hide_excel_button = 5001
'===========================================================================================================================

'Defining the excel files for when running the script
excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Time Tracking"

If user_ID_for_validation = "CALO001" Then
	t_drive_excel_file_path = excel_file_path & "\Casey Time Tracking.xlsx"
	my_docs_excel_file_path = user_myDocs_folder & "Casey Time Tracking.xlsx"
End If
If user_ID_for_validation = "ILFE001" Then
	t_drive_excel_file_path = excel_file_path & "\Ilse Time Tracking.xlsx"
	my_docs_excel_file_path = user_myDocs_folder & "Ilse Time Tracking.xlsx"
End If

view_excel = True		'this variable allows us to set
Call excel_open(my_docs_excel_file_path, view_excel, False, ObjExcel, objWorkbook)		'opening the excel file
ObjExcel.worksheets("Active Time Tracking").Activate

row_filled_with_end_time = " "
old_items_to_move = False

'Here we read the entire excel file and save it into an array
excel_row = 2			'start of the excel file information
activity_count = 0		'starting of the counter of the array
added_end_time_row_list = " "
Do
	ReDim Preserve TIME_TRACKING_ARRAY(last_const, activity_count)				'resize the array
	TIME_TRACKING_ARRAY(activity_date_const, activity_count) 	= ObjExcel.Cells(excel_row, 1).Value				'date of the activity
	TIME_TRACKING_ARRAY(activity_start_time, activity_count) 	= ObjExcel.Cells(excel_row, 2).Value				'start time of the activity
	TIME_TRACKING_ARRAY(activity_end_time, activity_count) 		= ObjExcel.Cells(excel_row, 3).Value				'the end time of the activity
	If TIME_TRACKING_ARRAY(activity_end_time, activity_count) = "" Then			'If there is no end time we put in the current time so that math works
		curr_hour = DatePart("h", time)
		curr_min = DatePart("n", time)
		ObjExcel.Cells(excel_row, 3).Value = TimeSerial(curr_hour, curr_min, 0)
		ObjExcel.Cells(excel_row, 4).Value = "=TEXT([@[End Time]]-[@[Start Time]],"+chr(34)+"h:mm"+chr(34)+")"		'adding the calculation of elapsed time for a line that didn't have an end time
		row_filled_with_end_time = row_filled_with_end_time & excel_row & " "	'saving this so we can remove it later if the file is left open
		TIME_TRACKING_ARRAY(activity_end_time, activity_count) 		= ObjExcel.Cells(excel_row, 3).Value
		added_end_time_row_list = added_end_time_row_list & excel_row & " "
	End If
	TIME_TRACKING_ARRAY(activity_time_spent, activity_count) 	= ObjExcel.Cells(excel_row, 4).Value				'the elapsed time in a format that can be read
	If TIME_TRACKING_ARRAY(activity_time_spent, activity_count) <> "" Then
		time_spent_hour = DatePart("h", TIME_TRACKING_ARRAY(activity_time_spent, activity_count))					'here we create a number of the time spend so we can add it together
		time_spent_min = DatePart("n", TIME_TRACKING_ARRAY(activity_time_spent, activity_count))
		time_spent_min = time_spent_min/60
		TIME_TRACKING_ARRAY(activity_time_spent_val, activity_count) = time_spent_hour + time_spent_min				'saving the time spent value into the array
	End If
	TIME_TRACKING_ARRAY(activity_category, activity_count) 		= ObjExcel.Cells(excel_row, 5).Value				'the activity category
	TIME_TRACKING_ARRAY(activity_meeting, activity_count) 		= ObjExcel.Cells(excel_row, 6).Value				'the Yes/No of if this activity is a meeting
	TIME_TRACKING_ARRAY(activity_detail, activity_count) 		= ObjExcel.Cells(excel_row, 7).Value				'the detail of the activity
	TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count) = trim(ObjExcel.Cells(excel_row, 8).Value)			'the Git Hub issue information
	If TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count) <> "" AND InStr(TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count), "&") = 0 AND InStr(TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count), ",") = 0 AND InStr(TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count), "/") = 0 AND InStr(TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count), "\") = 0 AND ucase(trim(TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count))) <> "MULTIPLE" Then
		' ObjExcel.Cells(excel_row, 8).Value = ""
		TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count) = replace(TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count), "#", "")			''here we are reading the GH Issue information and making sure we are reading only a number and then making it a URL
		TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count) = replace(TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count), "Issue", "")
		TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count) = trim(TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count))
		TIME_TRACKING_ARRAY(activity_gh_issue_url, activity_count) = "https://github.com/Hennepin-County/MAXIS-scripts/issues/" & TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count)
		' ObjExcel.Cells(excel_row, 8).Value = "=HYPERLINK(" & chr(34) & TIME_TRACKING_ARRAY(activity_gh_issue_url, activity_count) & chr(34) & ", " & chr(34) & TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count) & chr(34) & ")"
	End If
	TIME_TRACKING_ARRAY(activity_project, activity_count) 		= trim(ObjExcel.Cells(excel_row, 9).Value)			'the project of the activity
	TIME_TRACKING_ARRAY(activity_paid_yn, activity_count) 		= ObjExcel.Cells(excel_row, 10).Value				'if this is paid time
	TIME_TRACKING_ARRAY(moved_item, activity_count) = False
	TIME_TRACKING_ARRAY(item_xlrow, activity_count) = excel_row

	TIME_TRACKING_ARRAY(activity_date_const, activity_count) = DateAdd("d", 0, TIME_TRACKING_ARRAY(activity_date_const, activity_count))
	If DateDiff("d", TIME_TRACKING_ARRAY(activity_date_const, activity_count), active_period_start) > 0 Then old_items_to_move = True

	activity_count = activity_count + 1					'incrementing the array
	excel_row = excel_row + 1							'go to the next excel row
	next_row_date = ObjExcel.Cells(excel_row, 1).Value	'reading if there is more information on the next row.
Loop until next_row_date = ""

'now that the reading is done, we are going to make the Excel file visible.
'This is particularly for TESTING and we can remove this in the future
view_excel = True
objExcel.Visible = view_excel

If old_items_to_move = True Then			'If we find that some of the activities are older, we want to take them back.
	'This part is to ensure this script is futureproofed to add sheets for each year as it occurs
	current_year_worksheet_found = False										'defaulting that the current year sheet has been found or not.
	sheet_name = the_year & ""													'making the current year a string for a name usage
	last_year_sheet = DatePart("yyyy", DateAdd("yyyy", -1, date)) & ""			'also identifying last year's sheet name because that is what we use to move the new sheet correctly.
	For Each objWorkSheet In objWorkbook.Worksheets								'look at all the sheets that currently exist
		If objWorkSheet.Name = sheet_name Then current_year_worksheet_found = True		'If a sheet has been found with the current year as the anme, we set this boolean to true
	Next
	If current_year_worksheet_found = False Then								'if we looked at all of the sheets and the current year was not found, we will add it here
		ObjExcel.Worksheets.Add().Name = sheet_name								'adding a sheet with the current year as the name
		ObjExcel.worksheets(sheet_name).Move ObjExcel.worksheets(last_year_sheet)	'moving this sheet to be just before last year's sheet
	End If
	'now we are going to move individual activity entries that are 'too old' to the correct sheet by year.
	TIME_TRACKING_ARRAY(activity_date_const, 0) = DateAdd("d", 0, TIME_TRACKING_ARRAY(activity_date_const, 0))
	year_to_check = DatePart("yyyy", TIME_TRACKING_ARRAY(activity_date_const, 0))
	year_to_check = year_to_check & ""
	ObjExcel.worksheets(year_to_check).Activate
	year_sheet_xl_row = 2
	Do while ObjExcel.Cells(year_sheet_xl_row, 1).Value <> ""
		year_sheet_xl_row = year_sheet_xl_row + 1
	Loop

	For each_activity = 0 to UBound(TIME_TRACKING_ARRAY, 2)						'saving the information into Excel'
		TIME_TRACKING_ARRAY(activity_date_const, each_activity) = DateAdd("d", 0, TIME_TRACKING_ARRAY(activity_date_const, each_activity))
		If DateDiff("d", TIME_TRACKING_ARRAY(activity_date_const, each_activity), active_period_start) > 0 Then
			TIME_TRACKING_ARRAY(moved_item, each_activity) = True
			ObjExcel.Cells(year_sheet_xl_row, 1).Value = TIME_TRACKING_ARRAY(activity_date_const, each_activity)
			ObjExcel.Cells(year_sheet_xl_row, 2).Value = TIME_TRACKING_ARRAY(activity_start_time, each_activity)
			ObjExcel.Cells(year_sheet_xl_row, 3).Value = TIME_TRACKING_ARRAY(activity_end_time, each_activity)
			ObjExcel.Cells(year_sheet_xl_row, 4).Value = "=TEXT([@[End Time]]-[@[Start Time]],"+chr(34)+"h:mm"+chr(34)+")"
			ObjExcel.Cells(year_sheet_xl_row, 5).Value = TIME_TRACKING_ARRAY(activity_category, each_activity)
			ObjExcel.Cells(year_sheet_xl_row, 6).Value = TIME_TRACKING_ARRAY(activity_meeting, each_activity)
			ObjExcel.Cells(year_sheet_xl_row, 7).Value = TIME_TRACKING_ARRAY(activity_detail, each_activity)
			If TIME_TRACKING_ARRAY(activity_gh_issue_url, each_activity) = "" Then
				ObjExcel.Cells(year_sheet_xl_row, 8).Value = TIME_TRACKING_ARRAY(activity_gh_issue_numb, each_activity)
			Else
				If TIME_TRACKING_ARRAY(activity_gh_issue_numb, each_activity) <> "" Then
					ObjExcel.Cells(year_sheet_xl_row, 8).Value = "=HYPERLINK(" & chr(34) & "https://github.com/Hennepin-County/MAXIS-scripts/issues/" & TIME_TRACKING_ARRAY(activity_gh_issue_numb, each_activity) & chr(34) & ", " & chr(34) & TIME_TRACKING_ARRAY(activity_gh_issue_numb, each_activity) & chr(34) & ")"
				End If
			End If
			ObjExcel.Cells(year_sheet_xl_row, 9).Value = TIME_TRACKING_ARRAY(activity_project, each_activity)
			ObjExcel.Cells(year_sheet_xl_row, 10).Value= TIME_TRACKING_ARRAY(activity_paid_yn, each_activity)
			year_sheet_xl_row = year_sheet_xl_row + 1

		End If
	Next

	ObjExcel.worksheets("Active Time Tracking").Activate						'back to the main sheet'

	For each_activity = UBound(TIME_TRACKING_ARRAY, 2) to 0 Step -1				'deleting the old items that wwere copued FROM THE BOTTOM UP'
		If TIME_TRACKING_ARRAY(moved_item, each_activity) = True Then
			SET objRange = ObjExcel.Cells(TIME_TRACKING_ARRAY(item_xlrow, each_activity), 1).EntireRow
			objRange.Delete
		End If
	Next

	objWorkbook.Save									'saving the file to 'My Documents'
	objWorkbook.SaveAs (t_drive_excel_file_path)		'saving the file to the T Drive

	ReDim TIME_TRACKING_ARRAY(last_const, 0)									'reset the array because we are going to read it fresh after the deleting

	row_filled_with_end_time = " "

	'Here we read the entire excel file AGAIN and save it into an array
	excel_row = 2			'start of the excel file information
	activity_count = 0		'starting of the counter of the array
	added_end_time_row_list = " "
	Do
		ReDim Preserve TIME_TRACKING_ARRAY(last_const, activity_count)				'resize the array
		TIME_TRACKING_ARRAY(activity_date_const, activity_count) 	= ObjExcel.Cells(excel_row, 1).Value				'date of the activity
		TIME_TRACKING_ARRAY(activity_start_time, activity_count) 	= ObjExcel.Cells(excel_row, 2).Value				'start time of the activity
		TIME_TRACKING_ARRAY(activity_end_time, activity_count) 		= ObjExcel.Cells(excel_row, 3).Value				'the end time of the activity
		If TIME_TRACKING_ARRAY(activity_end_time, activity_count) = "" Then			'If there is no end time we put in the current time so that math works
			curr_hour = DatePart("h", time)
			curr_min = DatePart("n", time)
			ObjExcel.Cells(excel_row, 3).Value = TimeSerial(curr_hour, curr_min, 0)
			ObjExcel.Cells(excel_row, 4).Value = "=TEXT([@[End Time]]-[@[Start Time]],"+chr(34)+"h:mm"+chr(34)+")"		'adding the calculation of elapsed time for a line that didn't have an end time
			row_filled_with_end_time = row_filled_with_end_time & excel_row & " "	'saving this so we can remove it later if the file is left open
			TIME_TRACKING_ARRAY(activity_end_time, activity_count) 		= ObjExcel.Cells(excel_row, 3).Value
			added_end_time_row_list = added_end_time_row_list & excel_row & " "
		End If
		TIME_TRACKING_ARRAY(activity_time_spent, activity_count) 	= ObjExcel.Cells(excel_row, 4).Value				'the elapsed time in a format that can be read
		If TIME_TRACKING_ARRAY(activity_time_spent, activity_count) <> "" Then
			time_spent_hour = DatePart("h", TIME_TRACKING_ARRAY(activity_time_spent, activity_count))					'here we create a number of the time spend so we can add it together
			time_spent_min = DatePart("n", TIME_TRACKING_ARRAY(activity_time_spent, activity_count))
			time_spent_min = time_spent_min/60
			TIME_TRACKING_ARRAY(activity_time_spent_val, activity_count) = time_spent_hour + time_spent_min				'saving the time spent value into the array
		End If
		TIME_TRACKING_ARRAY(activity_category, activity_count) 		= ObjExcel.Cells(excel_row, 5).Value				'the activity category
		TIME_TRACKING_ARRAY(activity_meeting, activity_count) 		= ObjExcel.Cells(excel_row, 6).Value				'the Yes/No of if this activity is a meeting
		TIME_TRACKING_ARRAY(activity_detail, activity_count) 		= ObjExcel.Cells(excel_row, 7).Value				'the detail of the activity
		TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count) = trim(ObjExcel.Cells(excel_row, 8).Value)			'the Git Hub issue information
		If TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count) <> "" AND InStr(TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count), "&") = 0 AND InStr(TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count), ",") = 0 AND InStr(TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count), "/") = 0 AND InStr(TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count), "\") = 0 AND ucase(trim(TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count))) <> "MULTIPLE" Then
			' ObjExcel.Cells(excel_row, 8).Value = ""
			TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count) = replace(TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count), "#", "")			''here we are reading the GH Issue information and making sure we are reading only a number and then making it a URL
			TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count) = replace(TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count), "Issue", "")
			TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count) = trim(TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count))
			TIME_TRACKING_ARRAY(activity_gh_issue_url, activity_count) = "https://github.com/Hennepin-County/MAXIS-scripts/issues/" & TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count)
			' ObjExcel.Cells(excel_row, 8).Value = "=HYPERLINK(" & chr(34) & TIME_TRACKING_ARRAY(activity_gh_issue_url, activity_count) & chr(34) & ", " & chr(34) & TIME_TRACKING_ARRAY(activity_gh_issue_numb, activity_count) & chr(34) & ")"
		End If
		TIME_TRACKING_ARRAY(activity_project, activity_count) 		= trim(ObjExcel.Cells(excel_row, 9).Value)			'the project of the activity
		TIME_TRACKING_ARRAY(activity_paid_yn, activity_count) 		= ObjExcel.Cells(excel_row, 10).Value				'if this is paid time

		activity_count = activity_count + 1					'incrementing the array
		excel_row = excel_row + 1							'go to the next excel row
		next_row_date = ObjExcel.Cells(excel_row, 1).Value	'reading if there is more information on the next row.
	Loop until next_row_date = ""
End If

'Setting some defaults for the dialog
hours_in_time_pd = 0
hours_in_meetings_dur_time_pd = 0

current_day = date & ""
For each week_item in week_array
	temp_array = ""
	If Instr(week_item, " - ") <> 0 Then
		temp_array = split(week_item, " - ")
		temp_array(0) = DateAdd("d", 0, temp_array(0))
		temp_array(1) = DateAdd("d", 0, temp_array(1))
		If DateDiff("d", temp_array(0), date) >= 0 AND DateDiff("d", date, temp_array(1)) >= 0 Then current_week = week_item
	End If
Next
For each biweek_item in biweek_array
	temp_array = ""
	If Instr(biweek_item, " - ") <> 0 Then
		temp_array = split(biweek_item, " - ")
		temp_array(0) = DateAdd("d", 0, temp_array(0))
		temp_array(1) = DateAdd("d", 0, temp_array(1))
		If DateDiff("d", temp_array(0), date) >= 0 AND DateDiff("d", date, temp_array(1)) >= 0 Then current_pay_pd = biweek_item
	End If
Next
current_month = MonthName(DatePart("m", date))

selected_date = current_day
dialog_view = day_view
selected_sort = "DAY"
Call create_time_spent_totals(current_day, current_day, selected_sort, hours_in_time_pd, hours_in_meetings_dur_time_pd, CATEGORY_ARRAY)

'now we show the dialog.
Do
	err_msg = ""

	display_total_hours = hours_in_time_pd					'these are set to a new variable because they have to remain numbers
	display_meeting_hours = hours_in_meetings_dur_time_pd
	Call make_time_string(display_total_hours)				'this will make these strings in the format H hrs, M min
	Call make_time_string(display_meeting_hours)
	selected_start_date = selected_start_date & ""			'Making sure these are displayed
	selected_end_date = selected_end_date & ""
	selected_date = selected_date & ""

	on_current_time_pd = False
	If dialog_view = day_view AND selected_date = current_day Then on_current_time_pd = True			'determining if we need a plus button or not
	If dialog_view = week_view AND selected_date = current_week Then on_current_time_pd = True
	If dialog_view = biweek_view AND selected_date = current_pay_pd Then on_current_time_pd = True
	If dialog_view = month_view AND selected_date = current_month Then on_current_time_pd = True

	on_first_time_ppd = False
	If dialog_view = day_view AND selected_date = first_date Then on_first_time_ppd = True			'determining if we need a plus button or not
	If dialog_view = week_view AND selected_date = first_week Then on_first_time_ppd = True
	If dialog_view = biweek_view AND selected_date = first_pay_pd Then on_first_time_ppd = True
	If dialog_view = month_view AND selected_date = first_month Then on_first_time_ppd = True

	dlg_len = 140											'setting the lengths of the dialog and group boxes based on what the sort options are
	grp_1_len = 25
	grp_2_len = 105
	If selected_sort = "CATEGORY" Then
		For cat_item = 0 to UBound(CATEGORY_ARRAY, 2)
			dlg_len = dlg_len + 10
			grp_1_len = grp_1_len + 10
			grp_2_len = grp_2_len + 10
		Next
	End If
	If selected_sort = "PROJECT" Then
		For cat_item = 0 to UBound(PROJECT_ARRAY, 2)
			dlg_len = dlg_len + 10
			grp_1_len = grp_1_len + 10
			grp_2_len = grp_2_len + 10
		Next
	End If
	If selected_sort = "GITHUB ISSUE" Then
		For cat_item = 0 to UBound(GITHUB_ISSUE_ARRAY, 2)
			dlg_len = dlg_len + 10
			grp_1_len = grp_1_len + 10
			grp_2_len = grp_2_len + 10
		Next
	End If
	If selected_sort = "DAY" Then
		If dialog_view = day_view Then
			For logged_activity = 0 to UBound(TIME_TRACKING_ARRAY, 2)
				If DateDiff("d", selected_date, TIME_TRACKING_ARRAY(activity_date_const, logged_activity)) = 0 Then
					dlg_len = dlg_len + 10
					grp_1_len = grp_1_len + 10
					grp_2_len = grp_2_len + 10
				End If
			Next
		Else
			For cat_item = 0 to UBound(DAY_SORT_ARRAY, 2)
				dlg_len = dlg_len + 10
				grp_1_len = grp_1_len + 10
				grp_2_len = grp_2_len + 10
			Next
		End If
	End If

	'For real - the dialog
	BeginDialog Dialog1, 0, 0, 460, dlg_len, "View Hours and Activity"
	  ButtonGroup ButtonPressed
		GroupBox 5, 5, 385, grp_2_len, "Hours Breakdown"
		If dialog_view = day_view Then EditBox 280, 15, 50, 15, selected_date
		If dialog_view = week_view Then DropListBox 180, 15, 150, 45, week_list, selected_date
		If dialog_view = biweek_view Then DropListBox 180, 15, 150, 45, biweek_list, selected_date
		If dialog_view = month_view Then DropListBox 180, 15, 150, 45, month_list, selected_date
		If dialog_view = custom_view Then
			EditBox 230, 15, 40, 15, selected_start_date
			Text 270, 17, 10, 10, " - "
			EditBox 280, 15, 50, 15, selected_end_date
		Else
			If on_first_time_ppd = False Then
				If dialog_view = day_view Then
					PushButton 280, 33, 15, 12, "-", minus_button
				Else
					PushButton 180, 33, 15, 12, "-", minus_button
				End If
			End If
			If on_current_time_pd = False Then PushButton 315, 33, 15, 12, "+", plus_button
		End If
		PushButton 335, 15, 40, 15, "Switch", switch_button
		Text 85, 5, 110, 10, selected_date
		Text 15, 40, 165, 10, "Total Hours Logged: " & display_total_hours
		Text 15, 55, 165, 10, "Hours in Meetings: " & display_meeting_hours
		If dialog_view <> day_view Then PushButton 400, 10, 50, 15, "Day", day_button
		If dialog_view = day_view Then Text 418, 13, 30, 10, "Day"
		If dialog_view <> week_view Then PushButton 400, 25, 50, 15, "Week", week_button
		If dialog_view = week_view Then Text 415, 28, 35, 10, "Week"
		If dialog_view <> biweek_view Then PushButton 400, 40, 50, 15, "Pay Period", pay_period_button
		If dialog_view = biweek_view Then Text 405, 43, 43, 10, "Pay Period"
		If dialog_view <> month_view Then PushButton 400, 55, 50, 15, "Month", month_button
		If dialog_view = month_view Then Text 415, 58, 30, 10, "Month"
		If dialog_view <> custom_view Then PushButton 400, 70, 50, 15, "Custom", custom_time_button
		If dialog_view = custom_view Then Text 411, 73, 30, 10, "Custom"
		GroupBox 10, 75, 375, grp_1_len, "Hours by " & selected_sort
		y_pos = 90
		If selected_sort = "CATEGORY" Then
			For cat_item = 0 to UBound(CATEGORY_ARRAY, 2)
				Text 20, y_pos, 150, 10, CATEGORY_ARRAY(type_detail_const, cat_item) & ": "
				Text 270, y_pos, 50, 10, CATEGORY_ARRAY(total_hours_string_const, cat_item)
				y_pos = y_pos + 10
			Next
		End If
		If selected_sort = "PROJECT" Then
			For cat_item = 0 to UBound(PROJECT_ARRAY, 2)
				Text 20, y_pos, 150, 10, PROJECT_ARRAY(type_detail_const, cat_item) & ": "
				Text 170, y_pos, 50, 10, PROJECT_ARRAY(total_hours_string_const, cat_item)
				y_pos = y_pos + 10
			Next
		End If
		If selected_sort = "GITHUB ISSUE" Then
			For cat_item = 0 to UBound(GITHUB_ISSUE_ARRAY, 2)
				If GITHUB_ISSUE_ARRAY(type_detail_const, cat_item) <> "No Specified Issue" and Instr(GITHUB_ISSUE_ARRAY(type_detail_const, cat_item), "#") = 0 Then PushButton 20, y_pos, 65, 10, "Issue: " & GITHUB_ISSUE_ARRAY(type_detail_const, cat_item),  GITHUB_ISSUE_ARRAY(type_btn_const, cat_item)
				If Instr(GITHUB_ISSUE_ARRAY(type_detail_const, cat_item), "#") <> 0 Then Text 20, y_pos, 75, 10, "Issue: " & GITHUB_ISSUE_ARRAY(type_detail_const, cat_item) & ": "
				If GITHUB_ISSUE_ARRAY(type_detail_const, cat_item) = "No Specified Issue" Then Text 20, y_pos, 75, 10, GITHUB_ISSUE_ARRAY(type_detail_const, cat_item) & ": "
				Text 95, y_pos+1, 50, 10, GITHUB_ISSUE_ARRAY(total_hours_string_const, cat_item)
				y_pos = y_pos + 10
			Next
		End If
		If selected_sort = "DAY" Then
			If dialog_view = day_view Then
				For logged_activity = 0 to UBound(TIME_TRACKING_ARRAY, 2)
					If DateDiff("d", selected_date, TIME_TRACKING_ARRAY(activity_date_const, logged_activity)) = 0 Then
						If TIME_TRACKING_ARRAY(activity_paid_yn, logged_activity) = "Y" Then
							Text 15, y_pos, 350, 10, TIME_TRACKING_ARRAY(activity_category, logged_activity) & ": " & TIME_TRACKING_ARRAY(activity_detail, logged_activity)
							Text 365, y_pos, 15, 10, TIME_TRACKING_ARRAY(activity_time_spent, logged_activity)
						Else
							Text 15, y_pos, 340, 10, TIME_TRACKING_ARRAY(activity_category, logged_activity) & ": " & TIME_TRACKING_ARRAY(activity_detail, logged_activity)
							Text 355, y_pos, 25, 10, "UnPaid"
						End If
						y_pos = y_pos + 10
					End If
				Next
			Else
				For cat_item = 0 to UBound(DAY_SORT_ARRAY, 2)
					Text 20, y_pos, 90, 10, DAY_SORT_ARRAY(type_detail_const, cat_item) & ": "
					Text 110, y_pos, 50, 10, DAY_SORT_ARRAY(total_hours_string_const, cat_item)
					y_pos = y_pos + 10
				Next
			End If
		End If
		y_pos = y_pos + 5
		If selected_sort <> "CATEGORY" Then PushButton 15, y_pos, 60, 10, "CATEGORY", category_button
		If selected_sort <> "PROJECT" Then PushButton 75, y_pos, 60, 10, "PROJECT", project_button
		If selected_sort <> "GITHUB ISSUE" Then PushButton 135, y_pos, 60, 10, "GITHUB ISSUE", git_hub_issue_button
		If selected_sort <> "DAY" Then PushButton 195, y_pos, 30, 10, "DAY", day_sort_button
		y_pos = y_pos + 25
		If view_excel = False Then PushButton 5, y_pos, 100, 15, "Show Excel", show_excel_button
		If view_excel = True Then
			PushButton 5, y_pos, 100, 15, "Hide Excel", hide_excel_button
			CheckBox 110, y_pos + 5, 75, 10, "Leave Excel Open", leave_excel_open_checkbox
		End If
		Text 200, y_pos + 5, 150, 10, "Active Time Period: " &  active_period_start & " - " & active_period_end
		OkButton 405, y_pos, 50, 15
		' CancelButton 305, y_pos, 50, 15
	EndDialog

	dialog Dialog1
	If ButtonPressed = 0 Then ObjExcel.Quit			'If we press Cancel, it will close the file and stop the script
	cancel_without_confirmation

	err_msg = "MORE"								'this is always filled in unless the 'Enter' button is pressed.

	If ButtonPressed = show_excel_button Then view_excel = True					'Here we can change the showing of the Excel File
	If ButtonPressed = hide_excel_button Then view_excel = False
	If ButtonPressed = hide_excel_button OR ButtonPressed = show_excel_button Then objExcel.Visible = view_excel
	If selected_sort = "GITHUB ISSUE" Then										'If a GitHub Issue button is pressed, it will open Chrome with the Issue Number
		For cat_item = 0 to UBound(GITHUB_ISSUE_ARRAY, 2)
			If ButtonPressed = GITHUB_ISSUE_ARRAY(type_btn_const, cat_item) Then run "C:\Program Files\Google\Chrome\Application\chrome.exe https://github.com/Hennepin-County/MAXIS-scripts/issues/" & GITHUB_ISSUE_ARRAY(type_detail_const, cat_item)
		Next
	End If

	If ButtonPressed = day_button Then			'Chaning the view based on the buttons pressed
		dialog_view = day_view
		selected_date = current_day
	End If
	If ButtonPressed = week_button Then
		dialog_view = week_view
		selected_date = current_week
	End If
	If ButtonPressed = pay_period_button Then
		dialog_view = biweek_view
		selected_date = current_pay_pd
	End If
	If ButtonPressed = month_button Then
		dialog_view = month_view
		selected_date = current_month
	End If
	If ButtonPressed = custom_time_button Then
		dialog_view = custom_view
		selected_start_date = "1/1/2021"
		selected_end_date = "12/31/2021"
		selected_date = ""
	End If
	If ButtonPressed = plus_button Then
		If dialog_view = day_view Then
			selected_date = DateAdd("d", 1, selected_date)
		End If
		If dialog_view = week_view OR dialog_view = biweek_view Then
			temp_array = ""
			temp_array = split(selected_date, " - ")
			temp_array(0) = DateAdd("d", 0, temp_array(0))
			temp_array(1) = DateAdd("d", 0, temp_array(1))
			selected_start_date = temp_array(0)
			selected_end_date = temp_array(1)
			If dialog_view = week_view Then
				selected_start_date = DateAdd("d", 7, selected_start_date)
				selected_end_date = DateAdd("d", 7, selected_end_date)
			End if
			If dialog_view = biweek_view Then
				selected_start_date = DateAdd("d", 14, selected_start_date)
				selected_end_date = DateAdd("d", 14, selected_end_date)
			End If
			selected_date = selected_start_date & " - " & selected_end_date
		End If
		If dialog_view = month_view Then
			If selected_date = "January" Then selected_date = "February"
			If selected_date = "February" Then selected_date = "March"
			If selected_date = "March" Then selected_date = "April"
			If selected_date = "April" Then selected_date = "May"
			If selected_date = "May" Then selected_date = "June"
			If selected_date = "June" Then selected_date = "July"
			If selected_date = "July" Then selected_date = "August"
			If selected_date = "August" Then selected_date = "September"
			If selected_date = "September" Then selected_date = "October"
			If selected_date = "October" Then selected_date = "November"
			If selected_date = "November" Then selected_date = "December"
			If selected_date = "December" Then selected_date = "January"
		End If
	End If
	If ButtonPressed = minus_button Then
		If dialog_view = day_view Then
			selected_date = DateAdd("d", -1, selected_date)
		End If
		If dialog_view = week_view OR dialog_view = biweek_view Then
			temp_array = ""
			temp_array = split(selected_date, " - ")
			temp_array(0) = DateAdd("d", 0, temp_array(0))
			temp_array(1) = DateAdd("d", 0, temp_array(1))
			selected_start_date = temp_array(0)
			selected_end_date = temp_array(1)
			If dialog_view = week_view Then
				selected_start_date = DateAdd("d", -7, selected_start_date)
				selected_end_date = DateAdd("d", -7, selected_end_date)
			End if
			If dialog_view = biweek_view Then
				selected_start_date = DateAdd("d", -14, selected_start_date)
				selected_end_date = DateAdd("d", -14, selected_end_date)
			End If
			selected_date = selected_start_date & " - " & selected_end_date
		End If
		If dialog_view = month_view Then
			If selected_date = "January" Then selected_date = "December"
			If selected_date = "February" Then selected_date = "January"
			If selected_date = "March" Then selected_date = "February"
			If selected_date = "April" Then selected_date = "March"
			If selected_date = "May" Then selected_date = "April"
			If selected_date = "June" Then selected_date = "May"
			If selected_date = "July" Then selected_date = "June"
			If selected_date = "August" Then selected_date = "July"
			If selected_date = "September" Then selected_date = "August"
			If selected_date = "October" Then selected_date = "September"
			If selected_date = "November" Then selected_date = "October"
			If selected_date = "December" Then selected_date = "November"
		End If
	End If

	If dialog_view = day_view Then								'Based on the view selected, this will set the date(s) to be used when finding the right times/information to display from the array
		selected_date = DateAdd("d", 0, selected_date)
		selected_start_date = DateAdd("d", 0, selected_date)
		selected_end_date = DateAdd("d", 0, selected_date)
	End If
	If dialog_view = week_view or dialog_view = biweek_view Then
		temp_array = ""
		temp_array = split(selected_date, " - ")
		temp_array(0) = DateAdd("d", 0, temp_array(0))
		temp_array(1) = DateAdd("d", 0, temp_array(1))
		selected_start_date = temp_array(0)
		selected_end_date = temp_array(1)
	End If
	If dialog_view = month_view Then
		If selected_date = "January" Then
			selected_start_date = #1/1/2022#
			selected_end_date = #1/31/2022#
		End if
		If selected_date = "February" Then
			selected_start_date = #2/1/2022#
			selected_end_date = #2/28/2022#
		End if
		If selected_date = "March" Then
			selected_start_date = #3/1/2022#
			selected_end_date = #3/31/2022#
		End if
		If selected_date = "April" Then
			selected_start_date = #4/1/2022#
			selected_end_date = #4/30/2022#
		End if
		If selected_date = "May" Then
			selected_start_date = #5/1/2022#
			selected_end_date = #5/31/2022#
		End if
		If selected_date = "June" Then
			selected_start_date = #6/1/2022#
			selected_end_date = #6/30/2022#
		End if
		If selected_date = "July" Then
			selected_start_date = #7/1/2022#
			selected_end_date = #7/31/2022#
		End if
		If selected_date = "August" Then
			selected_start_date = #8/1/2022#
			selected_end_date = #8/31/2022#
		End if
		If selected_date = "September" Then
			selected_start_date = #9/1/2022#
			selected_end_date = #9/30/2022#
		End if
		If selected_date = "October" Then
			selected_start_date = #10/1/2022#
			selected_end_date = #10/31/2022#
		End if
		If selected_date = "November" Then
			selected_start_date = #11/1/2022#
			selected_end_date = #11/30/2022#
		End if
		If selected_date = "December" Then
			selected_start_date = #12/1/2022#
			selected_end_date = #12/31/2022#
		End if
	End If
	If dialog_view = custom_view Then
		selected_start_date = DateAdd("d", 0, selected_start_date)
		selected_end_date = DateAdd("d", 0, selected_end_date)
	End If

	If ButtonPressed = category_button Then selected_sort = "CATEGORY"			'Setting the right sort
	If ButtonPressed = project_button Then selected_sort = "PROJECT"
	If ButtonPressed = git_hub_issue_button Then selected_sort = "GITHUB ISSUE"
	If ButtonPressed = day_sort_button Then selected_sort = "DAY"

	'This is going to call the function that will fill the array information from the options selected so that the display in the dialog is correct.
	If selected_sort = "CATEGORY" Then Call create_time_spent_totals(selected_start_date, selected_end_date, selected_sort, hours_in_time_pd, hours_in_meetings_dur_time_pd, CATEGORY_ARRAY)
	If selected_sort = "PROJECT" Then Call create_time_spent_totals(selected_start_date, selected_end_date, selected_sort, hours_in_time_pd, hours_in_meetings_dur_time_pd, PROJECT_ARRAY)
	If selected_sort = "GITHUB ISSUE" Then Call create_time_spent_totals(selected_start_date, selected_end_date, selected_sort, hours_in_time_pd, hours_in_meetings_dur_time_pd, GITHUB_ISSUE_ARRAY)
	If selected_sort = "DAY" Then Call create_time_spent_totals(selected_start_date, selected_end_date, selected_sort, hours_in_time_pd, hours_in_meetings_dur_time_pd, DAY_SORT_ARRAY)
	' Call create_time_spent_totals(start_date, end_date, sort_type, total_hours, hours_in_meetings, TYPE_ARRAY)

	If ButtonPressed = -1 Then err_msg = ""		'blanking out the err_msg if the 'OK' button is pressed so we can leave the dialog loop
Loop until err_msg = ""

added_end_time_row_list = trim(added_end_time_row_list)
If added_end_time_row_list <> "" then
	array_of_rows_to_remove_end_time = split(added_end_time_row_list)
	For each excel_row in array_of_rows_to_remove_end_time
		MsgBox "~" & excel_row & "~"
		ObjExcel.Cells(excel_row, 3).Value = ""
	Next
	objWorkbook.Save									'saving the file to 'My Documents'
	objWorkbook.SaveAs (t_drive_excel_file_path)		'saving the file to the T Drive
End If

If view_excel = False Then leave_excel_open_checkbox = unchecked
If leave_excel_open_checkbox = checked Then				'If the checkbox is checked then we block out any row that was changed for math to work. This isn't needed if we aren't leaving it open then it closes without being saved.
	row_filled_with_end_time = trim(row_filled_with_end_time)
	If Instr(row_filled_with_end_time, " ") = 0 Then
		row_filled_with_end_time = Array(row_filled_with_end_time)
	Else
		row_filled_with_end_time = split(row_filled_with_end_time, " ")
	End If
	For each changed_row in row_filled_with_end_time
		If changed_row <> "" Then
			ObjExcel.Cells(changed_row, 3).Value = ""
			ObjExcel.Cells(changed_row, 4).Value = ""
		End If
	Next
End If
If leave_excel_open_checkbox = unchecked Then ObjExcel.Quit		'Closing the Excel file.
Call script_end_procedure("")
