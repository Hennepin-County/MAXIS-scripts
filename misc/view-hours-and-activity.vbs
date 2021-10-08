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
		If sort_type = "DAY" Then TYPE_ARRAY(type_detail_const, each_type) = TYPE_ARRAY(type_detail_const, each_type) & " - " & WeekdayName(WeekDay(TYPE_ARRAY(type_detail_const, each_type)))
	Next
end function
'===========================================================================================================================

'DECLARATIONS===============================================================================================================
'Lists for Dialogs
week_list = "Select"
' week_list = week_list+chr(9)+"12/27/2020 - 1/2/2021"
' week_list = week_list+chr(9)+"1/3/2021 - 1/9/2021"
' week_list = week_list+chr(9)+"1/10/2021 - 1/16/2021"
' week_list = week_list+chr(9)+"1/17/2021 - 1/23/2021"
' week_list = week_list+chr(9)+"1/24/2021 - 1/30/2021"
' week_list = week_list+chr(9)+"1/31/2021 - 2/6/2021"
' week_list = week_list+chr(9)+"2/7/2021 - 2/13/2021"
' week_list = week_list+chr(9)+"2/14/2021 - 2/20/2021"
' week_list = week_list+chr(9)+"2/21/2021 - 2/27/2021"
' week_list = week_list+chr(9)+"2/28/2021 - 3/6/2021"
' week_list = week_list+chr(9)+"3/7/2021 - 3/13/2021"
' week_list = week_list+chr(9)+"3/14/2021 - 3/20/2021"
' week_list = week_list+chr(9)+"3/21/2021 - 3/27/2021"
' week_list = week_list+chr(9)+"3/28/2021 - 4/3/2021"
' week_list = week_list+chr(9)+"4/4/2021 - 4/10/2021"
' week_list = week_list+chr(9)+"4/11/2021 - 4/17/2021"
' week_list = week_list+chr(9)+"4/18/2021 - 4/24/2021"
' week_list = week_list+chr(9)+"4/25/2021 - 5/1/2021"
' week_list = week_list+chr(9)+"5/2/2021 - 5/8/2021"
' week_list = week_list+chr(9)+"5/9/2021 - 5/15/2021"
' week_list = week_list+chr(9)+"5/16/2021 - 5/22/2021"
' week_list = week_list+chr(9)+"5/23/2021 - 5/29/2021"
' week_list = week_list+chr(9)+"5/30/2021 - 6/5/2021"
' week_list = week_list+chr(9)+"6/6/2021 - 6/12/2021"
' week_list = week_list+chr(9)+"6/13/2021 - 6/19/2021"
' week_list = week_list+chr(9)+"6/20/2021 - 6/26/2021"
week_list = week_list+chr(9)+"6/27/2021 - 7/3/2021"
week_list = week_list+chr(9)+"7/4/2021 - 7/10/2021"
week_list = week_list+chr(9)+"7/11/2021 - 7/17/2021"
week_list = week_list+chr(9)+"7/18/2021 - 7/24/2021"
week_list = week_list+chr(9)+"7/25/2021 - 7/31/2021"
week_list = week_list+chr(9)+"8/1/2021 - 8/7/2021"
week_list = week_list+chr(9)+"8/8/2021 - 8/14/2021"
week_list = week_list+chr(9)+"8/15/2021 - 8/21/2021"
week_list = week_list+chr(9)+"8/22/2021 - 8/28/2021"
week_list = week_list+chr(9)+"8/29/2021 - 9/4/2021"
week_list = week_list+chr(9)+"9/5/2021 - 9/11/2021"
week_list = week_list+chr(9)+"9/12/2021 - 9/18/2021"
week_list = week_list+chr(9)+"9/19/2021 - 9/25/2021"
week_list = week_list+chr(9)+"9/26/2021 - 10/2/2021"
week_list = week_list+chr(9)+"10/3/2021 - 10/9/2021"
week_list = week_list+chr(9)+"10/10/2021 - 10/16/2021"
week_list = week_list+chr(9)+"10/17/2021 - 10/23/2021"
week_list = week_list+chr(9)+"10/24/2021 - 10/30/2021"
week_list = week_list+chr(9)+"10/31/2021 - 11/6/2021"
week_list = week_list+chr(9)+"11/7/2021 - 11/13/2021"
week_list = week_list+chr(9)+"11/14/2021 - 11/20/2021"
week_list = week_list+chr(9)+"11/21/2021 - 11/27/2021"
week_list = week_list+chr(9)+"11/28/2021 - 12/4/2021"
week_list = week_list+chr(9)+"12/5/2021 - 12/11/2021"
week_list = week_list+chr(9)+"12/12/2021 - 12/18/2021"
week_list = week_list+chr(9)+"12/19/2021 - 12/25/2021"
week_list = week_list+chr(9)+"12/26/2021 - 1/1/2022"
week_list = week_list+chr(9)+"1/2/2022 - 1/9/2022"
week_array = split(week_list, chr(9))

biweek_list = "Select"
biweek_list = biweek_list+chr(9)+"12/20/2020 - 1/2/2021"
biweek_list = biweek_list+chr(9)+"1/3/2021 - 1/16/2021"
biweek_list = biweek_list+chr(9)+"1/17/2021 - 1/30/2021"
biweek_list = biweek_list+chr(9)+"1/31/2021 - 2/13/2021"
biweek_list = biweek_list+chr(9)+"2/14/2021 - 2/27/2021"
biweek_list = biweek_list+chr(9)+"2/28/2021 - 3/13/2021"
biweek_list = biweek_list+chr(9)+"3/14/2021 - 3/27/2021"
biweek_list = biweek_list+chr(9)+"3/28/2021 - 4/10/2021"
biweek_list = biweek_list+chr(9)+"4/11/2021 - 4/24/2021"
biweek_list = biweek_list+chr(9)+"4/25/2021 - 5/8/2021"
biweek_list = biweek_list+chr(9)+"5/9/2021 - 5/22/2021"
biweek_list = biweek_list+chr(9)+"5/23/2021 - 6/5/2021"
biweek_list = biweek_list+chr(9)+"6/6/2021 - 6/19/2021"
biweek_list = biweek_list+chr(9)+"6/20/2021 - 7/3/2021"
biweek_list = biweek_list+chr(9)+"7/4/2021 - 7/17/2021"
biweek_list = biweek_list+chr(9)+"7/18/2021 - 7/31/2021"
biweek_list = biweek_list+chr(9)+"8/1/2021 - 8/14/2021"
biweek_list = biweek_list+chr(9)+"8/15/2021 - 8/28/2021"
biweek_list = biweek_list+chr(9)+"8/29/2021 - 9/11/2021"
biweek_list = biweek_list+chr(9)+"9/12/2021 - 9/25/2021"
biweek_list = biweek_list+chr(9)+"9/26/2021 - 10/9/2021"
biweek_list = biweek_list+chr(9)+"10/10/2021 - 10/23/2021"
biweek_list = biweek_list+chr(9)+"10/24/2021 - 11/6/2021"
biweek_list = biweek_list+chr(9)+"11/7/2021 - 11/20/2021"
biweek_list = biweek_list+chr(9)+"11/21/2021 - 12/4/2021"
biweek_list = biweek_list+chr(9)+"12/5/2021 - 12/18/2021"
biweek_list = biweek_list+chr(9)+"12/19/2021 - 1/1/2022"
biweek_list = biweek_list+chr(9)+"1/2/2022 - 1/9/2022"
biweek_array = split(biweek_list, chr(9))

month_list = "Select"
month_list = month_list+chr(9)+"January"
month_list = month_list+chr(9)+"February"
month_list = month_list+chr(9)+"March"
month_list = month_list+chr(9)+"April"
month_list = month_list+chr(9)+"May"
month_list = month_list+chr(9)+"June"
month_list = month_list+chr(9)+"July"
month_list = month_list+chr(9)+"August"
month_list = month_list+chr(9)+"September"
month_list = month_list+chr(9)+"October"
month_list = month_list+chr(9)+"November"
month_list = month_list+chr(9)+"December"

'Constants and arrays
const activity_date_const 		= 00
const activity_start_time		= 01
const activity_end_time			= 02
const activity_time_spent		= 03
const activity_time_spent_val	= 11
const activity_category			= 04
const activity_meeting			= 05
const activity_detail			= 06
const activity_gh_issue_numb	= 07
const activity_gh_issue_url		= 08
const activity_project			= 09
const activity_paid_yn			= 10
const last_const 				= 12

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
	t_drive_excel_file_path = excel_file_path & "\Casey Time Tracking 2021.xlsx"
	my_docs_excel_file_path = user_myDocs_folder & "Casey Time Tracking 2021.xlsx"
End If
If user_ID_for_validation = "ILFE001" Then
	t_drive_excel_file_path = excel_file_path & "\Ilse Time Tracking 2021.xlsx"
	my_docs_excel_file_path = user_myDocs_folder & "Ilse Time Tracking 2021.xlsx"
End If
If user_ID_for_validation = "WFS395" Then
	t_drive_excel_file_path = excel_file_path & "\MiKayla Time Tracking 2021.xlsx"
	my_docs_excel_file_path = user_myDocs_folder & "MiKayla Time Tracking 2021.xlsx"
End If

view_excel = False		'this variable allows us to set
Call excel_open(t_drive_excel_file_path, view_excel, False, ObjExcel, objWorkbook)		'opening the excel file

row_filled_with_end_time = " "

'Here we read the entire excel file and save it into an array
excel_row = 2			'start of the excel file information
activity_count = 0		'starting of the counter of the array
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

'now that the reading is done, we are going to make the Excel file visible.
'This is particularly for TESTING and we can remove this in the future
view_excel = True
objExcel.Visible = view_excel

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
selected_sort = "CATEGORY"
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
	BeginDialog Dialog1, 0, 0, 361, dlg_len, "View Hours and Activity"
	  ButtonGroup ButtonPressed
		GroupBox 5, 5, 285, grp_2_len, "Hours Breakdown"
		If dialog_view = day_view Then EditBox 180, 15, 50, 15, selected_date
		If dialog_view = week_view Then DropListBox 80, 15, 150, 45, week_list, selected_date
		If dialog_view = biweek_view Then DropListBox 80, 15, 150, 45, biweek_list, selected_date
		If dialog_view = month_view Then DropListBox 80, 15, 150, 45, month_list, selected_date
		If dialog_view = custom_view Then
			EditBox 130, 15, 40, 15, selected_start_date
			Text 170, 17, 10, 10, " - "
			EditBox 180, 15, 50, 15, selected_end_date
		End If
		PushButton 235, 15, 40, 15, "Switch", switch_button
		Text 85, 5, 110, 10, selected_date
		Text 15, 40, 250, 10, "Total Hours Logged: " & display_total_hours
		Text 15, 55, 250, 10, "Hours in Meetings: " & display_meeting_hours
		If dialog_view <> day_view Then PushButton 300, 10, 50, 15, "Day", day_button
		If dialog_view = day_view Then Text 318, 13, 30, 10, "Day"
		If dialog_view <> week_view Then PushButton 300, 25, 50, 15, "Week", week_button
		If dialog_view = week_view Then Text 315, 28, 35, 10, "Week"
		If dialog_view <> biweek_view Then PushButton 300, 40, 50, 15, "Pay Period", pay_period_button
		If dialog_view = biweek_view Then Text 305, 43, 43, 10, "Pay Period"
		If dialog_view <> month_view Then PushButton 300, 55, 50, 15, "Month", month_button
		If dialog_view = month_view Then Text 315, 58, 30, 10, "Month"
		If dialog_view <> custom_view Then PushButton 300, 70, 50, 15, "Custom", custom_time_button
		If dialog_view = custom_view Then Text 311, 73, 30, 10, "Custom"
		GroupBox 10, 75, 275, grp_1_len, "Hours by " & selected_sort
		y_pos = 90
		If selected_sort = "CATEGORY" Then
			For cat_item = 0 to UBound(CATEGORY_ARRAY, 2)
				Text 20, y_pos, 150, 10, CATEGORY_ARRAY(type_detail_const, cat_item) & ": "
				Text 170, y_pos, 50, 10, CATEGORY_ARRAY(total_hours_string_const, cat_item)
				y_pos = y_pos + 10
			Next
		End If
		If selected_sort = "PROJECT" Then
			For cat_item = 0 to UBound(PROJECT_ARRAY, 2)
				Text 20, y_pos, 90, 10, PROJECT_ARRAY(type_detail_const, cat_item) & ": "
				Text 110, y_pos, 50, 10, PROJECT_ARRAY(total_hours_string_const, cat_item)
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
							Text 15, y_pos, 250, 10, TIME_TRACKING_ARRAY(activity_category, logged_activity) & ": " & TIME_TRACKING_ARRAY(activity_detail, logged_activity)
							Text 265, y_pos, 15, 10, TIME_TRACKING_ARRAY(activity_time_spent, logged_activity)
						Else
							Text 15, y_pos, 240, 10, TIME_TRACKING_ARRAY(activity_category, logged_activity) & ": " & TIME_TRACKING_ARRAY(activity_detail, logged_activity)
							Text 255, y_pos, 25, 10, "UnPaid"
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
		If view_excel = True Then PushButton 5, y_pos, 100, 15, "Hide Excel", hide_excel_button
		CheckBox 110, y_pos + 5, 100, 10, "Leave Excel Open", leave_excel_open_checkbox
		OkButton 305, y_pos, 50, 15
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
			selected_start_date = #1/1/2021#
			selected_end_date = #1/31/2021#
		End if
		If selected_date = "February" Then
			selected_start_date = #2/1/2021#
			selected_end_date = #2/28/2021#
		End if
		If selected_date = "March" Then
			selected_start_date = #3/1/2021#
			selected_end_date = #3/31/2021#
		End if
		If selected_date = "April" Then
			selected_start_date = #4/1/2021#
			selected_end_date = #4/30/2021#
		End if
		If selected_date = "May" Then
			selected_start_date = #5/1/2021#
			selected_end_date = #5/31/2021#
		End if
		If selected_date = "June" Then
			selected_start_date = #6/1/2021#
			selected_end_date = #6/30/2021#
		End if
		If selected_date = "July" Then
			selected_start_date = #7/1/2021#
			selected_end_date = #7/31/2021#
		End if
		If selected_date = "August" Then
			selected_start_date = #8/1/2021#
			selected_end_date = #8/31/2021#
		End if
		If selected_date = "September" Then
			selected_start_date = #9/1/2021#
			selected_end_date = #9/30/2021#
		End if
		If selected_date = "October" Then
			selected_start_date = #10/1/2021#
			selected_end_date = #10/31/2021#
		End if
		If selected_date = "November" Then
			selected_start_date = #11/1/2021#
			selected_end_date = #11/30/2021#
		End if
		If selected_date = "December" Then
			selected_start_date = #12/1/2021#
			selected_end_date = #12/31/2021#
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
If leave_excel_open_checkbox = checked Then				'If the checkbox is checked then we block out any row that was changed for math to work. This isn't needed if we aren't leaving it open then it closes without being saved.
	row_filled_with_end_time = trim(row_filled_with_end_time)
	If Instr(row_filled_with_end_time, " ") = 0 Then
		row_filled_with_end_time = Array(row_filled_with_end_time)
	Else
		row_filled_with_end_time = split(row_filled_with_end_time, " ")
	End If
	For each changed_row in row_filled_with_end_time
		ObjExcel.Cells(changed_row, 3).Value = ""
		ObjExcel.Cells(changed_row, 4).Value = ""
	Next
End If
If leave_excel_open_checkbox = unchecked Then ObjExcel.Quit		'Closing the Excel file.
Call script_end_procedure("")
