'Required for statistical purposes==========================================================================================
name_of_script = "ADMIN - SCRIPT USAGE REPORT.vbs"
start_time = timer
STATS_counter = 0                     	'sets the stats counter at one
STATS_manualtime = 240                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================
run_locally = TRUE
'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
    IF on_the_desert_island = TRUE Then
        FuncLib_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\desert-island\MASTER FUNCTIONS LIBRARY.vbs"
        Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
        Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
        text_from_the_other_script = fso_command.ReadAll
        fso_command.Close
        Execute text_from_the_other_script
    ELSEIF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
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
list_of_supervisors = list_of_supervisors+chr(9)+"Abdillahi, Hodan"
list_of_supervisors = list_of_supervisors+chr(9)+"Abdirahman, Mohamed"
list_of_supervisors = list_of_supervisors+chr(9)+"Adeniyi-Akins, Olukemi O"
list_of_supervisors = list_of_supervisors+chr(9)+"Alvarez, Claudia"
list_of_supervisors = list_of_supervisors+chr(9)+"Amadi, Tania L"
list_of_supervisors = list_of_supervisors+chr(9)+"Anderson, Marya D"
list_of_supervisors = list_of_supervisors+chr(9)+"Barnes, TyAnn"
list_of_supervisors = list_of_supervisors+chr(9)+"Berry, Jailon"
list_of_supervisors = list_of_supervisors+chr(9)+"Bradbury, Phillip M"
list_of_supervisors = list_of_supervisors+chr(9)+"Brown, Candace"
list_of_supervisors = list_of_supervisors+chr(9)+"Brown, Megan"
list_of_supervisors = list_of_supervisors+chr(9)+"Clifton, Doryan C"
list_of_supervisors = list_of_supervisors+chr(9)+"Coenen, Tammy L"
list_of_supervisors = list_of_supervisors+chr(9)+"Collins, Elizabeth Allison"
list_of_supervisors = list_of_supervisors+chr(9)+"Fleeman Schneider, Brianna Jo"
list_of_supervisors = list_of_supervisors+chr(9)+"Garrett, Twanda T"
list_of_supervisors = list_of_supervisors+chr(9)+"Grilley, Amy P"
list_of_supervisors = list_of_supervisors+chr(9)+"Hogan, Christopher M"
list_of_supervisors = list_of_supervisors+chr(9)+"Hughes, Nada B"
list_of_supervisors = list_of_supervisors+chr(9)+"Hurst-Baker, Valerie A"
list_of_supervisors = list_of_supervisors+chr(9)+"Lane, Matthew M"
list_of_supervisors = list_of_supervisors+chr(9)+"Lee, Payeng"
list_of_supervisors = list_of_supervisors+chr(9)+"Lucca, Jeremy T"
list_of_supervisors = list_of_supervisors+chr(9)+"Madison, Carlotta L"
list_of_supervisors = list_of_supervisors+chr(9)+"Manuel, Rashida R"
list_of_supervisors = list_of_supervisors+chr(9)+"Mcguinness, Mary V"
list_of_supervisors = list_of_supervisors+chr(9)+"Mohamed, Abdimalik A"
list_of_supervisors = list_of_supervisors+chr(9)+"Mui, Heather L"
list_of_supervisors = list_of_supervisors+chr(9)+"Nelson, Shawntel"
list_of_supervisors = list_of_supervisors+chr(9)+"Nur, Mohamed O"
list_of_supervisors = list_of_supervisors+chr(9)+"Otterness, Shauna S"
list_of_supervisors = list_of_supervisors+chr(9)+"Payne, Tanya L"
list_of_supervisors = list_of_supervisors+chr(9)+"Przybilla, Kristina M"
list_of_supervisors = list_of_supervisors+chr(9)+"Rubenstein, Daniel H F"
list_of_supervisors = list_of_supervisors+chr(9)+"Sanyal, Soumya G"
list_of_supervisors = list_of_supervisors+chr(9)+"Sebranek, Angela C"
list_of_supervisors = list_of_supervisors+chr(9)+"Socha, Monica M"
list_of_supervisors = list_of_supervisors+chr(9)+"Stone, Amber L"
list_of_supervisors = list_of_supervisors+chr(9)+"Teskey, Benjamin"
list_of_supervisors = list_of_supervisors+chr(9)+"Thyen, Benjamin R"
list_of_supervisors = list_of_supervisors+chr(9)+"Toolsie, Janelle L"
list_of_supervisors = list_of_supervisors+chr(9)+"Twomey, Susan R"
list_of_supervisors = list_of_supervisors+chr(9)+"Vogel, Susannah K"
list_of_supervisors = list_of_supervisors+chr(9)+"Vang, De"
list_of_supervisors = list_of_supervisors+chr(9)+"Williams, Natasha M"
list_of_supervisors = list_of_supervisors+chr(9)+"Wills, Angelia E"
list_of_supervisors = list_of_supervisors+chr(9)+"Yang, Alexander C"


grouping_drop_list = "Summary"
grouping_drop_list = grouping_drop_list+chr(9)+"Count by Script"
grouping_drop_list = grouping_drop_list+chr(9)+"Count by Case"
grouping_drop_list = grouping_drop_list+chr(9)+"Count by Category"
grouping_drop_list = grouping_drop_list+chr(9)+"List Only"

primary_checkbox = checked
star_checkbox = checked
admin_checkbox = checked
stats_checkbox = checked
limited_checkbox = checked
flag_checkbox = checked
uncategorized_checkbox = checked
team_checkbox = checked
team_admin_checkbox = checked
bzt_checkbox = checked

start_date = CM_minus_1_mo & "/1/" & CM_minus_1_yr
start_date = DateAdd("d", 0, start_date)
end_date = CM_mo & "/1/" & CM_yr
end_date = DateAdd("d", -1, end_date)

start_date = start_date & ""
end_date = end_date & ""

back_btn = 100
' For i = 0 to ubound(script_array)
' 	If script_array(i).usage_eval <> "" Then MsgBox "NAME - " & UCASE(script_array(i).category & " - " & script_array(i).script_name) & vbCr & "EVAL: " &script_array(i).usage_eval
' Next

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 276, 240, "Script Usage Reorts"
  GroupBox 10, 10, 255, 200, "Search Criteria"
  DropListBox 165, 35, 95, 45, "All Supervisors"+list_of_supervisors, supervisor_selection
  EditBox 165, 55, 70, 15, worker_windows_id
  EditBox 165, 90, 35, 15, start_date
  EditBox 210, 90, 35, 15, end_date
  'DropListBox 165, 120, 90, 45, grouping_drop_list, count_grouping					NOT SURE HOW WE ARE GOING TO DO THIS
  CheckBox 25, 160, 50, 10, "PRIMARY", primary_checkbox
  CheckBox 90, 160, 50, 10, "STAR", star_checkbox
  CheckBox 155, 160, 50, 10, "ADMIN", admin_checkbox
  CheckBox 220, 160, 40, 10, "STATS", stats_checkbox
  CheckBox 25, 175, 50, 10, "LIMITED", limited_checkbox
  CheckBox 90, 175, 50, 10, "FLAG", flag_checkbox
  CheckBox 155, 175, 75, 10, "UNCATEGORIZED", uncategorized_checkbox
  CheckBox 25, 190, 50, 10, "TEAM", team_checkbox
  CheckBox 90, 190, 55, 10, "TEAM-ADMIN", team_admin_checkbox
  CheckBox 155, 190, 50, 10, "BZT", bzt_checkbox
  ButtonGroup ButtonPressed
    OkButton 160, 215, 50, 15
    CancelButton 215, 215, 50, 15
    'PushButton 165, 5, 95, 15, "Run All Monthly Reports", monthly_reports_btn		NOT READY YET
  Text 65, 20, 145, 10, "Select the filters for the script udage serach."
  Text 20, 40, 145, 10, "Run for employees of a specific supervisor:"
  Text 80, 60, 85, 10, "Run for a specific worker:"
  Text 165, 70, 75, 10, "(Windows Logon ID)"
  Text 85, 95, 80, 10, "Select for these dates:"
  Text 200, 95, 10, 10, " to"
  Text 165, 105, 80, 10, "(Dates are Inclusive)"
  'Text 130, 125, 35, 10, "Count by:"
  Text 20, 140, 250, 10, "Select Script Usage Categories (check all that apply): - NOT FUNCTIONAL"
EndDialog

Do
	err_msg = ""

	dialog Dialog1
	cancel_without_confirmation

	worker_windows_id = trim(worker_windows_id)

	'Make sure a checkbox is checked
	'only allow either a supervisor or a windows user id - cannot do both
Loop until err_msg = ""


const user_name_const 		= 0
const user_number_const		= 1
const script_date_const		= 2
const script_name_const		= 3
const case_numb_const		= 4
const usage_category_const	= 5
const count_const			= 6
const script_array_index_const = 7
const script_category_const	= 8
' const _const
const usage_last_const		= 15

Dim SCRIPT_USAGE_ARRAY()
ReDim SCRIPT_USAGE_ARRAY(usage_last_const, 0)
use_count = 0

Dim SCRIPT_USED_ARRAY()
ReDim SCRIPT_USED_ARRAY(usage_last_const, 0)
unique_script_count = 0
' script_used_string = "~"

'NEED TO FIGURE OUT HOW TO TRACK BY CASE - MAYBE ALSO BY WORKER??? - MAYBE BY DATE???
Dim CASE_SCRIPT_USAGE()
ReDim CASE_SCRIPT_USAGE(usage_last_const, 0)
unique_case_count = 0
' case_numb_count = "~"

const wrkr_id_const			= 0
const wrkr_name_const		= 1
const NOTES_cnt_const		= 2
const ACTIONS_cnt_const		= 3
const NOTICES_cnt_const		= 4
const UTILITIES_cnt_const	= 5
const ADMIN_cnt_const		= 6
const BULK_cnt_const		= 7
const eval_PRIMARY_cnt_const= 8
const eval_STAR_cnt_const	= 9
const eval_LIMITED_cnt_const= 10
const eval_TEAM_cnt_const	= 11
const eval_ADMIN_cnt_const	= 12
const eval_FLAG_cnt_const	= 13
const eval_STATS_cnt_const	= 14
const eval_BLANK_cnt_const	= 15

const wrkr_ave_by_date		= 16
const wrkr_ave_by_case		= 17

const wrkr_suprvsr_id		= 20
const wrkr_suprvsr_name		= 21
const total_script_count	= 22
const wrkr_btn_const		= 23
const wrkr_last_const		= 30

Dim WORKER_ARRAY()
ReDim WORKER_ARRAY(wrkr_last_const, 0)
worker_count = 0
btn_placeholder = 500
If worker_windows_id <> "" Then
	WORKER_ARRAY(wrkr_id_const, 0) = worker_windows_id
	WORKER_ARRAY(wrkr_btn_const, 0) = btn_placeholder
	WORKER_ARRAY(total_script_count, 0) = 0
	WORKER_ARRAY(NOTES_cnt_const, 0) = 0
	WORKER_ARRAY(ACTIONS_cnt_const, 0) = 0
	WORKER_ARRAY(NOTICES_cnt_const, 0) = 0
	WORKER_ARRAY(UTILITIES_cnt_const, 0) = 0
	WORKER_ARRAY(ADMIN_cnt_const, 0) = 0
	WORKER_ARRAY(BULK_cnt_const, 0) = 0
	WORKER_ARRAY(eval_PRIMARY_cnt_const, 0) = 0
	WORKER_ARRAY(eval_STAR_cnt_const, 0) = 0
	WORKER_ARRAY(eval_LIMITED_cnt_const, 0) = 0
	WORKER_ARRAY(eval_TEAM_cnt_const, 0) = 0
	WORKER_ARRAY(eval_ADMIN_cnt_const, 0) = 0
	WORKER_ARRAY(eval_FLAG_cnt_const, 0) = 0
	WORKER_ARRAY(eval_STATS_cnt_const, 0) = 0
	WORKER_ARRAY(eval_BLANK_cnt_const, 0) = 0
End If

'HERE CREATE AN ARRAY OF ALL WORKERS UNDER A SUPERVISOR with counts for each worker of category types
If supervisor_selection <> "All Supervisors" Then

	Set objConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
	Set objRecordSet = CreateObject("ADODB.Recordset")

	objSQL = ""
	objSQL = "SELECT "
	objSQL = objSQL & "      [EmpFullName]"
	objSQL = objSQL & "      ,[L1Manager]"
	objSQL = objSQL & "      ,[EmployeeEmail]"
	objSQL = objSQL & "      ,[EmpLogOnID]"
	objSQL = objSQL & "  FROM [BlueZone_Statistics].[ES].[ES_StaffHierarchyDim]"
	objSQL = objSQL & "  WHERE L1Manager like '%" & supervisor_selection & "%'"

	'opening the connections and data table
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	Do While NOT objRecordSet.Eof
		ReDim preserve WORKER_ARRAY(wrkr_last_const, worker_count)

		WORKER_ARRAY(wrkr_id_const, worker_count) = objRecordSet("EmpLogOnID")
		WORKER_ARRAY(wrkr_name_const, worker_count) = objRecordSet("EmpFullName")
		WORKER_ARRAY(wrkr_suprvsr_name, worker_count) = objRecordSet("L1Manager")
		WORKER_ARRAY(wrkr_btn_const, worker_count) = btn_placeholder + worker_count
		WORKER_ARRAY(total_script_count, worker_count) = 0
		WORKER_ARRAY(NOTES_cnt_const, worker_count) = 0
		WORKER_ARRAY(ACTIONS_cnt_const, worker_count) = 0
		WORKER_ARRAY(NOTICES_cnt_const, worker_count) = 0
		WORKER_ARRAY(UTILITIES_cnt_const, worker_count) = 0
		WORKER_ARRAY(ADMIN_cnt_const, worker_count) = 0
		WORKER_ARRAY(BULK_cnt_const, worker_count) = 0
		WORKER_ARRAY(eval_PRIMARY_cnt_const, worker_count) = 0
		WORKER_ARRAY(eval_STAR_cnt_const, worker_count) = 0
		WORKER_ARRAY(eval_LIMITED_cnt_const, worker_count) = 0
		WORKER_ARRAY(eval_TEAM_cnt_const, worker_count) = 0
		WORKER_ARRAY(eval_ADMIN_cnt_const, worker_count) = 0
		WORKER_ARRAY(eval_FLAG_cnt_const, worker_count) = 0
		WORKER_ARRAY(eval_STATS_cnt_const, worker_count) = 0
		WORKER_ARRAY(eval_BLANK_cnt_const, worker_count) = 0

		worker_count = worker_count + 1
		objRecordSet.MoveNext
	Loop

	objRecordSet.Close			'Closing all the data connections
	objConnection.Close

	Set objRecordSet=nothing
	Set objConnection=nothing
End If

Set objConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
Set objRecordSet = CreateObject("ADODB.Recordset")

objSQL = ""
objSQL = objSQL & "SELECT ED.EmpFullName as ScriptUser"
objSQL = objSQL & ", BZS.USERNAME as WFNumber"
objSQL = objSQL & ", BZS.SDATE as ScriptRunDate"
objSQL = objSQL & ", BZS.SCRIPT_NAME as ScriptName"
objSQL = objSQL & ", BZS.CASE_NUMBER as CaseNum"
objSQL = objSQL & ", BZS.CLOSING_MSGBOX as EndMessage"
objSQL = objSQL & " FROM dbo.usage_log as BZS"
objSQL = objSQL & " LEFT JOIN ES.ES_StaffHierarchyDim as ED ON BZS.USERNAME = ED.EmpLogOnID"

objSQL = objSQL & " WHERE SDATE >= '" & start_date & "' and SDATE <= '" & end_date & "'"
objSQL = objSQL & " and CLOSING_MSGBOX not like '~PT%'"
If supervisor_selection <> "All Supervisors" Then
	objSQL = objSQL & "and L1Manager = '" & supervisor_selection & "'"
End If
If worker_windows_id <> "" Then objSQL = objSQL & "and USERNAME like '%" & worker_windows_id & "%'"

'opening the connections and data table
objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
objRecordSet.Open objSQL, objConnection

Do While NOT objRecordSet.Eof
	ReDim preserve SCRIPT_USAGE_ARRAY(usage_last_const, use_count)
	SCRIPT_USAGE_ARRAY(user_name_const, use_count) = objRecordSet("ScriptUser")
	SCRIPT_USAGE_ARRAY(user_number_const, use_count) = objRecordSet("WFNumber")
	SCRIPT_USAGE_ARRAY(script_date_const, use_count) = objRecordSet("ScriptRunDate")
	SCRIPT_USAGE_ARRAY(script_name_const, use_count) = objRecordSet("ScriptName")
	SCRIPT_USAGE_ARRAY(case_numb_const, use_count) = objRecordSet("CaseNum")
	use_count = use_count + 1
	objRecordSet.MoveNext
Loop

objRecordSet.Close			'Closing all the data connections
objConnection.Close

Set objRecordSet=nothing
Set objConnection=nothing

Set objConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
Set objRecordSet = CreateObject("ADODB.Recordset")


objSQL = ""
objSQL = "SELECT "
objSQL = objSQL & " BZS.USERNAME as WFNumber,"
objSQL = objSQL & "	BZS.SCRIPT_NAME as ScriptName,"
objSQL = objSQL & "	COUNT(BZS.SCRIPT_NAME) as ScriptRuns"
objSQL = objSQL & " FROM 	"
objSQL = objSQL & "	BlueZone_Statistics.ES.ES_StaffHierarchyDim as ED"
objSQL = objSQL & "	LEFT JOIN BlueZone_Statistics.dbo.usage_log as BZS"
objSQL = objSQL & "	ON  ED.EmpLogOnID	= BZS.USERNAME"
objSQL = objSQL & "	"
objSQL = objSQL & " WHERE SDATE >= '" & start_date & "' and SDATE <= '" & end_date & "'"
objSQL = objSQL & " and CLOSING_MSGBOX not like '~PT%' "
If supervisor_selection <> "All Supervisors" Then
	objSQL = objSQL & " and L1Manager = '" & supervisor_selection & "' "
End If
If worker_windows_id <> "" Then objSQL = objSQL & " and USERNAME like '%" & worker_windows_id & "%' "
objSQL = objSQL & " GROUP BY BZS.USERNAME, BZS.SCRIPT_NAME "
objSQL = objSQL & " Order by 3 DESC"

'opening the connections and data table
objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
objRecordSet.Open objSQL, objConnection

Do While NOT objRecordSet.Eof
	' MsgBox "SCRIPT - " & objRecordSet("ScriptName") & vbCr & "Count - " & objRecordSet("ScriptRuns")
	ReDim preserve SCRIPT_USED_ARRAY(usage_last_const, unique_script_count)
	SCRIPT_USED_ARRAY(script_name_const, unique_script_count) = objRecordSet("ScriptName")
	SCRIPT_USED_ARRAY(user_number_const, unique_script_count) = objRecordSet("WFNumber")
	' SCRIPT_USED_ARRAY(user_name_const, unique_script_count) = objRecordSet("ScriptUser")
	SCRIPT_USED_ARRAY(count_const, unique_script_count) = objRecordSet("ScriptRuns")
	unique_script_count = unique_script_count + 1
	objRecordSet.MoveNext

Loop

objRecordSet.Close			'Closing all the data connections
objConnection.Close


Set objRecordSet=nothing
Set objConnection=nothing
' MsgBox "QUERY 2 DONE"

Set objConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
Set objRecordSet = CreateObject("ADODB.Recordset")

objSQL = ""
objSQL = "SELECT "
' objSQL = objSQL & "	ED.EmpFullName as ScriptUser,"
objSQL = objSQL & " BZS.USERNAME as WFNumber,"
objSQL = objSQL & "	BZS.CASE_NUMBER as CaseNumb,"
objSQL = objSQL & " BZS.SDATE as ScriptDate,"
objSQL = objSQL & "	COUNT(BZS.SCRIPT_NAME) as ScriptRuns"
objSQL = objSQL & " FROM 	"
objSQL = objSQL & "	BlueZone_Statistics.ES.ES_StaffHierarchyDim as ED"
objSQL = objSQL & "	LEFT JOIN BlueZone_Statistics.dbo.usage_log as BZS"
objSQL = objSQL & "	ON  ED.EmpLogOnID	= BZS.USERNAME"
objSQL = objSQL & "	"
objSQL = objSQL & " WHERE SDATE >= '" & start_date & "' and SDATE <= '" & end_date & "'"
objSQL = objSQL & " and DATALENGTH(BZS.CASE_NUMBER) > 0"
objSQL = objSQL & " and CLOSING_MSGBOX not like '~PT%' "
If supervisor_selection <> "All Supervisors" Then
	objSQL = objSQL & " and L1Manager = '" & supervisor_selection & "' "
End If
If worker_windows_id <> "" Then objSQL = objSQL & "and USERNAME like '%" & worker_windows_id & "%' "
objSQL = objSQL & " GROUP BY BZS.USERNAME, BZS.CASE_NUMBER, BZS.SDATE"
objSQL = objSQL & " Order by 3"

'opening the connections and data table
objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
objRecordSet.Open objSQL, objConnection

Do While NOT objRecordSet.Eof

	ReDim preserve CASE_SCRIPT_USAGE(usage_last_const, unique_case_count)
	CASE_SCRIPT_USAGE(case_numb_const, unique_case_count) = objRecordSet("CaseNumb")
	CASE_SCRIPT_USAGE(user_number_const, unique_case_count) = objRecordSet("WFNumber")
	CASE_SCRIPT_USAGE(script_date_const, unique_case_count) = objRecordSet("ScriptDate")
	' CASE_SCRIPT_USAGE(user_name_const, unique_case_count) = objRecordSet("ScriptUser")
	CASE_SCRIPT_USAGE(count_const, unique_case_count) = objRecordSet("ScriptRuns")
	unique_case_count = unique_case_count + 1
	objRecordSet.MoveNext

Loop

objRecordSet.Close			'Closing all the data connections
objConnection.Close


Set objRecordSet=nothing
Set objConnection=nothing


'ADD QUERY FOR AVERAGES
'Average script run per case by worker
'Average script run by day per worker
'MAX script count by day and by case
'Number of cases with 1 script run




For curr_counted = 0 to UBound(SCRIPT_USED_ARRAY, 2)
	For i = 0 to ubound(script_array)
		If InStr(UCASE(SCRIPT_USED_ARRAY(script_name_const, curr_counted)), UCASE(script_array(i).category & " - " & script_array(i).script_name)) <> 0 Then
			SCRIPT_USED_ARRAY(usage_category_const, curr_counted) = script_array(i).usage_eval
			SCRIPT_USED_ARRAY(script_category_const, curr_counted) = script_array(i).category
			SCRIPT_USED_ARRAY(script_array_index_const, curr_counted) = i
			' MsgBox script_array(i).usage_eval & vbCr & "Array name - " & SCRIPT_USED_ARRAY(script_name_const, curr_counted) & vbCr & "CLOS Name - " & UCASE(script_array(i).category & " - " & script_array(i).script_name)
			Exit For
		End If
	Next
Next

For use_case = 0 to UBound(SCRIPT_USAGE_ARRAY, 2)
	For i = 0 to ubound(script_array)
		If InStr(UCASE(SCRIPT_USAGE_ARRAY(script_name_const, use_case)), UCASE(script_array(i).category & " - " & script_array(i).script_name)) <> 0 Then
			SCRIPT_USAGE_ARRAY(usage_category_const, use_case) = script_array(i).usage_eval
			SCRIPT_USAGE_ARRAY(script_category_const, use_case) = script_array(i).category
			SCRIPT_USAGE_ARRAY(script_array_index_const, use_case) = i
			Exit For
		End If
	Next
	For each_worker = 0 to UBound (WORKER_ARRAY, 2)
		If UCASE(SCRIPT_USAGE_ARRAY(user_number_const, use_case)) = UCASE(WORKER_ARRAY(wrkr_id_const, each_worker)) Then
			If WORKER_ARRAY(wrkr_name_const, each_worker) = "" Then WORKER_ARRAY(wrkr_name_const, each_worker) = SCRIPT_USAGE_ARRAY(user_name_const, use_case)
			WORKER_ARRAY(total_script_count, each_worker) = WORKER_ARRAY(total_script_count, each_worker) + 1
			' MsgBox "Total: " & WORKER_ARRAY(total_script_count, each_worker) & vbCr & "CATEGORIES: " & SCRIPT_USAGE_ARRAY(script_category_const, use_case) & vbCr & "USAGE EVAL: " & SCRIPT_USAGE_ARRAY(usage_category_const, use_case)
			If SCRIPT_USAGE_ARRAY(script_category_const, use_case) = "NOTES" Then WORKER_ARRAY(NOTES_cnt_const, each_worker) = WORKER_ARRAY(NOTES_cnt_const, each_worker) + 1
			If SCRIPT_USAGE_ARRAY(script_category_const, use_case) = "ACTIONS" Then WORKER_ARRAY(ACTIONS_cnt_const, each_worker) = WORKER_ARRAY(ACTIONS_cnt_const, each_worker) + 1
			If SCRIPT_USAGE_ARRAY(script_category_const, use_case) = "NOTICES" Then WORKER_ARRAY(NOTICES_cnt_const, each_worker) = WORKER_ARRAY(NOTICES_cnt_const, each_worker) + 1
			If SCRIPT_USAGE_ARRAY(script_category_const, use_case) = "UTILITIES" Then WORKER_ARRAY(UTILITIES_cnt_const, each_worker) = WORKER_ARRAY(UTILITIES_cnt_const, each_worker) + 1
			If SCRIPT_USAGE_ARRAY(script_category_const, use_case) = "ADMIN" Then WORKER_ARRAY(ADMIN_cnt_const, each_worker) = WORKER_ARRAY(ADMIN_cnt_const, each_worker) + 1
			If SCRIPT_USAGE_ARRAY(script_category_const, use_case) = "BULK" Then WORKER_ARRAY(BULK_cnt_const, each_worker) = WORKER_ARRAY(BULK_cnt_const, each_worker) + 1
			If SCRIPT_USAGE_ARRAY(usage_category_const, use_case) = "PRIMARY" Then WORKER_ARRAY(eval_PRIMARY_cnt_const, each_worker) = WORKER_ARRAY(eval_PRIMARY_cnt_const, each_worker) + 1
			If SCRIPT_USAGE_ARRAY(usage_category_const, use_case) = "STAR" Then WORKER_ARRAY(eval_STAR_cnt_const, each_worker) = WORKER_ARRAY(eval_STAR_cnt_const, each_worker) + 1
			If SCRIPT_USAGE_ARRAY(usage_category_const, use_case) = "LIMITED" Then WORKER_ARRAY(eval_LIMITED_cnt_const, each_worker) = WORKER_ARRAY(eval_LIMITED_cnt_const, each_worker) + 1
			If SCRIPT_USAGE_ARRAY(usage_category_const, use_case) = "TEAM" Then WORKER_ARRAY(eval_TEAM_cnt_const, each_worker) = WORKER_ARRAY(eval_TEAM_cnt_const, each_worker) + 1
			If SCRIPT_USAGE_ARRAY(usage_category_const, use_case) = "ADMIN" Then WORKER_ARRAY(eval_ADMIN_cnt_const, each_worker) = WORKER_ARRAY(eval_ADMIN_cnt_const, each_worker) + 1
			If SCRIPT_USAGE_ARRAY(usage_category_const, use_case) = "FLAG" Then WORKER_ARRAY(eval_FLAG_cnt_const, each_worker) = WORKER_ARRAY(eval_FLAG_cnt_const, each_worker) + 1
			If SCRIPT_USAGE_ARRAY(usage_category_const, use_case) = "STATS" Then WORKER_ARRAY(eval_STATS_cnt_const, each_worker) = WORKER_ARRAY(eval_STATS_cnt_const, each_worker) + 1
			If SCRIPT_USAGE_ARRAY(usage_category_const, use_case) = "" Then WORKER_ARRAY(eval_BLANK_cnt_const, each_worker) = WORKER_ARRAY(eval_BLANK_cnt_const, each_worker) + 1
			'CREATE a string of all cases with DATE  with ELIG Summ run on them by worker name
		End If
	Next
Next

'Loop through count by case and date for each worker - if the case is on the list of cases where wlig summ was run - then we count the runs for that case.
'track the number of cases with ONLY Elig Summ run
'Track the number of cases with ELIG Summ and at least one other script

Do

	sup_dlg_len = 50 + (UBound(WORKER_ARRAY, 2)+1)*15
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 396, sup_dlg_len, "Worker Script Runs"
		ButtonGroup ButtonPressed
			OkButton 340, sup_dlg_len-20, 50, 15
			Text 10, 10, 70, 10, "WORKER NAME"
			Text 125, 10, 30, 10, "ID"
			Text 170, 10, 25, 10, "Total"
			Text 200, 10, 30, 10, "PRIMARY"
			Text 230, 10, 25, 10, "STAR"
			Text 255, 10, 25, 10, "ADMIN"
			Text 280, 10, 30, 10, "LIMITED"
			Text 305, 10, 25, 10, "FLAG"
			Text 335, 10, 25, 10, "STATS"
			Text 365, 10, 25, 10, "TEAM"
			y_pos = 25
			For each_worker = 0 to UBound(WORKER_ARRAY, 2)
				PushButton 15, y_pos-2, 100, 13, WORKER_ARRAY(wrkr_name_const, each_worker), WORKER_ARRAY(wrkr_btn_const, each_worker)
				Text 125, y_pos, 40, 10, WORKER_ARRAY(wrkr_id_const, each_worker)
				Text 170, y_pos, 15, 10, WORKER_ARRAY(total_script_count, each_worker)
				Text 200, y_pos, 15, 10, WORKER_ARRAY(eval_PRIMARY_cnt_const, each_worker)
				Text 230, y_pos, 15, 10, WORKER_ARRAY(eval_STAR_cnt_const, each_worker)
				Text 255, y_pos, 15, 10, WORKER_ARRAY(eval_ADMIN_cnt_const, each_worker)
				Text 280, y_pos, 15, 10, WORKER_ARRAY(eval_LIMITED_cnt_const, each_worker)
				Text 305, y_pos, 15, 10, WORKER_ARRAY(eval_FLAG_cnt_const, each_worker)
				Text 335, y_pos, 15, 10, WORKER_ARRAY(eval_STATS_cnt_const, each_worker)
				Text 365, y_pos, 15, 10, WORKER_ARRAY(eval_TEAM_cnt_const, each_worker)
				y_pos = y_pos + 15
			Next
	EndDialog

	err_msg = "Loop"

	dialog Dialog1
	cancel_confirmation

	For each_worker = 0 to UBound(WORKER_ARRAY, 2)
		If ButtonPressed = WORKER_ARRAY(wrkr_btn_const, each_worker) Then
			CALL worker_detail_dlg(each_worker)
		End If
	Next
	If ButtonPressed = -1 Then err_msg = ""
Loop until err_msg = ""


function worker_detail_dlg(worker_selection)
	grp_len_script = 20
	For curr_counted = 0 to UBound(SCRIPT_USED_ARRAY, 2)
		If UCASE(SCRIPT_USED_ARRAY(user_number_const, curr_counted)) =UCASE(WORKER_ARRAY(wrkr_id_const, worker_selection)) Then grp_len_script = grp_len_script + 10
	Next
	If grp_len_script = 20 Then grp_len_script = 30

	grp_len_case = 20
	For each_case = 0 to UBound(CASE_SCRIPT_USAGE, 2)
		If UCASE(CASE_SCRIPT_USAGE(user_number_const, each_case)) = UCASE(WORKER_ARRAY(wrkr_id_const, worker_selection)) Then grp_len_case = grp_len_case + 10
	Next
	If grp_len_case = 20 Then grp_len_case = 30
	If grp_len_case > 320 Then grp_len_case = 320

	If grp_len_case > grp_len_script Then
		dlg_len = 80 + grp_len_case
	Else
		dlg_len = 80 + grp_len_script
	End If

	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 431, dlg_len, "Script Use Count for " & WORKER_ARRAY(wrkr_name_const, worker_selection)
		ButtonGroup ButtonPressed
			PushButton 375, dlg_len-20, 50, 15, "BACK", back_btn
		Text 10, 10, 80, 10, "Total Usage: " & WORKER_ARRAY(total_script_count, worker_selection)
		Text 105, 10, 70, 10, "NOTES Usage: " & WORKER_ARRAY(NOTES_cnt_const, worker_selection)
		Text 105, 20, 70, 10, "ACTIONS Usage: " & WORKER_ARRAY(ACTIONS_cnt_const, worker_selection)
		Text 105, 30, 70, 10, "NOTICES Usage: " & WORKER_ARRAY(NOTICES_cnt_const, worker_selection)
		Text 180, 10, 70, 10, "UTILITIES Usage: " & WORKER_ARRAY(UTILITIES_cnt_const, worker_selection)
		Text 180, 20, 70, 10, "BULK Usage: " & WORKER_ARRAY(BULK_cnt_const, worker_selection)
		Text 180, 30, 70, 10, "ADMIN Usage: " & WORKER_ARRAY(ADMIN_cnt_const, worker_selection)
		Text 275, 10, 75, 10, "PRIMARY Usage: " & WORKER_ARRAY(eval_PRIMARY_cnt_const, worker_selection)
		Text 275, 20, 75, 10, "STAR Usage: " & WORKER_ARRAY(eval_STAR_cnt_const, worker_selection)
		Text 275, 30, 75, 10, "LIMITED  Usage: " & WORKER_ARRAY(eval_LIMITED_cnt_const, worker_selection)
		Text 275, 40, 75, 10, "TEAM Usage: " & WORKER_ARRAY(eval_TEAM_cnt_const, worker_selection)
		Text 350, 10, 70, 10, "ADMIN Usage: " & WORKER_ARRAY(eval_ADMIN_cnt_const, worker_selection)
		Text 350, 20, 70, 10, "FLAG Usage: " & WORKER_ARRAY(eval_FLAG_cnt_const, worker_selection)
		Text 350, 30, 70, 10, "STATS Usage: " & WORKER_ARRAY(eval_STATS_cnt_const, worker_selection)
		y_pos = 70
		For curr_counted = 0 to UBound(SCRIPT_USED_ARRAY, 2)
			If UCASE(SCRIPT_USED_ARRAY(user_number_const, curr_counted)) =UCASE(WORKER_ARRAY(wrkr_id_const, worker_selection)) Then
				Text 15, y_pos, 200, 10, SCRIPT_USED_ARRAY(script_name_const, curr_counted)
				Text 215, y_pos, 50, 10, SCRIPT_USED_ARRAY(usage_category_const, curr_counted)
				Text 265, y_pos, 20, 10, SCRIPT_USED_ARRAY(count_const, curr_counted)
				y_pos = y_pos + 10
			End If
			If y_pos >360 Then
				If curr_counted < UBound(SCRIPT_USED_ARRAY, 2) Then
					Text 300, y_pos, 50, 10, "MORE +  "
				End If
				Exit For
			End If
		Next
		If y_pos = 70 Then Text 15, 70, 120, 10, "NO SCRIPT RUNS FOUND"

		y_pos = 70
		For each_case = 0 to UBound(CASE_SCRIPT_USAGE, 2)
			If UCASE(CASE_SCRIPT_USAGE(user_number_const, each_case)) = UCASE(WORKER_ARRAY(wrkr_id_const, worker_selection)) Then
				Text 300, y_pos, 50, 10, CASE_SCRIPT_USAGE(case_numb_const, each_case)
				Text 350, y_pos, 50, 10, CASE_SCRIPT_USAGE(script_date_const, each_case)
				Text 400, y_pos, 15, 10, CASE_SCRIPT_USAGE(count_const, each_case)
				y_pos = y_pos + 10
			End If
			If y_pos >360 Then
				If each_case < UBound(CASE_SCRIPT_USAGE, 2) Then
					Text 300, y_pos, 50, 10, "MORE +  "
				End If
				Exit For
			End If
		Next
		If y_pos = 70 Then Text 235, 70, 150, 10, "NO CASE SPECIFIC SCRIPT RUNS FOUND"

		GroupBox 10, 55, 275, grp_len_script, "Count by Script"
		GroupBox 295, 55, 115, grp_len_case, "Count by Case Number"
	EndDialog

	Dialog Dialog1

	ButtonPressed = back_btn
end function


Call script_end_procedure("That's It")