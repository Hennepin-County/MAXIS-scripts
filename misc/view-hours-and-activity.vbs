'Required for statistical purposes==========================================================================================
name_of_script = "TASKS - DASH.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================

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

excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Time Tracking"

If user_ID_for_validation = "CALO001" Then excel_file_path = excel_file_path & "\Casey Time Tracking 2021.xlsx"
If user_ID_for_validation = "ILFE001" Then excel_file_path = excel_file_path & "\Ilse Time Tracking 2021.xlsx"
If user_ID_for_validation = "WFS395" Then excel_file_path = excel_file_path & "\MiKayla Time Tracking 2021.xlsx"

Call excel_open(excel_file_path, False, False, ObjExcel, objWorkbook)

on_task = False

'find the open ended line
excel_row = 2
Do
	row_date = ObjExcel.Cells(excel_row, 1).Value
	row_start_time = ObjExcel.Cells(excel_row, 2).Value
	row_end_time = ObjExcel.Cells(excel_row, 3).Value
	row_start_time = row_start_time * 24
	row_end_time = row_end_time * 24

	MsgBox "Date - " & row_date & vbCr & "Start time - " & row_start_time & vbCr & "End time - " & row_end_time

	excel_row = excel_row + 1
Loop until row_date = ""
ObjExcel.Quit


stopscript

BeginDialog Dialog1, 0, 0, 361, 245, "View Hours and Activity"
  EditBox 180, 15, 50, 15, selected_date
  ButtonGroup ButtonPressed
    CancelButton 305, 225, 50, 15
  GroupBox 5, 5, 275, 210, "Hours Breakdown"
  Text 85, 5, 110, 10, "TIME PERIOD"
  Text 15, 35, 70, 10, "Total Hours Logged:"
  Text 15, 55, 65, 10, "Hours in Meetings:"
  GroupBox 15, 75, 255, 130, "Hours by "
  ButtonGroup ButtonPressed
    PushButton 235, 15, 40, 15, "Switch", switch_button
    PushButton 15, 200, 60, 10, "CATEGORY", category_button
    PushButton 75, 200, 60, 10, "PROJECT", project_button
    PushButton 135, 200, 60, 10, "GITHUB ISSUE", git_hub_issue_button
    PushButton 300, 10, 50, 15, "Today", today_button
    PushButton 300, 25, 50, 15, "Week", week_button
    PushButton 300, 40, 50, 15, "Pay Period", pay_period_button
    PushButton 300, 55, 50, 15, "Month", month_button
    PushButton 300, 70, 50, 15, "Custom", custom_time_button
    OkButton 255, 225, 50, 15
EndDialog
