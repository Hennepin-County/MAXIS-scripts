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


BeginDialog Dialog1, 0, 0, 361, 135, "Log Activity"
  GroupBox 10, 25, 345, 85, "Activity in Progress"
  Text 25, 40, 85, 10, "Date: DATE"
  Text 25, 50, 85, 10, "Start Time: TIME"
  Text 25, 70, 190, 10, "Category: CATEGORY"
  Text 25, 80, 185, 10, "Detail: DETAIL"
  Text 230, 40, 65, 10, "Meeting? YES"
  Text 230, 55, 115, 10, "Project: PROJECT"
  Text 230, 95, 95, 10, "Elapsed Time: HH:MM"
  ButtonGroup ButtonPressed
    PushButton 260, 75, 85, 15, "GitHub Issue #", git_hub_issue_button
    PushButton 135, 5, 65, 15, "Switch Activity", switch_activity_button
    PushButton 205, 5, 60, 15, "Start Break", start_break_button
    PushButton 270, 5, 85, 15, "End Work Day", end_work_day_button
    OkButton 255, 115, 50, 15
    CancelButton 305, 115, 50, 15
EndDialog
