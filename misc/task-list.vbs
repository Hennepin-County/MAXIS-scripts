'Required for statistical purposes==========================================================================================
name_of_script = "TASKS - Task List.vbs"
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

date_to_review = date & ""

Do
	BeginDialog Dialog1, 0, 0, 356, 150, "Dialog"
	  EditBox 300, 5, 50, 15, date_to_review
	  ButtonGroup ButtonPressed
	    PushButton 305, 25, 45, 10, "UPDATE", Button3
	    ' CancelButton 245, 130, 50, 15
	    OkButton 300, 130, 50, 15
	  Text 5, 10, 135, 10, "Task Completion for WORKER NAME"
	  Text 245, 10, 50, 10, "date to review:"
	  GroupBox 5, 40, 345, 70, "Tasks completed on " & date_to_review
	  Text 15, 50, 25, 10, "Case #"
	  Text 65, 50, 40, 10, "Time Spent"
	  Text 130, 50, 40, 10, "Interview"
	  Text 190, 50, 35, 10, "Approved"
	  Text 250, 50, 40, 10, "CASE:NOTE"
	  Text 15, 65, 40, 10, "-XXXXXXX"
	  Text 75, 65, 25, 10, "h:mm"
	  Text 140, 65, 15, 10, "No"
	  Text 200, 65, 15, 10, "Yes"
	  Text 265, 65, 15, 10, "Yes"
	  ButtonGroup ButtonPressed
	    PushButton 295, 65, 50, 10, "CHANGE", Button5
	  Text 15, 75, 40, 10, "-XXXXXXX"
	  Text 75, 75, 25, 10, "h:mm"
	  Text 140, 75, 15, 10, "Yes"
	  Text 200, 75, 15, 10, "Yes"
	  Text 265, 75, 15, 10, "Yes"
	  ButtonGroup ButtonPressed
	    PushButton 295, 75, 50, 10, "CHANGE", Button11
	  Text 15, 85, 40, 10, "-XXXXXXX"
	  Text 75, 85, 25, 10, "h:mm"
	  Text 140, 85, 15, 10, "Yes"
	  Text 200, 85, 15, 10, "No"
	  Text 265, 85, 15, 10, "Yes"
	  ButtonGroup ButtonPressed
	    PushButton 295, 85, 50, 10, "CHANGE", Button12
	  Text 15, 95, 40, 10, "-XXXXXXX"
	  Text 75, 95, 25, 10, "h:mm"
	  Text 140, 95, 15, 10, "No"
	  Text 200, 95, 15, 10, "Yes"
	  Text 265, 95, 15, 10, "Yes"
	  ButtonGroup ButtonPressed
	    PushButton 295, 95, 50, 10, "CHANGE", Button13
	  GroupBox 10, 115, 180, 30, "Counts by case"
	  Text 15, 130, 50, 10, "Interviews: 2"
	  Text 70, 130, 50, 10, "Approvals: 3"
	  Text 130, 130, 50, 10, "CASE:NOTE: 4"
	EndDialog

	dialog Dialog1
	' cancel_without_confirmation
Loop until ButtonPressed = -1

call script_end_procedure("")
