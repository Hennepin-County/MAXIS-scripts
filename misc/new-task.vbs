'Required for statistical purposes==========================================================================================
name_of_script = "TASKS - New Task.vbs"
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

EMConnect ""

BeginDialog Dialog1, 0, 0, 296, 165, "New Task Assignment"
  Text 10, 10, 135, 10, "Thank you for requesting a new task!"
  Text 30, 25, 65, 10, "Case: XXXXXXX"
  Text 50, 35, 125, 10, "(Case has been entered into MAXIS)"
  GroupBox 5, 50, 285, 90, "Case Overview"
  Text 15, 65, 265, 10, "Active Programs: SNAP"
  Text 15, 80, 265, 10, "Pending Programs: Cash"
  Text 15, 95, 65, 10, "HH Members; 2"
  Text 15, 110, 65, 10, "DAILs found: 5"
  Text 15, 125, 100, 10, "REVW Month: SNAP - 08/21"
  Text 215, 10, 70, 10, "Assignment Details:"
  Text 220, 25, 75, 10, "- Assigned on m/d/yy"
  Text 220, 35, 70, 10, "- Assigned at hh:mm"
  ButtonGroup ButtonPressed
    ' CancelButton 185, 145, 50, 15
    OkButton 240, 145, 50, 15
EndDialog

Do
	dialog Dialog1
	' cancel_without_confirmation
Loop until ButtonPressed = -1

call script_end_procedure("")
