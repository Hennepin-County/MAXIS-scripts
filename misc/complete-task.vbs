'Required for statistical purposes==========================================================================================
name_of_script = "TASKS - Complete Task.vbs"
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


BeginDialog Dialog1, 0, 0, 336, 210, "Task Completion Information"
  DropListBox 170, 20, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", List2
  DropListBox 170, 35, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", List3
  DropListBox 15, 65, 180, 45, "No follow up needed"+chr(9)+"Yes - policy and process questions"+chr(9)+"Yes - Unique/specific task assignment"+chr(9)+"Yes - Assingment/work question", List1
  CheckBox 20, 100, 70, 10, "MFIP Sanctions", Check1
  CheckBox 95, 100, 70, 10, "Immigration", Check3
  CheckBox 175, 100, 70, 10, "More...", Check5
  CheckBox 20, 115, 70, 10, "Facility", Check2
  CheckBox 95, 115, 70, 10, "Overpayment", Check4
  TextBox 175, 115, 100, 15, more_detail_var
  DropListBox 130, 130, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", List4
  DropListBox 130, 145, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", List5
  DropListBox 130, 160, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", List6
  DropListBox 140, 175, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", List7
  CheckBox 15, 195, 90, 10, "Send updates to METS", Check6
  ButtonGroup ButtonPressed
    ' CancelButton 225, 190, 50, 15
    OkButton 280, 190, 50, 15
  Text 5, 10, 135, 10, "Task Completion for Case: XXXXXXX"
  Text 10, 25, 155, 10, "Was there work to be completed on this case?"
  Text 10, 40, 150, 10, "Were you able to complete all the work?"
  Text 10, 55, 105, 10, "Does this case need follow up?"
  Text 15, 85, 290, 10, "If this task should be handled by a  specialty group, select the appropriate actions here:"
  Text 15, 135, 110, 10, "Did you complete an interview?"
  Text 15, 150, 110, 10, "Did you 'APP' in ELIG?"
  Text 15, 165, 110, 10, "Did you CASE:NOTE?"
  Text 15, 180, 125, 10, "Did you send ECF Docs to the client?"
  Text 245, 10, 70, 10, "Assignment Details:"
  Text 250, 25, 75, 10, "- Completed on m/d/yy"
  Text 250, 35, 80, 10, "- Completed at hh:mm"
  ' Text 250, 45, 80, 10, "- Time Spent - h:mm"
EndDialog
Do
	dialog Dialog1
	' cancel_without_confirmation
Loop until ButtonPressed = -1

call script_end_procedure("")
