'STATS GATHERING=============================================================================================================
name_of_script = "ADMIN - FIND HIDDEN EXCEL.vbs"       'Replace TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer
STATS_counter = 0               'sets the stats counter at one
STATS_manualtime = 90            'manual run time in seconds
STATS_denomination = "I"        'C is for each case
'END OF stats block==========================================================================================================

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


'THE SCRIPT==================================================================================================================
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 180, "Find Hidden Excel Files"
  ButtonGroup ButtonPressed
    OkButton 155, 160, 50, 15
    CancelButton 210, 160, 50, 15
  Text 10, 10, 255, 10, "This script is intended to find any Excel files that are open on your computer."
  Text 10, 25, 250, 20, "Sometimes Excel files can be open and invisible on your computer. It is difficult to find these files manually. "
  Text 10, 50, 245, 20, "The script will find the files one at a time and make them visible, so you can decide what to do with them. "
  GroupBox 10, 75, 245, 75, "IMPORTANT"
  Text 20, 90, 230, 10, "This script works best when there are NO VISIBLE EXCEL Files open."
  Text 20, 105, 145, 10, "Close all Excel Files now.Save as needed."
  Text 20, 120, 220, 25, "As the script makes Excel Files visible, take required action (save as needed) and close those files. The script will search again for more files until none are found."
EndDialog

dialog Dialog1

On Error Resume Next
Do
	STATS_counter = STATS_counter + 1
	Set objXl = GetObject(, "Excel.Application")
	file_name = objXL.ActiveWorkbook.Name
	If Err Then
	    ' MsgBox "Excel NOT Running", vbInformation, "Excel Status"
	    WScript.Quit(-1)
		Call script_end_procedure("No Excel File was found open on your computer. The script will end, there are no hidden or visible Excel Files the script can find.")
	End If

	If Not TypeName(objXL) = "Empty" then
		objXl.Visible = True
		objXl.WindowState = -4137			'Excel Ennumeration can be found here -  https://docs.microsoft.com/en-us/office/vba/api/excel.xlwindowstate
		MsgBox "Excel Running - " & file_name & " is active" & vbCr & vbCr & "It has been made visible." & vbCr & vbCr & "Review the file, save as needed, and close it now. Only press OK once the files is closed."
	End If
	' Set objXl = ""
Loop

Call script_end_procedure("All Excel Files found.")
