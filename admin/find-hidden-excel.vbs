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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("05/26/2022", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT==================================================================================================================
'This dialog doesn't capture any variables - there is no input needed for this script. This is simply for information.
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 180, "Find Hidden Excel Files"
  ButtonGroup ButtonPressed
    OkButton 155, 160, 50, 15
    CancelButton 210, 160, 50, 15
	PushButton 10, 160, 80, 15, "INSTRUCTIONS", instructions_btn
  Text 10, 10, 255, 10, "This script is intended to find any Excel files that are open on your computer."
  Text 10, 25, 250, 20, "Sometimes Excel files can be open and invisible on your computer. It is difficult to find these files manually. "
  Text 10, 50, 245, 20, "The script will find the files one at a time and make them visible, so you can decide what to do with them. "
  GroupBox 10, 75, 245, 75, "IMPORTANT"
  Text 20, 90, 230, 10, "This script works best when there are NO VISIBLE EXCEL Files open."
  Text 20, 105, 145, 10, "Close all Excel Files now.Save as needed."
  Text 20, 120, 220, 25, "As the script makes Excel Files visible, take required action (save as needed) and close those files. The script will search again for more files until none are found."
EndDialog

Do
	dialog Dialog1					''showing the dialog
	cancel_without_confirmation
	If ButtonPressed = instructions_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3315A-ENG"
Loop until ButtonPressed = -1

On Error Resume Next			'this is needed because an error indicates that no excel files have been found. Instead of throwing an error, the script is coded to script_end when the error hits
Do
	STATS_counter = STATS_counter + 1					'incrementing the stats
	Set objXl = GetObject(, "Excel.Application")		'try to find an excel file
	file_name = objXL.ActiveWorkbook.Name				'set the name to a variable - THIS IS WHAT WILL THROW AN ERROR IF NO EXCEL IS OPEN
	If Err Then											'the script knows an error has been thrown and will stop the script
	    WScript.Quit(-1)
		Call script_end_procedure("No Excel File was found open on your computer. The script will end, there are no hidden or visible Excel Files the script can find.")
	End If

	If Not TypeName(objXL) = "Empty" then				'If the type is not empty - then an excel exists
		objXl.Visible = True							'set to visible and maximize the window
		objXl.WindowState = -4137			'Excel Ennumeration can be found here -  https://docs.microsoft.com/en-us/office/vba/api/excel.xlwindowstate
		'this message is the stop point in the script to let the user know to address the Excel File
		continue_msg = MsgBox("Excel Running - " & file_name & " is active" & vbCr & vbCr & "It has been made visible." & vbCr & vbCr & "Review the file, save as needed, and close it now. Only press OK once the files is closed.", vbImportant + VBOkCancel, "Excel File Found")
		If continue_msg = VBCancel Then script_end_procedure("Review of Excel Files cancelled. You have cancelled the search for more Excel Files, there may be more still open but hidden. You can run the script again if needed.")
	End If
Loop				'There is no end condition on this loop because we will always hit a script end procedure

Call script_end_procedure("All Excel Files found.")
