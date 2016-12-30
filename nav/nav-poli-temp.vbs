'Required for statistical purposes==========================================================================================
name_of_script = "NAV - POLI-TEMP.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 20                      'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS--------------------------------------------------
BeginDialog POLI_TEMP_dialog, 0, 0, 211, 75, "POLI/TEMP dialog"
  DropListBox 35, 25, 55, 45, "TABLE"+chr(9)+"INDEX", Temp_table_index
  ButtonGroup ButtonPressed
    OkButton 95, 55, 50, 15
    CancelButton 155, 55, 50, 15
  Text 5, 10, 140, 15, "What area of POLI/TEMP you want to go?"
  Text 5, 25, 30, 10, "Select:"
  Text 95, 25, 105, 10, "TABLE - Search by TEMP code"
  Text 95, 35, 115, 10, "INDEX - Search by a word or topic"
EndDialog

'THE SCRIPT

'Displays dialog
Dialog POLI_TEMP_dialog
If buttonpressed = cancel then stopscript

'Determines which POLI/TEMP section to go to, using the dropdown list outcome to decide
If Temp_table_index = "TABLE" then
	panel_title = "TABLE"
ElseIf Temp_table_index = "INDEX" then
	panel_title = "INDEX"
End if

'Connects to BlueZone
EMConnect ""

'call screen back to SELF screen to proceed onward with POLI
'navigating back to SELF menu, since back_to_SELF does not work in POLI function
DO
	PF3
	EMReadScreen SELF_check, 4, 2, 50
Loop until SELF_check = "SELF"

'Checks to make sure we're in MAXIS
call check_for_MAXIS(True)

'Navigates to POLI (can't direct navigate to TEMP)
call navigate_to_MAXIS_screen("POLI", "____")

'Writes TEMP
EMWriteScreen "TEMP", 5, 40

'Writes the panel_title selection
EMWriteScreen panel_title, 21, 71

'Transmits
transmit

script_end_procedure("")
