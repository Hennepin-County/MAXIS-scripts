'Required for statistical purposes===============================================================================
name_of_script = "UTILITIES - PRISM SCREEN FINDER.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 10                      'manual run time in seconds
STATS_denomination = "I"                   'I is for each instance

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

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog Dialog1, 0, 0, 261, 135, "PRISM screen finder"
  ButtonGroup ButtonPressed
    CancelButton 210, 120, 50, 15
    PushButton 140, 70, 45, 10, "DDPL", DDPL_button
    PushButton 140, 40, 45, 10, "CAAD", CAAD_button
    PushButton 140, 55, 45, 10, "CAFS", CAFS_button
    PushButton 140, 85, 45, 10, "GCSC", GCSC_button
    PushButton 140, 115, 45, 10, "PESE", PESE_button

  Text 35, 70, 90, 10, "Direct deposit listing:"
  Text 35, 40, 65, 10, "Case notes:"
  Text 35, 55, 100, 10, "Case financial summary:"
  Text 35, 85, 100, 10, "Good cause/safety concerns:"
  Text 35, 115, 65, 10, "Person search:"
  Text 10, 0, 250, 25, "Press a button below to navigate to PRISM screens.  Then press F1 in the case number or MCI number field to select the participant or case information you are looking for."
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connect to BlueZone
EMConnect ""

CALL check_for_PRISM(FALSE)

DO

	Dialog

	'Now it'll navigate to any of the screens chosen
	If buttonpressed = DDPL_button then call navigate_to_PRISM_screen("DDPL")
	If buttonpressed = CAAD_button then call navigate_to_PRISM_screen("CAAD")
	If buttonpressed = CAFS_button then call navigate_to_PRISM_screen("CAFS")
	If buttonpressed = GCSC_button then call navigate_to_PRISM_screen("GCSC")
	If buttonpressed = PESE_button then call navigate_to_PRISM_screen("PESE")
LOOP until buttonpressed = 0

script_end_procedure("")
