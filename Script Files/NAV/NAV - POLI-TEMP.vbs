'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NAV - POLI-TEMP.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" OR default_directory = "" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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

'DIALOGS--------------------------------------------------
BeginDialog POLI_TEMP_dialog, 0, 0, 256, 60, "POLI/TEMP dialog"
  OptionGroup RadioGroup1
    RadioButton 5, 30, 175, 10, "Table of Contents (search by TEMP section code)", table_radio
    RadioButton 5, 45, 150, 10, "Index of Topics (search by a word or topic)", index_radio
  ButtonGroup ButtonPressed
    OkButton 195, 10, 50, 15
    CancelButton 195, 30, 50, 15
  Text 10, 10, 160, 10, "What area of POLI/TEMP do you want to go to?"
EndDialog


'THE SCRIPT

'Displays dialog
Dialog POLI_TEMP_dialog
If buttonpressed = cancel then stopscript

'Determines which POLI/TEMP section to go to, using the radioboxes outcome to decide
If radiogroup1 = table_radio then 
	panel_title = "TABLE"
ElseIf radiogroup1 = index_radio then
	panel_title = "INDEX"
End if


'Connects to BlueZone
EMConnect ""

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
