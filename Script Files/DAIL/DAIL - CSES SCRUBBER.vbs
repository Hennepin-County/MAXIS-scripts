'GATHERING STATS===========================================================================================================
name_of_script = "DAIL - CSES SCRUBBER.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY==========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
'END FUNCTIONS LIBRARY BLOCK===============================================================================================

'Required for statistical purposes=========================================================================================
STATS_counter = 0              'sets the stats counter at 0 because each iteration of the loop which counts the dail messages adds 1 to the counter.
STATS_manualtime = 54          'manual run time in seconds
STATS_denomination = "I"       'I is for each dail message
'END OF stats block========================================================================================================

'DIALOGS===================================================================================================================
BeginDialog CSES_initial_dialog, 0, 0, 296, 40, "CSES Dialog"
  CheckBox 5, 5, 290, 10, "Check here if you would like to see an Excel sheet of the CSES scrubber calculations.", excel_visible_checkbox
  EditBox 70, 20, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 185, 20, 50, 15
    CancelButton 240, 20, 50, 15
  Text 5, 25, 65, 10, "Worker signature:"
EndDialog
'END DIALOGS===============================================================================================================
'THE SCRIPT================================================================================================================

'Connects to MAXIS
EMConnect ""

'Displays initial dialog (script assumes you're on a CSES message by virtue of the DAIL scrubber). Dialog has no mandatory fields, so there's no loop.
Dialog CSES_initial_dialog
If ButtonPressed = cancel then stopscript

'If the worker signature is the Konami code (UUDDLRLRBA), developer mode will be triggered
If worker_signature = "UUDDLRLRBA" then
    MsgBox "Developer mode: ACTIVATED!"
    developer_mode = true           'This will be helpful later
    collecting_statistics = false   'Lets not collect this, shall we?
    excel_visible_checkbox = 1      'Forces this to be checked
End if

'If excel_visible_checkbox is checked, it'll set the visibility for Excel to "true". Otherwise it'll be false.
If excel_visible_checkbox = 1 then
    excel_visibility = true
Else
    excel_visibility = false
End if

'Checks if you're on a TIKL, and asks if this is a training scenario
EMReadScreen TIKL_check, 4, 6, 6
If TIKL_check = "TIKL" then
    TIKL_processing_confirmation = MsgBox("You seem to be running this on a TIKL. Are you testing the script? If you aren't, something has gone wrong.", vbYesNo)
    If TIKL_processing_confirmation = vbNo then stopscript
End if

'~~~~~~~~~~~~~~~~~~~~Reads each message

'~~~~~~~~~~~~~~~~~~~~Sorts each message into Excel column by PMI (and divides each into a share of each message based on total PMIs)

'~~~~~~~~~~~~~~~~~~~~Navigates to CASE/CURR to determine programs open

'~~~~~~~~~~~~~~~~~~~~Decision: Is MFIP/DWP open? IF YES...

    '~~~~~~~~~~~~~~~~~~~~Displays prospective estimated budget based on DAILs received

    '~~~~~~~~~~~~~~~~~~~~Decision: Does user want to update? IF YES...

        '~~~~~~~~~~~~~~~~~~~~Script updates UNEA for all messages with prospective amounts and actual amounts for retrospective budgeting

'~~~~~~~~~~~~~~~~~~~~Decision: Is SNAP open? IF YES...

    '~~~~~~~~~~~~~~~~~~~~Displays total and current PIC, user decides if it’s within the realm for each message

'~~~~~~~~~~~~~~~~~~~~Decision: Any panels updated/checked for either SNAP or MFIP? IF YES...

    '~~~~~~~~~~~~~~~~~~~~Case note details from Excel sheet, and all panels updated

'~~~~~~~~~~~~~~~~~~~~Decision: Does user want Excel breakdown of info? IF YES...

    '~~~~~~~~~~~~~~~~~~~~Make Excel visible

script_end_procedure("")
