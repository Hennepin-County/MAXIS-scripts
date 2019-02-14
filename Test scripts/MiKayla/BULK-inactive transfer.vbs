'Required for statistical purposes==========================================================================================
name_of_script = "BULK - INACTIVE TRANSFER.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 229                	'manual run time in seconds
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK=================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("02/14/2019", "Initial version.", "MiKayla Handley")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'------------------------------------------------------------------------THE SCRIPT
EMConnect ""

check_for_MAXIS(True)

BeginDialog transfer_dialog, 0, 0, 171, 70, "Transfer"
  EditBox 55, 5, 35, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    PushButton 110, 5, 50, 15, "Geocoder", Geo_coder_button
  EditBox 55, 25, 20, 15, spec_xfer_worker
  ButtonGroup ButtonPressed
    OkButton 65, 50, 45, 15
    CancelButton 115, 50, 45, 15
  Text 5, 30, 40, 10, "Transfer to:"
  Text 5, 10, 50, 10, "Case Number:"
  Text 80, 30, 60, 10, " (last 3 digit of X#)"
EndDialog


DIALOG xfer_menu_dialog
cancel_confirmation

call MAXIS_case_number_finder(MAXIS_case_number)

'Transfers case
back_to_self
EMWriteScreen "spec", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen MAXIS_case_number, 18, 43
EMWriteScreen "xfer", 21, 70


'Dummy case number initially
EMWriteScreen "X" 11, 16 'Transfer case load same county
transmit
Emreadscreen current_worker 7, 04, 18
IF current_worker = "" then pf3
EMWriteScreen inactive_case "x" 11, 13
'Change the footer month
EMWriteScreen transfer_to 15, 13
Transmit
Row = 7
DO
	EMWriteScreen "1", row, 7
'12 1 at 7, and last row is 18
PF8
if row = "M"
Emreadscreen for erorr msg  30, 24, 02
'MAXIMUM PAGES ALLOWED IS 100
Pf3 to save



DO
	DO
		DIALOG out_of_county_dlg
			cancel_confirmation
			IF ButtonPressed = nav_to_xfer_button THEN
				CALL navigate_to_MAXIS_screen("SPEC", "XFER")
				EMWriteScreen "X", 9, 16
				transmit
			END IF
	LOOP UNTIL ButtonPressed = -1
		last_chance = MsgBox("Do you want to continue? NOTE: You will get a chance to review SPEC/XFER before transmitting to transfer.", vbYesNo)
LOOP UNTIL last_chance = vbYes

script_end_procedure("Success!")
