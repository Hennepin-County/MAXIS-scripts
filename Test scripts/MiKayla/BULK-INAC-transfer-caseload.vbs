'Required for statistical purposes==========================================================================================
name_of_script = "BULK - CASELOAD INACTIVE TRANSFER.vbs"
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

current_worker = ""
new_worker = "x127CCL"
MAXIS_case_number = "2335052"

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail

BeginDialog Dialog1, 0, 0, 131, 85, "BULK INACTIVE"
  EditBox 60, 5, 65, 15, MAXIS_case_number
  EditBox 60, 25, 65, 15, current_worker
  EditBox 60, 45, 65, 15, new_worker
  ButtonGroup ButtonPressed
    OkButton 40, 65, 40, 15
    CancelButton 85, 65, 40, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 30, 50, 10, "Current worker:"
  Text 5, 50, 45, 10, "New worker:"
EndDialog



DIALOG Dialog1
cancel_confirmation

CALL navigate_to_MAXIS_screen ("SPEC", "XFER")
EMWriteScreen "2335052", 18, 43 'MAXIS_case_number'
TRANSMIT
'Dummy case number initially
EMWriteScreen "X", 11, 16 'Transfer case load same county
TRANSMIT
'-----------------------------------------------XCLD
'msgbox "where am I now"
EMWriteScreen current_worker, 04, 18
TRANSMIT
'IF current_worker = "" then pf3
EMWriteScreen "X127CCL", 15, 13 'new_worker'
TRANSMIT
EMWriteScreen "x", 11, 13 'inactive_case
'Change the footer month
TRANSMIT 'REVIEW NEW WORKER NAME AND PRESS ENTER TO VIEW DETAILS
TRANSMIT
'msgbox "AM I moving"
'-------------------------------------------XFER
row = 7
DO
	DO
		EMReadScreen previous_number, 7, row, 5         'First it reads the case number, name, date they closed, and the APPL date.

		IF previous_number <> "" THEN
			case_found = TRUE
			EMWriteScreen "1", row, 3
			row = row + 1
		END IF
	Loop until row = 19
	row = 7 'Setting the variable for when the do...loop restarts
	PF8
	IF previous_number = "" THEN case_found = FALSE
	EMReadScreen last_page_check, 4, 24, 2 'checks for "THIS IS THE LAST PAGE"
	IF last_page_check = "THIS" or last_page_check = "MAXI" THEN case_found = FALSE
LOOP UNTIL case_found = FALSE
	'IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "What does it say?"
'LOOP UNTIL err_msg = ""
PF3 'to save
'MAXIMUM PAGES ALLOWED IS 100
'THIS IS LAST PAGE- PF3 FOR THE NEXT CASE STATUS OR ENTER START NAME

script_end_procedure("Success!")
