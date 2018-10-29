'Required for statistical purposes==========================================================================================
name_of_script = "AUTOMATE - Paperless IR.vbs"
start_time = timer
STATS_counter = 0                          'sets the stats counter at one
STATS_manualtime = 1                       'manual run time in seconds
STATS_denomination = "M"       							'C is for each CASE
'END OF stats block==============================================================================================

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
call changelog_update("09/21/2018", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'SCRIPT MAP ================================================================================================================

'2 options - 1 to collect the list and 1 to process the list
'Have row to start and row/hour limiters to make sure the cases can all be acted upon.

'OPTION 1 - Gather the List
'Use the BULK - Paperless IR script to gather a list from REPT/REVW of all of the IR cases - put in an Array
'Confirm no EARNED income and no varying income
'Confirm they are at IR and not ER
'Dump all of the array into an Excel spreadsheet

'OPTION 2 - Act on the cases
'Dialog to select ecxel, set start row and limit run hours or rows
'Create an array of each case in the spreadsheet
'Go in to each case and confirm no earned income and no varying income - take from BULK Paperless IR pluss look at UNEA for income that is possibly varrying
'Confirm the number of people in the case that are on HC
'Look for Waiver/LTC cases identifiers
'Look to see if MEDI exists for any HH member
'Send the case through background - updating REVW
'Confirm no STAT edits exist
'Go to HC WLIG
    'Confirm that there is a budget for each HH member
    'Confirm there is a MSP for any HH member with MEDI
    'Confirm the budget starts with CM + 1 and has 6 months - if not possibly update budget OR add to manual action to have budget reviewed
    'approve each budget
'Send a MEMO ????
'Confirm the memo
'Case note the approval
'Review STAT to make sure REVW has been updated
'check MMIS to make sure the spans match
'Add all case action information to Excel Spreadsheet
'Go to the next case

'Option 2 will reattempt any case that was not successfully approved when re run so that we can make small changes and still auto approve


'===========================================================================================================================
