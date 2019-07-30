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
call changelog_update("07/30/2019", "Added X127EW7 & X127EW8 per request from Christine Glisczinski. MiKayla Handley")
call changelog_update("02/14/2019", "Initial version.", "MiKayla Handley")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'-------------------------------------------------------------------------------------DIALOGS
BeginDialog pull_REPT_data, 0, 0, 301, 105, "Confirm the footer month for INAC transfer"
  CheckBox 85, 5, 150, 10, "Check here to run this query county-wide.", all_workers_check
  EditBox 55, 15, 20, 15, inac_month
  EditBox 55, 35, 20, 15, inac_year
  EditBox 155, 20, 140, 15, worker_number
  ButtonGroup ButtonPressed
    OkButton 185, 85, 50, 15
    CancelButton 240, 85, 50, 15
  Text 5, 70, 290, 10, "NOTE: running queries county-wide can take a significant amount of time and resources."
  Text 85, 40, 210, 20, "Enter all 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
  GroupBox 5, 5, 75, 60, "Month to scan"
  Text 85, 25, 65, 10, "Worker(s) to check:"
  Text 10, 20, 40, 10, "Month (MM):"
  Text 10, 40, 35, 10, "Year (YY):"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------
'inserting month - 1 into footer month section as this is likely the most commonly needed inac month.
inac_month = datepart("m", dateadd("m", -1, date))
inac_year = right(dateadd("m", -1, date), 2)
If len(inac_month) = 1 then inac_month = "0" & inac_month

'Connects to BlueZone
EMConnect ""

all_workers_check = CHECKED
'Shows dialog
Dialog pull_REPT_data

cancel_confirmation

If len(inac_month) = 1 then inac_month = "0" & inac_month

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Checking for MAXIS
Call check_for_MAXIS(false)

EMWriteScreen inac_month, 20, 54
EMWriteScreen inac_year, 20, 57
TRANSMIT
'msgbox inac_month

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = CHECKED then
	worker = ""
	new_worker = "X127CCL"
	MAXIS_case_number = "2335052"
	excluded_array = array("P927079X", "P927091X", "P927152X", "P927161X", "P927252X", "PW35DI01", "PWAT072", "PWAT075", "PWAT231", "PWAT352", "PWPCT01", "PWPCT02", "PWPCT03", "PWTST40", "PWTST41", "PWTST49", "PWTST58", "PWTST64", "PWTST92", "X1274EC ", "X127966", "X127AN1 ", "X127AP7", "X127CCA", "X127CCL", "X127CCR", "X127CSS", "X127EF8", "X127EF9", "X127EM3", "X127EM4", "X127EN8", "X127EN9", "X127EP1", "X127EP2","X127EQ6", "X127EQ7", "X127EW4", "X127EW6 ","X127EX4", "X127EX5", "X127F3E", "X127F3F", "X127F3J", "X127F3K", "X127F3N", "X127F3P", "X127F4A", "X127F4B", "X127FB1 ", "X127FE2", "X127FE3", "X127FF1", "X127FF2", "X127FF4", "X127FF5", "X127FF6 ", "X127FF9 ", "X127FG1", "X127FG2", "X127FG5", "X127FG6", "X127FG7", "X127FG9", "X127FH3", "X127FI1", "X127FI3", "X127FI6", "X127GF5", "X127LE1", "X127NP0", "X127NPC", "X127NPC ", "X127Q95", "X127EW7", "X127EW8")

	Call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
	FOR EACH worker in excluded_array
		Filter_array = Filter(worker_array, worker, FALSE)
		worker_array = Filter_array
	NEXT
Else		'If worker numbers are litsted - this will create an array of workers to check
	x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas

	'formatting array
	For each x1_number in x1s_from_dialog
		x1_number = trim(ucase(x1_number))					'Formatting the x numbers so there are no errors
		Call navigate_to_MAXIS_screen ("REPT", "USER")		'This part will check to see if the x number entered is a supervisor of anyone
		PF5
		PF5
		EMWriteScreen x1_number, 21, 12
		transmit
		EMReadScreen sup_id_check, 7, 7, 5					'This is the spot where the first person is listed under this supervisor
		IF sup_id_check <> "       " Then 					'If this frist one is not blank then this person is a supervisor
			supervisor_array = trim(supervisor_array & " " & x1_number)		'The script will add this x number to a list of supervisors
		Else
			If worker_array = "" then						'Otherwise this x number is added to a list of workers to run the script on
				worker_array = trim(x1_number)
			Else
				worker_array = worker_array & ", " & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
			End if
		End If
		PF3
	Next

	If supervisor_array <> "" Then 				'If there are any x numbers identified as a supervisor, the script will run the function above
		Call create_array_of_all_active_x_numbers_by_supervisor (more_workers_array, supervisor_array)
		workers_to_add = join(more_workers_array, ", ")
		If worker_array = "" then				'Adding all x numbers listed under the supervisor to the worker array
			worker_array = workers_to_add
		Else
			worker_array = worker_array & ", " & trim(ucase(workers_to_add))
		End if
	End If

	'Split worker_array
	worker_array = split(worker_array, ", ")
End if

'navigating to start the transfer this goes to INACTIVE cases only!
CALL navigate_to_MAXIS_screen ("SPEC", "XFER")
EMWriteScreen inac_month, 20, 54
EMWriteScreen inac_year, 20, 57
'Dummy case number initially
EMWriteScreen "2335052", 18, 43 'MAXIS_case_number'
TRANSMIT
EMWriteScreen "X", 11, 16 'Transfer case load same county
TRANSMIT
'-----------------------------------------------XCLD
    For each worker in worker_array
        'msgbox "where am I now"
        EMWriteScreen worker, 04, 18
        TRANSMIT
        'IF current_worker = "" then pf3
        EMWriteScreen "X127CCL", 15, 13 'new_worker'
        TRANSMIT
        EMWriteScreen "x", 11, 13 'inactive_case
        'Change the footer month
        TRANSMIT 'REVIEW NEW WORKER NAME AND PRESS ENTER TO VIEW DETAILS
        TRANSMIT
        '-------------------------------------------XFER
        row = 7
        DO
        	DO
        		EMReadScreen previous_number, 7, row, 5
        		IF previous_number <> "" THEN
        			case_found = TRUE
        			EMWriteScreen "1", row, 3
        			row = row + 1
					STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
        		END IF
        	Loop until row = 19
        	row = 7 'Setting the variable for when the do...loop restarts
        	PF8
        	IF previous_number = "" THEN case_found = FALSE
        	EMReadScreen last_page_check, 4, 24, 2 'checks for "THIS IS THE LAST PAGE"
        	IF last_page_check = "THIS" or last_page_check = "MAXI" THEN 'MAXIMUM PAGES ALLOWED IS 100 or 'THIS IS LAST PAGE- PF3 FOR THE NEXT CASE STATUS OR ENTER START NAME
    			case_found = FALSE
    			EXIT DO
    		END IF
        LOOP UNTIL case_found = FALSE
        	'IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "What does it say?"
        'LOOP UNTIL err_msg = ""
        PF3 'to save
    Next
	'Logging usage stats
	STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
MsgBox STATS_COUNTER
script_end_procedure("Success!")
