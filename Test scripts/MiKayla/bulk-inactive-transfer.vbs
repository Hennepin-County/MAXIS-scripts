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

worker = ""
new_worker = "x127CCL"
MAXIS_case_number = "2335052"
excluded_array = array("X127CCL", "P927079X", "P927091X", "P927152X", "P927161X", "P927252X", "PW35DI01", "PWAT072", "PWAT075", "PWAT231", "PWAT352", "PWPCT01", "PWPCT02", "PWPCT03", "PWTST40", "PWTST41", "PWTST49", "PWTST58", "PWTST64", "PWTST92", "X127EN8", "X127EN9", "X127EP1", "X127EP2", "X127EQ6", "X127EQ7", "X127EX4", "X127EX5", "X127F3E", "X127F3J", "X127F3N", "X127F4A", "X127F4B", "X127FE2", "X127FE3", "X127FF1", "X127FF2", "X127FG5", "X127FG9", "X127FH3", "X127FI1", "X127FI3", "X127FI6", "X127EJ6", "X127FE5", "X127EK3", "X127EK1", "X127EK2", "X127EJ7", "X127EJ8", "X127EJ5", "X127EH6", "X127EM1", "X127FE1", "X127FI7", "X127FH3","X127F3E", "X127F3J", "X127F3N", "X127FI6", "X127EK9", "X127FH5", "X127EK5", "X127EN7", "X127EK6", "X127EK4", "X127EN6", "X127EL1", "X127ER6", "X127EP8", "X127EQ3", "X127FG9", "X127FI3", "X127EM7", "X127FI2", "X127FG3", "X127EM8", "X127EM9", "X127EJ4", "X127EH1", "X127EH7", "X127EH2", "X127EH3", "X127FH4", "X127FI1", "X127EP3", "X127EP4", "X127EP5", "X127EP9", "X127F3P", "X127F3K", "X127F3F", "X127EM3",  "X127EM4", "X127FG6",  "X127FG7", "X127LE1", "X127NP0", "X127NPC", "X127FF4", "X127FF5")

Call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
FOR EACH worker in excluded_array	'This will remove any counted month that was actually a banked month'
	Filter_array = Filter(worker_array, worker, FALSE)
	worker_array = Filter_array
NEXT
'navigating to start the transfer'
CALL navigate_to_MAXIS_screen ("SPEC", "XFER")
EMWriteScreen "2335052", 18, 43 'MAXIS_case_number'
TRANSMIT
'Dummy case number initially
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
        	IF last_page_check = "THIS" or last_page_check = "MAXI" THEN 'MAXIMUM PAGES ALLOWED IS 100 or 'THIS IS LAST PAGE- PF3 FOR THE NEXT CASE STATUS OR ENTER START NAME
    			case_found = FALSE
    			EXIT DO
    		END IF
        LOOP UNTIL case_found = FALSE
        	'IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "What does it say?"
        'LOOP UNTIL err_msg = ""
        PF3 'to save
    Next

script_end_procedure("Success!")
