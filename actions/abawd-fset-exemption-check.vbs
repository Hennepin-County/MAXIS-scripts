'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - ABAWD FSET EXEMPTION CHECK.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 98                	'manual run time in seconds
STATS_denomination = "M"       		'M is for each MEMBER
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
call changelog_update("08/19/2019", "Updated script so that if started from the ABAWD Tracking Record pop-up on WREG, the script will read where the cursor is placed in the tracking record and if placed on a specific month, the script will autofill that footer month.", "Casey Love, Hennepin County")
call changelog_update("05/07/2018", "Updated universal ABWAWD function.", "Ilse Ferris, Hennepin County")
call changelog_update("04/25/2018", "Updated SCHL exemption coding.", "Ilse Ferris, Hennepin County")
call changelog_update("04/16/2018", "Updated output of potential exemptions for readability.", "Ilse Ferris, Hennepin County")
call changelog_update("04/10/2018", "Enhanced to check cases coded for homelessness for the 'Unfit for Employment' expansion. Also removed code that checked for SSI applying/appealing as this is no longer an exemption reason.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'The script----------------------------------------------------------------------------------------------------
'Connecting to MAXIS, and grabbing the case number and current footer month/year
EMConnect ""

Function ABAWD_FSET_exemption_finder_test()
'excluding matching grant and participating in CD treatment due to non-MAXIS indicators.
'excluding armed forces participation dur to non-MAXIS indicators. 
'----------------------------------------------------------------------------------------------------Determining the EATS Household

    Dim eats_group_array()
    ReDim eats_group_array(memb_verified_abawd_const,0)

    'constants for array
    const memb_name_const = 0
    const memb_number_const = 1
    const verified_exemption_const = 2
    const memb_potential_exempt_const = 3
    const memb_verified_wreg_const = 4
    const memb_verified_abawd_const = 5
    

    entry_record = 0
    case_based_exemptions = ""
    eats_HH_count = 0

    CALL navigate_to_MAXIS_screen("STAT", "EATS")
    eats_group_members = ""
    memb_found = True
    EMReadScreen all_eat_together, 1, 4, 72

    IF all_eat_together = "_" THEN
        eats_group_members = "01" & "," 'single member HH's
		eats_HH_count = 1
    ELSEIF all_eat_together = "Y" THEN
    'HH's where all members eat together
        eats_row = 5
        DO
            EMReadScreen eats_pers, 2, eats_row, 3
            eats_pers = replace(eats_pers, " ", "")
            IF eats_pers <> "" THEN
                eats_group_members = eats_group_members & eats_pers & ","
				eats_HH_count = eats_HH_count  + 1
                eats_row = eats_row + 1
            END IF
        LOOP UNTIL eats_pers = ""
    ELSEIF all_eat_together = "N" THEN
    'multiple eats HH cases - we are only caring about the 1st eats group that contains MEMB 01.
        eats_row = 13
        DO
            EMReadScreen eats_group, 38, eats_row, 39
            find_memb01 = InStr(eats_group, eats_pers)
            IF find_memb01 = 0 THEN
                eats_row = eats_row + 1
                IF eats_row = 18 THEN
                    memb_found = False
                    EXIT DO
                END IF
            END IF
        LOOP UNTIL find_memb01 <> 0

        'Gathering the eats group members
        eats_col = 39
        DO
            EMReadScreen eats_group, 2, eats_row, eats_col
            IF eats_group <> "__" THEN
                eats_group_members = eats_group_members & eats_group & ","
                eats_col = eats_col + 4
				eats_HH_count = eats_HH_count  + 1
            END IF
        LOOP UNTIL eats_group = "__"
    END IF

    eats_group_members = trim(eats_group_members)
    eats_group_members = split(eats_group_members, ",")

    For each memb in eats_group_members    
    	ReDim Preserve eats_group_array(memb_verified_abawd_const, entry_record)	'This resizes the array based on the number of rows in the Excel File'
    	eats_group_array(memb_number_const, entry_record) = memb
    	entry_record = entry_record + 1			'This increments to the next entry in the array'
    	stats_counter = stats_counter + 1
    Next 

    msgbox entry_entry
    'For item = 0 to UBound(eats_group_array, 2)