'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - Approve Banked Months.vbs"
start_time = timer
STATS_counter = 0			 'sets the stats counter at one
STATS_manualtime = 0			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
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
call changelog_update("07/11/2018", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


'NEED A SCRIPT THAT WILL OPERATE OFF OF THE DAIL - PEPR (this is from a list generated with BULK-Dail)
    'FS ABAWD ELIGIBILITY HAS EXPIRED, APPROVE NEW ELIG RESULTS
    'Review case for possible ABAWD exemptions, 2nd set, then finally Banked Months
        'if banked - add to working list
    'Gathers MEMB number and which months are banked
    'Updates ABAWD tracking record and WREG
    'approves case
    'CASE NOTES

'NEED A SCRIPT TO ASSES AND UPDATE A WORKING EXCEL
    'There is a BOBI list of all clients on SNAP
    'It should compare to a working list and for any not on the working list
    'For any not on the list, asses for potential banked months cases

'NEED A SCRIPT TO REVIEW ALL THE CASES on the BANKED MONTHS LIST'

'THOUGHTS
'Use case notes instead of person notes to document used Banked Months as that way we are using a form people are more comfortable with.

'CONSTANTS=================================================================================================================

'THE COLUMNS IN THE WORKING EXCEL
Const case_nbr_col      = 1
Const memb_nrb_col      = 2
Const last_name_col     = 3
Const first_name_col    = 4
Const notes_col         = 5
Const first_mo_col      = 6
Const scnd_mo_col       = 7
Const third_mo_col      = 8
Const fourth_mo_col     = 9
Const fifth_mo_col      = 10
Const sixth_mo_col      = 11
Const svnth_mo_col      = 12
Const eighth_mo_col     = 13
Const ninth_mo_col      = 14
Const curr_mo_stat_col  = 15


'==========================================================================================================================


'THE SCRIPT================================================================================================================

'Connects to BlueZone
EMConnect ""

'Initial Dialog will have worker select which option is going to be run at this time
    'Assess Banked Month cases from DAIL PEPR List
    'Review monthly BOBI report of all SNAP clients
    'Review of Banked Months cases
    'Approve ongoing Banked Month Cases
    'HAVE DEVELOPER MODE

'IF NOT in Developer Mode, check to be sure we are in production

'DAIL PEPR Option
    'Dialog to select the Excel list that has the DAILs
    'add all to an array
    'Compare the array to the Working list
        'add to working list if not already there

'BOBI Report Option
    'Check each person on the BOBI list in MAXIS
        'exclude clients with obvious exclusions (?? age)
        'should we actually check MAXIS to see if it is coded correctly?
    'add each to the array
    'compare the array to the working list
        'if not already on the list, check WREG for 30/13
        

'Review of cases
    'Open the working Excel sheet
    'Have worker confirm the correct sheet opened
    'Read all the cases from the spreadsheet and add to an array

    'Check CASE CURR to see if case and person are still active SNAP
        'If closed need to review if the closure was correct
    'Check WREG
        'Confirm case is coded as 30/13
        'Confirm ABAWD months have been used
    'GET Code from UTILITIES - COUNTED ABAWD MONTHS - need to confirm that counted months are correct
    'GET Code from ACTIONS - ABAWD FSET EXEMPTION CHECK and run it on every SNAP month to check the counting
    'Update MAXIS panels/WREG/ABAWD Tracking Record as determined by other runs
        'may need to do person search to see if there was SNAP on another case that caused the counted month
        'If any month is confusing then use code from NOTES - ABAWD TRACKING RECORD to coordinate
        'MAY need dialog for worker to confirm confusing months
    'Need to check ECF - create a dialog to allow worker to review ECF information
    '????'

'Approve ongoing cases
    'Open the working Excel sheet
    'Have worker confirm the correct sheet opened
    'Read all the cases from the spreadsheet and add to an array

    'Read PROG and ELIG to confirm client is still active SNAP on this case
    'Check for possible exemption in STAT
    'Review Case Notes to see if there are any case notes that need to be assessed
        'Have a series of case notes that can be ignored
        'Look just to the last BM case note
        'Have a dialog for the worker to review the case notes if anything appears indicating a change may have happened
        'Worker can confirm that the BM coding is correct or adjust in the dialog
    'Go to WREG
    'Check tracker to see if any ABAWD months have fallen off of the 36 month look back period
    'Update WREG with any information found
        'If exempt - update exemption coding
        'If still BM ensure coding is 30/13 and update the BM counter
    'Review case and update other STAT panels if eneeded (JOBS dates)
    'Review ELIG and approve
    'Update Excel
'
