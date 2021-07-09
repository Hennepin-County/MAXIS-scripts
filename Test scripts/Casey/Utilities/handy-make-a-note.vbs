'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - SELECT WCOM.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 120                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
'END OF stats block==============================================================================================
run_locally = TRUE
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

Call MAXIS_case_number_finder(MAXIS_case_number)

Call start_a_blank_CASE_NOTE

Call write_variable_in_CASE_NOTE("*** SNAP Approved starting in ")
Call write_variable_in_CASE_NOTE("* SNAP approved for 01/19")
Call write_variable_in_CASE_NOTE("    Eligible Household Members: 01,")
Call write_variable_in_CASE_NOTE("    Income: Earned: $166.00 Unearned: $0.00")
Call write_variable_in_CASE_NOTE("    Shelter Costs: $0.00")
Call write_variable_in_CASE_NOTE("    SNAP BENEFTIT: $192.00 Reporting Status: SIX MONTH")
Call write_variable_in_CASE_NOTE("* SNAP approved for 02/19")
Call write_variable_in_CASE_NOTE("    Eligible Household Members: 01,")
Call write_variable_in_CASE_NOTE("    Income: Earned: $166.00 Unearned: $0.00")
Call write_variable_in_CASE_NOTE("    Shelter Costs: $0.00")
Call write_variable_in_CASE_NOTE("    SNAP BENEFTIT: $192.00 Reporting Status: SIX MONTH")
Call write_variable_in_CASE_NOTE("THIS IS NEW INFORMATION")
Call write_variable_with_indent_in_CASE_NOTE("using a new function")
Call write_bullet_and_variable_in_CASE_NOTE("Notes", "01/19 for 01 is BANKED MONTH - Banked Month: 6.; 02/19 for 01 is BANKED MONTH - Banked Month: 7.")
Call write_variable_in_CASE_NOTE("C. Love ~ BlueZone Scripts Project ~ EWS QI")
