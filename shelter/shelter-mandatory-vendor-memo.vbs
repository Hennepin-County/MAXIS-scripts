'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - SHELTER-MAND VENDOR MEMO.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 0         	'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("12/29/2017", "Coordinates for sending MEMO's has changed in SPEC/MEMO. Updated script to support change.", "Ilse Ferris, Hennepin County")
CALL changelog_update("12/01/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog Mand_vendor_dialog, 0, 0, 306, 195, "Mandatory Vendor SPEC/MEMO"
  EditBox 60, 5, 55, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 180, 175, 50, 15
    CancelButton 235, 175, 50, 15
  Text 10, 10, 45, 10, "Case number:"
  Text 25, 50, 250, 25, "Call 1 (888) 577-2227 to get more information or enroll. You will be placed on mandatory vendor because you have used shelter or have requested assistance for housing issues. "
  Text 25, 80, 250, 25, "You will remain on mandatory vendor for 12 months. If you move, or your rent changes you must let your team know at least 15 days before the end of the month to make this change."
  Text 25, 110, 250, 35, "Call your Human Service Representative Team at the end of this 12 month period if you want them to stop vendoring your rent at that time. Budgeting classes are free to you and available through the Lutheran Social Services. If you have any questions call the Shelter Team at (612)-348-9410."
  Text 25, 150, 80, 10, "Sincerely, Shelter Team"
  GroupBox 15, 30, 270, 135, "This is the following text that will be added to the SPEC/MEMO:"
  Text 125, 10, 170, 10, "*NOTE: MEMO will also be sent to AREP and SWKR"
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog Mand_vendor_dialog
        cancel_confirmation
		IF len(case_number) > 8 or IsNumeric(case_number) = False THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."		
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
 Call check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False
 
back_to_SELF
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen CM_mo, 20, 43	'entering current footer month/year
EMWriteScreen CM_yr, 20, 46

call navigate_to_MAXIS_screen("SPEC", "MEMO")		'Navigating to SPEC/MEMO
'Creates a new MEMO. If it's unable the script will stop.
PF5
EMReadScreen memo_display_check, 12, 2, 33
If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")

'Checking for an AREP. If there's an AREP it'll navigate to STAT/AREP, check to see if the forms go to the AREP. If they do, it'll write X's in those fields below.
row = 4                             'Defining row and col for the search feature.
col = 1
EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
IF row > 4 THEN                     'If it isn't 4, that means it was found.
	arep_row = row                                          'Logs the row it found the ALTREP string as arep_row
	call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
	EMReadscreen forms_to_arep, 1, 10, 45                   'Reads for the "Forms to AREP?" Y/N response on the panel.
	call navigate_to_MAXIS_screen("SPEC", "MEMO")           'Navigates back to SPEC/MEMO
	PF5                                                     'PF5s again to initiate the new memo process
END IF
	
'Checking for SWKR
row = 4                             'Defining row and col for the search feature.
col = 1
EMSearch "SOCWKR", row, col         'Row and col are variables which change from their above declarations if "SOCWKR" string is found.
IF row > 4 THEN                     'If it isn't 4, that means it was found.
	swkr_row = row                                          'Logs the row it found the SOCWKR string as swkr_row
	call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
	EMReadscreen forms_to_swkr, 1, 15, 63                'Reads for the "Forms to SWKR?" Y/N response on the panel.
	call navigate_to_MAXIS_screen("SPEC", "MEMO")         'Navigates back to SPEC/MEMO
	PF5                                           'PF5s again to initiate the new memo process
END IF
EMWriteScreen "x", 5, 12                                        'Initiates new memo to client
IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 12     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 12     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
transmit  
'writing the MEMO'
Call write_variable_in_SPEC_MEMO("************************************************************")
Call write_variable_in_SPEC_MEMO("Call 1 (888) 577-2227 to get more information or enroll. You will be placed on mandatory vendor because you have used shelter or have requested assistance for housing issues.")   
Call write_variable_in_SPEC_MEMO(" ")
Call write_variable_in_SPEC_MEMO("You will remain on mandatory vendor for 12 months. If you move, or your rent changes you must let your team know at least 15 days before the end of the month to make this change.")                                 
Call write_variable_in_SPEC_MEMO(" ")
Call write_variable_in_SPEC_MEMO("Call your Human Service Representative Team at the end of this 12 month period if you want them to stop vendoring your rent at that time. Budgeting classes are free to you and available through the Lutheran Social Services. If you have any questions call the Shelter Team at (612)-348-9410.")
Call write_variable_in_SPEC_MEMO(" ")
Call write_variable_in_SPEC_MEMO("Sincerely, Shelter Team")	
Call write_variable_in_SPEC_MEMO("************************************************************")
PF4

script_end_procedure("Success! Your MEMO has been sent.")	