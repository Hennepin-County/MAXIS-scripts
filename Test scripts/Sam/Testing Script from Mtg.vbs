'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - DISASTER FOOD REPLACEMENT.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 150                	'manual run time in seconds - INCLUDES A POLICY LOOKUP
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
CALL changelog_update("10/01/2025", "Update language to reflect script use for food destroyed in disasters AND misfortunes.", "Mark Riegel, Hennepin County") '#2434
CALL changelog_update("09/10/2024", "Update to align with 06/2024 POLI/TEMP regarding denial of replacement requests.", "Mark Riegel, Hennepin County") '#1848
CALL changelog_update("05/07/2024", "Update to align with updated 02/2024 POLI/TEMP.", "Mark Riegel, Hennepin County") '#1796
CALL changelog_update("09/19/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("11/27/2019", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Connecting to BlueZone
EMConnect ""

'Gather case details as applicable
get_county_code
Call check_for_MAXIS(False)
CALL MAXIS_case_number_finder(MAXIS_case_number)


msgbox "Welcome to Sam's practice script!"

script_end_procedure_with_error_report("Ta da!! Yay script!")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------05/08/2024
'--Tab orders reviewed & confirmed----------------------------------------------09/10/2024
'--Mandatory fields all present & Reviewed--------------------------------------09/10/2024
'--All variables in dialog match mandatory fields-------------------------------05/08/2024
'Review dialog names for content and content fit in dialog----------------------05/08/2024
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------09/10/2024
'--CASE:NOTE Header doesn't look funky------------------------------------------09/10/2024
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------09/10/2024
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------05/08/2024
'--MAXIS_background_check reviewed (if applicable)------------------------------05/08/2024
'--PRIV Case handling reviewed -------------------------------------------------05/08/2024
'--Out-of-County handling reviewed----------------------------------------------05/08/2024
'--script_end_procedures (w/ or w/o error messaging)----------------------------05/08/2024
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------05/08/2024
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------05/08/2024
'--Incrementors reviewed (if necessary)-----------------------------------------05/08/2024
'--Denomination reviewed -------------------------------------------------------05/08/2024
'--Script name reviewed---------------------------------------------------------05/08/2024
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------09/10/2024
'--comment Code-----------------------------------------------------------------09/10/2024
'--Update Changelog for release/update------------------------------------------09/10/2024
'--Remove testing message boxes-------------------------------------------------09/10/2024
'--Remove testing code/unnecessary code-----------------------------------------09/10/2024
'--Review/update SharePoint instructions----------------------------------------09/10/2024
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------05/08/2024
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------05/08/2024
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------05/08/2024
'--Complete misc. documentation (if applicable)---------------------------------05/08/2024
'--Update project team/issue contact (if applicable)----------------------------05/08/2024