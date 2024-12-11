'Required for statistical purposes===============================================================================
name_of_script = "NAV - FIND MAXIS CASE IN MMIS.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 40                      'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
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
call changelog_update("12/11/2024", "Updated script to navigate from MAXIS training to MMIS training region. ", "Mark Riegel, Hennepin County")
call changelog_update("01/31/2019", "Streamlined script to navigate to MMIS more effectively.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""

'First checks to make sure you're in MAXIS.
check_for_MAXIS(False)

'Reading the case number, then removing spaces and underscores, and adding the leading zeroes for MMIS.
Call MAXIS_case_number_finder(maxis_case_number)
MAXIS_case_number = right("00000000" & MAXIS_case_number, 8)

'Checking to see if we are on the HC/APP screen, which is not supported at this time (case number is in different place)
EMReadScreen HC_app_check, 16, 3, 33
If HC_app_check = "Approval Package" then script_end_procedure("The script needs to be on the previous or next screen to process this.")

'Navigate to MMIS 
Call navigate_to_MMIS_region("CTY ELIG STAFF/UPDATE")	'function to navigate into MMIS, select the HC realm, and enters the prior autorization area

'Now we are in RKEY, and it navigates into the case, transmits, and makes sure we've moved to the next screen.
EMWriteScreen "I", 2, 19    'enter into case in MMIS in INQUIRY mode 
Call write_value_and_transmit(MAXIS_case_number, 9, 19)

EMReadscreen RKEY_check, 4, 1, 52
If RKEY_check = "RKEY" then script_end_procedure("A correct case number was not taken from MAXIS. Check your case number and try again.")

'Now it gets to RELG for member 01 of this case.
Call write_value_and_transmit("RCIN", 1, 8)
EMWriteScreen "X", 11, 2                            'selecting MEMB 01 on the case 
Call write_value_and_transmit("RELG", 1, 8)

script_end_procedure("Success!")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 05/23/2024
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------N/A
'--Tab orders reviewed & confirmed----------------------------------------------N/A
'--Mandatory fields all present & Reviewed--------------------------------------N/A
'--All variables in dialog match mandatory fields-------------------------------N/A
'Review dialog names for content and content fit in dialog----------------------N/A
'--FIRST DIALOG--NEW EFF 5/23/2024----------------------------------------------N/A
'--Include script category and name somewhere on first dialog-------------------N/A
'--Create a button to reference instructions------------------------------------N/A
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'--write_variable_in_CASE_NOTE function: confirm proper punctuation is used-----N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------12/11/2024
'--MAXIS_background_check reviewed (if applicable)------------------------------12/11/2024
'--PRIV Case handling reviewed -------------------------------------------------12/11/2024
'--Out-of-County handling reviewed----------------------------------------------12/11/2024
'--script_end_procedures (w/ or w/o error messaging)----------------------------12/11/2024
'--BULK - review output of statistics and run time/count (if applicable)--------12/11/2024
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------12/11/2024
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------12/11/2024
'--Incrementors reviewed (if necessary)-----------------------------------------12/11/2024
'--Denomination reviewed -------------------------------------------------------12/11/2024
'--Script name reviewed---------------------------------------------------------12/11/2024
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------12/11/2024
'--comment Code-----------------------------------------------------------------12/11/2024
'--Update Changelog for release/update------------------------------------------12/11/2024
'--Remove testing message boxes-------------------------------------------------12/11/2024
'--Remove testing code/unnecessary code-----------------------------------------12/11/2024
'--Review/update SharePoint instructions----------------------------------------12/11/2024
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------12/11/2024
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------12/11/2024
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------12/11/2024
'--Complete misc. documentation (if applicable)---------------------------------12/11/2024
'--Update project team/issue contact (if applicable)----------------------------12/11/2024
