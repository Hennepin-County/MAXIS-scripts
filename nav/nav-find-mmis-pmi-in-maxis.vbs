'Required for statistical purposes===============================================================================
name_of_script = "NAV - FIND MMIS PMI IN MAXIS.vbs"
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
call changelog_update("12/10/2024", "Improved error handling and functionality.", "Mark Riegel, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

EMConnect ""

'Read the PMI number depending on which MMIS screen the user is on
EmReadScreen MMIS_panel_code, 79, 1, 2
If InStr(MMIS_panel_code, "-RBEN") OR _
   instr(MMIS_panel_code, "-RBYD") OR _
   instr(MMIS_panel_code, "-RTCP") OR _
   instr(MMIS_panel_code, "-RELG") OR _
   instr(MMIS_panel_code, "-RUNE") OR _
   instr(MMIS_panel_code, "-RBUY") OR _
   instr(MMIS_panel_code, "-RCAP") OR _
   instr(MMIS_panel_code, "-RCIP") OR _
   instr(MMIS_panel_code, "-REMP") OR _
   instr(MMIS_panel_code, "-RSLF") OR _
   instr(MMIS_panel_code, "-RBYB") OR _
   instr(MMIS_panel_code, "-RCPC") OR _
   instr(MMIS_panel_code, "-RHCI") OR _
   instr(MMIS_panel_code, "-RJOB") OR _
   instr(MMIS_panel_code, "-RSUM") OR _
   instr(MMIS_panel_code, "-RFED") Then
	EMReadScreen PMI_number, 8, 2, 2
Else
  script_end_procedure("A PMI number could not be found on this screen. Please navigate to a screen that provides a PMI number.")
End if

'Now it checks to make sure MAXIS production (or training) is running on this screen. If both are running the script will stop.
attn
EMReadScreen training_check, 7, 8, 15
EMReadScreen production_check, 7, 6, 15
If training_check = "RUNNING" and production_check = "RUNNING" then script_end_procedure("You have production and training both running. Close one before proceeding.")
If training_check <> "RUNNING" and production_check <> "RUNNING" then script_end_procedure("You need to run this script on the window that has MAXIS production on it. Please try again.")
If training_check = "RUNNING" then call write_value_and_transmit("S", 8, 2)
If production_check = "RUNNING" then call write_value_and_transmit("S", 6, 2)

'Ensure MAXIS was opened properly
check_for_MAXIS(False)

'Navigates to SELF
back_to_SELF

'Navigates to person search
Call navigate_to_MAXIS_screen("PERS", "    ")
EmWriteScreen PMI_number, 15, 36
transmit
EMReadScreen MTCH_check, 4, 2, 51
If MTCH_check <> "MTCH" then script_end_procedure("Unable to navigate to MTCH panel. Script will now end.")
EMWriteScreen "X", 8, 5
transmit
Do
  row = 1
  col = 1
  EMSearch "  Y    ", row, col
  If row = 0 then
    PF8
  end if
  EMReadScreen page_check, 21, 24, 2
  If page_check = "THIS IS THE ONLY PAGE" or page_check = "THIS IS THE LAST PAGE" then script_end_procedure("A case could not be found for this PMI. They could be a spouse or other member on an existing case.")
Loop until row <> 0
EMWriteScreen "X", row, 4
transmit

Call MAXIS_case_number_finder(MAXIS_case_number)

Call navigate_to_MAXIS_screen("CASE", "NOTE")

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
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------12/11/2024
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------12/11/2024
'--Incrementors reviewed (if necessary)-----------------------------------------12/11/2024
'--Denomination reviewed -------------------------------------------------------12/11/2024
'--Script name reviewed---------------------------------------------------------12/11/2024
'--BULK - remove 1 incrementor at end of script reviewed------------------------12/11/2024

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
