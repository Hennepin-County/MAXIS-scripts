'STATS GATHERING=============================================================================================================
name_of_script = "DAIL - SDX MATCH.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 1            'manual run time in seconds  -----REPLACE STATS_MANUALTIME = 1 with the anctual manualtime based on time study
STATS_denomination = "C"        'C is for each case; I is for Instance, M is for member; REPLACE with the denomonation appliable to your script.
'END OF stats block==========================================================================================================

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

'CHANGELOG BLOCK===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County
CALL changelog_update("10/14/25", "Initial version.", "Mark Riegel, Hennepin County") 'REPLACE with release date and your name.

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT==================================================================================================================
EMConnect "" 'Connects to BlueZone
'Read the dail message
EMReadScreen full_message, 60, 6, 20
full_message = trim(full_message)

Dialog1 = "" 'blanking out dialog name

BeginDialog Dialog1, 0, 0, 311, 155, "DAIL - 1619 Status"
  ButtonGroup ButtonPressed
    PushButton 5, 45, 65, 15, "ONEsource", onesource_manual_btn
    PushButton 5, 65, 65, 15, "TE 02.07.259", poli_temp_btn
    PushButton 5, 85, 65, 15, "HSR Manual", hsr_manual_btn
    PushButton 5, 105, 65, 15, "Script Instructions", script_instructions_btn
    OkButton 205, 135, 50, 15
    CancelButton 255, 135, 50, 15
  Text 5, 5, 55, 10, "DAIL Message - "
  Text 60, 5, 245, 10, full_message
  Text 5, 20, 300, 20, "This DAIL message is not currently supported by scripts. Please see the following policies/ procedures for information on how to process:"
  Text 75, 50, 95, 10, "Link to ONEsource"
  Text 75, 70, 75, 10, "Link to POLI/TEMP"
  Text 75, 90, 85, 10, "Link to HSR Manual"
  Text 75, 110, 85, 10, "Link to Script Instructions"
EndDialog


DO
    Do
        err_msg = ""    'This is the error message handling
        Dialog Dialog1
        cancel_without_confirmation
		If ButtonPressed = onesource_manual_btn Then
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=ONESOURCE-170311"
			err_msg = "LOOP"
		End If
		If ButtonPressed = poli_temp_btn Then
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:b:/r/sites/hs-es-poli-temp/Documents%203/TE%2002.07.259%201619%20A%20AND%20B%20STATUS.pdf"
			err_msg = "LOOP"
		End If
		If ButtonPressed = hsr_manual_btn Then 
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/INFO.aspx#prsn-01-1619-status-updated-on-disa-check-ma-elig"
			err_msg = "LOOP"
		End If
		If ButtonPressed = script_instructions_btn Then 
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/DAIL/DAIL%20-%201619%20Status.docx"
			err_msg = "LOOP"
		End If
    Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
'End dialog section-----------------------------------------------------------------------------------------------

'End the script.
script_end_procedure("Please follow the instructions provided in ONEsource, POLI/TEMP, and/or the HSR Manual. The script will now end.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 05/23/2024
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------10/14/2025
'--Tab orders reviewed & confirmed----------------------------------------------10/14/2025
'--Mandatory fields all present & Reviewed--------------------------------------10/14/2025
'--All variables in dialog match mandatory fields-------------------------------10/14/2025
'Review dialog names for content and content fit in dialog----------------------10/14/2025
'--FIRST DIALOG--NEW EFF 5/23/2024----------------------------------------------10/14/2025
'--Include script category and name somewhere on first dialog-------------------10/14/2025
'--Create a button to reference instructions------------------------------------10/14/2025
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------10/14/2025
'--MAXIS_background_check reviewed (if applicable)------------------------------10/14/2025
'--PRIV Case handling reviewed -------------------------------------------------10/14/2025
'--Out-of-County handling reviewed----------------------------------------------10/14/2025
'--script_end_procedures (w/ or w/o error messaging)----------------------------10/14/2025
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------10/14/2025
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------10/14/2025
'--Incrementors reviewed (if necessary)-----------------------------------------10/14/2025
'--Denomination reviewed -------------------------------------------------------10/14/2025
'--Script name reviewed---------------------------------------------------------10/14/2025
'--BULK - remove 1 incrementor at end of script reviewed------------------------10/14/2025

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------10/14/2025
'--comment Code-----------------------------------------------------------------10/14/2025
'--Update Changelog for release/update------------------------------------------10/14/2025
'--Remove testing message boxes-------------------------------------------------10/14/2025
'--Remove testing code/unnecessary code-----------------------------------------10/14/2025
'--Review/update SharePoint instructions----------------------------------------10/14/2025
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------10/14/2025
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------10/14/2025
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------10/14/2025
'--Complete misc. documentation (if applicable)---------------------------------10/14/2025
'--Update project team/issue contact (if applicable)----------------------------10/14/2025