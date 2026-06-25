'STATS GATHERING=============================================================================================================
name_of_script = "TYPE - NEW SCRIPT TEMPLATE.vbs"       'REPLACE TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
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
CALL changelog_update("01/01/01", "Initial version.", "Ilse Ferris, Hennepin County") 'REPLACE with release date and your name.



'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Write Screen Function
function write_value(input_value, row, col)
'--- This function writes a specific value and transmits.
'~~~~~ input_value: information to be entered
'~~~~~ row: row to write the input_value
'~~~~~ col: column to write the input_value
'===== Keywords: MAXIS, PRISM, case note, three columns, format
	EMWriteScreen input_value, row, col
end function





'THE SCRIPT==================================================================================================================
EMConnect "" 'Connects to BlueZone

CALL MAXIS_case_number_finder(MAXIS_case_number)    				'Grabs the MAXIS case number automatically
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)		'Grabs the footer month and year from MAXIS
'  -- OR --
MAXIS_footer_month = CM_plus_1_mo									'Directly assigns a footer month based on the current month
MAXIS_footer_year = CM_plus_1_yr

Dialog1 = "" 'blanking out dialog name
'Add dialog here: Add the dialog just before calling the dialog below unless you need it in the dialog due to using COMBO Boxes or other looping reasons. Blank out the dialog name with Dialog1 = "" before adding dialog.
'    Some Dialog Elements:  Initial Dialog Header: 			BeginDialog Dialog1, 0, 0, 191, 105, "CATEGORY - NAME Case Number Dialog"  				-- Use CATEGORY - NAME somewhere in the header
'							Script Instructions Button:		PushButton 135, 5, 50, 15, "Instructions", script_instructions_btn						-- Have a button to open the instructions
'							Script Purpose/Overview: 		Text 10, 70, 120, 30, "Here is a quick summary of the purpose of the script."			-- Give the worker a little guidance
'							Include edit boxes for necessary details like Case Number, Footer Month, Footer Year, and Worker Signature
'Shows dialog (replace "sample_dialog" with the actual dialog you entered above)----------------------------------


BeginDialog Dialog1, 0, 0, 301, 240, "Dialog"
  ButtonGroup ButtonPressed
    OkButton 185, 215, 50, 15
    CancelButton 240, 215, 50, 15
  EditBox 5, 210, 75, 15, MAXIS_case_number
  EditBox 90, 210, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    PushButton 15, 15, 75, 15, "Person Search", person_search
EndDialog


DO
	DO
	    err_msg = ""
	    Dialog Dialog1
	    cancel_confirmation
			if buttonpressed = person_search Then
				Call navigate_to_MAXIS_screen("PERS", "") 
			end if


	    IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "You must enter a worker signature."
	    IF err_msg <> "" THEN msgbox "*** Notice!!! ***" & vbNewLine & err_msg
 	LOOP UNTIL err_msg = ""
	CALL check_for_password_without_transmit(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = false

Dialog1 = ""

BeginDialog Dialog1, 0, 0, 191, 105, "Dialog"
  EditBox 15, 60, 90, 15, first_name
  Text 15, 10, 50, 10, "Last Name"

  EditBox 15, 20, 90, 15, last_name
  Text 15, 45, 50, 10, "First Name"  
  ButtonGroup ButtonPressed
	PushButton 120, 20, 50, 15, "Search", search_button
EndDialog



DO
	DO
	    err_msg = ""
	    Dialog Dialog1
	    cancel_confirmation
			If ButtonPressed = search_button Then
				If last_name = "" Or first_name = "" Then
					err_msg = err_msg & vbNewLine & "You must enter both a last name and first name to search."
				Else
					Call write_value(last_name, 4, 36)
					Call write_value(first_name, 10, 36)
					Call transmit()
				End If
			end if

	    IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "You must enter a worker signature."
	    IF err_msg <> "" THEN msgbox "*** Notice!!! ***" & vbNewLine & err_msg
 	LOOP UNTIL err_msg = ""
	CALL check_for_password_without_transmit(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = false




script_end_procedure_with_error_report("SUCCESS")



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