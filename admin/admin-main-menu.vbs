'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - MAIN MENU.vbs"
start_time = timer

'=====User ID's======
'Faughn	    WFX901
'Jennifer	WFU851
'Ilse       ILFE001
'MiKayla	WFS395
'Casey	    CALO001
'Brenda	    WFI021
'Brooke	    WFU161
'Charles	WF7638
'Deb	    WFP106
'Hannah	    WFQ898
'Jessica	WFK093
'Louise	    WF1875
'Mandora	WFM207
'Melissa F.	WFG492
'Melissa M	WFP803
'Molly  	WFX490

'The following code looks to find the user name of the user running the script---------------------------------------------------------------------------------------------
'This is used in arrays that specify functionality to specific workers
Set objNet = CreateObject("WScript.NetWork")
windows_user_ID = objNet.UserName
user_ID_for_validation= ucase(windows_user_ID)

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
call changelog_update("08/06/2019", "Added a new script to create an Excel List of MAXIS User detail.", "Casey Love, Hennepin County")
call changelog_update("04/12/2019", "Updated backend fuctionality. If you are on the QI team, and cannot access the QI scripts, contact me right away.", "Ilse Ferris, Hennepin County")
call changelog_update("06/21/2018", "Added QI specific scripts and sub menu.", "Ilse Ferris, Hennepin County")
call changelog_update("03/12/2018", "Added ODW Application.", "MiKayla Handley, Hennepin County")
call changelog_update("03/12/2018", "Removed DAIL report. BULK DAIL REPORT was updated in its place.", "Ilse Ferris, Hennepin County")
call changelog_update("11/30/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'CUSTOM FUNCTIONS===========================================================================================================
Function declare_admin_menu_dialog(script_array)
    BeginDialog admin_dialog, 0, 0, 516, 320, "Admin Scripts"
    Text 5, 5, 516, 300, "Admin scripts main menu: select the script to run from the choices below."
    ButtonGroup ButtonPressed
		 	PushButton 015, 35, 30, 15, "ADMIN", 			admin_main_button
		 	If show_QI_button = True then PushButton 045, 35, 30, 15, "QI", 			    QI_button
            If show_BZ_button = TRUE Then PushButton 075, 35, 30, 15, "BZ",                 BZ_button
            If show_tasks_button = TRUE Then PushButton 105, 35, 50, 15, " Monthly Tasks ", tasks_button
		'This starts here, but it shouldn't end here :)
		vert_button_position = 70

		For current_script = 0 to ubound(script_array)
			'Displays the button and text description-----------------------------------------------------------------------------------------------------------------------------
			'FUNCTION		HORIZ. ITEM POSITION								VERT. ITEM POSITION		ITEM WIDTH									ITEM HEIGHT		ITEM TEXT/LABEL										BUTTON VARIABLE
			PushButton 		5, 													vert_button_position, 	script_array(current_script).button_size, 	10, 			script_array(current_script).script_name, 			button_placeholder
			Text 			script_array(current_script).button_size + 10, 		vert_button_position, 	500, 										10, 			"--- " & script_array(current_script).description
			'----------
			vert_button_position = vert_button_position + 15	'Needs to increment the vert_button_position by 15px (used by both the text and buttons)
			'----------
			script_array(current_script).button = button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
			button_placeholder = button_placeholder + 1
		next

		CancelButton 455, 300, 50, 15
        If show_BZ_button = TRUE Then
		    GroupBox 5, 20, 155, 35, "Admin Sub-Menus"
        Else
            GroupBox 5, 20, 100, 35, "Admin Sub-Menus"
        End If
	EndDialog
End function
'END CUSTOM FUNCTIONS=======================================================================================================

'VARIABLES TO DECLARE=======================================================================================================

'Declaring the variable names to cut down on the number of arguments that need to be passed through the function.
DIM ButtonPressed
'DIM SIR_instructions_button
Dim admin_dialog
'Dim script_array_admin_main
'Dim script_array_QI_list
'Dim script_array_BZ_list
'Dim script_array_task_list

script_array_admin_main = array()
script_array_QI_list = array()
script_array_BZ_list = array()
script_array_task_list = array()

'END VARIABLES TO DECLARE===================================================================================================

'LIST OF SCRIPTS================================================================================================================
'INSTRUCTIONS: simply add your new script below. Scripts are listed in alphabetical order. Copy a block of code from above and paste your script info in. The function does the rest.

'ADMIN scripts-------------------------------------------------------------------------------------------------------------------------
'Resetting the variable

script_num = 0
ReDim Preserve script_array_admin_main(script_num)
Set script_array_admin_main(script_num) = new script
script_array_admin_main(script_num).script_name 		= "Add GRH Rate 2 to MMIS"											'Script name
script_array_admin_main(script_num).file_name 			= "add-grh-rate-2-to-mmis.vbs"										'Script URL
script_array_admin_main(script_num).description 		= "ACTION script adds GRH Rate 2 SSR's to MMIS. This version without Rate 2 in elig results for error cases."

script_num = script_num + 1
ReDim Preserve script_array_admin_main(script_num)
Set script_array_admin_main(script_num) = new script
script_array_admin_main(script_num).script_name 		= "Application Received"											'Script name
script_array_admin_main(script_num).file_name 			= "ca-application-received.vbs"										'Script URL
script_array_admin_main(script_num).description 		= "ACTION script that will case note application received. This version was created for the Request for APPL process."

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array_admin_main(script_num)	    'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script    'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name		    = "BULK - REPT USER List"
script_array_admin_main(script_num).file_name			= "bulk-rept-user-list.vbs"
script_array_admin_main(script_num).description		    = "Report to pull MAXIS USER detail into an Excel Spreadsheet."

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array_admin_main(script_num)	    'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script    'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name		    = "Copy Case Data for Training"
script_array_admin_main(script_num).file_name			= "copy-case-data-for-training.vbs"
script_array_admin_main(script_num).description		    = "Copies data from a case to a spreadsheet to be run on the Training Case Generator."

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array_admin_main(script_num)	    'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script    'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name		    = "Copy Panels to Word"
script_array_admin_main(script_num).file_name			= "copy-panels-to-word.vbs"
script_array_admin_main(script_num).description		    = "Copies MAXIS panels to Word en masse for a case for easier review."

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array_admin_main(script_num)	    'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script    'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name		    = "Drug Felon list"
script_array_admin_main(script_num).file_name			= "drug-felon-list.vbs"
script_array_admin_main(script_num).description		    = "Reviews the Drug Felon list from DHS to update these cases."

script_num = script_num + 1
ReDim Preserve script_array_admin_main(script_num)
Set script_array_admin_main(script_num) = new script
script_array_admin_main(script_num).script_name 		= "CS Good Cause "											'Script name
script_array_admin_main(script_num).file_name 			= "cs-good cause.vbs"										'Script URL
script_array_admin_main(script_num).description 		= "Completes updates to ABPS and case notes actions taken."

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array_admin_main(script_num)	    'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script    'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name		    = "Drug Felon list"
script_array_admin_main(script_num).file_name			= "drug-felon-list.vbs"
script_array_admin_main(script_num).description		    = "Reviews the Drug Felon list from DHS to update these cases."

script_num = script_num + 1
ReDim Preserve script_array_admin_main(script_num)
Set script_array_admin_main(script_num) = new script
script_array_admin_main(script_num).script_name 		= "Earned Income Budgeting"											'Script name
script_array_admin_main(script_num).file_name 			= "earned-income-budgeting.vbs"										'Script URL
script_array_admin_main(script_num).description 		= "Assists with determination and entry of income information."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_admin_main(script_num)		'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name			= "Language Stats"													'Script name
script_array_admin_main(script_num).file_name			= "language-stats.vbs"												'Script URL
script_array_admin_main(script_num).description			= "Collects language statistics by language and region. Take approximately 10 hours to run."

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array_admin_main(script_num)	   'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script   'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name		   = "MEMO from List"
script_array_admin_main(script_num).file_name		   = "memo-from-list.vbs"
script_array_admin_main(script_num).description		   = "Creates the same MEMO on cases listed in REPT/ACTV, manually entered, or from an Excel spreadsheet of your choice."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_admin_main(script_num)		'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name			= "MFIP Sanction FIATer"											'Script name
script_array_admin_main(script_num).file_name			= "mfip-sanction-fiater.vbs"										'Script URL
script_array_admin_main(script_num).description			= "FIATs MFIP sanction actions for CS, ES and both types of sanctions."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_admin_main(script_num)		'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name			= "Sanction Member Info"										'Script name
script_array_admin_main(script_num).file_name			= "sanction-member-info.vbs"									'Script URL
script_array_admin_main(script_num).description			= "BULK script to gather information for for MFIP participants on REPT/MFCM."

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array_admin_main(script_num)	    'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script    'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name		    = "Training Case Creator"
script_array_admin_main(script_num).file_name			= "training-case-creator.vbs"
script_array_admin_main(script_num).description		    = "Creates training case scenarios en masse and XFERs them to workers."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_admin_main(script_num)		'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name			= "WF1 Case Status"													'Script name
script_array_admin_main(script_num).file_name			= "wf1-case-status.vbs"												'Script URL
script_array_admin_main(script_num).description			= "Updates a list of cases from Excel with current case and ABAWD status information."

'QI scripts----------------------------------------------------------------------------------------------------
script_num = 0
ReDim Preserve script_array_QI_list(script_num)		'Resets the array to add one more element to it
Set script_array_QI_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_QI_list(script_num).script_name		= "Banked Months Review"													'Script name
script_array_QI_list(script_num).file_name			= "banked-months-review.vbs"												'Script URL
script_array_QI_list(script_num).description		= "Script to assist in the review and processing of SNAP Banked Months."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_QI_list(script_num)		'Resets the array to add one more element to it
Set script_array_QI_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_QI_list(script_num).script_name		= "Banked Months Individual Case Notes"													'Script name
script_array_QI_list(script_num).file_name			= "individual-banked-note.vbs"												'Script URL
script_array_QI_list(script_num).description		= "Script to enter case notes in line with BULK Processing script."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_QI_list(script_num)
Set script_array_QI_list(script_num) = new script
script_array_QI_list(script_num).script_name 		= "Budget Estimator"											'Script name
script_array_QI_list(script_num).file_name 			= "budget-estimator.vbs"										'Script URL
script_array_QI_list(script_num).description 		= "UTILITIES script that can be used to calculate an expected budget outside of MAXIS."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_QI_list(script_num)		'Resets the array to add one more element to it
Set script_array_QI_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_QI_list(script_num).script_name		= "Individual Appointment Notice"													'Script name
script_array_QI_list(script_num).file_name			= "individual-appointment-letter.vbs"												'Script URL
script_array_QI_list(script_num).description		= "Sends an appointment letter for a single case, with the same wording as On Demand Applications"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_QI_list(script_num)		'Resets the array to add one more element to it
Set script_array_QI_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_QI_list(script_num).script_name		= "Individual NOMI"													'Script name
script_array_QI_list(script_num).file_name			= "individual-nomi.vbs"												'Script URL
script_array_QI_list(script_num).description		= "Sends a NOMI for a single case, with the same wording as On Demand Applications"

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array_QI_list(script_num)		 'Resets the array to add one more element to it
Set script_array_QI_list(script_num) = new script	 'Set this array element to be a new script. Script details below...
script_array_QI_list(script_num).script_name		 = "On Demand Waiver - Applications"									'Script name
script_array_QI_list(script_num).file_name			 = "bulk-applications.vbs"												'Script URL
script_array_QI_list(script_num).description		 = "BULK script to collect information for cases that require an interview for the On Demand Waiver."

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array_QI_list(script_num)		 'Resets the array to add one more element to it
Set script_array_QI_list(script_num) = new script	 'Set this array element to be a new script. Script details below...
script_array_QI_list(script_num).script_name		 = "QC Results"									'Script name
script_array_QI_list(script_num).file_name			 = "qc-review.vbs"												'Script URL
script_array_QI_list(script_num).description		 = "Case note and WCOM script to support the DHS QC process."

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array_QI_list(script_num)		'Resets the array to add one more element to it
Set script_array_QI_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_QI_list(script_num).script_name		= "QI Renewal Accuracy"                                              'Script name
script_array_QI_list(script_num).file_name			= "qi-renewal-accuracy.vbs"                                          'Script URL
script_array_QI_list(script_num).description		= "Template for documenting specific renewal information that has been reviewed by policy experts."

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array_QI_list(script_num)		'Resets the array to add one more element to it
Set script_array_QI_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_QI_list(script_num).script_name		= "Targeted SNAP Review Selection"
script_array_QI_list(script_num).file_name			= "targeted-snap-review-selection.vbs"
script_array_QI_list(script_num).description		= "Creates a list of SNAP cases meeting review criteria and selects a random sample for review."

'BZST scripts----------------------------------------------------------------------------------------------------
script_num = 0
ReDim Preserve script_array_BZ_list(script_num)		'Resets the array to add one more element to it
Set script_array_BZ_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BZ_list(script_num).script_name		= "AVS Report"													'Script name
script_array_BZ_list(script_num).file_name			= "avs-report.vbs"												'Script URL
script_array_BZ_list(script_num).description		= "BULK script that supports the AVS processing needs for active MA recipients."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BZ_list(script_num)		'Resets the array to add one more element to it
Set script_array_BZ_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BZ_list(script_num).script_name		= "COLA Decimator"													'Script name
script_array_BZ_list(script_num).file_name			= "cola-decimator.vbs"												'Script URL
script_array_BZ_list(script_num).description		= "BULK script that deletes and case notes auto-approval COLA messages."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BZ_list(script_num)		'Resets the array to add one more element to it
Set script_array_BZ_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BZ_list(script_num).script_name		= "DAIL CCD"													'Script name
script_array_BZ_list(script_num).file_name			= "dail-case-note-capture.vbs"												'Script URL
script_array_BZ_list(script_num).description		= "BULK script that captures, case notes and deletes specific DAILS based on content, and collects them into an Excel spreadsheet."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BZ_list(script_num)		'Resets the array to add one more element to it
Set script_array_BZ_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BZ_list(script_num).script_name		= "DAIL Decimator"													'Script name
script_array_BZ_list(script_num).file_name			= "dail-decimator.vbs"												'Script URL
script_array_BZ_list(script_num).description		= "BULK script that deletes specific DAILS based on content, and collects them into an Excel spreadsheet."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BZ_list(script_num)		'Resets the array to add one more element to it
Set script_array_BZ_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BZ_list(script_num).script_name		= "DAIL Decimator - TB Edition"													'Script name
script_array_BZ_list(script_num).file_name			= "dail-decimator-task-based-edition.vbs"												'Script URL
script_array_BZ_list(script_num).description		= "BULK script that deletes specific DAILS for identified populations, and outputs them into a SQL Database."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BZ_list(script_num)		'Resets the array to add one more element to it
Set script_array_BZ_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BZ_list(script_num).script_name		= "DAIL FAD Task Based Assignments"													'Script name
script_array_BZ_list(script_num).file_name			= "dail-task-based-assignments.vbs"												'Script URL
script_array_BZ_list(script_num).description		= "BULK script completed DAIL DECIMATOR funtionality for families-population baskets and specific DAIL types for assignment."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BZ_list(script_num)		'Resets the array to add one more element to it
Set script_array_BZ_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BZ_list(script_num).script_name		= "Get basket number"													'Script name
script_array_BZ_list(script_num).file_name			= "get-basket-number.vbs"												'Script URL
script_array_BZ_list(script_num).description		= "BULK script that will obtain the basket number and population."

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array_BZ_list(script_num)		'Resets the array to add one more element to it
Set script_array_BZ_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BZ_list(script_num).script_name		= "Send CBO Manual Referrals"										'Script name
script_array_BZ_list(script_num).file_name			= "send-cbo-manual-referrals.vbs"									'Script URL
script_array_BZ_list(script_num).description		= "Sends manual referrals for a list of cases provided by Employment and Training."

'MONTHLY TASKS scripts----------------------------------------------------------------------------------------------------
script_num = 0
ReDim Preserve script_array_task_list(script_num)   'Resets the array to add one more element to it
Set script_array_task_list(script_num) = new script 'Set this array element to be a new script. Script details below...
script_array_task_list(script_num).script_name 		= " ABAWD Report "											'Script name
script_array_task_list(script_num).file_name 		= "abawd-report.vbs"										'Script URL
script_array_task_list(script_num).description 		= "BULK script that gathers ABAWD/FSET codes for members on SNAP/MFIP active cases."

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array_task_list(script_num)   'Resets the array to add one more element to it
Set script_array_task_list(script_num) = new script 'Set this array element to be a new script. Script details below...
script_array_task_list(script_num).script_name 		= "Auto-Dialer Case Status"											'Script name
script_array_task_list(script_num).file_name 		= "auto-dialer-case-status.vbs"										'Script URL
script_array_task_list(script_num).description 		= "BULK script that gathers case status for cases with recerts for SNAP/MFIP the previous month."

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array_task_list(script_num)	'Resets the array to add one more element to it
Set script_array_task_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_task_list(script_num).script_name		= "Close GRH Rate 2 in MMIS"													'Script name
script_array_task_list(script_num).file_name		= "close-GRH-rate-2-in-MMIS.vbs"												'Script URL
script_array_task_list(script_num).description		= "Script to assist in closing SSR agreements in MMIS for GRH Rate 2 cases."

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array_task_list(script_num)	'Resets the array to add one more element to it
Set script_array_task_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_task_list(script_num).script_name		= "Gather Banked Months Cases"													'Script name
script_array_task_list(script_num).file_name		= "gather-banked-months-cases.vbs"												'Script URL
script_array_task_list(script_num).description		= "Script to determine which cases are new banked months cases to be reviewed by QI."

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array_task_list(script_num)	'Resets the array to add one more element to it
Set script_array_task_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_task_list(script_num).script_name		= "Inactive Transfer"													'Script name
script_array_task_list(script_num).file_name		= "bulk-inactive-transfer.vbs"												'Script URL
script_array_task_list(script_num).description		= "Script to transfer inactive cases via SPEC/XFER"

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array_task_list(script_num)	'Resets the array to add one more element to it
Set script_array_task_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_task_list(script_num).script_name		= "Individual Recertification Notices"													'Script name
script_array_task_list(script_num).file_name		= "individual-recertification-notices.vbs"												'Script URL
script_array_task_list(script_num).description		= "NOTICES Script that will send ODW Recert Appointment Letter or NOMI on a single case."

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array_task_list(script_num)	'Resets the array to add one more element to it
Set script_array_task_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_task_list(script_num).script_name		= "Interview Required"													'Script name
script_array_task_list(script_num).file_name		= "interview-required.vbs"												'Script URL
script_array_task_list(script_num).description		= "BULK script to collect case information for cases that require an interview for SNAP/MFIP."

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array_task_list(script_num)	'Resets the array to add one more element to it
Set script_array_task_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_task_list(script_num).script_name		= "MAXIS to METS Conversion"													'Script name
script_array_task_list(script_num).file_name		= "maxis-to-mets-conversion.vbs"												'Script URL
script_array_task_list(script_num).description		= "BULK script to collect case information for cases that may need to convert from MAXIS to METS."

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array_task_list(script_num)   'Resets the array to add one more element to it
Set script_array_task_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_task_list(script_num).script_name		 = "On Demand Waiver - Recertifications"													'Script name
script_array_task_list(script_num).file_name		 = "bulk-recertifications.vbs"												'Script URL
script_array_task_list(script_num).description		 = "BULK script to send notices for cases at recertification that require an interview for the On Demand Waiver."

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array_task_list(script_num)	'Resets the array to add one more element to it
Set script_array_task_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_task_list(script_num).script_name		= "Paperless IR"                                                       'Script name
script_array_task_list(script_num).file_name		= "new-paperless-ir.vbs"                                                     'Script URL
script_array_task_list(script_num).description		= "Updates cases on a caseload(s) that require paperless IR processing. Does not approve cases."

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array_task_list(script_num)   'Resets the array to add one more element to it
Set script_array_task_list(script_num) = new script	 'Set this array element to be a new script. Script details below...
script_array_task_list(script_num).script_name		 = " Resolve HC EOMC in MMIS "													'Script name
script_array_task_list(script_num).file_name		 = "resolve-hc-eomc-in-mmis.vbs"												'Script URL
script_array_task_list(script_num).description		 = "BULK script that checks MMIS for all cases on EOMC for HC to ensure MMIS is set to close."

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array_task_list(script_num)	'Resets the array to add one more element to it
Set script_array_task_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_task_list(script_num).script_name		= " REVW/MONT Closures "													'Script name
script_array_task_list(script_num).file_name		= "revw-mont-closures.vbs"												'Script URL
script_array_task_list(script_num).description		= "Case notes all cases on REPT/REVW or REPT/MONT that are closing for missing or incomplete CAF/HRF/CSR/HC ER."

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array_task_list(script_num)	'Resets the array to add one more element to it
Set script_array_task_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_task_list(script_num).script_name		= "UNEA Updater"										'Script name
script_array_task_list(script_num).file_name		= "unea-updater.vbs"									'Script URL
script_array_task_list(script_num).description		= "BULK script that updates UNEA information and sends SPEC/MEMO for VA cases at ER."

'Starting these with a very high number, higher than the normal possible amount of buttons.
'We're doing this because we want to assign a value to each button pressed, and we want
'that value to change with each button. The button_placeholder will be placed in the .button
'property for each script item. This allows it to both escape the Function and resize
'near infinitely. We use dummy numbers for the other selector buttons for much the same reason,
'to force the value of ButtonPressed to hold in near infinite iterations.

button_placeholder 	= 24601
admin_main_button	= 1000
QI_button	        = 2000
BZ_button           = 3000
tasks_button        = 4000

show_QI_button = False
show_BZ_button = False
show_tasks_button = False

'Displays the dialog
Do
    'BZST scripts menu authorization
    If user_ID_for_validation = "ILFE001" OR user_ID_for_validation = "WFS395" OR user_ID_for_validation = "CALO001" OR user_ID_for_validation = "WFX901" OR user_ID_for_validation = "WFU851" then
        show_BZ_button = True
        show_QI_button = True
        show_tasks_button = True
    End if
    'QI scripts menu authorization
    If user_ID_for_validation = "WFI021" OR user_ID_for_validation = "WFU161" OR user_ID_for_validation = "WF7638" OR user_ID_for_validation = "WFP106" OR user_ID_for_validation = "WFQ898" OR user_ID_for_validation = "WFK093" OR _
    user_ID_for_validation = "WF1875" OR user_ID_for_validation = "WFM207" OR user_ID_for_validation = "WFG492" OR user_ID_for_validation = "WFP803" OR user_ID_for_validation = "WFX490" then show_QI_button = TRUE

	If ButtonPressed = "" or ButtonPressed = admin_main_button then
        declare_admin_menu_dialog(script_array_admin_main)
	elseif ButtonPressed = QI_button then
        If show_QI_button = True then declare_admin_menu_dialog(script_array_QI_list)
    elseif ButtonPressed = BZ_button then
        If show_BZ_button = True then declare_admin_menu_dialog(script_array_BZ_list)
    elseif ButtonPressed = tasks_button then
        If show_tasks_button = True then declare_admin_menu_dialog(script_array_task_list)
    end if

    dialog admin_dialog
	If ButtonPressed = 0 then stopscript

    'Opening the SIR Instructions
	'IF buttonpressed = SIR_instructions_button then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/Script%20Instructions%20Wiki/Notices%20scripts.aspx")
Loop until ButtonPressed <> admin_main_button and _
			ButtonPressed <> QI_button and _
            ButtonPressed <> BZ_button and _
            ButtonPressed <> tasks_button

'Runs through each script in the array... if the selected script (buttonpressed) is in the array, it'll run_from_GitHub
For i = 0 to ubound(script_array_admin_main)
	If ButtonPressed = script_array_admin_main(i).button then call run_from_GitHub(script_repository & "admin/" & script_array_admin_main(i).file_name)
Next

For i = 0 to ubound(script_array_QI_list)
	If ButtonPressed = script_array_QI_list(i).button then call run_from_GitHub(script_repository & "admin/" & script_array_QI_list(i).file_name)
Next

For i = 0 to ubound(script_array_BZ_list)
	If ButtonPressed = script_array_BZ_list(i).button then call run_from_GitHub(script_repository & "admin/" & script_array_BZ_list(i).file_name)
Next

For i = 0 to ubound(script_array_task_list)
	If ButtonPressed = script_array_task_list(i).button then call run_from_GitHub(script_repository & "admin/" & script_array_task_list(i).file_name)
Next

stopscript
