'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - MAIN MENU.vbs"
start_time = timer

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
call changelog_update("04/29/2019", "Retired the following BULK scripts: CEI premium noter, Housing Grant Exemption Finder, INAC scrubber, NON-MAGI HC Info, REPT-ARST, Returned Mail, Spenddown Report, SWKR List Generator, TIKL From List and Update EOMC List. ", "Ilse Ferris, Hennepin County")
call changelog_update("09/25/2017", "Added the following BULK scripts: 7th Sanction Identifier, EMPS, FSS Info, GA Advanced Age Identifier, GRH Professional Need, Homesless Discrepancy, MFIP Sanction and Paperless IR. Removed REPT GRMR, this is no longer a report option in MAXIS.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'CUSTOM FUNCTIONS===========================================================================================================
Function declare_BULK_menu_dialog(script_array)
	BeginDialog BULK_dialog, 0, 0, 545, 380, "BULK Scripts"
	 	Text 5, 5, 435, 10, "Bulk scripts main menu: select the script to run from the choices below."
	  	ButtonGroup ButtonPressed
		 	PushButton 015, 35, 60, 15, "BULK ACTIONS", 		BULK_main_button
		 	PushButton 075, 35, 50, 15, "BULK LISTS",			BULK_lists_button
			PushButton 125, 35, 80, 15, "ENHANCED LISTS", 		BULK_enhanced_lists_button
		 	PushButton 445, 10, 65, 10, "Instructions", 	Instructions_button

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

		CancelButton 480, 360, 50, 15
		GroupBox 5, 20, 205, 35, "BULK Sub-Menus"
	EndDialog
End function
'END CUSTOM FUNCTIONS=======================================================================================================

'VARIABLES TO DECLARE=======================================================================================================

'Declaring the variable names to cut down on the number of arguments that need to be passed through the function.
DIM ButtonPressed
DIM Instructions_button
dim BULK_dialog

script_array_BULK_main = array()
script_array_BULK_enhanced_list = array()
script_array_BULK_stat_list = array()

'END VARIABLES TO DECLARE===================================================================================================

'LIST OF SCRIPTS================================================================================================================

'INSTRUCTIONS: simply add your new script below. Scripts are listed in alphabetical order. Copy a block of code from above and paste your script info in. The function does the rest.

'-------------------------------------------------------------------------------------------------------------------------BULK MAIN MENU
'Resetting the variable
script_num = 0											'sets to 0, 1st script in array
ReDim Preserve script_array_BULK_main(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_main(script_num).script_name 			= "CASE/NOTE from List"																	'Script name
script_array_BULK_main(script_num).file_name 			= "case-note-from-list.vbs"																'Script URL
script_array_BULK_main(script_num).description 			= "Creates the same case note on cases listed in REPT/ACTV, manually entered, or from an Excel spreadsheet of your choice."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_main(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_main(script_num).script_name 			= "Case Transfer"																		'Script name
script_array_BULK_main(script_num).file_name 			= "case-transfer.vbs"																	'Script URL
script_array_BULK_main(script_num).description 			= "Searches caseload(s) by selected parameters. Transfers a specified number of those cases to another worker. Creates list of these cases."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_main(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_main(script_num).script_name			= "Drug Felon list"
script_array_BULK_main(script_num).file_name			= "drug-felon-list.vbs"
script_array_BULK_main(script_num).description			= "Reviews the Drug Felon list from DHS to update these cases."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_main(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_main(script_num).script_name			= "MEMO from List"
script_array_BULK_main(script_num).file_name			= "memo-from-list.vbs"
script_array_BULK_main(script_num).description			= "Creates the same MEMO on cases listed in REPT/ACTV, manually entered, or from an Excel spreadsheet of your choice."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_main(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_main(script_num).script_name			= " REVW/MONT Closures "													'needs spaces to generate button width properly.
script_array_BULK_main(script_num).file_name			= "revw-mont-closures.vbs"
script_array_BULK_main(script_num).description			= "Case notes all cases on REPT/REVW or REPT/MONT that are closing for missing or incomplete CAF/HRF/CSR/HC ER."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_main(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_main(script_num).script_name			= "Targeted SNAP Review Selection"
script_array_BULK_main(script_num).file_name			= "targeted-snap-review-selection.vbs"
script_array_BULK_main(script_num).description			= "Creates a list of SNAP cases meeting review criteria and selects a random sample for review."

'-------------------------------------------------------------------------------------------------------------------------ENHANCED LISTS
'Resetting the variable

script_num = 0												'sets to 0, 1st script in array
ReDim Preserve script_array_BULK_enhanced_list(script_num)	'Resets the array to add one more element to t
Set script_array_BULK_enhanced_list(script_num) = new script'Set this array element to be a new script. Script details below...
script_array_BULK_enhanced_list(script_num).script_name 	= "7th Sanction Identifier"																'Script name
script_array_BULK_enhanced_list(script_num).file_name		= "7th-sanction-identifier.vbs"															'Script URL
script_array_BULK_enhanced_list(script_num).description 	= "Creates a list of active MFIP cases that may meet 7th sanction criteria into an Excel spreadsheet"

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_enhanced_list(script_num)	'Resets the array to add one more element to t
Set script_array_BULK_enhanced_list(script_num) = new script'Set this array element to be a new script. Script details below...
script_array_BULK_enhanced_list(script_num).script_name 	= " Check SNAP for GA/RCA "													'needs spaces to generate button width properly.
script_array_BULK_enhanced_list(script_num).file_name		= "check-snap-for-ga-rca.vbs"
script_array_BULK_enhanced_list(script_num).description 	= "Compares the amount of GA and RCA FIAT'd into SNAP and creates a list of the results."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_enhanced_list(script_num)	'Resets the array to add one more element to it
Set script_array_BULK_enhanced_list(script_num) = new script'Set this array element to be a new script. Script details below...
script_array_BULK_enhanced_list(script_num).script_name 	= " EXP SNAP Review "
script_array_BULK_enhanced_list(script_num).file_name		= "exp-snap-review.vbs"
script_array_BULK_enhanced_list(script_num).description 	= "Creates a list of PND1/PND2 cases that need to reviewed for EXP SNAP criteria."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_enhanced_list(script_num)	'Resets the array to add one more element to it
Set script_array_BULK_enhanced_list(script_num) = new script'Set this array element to be a new script. Script details below...
script_array_BULK_enhanced_list(script_num).script_name 	= "Find Updated Panels"
script_array_BULK_enhanced_list(script_num).file_name		= "find-panel-update-date.vbs"
script_array_BULK_enhanced_list(script_num).description 	= "Creates a list of cases from a caseload(s) showing when selected panels have been updated."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_enhanced_list(script_num)	'Resets the array to add one more element to it
Set script_array_BULK_enhanced_list(script_num) = new script'Set this array element to be a new script. Script details below...
script_array_BULK_enhanced_list(script_num).script_name 	= "FSS Info"
script_array_BULK_enhanced_list(script_num).file_name		= "fss-info.vbs"
script_array_BULK_enhanced_list(script_num).description 	= "Creates a list of FSS identified info from EMPS and DISA into an Excel spreadsheet."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_enhanced_list(script_num)	'Resets the array to add one more element to it
Set script_array_BULK_enhanced_list(script_num) = new script'Set this array element to be a new script. Script details below...
script_array_BULK_enhanced_list(script_num).script_name 	= "GA Adv. Age Identifier"
script_array_BULK_enhanced_list(script_num).file_name		= "ga-advanced-age-identifier.vbs"
script_array_BULK_enhanced_list(script_num).description 	= "Creates a list of GA adv. age identified info into an Excel spreadsheet."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_enhanced_list(script_num)	'Resets the array to add one more element to it
Set script_array_BULK_enhanced_list(script_num) = new script'Set this array element to be a new script. Script details below...
script_array_BULK_enhanced_list(script_num).script_name 	= "GRH Professional Need"
script_array_BULK_enhanced_list(script_num).file_name		= "grh-professional-need.vbs"
script_array_BULK_enhanced_list(script_num).description 	= "Creates a list of active GRH cases and identified info into an Excel spreadsheet."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_enhanced_list(script_num)	'Resets the array to add one more element to it
Set script_array_BULK_enhanced_list(script_num) = new script'Set this array element to be a new script. Script details below...
script_array_BULK_enhanced_list(script_num).script_name 	= "Homeless Discrepancy"
script_array_BULK_enhanced_list(script_num).file_name		= "homeless-discprepancy.vbs"
script_array_BULK_enhanced_list(script_num).description 	= "Creates a list of active SNAP/MFIP cases with identified info into an Excel spreadsheet."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_enhanced_list(script_num)	'Resets the array to add one more element to it
Set script_array_BULK_enhanced_list(script_num) = new script'Set this array element to be a new script. Script details below...
script_array_BULK_enhanced_list(script_num).script_name 	= "MA-EPD/Medi Pt B CEI"
script_array_BULK_enhanced_list(script_num).file_name		= "find-maepd-medi-cei.vbs"
script_array_BULK_enhanced_list(script_num).description 	= "Creates a list of cases and clients active on MA-EPD and Medicare Part B that are eligible for Part B reimbursement."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_enhanced_list(script_num)	'Resets the array to add one more element to it
Set script_array_BULK_enhanced_list(script_num) = new script'Set this array element to be a new script. Script details below...
script_array_BULK_enhanced_list(script_num).script_name 	= "MFIP Sanction"
script_array_BULK_enhanced_list(script_num).file_name		= "mfip-sanction.vbs"
script_array_BULK_enhanced_list(script_num).description 	= "Creates a list of active MFIP cases with identified info into an Excel spreadsheet."

'-------------------------------------------------------------------------------------------------------------------------BULK STAT list
'Resetting the variable
script_num = 0
ReDim Preserve script_array_BULK_stat_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_stat_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BULK_stat_list(script_num).script_name 		= "ADDR"																		'Script name
script_array_BULK_stat_list(script_num).file_name			= "address-report.vbs"
script_array_BULK_stat_list(script_num).description 		= "Creates a list of all addresses from a caseload(or entire county)."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_stat_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_stat_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BULK_stat_list(script_num).script_name 		= "ACTV"
script_array_BULK_stat_list(script_num).file_name			= "rept-actv-list.vbs"
script_array_BULK_stat_list(script_num).description 		= "Pulls a list of cases in REPT/ACTV into an Excel spreadsheet."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_stat_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_stat_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BULK_stat_list(script_num).script_name 		= "DAIL"
script_array_BULK_stat_list(script_num).file_name			= "dail-report.vbs"
script_array_BULK_stat_list(script_num).description 		= "Pulls a list of DAILS in DAIL/DAIL into an Excel spreadsheet."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_stat_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_stat_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BULK_stat_list(script_num).script_name 		= "EMPS"
script_array_BULK_stat_list(script_num).file_name			= "emps.vbs"
script_array_BULK_stat_list(script_num).description 		= "Pulls a list of STAT/EMPS information into an Excel spreadsheet."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_stat_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_stat_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BULK_stat_list(script_num).script_name 		= " EOMC "													'needs spaces to generate button width properly.
script_array_BULK_stat_list(script_num).file_name			= "rept-eomc-list.vbs"
script_array_BULK_stat_list(script_num).description 		= "Pulls a list of cases in REPT/EOMC into an Excel spreadsheet."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_stat_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_stat_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BULK_stat_list(script_num).script_name 		= "IEVC"
script_array_BULK_stat_list(script_num).file_name			= "rept-ievc-list.vbs"
script_array_BULK_stat_list(script_num).description 		= "Pulls a list of cases in REPT/IEVC into an Excel spreadsheet."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_stat_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_stat_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BULK_stat_list(script_num).script_name 		= "INAC"
script_array_BULK_stat_list(script_num).file_name			= "rept-inac-list.vbs"
script_array_BULK_stat_list(script_num).description 		= "Pulls a list of cases in REPT/INAC into an Excel spreadsheet."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_stat_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_stat_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BULK_stat_list(script_num).script_name 		= "INTR"
script_array_BULK_stat_list(script_num).file_name			= "rept-intr-list.vbs"
script_array_BULK_stat_list(script_num).description 		= "Pulls a list of cases in REPT/INTR into an Excel spreadsheet."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_stat_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_stat_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BULK_stat_list(script_num).script_name 		= "LTC-GRH List Gen"
script_array_BULK_stat_list(script_num).file_name			= "ltc-grh-list-generator.vbs"
script_array_BULK_stat_list(script_num).description 		= "Creates a list of FACIs, AREPs, and waiver types assigned to the various cases in a caseload(s)."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_stat_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_stat_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BULK_stat_list(script_num).script_name 		= " MAMS "													'needs spaces to generate button width properly.
script_array_BULK_stat_list(script_num).file_name			= "rept-mams-list.vbs"
script_array_BULK_stat_list(script_num).description 		= "Pulls a list of cases in REPT/MAMS into an Excel spreadsheet."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_stat_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_stat_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BULK_stat_list(script_num).script_name 		= " MFCM "													'needs spaces to generate button width properly.
script_array_BULK_stat_list(script_num).file_name			= "rept-mfcm-list.vbs"
script_array_BULK_stat_list(script_num).description 		= "Pulls a list of cases in REPT/MFCM into an Excel spreadsheet."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_stat_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_stat_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BULK_stat_list(script_num).script_name 		= " MONT "													'needs spaces to generate button width properly.
script_array_BULK_stat_list(script_num).file_name			= "rept-mont-list.vbs"
script_array_BULK_stat_list(script_num).description 		= "Pulls a list of cases in REPT/MONT into an Excel spreadsheet."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_stat_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_stat_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BULK_stat_list(script_num).script_name 		= " MRSR "													'needs spaces to generate button width properly.
script_array_BULK_stat_list(script_num).file_name			= "rept-mrsr-list.vbs"
script_array_BULK_stat_list(script_num).description 		= "Pulls a list of cases in REPT/MRSR into an Excel spreadsheet."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_stat_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_stat_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BULK_stat_list(script_num).script_name 		= "PND1"
script_array_BULK_stat_list(script_num).file_name			= "rept-pnd1-list.vbs"
script_array_BULK_stat_list(script_num).description 		= "Pulls a list of cases in REPT/PND1 into an Excel spreadsheet."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_stat_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_stat_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BULK_stat_list(script_num).script_name 		= "PND2"
script_array_BULK_stat_list(script_num).file_name			= "rept-pnd2-list.vbs"
script_array_BULK_stat_list(script_num).description 		= "Pulls a list of cases in REPT/PND2 into an Excel spreadsheet."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_stat_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_stat_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BULK_stat_list(script_num).script_name			= " REVS "													'needs spaces to generate button width properly.
script_array_BULK_stat_list(script_num).file_name			= "rept-revs-list.vbs"
script_array_BULK_stat_list(script_num).description 		= "Pulls a list of cases in REPT/REVS into an Excel spreadsheet."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_BULK_stat_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_stat_list(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_BULK_stat_list(script_num).script_name 		= " REVW "													'needs spaces to generate button width properly.
script_array_BULK_stat_list(script_num).file_name			= "rept-revw-list.vbs"
script_array_BULK_stat_list(script_num).description 		= "Pulls a list of cases in REPT/REVW into an Excel spreadsheet."


'Starting these with a very high number, higher than the normal possible amount of buttons.
'	We're doing this because we want to assign a value to each button pressed, and we want
'	that value to change with each button. The button_placeholder will be placed in the .button
'	property for each script item. This allows it to both escape the Function and resize
'	near infinitely. We use dummy numbers for the other selector buttons for much the same reason,
'	to force the value of ButtonPressed to hold in near infinite iterations.
button_placeholder 			= 24601
BULK_main_button			= 1000
BULK_enhanced_lists_button  = 2000
BULK_lists_button			= 3000

'Displays the dialog
Do
	If ButtonPressed = "" or ButtonPressed = BULK_main_button then
		declare_BULK_menu_dialog(script_array_BULK_main)
	ELSEif ButtonPressed = BULK_enhanced_lists_button then
		declare_BULK_menu_dialog(script_array_BULK_enhanced_list)
	ElseIf ButtonPressed = BULK_lists_button then
		declare_BULK_menu_dialog(script_array_BULK_stat_list)
	End if

	dialog BULK_dialog

	If ButtonPressed = 0 then stopscript
    'Opening the Instructions
	IF buttonpressed = Instructions_button then CreateObject("WScript.Shell").Run("https://dept.hennepin.us/hsphd/sa/ews/BlueZone_Script_Instructions/Forms/AllItems.aspx?RootFolder=%2Fhsphd%2Fsa%2Fews%2FBlueZone%5FScript%5FInstructions%2FBULK&FolderCTID=0x012000A05B86818A1703428050D2E34B3E8EA1&View=%7BFFD55BF9%2D6CDF%2D4B5C%2DB47B%2D3701445A9B34%7D")
Loop until 	ButtonPressed <> Instructions_button and _
			ButtonPressed <> BULK_main_button and _
			ButtonPressed <> BULK_enhanced_lists_button and _
			ButtonPressed <> BULK_lists_button

'MsgBox buttonpressed = script_array_BULK_main(0).button

'Runs through each script in the array... if the selected script (buttonpressed) is in the array, it'll run_from_GitHub
For i = 0 to ubound(script_array_BULK_main)
	If ButtonPressed = script_array_BULK_main(i).button then call run_from_GitHub(script_repository & "bulk/" & script_array_BULK_main(i).file_name)
Next

For i = 0 to ubound(script_array_BULK_enhanced_list)
	If ButtonPressed = script_array_BULK_enhanced_list(i).button then call run_from_GitHub(script_repository & "bulk/" & script_array_BULK_enhanced_list(i).file_name)
Next

For i = 0 to ubound(script_array_BULK_stat_list)
	If ButtonPressed = script_array_BULK_stat_list(i).button then call run_from_GitHub(script_repository & "bulk/" & script_array_BULK_stat_list(i).file_name)
Next
stopscript
