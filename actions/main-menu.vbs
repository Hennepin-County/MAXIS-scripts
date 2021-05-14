'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - MAIN MENU.vbs"
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
Call changelog_update("04/02/2021", "NEW SCRIPT ADDED##~## ##~##MA FIATER for GRH MSA - script to FIAT the Method X budget to be a Method B budget for MSA and GRH cases because these programs no longer grant automatic MA eligibility.##~## ##~##See the OneSource page 'Health Care Eligibility Determination for a GRH or MSA MAXIS case' for details on this process.##~##**This script is in testing at this time, please report any issues or concerns to the BlueZone Script Team.##~##", "Casey Love, Hennepin County")
Call changelog_update("09/25/2020", "Script has been retired.##~## ##~##ACTIONS - New Job Reported is retired and no longer available.##~## ##~##Job Change Reported now has the functionality from New Job Reported.##~##", "09/25/2020")
Call changelog_update("09/23/2020", "!!! NEW SCRIPT !!! ##~## ##~##Job Change Reported##~## ##~##This script will replace 'New Job Reported', it handles the functionality of a new job reported and adds functionality to handle for changes to existing jobs and stop work reports.##~##This script is not intended for budgeting with received verifications, it is used for the initial report when no verif have been received.##~## ##~##Expect 'New Job Reported' to be retired soon.##~##", "Casey Love, Hennepin County")
call changelog_update("06/03/2020", "Removed the script 'ABAWD FIATer' as there is currently an ABAWD waiver in effect due to the Coronavirus Response Act. Processing cases as ABAWDs, and so FIATing is unnecessary.", "Casey Love, Hennepin County")
call changelog_update("07/31/2019", "Removed the script ACCT Panel Updater. This functionality is now available in the script NOTES - Docs Received.", "Casey Love, Hennepin County")
call changelog_update("03/13/2019", "Removed Paystubs Received. Use EARNED INCOME BUDGETING instead to update a case with paystubs received.", "Casey Love, Hennepin County")
call changelog_update("03/05/2019", "Added EARNED INCOME BUDGETING Script.", "Casey Love, Hennepin County")
call changelog_update("12/12/2018", "Added COUNTED ABAWD MONTHS script under ABAWD sub-menu.", "Ilse Ferris, Hennepin County")
call changelog_update("10/29/2018", "Removed FSET SANCTION script. Cases no longer require FSET sanctions (voluntary compliance).", "Ilse Ferris, Hennepin County")
call changelog_update("02/09/2018", "Added ADD GRH RATE 2 TO MMIS script for GRH Rate 2 case.", "Ilse Ferris, Hennepin County")
call changelog_update("01/12/2018", "Retired ACTIONS - ABAWD MINOR CHILD EXEMPTION FIATER'. MAXIS has been updated to support this process.", "Ilse Ferris, Hennepin County")
call changelog_update("01/17/2017", "Added new ACTION script 'ABAWD FIATER'.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

class subcat
	public subcat_name
	public subcat_button
End class

Function declare_main_menu_dialog(script_category)

	'Runs through each script in the array and generates a list of subcategories based on the category located in the function. Also modifies the script description if it's from the last two months, to include a "NEW!!!" notification.
	For current_script = 0 to ubound(script_array)
		'Subcategory handling (creating a second list as a string which gets converted later to an array)
		If ucase(script_array(current_script).category) = ucase(script_category) then																								'If the script in the array is of the correct category (ACTIONS/NOTES/ETC)...
			For each listed_subcategory in script_array(current_script).subcategory																									'...then iterate through each listed subcategory, and...
				If listed_subcategory <> "" and InStr(subcategory_list, ucase(listed_subcategory)) = 0 then subcategory_list = subcategory_list & "|" & ucase(listed_subcategory)	'...if the listed subcategory isn't blank and isn't already in the list, then add it to our handy-dandy list.
			Next
		End if
		' 'Adds a "NEW!!!" notification to the description if the script is from the last two months.
		' If DateDiff("m", script_array(current_script).release_date, DateAdd("m", -2, date)) <= 0 then
		' 	script_array(current_script).description = "NEW " & script_array(current_script).release_date & "!!! --- " & script_array(current_script).description
		' 	script_array(current_script).release_date = "12/12/1999" 'backs this out and makes it really old so it doesn't repeat each time the dialog loops. This prevents NEW!!!... from showing multiple times in the description.
		' End if

	Next

	subcategory_list = split(subcategory_list, "|")

	For i = 0 to ubound(subcategory_list)
		ReDim Preserve subcategory_array(i)
		set subcategory_array(i) = new subcat
		If subcategory_list(i) = "" then subcategory_list(i) = "MAIN"
		subcategory_array(i).subcat_name = subcategory_list(i)
	Next

    dlg_len = 60
    For current_script = 0 to ubound(script_array)
        script_array(current_script).show_script = FALSE
        If ucase(script_array(current_script).category) = ucase(script_category) then

            '<<<<<<RIGHT HERE IT SHOULD ITERATE THROUGH SUBCATEGORIES AND BUTTONS PRESSED TO DETERMINE WHAT THE CURRENTLY DISPLAYED SUBCATEGORY SHOULD BE, THEN ONLY DISPLAY SCRIPTS THAT MATCH THAT CRITERIA
            'Joins all subcategories together
            subcategory_string = ucase(join(script_array(current_script).subcategory))

            'Accounts for scripts without subcategories
            If subcategory_string = "" then subcategory_string = "MAIN"		'<<<THIS COULD BE A PROPERTY OF THE CLASS


            'If the selected subcategory is in the subcategory string, it will display those scripts
            If InStr(subcategory_string, subcategory_selected) <> 0 then script_array(current_script).show_script = TRUE

            If IsDate(script_array(current_script).retirement_date) = TRUE Then
                If DateDiff("d", date, script_array(current_script).retirement_date) =< 0 Then script_array(current_script).show_script = FALSE
            End If
			Call script_array(current_script).show_button(see_the_button)
			If see_the_button = FALSE Then script_array(current_script).show_script = FALSE

            If script_array(current_script).show_script = TRUE Then dlg_len = dlg_len + 15
        End if
    next

    dialog1 = ""
	BeginDialog dialog1, 0, 0, 600, dlg_len, script_category & " scripts main menu dialog"
	 	Text 5, 5, 435, 10, script_category & " scripts main menu: select the script to run from the choices below."
	  	ButtonGroup ButtonPressed


		'SUBCATEGORY HANDLING--------------------------------------------

		subcat_button_position = 5

		For i = 0 to ubound(subcategory_array)

			'Displays the button and text description-----------------------------------------------------------------------------------------------------------------------------
			'FUNCTION		HORIZ. ITEM POSITION	VERT. ITEM POSITION		ITEM WIDTH	ITEM HEIGHT		ITEM TEXT/LABEL										BUTTON VARIABLE
			PushButton 		subcat_button_position, 20, 					50, 		15, 			subcategory_array(i).subcat_name, 					subcat_button_placeholder

			subcategory_array(i).subcat_button = subcat_button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
			subcat_button_position = subcat_button_position + 50
			subcat_button_placeholder = subcat_button_placeholder + 1
		Next


		'SCRIPT LIST HANDLING--------------------------------------------

		'' 	PushButton 445, 10, 65, 10, "SIR instructions", 	SIR_instructions_button
		'This starts here, but it shouldn't end here :)
		vert_button_position = 50

		For current_script = 0 to ubound(script_array)


            If script_array(current_script).show_script = TRUE Then

				SIR_button_placeholder = button_placeholder + 1	'We always want this to be one more than the button_placeholder

				'Displays the button and text description-----------------------------------------------------------------------------------------------------------------------------
				'FUNCTION		HORIZ. ITEM POSITION	VERT. ITEM POSITION		ITEM WIDTH	ITEM HEIGHT		ITEM TEXT/LABEL										BUTTON VARIABLE
				PushButton 		5, 						vert_button_position, 	10, 		12, 			"?", 												SIR_button_placeholder
				PushButton 		18,						vert_button_position, 	120, 		12, 			script_array(current_script).script_name, 			button_placeholder
				Text 			120 + 23, 				vert_button_position+1, 500, 		14, 			"--- " & script_array(current_script).description
				'----------
				vert_button_position = vert_button_position + 15	'Needs to increment the vert_button_position by 15px (used by both the text and buttons)
				'----------
				script_array(current_script).button = button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
				script_array(current_script).SIR_instructions_button = SIR_button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
				button_placeholder = button_placeholder + 2
			End if

		next

		CancelButton 540, dlg_len - 20, 50, 15
	EndDialog
End function

'Starting these with a very high number, higher than the normal possible amount of buttons.
'	We're doing this because we want to assign a value to each button pressed, and we want
'	that value to change with each button. The button_placeholder will be placed in the .button
'	property for each script item. This allows it to both escape the Function and resize
'	near infinitely. We use dummy numbers for the other selector buttons for much the same reason,
'	to force the value of ButtonPressed to hold in near infinite iterations.
button_placeholder 			= 24601
subcat_button_placeholder 	= 1701

'Other pre-loop and pre-function declarations
subcategory_array = array()
subcategory_string = ""
subcategory_selected = "MAIN"

'Displays the dialog
Do

	'Creates the dialog
	call declare_main_menu_dialog("Actions")

	'At the beginning of the loop, we are not ready to exit it. Conditions later on will impact this.
	ready_to_exit_loop = false

	'Displays dialog, if cancel is pressed then stopscript
	dialog
	If ButtonPressed = 0 then stopscript

	'Determines the subcategory if a subcategory button was selected.
	For i = 0 to ubound(subcategory_array)
		If ButtonPressed = subcategory_array(i).subcat_button then subcategory_selected = subcategory_array(i).subcat_name
	Next

	'Runs through each script in the array... if the user selected script instructions (via ButtonPressed) it'll open_URL_in_browser to those instructions
	For i = 0 to ubound(script_array)
		If ButtonPressed = script_array(i).SIR_instructions_button then call open_URL_in_browser(script_array(i).SharePoint_instructions_URL)
	Next

	'Runs through each script in the array... if the user selected the actual script (via ButtonPressed), it'll run_from_GitHub
	For i = 0 to ubound(script_array)
		If ButtonPressed = script_array(i).button then
			ready_to_exit_loop = true		'Doing this just in case a stopscript or script_end_procedure is missing from the script in question
			script_to_run = script_array(i).script_URL
			Exit for
		End if
	Next


Loop until ready_to_exit_loop = true

call run_from_GitHub(script_to_run)

stopscript
