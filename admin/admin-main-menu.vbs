'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - MAIN MENU.vbs"
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
call changelog_update("11/01/2020", "Retired ADMIN - INTERVIEW REQUIRED. ADMIN - REVIEW REPORT replaces this.", "Ilse Ferris, Hennepin County")
call changelog_update("07/24/2020", "Removed the script 'Update Check Dates' from the ADMIN menu, it is currently available in the 'UTILITIES' Menu.", "Casey Love, Hennepin County")
call changelog_update("01/27/2020", "Added new HSR's to the QI script access menu. Welcome Keith, Kerry & Tanya!", "Ilse Ferris, Hennepin County")
call changelog_update("12/09/2019", "Added Jacob to the QI script access menu. Welcome Jacob!", "Ilse Ferris, Hennepin County")
call changelog_update("10/05/2019", "Remove CA Application Received.", "MiKayla Handley, Hennepin County")
call changelog_update("08/06/2019", "Added a new script to create an Excel List of MAXIS User detail.", "Casey Love, Hennepin County")
call changelog_update("04/12/2019", "Updated backend fuctionality. If you are on the QI team, and cannot access the QI scripts, contact me right away.", "Ilse Ferris, Hennepin County")
call changelog_update("06/21/2018", "Added QI specific scripts and sub menu.", "Ilse Ferris, Hennepin County")
call changelog_update("03/12/2018", "Added ODW Application.", "MiKayla Handley, Hennepin County")
call changelog_update("03/12/2018", "Removed DAIL report. BULK DAIL REPORT was updated in its place.", "Ilse Ferris, Hennepin County")
call changelog_update("11/30/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

tester_found = FALSE
qi_staff = FALSE
bz_staff = FALSE
For each tester in tester_array
    If user_ID_for_validation = tester.tester_id_number Then
        tester_found = TRUE
        worker_full_name            = tester.tester_full_name
        worker_first_name           = tester.tester_first_name
        worker_last_name            = tester.tester_last_name
        worker_email                = tester.tester_email
        worker_id_number            = tester.tester_id_number
        worker_x_number             = tester.tester_x_number
        worker_supervisor           = tester.tester_supervisor_name
        worker_supervisor_email     = tester.tester_supervisor_email
        worker_population           = tester.tester_population
        worker_region               = tester.tester_region
        worker_test_groups          = tester.tester_groups
        worker_test_scripts         = tester.tester_scripts
        For each group in worker_test_groups
            If group = "QI" Then
                qi_staff = TRUE
            End If
            If group = "BZ" Then
                qi_staff = TRUE
                bz_staff = TRUE
            End If
        Next
    End If
Next


Function declare_main_menu_dialog(script_category)

	'Runs through each script in the array and generates a list of subcategories based on the category located in the function. Also modifies the script description if it's from the last two months, to include a "NEW!!!" notification.
	For current_script = 0 to ubound(script_array)
		'Adds a "NEW!!!" notification to the description if the script is from the last two months.
		If DateDiff("m", script_array(current_script).release_date, DateAdd("m", -2, date)) <= 0 then
			script_array(current_script).description = "NEW " & script_array(current_script).release_date & "!!! --- " & script_array(current_script).description
			script_array(current_script).release_date = "12/12/1999" 'backs this out and makes it really old so it doesn't repeat each time the dialog loops. This prevents NEW!!!... from showing multiple times in the description.
		End if

	Next

    ' MsgBox "In fn - " & ButtonPressed

    dlg_len = 60
    For current_script = 0 to ubound(script_array)
        script_array(current_script).show_script = FALSE
        If ucase(script_array(current_script).category) = ucase(script_category) then

            '<<<<<<RIGHT HERE IT SHOULD ITERATE THROUGH SUBCATEGORIES AND BUTTONS PRESSED TO DETERMINE WHAT THE CURRENTLY DISPLAYED SUBCATEGORY SHOULD BE, THEN ONLY DISPLAY SCRIPTS THAT MATCH THAT CRITERIA

            'Accounts for scripts without subcategories
            If subcategory_string = "" then subcategory_string = "MAIN"		'<<<THIS COULD BE A PROPERTY OF THE CLASS

            For each group in script_array(current_script).tags
                ' MsgBox script_array(current_script).script_name & vbNewLine & group
                If ButtonPressed = menu_admin_button Then
                    If group = "" Then script_array(current_script).show_script = TRUE
                    show_question_mark = TRUE
                ElseIf ButtonPressed = menu_QI_button Then
                    If group = "QI" Then script_array(current_script).show_script = TRUE
                    show_question_mark = FALSE
                ElseIf ButtonPressed = menu_BZ_button Then
                    If group = "BZ" Then script_array(current_script).show_script = TRUE
                    If group = "Monthly Tasks" Then script_array(current_script).show_script = FALSE
                    show_question_mark = FALSE
                ElseIf ButtonPressed = menu_monthly_tasks_button Then
                    If group = "Monthly Tasks" Then script_array(current_script).show_script = TRUE
                    show_question_mark = FALSE
                End If
                ' MsgBox script_array(current_script).show_script
            Next

            If IsDate(script_array(current_script).retirement_date) = TRUE Then
                If DateDiff("d", date, script_array(current_script).retirement_date) =< 0 Then script_array(current_script).show_script = FALSE
            End If

            If script_array(current_script).show_script = TRUE Then dlg_len = dlg_len + 15
        End if
    next


	BeginDialog dialog1, 0, 0, 600, dlg_len, script_category & " scripts main menu dialog"
	 	Text 5, 5, 435, 10, script_category & " scripts main menu: select the script to run from the choices below."
	  	ButtonGroup ButtonPressed


		'SUBCATEGORY HANDLING--------------------------------------------

            'Displays the button and text description-----------------------------------------------------------------------------------------------------------------------------
    		'FUNCTION		HORIZ. ITEM POSITION	VERT. ITEM POSITION		ITEM WIDTH	ITEM HEIGHT		ITEM TEXT/LABEL				BUTTON VARIABLE
        If qi_staff = TRUE OR bz_staff = TRUE Then
        	PushButton 		5,                      20, 					50, 		15, 			"ADMIN", 					menu_admin_button
        End If
        If qi_staff = TRUE Then
            PushButton 		65,                     20, 					40, 		15, 			"QI", 					    menu_QI_button
        End If
        If bz_staff = TRUE Then
            PushButton 		105,                    20, 					40, 		15, 			"BZ", 					    menu_BZ_button
            PushButton 		145,                    20, 					60, 		15, 			"Monthly Tasks", 			menu_monthly_tasks_button
        End If


		'SCRIPT LIST HANDLING--------------------------------------------

		'' 	PushButton 445, 10, 65, 10, "SIR instructions", 	SIR_instructions_button
		'This starts here, but it shouldn't end here :)
		vert_button_position = 50

		For current_script = 0 to ubound(script_array)


            If script_array(current_script).show_script = TRUE Then

				SIR_button_placeholder = button_placeholder + 1	'We always want this to be one more than the button_placeholder

				'Displays the button and text description-----------------------------------------------------------------------------------------------------------------------------
				'FUNCTION		HORIZ. ITEM POSITION	VERT. ITEM POSITION		ITEM WIDTH	ITEM HEIGHT		ITEM TEXT/LABEL										BUTTON VARIABLE
				If show_question_mark = TRUE Then
                PushButton 		5, 						vert_button_position, 	10, 		12, 			"?", 												SIR_button_placeholder
                End If
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
menu_admin_button           = 110
menu_QI_button              = 120
menu_BZ_button              = 130
menu_monthly_tasks_button   = 140

ButtonPressed = menu_admin_button
show_question_mark = TRUE

dialog1 = ""
'Displays the dialog
Do
    last_button = ButtonPressed
    ' MsgBox "Before - " & ButtonPressed

	'Creates the dialog
	call declare_main_menu_dialog("Admin")

	'At the beginning of the loop, we are not ready to exit it. Conditions later on will impact this.
	ready_to_exit_loop = false

	'Displays dialog, if cancel is pressed then stopscript
	dialog dialog1
	If ButtonPressed = 0 then stopscript

	'Runs through each script in the array... if the user selected script instructions (via ButtonPressed) it'll open_URL_in_browser to those instructions
	For i = 0 to ubound(script_array)
		If ButtonPressed = script_array(i).SIR_instructions_button then
            ' MsgBox script_array(i).SharePoint_instructions_URL
            call open_URL_in_browser(script_array(i).SharePoint_instructions_URL)
            ButtonPressed = last_button
        End If
	Next

	'Runs through each script in the array... if the user selected the actual script (via ButtonPressed), it'll run_from_GitHub
	For i = 0 to ubound(script_array)
		If ButtonPressed = script_array(i).button then
			ready_to_exit_loop = true		'Doing this just in case a stopscript or script_end_procedure is missing from the script in question
			script_to_run = script_array(i).script_URL
            ' MsgBox script_to_run
			Exit for
		End if
	Next

    ' MsgBox "After - " & ButtonPressed
    dialog1 = ""
Loop until ready_to_exit_loop = true

call run_from_GitHub(script_to_run)

stopscript
