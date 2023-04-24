'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "SHELTER - MAIN MENU.vbs"
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
call changelog_update("04/22/2022", "Retired SELF PAY, CASH CUTOFF, & MONEY MISMANAGEMENT scripts and removed all other mention of self pay from other scripts.", "MiKayla Handley, Hennepin County")
call changelog_update("06/04/2021", "Retired GRH APPROVAL and SINGLE CLIENT INTERVIEW scripts.", "Ilse Ferris, Hennepin County")
call changelog_update("07/05/2018", "Updates to add scripts per shelter team request..", "MiKayla Handley")
call changelog_update("01/05/2018", "Updates to CES-Screening Appt per shelter team request..", "MiKayla Handley")
call changelog_update("09/23/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

class subcat
	public subcat_name
	public subcat_button
End class

Function declare_main_menu_dialog(script_category)
	show_0_l_btn = True
	show_m_z_btn = True

	dlg_len = 60
    For current_script = 0 to ubound(script_array)
        script_array(current_script).show_script = FALSE
        If ucase(script_array(current_script).category) = ucase(script_category) then
            If ButtonPressed = menu_0_to_L_button Then
                If IsNumeric(left(script_array(current_script).script_name, 1)) = TRUE Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "A" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "B" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "C" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "D" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "E" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "F" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "G" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "H" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "I" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "J" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "K" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "L" Then script_array(current_script).show_script = TRUE
				show_0_l_btn = False
            ElseIf ButtonPressed = menu_M_to_Z_button Then
                If left(script_array(current_script).script_name, 1) = "M" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "N" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "O" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "P" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "Q" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "R" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "S" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "T" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "U" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "V" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "W" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "X" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "Y" Then script_array(current_script).show_script = TRUE
                If left(script_array(current_script).script_name, 1) = "Z" Then script_array(current_script).show_script = TRUE
				show_m_z_btn = False
			End If
            If IsDate(script_array(current_script).retirement_date) = TRUE Then
                If DateDiff("d", date, script_array(current_script).retirement_date) =< 0 Then script_array(current_script).show_script = FALSE
            End If
			Call script_array(current_script).show_button(see_the_button)
			If see_the_button = FALSE Then script_array(current_script).show_script = FALSE

            If script_array(current_script).show_script = TRUE Then dlg_len = dlg_len + 15
        End If
    next

	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 600, dlg_len, script_category & " scripts main menu dialog"
	 	Text 5, 5, 435, 10, script_category & " scripts main menu: select the script to run from the choices below."
		EditBox 700, 700, 50, 15, holderbox				'This sits here as the first control element so the fisrt button listed doesn't have the blue box around it.
	  	ButtonGroup ButtonPressed

		'SUBCATEGORY HANDLING--------------------------------------------

		'Displays the button and text description-----------------------------------------------------------------------------------------------------------------------------
			'FUNCTION		HORIZ. ITEM POSITION	VERT. ITEM POSITION		ITEM WIDTH	ITEM HEIGHT		ITEM TEXT/LABEL				BUTTON VARIABLE
		If show_0_l_btn = True  Then
			PushButton 		5,                      20, 					50, 		15, 			" # - L ", 					menu_0_to_L_button
		Else
			Text 			20,                      23, 					40, 		15, 			" # - L "
		End If
		If show_m_z_btn = True  Then
        	PushButton 		55,                     20, 					50, 		15, 			" M - Z ", 					menu_M_to_Z_button
		Else
			Text 			70,                     23, 					40, 		15, 			" M - Z "
		End If

		'SCRIPT LIST HANDLING--------------------------------------------

		'This starts here, but it shouldn't end here :)
		vert_button_position = 50

		For current_script = 0 to ubound(script_array)


            If script_array(current_script).show_script = TRUE Then

				SIR_button_placeholder = button_placeholder + 1	'We always want this to be one more than the button_placeholder

				'Displays the button and text description-----------------------------------------------------------------------------------------------------------------------------
				'FUNCTION		HORIZ. ITEM POSITION	VERT. ITEM POSITION		ITEM WIDTH	ITEM HEIGHT		ITEM TEXT/LABEL										BUTTON VARIABLE
				' PushButton 		5, 						vert_button_position, 	10, 		12, 			"?", 												SIR_button_placeholder
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
button_placeholder 		= 24601
menu_0_to_L_button      = 110
menu_M_to_Z_button      = 120

subcategory_selected = "# - L"

'Displays the dialog
ButtonPressed = menu_0_to_L_button
Do
	'Creates the dialog
	call declare_main_menu_dialog("SHELTER")

	'At the beginning of the loop, we are not ready to exit it. Conditions later on will impact this.
	ready_to_exit_loop = false

	'Displays dialog, if cancel is pressed then stopscript
	dialog Dialog1
	If ButtonPressed = 0 then stopscript

	'Runs through each script in the array... if the user selected script instructions (via ButtonPressed) it'll open_URL_in_browser to those instructions
	For i = 0 to ubound(script_array)
		If ButtonPressed = script_array(i).SIR_instructions_button then
            ' MsgBox script_array(i).SharePoint_instructions_URL
            call open_URL_in_browser(script_array(i).SharePoint_instructions_URL)
        End If
	Next

	'Runs through each script in the array... if the user selected the actual script (via ButtonPressed), it'll run_from_GitHub
	For i = 0 to ubound(script_array)
		If ButtonPressed = script_array(i).button then
			ready_to_exit_loop = true		'Doing this just in case a stopscript or script_end_procedure is missing from the script in question
			script_to_run = script_array(i).script_URL
			Exit for
		End if
	Next

    ' MsgBox script_to_run
Loop until ready_to_exit_loop = true

call run_from_GitHub(script_to_run)

stopscript
