'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "TEST - MAIN MENU.vbs"
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

'----------------------------------------------------------------------------------------------------This is the list of scripts that are held locally
' IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
' 	IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
' 		script_list_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/COMPLETE%20LIST%20OF%20SCRIPTS.vbs"
' 	Else											'Everyone else should use the release branch.
' 		script_list_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/COMPLETE%20LIST%20OF%20SCRIPTS.vbs"
' 	End if
'
' 	SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a script_list_URL
' 	req.open "GET", script_list_URL, FALSE							'Attempts to open the script_list_URL
' 	req.send													'Sends request
' 	IF req.Status = 200 THEN									'200 means great success
' 		Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
' 		Execute req.responseText								'Executes the script code
' 	ELSE														'Error message
' 		critical_error_msgbox = MsgBox ("Something has gone wrong. The script list code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
'                                         "Script list URL: " & script_list_URL & vbNewLine & vbNewLine &_
'                                         "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
'                                         vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
'         StopScript
' 	END IF
' ELSE
' 	script_list_URL = "C:\MAXIS-scripts\COMPLETE LIST OF SCRIPTS.vbs"
' 	Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
' 	Set fso_command = run_another_script_fso.OpenTextFile(script_list_URL)
' 	text_from_the_other_script = fso_command.ReadAll
' 	fso_command.Close
' 	Execute text_from_the_other_script
' End if

' script_list_URL = "C:\MAXIS-scripts\Test scripts\Casey\Tabs\COMPLETE LIST OF SCRIPTS.vbs"
' If run_locally = TRUE Then
'
' Else
'
' End If

' tags_coming_soon = MsgBox("***            Coming soon!            ***" & vbNewLine & vbNewLine & "We are updating how we engage with the script tools. One of these ways is with a new system of tagging." & vbNewLine & "This button will have functionality to preview the new menu using these tags. It is not available just yet as we develop and test the functionality." & vbNewLine & vbNewLine & "Come back later to try this new functionality.", vbOk, "New Tags Menu Coming Soon.")
' script_end_procedure("")


' script_repository = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/"
' script_list_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/COMPLETE%20LIST%20OF%20SCRIPTS.vbs"
' script_list_URL = "C:\MAXIS-scripts\COMPLETE LIST OF SCRIPTS.vbs"


testers_script_list_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\COMPLETE LIST OF TESTERS.vbs"        'Opening the list of testers - which is saved locally for security
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile(testers_script_list_URL)
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

Set objNet = CreateObject("WScript.NetWork")
windows_user_ID = objNet.UserName
user_ID_for_validation = ucase(windows_user_ID)

tester_found = FALSE
For each tester in tester_array
    If user_ID_for_validation = tester.tester_id_number Then
        tester_found = TRUE
    End If
Next

If tester_found = FALSE Then
    tags_coming_soon = MsgBox("***            Coming soon!            ***" & vbNewLine & vbNewLine & "We are updating how we engage with the script tools. One of these ways is with a new system of tagging." & vbNewLine & "This button will have functionality to preview the new menu using these tags. It is not available just yet as we develop and test the functionality." & vbNewLine & vbNewLine & "Come back later to try this new functionality.", vbOk, "New Tags Menu Coming Soon.")
    stopscript
End If

If script_repository = "" Then script_repository = "C:\MAXIS-scripts\"
script_list_URL = script_repository & "/COMPLETE LIST OF SCRIPTS.vbs"
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile(script_list_URL)
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("10/22/2019", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
' Dialog dialog_test
class subcat
	public subcat_name
	public subcat_button
End class

function declare_tabbed_menu(tab_selected)

        Dialog1 = ""
        tab_selected = trim(tab_selected)
        If right(tab_selected, 1) = "~" Then tab_selected = left(tab_selected, len(tab_selected) - 1)
        If left(tab_selected, 1) = "~" Then tab_selected = right(tab_selected, len(tab_selected) - 1)
        tags_array = split(tab_selected, "~")
        scripts_included = 0

        dlg_len = 80
        show_dail_scrubber = FALSE
        'Runs through each script in the array and generates a list of subcategories based on the category located in the function. Also modifies the script description if it's from the last two months, to include a "NEW!!!" notification.
        For current_script = 0 to ubound(script_array)
            script_array(current_script).show_script = TRUE
            If tab_selected <> "" Then
                If script_array(current_script).show_script = TRUE Then
                    For each selected_tag in tags_array
                        If selected_tag <> "" Then
                            ' MsgBox script_array(current_script).script_name & vbNewLine' & script_array(current_script).tags
                            For each listed_tag in script_array(current_script).tags
                                If listed_tag <> "" Then
                                    tag_matched = FALSE

                                    ' MsgBox "selected tag - " & selected_tag & vbNewLine & "listed tag - " & listed_tag

                                    If UCase(selected_tag) = UCase(listed_tag) Then
                                        tag_matched = TRUE
                                        ' MsgBox "selected tag - " & selected_tag & vbNewLine & "listed tag - " & listed_tag & vbNewLine & "tag matched - " & tag_matched & vbNewLine & script_array(current_script).script_name & vbNewLine & "list this script - " & list_this_script
                                        Exit For
                                    End If
                                Else
                                    script_array(current_script).show_script = FALSE
                                End If
                                ' MsgBox "Tag matched - " & tag_matched
                            Next
                            If tag_matched = FALSE Then script_array(current_script).show_script = FALSE
                            ' If tag_matched = TRUE Then MsgBox script_array(current_script).script_name & vbNewLine & "list this script - " & list_this_script
                        Else
                            script_array(current_script).show_script = FALSE
                        End If
                    Next
                End If
                If script_array(current_script).show_script = TRUE Then
                    Call script_array(current_script).show_button(use_this_button)
                    If use_this_button = FALSE Then script_array(current_script).show_script = FALSE
                    If script_array(current_script).in_testing = TRUE Then script_array(current_script).description = "IN TESTING - " & script_array(current_script).description
                    ' MsgBox script_array(current_script).script_name & vbNewLine & "Use this button - " & use_this_button & vbNewLine & "show script - " & script_array(current_script).show_script
                End If
            Else
                script_array(current_script).show_script = FALSE
            End If

            If script_array(current_script).show_script = TRUE Then
                dlg_len = dlg_len + 15
                scripts_included = scripts_included + 1
                ' MsgBox "script - " & script_array(current_script).script_name & vbNewLine & "tags - " & Join(script_array(current_script).tags, ", ")
                If script_array(current_script).category = "DAIL" Then
                    script_array(current_script).show_script = FALSE
                    show_dail_scrubber = TRUE
                    dlg_len = dlg_len - 15
                    dail_scrubber_functionality = dail_scrubber_functionality & " : " & script_array(current_script).script_name
                    scripts_included = scripts_included - 1
                End If
            End If
            ' 'Subcategory handling (creating a second list as a string which gets converted later to an array)
            ' If ucase(script_array(current_script).category) = ucase(script_category) then																								'If the script in the array is of the correct category (ACTIONS/NOTES/ETC)...
            '     For each listed_subcategory in script_array(current_script).subcategory																									'...then iterate through each listed subcategory, and...
            '         If listed_subcategory <> "" and InStr(subcategory_list, ucase(listed_subcategory)) = 0 then subcategory_list = subcategory_list & "|" & ucase(listed_subcategory)	'...if the listed subcategory isn't blank and isn't already in the list, then add it to our handy-dandy list.
            '     Next
            ' End if
            'Adds a "NEW!!!" notification to the description if the script is from the last two months.
            If DateDiff("m", script_array(current_script).release_date, DateAdd("m", -2, date)) <= 0 then
                script_array(current_script).description = "NEW " & script_array(current_script).release_date & "!!! --- " & script_array(current_script).description
                script_array(current_script).release_date = "12/12/1999" 'backs this out and makes it really old so it doesn't repeat each time the dialog loops. This prevents NEW!!!... from showing multiple times in the description.
            End if

        Next
        If show_dail_scrubber = TRUE Then dlg_len = dlg_len + 15
        If current_page = "" Then current_page = "One"
        If current_page = "One" AND scripts_included > 20 Then dlg_len = 385
        If current_page = "Two" Then
            dlg_len = 80 + 15 * (scripts_included - 20)
            If scripts_included > 40 Then dlg_len = 385
        End If
        If current_page = "Three" Then
            dlg_len = 80 + 15 * (scripts_included - 40)
        End If
        If show_dail_scrubber = TRUE then dlg_len = dlg_len + 10
        dail_scrubber_functionality = trim(dail_scrubber_functionality)
        If dail_scrubber_functionality <> "" Then dail_scrubber_functionality = right(dail_scrubber_functionality, len(dail_scrubber_functionality)  - 2)
        If dlg_len < 240 Then dlg_len = 240

        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 750, dlg_len, "Select Script to Run"
          GroupBox 550, 5, 185, 40, "Selected TAGS"
          Text 555, 15, 175, 25, Join(tags_array, ", ")
          ' Text 650, 10, 50, 10, "Keywords:"
          ' EditBox 650, 20, 50, 15, search_keywords
          GroupBox 630, 50, 105, 125, "Key Codes"
          Text 635, 60, 20, 10, "Cn  - "
          Text 655, 60, 70, 10, "Case Notes"
          Text 635, 70, 20, 10, "Ex  - "
          Text 655, 70, 70, 10, "Excel"
          Text 635, 80, 20, 10, "Exp -"
          Text 655, 80, 70, 10, "Expedited SNAP"
          Text 635, 90, 20, 10, "Fi  - "
          Text 655, 90, 70, 10, "FIATs"
          Text 635, 100, 20, 10, "Oa  -"
          Text 655, 100, 70, 10, "Outlook Appointment"
          Text 635, 110, 20, 10, "Oe  -"
          Text 655, 110, 70, 10, "Outlook Email"
          Text 635, 120, 20, 10, "Sm  -"
          Text 655, 120, 70, 20, "SPEC/ MEMO"
          Text 635, 130, 20, 10, "Sw  -"
          Text 655, 130, 70, 10, "SPEC/ WCOM"
          Text 635, 140, 20, 10, "Tk  - "
          Text 655, 140, 70, 10, "TIKL"
          Text 635, 150, 20, 10, "Up  - "
          Text 655, 150, 70, 10, "Updates Panel"
          Text 635, 160, 20, 10, "Wrd - "
          Text 655, 160, 70, 10, "Word"

          GroupBox 630, 175, 105, 40, "Button Functions"
          ' Text 635, 185, 20, 10, "?"
          Text 655, 190, 70, 10, "View Instructions"
          ' Text 635, 195, 20, 10, "+"
          Text 655, 200, 70, 10, "Add to Favorites"

          ButtonGroup ButtonPressed
            PushButton 5, 10, 60, 15, "SNAP", snap_btn
            PushButton 65, 10, 60, 15, "MFIP", mfip_btn
            PushButton 125, 10, 60, 15, "DWP", dwp_btn
            PushButton 185, 10, 60, 15, "Health Care", hc_btn
            PushButton 245, 10, 60, 15, "HS/GRH", grh_btn
            PushButton 305, 10, 60, 15, "Adult Cash", ga_btn
            PushButton 365, 10, 60, 15, "EMER", emer_btn
            PushButton 425, 10, 60, 15, "LTC", ltc_btn
            PushButton 485, 10, 60, 15, "ABAWD", abawd_btn
            PushButton 5, 25, 60, 15, "Income", income_btn
            PushButton 65, 25, 60, 15, "Assets", asset_btn
            PushButton 125, 25, 60, 15, "Deductions", deductions_btn
            PushButton 185, 25, 60, 15, "Applications", application_btn
            PushButton 245, 25, 60, 15, "Reviews", review_btn
            PushButton 305, 25, 60, 15, "Communication", communication_btn
            PushButton 365, 25, 60, 15, "Utility", utility_btn
            PushButton 425, 25, 60, 15, "Reports", reports_btn
            PushButton 485, 25, 60, 15, "Resources", resources_btn
            ' PushButton 15, 60, 10, 10, "?", instructions_btn
            ' PushButton 30, 60, 95, 10, "script name", name_btn
            PushButton 640, 190, 10, 10, "?", explain_questionmark_btn
            PushButton 640, 200, 10, 10, "+", explain_plus_btn


            vert_button_position = 50
            list_counter = 0
            For current_script = 0 to ubound(script_array)
                If script_array(current_script).show_script = TRUE Then
                    show_this_one = FALSE
                    If current_page = "One" AND list_counter < 20 Then show_this_one = TRUE
                    If current_page = "Two" AND list_counter >= 20 AND list_counter < 40 Then show_this_one = TRUE
                    If current_page = "Three" AND list_counter >= 40 Then show_this_one = TRUE
                    ' MsgBox "Current page - " & current_page & vbNewLine & "list_counter - " & list_counter & vbNewLine & "show this one - " & show_this_one
                    If show_this_one = TRUE Then
                    ' If tab_selected <> "" Then
                    '     For each listed_tag in script_array(current_script).tags
                    '         If listed_tag <> "" Then
                    '             If UCase(listed_tag) = UCase(tab_selected) then
                                    SIR_button_placeholder = button_placeholder + 1	'We always want this to be one more than the button_placeholder
                                    add_to_favorites_button_placeholder = button_placeholder + 2
                                    script_keys_combine = ""
                                    If script_array(current_script).dlg_keys(0) <> "" Then script_keys_combine = Join(script_array(current_script).dlg_keys, ":")

                                    'Displays the button and text description-----------------------------------------------------------------------------------------------------------------------------
                                    'FUNCTION		HORIZ. ITEM POSITION	VERT. ITEM POSITION		ITEM WIDTH	ITEM HEIGHT		ITEM TEXT/LABEL										BUTTON VARIABLE
                                    PushButton 		5, 						vert_button_position, 	10, 		10, 			"?", 												SIR_button_placeholder
                                    PushButton 		18,						vert_button_position, 	120, 		10, 			script_array(current_script).script_name, 			button_placeholder
                                    PushButton      140,                    vert_button_position,   10,         10,             "+",                                                add_to_favorites_button_placeholder
                                    Text 			150, 				    vert_button_position, 	65, 		10, 			"-- " & script_keys_combine & " --"
                                    Text            210,                    vert_button_position,   425,        10,             script_array(current_script).description
                                    '----------
                                    vert_button_position = vert_button_position + 15	'Needs to increment the vert_button_position by 15px (used by both the text and buttons)
                                    '----------
                                    script_array(current_script).button = button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
                                    script_array(current_script).SIR_instructions_button = SIR_button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
                                    script_array(current_script).fav_add_button = add_to_favorites_button_placeholder
                                    button_placeholder = button_placeholder + 3

                    '             End If
                    '         End If
                    '     Next
                    ' End If
                    End If
                    list_counter = list_counter + 1
                End If
            Next
            If show_dail_scrubber = TRUE Then
                PushButton 5, vert_button_position, 10, 15, "?", dail_scrubber_instructions_btn
                PushButton 18, vert_button_position, 120, 15, "DAIL Scrubber", dail_scrubber_script_button
                ' Text 143, vert_button_position, 40, 10, dail_keys
                Text 143, vert_button_position, 500, 20, dail_scrubber_functionality
                vert_button_position = vert_button_position + 20
            End If
            vert_button_position = vert_button_position + 5
            If scripts_included > 20 Then
                Text 520, vert_button_position + 5, 20, 10, "Page:"
                If current_page = "One" AND scripts_included > 20 Then
                    Text 545, vert_button_position + 5, 10, 10, "1"
                    PushButton 550, vert_button_position + 5, 10, 10, "2", page_two_btn
                    If scripts_included > 40 Then PushButton 560, vert_button_position + 5, 10, 10, "3", page_three_btn
                ElseIf current_page = "Two" AND scripts_included > 20 Then
                    PushButton 540, vert_button_position + 5, 10, 10, "1", page_one_btn
                    Text 555, vert_button_position + 5, 5, 10, "2"
                    If scripts_included > 40 Then PushButton 560, vert_button_position + 5, 10, 10, "3", page_three_btn
                ElseIf current_page = "Three" AND scripts_included > 40 Then
                    PushButton 540, vert_button_position + 5, 10, 10, "1", page_one_btn
                    PushButton 550, vert_button_position + 5, 10, 10, "2", page_two_btn
                    Text 565, vert_button_position + 5, 5, 10, "3"
                End If
            End If
          '   If vert_button_position < 200 Then vert_button_position = 200
          '   PushButton 10, vert_button_position, 100, 15, "Clear TAG Selection", clear_selection_btn
          '   CancelButton 690, vert_button_position, 50, 15
          ' ' Text 120, vert_button_position, 380, 15, "C - Case Notes ... E - Excel ... EXP - Expedited SNAP ... F - FIATs ... OA - Outlook Appointment ... OE - Outlook Email ... SM - SPEC/MEMO ... SW - SPEC/WCOM ... T - TIKL ... U - Updates Panel ... W - Word"
          ' Text 120, vert_button_position + 5, 40, 10, "Keywords:"
          ' EditBox 165, vert_button_position, 200, 15, search_keywords
          ' Text 130, 60, 30, 10, "- KEYS -"
          ' Text 170, 60, 170, 10, "Description"p

            PushButton 10, dlg_len - 20, 100, 15, "Clear TAG Selection", clear_selection_btn
            CancelButton 690, dlg_len - 20, 50, 15
            ' Text 120, vert_button_position, 380, 15, "C - Case Notes ... E - Excel ... EXP - Expedited SNAP ... F - FIATs ... OA - Outlook Appointment ... OE - Outlook Email ... SM - SPEC/MEMO ... SW - SPEC/WCOM ... T - TIKL ... U - Updates Panel ... W - Word"
          ' Text 120, dlg_len - 15, 40, 10, "Keywords:"                 'commented out because we don't have keywordds
          ' EditBox 165, dlg_len - 20, 200, 15, search_keywords
        EndDialog


end function

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
select_tab = ""
Dim snap_btn, mfip_btn, dwp_btn, hc_btn, grh_btn, ga_btn, emer_btn, ltc_btn, abawd_btn, income_btn, asset_btn, deductions_btn, application_btn, review_btn, communication_btn, utility_btn, reports_btn, resources_btn, clear_selection_btn
Dim page_one_btn, page_two_btn, page_three_btn, current_page

Do
    leave_loop = ""

    Call declare_tabbed_menu(select_tab)
    dialog Dialog1

    cancel_without_confirmation

    If ButtonPressed = explain_questionmark_btn Then
        explain_questionmark_msg = MsgBox("See all the Question Mark Buttons?" & vbCr & vbCr & "Look to the left of the script name button. Each script has a button with a question mark - ? - next to it." & vbCr & vbCr & "Press this button and the instructions for that script will be opened. This is an easy way to see how a script functions, or when to use it.", vbQuestion, "What's with the Question Marks?")
    End If
    If ButtonPressed = explain_plus_btn Then
        explain_plus_msg = MsgBox("What are these plus buttons?" & vbCr & vbCr & "Just to the right of the script name button is a button with a plus sign (+)." & vbCr & vbCr & "This button will add this script to your list of favorite scripts for access inthe favorites menu.", vbQuestion, "What are the Plus Signs for?")
    End If


    If ButtonPressed = clear_selection_btn Then
        leave_loop = FALSE
        select_tab = ""
        current_page = "One"
    End If

    If ButtonPressed = page_one_btn Then
        leave_loop = FALSE
        current_page = "One"
    End If
    If ButtonPressed = page_two_btn Then
        leave_loop = FALSE
        current_page = "Two"
    End If
    If ButtonPressed = page_three_btn Then
        leave_loop = FALSE
        current_page = "Three"
    End If
    If ButtonPressed = dail_scrubber_instructions_btn Then
        Call open_URL_in_browser("https://dept.hennepin.us/hsphd/sa/ews/BlueZone_Script_Instructions/DAIL/ALL%20DAIL%20SCRIPTS.docx")
        leave_loop = FALSE
    End If

    If ButtonPressed = snap_btn OR ButtonPressed = mfip_btn OR ButtonPressed = dwp_btn OR ButtonPressed = hc_btn OR ButtonPressed = grh_btn OR ButtonPressed = ga_btn OR ButtonPressed = emer_btn OR ButtonPressed = ltc_btn OR ButtonPressed = abawd_btn OR ButtonPressed = income_btn OR ButtonPressed = asset_btn OR ButtonPressed = deductions_btn OR ButtonPressed = application_btn OR ButtonPressed = review_btn OR ButtonPressed = communication_btn OR ButtonPressed = utility_btn OR ButtonPressed = reports_btn OR ButtonPressed = resources_btn Then

        leave_loop = FALSE
        current_page = "One"
        If ButtonPressed = snap_btn Then select_tab = select_tab & "~" & "SNAP"
        If ButtonPressed = mfip_btn Then select_tab = select_tab & "~" & "MFIP"
        If ButtonPressed = dwp_btn Then select_tab = select_tab & "~" & "DWP"
        If ButtonPressed = hc_btn Then select_tab = select_tab & "~" & "Health Care"
        If ButtonPressed = grh_btn Then select_tab = select_tab & "~" & "HS/GRH"
        If ButtonPressed = ga_btn Then select_tab = select_tab & "~" & "Adult Cash"
        If ButtonPressed = emer_btn Then select_tab = select_tab & "~" & "EMER"
        If ButtonPressed = ltc_btn Then select_tab = select_tab & "~" & "LTC"
        If ButtonPressed = abawd_btn Then select_tab = select_tab & "~" & "ABAWD"
        If ButtonPressed = income_btn Then select_tab = select_tab & "~" & "Income"
        If ButtonPressed = asset_btn Then select_tab = select_tab & "~" & "Assets"
        If ButtonPressed = deductions_btn Then select_tab = select_tab & "~" & "Deductions"
        If ButtonPressed = application_btn Then select_tab = select_tab & "~" & "Application"
        If ButtonPressed = review_btn Then select_tab = select_tab & "~" & "Reviews"
        If ButtonPressed = communication_btn Then select_tab = select_tab & "~" & "Communication"
        If ButtonPressed = utility_btn Then select_tab = select_tab & "~" & "Utility"
        If ButtonPressed = reports_btn Then select_tab = select_tab & "~" & "Reports"
        If ButtonPressed = resources_btn Then select_tab = select_tab & "~" & "Resources"
    End If

    For i = 0 to ubound(script_array)
		If ButtonPressed = script_array(i).SIR_instructions_button then
            call open_URL_in_browser(script_array(i).SharePoint_instructions_URL)
            leave_loop = FALSE
        End If
        If ButtonPressed = script_array(i).fav_add_button then
            ' MsgBox "Script in favorites - " & script_array(i).script_in_favorites
            If script_array(i).script_in_favorites = TRUE Then
                MsgBox "The script " & script_array(i).category & "-" & script_array(i).script_name & " is already listed in favorites."
            Else
                new_favorite = script_array(i).category & "/" & script_array(i).script_name
                If all_favorites = "" Then
                    all_favorites = join(favorites_text_file_array, vbNewLine)
                End If
                all_favorites = all_favorites & vbNewLine & new_favorite

                SET updated_fav_scripts_fso = CreateObject("Scripting.FileSystemObject")
                SET updated_fav_scripts_command = updated_fav_scripts_fso.CreateTextFile(favorites_text_file_location, 2)
                updated_fav_scripts_command.Write(all_favorites)
                updated_fav_scripts_command.Close

                MsgBox "The script " & script_array(i).category & "-" & script_array(i).script_name & " has been added to your list of favorites."

            End If
        End If
	Next

	'Runs through each script in the array... if the user selected the actual script (via ButtonPressed), it'll run_from_GitHub
	For i = 0 to ubound(script_array)
		If ButtonPressed = script_array(i).button then
			leave_loop = true		'Doing this just in case a stopscript or script_end_procedure is missing from the script in question
			script_to_run = script_array(i).script_URL
			Exit for
		End if
	Next

    If ButtonPressed = dail_scrubber_script_button Then
        leave_loop = TRUE
        script_to_run = script_array(0).script_URL
    End If
Loop Until leave_loop = TRUE

call run_from_GitHub(script_to_run)

stopscript
