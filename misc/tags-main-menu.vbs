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

'This block is to lock out non-testers from using TAGS.
testers_script_list_URL = t_drive & "\Eligibility Support\Scripts\Script Files\COMPLETE LIST OF TESTERS.vbs"        'Opening the list of testers - which is saved locally for security
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile(testers_script_list_URL)
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

Set objNet = CreateObject("WScript.NetWork")
windows_user_ID = objNet.UserName
user_ID_for_validation = ucase(windows_user_ID)

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

If tester_found = FALSE Then
    tags_coming_soon = MsgBox("***            Coming soon!            ***" & vbNewLine & vbNewLine & "We are updating how we engage with the script tools. One of these ways is with a new system of tagging." & vbNewLine & "This button will have functionality to preview the new menu using these tags. It is not available just yet as we develop and test the functionality." & vbNewLine & vbNewLine & "Come back later to try this new functionality.", vbOk, "New Tags Menu Coming Soon.")
    stopscript
End If

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
        dlg_len = 80
        scripts_included = 0

        tab_selected = trim(tab_selected)
        If right(tab_selected, 1) = "~" Then tab_selected = left(tab_selected, len(tab_selected) - 1)
        If left(tab_selected, 1) = "~" Then tab_selected = right(tab_selected, len(tab_selected) - 1)
        tags_array = split(tab_selected, "~")
        one_month_ago = DateAdd("m", -1, date)
        two_months_ago = DateAdd("m", -2, date)

        show_dail_scrubber = FALSE
        new_script_to_list = FALSE
        hot_topic_script_to_list = FALSE
        'Runs through each script in the array and generates a list of subcategories based on the category located in the function. Also modifies the script description if it's from the last two months, to include a "NEW!!!" notification.
        For current_script = 0 to ubound(script_array)
            script_array(current_script).show_script = TRUE
            If show_resources = FALSE AND qi_menu = FALSE and bz_menu = FALSE and task_menu = FALSE Then
                If tab_selected <> "" Then
                    If script_array(current_script).show_script = TRUE Then
                        For each selected_tag in tags_array
                            If selected_tag <> "" Then
                                ' MsgBox script_array(current_script).script_name & vbNewLine' & script_array(current_script).tags
                                For each listed_tag in script_array(current_script).tags
                                    If listed_tag <> "" Then
                                        tag_matched = FALSE

                                        ' If listed_tag = "Support" Then MsgBox "selected tag - " & selected_tag & vbNewLine & "listed tag - " & listed_tag

                                        If UCase(selected_tag) = UCase(listed_tag) Then
                                            tag_matched = TRUE
                                            ' If listed_tag = "Support" Then MsgBox "selected tag - " & selected_tag & vbNewLine & "listed tag - " & listed_tag & vbNewLine & "tag matched - " & tag_matched & vbNewLine & script_array(current_script).script_name & vbNewLine & "list this script - " & list_this_script
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
                        ' If script_array(current_script).in_testing = TRUE Then script_array(current_script).description = "IN TESTING - " & script_array(current_script).description
                        ' MsgBox script_array(current_script).script_name & vbNewLine & "Use this button - " & use_this_button & vbNewLine & "show script - " & script_array(current_script).show_script
                    End If
                Else
                    script_array(current_script).show_script = FALSE
                End If


            ElseIf show_resources = TRUE Then
                script_array(current_script).show_script = FALSE
                ' MsgBox "Script - " & script_array(current_script).script_name & vbCr & "Release Date - " & script_array(current_script).release_date & vbCr & "Diff - " & DateDiff("d", script_array(current_script).release_date, two_months_ago)
                If DateDiff("d", script_array(current_script).release_date, two_months_ago) <= 0 Then
                    script_array(current_script).show_script = TRUE
                    new_script_to_list = TRUE
                    scripts_included = scripts_included + 1
                End If
                If IsDate(script_array(current_script).hot_topic_date) = TRUE Then
                    If DateDiff("d", script_array(current_script).hot_topic_date, two_months_ago) <= 0 Then
                        script_array(current_script).show_script = TRUE
                        hot_topic_script_to_list = TRUE
                        scripts_included = scripts_included + 1
                    End If
                End If
            ElseIf qi_menu = TRUE then
                script_array(current_script).show_script = FALSE
                ' MsgBox script_array(current_script).script_name & vbNewLine & JOIN(script_array(current_script).tags, ", ")
                For each listed_tag in script_array(current_script).tags
                    If listed_tag = "QI" Then
                        ' MsgBox "Script saved"
                        script_array(current_script).show_script = TRUE
                        Call script_array(current_script).show_button(use_this_button)
                        If use_this_button = FALSE Then script_array(current_script).show_script = FALSE
                        ' If script_array(current_script).in_testing = TRUE Then script_array(current_script).description = "IN TESTING - " & script_array(current_script).description
                    End If
                Next
                ' MsgBox script_array(current_script).show_script
            ElseIf bz_menu = TRUE Then
                script_array(current_script).show_script = FALSE
                ' MsgBox script_array(current_script).script_name & vbNewLine & JOIN(script_array(current_script).tags, ", ")
                For each listed_tag in script_array(current_script).tags
                    If listed_tag = "BZ" Then
                        ' MsgBox "Script saved"
                        script_array(current_script).show_script = TRUE
                    ElseIf listed_tag = "Monthly Tasks" Then
                        script_array(current_script).show_script = FALSE

                    End If
                Next
                If script_array(current_script).show_script = TRUE Then
                    Call script_array(current_script).show_button(use_this_button)
                    If use_this_button = FALSE Then script_array(current_script).show_script = FALSE
                    ' If script_array(current_script).in_testing = TRUE Then script_array(current_script).description = "IN TESTING - " & script_array(current_script).description
                End If
                ' MsgBox script_array(current_script).show_script
            ElseIf task_menu = TRUE Then
                script_array(current_script).show_script = FALSE
                ' MsgBox script_array(current_script).script_name & vbNewLine & JOIN(script_array(current_script).tags, ", ")
                For each listed_tag in script_array(current_script).tags
                    If listed_tag = "Monthly Tasks" Then
                        ' MsgBox "Script saved"
                        script_array(current_script).show_script = TRUE
                        Call script_array(current_script).show_button(use_this_button)
                        If use_this_button = FALSE Then script_array(current_script).show_script = FALSE
                        ' If script_array(current_script).in_testing = TRUE Then script_array(current_script).description = "IN TESTING - " & script_array(current_script).description
                    End If
                Next
                ' MsgBox script_array(current_script).show_script
            End If

            If show_resources = FALSE Then
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
            End If

            ' 'Subcategory handling (creating a second list as a string which gets converted later to an array)
            ' If ucase(script_array(current_script).category) = ucase(script_category) then																								'If the script in the array is of the correct category (ACTIONS/NOTES/ETC)...
            '     For each listed_subcategory in script_array(current_script).subcategory																									'...then iterate through each listed subcategory, and...
            '         If listed_subcategory <> "" and InStr(subcategory_list, ucase(listed_subcategory)) = 0 then subcategory_list = subcategory_list & "|" & ucase(listed_subcategory)	'...if the listed subcategory isn't blank and isn't already in the list, then add it to our handy-dandy list.
            '     Next
            ' End if
            'Adds a "NEW!!!" notification to the description if the script is from the last two months.
            ' If DateDiff("m", script_array(current_script).release_date, DateAdd("m", -2, date)) <= 0 then
            '     If left(script_array(current_script).description, 3) <> "NEW" Then script_array(current_script).description = "NEW " & script_array(current_script).release_date & "!!! --- " & script_array(current_script).description
            '     ' script_array(current_script).release_date = "12/12/1999" 'backs this out and makes it really old so it doesn't repeat each time the dialog loops. This prevents NEW!!!... from showing multiple times in the description.
            ' End if

        Next

        If show_resources = TRUE Then
            dlg_len = dlg_len + 50

            If new_script_to_list = TRUE Then dlg_len = dlg_len + 10
            If hot_topic_script_to_list = TRUE Then dlg_len = dlg_len + 10

            For current_script = 0 to ubound(script_array)
                If DateDiff("d", script_array(current_script).release_date, two_months_ago) <= 0 Then
                    dlg_len = dlg_len + 15
                End If
                If IsDate(script_array(current_script).hot_topic_date) = TRUE Then
                    If DateDiff("d", script_array(current_script).hot_topic_date, two_months_ago) <= 0 Then
                        dlg_len = dlg_len + 15
                    End If
                End If
            Next
            If dlg_len > 385 Then dlg_len = 385
        End If

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
            If qi_staff = FALSE Then
                PushButton 5, 10, 60, 15,   button_name_top_1, button_clik_top_1
                PushButton 65, 10, 60, 15,  button_name_top_2, button_clik_top_2
                PushButton 125, 10, 60, 15, button_name_top_3, button_clik_top_3
                PushButton 185, 10, 60, 15, button_name_top_4, button_clik_top_4
                PushButton 245, 10, 60, 15, button_name_top_5, button_clik_top_5
                PushButton 305, 10, 60, 15, button_name_top_6, button_clik_top_6
                PushButton 365, 10, 60, 15, button_name_top_7, button_clik_top_7
                PushButton 425, 10, 60, 15, button_name_top_8, button_clik_top_8
                PushButton 485, 10, 60, 15, button_name_top_9, button_clik_top_0

            Else

                PushButton 5, 10, 60, 15,   button_name_top_1, button_clik_top_1
                PushButton 65, 10, 60, 15,  button_name_top_2, button_clik_top_2
                PushButton 125, 10, 60, 15, button_name_top_3, button_clik_top_3
                PushButton 185, 10, 60, 15, button_name_top_4, button_clik_top_4
                PushButton 245, 10, 60, 15, button_name_top_5, button_clik_top_5
                PushButton 305, 10, 60, 15, button_name_top_6, button_clik_top_6
                PushButton 365, 10, 60, 15, button_name_top_7, button_clik_top_7
                PushButton 425, 10, 60, 15, button_name_top_8, button_clik_top_8
                PushButton 485, 10, 30, 15, button_name_top_9, button_clik_top_9

                PushButton 515, 10, 30, 15, button_name_top_0, button_clik_top_0
            End If
            PushButton 5, 25, 60, 15,   button_name_bottom_1, button_clik_bottom_1
            PushButton 65, 25, 60, 15,  button_name_bottom_2, button_clik_bottom_2
            PushButton 125, 25, 60, 15, button_name_bottom_3, button_clik_bottom_3
            PushButton 185, 25, 60, 15, button_name_bottom_4, button_clik_bottom_4
            PushButton 245, 25, 60, 15, button_name_bottom_5, button_clik_bottom_5
            PushButton 305, 25, 60, 15, button_name_bottom_6, button_clik_bottom_6
            PushButton 365, 25, 60, 15, button_name_bottom_7, button_clik_bottom_7
            PushButton 425, 25, 60, 15, button_name_bottom_8, button_clik_bottom_8
            PushButton 485, 25, 60, 15, button_name_bottom_9, button_clik_bottom_9

            PushButton 640, 190, 10, 10, "?", explain_questionmark_btn
            PushButton 640, 200, 10, 10, "+", explain_plus_btn

            vert_button_position = 50
            If show_resources = False Then
                list_counter = 0
                For current_script = 0 to ubound(script_array)
                    If script_array(current_script).show_script = TRUE Then
                        show_this_one = FALSE
                        If current_page = "One" AND list_counter < 20 Then show_this_one = TRUE
                        If current_page = "Two" AND list_counter >= 20 AND list_counter < 40 Then show_this_one = TRUE
                        If current_page = "Three" AND list_counter >= 40 Then show_this_one = TRUE
                        ' MsgBox "Script - " & script_array(current_script).script_name & vbNewLine & "Current page - " & current_page & vbNewLine & "list_counter - " & list_counter & vbNewLine & "show this one - " & show_this_one
                        If show_this_one = TRUE Then
                        ' If tab_selected <> "" Then
                        '     For each listed_tag in script_array(current_script).tags
                        '         If listed_tag <> "" Then
                        '             If UCase(listed_tag) = UCase(tab_selected) then
                                        SIR_button_placeholder = button_placeholder + 1	'We always want this to be one more than the button_placeholder
                                        add_to_favorites_button_placeholder = button_placeholder + 2
										ht_button_placeholder = button_placeholder + 3
                                        script_keys_combine = ""
                                        If script_array(current_script).dlg_keys(0) <> "" Then script_keys_combine = Join(script_array(current_script).dlg_keys, ":")

                                        'Displays the button and text description-----------------------------------------------------------------------------------------------------------------------------
                                        'FUNCTION		HORIZ. ITEM POSITION	VERT. ITEM POSITION		ITEM WIDTH	ITEM HEIGHT		ITEM TEXT/LABEL										BUTTON VARIABLE
                                        PushButton 		5, 						vert_button_position, 	10, 		12, 			"?", 												SIR_button_placeholder
                                        PushButton 		18,						vert_button_position, 	120, 		12, 			script_array(current_script).script_name, 			button_placeholder
                                        PushButton      140,                    vert_button_position,   10,         12,             "+",                                                add_to_favorites_button_placeholder
                                        Text 			150, 				    vert_button_position+1, 65, 		14, 			"-- " & script_keys_combine & " --"
									If script_array(current_script).hot_topic_link = "" Then
										Text            210,                    vert_button_position+1, 425,        14,             script_array(current_script).description
									Else
										PushButton		210,					vert_button_position, 	15, 		12, 			"HT",												ht_button_placeholder
										Text            225,                    vert_button_position+1, 425,        14,             script_array(current_script).description
									End If                                        '----------
                                        vert_button_position = vert_button_position + 15	'Needs to increment the vert_button_position by 15px (used by both the text and buttons)
                                        '----------
                                        script_array(current_script).button = button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
                                        script_array(current_script).SIR_instructions_button = SIR_button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
                                        script_array(current_script).fav_add_button = add_to_favorites_button_placeholder
										script_array(current_script).script_btn_one = ht_button_placeholder
                                        button_placeholder = button_placeholder + 4

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
            ElseIf show_resources = TRUE Then
                Text 15, vert_button_position, 500, 15, "This is the resources tab. This tab will provide you access to information about scripts, including highlighted scripts and new scripts. You can also find links here to report issues or contact the BlueZone Script Team."
                vert_button_position = vert_button_position + 20

                PushButton 15, vert_button_position, 75, 15, "Script Error Report", script_error_report_btn
                ' PushButton 85, vert_button_position, 100, 15, "Script Idea or Enhancement", script_idea_report_btn						'Removing these for now because we don't have functionality for this
                ' PushButton 185, vert_button_position, 95, 15, "Sign up for a Script Demo", script_demo_btn
                PushButton 450, vert_button_position, 75, 15, "Email BZST", email_bzst_btn
                PushButton 525, vert_button_position, 75, 15, "Email QI", email_qi_btn
                vert_button_position = vert_button_position + 25
                'Add buttons - Report Error, Email BZST, Email QI, Submit a script Idea or Enhancement, Sign up for Script Demos'
                If new_script_to_list = TRUE Then
                    Text 15, vert_button_position, 600, 10, "------------------------------------------------------------ NEW SCRIPTS ------------------------------------------------------------                                                                                          Added within the past two months"
                    vert_button_position = vert_button_position + 10
                    For current_script = 0 to ubound(script_array)
                        If DateDiff("d", script_array(current_script).release_date, two_months_ago) <= 0 Then
							show_this_one = TRUE
							If script_array(current_script).category = "ADMIN" Then
								show_this_one = FALSE
								For each review_group in script_array(current_script).tags
									If bz_staff = TRUE AND review_group = "BZ" Then show_this_one = TRUE
									If qi_staff = TRUE AND review_group = "QI" Then show_this_one = TRUE
								Next
							End If
							If show_this_one = TRUE Then
	                            SIR_button_placeholder = button_placeholder + 1	'We always want this to be one more than the button_placeholder
	                            add_to_favorites_button_placeholder = button_placeholder + 2
								ht_button_placeholder = button_placeholder + 3
	                            script_keys_combine = ""
	                            If script_array(current_script).dlg_keys(0) <> "" Then script_keys_combine = Join(script_array(current_script).dlg_keys, ":")

	                            'Displays the button and text description-----------------------------------------------------------------------------------------------------------------------------
	                            'FUNCTION		HORIZ. ITEM POSITION	VERT. ITEM POSITION		ITEM WIDTH	ITEM HEIGHT		ITEM TEXT/LABEL										BUTTON VARIABLE
	                            PushButton 		5, 						vert_button_position, 	10, 		12, 			"?", 												SIR_button_placeholder
	                            PushButton 		18,						vert_button_position, 	120, 		12, 			script_array(current_script).script_name, 			button_placeholder
	                            PushButton      140,                    vert_button_position,   10,         12,             "+",                                                add_to_favorites_button_placeholder
	                            Text 			150, 				    vert_button_position+1, 65, 		14, 			"-- " & script_keys_combine & " --"
							If script_array(current_script).hot_topic_link = "" Then
								Text            210,                    vert_button_position+1, 425,        14,             script_array(current_script).description
							Else
								PushButton		210,					vert_button_position, 	15, 		12, 			"HT",												ht_button_placeholder
								Text            225,                    vert_button_position+1, 425,        14,             script_array(current_script).description
							End If	                            '----------
	                            vert_button_position = vert_button_position + 15	'Needs to increment the vert_button_position by 15px (used by both the text and buttons)
	                            '----------
	                            script_array(current_script).button = button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
	                            script_array(current_script).SIR_instructions_button = SIR_button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
	                            script_array(current_script).fav_add_button = add_to_favorites_button_placeholder
								script_array(current_script).script_btn_one = ht_button_placeholder
								button_placeholder = button_placeholder + 4
							End If
                        End If
                    Next
                    vert_button_position = vert_button_position + 10
                End If
                If hot_topic_script_to_list = TRUE Then
                    Text 15, vert_button_position, 500, 10, "------------------------------------------------------------ FEATURED SCRIPTS ------------------------------------------------------------"
                    PushButton 515, vert_button_position, 75, 10, "See HOT TOPICS", hot_topics_btn
                    vert_button_position = vert_button_position + 10
                    For current_script = 0 to ubound(script_array)
                        If IsDate(script_array(current_script).hot_topic_date) = TRUE Then
                            If DateDiff("d", script_array(current_script).hot_topic_date, two_months_ago) <= 0 Then
                                SIR_button_placeholder = button_placeholder + 1	'We always want this to be one more than the button_placeholder
                                add_to_favorites_button_placeholder = button_placeholder + 2
								ht_button_placeholder = button_placeholder + 3
                                script_keys_combine = ""
                                If script_array(current_script).dlg_keys(0) <> "" Then script_keys_combine = Join(script_array(current_script).dlg_keys, ":")

                                'Displays the button and text description-----------------------------------------------------------------------------------------------------------------------------
                                'FUNCTION		HORIZ. ITEM POSITION	VERT. ITEM POSITION		ITEM WIDTH	ITEM HEIGHT		ITEM TEXT/LABEL										BUTTON VARIABLE
                                PushButton 		5, 						vert_button_position, 	10, 		12, 			"?", 												SIR_button_placeholder
                                PushButton 		18,						vert_button_position, 	120, 		12, 			script_array(current_script).script_name, 			button_placeholder
                                PushButton      140,                    vert_button_position,   10,         12,             "+",                                                add_to_favorites_button_placeholder
                                Text 			150, 				    vert_button_position+1, 65, 		14, 			"-- " & script_keys_combine & " --"
							If script_array(current_script).hot_topic_link = "" Then
								Text            210,                    vert_button_position+1, 425,        14,             script_array(current_script).description
							Else
								PushButton		210,					vert_button_position, 	15, 		12, 			"HT",												ht_button_placeholder
								Text            225,                    vert_button_position+1, 425,        14,             script_array(current_script).description
							End If
                                '----------
                                vert_button_position = vert_button_position + 15	'Needs to increment the vert_button_position by 15px (used by both the text and buttons)
                                '----------
                                script_array(current_script).button = button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
                                script_array(current_script).SIR_instructions_button = SIR_button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
                                script_array(current_script).fav_add_button = add_to_favorites_button_placeholder
								script_array(current_script).script_btn_one = ht_button_placeholder
                                button_placeholder = button_placeholder + 4
                            End If
                        End If
                    Next
                End If
                ' If IsDate(script_array(current_script).hot_topic_date) = TRUE Then
                '     If script_array(current_script).hot_topic_date < one_month_ago Then script_array(current_script).show_script = FALSE
                ' End If
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
			PushButton 110, dlg_len - 20, 100, 15, "Reources", resources_btn
            If bz_staff = TRUE Then
                PushButton 595, dlg_len - 20, 55, 15, "Monthly Tasks", monthly_task_btn
                PushButton 650, dlg_len - 20, 40, 15, "BZST", bz_btn
            End If
            CancelButton 690, dlg_len - 20, 50, 15
            ' Text 120, vert_button_position, 380, 15, "C - Case Notes ... E - Excel ... EXP - Expedited SNAP ... F - FIATs ... OA - Outlook Appointment ... OE - Outlook Email ... SM - SPEC/MEMO ... SW - SPEC/WCOM ... T - TIKL ... U - Updates Panel ... W - Word"
          ' Text 120, dlg_len - 15, 40, 10, "Keywords:"                 'commented out because we don't have keywordds
          ' EditBox 165, dlg_len - 20, 200, 15, search_keywords
        EndDialog


end function

function send_script_error()
	Do
		confirm_err = ""

		case_note_checkbox = unchecked
		stat_update_checkbox = unchecked
		date_checkbox = unchecked
		math_checkbox = unchecked
		tikl_checkbox = unchecked
		memo_wcom_checkbox = unchecked
		document_checkbox = unchecked
		missing_spot_checkbox = unchecked

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 401, 185, "Report Error Detail"
		  EditBox 60, 25, 55, 15, MAXIS_case_number
		  ComboBox 220, 25, 175, 45, error_type+chr(9)+"BUG - something happened that was wrong"+chr(9)+"ENHANCEMENT - something could be done better"+chr(9)+"TYPO - grammatical/spelling type errors"+chr(9)+"DAIL - add support for this DAIL message.", error_type
		  EditBox 85, 45, 310, 15, script_names
		  EditBox 65, 65, 330, 15, error_detail
		  CheckBox 20, 115, 65, 10, "CASE/NOTE", case_note_checkbox
		  CheckBox 95, 115, 65, 10, "Update in STAT", stat_update_checkbox
		  CheckBox 170, 115, 75, 10, "Problems with Dates", date_checkbox
		  CheckBox 265, 115, 65, 10, "Math is incorrect", math_checkbox
		  CheckBox 20, 130, 65, 10, "TIKL is incorrect", tikl_checkbox
		  CheckBox 95, 130, 65, 10, "MEMO or WCOM", memo_wcom_checkbox
		  CheckBox 170, 130, 75, 10, "Created Document", document_checkbox
		  CheckBox 265, 130, 115, 10, "Missing a place for Information", missing_spot_checkbox
		  EditBox 60, 155, 165, 15, email_signature
		  ButtonGroup ButtonPressed
		    OkButton 290, 155, 50, 15
		    CancelButton 345, 155, 50, 15
		  Text 10, 10, 300, 10, "Information is needed about the error for our scriptwriters to review and resolve the issue. "
		  Text 5, 30, 50, 10, "Case Number:"
		  Text 125, 30, 95, 10, "What type of error occured?"
		  Text 5, 50, 75, 10, "Script(s) with an Error:"
		  Text 5, 70, 60, 10, "Explain in detail:"
		  GroupBox 10, 90, 380, 60, "Common areas of issue"
		  Text 20, 100, 200, 10, "Check any that were impacted by the error you are reporting."
		  Text 10, 160, 50, 10, "Worker Name:"
		  Text 25, 175, 335, 10, "*** Remember to leave the case as is if possible. We can resolve error better when in a live case. ***"
		EndDialog

		Dialog Dialog1

		If ButtonPressed = 0 Then
			cancel_confirm_msg = MsgBox("An Error Report will NOT be sent as you pressed 'Cancel'." & vbNewLine & vbNewLine & "Is this what you would like to do?", vbQuestion + vbYesNo, "Confirm Cancel")
			If cancel_confirm_msg = vbYes Then confirm_err = ""
			If cancel_confirm_msg = vbNo Then confirm_err = "LOOP" & vbNewLine & confirm_err
		End If

		If ButtonPressed = -1 Then
			full_text = "Error occurred on " & date & " at " & time
			full_text = full_text & vbCr & "Error type - " & error_type
			full_text = full_text & vbCr & "Script name - " & script_names & " was run on Case #" & MAXIS_case_number & " with a runtime of " & script_run_time & " seconds."
			full_text = full_text & vbCr & "Information: " & error_detail
			If case_note_checkbox = checked OR stat_update_checkbox = checked OR date_checkbox = checked OR math_checkbox = checked OR tikl_checkbox = checked OR memo_wcom_checkbox = checked OR document_checkbox = checked OR missing_spot_checkbox = checked Then full_text = full_text & vbCr & vbCr & "Script has issues/concerns in the following areas:"

			If case_note_checkbox = checked Then full_text = full_text & vbCr & " - CASE/NOTE"
			If stat_update_checkbox = checked Then full_text = full_text & vbCr & " - Update in STAT"
			If date_checkbox = checked Then full_text = full_text & vbCr & " - Dates are incorrect"
			If math_checkbox = checked Then full_text = full_text & vbCr & " - Math is incorrect"
			If tikl_checkbox = checked Then full_text = full_text & vbCr & " - TIKL"
			If memo_wcom_checkbox = checked Then full_text = full_text & vbCr & " - NOTICES (WCOM/MEMO)"
			If document_checkbox = checked Then full_text = full_text & vbCr & " - The Excel or Word Document"
			If missing_spot_checkbox = checked Then full_text = full_text & vbCr & " - There is no space to enter particular information"

			full_text = full_text & vbCr & "Closing message: " & closing_message
			full_text = full_text & vbCr & vbCr & "Sent by: " & worker_signature

			send_confirm_msg = MsgBox("** This is what will be sent as an email to the BlueZone Script team:" & vbNewLine & vbNewLine & full_text & vbNewLine & vbNewLine & "*** Is this what you want to send? ***", vbQuestion + vbYesNo, "Confirm Error Report")

			If send_confirm_msg = vbYes Then confirm_err = ""
			If send_confirm_msg = vbNo Then confirm_err = "LOOP" & vbNewLine & confirm_err
		End If
	Loop until confirm_err = ""

	full_text = ""
	If ButtonPressed = -1 Then
		bzt_email = "HSPH.EWS.BlueZoneScripts@hennepin.us"
		subject_of_email = "Script Error -- " & script_names & " (Automated Report)"

		full_text = "Error reported from TAGS Menu on " & date & " at " & time
		full_text = full_text & vbCr & "Error type - " & error_type
		full_text = full_text & vbCr & "Script name - " & script_names & " was run on Case #" & MAXIS_case_number & " with a runtime of " & script_run_time & " seconds."
		full_text = full_text & vbCr & "Information: " & error_detail
		If case_note_checkbox = checked OR stat_update_checkbox = checked OR date_checkbox = checked OR math_checkbox = checked OR tikl_checkbox = checked OR memo_wcom_checkbox = checked OR document_checkbox = checked OR missing_spot_checkbox = checked Then full_text = full_text & vbCr & vbCr & "Script has issues/concerns in the following areas:"

		If case_note_checkbox = checked Then full_text = full_text & vbCr & " - CASE/NOTE"
		If stat_update_checkbox = checked Then full_text = full_text & vbCr & " - Update in STAT"
		If date_checkbox = checked Then full_text = full_text & vbCr & " - Dates are incorrect"
		If math_checkbox = checked Then full_text = full_text & vbCr & " - Math is incorrect"
		If tikl_checkbox = checked Then full_text = full_text & vbCr & " - TIKL"
		If memo_wcom_checkbox = checked Then full_text = full_text & vbCr & " - NOTICES (WCOM/MEMO)"
		If document_checkbox = checked Then full_text = full_text & vbCr & " - The Excel or Word Document"
		If missing_spot_checkbox = checked Then full_text = full_text & vbCr & " - There is no space to enter particular information"

		full_text = full_text & vbCr & "Closing message: " & closing_message
		full_text = full_text & vbCr & vbCr & "Sent by: " & worker_signature

		If script_run_lowdown <> "" Then full_text = full_text & vbCr & vbCr & "All Script Run Details:" & vbCr & script_run_lowdown

		Call create_outlook_email(bzt_email, "", subject_of_email, full_text, "", true)

		MsgBox "Error Report completed!" & vbNewLine & vbNewLine & "Thank you for working with us for Continuous Improvement."
	Else
		MsgBox "Your error report has been cancelled and has NOT been sent to the BlueZone Script Team"
	End If
end function

'Starting these with a very high number, higher than the normal possible amount of buttons.
'	We're doing this because we want to assign a value to each button pressed, and we want
'	that value to change with each button. The button_placeholder will be placed in the .button
'	property for each script item. This allows it to both escape the Function and resize
'	near infinitely. We use dummy numbers for the other selector buttons for much the same reason,
'	to force the value of ButtonPressed to hold in near infinite iterations.
button_placeholder 			= 24601

'Other pre-loop and pre-function declarations
subcategory_array = array()
subcategory_string = ""
subcategory_selected = "MAIN"
select_tab = ""
email_signature = worker_signature

' Dim snap_btn, mfip_btn, dwp_btn, hc_btn, grh_btn, ga_btn, emer_btn, ltc_btn, abawd_btn, income_btn, asset_btn, deductions_btn, application_btn, review_btn, communication_btn, utility_btn, reports_btn, resources_btn, clear_selection_btn
' Dim page_one_btn, page_two_btn, page_three_btn, current_page, hot_topics_btn, script_error_report_btn, script_idea_report_btn, script_demo_btn, email_bzst_btn, email_qi_btn
' Dim button_clik_top_1, button_clik_top_2, button_clik_top_3, button_clik_top_4, button_clik_top_5, button_clik_top_6, button_clik_top_7, button_clik_top_8, button_clik_top_0

dIM current_page
snap_btn = 100
mfip_btn = 200
dwp_btn = 300
hc_btn = 400
grh_btn = 500
ga_btn = 600
emer_btn = 700
ltc_btn = 800
abawd_btn = 900
income_btn = 1000
asset_btn = 1100
deductions_btn = 1200
application_btn = 1300
review_btn = 1400
communication_btn = 1500
utility_btn = 1600
reports_btn = 1700
support_btn = 1750
resources_btn = 1800
clear_selection_btn = 1900
page_one_btn = 2000
page_two_btn = 2100
page_three_btn = 2200
' current_page = 2300
monthly_task_btn = 2300
hot_topics_btn = 2400
script_error_report_btn = 2500
script_idea_report_btn = 2600
script_demo_btn = 2700
email_bzst_btn = 2800
email_qi_btn = 2900
button_clik_top_1 = 3000
button_clik_top_2 = 3100
button_clik_top_3 = 3200
button_clik_top_4 = 3300
button_clik_top_5 = 3400
button_clik_top_6 = 3500
button_clik_top_7 = 3600
button_clik_top_8 = 3700
button_clik_top_9 = 3800
button_clik_top_0 = 3900
qi_btn = 4000
bz_btn = 4100
button_clik_bottom_1 = 4200
button_clik_bottom_2 = 4300
button_clik_bottom_3 = 4400
button_clik_bottom_4 = 4500
button_clik_bottom_5 = 4600
button_clik_bottom_6 = 4700
button_clik_bottom_7 = 4800
button_clik_bottom_8 = 4900
button_clik_bottom_9 = 5000
button_clik_bottom_0 = 5100

button_name_top_1 = "SNAP"
button_name_top_2 = "MFIP"
button_name_top_3 = "DWP"
button_name_top_4 = "Health Care"
button_name_top_5 = "HS/GRH"
button_name_top_6 = "Adult Cash"
button_name_top_7 = "EMER"
button_name_top_8 = "ABAWD"
button_name_top_9 = "LTC"
button_name_top_0 = "QI"

button_name_bottom_1 = "Income"
button_name_bottom_2 = "Assets"
button_name_bottom_3 = "Deductions"
button_name_bottom_4 = "Applications"
button_name_bottom_5 = "Reviews"
button_name_bottom_6 = "Communication"
button_name_bottom_7 = "Utility"
button_name_bottom_8 = "Reports"
button_name_bottom_9 = "Support"
button_name_bottom_0 = ""

qi_menu = FALSE
bz_menu = FALSE
task_menu = FALSE
show_resources = FALSE
Do
    leave_loop = ""

    Call declare_tabbed_menu(select_tab)
    dialog Dialog1

    cancel_without_confirmation

    leave_loop = FALSE

    If ButtonPressed = button_clik_top_1 Then ButtonPressed = snap_btn
    If ButtonPressed = button_clik_top_2 Then ButtonPressed = mfip_btn
    If ButtonPressed = button_clik_top_3 Then ButtonPressed = dwp_btn
    If ButtonPressed = button_clik_top_4 Then ButtonPressed = hc_btn
    If ButtonPressed = button_clik_top_5 Then ButtonPressed = grh_btn
    If ButtonPressed = button_clik_top_6 Then ButtonPressed = ga_btn
    If ButtonPressed = button_clik_top_7 Then ButtonPressed = emer_btn
    If ButtonPressed = button_clik_top_8 Then ButtonPressed = abawd_btn
    If ButtonPressed = button_clik_top_9 Then ButtonPressed = ltc_btn
    If ButtonPressed = button_clik_top_0 Then ButtonPressed = qi_btn

    If ButtonPressed = button_clik_bottom_1 Then ButtonPressed = income_btn
    If ButtonPressed = button_clik_bottom_2 Then ButtonPressed = asset_btn
    If ButtonPressed = button_clik_bottom_3 Then ButtonPressed = deductions_btn
    If ButtonPressed = button_clik_bottom_4 Then ButtonPressed = application_btn
    If ButtonPressed = button_clik_bottom_5 Then ButtonPressed = review_btn
    If ButtonPressed = button_clik_bottom_6 Then ButtonPressed = communication_btn
    If ButtonPressed = button_clik_bottom_7 Then ButtonPressed = utility_btn
    If ButtonPressed = button_clik_bottom_8 Then ButtonPressed = reports_btn
    If ButtonPressed = button_clik_bottom_9 Then ButtonPressed = support_btn

    If ButtonPressed = explain_questionmark_btn Then
        explain_questionmark_msg = MsgBox("See all the Question Mark Buttons?" & vbCr & vbCr & "Look to the left of the script name button. Each script has a button with a question mark - ? - next to it." & vbCr & vbCr & "Press this button and the instructions for that script will be opened. This is an easy way to see how a script functions, or when to use it.", vbQuestion, "What's with the Question Marks?")
    End If
    If ButtonPressed = explain_plus_btn Then
        explain_plus_msg = MsgBox("What are these plus buttons?" & vbCr & vbCr & "Just to the right of the script name button is a button with a plus sign (+)." & vbCr & vbCr & "This button will add this script to your list of favorite scripts for access inthe favorites menu.", vbQuestion, "What are the Plus Signs for?")
    End If

    If ButtonPressed = qi_btn Then
        qi_menu = TRUE
        bz_menu = FALSE
        task_menu = FALSE
    End If
    If ButtonPressed = bz_btn Then
        qi_menu = FALSE
        bz_menu = TRUE
        task_menu = FALSE
    End If
    If ButtonPressed = monthly_task_btn Then
        qi_menu = FALSE
        bz_menu = FALSE
        task_menu = TRUE
    End If
    If ButtonPressed = clear_selection_btn Then
        ' leave_loop = FALSE
        select_tab = ""
        current_page = "One"
    End If

    If ButtonPressed = page_one_btn Then
        ' leave_loop = FALSE
        current_page = "One"
    End If
    If ButtonPressed = page_two_btn Then
        ' leave_loop = FALSE
        current_page = "Two"
    End If
    If ButtonPressed = page_three_btn Then
        ' leave_loop = FALSE
        current_page = "Three"
    End If
    If ButtonPressed = dail_scrubber_instructions_btn Then
        Call open_URL_in_browser("https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/DAIL/ALL%20DAIL%20SCRIPTS.docx")
        ' leave_loop = FALSE
    End If
    If ButtonPressed = hot_topics_btn OR ButtonPressed = script_error_report_btn OR ButtonPressed = script_idea_report_btn OR ButtonPressed = script_demo_btn OR ButtonPressed = email_bzst_btn OR ButtonPressed = email_qi_btn Then
        If ButtonPressed = hot_topics_btn Then Call open_URL_in_browser("https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/Economic_Supports_ES_Zone.aspx")
        If ButtonPressed = script_error_report_btn Then Call send_script_error
        ' If ButtonPressed = script_idea_report_btn Then
        ' If ButtonPressed = script_demo_btn Then
        If ButtonPressed = email_bzst_btn Then
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 431, 265, "Email to BZST"
			  EditBox 25, 20, 150, 15, email_subject_line
			  CheckBox 60, 50, 125, 10, "Check here if this email is Urgent.", urgent_email_checkbox
			  EditBox 10, 65, 415, 15, email_body_line_one
			  EditBox 10, 85, 415, 15, email_body_line_two
			  EditBox 10, 105, 415, 15, email_body_line_three
			  EditBox 10, 125, 415, 15, email_body_line_four
			  EditBox 10, 145, 415, 15, email_body_line_five
			  EditBox 10, 165, 415, 15, email_body_line_six
			  EditBox 10, 205, 415, 15, url_line
			  EditBox 225, 225, 200, 15, email_signature
			  CheckBox 15, 250, 215, 10, "Check here if you would like an email response from the BZST.", response_needed_checkbox
			  ButtonGroup ButtonPressed
			    OkButton 325, 245, 50, 15
			    CancelButton 375, 245, 50, 15
			  Text 5, 10, 50, 10, "Subject Line"
			  Text 10, 25, 15, 10, "RE:"
			  Text 10, 50, 50, 10, "Email Body"
			  Text 10, 190, 195, 10, "Any policy/procedure to reference? Copy and paste it here:"
			  Text 170, 230, 55, 10, "Sign your Email:"
			  Text 275, 10, 145, 40, "This functionality will automate an email to the BlueZone Script Team at HSPS.EWS.BlueZoneScripts@hennepin.us."
			EndDialog

			Do
				err_msg = ""

				dialog Dialog1

				email_subject_line = trim(email_subject_line)
  			  	email_body_line_one = trim(email_body_line_one)
  			  	email_body_line_two = trim(email_body_line_two)
  			  	email_body_line_three = trim(email_body_line_three)
  			  	email_body_line_four = trim(email_body_line_four)
  			  	email_body_line_five = trim(email_body_line_five)
  			  	email_body_line_six = trim(email_body_line_six)
				url_line = trim(url_line)
				email_signature = trim(email_signature)

				If email_subject_line = "" Then err_msg = err_msg & vbNewLine & "* Enter something in the subject header line."
  			  	If email_body_line_one = "" AND email_body_line_two = "" AND email_body_line_three = "" AND email_body_line_four = "" AND email_body_line_five = "" AND email_body_line_six = "" Then err_msg = err_msg & vbNewLine & "* Enter information in at least one line of the email body."
				If email_signature = "" Then err_msg = err_msg & vbNewLine & "* Enter your name/signature for the email."

				If ButtonPressed = 0 Then err_msg = ""

				If err_msg <> "" Then MsgBox "*** Please Resolve to Continue:" & vbNewLine & err_msg
			Loop until err_msg = ""
			If ButtonPressed = -1 Then

				If urgent_email_checkbox = checked Then email_subject_line = "URGENT!  " & email_subject_line
				If email_body_line_one <> "" Then   email_body_lines = email_body_lines & email_body_line_one & vbCr & vbCr
				If email_body_line_two <> "" Then   email_body_lines = email_body_lines & email_body_line_two & vbCr & vbCr
				If email_body_line_three <> "" Then email_body_lines = email_body_lines & email_body_line_three & vbCr & vbCr
				If email_body_line_four <> "" Then  email_body_lines = email_body_lines & email_body_line_four & vbCr & vbCr
				If email_body_line_five <> "" Then  email_body_lines = email_body_lines & email_body_line_five & vbCr & vbCr
				If email_body_line_six <> "" Then   email_body_lines = email_body_lines & email_body_line_six & vbCr & vbCr
				If response_needed_checkbox = checked Then email_body_lines = "RESPONSE REQUESTED" & vbCr & vbCr & email_body_lines
				If url_line <> "" Then email_body_lines = email_body_lines & vbCr & vbCr & "Referenced Link: " & url_line
				email_body_lines = email_body_lines & vbCr & vbCr & "---" & vbCr & email_signature

				Call create_outlook_email("HSPH.EWS.BlueZoneScripts@hennepin.us", "", email_subject_line, email_body_lines, "", TRUE)

				MsgBox "Email Sent to BZST" & vbNewLine & "----------------------------" & vbNewLine & "Subject: " & email_subject_line & vbNewLine & vbNewLine & email_body_lines
			Else
				MsgBox "Email to BZST has been cancelled."
			End If
		End If
        If ButtonPressed = email_qi_btn Then
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 431, 265, "Email to QI"
			  EditBox 25, 20, 150, 15, email_subject_line
			  CheckBox 60, 50, 125, 10, "Check here if this email is Urgent.", urgent_email_checkbox
			  EditBox 10, 65, 415, 15, email_body_line_one
			  EditBox 10, 85, 415, 15, email_body_line_two
			  EditBox 10, 105, 415, 15, email_body_line_three
			  EditBox 10, 125, 415, 15, email_body_line_four
			  EditBox 10, 145, 415, 15, email_body_line_five
			  EditBox 10, 165, 415, 15, email_body_line_six
			  EditBox 10, 205, 415, 15, url_line
			  EditBox 225, 225, 200, 15, email_signature
			  CheckBox 15, 250, 215, 10, "Check here if you would like an email response from the QI Team.", response_needed_checkbox
			  ButtonGroup ButtonPressed
				OkButton 325, 245, 50, 15
				CancelButton 375, 245, 50, 15
			  Text 5, 10, 50, 10, "Subject Line"
			  Text 10, 25, 15, 10, "RE:"
			  Text 10, 50, 50, 10, "Email Body"
			  Text 10, 190, 195, 10, "Any policy/procedure to reference? Copy and paste it here:"
			  Text 170, 230, 55, 10, "Sign your Email:"
			  Text 275, 10, 145, 40, "This functionality will automate an email to the BlueZone Script Team at HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us."
			EndDialog

			Do
				err_msg = ""

				dialog Dialog1

				email_subject_line = trim(email_subject_line)
				email_body_line_one = trim(email_body_line_one)
				email_body_line_two = trim(email_body_line_two)
				email_body_line_three = trim(email_body_line_three)
				email_body_line_four = trim(email_body_line_four)
				email_body_line_five = trim(email_body_line_five)
				email_body_line_six = trim(email_body_line_six)
				url_line = trim(url_line)
				email_signature = trim(email_signature)

				If email_subject_line = "" Then err_msg = err_msg & vbNewLine & "* Enter something in the subject header line."
				If email_body_line_one = "" AND email_body_line_two = "" AND email_body_line_three = "" AND email_body_line_four = "" AND email_body_line_five = "" AND email_body_line_six = "" Then err_msg = err_msg & vbNewLine & "* Enter information in at least one line of the email body."
				If email_signature = "" Then err_msg = err_msg & vbNewLine & "* Enter your name/signature for the email."

				If ButtonPressed = 0 Then err_msg = ""

				If err_msg <> "" Then MsgBox "*** Please Resolve to Continue:" & vbNewLine & err_msg
			Loop until err_msg = ""
			If ButtonPressed = -1 Then

				If urgent_email_checkbox = checked Then email_subject_line = "URGENT!  " & email_subject_line
				If email_body_line_one <> "" Then   email_body_lines = email_body_lines & email_body_line_one & vbCr & vbCr
				If email_body_line_two <> "" Then   email_body_lines = email_body_lines & email_body_line_two & vbCr & vbCr
				If email_body_line_three <> "" Then email_body_lines = email_body_lines & email_body_line_three & vbCr & vbCr
				If email_body_line_four <> "" Then  email_body_lines = email_body_lines & email_body_line_four & vbCr & vbCr
				If email_body_line_five <> "" Then  email_body_lines = email_body_lines & email_body_line_five & vbCr & vbCr
				If email_body_line_six <> "" Then   email_body_lines = email_body_lines & email_body_line_six & vbCr & vbCr
				If response_needed_checkbox = checked Then email_body_lines = "RESPONSE REQUESTED" & vbCr & vbCr & email_body_lines
				If url_line <> "" Then email_body_lines = email_body_lines & vbCr & vbCr & "Referenced Link: " & url_line
				email_body_lines = email_body_lines & vbCr & vbCr & "---" & vbCr & email_signature

				Call create_outlook_email("HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us", "", email_subject_line, email_body_lines, "", TRUE)

				MsgBox "Email Sent to Quality Improvement Team" & vbNewLine & "----------------------------" & vbNewLine & "Subject: " & email_subject_line & vbNewLine & vbNewLine & email_body_lines
			Else
				MsgBox "Email to Quality Improvement Team has been cancelled."
			End If
		End If
        ' leave_loop = FALSE
        ButtonPressed = resources_btn
    End If

    If ButtonPressed = snap_btn OR ButtonPressed = mfip_btn OR ButtonPressed = dwp_btn OR ButtonPressed = hc_btn OR ButtonPressed = grh_btn OR ButtonPressed = ga_btn OR ButtonPressed = emer_btn OR ButtonPressed = ltc_btn OR ButtonPressed = abawd_btn OR ButtonPressed = income_btn OR ButtonPressed = asset_btn OR ButtonPressed = deductions_btn OR ButtonPressed = application_btn OR ButtonPressed = review_btn OR ButtonPressed = communication_btn OR ButtonPressed = utility_btn OR ButtonPressed = reports_btn OR ButtonPressed = support_btn OR ButtonPressed = resources_btn Then

        ' leave_loop = FALSE
        qi_menu = FALSE
        bz_menu = FALSE
        task_menu = FALSE
        current_page = "One"
        If ButtonPressed = snap_btn             AND InStr(select_tab, "SNAP") = 0           Then select_tab = select_tab & "~" & "SNAP"
        If ButtonPressed = mfip_btn             AND InStr(select_tab, "MFIP") = 0           Then select_tab = select_tab & "~" & "MFIP"
        If ButtonPressed = dwp_btn              AND InStr(select_tab, "DWP") = 0            Then select_tab = select_tab & "~" & "DWP"
        If ButtonPressed = hc_btn               AND InStr(select_tab, "Health Care") = 0    Then select_tab = select_tab & "~" & "Health Care"
        If ButtonPressed = grh_btn              AND InStr(select_tab, "HS/GRH") = 0         Then select_tab = select_tab & "~" & "HS/GRH"
        If ButtonPressed = ga_btn               AND InStr(select_tab, "Adult Cash") = 0     Then select_tab = select_tab & "~" & "Adult Cash"
        If ButtonPressed = emer_btn             AND InStr(select_tab, "EMER") = 0           Then select_tab = select_tab & "~" & "EMER"
        If ButtonPressed = ltc_btn              AND InStr(select_tab, "LTC") = 0            Then select_tab = select_tab & "~" & "LTC"
        If ButtonPressed = abawd_btn            AND InStr(select_tab, "ABAWD") = 0          Then select_tab = select_tab & "~" & "ABAWD"
        If ButtonPressed = income_btn           AND InStr(select_tab, "Income") = 0         Then select_tab = select_tab & "~" & "Income"
        If ButtonPressed = asset_btn            AND InStr(select_tab, "Assets") = 0         Then select_tab = select_tab & "~" & "Assets"
        If ButtonPressed = deductions_btn       AND InStr(select_tab, "Deductions") = 0     Then select_tab = select_tab & "~" & "Deductions"
        If ButtonPressed = application_btn      AND InStr(select_tab, "Application") = 0    Then select_tab = select_tab & "~" & "Application"
        If ButtonPressed = review_btn           AND InStr(select_tab, "Reviews") = 0        Then select_tab = select_tab & "~" & "Reviews"
        If ButtonPressed = communication_btn    AND InStr(select_tab, "Communication") = 0  Then select_tab = select_tab & "~" & "Communication"
        If ButtonPressed = utility_btn          AND InStr(select_tab, "Utility") = 0        Then select_tab = select_tab & "~" & "Utility"
		If ButtonPressed = support_btn 			AND InStr(select_tab, "Support") = 0		Then select_tab = select_tab & "~" & "Support"
        If ButtonPressed = reports_btn          AND InStr(select_tab, "Reports") = 0        Then select_tab = select_tab & "~" & "Reports"
    End If
    If ButtonPressed = resources_btn Then
        show_resources = TRUE
    Else
        show_resources = FALSE
    End If

    For i = 0 to ubound(script_array)
		If ButtonPressed = script_array(i).SIR_instructions_button then
            call open_URL_in_browser(script_array(i).SharePoint_instructions_URL)
            ' leave_loop = FALSE
        End If
		If ButtonPressed = script_array(i).script_btn_one then
			call open_URL_in_browser(script_array(i).hot_topic_link)
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
