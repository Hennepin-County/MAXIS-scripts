'Required for statistical purposes==========================================================================================
name_of_script = "UTILITIES - ALL SCRIPTS.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
STATS_denomination = "I"       		'C is for each CASE
'END OF stats block=========================================================================================================

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
function add_page_buttons_to_dialog(page_variable, items_per_page, total_items, dlg_vert)
    '--- This function adds numbered buttons to the bottom of a dinamic dialog if there is a list that is too long to display in one dialog, this can be used to access the pages BUT this is ONLY the buttons to display not the functionality to switch between them.
    '~~~~~ page_variable: the name of the variable used to identify which page is being viewed
    '~~~~~ items_per_page: this must be a number and is now many items can be shown on one instance of the dialog
    '~~~~~ total_items: how many items exist in the list in total
    '~~~~~ dlg_vert: the variable used to define where elements of the dialog are
    '===== Keywords: MAXIS, dialog, list, dynamic, information

	If page <> 1 Then PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
	If page <> 2  AND total_items > items_per_page    Then PushButton 20,  dlg_vert, 10, 10, "2",  page_two_btn
	If page <> 3  AND total_items > items_per_page*2  Then PushButton 30,  dlg_vert, 10, 10, "3",  page_three_btn
	If page <> 4  AND total_items > items_per_page*3  Then PushButton 40,  dlg_vert, 10, 10, "4",  page_four_btn
	If page <> 5  AND total_items > items_per_page*4  Then PushButton 50,  dlg_vert, 10, 10, "5",  page_five_btn
	If page <> 6  AND total_items > items_per_page*5  Then PushButton 60,  dlg_vert, 10, 10, "6",  page_six_btn
	If page <> 7  AND total_items > items_per_page*6  Then PushButton 70,  dlg_vert, 10, 10, "7",  page_seven_btn
	If page <> 8  AND total_items > items_per_page*7  Then PushButton 80,  dlg_vert, 10, 10, "8",  page_eight_btn
	If page <> 9  AND total_items > items_per_page*8  Then PushButton 90,  dlg_vert, 10, 10, "9",  page_nine_btn
	If page <> 10 AND total_items > items_per_page*9  Then PushButton 100, dlg_vert, 12, 10, "10", page_ten_btn
	If page <> 11 AND total_items > items_per_page*10 Then PushButton 112, dlg_vert, 12, 10, "11", page_eleven_btn
	If page <> 12 AND total_items > items_per_page*11 Then PushButton 124, dlg_vert, 12, 10, "12", page_twelve_btn
	If page <> 13 AND total_items > items_per_page*12 Then PushButton 136, dlg_vert, 12, 10, "13", page_thirteen_btn
	If page <> 14 AND total_items > items_per_page*13 Then PushButton 148, dlg_vert, 12, 10, "14", page_fourteen_btn
	If page <> 15 AND total_items > items_per_page*14 Then PushButton 160, dlg_vert, 12, 10, "15", page_fifteen_btn
	If page <> 16 AND total_items > items_per_page*15 Then PushButton 172, dlg_vert, 12, 10, "16", page_sixteen_btn
	If page <> 17 AND total_items > items_per_page*16 Then PushButton 184, dlg_vert, 12, 10, "17", page_seventeen_btn
	If page <> 18 AND total_items > items_per_page*17 Then PushButton 196, dlg_vert, 12, 10, "18", page_eightteen_btn

	If page = 1  Then Text 12,  dlg_vert + 1, 8,  10, "1"
	If page = 2  Then Text 22,  dlg_vert + 1, 8,  10, "2"
	If page = 3  Then Text 32,  dlg_vert + 1, 8,  10, "3"
	If page = 4  Then Text 42,  dlg_vert + 1, 8,  10, "4"
	If page = 5  Then Text 52,  dlg_vert + 1, 8,  10, "5"
	If page = 6  Then Text 62,  dlg_vert + 1, 8,  10, "6"
	If page = 7  Then Text 72,  dlg_vert + 1, 8,  10, "7"
	If page = 8  Then Text 82,  dlg_vert + 1, 8,  10, "8"
	If page = 9  Then Text 92,  dlg_vert + 1, 8,  10, "9"
	If page = 10 Then Text 101, dlg_vert + 1, 10, 10, "10"
	If page = 11 Then Text 113, dlg_vert + 1, 10, 10, "11"
	If page = 12 Then Text 125, dlg_vert + 1, 10, 10, "12"
	If page = 13 Then Text 137, dlg_vert + 1, 10, 10, "13"
	If page = 14 Then Text 149, dlg_vert + 1, 10, 10, "14"
	If page = 15 Then Text 161, dlg_vert + 1, 10, 10, "15"
	If page = 16 Then Text 173, dlg_vert + 1, 10, 10, "16"
	If page = 17 Then Text 185, dlg_vert + 1, 10, 10, "17"
	If page = 18 Then Text 197, dlg_vert + 1, 10, 10, "18"

end function

'These are defined in the function above. They have to be defined outside of the function so we don't break all the things.
'These should be moved to the MASTER FUNCTION LIBRARY when the function is.
Dim page_one_btn, page_two_btn, page_three_btn, page_four_btn, page_five_btn, page_six_btn, page_seven_btn, page_eight_btn, page_nine_btn, page_ten_btn, page_eleven_btn, page_twelve_btn, page_thirteen_btn, page_fourteen_btn, page_fifteen_btn, page_sixteen_btn, page_seventeen_btn, page_eightteen_btn

user_is_tester = False
user_is_QI = False
user_is_BZ = False

For each tester in tester_array                         'looping through all of the testers
	If user_ID_for_validation = tester.tester_id_number Then             'If the person who is running the script is a tester
		user_is_tester = True
		If tester.tester_population = "BZ" Then user_is_BZ = True
		For each grp in tester.tester_groups
			If grp = "QI" Then user_is_QI = True
		Next
	End If
Next
MsgBox "Tester - " & user_is_tester & vbCr & "QI - " & user_is_QI & vbCr & "BZ - " & user_is_BZ

excel_created = FALSE           'setting this boolean at the beginning - this will later determine if an excel workbook is already open if exporting more than once

script_selection = "Select One..."          'Defaulting the script selection for the first run of the dialog

button_counter = 5001
For each script_item in script_array      'now we look at each script
	script_item.script_btn_one = button_counter
	button_counter = button_counter + 1
Next

page = 1            'defining the page we are starting on so everything actually works
script_selection = "All"
Do
    dlg_len = 80                'set the dialog length to start - this will be adjusted as the script reads the array
    dlg_width = 815             'This is how wide the dialog usually is
    button_pos = 630            'This is where the bottom 3 buttons would be (Export, Search, and Done)
    If user_is_BZ = False Then     'For searching for testing information, the dialog will be wider and the buttons will be more right... which may not fit
        dlg_width = 770
        button_pos = 585
    End If

    old_detail = detail_edit        'this saves the detail of the search criteria from the last run to see if it changed.
    total_scripts = 0               'setting the number of scripts at the beginning of each loop
    script_counter = 0              'setting the start of the counter at the beginning of each loop

    detail_operator = ""                    'Maybe we want to be able to select and or or when listing options. Discussion with MiKayla and Ilse'
    'The details of the search criteria need to be made into an array - even if there is only one thing listed because we have to loop through them
    If Instr(detail_edit, ",") <> 0 Then
        detail_array = split(detail_edit, ",")
    ' ElseIf Instr(detail_edit, "AND") <> 0 Then        'We may want to choose if wmultiple criteria should be inclusive or exclusive
    '     detail_array = split(detail_edit, "AND")
    '     detail_operator = "AND"
    ' ElseIf Instr(detail_edit, "OR") <> 0 Then
    '     detail_array = split(detail_edit, "OR")
    '     detail_operator = "OR"
    Else
        detail_array = ARRAY(detail_edit)   'this makes a single thing an array'
    End If

    'Now we have to loop through all of the scripts in the list of scripts to see if they meet the selected criteria
    For each script_item in script_array
        script_item.show_script = FALSE     'this determines if the script should be displayed in the dialog and is set to false to start with every time. (This is a class property also used for testing)
		qi_only_script = False
		bz_only_script = False
		for each tag in script_item.tags
			If tag = "QI" Then qi_only_script = True
			If tag = "BZ" Then bz_only_script = True
			If tag = "Monthly Tasks" Then bz_only_script = True
		Next
		is_script_retired = False
		If IsDate(script_item.retirement_date) = True Then
			If DateDiff("d", script_item.retirement_date, date) >= 0 Then is_script_retired = True
		End If
		If script_item.category = "" Then script_item.category = "NOTICES"
		If is_script_retired = False OR user_is_BZ = True Then
			Select Case script_selection        'These are all of the options for how we sort through the scripts
	            Case "All"  'All scripts listed
					If user_is_BZ = True Then
		                dlg_len = dlg_len + 20                  'Make the dialog larger
		                total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
						script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
					Else
						If script_item.in_testing <> True AND qi_only_script = False AND bz_only_script = False Then
							dlg_len = dlg_len + 20                  'Make the dialog larger
							total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
							script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
						ElseIf script_item.in_testing = True AND user_is_tester = True Then
							dlg_len = dlg_len + 20                  'Make the dialog larger
			                total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
							script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
						ElseIf qi_only_script = True AND user_is_QI = True Then
							dlg_len = dlg_len + 20                  'Make the dialog larger
							total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
							script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
						End If
					End If
	            Case "All in Testing"       'Any scripts that indicate 'testing' is true
	                If script_item.in_testing = TRUE AND user_is_tester = True Then
	                    dlg_len = dlg_len + 20                  'Make the dialog larger
	                    total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
	                    script_item.show_script = TRUE
	                End If
	            Case "Tags"           'Search based upon tags listed for the script'
	                For each tag_to_see in detail_array     'looking at each tag listed in the dialog selection
	                    tag_to_see = trim(tag_to_see)       'taking the spaces from the front and back of the listed tag
	                    tag_to_see = UCase(tag_to_see)      'making the tag uppercase
	                    For each script_tag in script_item.tags     'Now we look at each of the tags listed for the script
	                        script_tag = trim(script_tag)
	                        script_tag = UCase(script_tag)
	                        ' MsgBox script_item.script_name & vbNewLine & "Detail Edit - " & detail_edit & vbNewLine & "Tag to see - " & tag_to_see & vbNewLine & "Script tag - " & script_tag
	                        If script_tag = tag_to_see Then             'If the tag listed in the script array matches the one indicated in the dialog - we want to show this script
	                            dlg_len = dlg_len + 20                  'Make the dialog larger
	                            total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
	                            script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
	                        End If
	                    Next
	                Next
	            Case "Key Codes"        'Seach based upon specific keys
	                For each key_code_to_see in detail_array        'Looking at each key listed in the dialog selection'
	                    key_code_to_see = trim(key_code_to_see)     'taking the spaces from the front and back of the listed key
	                    key_code_to_see = UCase(key_code_to_see)    'making the key uppercase
	                    For each script_key_code in script_item.dlg_keys     'Now we look at each of the keys listed for the script
	                        script_key_code = trim(script_key_code)
	                        script_key_code = UCase(script_key_code)
	                        ' MsgBox script_item.script_name & vbNewLine & "Detail Edit - " & detail_edit & vbNewLine & "Tag to see - " & tag_to_see & vbNewLine & "Script tag - " & script_tag
	                        If script_key_code = key_code_to_see Then   'If the key code listed in the script array matches the one indicated in the dialog - we want to show this script
								If user_is_BZ = True Then
					                dlg_len = dlg_len + 20                  'Make the dialog larger
					                total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
									script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
								Else
									If script_item.in_testing <> True AND qi_only_script = False AND bz_only_script = False Then
										dlg_len = dlg_len + 20                  'Make the dialog larger
										total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
										script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
									ElseIf script_item.in_testing = True AND user_is_tester = True Then
										dlg_len = dlg_len + 20                  'Make the dialog larger
						                total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
										script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
									ElseIf qi_only_script = True AND user_is_QI = True Then
										dlg_len = dlg_len + 20                  'Make the dialog larger
										total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
										script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
									End If
								End If
	                        End If
	                    Next
	                Next
	            Case "Category"         'Select the scripts by category
	                For each category_to_see in detail_array
	                    category_to_see = trim(category_to_see)
	                    category_to_see = UCase(category_to_see)
	                    If category_to_see = script_item.category Then  'If the category listed in the script array matches the one indicated in the dialog - we want to show this script
							If user_is_BZ = True Then
				                dlg_len = dlg_len + 20                  'Make the dialog larger
				                total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
								script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
							Else
								If script_item.in_testing <> True AND qi_only_script = False AND bz_only_script = False Then
									dlg_len = dlg_len + 20                  'Make the dialog larger
									total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
									script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
								ElseIf script_item.in_testing = True AND user_is_tester = True Then
									dlg_len = dlg_len + 20                  'Make the dialog larger
					                total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
									script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
								ElseIf qi_only_script = True AND user_is_QI = True Then
									dlg_len = dlg_len + 20                  'Make the dialog larger
									total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
									script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
								End If
							End If
	                    End If
	                Next
	            Case "Subcategory"          'Select based upon subcategories
	                For each subcategory_to_see in detail_array
	                    subcategory_to_see = trim(subcategory_to_see)
	                    subcategory_to_see = UCase(subcategory_to_see)
	                    For each script_subcategory in script_item.subcategory
	                        script_subcategory = trim(script_subcategory)
	                        script_subcategory = UCase(script_subcategory)
	                        ' MsgBox script_item.script_name & vbNewLine & "Detail Edit - " & detail_edit & vbNewLine & "Tag to see - " & tag_to_see & vbNewLine & "Script tag - " & script_tag
	                        If script_subcategory = subcategory_to_see Then     'If the subcategory listed in the script array matches the one indicated in the dialog - we want to show this script
								If user_is_BZ = True Then
					                dlg_len = dlg_len + 20                  'Make the dialog larger
					                total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
									script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
								Else
									If script_item.in_testing <> True AND qi_only_script = False AND bz_only_script = False Then
										dlg_len = dlg_len + 20                  'Make the dialog larger
										total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
										script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
									ElseIf script_item.in_testing = True AND user_is_tester = True Then
										dlg_len = dlg_len + 20                  'Make the dialog larger
						                total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
										script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
									ElseIf qi_only_script = True AND user_is_QI = True Then
										dlg_len = dlg_len + 20                  'Make the dialog larger
										total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
										script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
									End If
								End If
	                        End If
	                    Next
	                Next
	            Case "Release Before"           'Select if release date is before specified date
	                If IsDate(script_item.release_date) = TRUE Then     'If there is a valid date listed in the array of scripts we can compare dates
	                    If DateDiff("d", detail_edit, script_item.release_date) < 0 Then        'if the date listed in the script array is before the one listed in the dialog, this comparisson will be negative
							If user_is_BZ = True Then
				                dlg_len = dlg_len + 20                  'Make the dialog larger
				                total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
								script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
							Else
								If script_item.in_testing <> True AND qi_only_script = False AND bz_only_script = False Then
									dlg_len = dlg_len + 20                  'Make the dialog larger
									total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
									script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
								ElseIf script_item.in_testing = True AND user_is_tester = True Then
									dlg_len = dlg_len + 20                  'Make the dialog larger
					                total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
									script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
								ElseIf qi_only_script = True AND user_is_QI = True Then
									dlg_len = dlg_len + 20                  'Make the dialog larger
									total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
									script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
								End If
							End If
	                    End If
	                End If
	            Case "Release After"            'Select if release date is after specified date'
	                If IsDate(script_item.release_date) = TRUE Then     'If there is a valid date listed in the array of scripts we can compare dates
	                    If DateDiff("d", detail_edit, script_item.release_date) > 0 Then        'If the date listed in the dialog is before the one in the script array, this comparrison would be positive
							If user_is_BZ = True Then
				                dlg_len = dlg_len + 20                  'Make the dialog larger
				                total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
								script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
							Else
								If script_item.in_testing <> True AND qi_only_script = False AND bz_only_script = False Then
									dlg_len = dlg_len + 20                  'Make the dialog larger
									total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
									script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
								ElseIf script_item.in_testing = True AND user_is_tester = True Then
									dlg_len = dlg_len + 20                  'Make the dialog larger
					                total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
									script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
								ElseIf qi_only_script = True AND user_is_QI = True Then
									dlg_len = dlg_len + 20                  'Make the dialog larger
									total_scripts = total_scripts + 1       'increase the number of total scripts that meet the requirement
									script_item.show_script = TRUE          'Set this to TRUE so that when we loop in the dialog, the script knows to show this one
								End If
							End If
	                    End If
	                End If
	        End Select
			If script_item.show_script = True Then
				If script_item.category = "NAV" and user_is_BZ = False Then
					script_item.show_script = False
					dlg_len = dlg_len - 20
					total_scripts = total_scripts - 1
				End If
				If script_item.category = "DAIL" AND script_item.script_name <> "DAIL Scrubber" and user_is_BZ = False Then
					script_item.show_script = False
					dlg_len = dlg_len - 20
					total_scripts = total_scripts - 1
				End If
			End If
		End If
    Next

    If dlg_len > 385 Then dlg_len = 385     'If there are a LOT of scripts (more than 15) this number could get really bit and be too tall for the monitor so we need to reset it.
    If dlg_len = 80 Then dlg_len = 100      'If there are NO scripts that will be displayed the original dialog height is a little small and we need to reset it.

    Dialog1 = ""        'because we use all the dialogs we reset this before degining a dialog
    'Now comes the dialog creation - which is what this is all about really
    BeginDialog Dialog1, 0, 0, dlg_width, dlg_len, "Detailed Script Information"
      ButtonGroup ButtonPressed
		Text 10, 15, 170, 10, "Select what information you want to review/gather."        'These are the selection parts
		Text 190, 15, 55, 10, "Script Selection:"
		If user_is_tester = False Then DropListBox 260, 10, 130, 45, "Select One..."+chr(9)+"All"+chr(9)+"Tags"+chr(9)+"Key Codes"+chr(9)+"Category"+chr(9)+"Subcategory"+chr(9)+"Release Before"+chr(9)+"Release After", script_selection
		If user_is_tester = True Then DropListBox 260, 10, 130, 45, "Select One..."+chr(9)+"All"+chr(9)+"All in Testing"+chr(9)+"Tags"+chr(9)+"Key Codes"+chr(9)+"Category"+chr(9)+"Subcategory"+chr(9)+"Release Before"+chr(9)+"Release After", script_selection
		Text 400, 15, 30, 10, "which is:"
		EditBox 440, 10, 145, 15, detail_edit
		Text 600, 15, 110, 10, "Scripts Found: " & total_scripts
		Text 445, 30, 95, 10, "If a list, separate by commas"

		Text 10, 50, 45, 10, "Script Name"                        'This is the script header information
		Text 145, 50, 40, 10, "Description"
		Text 390, 50, 20, 10, "Tags"
		Text 535, 50, 40, 10, "Key Codes"
		Text 590, 50, 50, 10, "Subcategories"
		Text 670, 50, 30, 10, "Release"
		Text 715, 50, 35, 10, "Hot Topic"
		If user_is_BZ = True Then Text 760, 50, 35, 10, "Retirement"
		If user_is_BZ = True Then GroupBox 665, 40, 145, 20, "Dates"
		If user_is_BZ = False Then GroupBox 665, 40, 100, 20, "Dates"
		Text 815, 50, 50, 10, "Keywords"
		If script_selection = "All in Testing" Then Text 875, 50, 100, 10, "Testing Type and criteria"        'this will be off the dialog if 'All in Testing' has not been selected so we only show it if that is the selection

		y_pos = 65        'This is the incrementer
		For each script_item in script_array      'now we look at each script
		  skip_this_script = FALSE              'we are going to assume that we aren't skipping a script as we look at each one
		  'This will determine where in the list we should start displaying the scripts, because if we are on a later page, we don't want the first 15 to show.
		  'Because we do not increase the counter until AFTER we look at the script and we start at 0, we use the number from where the previous page left off to figure where we should skip
		  If page = 2 and script_counter < 15 Then skip_this_script = TRUE      'any script that is counted 0 - 14 should be skipped if we are on page 2
		  If page = 3 and script_counter < 30 Then skip_this_script = TRUE      'any script that is counted 0 - 29 should be skipped if we are on page 3
		  If page = 4 and script_counter < 45 Then skip_this_script = TRUE      'any script that is counted 0 - 44 should be skipped if we are on page 4
		  If page = 5 and script_counter < 60 Then skip_this_script = TRUE      'any script that is counted 0 - 59 should be skipped if we are on page 5
		  If page = 6 and script_counter < 75 Then skip_this_script = TRUE      'any script that is counted 0 - 74 should be skipped if we are on page 6
		  If page = 7 and script_counter < 90 Then skip_this_script = TRUE      'any script that is counted 0 - 89 should be skipped if we are on page 7
		  If page = 8 and script_counter < 105 Then skip_this_script = TRUE     'any script that is counted 0 - 104 should be skipped if we are on page 8
		  If page = 9 and script_counter < 120 Then skip_this_script = TRUE     'any script that is counted 0 - 119 should be skipped if we are on page 9
		  If page = 10 and script_counter < 135 Then skip_this_script = TRUE    'any script that is counted 0 - 134 should be skipped if we are on page 10
		  If page = 11 and script_counter < 150 Then skip_this_script = TRUE    'any script that is counted 0 - 149 should be skipped if we are on page 11
		  If page = 12 and script_counter < 165 Then skip_this_script = TRUE    'any script that is counted 0 - 164 should be skipped if we are on page 12
		  If page = 13 and script_counter < 180 Then skip_this_script = TRUE    'any script that is counted 0 - 179 should be skipped if we are on page 13
		  If page = 14 and script_counter < 195 Then skip_this_script = TRUE    'any script that is counted 0 - 194 should be skipped if we are on page 14
		  If page = 15 and script_counter < 210 Then skip_this_script = TRUE    'any script that is counted 0 - 194 should be skipped if we are on page 14
		  If page = 16 and script_counter < 225 Then skip_this_script = TRUE    'any script that is counted 0 - 194 should be skipped if we are on page 14
		  If page = 17 and script_counter < 240 Then skip_this_script = TRUE    'any script that is counted 0 - 194 should be skipped if we are on page 14
		  If page = 18 and script_counter < 255 Then skip_this_script = TRUE    'any script that is counted 0 - 194 should be skipped if we are on page 14

		  If script_item.show_script = TRUE Then            'If the logic above inidcates this is a script that meets the criteria then we will show this script
		      If skip_this_script = TRUE Then               'If the above inidcates we should skip this one due to which page we are on then the dialog won't list
		          script_counter = script_counter + 1       'Still need to increment or we ALWAYS be on counter 0
		      Else
		          ' MsgBox "BEFORE" & vbNewLine & "Page - " & page & vbNewLine & "Script COunter - " & script_counter
				  If user_is_BZ = False Then PushButton 5, y_pos-2, 10, 13, "?", script_item.script_btn_one
		          If script_item.in_testing = TRUE Then     'If the script is in testing, we add that detail to the name so we can tell
		              Text 17, y_pos, 120, 20, "TESTING - " & script_item.category & " - " & script_item.script_name
				  ElseIf script_item.category = "" Then
				  	  Text 17, y_pos, 120, 20, script_item.script_name
		          Else
		              Text 17, y_pos, 120, 20, script_item.category & " - " & script_item.script_name
		          End If
				  display_description = replace(script_item.description, "IN TESTING - ", "")
				  ' If script_item.in_testing = TRUE Then MsgBox display_description & vbCr & "~" & left(display_description, 4) & "~"
				  If left(display_description, 4) = "--- " Then display_description = right(display_description, len(display_description) - 4)
				  display_description = trim(display_description)
		          Text 145, y_pos, 235, 20, display_description
		          all_the_tags = join(script_item.tags, ", ")           'Displaying the array as a string
				  If Instr(all_the_tags, "SNAP") <> 0 AND Instr(all_the_tags, "MFIP") <> 0 AND Instr(all_the_tags, "DWP") <> 0 AND Instr(all_the_tags, "HS/GRH") <> 0 AND Instr(all_the_tags, "Adult Cash") <> 0 AND Instr(all_the_tags, "Health Care") <> 0 AND Instr(all_the_tags, "EMER") <> 0 Then
					  all_the_tags = replace(all_the_tags, ", SNAP", "")
					  all_the_tags = replace(all_the_tags, ", MFIP", "")
					  all_the_tags = replace(all_the_tags, ", DWP", "")
					  all_the_tags = replace(all_the_tags, ", HS/GRH", "")
					  all_the_tags = replace(all_the_tags, ", Adult Cash", "")
					  all_the_tags = replace(all_the_tags, ", Health Care", "")
					  all_the_tags = replace(all_the_tags, ", EMER", "")
					  all_the_tags = replace(all_the_tags, ", LTC", "")
					  all_the_tags = replace(all_the_tags, "Adult Cash, ", "")
					  all_the_tags = all_the_tags + ", All Programs"
				  ElseIf Instr(all_the_tags, "MFIP") <> 0 AND Instr(all_the_tags, "DWP") <> 0 AND Instr(all_the_tags, "Adult Cash") <> 0 Then
					  all_the_tags = replace(all_the_tags, ", MFIP", "")
					  all_the_tags = replace(all_the_tags, ", DWP", "")
					  all_the_tags = replace(all_the_tags, ", Adult Cash", "")
					  all_the_tags = all_the_tags + ", All Cash"
				  End If
				  all_the_tags = replace(all_the_tags, "Application", "APPL")
				  all_the_tags = replace(all_the_tags, "Health Care", "HC")
				  all_the_tags = replace(all_the_tags, "Reviews", "REVW")
				  all_the_tags = replace(all_the_tags, "Reviews", "REVW")
		          Text 390, y_pos, 140, 20, all_the_tags

		          all_the_keys = join(script_item.dlg_keys, ", ")       'Displaying the array as a string
		          Text 535, y_pos, 50, 20, all_the_keys

		          all_the_subcats = join(script_item.subcategory, ", ")     'Displaying the array as a strink
		          Text 590, y_pos, 75, 20, all_the_subcats
				  If DateDiff("d", script_item.release_date, #10/01/2000#) = 0 Then
					  Text 670, y_pos, 40, 10, "Pre-2016"
				  Else
			          Text 670, y_pos, 40, 10, script_item.release_date
				  End If
		          Text 715, y_pos, 40, 10, script_item.hot_topic_date
		          Text 760, y_pos, 40, 10, script_item.retirement_date

		          ' all_the_keywords = join(script_item.keywords , ", ")                'This isn't in the complete list yet but when it is - we are ready
		          Text 815, y_pos, 50, 15, all_the_keywords

		          If script_selection = "All in Testing" Then           'Adding more fields if the testing cases are selected

		              If IsArray(script_item.testing_criteria) = TRUE Then
		                all_the_test_criteria = join(script_item.testing_criteria, ", ")
		              Else
		                all_the_test_criteria = ""
		              End If
		              Text 875, y_pos, 100, 10, script_item.testing_category & " - " & all_the_test_criteria
		              ' Text 850, y_pos, 50, 10, all_the_test_criteria

		          End If
		          script_counter = script_counter + 1       'increment the counter so we know where we are'
		          y_pos = y_pos + 20                        'move down in the dialog
		      End If
		  End If

		  'This will determine if we should stop looping through the scripts because we have reached the max of 15 per page
		  If page = 1 and script_counter = 15 Then Exit For
		  If page = 2 and script_counter = 30 Then Exit For
		  If page = 3 and script_counter = 45 Then Exit For
		  If page = 4 and script_counter = 60 Then Exit For
		  If page = 5 and script_counter = 75 Then Exit For
		  If page = 6 and script_counter = 90 Then Exit For
		  If page = 7 and script_counter = 105 Then Exit For
		  If page = 8 and script_counter = 120 Then Exit For
		  If page = 9 and script_counter = 135 Then Exit For
		  If page = 10 and script_counter = 150 Then Exit For
		  If page = 11 and script_counter = 165 Then Exit For
		  If page = 12 and script_counter = 180 Then Exit For
		  If page = 13 and script_counter = 195 Then Exit For
		  If page = 14 and script_counter = 210 Then Exit For
		  If page = 15 and script_counter = 225 Then Exit For
		  If page = 16 and script_counter = 240 Then Exit For
		  If page = 17 and script_counter = 255 Then Exit For

		Next

		If y_pos = 65 Then y_pos = 75     'If there were no scripts, we need to move the buttons down a little

        call add_page_buttons_to_dialog(page, 15, total_scripts, y_pos)     'This is the function to call the page buttons - it's like 1000 lines of code because it has possibilities for each page
		' add_page_buttons_to_dialog(page_variable, items_per_page, total_items, dlg_vert)
        PushButton button_pos, y_pos, 70, 15, "Export to EXCEL", export_btn
        PushButton button_pos + 75, y_pos, 50, 15, "Search", search_btn
        PushButton button_pos + 130, y_pos, 50, 15, "DONE", done_btn
    EndDialog

    Dialog Dialog1      'actually displaying the dialog'

    'now we figure out which page we should be at
    page = 1                                                'we start at page 1 always - it will stay at page 1 unless a page button is pushed.
    If ButtonPressed = page_one_btn Then page = 1
    If ButtonPressed = page_two_btn Then page = 2
    If ButtonPressed = page_three_btn Then page = 3
    If ButtonPressed = page_four_btn Then page = 4
    If ButtonPressed = page_five_btn Then page = 5
    If ButtonPressed = page_six_btn Then page = 6
    If ButtonPressed = page_seven_btn Then page = 7
    If ButtonPressed = page_eight_btn Then page = 8
    If ButtonPressed = page_nine_btn Then page = 9
    If ButtonPressed = page_ten_btn Then page = 10
    If ButtonPressed = page_eleven_btn Then page = 11
    If ButtonPressed = page_twelve_btn Then page = 12
    If ButtonPressed = page_thirteen_btn Then page = 13
    If ButtonPressed = page_fourteen_btn Then page = 14
	If ButtonPressed = page_fifteen_btn Then page = 15
	If ButtonPressed = page_sixteen_btn Then page = 16
	If ButtonPressed = page_seventeen_btn Then page = 17
	If ButtonPressed = page_eightteen_btn Then page = 18

	For each script_item in script_array      'now we look at each script
		If ButtonPressed = script_item.script_btn_one Then call open_URL_in_browser(script_item.SharePoint_instructions_URL)
	Next

    If ButtonPressed = 0 Then ButtonPressed = done_btn          'default 'ESC' tp done
    If ButtonPressed = -1 Then ButtonPressed = search_btn       'default 'ENTER' to search

    If old_detail <> detail_edit Then page = 1

    ' MsgBox "The button pressed was - " & ButtonPressed
    'If we select the ones that use dates, we need to make sure the criteria is a date, or the whole thing breaks
    If script_selection = "Release Before" OR script_selection = "Release After" Then       'these are the only options that have date requirements
        If IsDate(detail_edit) = FALSE Then         'if this is NOT a date the script will reset and alert you to the change
            MsgBox "You have selected 'Release Before' or 'Release After' but ahve not provided a date to compare." & vbNewLine & vbNewLine & "The script has defaulted to 'ALL' and you can re-enter the selection and detail. If using a date specific selection be sure to enter a valid date."
            script_selection = "All"
            detail_edit = ""
            ButtonPressed = search_btn
        End If
    End If
    If ButtonPressed = export_btn Then          'If we pressed the button for export to excel - here we go to excel

        'This is a repeat of the logicc above because if someone changes the search information and presses 'Export' instead of 'Search' everything will be wrong
        detail_operator = ""                    'Maybe we want to be able to select and or or when listing options. Discussion with MiKayla and Ilse'
        If Instr(detail_edit, ",") <> 0 Then
            detail_array = split(detail_edit, ",")
        ' ElseIf Instr(detail_edit, "AND") <> 0 Then
        '     detail_array = split(detail_edit, "AND")
        '     detail_operator = "AND"
        ' ElseIf Instr(detail_edit, "OR") <> 0 Then
        '     detail_array = split(detail_edit, "OR")
        '     detail_operator = "OR"
        Else
            detail_array = ARRAY(detail_edit)
        End If

        For each script_item in script_array
            script_item.show_script = FALSE
            Select Case script_selection
                Case "All"
                    dlg_len = dlg_len + 20
                    total_scripts = total_scripts + 1
                    script_item.show_script = TRUE
                Case "All in Testing"
                    If script_item.in_testing = TRUE Then
                        dlg_len = dlg_len + 20
                        total_scripts = total_scripts + 1
                        script_item.show_script = TRUE
                    End If
                Case "Tags"
                    For each tag_to_see in detail_array
                        tag_to_see = trim(tag_to_see)
                        tag_to_see = UCase(tag_to_see)
                        For each script_tag in script_item.tags
                            script_tag = trim(script_tag)
                            script_tag = UCase(script_tag)
                            ' MsgBox script_item.script_name & vbNewLine & "Detail Edit - " & detail_edit & vbNewLine & "Tag to see - " & tag_to_see & vbNewLine & "Script tag - " & script_tag
                            If script_tag = tag_to_see Then
                                dlg_len = dlg_len + 20
                                total_scripts = total_scripts + 1
                                script_item.show_script = TRUE
                            End If
                        Next
                    Next
                Case "Key Codes"
                    For each key_code_to_see in detail_array
                        key_code_to_see = trim(key_code_to_see)
                        key_code_to_see = UCase(key_code_to_see)
                        For each script_key_code in script_item.dlg_keys
                            script_key_code = trim(script_key_code)
                            script_key_code = UCase(script_key_code)
                            ' MsgBox script_item.script_name & vbNewLine & "Detail Edit - " & detail_edit & vbNewLine & "Tag to see - " & tag_to_see & vbNewLine & "Script tag - " & script_tag
                            If script_key_code = key_code_to_see Then
                                dlg_len = dlg_len + 20
                                total_scripts = total_scripts + 1
                                script_item.show_script = TRUE
                            End If
                        Next
                    Next
                Case "Category"
                    For each category_to_see in detail_array
                        category_to_see = trim(category_to_see)
                        category_to_see = UCase(category_to_see)
                        If category_to_see = script_item.category Then
                            dlg_len = dlg_len + 20
                            total_scripts = total_scripts + 1
                            script_item.show_script = TRUE
                        End If
                    Next
                Case "Subcategory"
                    For each subcategory_to_see in detail_array
                        subcategory_to_see = trim(subcategory_to_see)
                        subcategory_to_see = UCase(subcategory_to_see)
                        For each script_subcategory in script_item.subcategory
                            script_subcategory = trim(script_subcategory)
                            script_subcategory = UCase(script_subcategory)
                            ' MsgBox script_item.script_name & vbNewLine & "Detail Edit - " & detail_edit & vbNewLine & "Tag to see - " & tag_to_see & vbNewLine & "Script tag - " & script_tag
                            If script_subcategory = subcategory_to_see Then
                                dlg_len = dlg_len + 20
                                total_scripts = total_scripts + 1
                                script_item.show_script = TRUE
                            End If
                        Next
                    Next
                Case "Release Before"
                    If IsDate(script_item.release_date) = TRUE Then
                        If DateDiff("d", detail_edit, script_item.release_date) < 0 Then
                            dlg_len = dlg_len + 20
                            total_scripts = total_scripts + 1
                            script_item.show_script = TRUE
                        End If
                    End If
                Case "Release After"
                    If IsDate(script_item.release_date) = TRUE Then
                        If DateDiff("d", detail_edit, script_item.release_date) > 0 Then
                            dlg_len = dlg_len + 20
                            total_scripts = total_scripts + 1
                            script_item.show_script = TRUE
                        End If
                    End If
            End Select
        Next

        'Now comes the excel bit
        sheet_title = "SCRIPTS sorted " & script_selection              'setting the sheet name
        If excel_created = FALSE Then                                   'If this is the first time in this run we have exported to excel, a new workbook will be opened
            'Opening a new Excel file
            Set ObjExcel = CreateObject("Excel.Application")
            ObjExcel.Visible = True
            Set objWorkbook = ObjExcel.Workbooks.Add()
            ObjExcel.DisplayAlerts = True

            excel_created = TRUE                                        'telling the script this is no longer the first time excel has been used
        Else
            ObjExcel.Worksheets.Add().Name = sheet_title                'If this is NOT the first time excel has been used in this script run, a new worksheet will be added to the existing workbook
        End If

        ObjExcel.ActiveSheet.Name = sheet_title                         'setting the name of the worksheet

        'Here we add the headers
        ObjExcel.Cells(1, 1).Value = "Script Category"
        ObjExcel.Cells(1, 2).Value = "Script Name"
        ObjExcel.Cells(1, 3).Value = "Description"
		ObjExcel.Cells(1, 4).Value = "INSTRUCTIONS"
        ObjExcel.Cells(1, 5).Value = "Tags"
        ObjExcel.Cells(1, 6).Value = "Key Codes"
        ObjExcel.Cells(1, 7).Value = "Subcategory"
        ObjExcel.Cells(1, 8).Value = "Keywords"
        ObjExcel.Cells(1, 9).Value = "Release Date"
        ObjExcel.Cells(1, 10).Value = "Hot Topic Date"
        If user_is_tester = True Then
			ObjExcel.Cells(1, 11).Value = "In Testing"
	        ObjExcel.Cells(1, 12).Value = "Testing Category"
	        ObjExcel.Cells(1, 13).Value = "Testing Criteria"
		End If
		If user_is_BZ = True Then ObjExcel.Cells(1, 14).Value = "Retired Date"
        'ADD MORE PROPERTIES HERE - these more properties will likely NOT be on the dialog

        ObjExcel.Rows(1).Font.Bold = TRUE           'format the headers to be bold

        row_to_use = 2                              'start at row 2 with information

        For each script_item in script_array        'look at each script
            If script_item.show_script = TRUE Then  'if in the logic above they have been determined to meet the critera this will be set to true and we will add the detail of that script to the excel
                ObjExcel.Cells(row_to_use, 1).Value = script_item.category              'this is adding each script that is selected to the excel
                ObjExcel.Cells(row_to_use, 2).Value = script_item.script_name
                ObjExcel.Cells(row_to_use, 3).Value = script_item.description
				ObjExcel.Cells(row_to_use, 4).Value = "=HYPERLINK(" & chr(34) &  script_item.SharePoint_instructions_URL & chr(34) & ", " & chr(34) & script_item.script_name & chr(34) & ")"
				' "=HYPERLINK(""" & sLinkAddress & """,""" & sFriendly & """)"
                ObjExcel.Cells(row_to_use, 5).Value = join(script_item.tags, ", ")
                ObjExcel.Cells(row_to_use, 6).Value = join(script_item.dlg_keys, ", ")
                ObjExcel.Cells(row_to_use, 7).Value = join(script_item.subcategory, ", ")
                ' ObjExcel.Cells(row_to_use, 8).Value = join(script_item.keywords, ", ")
                ObjExcel.Cells(row_to_use, 9).Value = script_item.release_date
                ObjExcel.Cells(row_to_use, 10).Value = script_item.hot_topic_date
				If user_is_tester = True Then
	                ObjExcel.Cells(row_to_use, 11).Value = script_item.in_testing
	                ObjExcel.Cells(row_to_use, 12).Value = script_item.testing_category
	                If IsArray(script_item.testing_criteria) = TRUE Then ObjExcel.Cells(row_to_use, 13).Value = join(script_item.testing_criteria, ", ")
				End If
				If user_is_BZ = True Then ObjExcel.Cells(row_to_use, 14).Value = script_item.retirement_date

                row_to_use = row_to_use + 1     'go to the next row for the next script
            End If
        Next

        'Autofitting columns
        For col_to_autofit = 1 to 14
            ObjExcel.columns(col_to_autofit).AutoFit()
        Next
    End If

Loop until ButtonPressed = done_btn     'The dialog will keep appearing for a different search until you press 'Done' or 'ESC'

Call script_end_procedure("") 'The End
