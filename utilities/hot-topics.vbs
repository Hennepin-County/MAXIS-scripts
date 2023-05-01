'STATS GATHERING--------------------------------------------------------------------------------------------------------------
name_of_script = "UTILITIES - HOT TOPICS.vbs"
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
call changelog_update("06/02/2021", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

Const hot_topic_link_const	= 0
Const hot_topic_date_const 	= 1
Const script_category_const	= 2
Const script_name_const 	= 3
Const article_btn_const		= 4
Const instructions_btn_const= 5
Const list_order_const 		= 6
Const multiple_scripts_TF	= 7
Const run_script_btn 		= 8
Const add_to_favorites_btn	= 9
Const script_displayed 		= 10
Const hot_topic_name_const 	= 12
Const script_instructions_url_const 	= 13
Const script_url_const 		= 14
Const last_const 			= 15

Dim HOT_TOPIC_ARRAY()
ReDim HOT_TOPIC_ARRAY(last_const, 0)
ht_dates = " "

article_count = 0
For current_script = 0 to ubound(script_array)
	If script_array(current_script).hot_topic_date <> "" Then
		' MsgBox script_array(current_script).hot_topic_date
		HT_already_in_list = False
		For known_ht = 0 to UBound(HOT_TOPIC_ARRAY, 2)
			If HOT_TOPIC_ARRAY(hot_topic_link_const, known_ht) = script_array(current_script).hot_topic_link Then
				HOT_TOPIC_ARRAY(multiple_scripts_TF, known_ht) = True
				HT_already_in_list = True
			End If
		Next
		If InStr(ht_dates, script_array(current_script).hot_topic_date) = 0 Then ht_dates = ht_dates & script_array(current_script).hot_topic_date & " "

		ReDim Preserve HOT_TOPIC_ARRAY(last_const, article_count)

		HOT_TOPIC_ARRAY(hot_topic_link_const, article_count) = script_array(current_script).hot_topic_link
		HOT_TOPIC_ARRAY(hot_topic_name_const, article_count) = replace(script_array(current_script).hot_topic_link, "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/", "")
		HOT_TOPIC_ARRAY(hot_topic_name_const, article_count) = replace(HOT_TOPIC_ARRAY(hot_topic_name_const, article_count), ".aspx", "")
		HOT_TOPIC_ARRAY(hot_topic_name_const, article_count) = replace(HOT_TOPIC_ARRAY(hot_topic_name_const, article_count), "-%E2%80%93-", "   ")
		HOT_TOPIC_ARRAY(hot_topic_name_const, article_count) = replace(HOT_TOPIC_ARRAY(hot_topic_name_const, article_count), "-", " ")
		HOT_TOPIC_ARRAY(hot_topic_name_const, article_count) = replace(HOT_TOPIC_ARRAY(hot_topic_name_const, article_count), "   ", " - ")
		HOT_TOPIC_ARRAY(hot_topic_date_const, article_count) = script_array(current_script).hot_topic_date
		HOT_TOPIC_ARRAY(script_category_const, article_count) = script_array(current_script).category
		HOT_TOPIC_ARRAY(script_name_const, article_count) = script_array(current_script).script_name
		HOT_TOPIC_ARRAY(script_instructions_url_const, article_count) = script_array(current_script).SharePoint_instructions_URL
		HOT_TOPIC_ARRAY(script_url_const, article_count) = script_array(current_script).script_URL
		HOT_TOPIC_ARRAY(multiple_scripts_TF, article_count)  = False
		HOT_TOPIC_ARRAY(script_displayed, article_count)  = False
		If HT_already_in_list = True Then HOT_TOPIC_ARRAY(multiple_scripts_TF, article_count) = True

		HOT_TOPIC_ARRAY(article_btn_const, article_count) = 500 + article_count
		HOT_TOPIC_ARRAY(instructions_btn_const, article_count) = 1000 + article_count
		HOT_TOPIC_ARRAY(run_script_btn, article_count) = 1500 + article_count
		HOT_TOPIC_ARRAY(add_to_favorites_btn, article_count) = 2000 + article_count



		article_count = article_count + 1


	End If
Next
ht_dates = trim(ht_dates)
' MsgBox "ht_dates - " & ht_dates
ht_dates_array = split(ht_dates)

bzst_hot_topics_page_btn = 100
report_to_BZST_btn = 200

' MsgBox "UBOUND ht_dates_array - " & UBOUND(ht_dates_array) & vbCr & "UBOUND HOT_TOPIC_ARRAY - " & UBOUND(HOT_TOPIC_ARRAY, 2)

Call sort_dates(ht_dates_array)		'This function takes all the dates in an array and put them in order from oldest to newest

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 570, 95 + ((UBOUND(HOT_TOPIC_ARRAY, 2)+1) * 20), "BlueZone Script Hot Topics"
ButtonGroup ButtonPressed
	Text 10, 10, 415, 20, "These are the most recent Hot Topic Articles from the BlueZone Script Team. From this menu, you can access the article information, run the script, read the instructions for the scipt or add the script to your favorites menu."
	PushButton 400, 15, 160, 15, "Open the BZST Hot Topics SharePoint Page", bzst_hot_topics_page_btn
	Text 25, 50, 50, 10, "Date"
	Text 75, 50, 40, 10, "Article"
	Text 365, 50, 110, 10, "Script (press to run the script)"
	Text 500, 45, 75, 10, "Instructions"
	Text 525, 55, 30, 10, "Favorite"

	y_pos = 70
	For each hot_topic_date in ht_dates_array
		hot_topic_date = DateAdd("d", 0, hot_topic_date)
		Text 25, y_pos, 50, 10, hot_topic_date
		For article = 0 to UBound(HOT_TOPIC_ARRAY, 2)
			HOT_TOPIC_ARRAY(hot_topic_date_const, article) = DateAdd("d", 0, HOT_TOPIC_ARRAY(hot_topic_date_const, article))
			' MsgBox "hot_topic_date - " & hot_topic_date & vbCr & "HOT_TOPIC_ARRAY(hot_topic_link_const, article) - " & HOT_TOPIC_ARRAY(hot_topic_link_const, article) & vbCr & "HOT_TOPIC_ARRAY(hot_topic_date_const, article) - " & HOT_TOPIC_ARRAY(hot_topic_date_const, article) & vbCr & "HOT_TOPIC_ARRAY(script_displayed, article) - " & HOT_TOPIC_ARRAY(script_displayed, article) & vbCr & "y_pos - " & y_pos
			If DateDiff("d", hot_topic_date, HOT_TOPIC_ARRAY(hot_topic_date_const, article)) = 0 and HOT_TOPIC_ARRAY(script_displayed, article)  = False Then
				' Text 20, y_pos, 55, 10, HOT_TOPIC_ARRAY(hot_topic_date_const, article)
				' MsgBox "y_pos - " & y_pos
				PushButton 75, y_pos-3, 280, 13, HOT_TOPIC_ARRAY(hot_topic_name_const, article), HOT_TOPIC_ARRAY(article_btn_const, article)
				PushButton 365, y_pos-3, 140, 13, HOT_TOPIC_ARRAY(script_category_const, article) & " - " & HOT_TOPIC_ARRAY(script_name_const, article), HOT_TOPIC_ARRAY(run_script_btn, article)
				PushButton 510, y_pos-3, 15, 15, "?", HOT_TOPIC_ARRAY(instructions_btn_const, article)
				PushButton 530, y_pos-3, 15, 15, "+", HOT_TOPIC_ARRAY(add_to_favorites_btn, article)
				y_pos = y_pos + 20
				HOT_TOPIC_ARRAY(script_displayed, article)  = True

				If HOT_TOPIC_ARRAY(multiple_scripts_TF, article) = True Then
					For second_article = 0 to UBound(HOT_TOPIC_ARRAY, 2)
						If HOT_TOPIC_ARRAY(hot_topic_link_const, second_article) = HOT_TOPIC_ARRAY(hot_topic_link_const, article) and HOT_TOPIC_ARRAY(script_displayed, second_article)  = False Then
							' MsgBox "y_pos - " & y_pos & " SECOND"
							PushButton 365, y_pos-3, 140, 13, HOT_TOPIC_ARRAY(script_category_const, second_article) & " - " & HOT_TOPIC_ARRAY(script_name_const, second_article), HOT_TOPIC_ARRAY(run_script_btn, second_article)
							PushButton 510, y_pos-3, 15, 15, "?", HOT_TOPIC_ARRAY(instructions_btn_const, second_article)
							PushButton 530, y_pos-3, 15, 15, "+", HOT_TOPIC_ARRAY(add_to_favorites_btn, second_article)
							y_pos = y_pos + 20
							HOT_TOPIC_ARRAY(script_displayed, second_article)  = True
						End if
					Next
				End If
			End If
		Next
	Next
	GroupBox 10, 35, 555, 35 + ((UBOUND(HOT_TOPIC_ARRAY, 2)+1) * 20), "Hot Topics List"
	y_pos = y_pos + 10

	Text 10, y_pos, 140, 10, "Do you have another question or an idea?"
	PushButton 150, y_pos-5, 95, 15, "Report to the BZST", report_to_BZST_btn
	OkButton 460, y_pos-5, 50, 15
	CancelButton 515, y_pos-5, 50, 15
EndDialog


Do

	dialog Dialog1
	cancel_without_confirmation

	If ButtonPressed = bzst_hot_topics_page_btn Then Call open_URL_in_browser("https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/BlueZone-Scripts.aspx")
	If ButtonPressed = report_to_BZST_btn Then Call run_from_GitHub(script_repository & "utilities/report-to-the-bzst.vbs")

	For article = 0 to UBound(HOT_TOPIC_ARRAY, 2)
		If ButtonPressed = HOT_TOPIC_ARRAY(article_btn_const, article) Then Call open_URL_in_browser(HOT_TOPIC_ARRAY(hot_topic_link_const, article))
		If ButtonPressed = HOT_TOPIC_ARRAY(run_script_btn, article) Then Call run_from_GitHub(HOT_TOPIC_ARRAY(script_url_const, article))
		If ButtonPressed = HOT_TOPIC_ARRAY(instructions_btn_const, article) Then Call open_URL_in_browser(HOT_TOPIC_ARRAY(script_instructions_url_const, article))
		If ButtonPressed = HOT_TOPIC_ARRAY(add_to_favorites_btn, article) Then
			For i = 0 to ubound(script_array)
				If script_array(i).script_URL = HOT_TOPIC_ARRAY(script_url_const, article) Then
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
				End if
			Next
		End If
	Next
Loop until ButtonPressed = OK

'Script ends
script_end_procedure("")
