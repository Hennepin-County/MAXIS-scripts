'Required for statistical purposes==========================================================================================
name_of_script = "UTILITIES - HOT TOPICS.vbs"
start_time = timer
STATS_counter = 1					'sets the stats counter at one
STATS_manualtime = 60				'manual run time in seconds
STATS_denomination = "I"			'C is for each CASE
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("06/02/2021", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'This is the formatting to turn the hot topic URL into a more readable name of the article
function find_hot_topic_name(ht_link, ht_name)
	ht_name = replace(ht_link, "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/", "")
	ht_name = replace(ht_name, "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/sitepages/", "")
	ht_name = replace(ht_name, ".aspx", "")
	ht_name = replace(ht_name, "-%E2%80%93-", "   ")
	ht_name = replace(ht_name, "-%e2%80%93-", "   ")
	ht_name = replace(ht_name, "%20%E2%80%93%20", " - ")
	ht_name = replace(ht_name, "-", " ")
	ht_name = replace(ht_name, "   ", " - ")
	ht_name = replace(ht_name, "%20", " ")
	ht_name = replace(ht_name, "20%", " ")
end function

'declaring some constants for the array of the hot topic articles we are going to use
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
Const hot_topic_name_const 	= 11
Const script_instructions_url_const = 12
Const script_url_const 		= 13
Const last_ht_const 		= 14

'declaring the array
Dim HOT_TOPIC_ARRAY()
ReDim HOT_TOPIC_ARRAY(last_ht_const, 0)
ht_dates = " "				'starting this string with a space as that is what we are going to use as a delimiter.

article_count = 0			'This is the incrementor we are going to use to add to the array

'Here we manually add any Hot Topic articles that are not tied to a specific script (or scripts) and are more general.
'There isn't another place to store these and they will just need to be maintained in this script.
ReDim Preserve HOT_TOPIC_ARRAY(last_ht_const, article_count)
HOT_TOPIC_ARRAY(hot_topic_link_const, article_count) = "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/sitepages/COLA-Processing-with-Bluezone-Scripts.aspx"
Call find_hot_topic_name(HOT_TOPIC_ARRAY(hot_topic_link_const, article_count), HOT_TOPIC_ARRAY(hot_topic_name_const, article_count))
HOT_TOPIC_ARRAY(hot_topic_date_const, article_count) = #12/08/2020#
HOT_TOPIC_ARRAY(multiple_scripts_TF, article_count)  = False
HOT_TOPIC_ARRAY(script_displayed, article_count)  = False
HOT_TOPIC_ARRAY(article_btn_const, article_count) = 500 + article_count
If InStr(ht_dates, HOT_TOPIC_ARRAY(hot_topic_date_const, article_count)) = 0 Then ht_dates = ht_dates & HOT_TOPIC_ARRAY(hot_topic_date_const, article_count) & " "
article_count = article_count + 1

ReDim Preserve HOT_TOPIC_ARRAY(last_ht_const, article_count)
HOT_TOPIC_ARRAY(hot_topic_link_const, article_count) = "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/sitepages/Script-Highlight-%e2%80%93-Postponed-Case-Actions.aspx"
Call find_hot_topic_name(HOT_TOPIC_ARRAY(hot_topic_link_const, article_count), HOT_TOPIC_ARRAY(hot_topic_name_const, article_count))
HOT_TOPIC_ARRAY(hot_topic_date_const, article_count) = #01/26/2021#
HOT_TOPIC_ARRAY(multiple_scripts_TF, article_count)  = False
HOT_TOPIC_ARRAY(script_displayed, article_count)  = False
HOT_TOPIC_ARRAY(article_btn_const, article_count) = 500 + article_count
If InStr(ht_dates, HOT_TOPIC_ARRAY(hot_topic_date_const, article_count)) = 0 Then ht_dates = ht_dates & HOT_TOPIC_ARRAY(hot_topic_date_const, article_count) & " "
article_count = article_count + 1

ReDim Preserve HOT_TOPIC_ARRAY(last_ht_const, article_count)
HOT_TOPIC_ARRAY(hot_topic_link_const, article_count) = "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/sitepages/New-Bluezone-Script-Installer-Applu.aspx"
Call find_hot_topic_name(HOT_TOPIC_ARRAY(hot_topic_link_const, article_count), HOT_TOPIC_ARRAY(hot_topic_name_const, article_count))
HOT_TOPIC_ARRAY(hot_topic_date_const, article_count) = #04/06/2021#
HOT_TOPIC_ARRAY(multiple_scripts_TF, article_count)  = False
HOT_TOPIC_ARRAY(script_displayed, article_count)  = False
HOT_TOPIC_ARRAY(article_btn_const, article_count) = 500 + article_count
If InStr(ht_dates, HOT_TOPIC_ARRAY(hot_topic_date_const, article_count)) = 0 Then ht_dates = ht_dates & HOT_TOPIC_ARRAY(hot_topic_date_const, article_count) & " "
article_count = article_count + 1

ReDim Preserve HOT_TOPIC_ARRAY(last_ht_const, article_count)
HOT_TOPIC_ARRAY(hot_topic_link_const, article_count) = "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/sitepages/The-Bluezone-Scripts-Power-Pad-got-a-New-Look.aspx"
Call find_hot_topic_name(HOT_TOPIC_ARRAY(hot_topic_link_const, article_count), HOT_TOPIC_ARRAY(hot_topic_name_const, article_count))
HOT_TOPIC_ARRAY(hot_topic_date_const, article_count) = #06/08/2021#
HOT_TOPIC_ARRAY(multiple_scripts_TF, article_count)  = False
HOT_TOPIC_ARRAY(script_displayed, article_count)  = False
HOT_TOPIC_ARRAY(article_btn_const, article_count) = 500 + article_count
If InStr(ht_dates, HOT_TOPIC_ARRAY(hot_topic_date_const, article_count)) = 0 Then ht_dates = ht_dates & HOT_TOPIC_ARRAY(hot_topic_date_const, article_count) & " "
article_count = article_count + 1

ReDim Preserve HOT_TOPIC_ARRAY(last_ht_const, article_count)
HOT_TOPIC_ARRAY(hot_topic_link_const, article_count) = "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/sitepages/Bluezone-Scripts-News-and-Updates.aspx"
Call find_hot_topic_name(HOT_TOPIC_ARRAY(hot_topic_link_const, article_count), HOT_TOPIC_ARRAY(hot_topic_name_const, article_count))
HOT_TOPIC_ARRAY(hot_topic_date_const, article_count) = #08/17/2021#
HOT_TOPIC_ARRAY(multiple_scripts_TF, article_count)  = False
HOT_TOPIC_ARRAY(script_displayed, article_count)  = False
HOT_TOPIC_ARRAY(article_btn_const, article_count) = 500 + article_count
If InStr(ht_dates, HOT_TOPIC_ARRAY(hot_topic_date_const, article_count)) = 0 Then ht_dates = ht_dates & HOT_TOPIC_ARRAY(hot_topic_date_const, article_count) & " "
article_count = article_count + 1

ReDim Preserve HOT_TOPIC_ARRAY(last_ht_const, article_count)
HOT_TOPIC_ARRAY(hot_topic_link_const, article_count) = "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/sitepages/Customized-Access-to-Scripts.aspx"
Call find_hot_topic_name(HOT_TOPIC_ARRAY(hot_topic_link_const, article_count), HOT_TOPIC_ARRAY(hot_topic_name_const, article_count))
HOT_TOPIC_ARRAY(hot_topic_date_const, article_count) = #10/12/2021#
HOT_TOPIC_ARRAY(multiple_scripts_TF, article_count)  = False
HOT_TOPIC_ARRAY(script_displayed, article_count)  = False
HOT_TOPIC_ARRAY(article_btn_const, article_count) = 500 + article_count
If InStr(ht_dates, HOT_TOPIC_ARRAY(hot_topic_date_const, article_count)) = 0 Then ht_dates = ht_dates & HOT_TOPIC_ARRAY(hot_topic_date_const, article_count) & " "
article_count = article_count + 1

ReDim Preserve HOT_TOPIC_ARRAY(last_ht_const, article_count)
HOT_TOPIC_ARRAY(hot_topic_link_const, article_count) = "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/April-2023-Staffing-Announcements.aspx"
Call find_hot_topic_name(HOT_TOPIC_ARRAY(hot_topic_link_const, article_count), HOT_TOPIC_ARRAY(hot_topic_name_const, article_count))
HOT_TOPIC_ARRAY(hot_topic_date_const, article_count) = #04/11/2023#
HOT_TOPIC_ARRAY(multiple_scripts_TF, article_count)  = False
HOT_TOPIC_ARRAY(script_displayed, article_count)  = False
HOT_TOPIC_ARRAY(article_btn_const, article_count) = 500 + article_count
If InStr(ht_dates, HOT_TOPIC_ARRAY(hot_topic_date_const, article_count)) = 0 Then ht_dates = ht_dates & HOT_TOPIC_ARRAY(hot_topic_date_const, article_count) & " "
article_count = article_count + 1

ReDim Preserve HOT_TOPIC_ARRAY(last_ht_const, article_count)
HOT_TOPIC_ARRAY(hot_topic_link_const, article_count) = "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/sitepages/Power-Pad-and-Health-Care-Script-Updates.aspx"
Call find_hot_topic_name(HOT_TOPIC_ARRAY(hot_topic_link_const, article_count), HOT_TOPIC_ARRAY(hot_topic_name_const, article_count))
HOT_TOPIC_ARRAY(hot_topic_date_const, article_count) = #04/18/2023#
HOT_TOPIC_ARRAY(multiple_scripts_TF, article_count)  = False
HOT_TOPIC_ARRAY(script_displayed, article_count)  = False
HOT_TOPIC_ARRAY(article_btn_const, article_count) = 500 + article_count
If InStr(ht_dates, HOT_TOPIC_ARRAY(hot_topic_date_const, article_count)) = 0 Then ht_dates = ht_dates & HOT_TOPIC_ARRAY(hot_topic_date_const, article_count) & " "
article_count = article_count + 1

'At this point, we loop through all of the scripts from the Complete List of Scripts
'Hot Topic articles are listed in the the script class and we are going to add the hot topics to the array of all the hot topics.
'NOTE that a HT article might be have more than one instance in this array if it is associated with more than one script, the matching url information is how this script will identify that it is duplicated
For current_script = 0 to ubound(script_array)
	If script_array(current_script).hot_topic_date <> "" Then			'if a hot topic date has been entered in the CLOS, the information will be added to the array
		HT_already_in_list = False										'this is a default for if the listed HT article has already been added to the array for this script
		For known_ht = 0 to UBound(HOT_TOPIC_ARRAY, 2)					'we need to look through all of the HT articles that were already added to see if this link was listed associated with another script
			If HOT_TOPIC_ARRAY(hot_topic_link_const, known_ht) = script_array(current_script).hot_topic_link Then		'if the script matches one already added, we need to note in the array that it is associated
				HOT_TOPIC_ARRAY(multiple_scripts_TF, known_ht) = True													'saving this detail in the array - that this HT article is associated with multiple scripts
				HT_already_in_list = True																				'saving this detail for this loop
			End If
		Next
		'here we need to add the date to a list of all the dates, this will be used to order the display of the HT articles
		If InStr(ht_dates, script_array(current_script).hot_topic_date) = 0 Then ht_dates = ht_dates & script_array(current_script).hot_topic_date & " "

		ReDim Preserve HOT_TOPIC_ARRAY(last_ht_const, article_count)		'resize the array
		HOT_TOPIC_ARRAY(hot_topic_link_const, article_count) = script_array(current_script).hot_topic_link										'save the link to the array
		Call find_hot_topic_name(HOT_TOPIC_ARRAY(hot_topic_link_const, article_count), HOT_TOPIC_ARRAY(hot_topic_name_const, article_count))	'format the article name
		HOT_TOPIC_ARRAY(hot_topic_date_const, article_count) = script_array(current_script).hot_topic_date										'save the date from the CLOS
		HOT_TOPIC_ARRAY(script_category_const, article_count) = script_array(current_script).category											'save the script category associated with the HT Article
		HOT_TOPIC_ARRAY(script_name_const, article_count) = script_array(current_script).script_name											'save the script name associated with the HT Article
		HOT_TOPIC_ARRAY(script_instructions_url_const, article_count) = script_array(current_script).SharePoint_instructions_URL				'save the URL to the instructions for the script
		HOT_TOPIC_ARRAY(script_url_const, article_count) = script_array(current_script).script_URL												'save the URL that can be used to run the script
		HOT_TOPIC_ARRAY(multiple_scripts_TF, article_count)  = False																			'default that this HT article is only for one script.
		HOT_TOPIC_ARRAY(script_displayed, article_count)  = False																				'set that this script isn't displayed, it will change once displayed in the dialog
		If HT_already_in_list = True Then HOT_TOPIC_ARRAY(multiple_scripts_TF, article_count) = True											'if this HT article was already found in the array, changes that indicator for this instance

		HOT_TOPIC_ARRAY(article_btn_const, article_count) = 500 + article_count					'button settings
		HOT_TOPIC_ARRAY(instructions_btn_const, article_count) = 1000 + article_count
		HOT_TOPIC_ARRAY(run_script_btn, article_count) = 1500 + article_count
		HOT_TOPIC_ARRAY(add_to_favorites_btn, article_count) = 2000 + article_count

		article_count = article_count + 1														'increment up for the next loop
	End If
Next
ht_dates = trim(ht_dates)						'formatting the list of dates to remove the spaces on the ends
ht_dates_array = split(ht_dates)				'creating an array of all of the dates for HT articles that were found

bzst_hot_topics_page_btn = 100					'button definitions
report_to_BZST_btn = 200

Call sort_dates(ht_dates_array)					'This function takes all the dates in an array and put them in order from oldest to newest

dlg_len = 85 + ((UBOUND(HOT_TOPIC_ARRAY, 2)+1) * 15)			'using math to determine the size of the dialog
If dlg_len > 390 Then dlg_len = 400

'creating the dialog
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 570, dlg_len, "BlueZone Script Hot Topics"
ButtonGroup ButtonPressed
	Text 10, 10, 415, 20, "These are the most recent Hot Topic Articles from the BlueZone Script Team. From this menu, you can access the article information, run the script, read the instructions for the scipt or add the script to your favorites menu."
	PushButton 400, 15, 160, 15, "Open the BZST Hot Topics SharePoint Page", bzst_hot_topics_page_btn
	Text 25, 45, 50, 10, "Date"
	Text 75, 45, 40, 10, "Article"
	Text 365, 45, 110, 10, "Script (press to run the script)"
	Text 510, 45, 15, 10, "Instr"
	Text 530, 45, 30, 10, "Fav"

	y_pos = 60
	For hot_topic_date = UBound(ht_dates_array) to 0 Step -1										'This loops through all of the dates in the array that was sorted, starting from the bottom (the newest)
		ht_dates_array(hot_topic_date) = DateAdd("d", 0, ht_dates_array(hot_topic_date))			'making sure this is a date
		Text 25, y_pos, 50, 10, ht_dates_array(hot_topic_date)										'display the date
		For article = 0 to UBound(HOT_TOPIC_ARRAY, 2)												'looping through the array of all of HT articles
			HOT_TOPIC_ARRAY(hot_topic_date_const, article) = DateAdd("d", 0, HOT_TOPIC_ARRAY(hot_topic_date_const, article))						'making sure this is a date
			'looking to see if the item in the HT array matches the current date being displayed - if they match, it will display the HT article information
			If DateDiff("d", ht_dates_array(hot_topic_date), HOT_TOPIC_ARRAY(hot_topic_date_const, article)) = 0 and HOT_TOPIC_ARRAY(script_displayed, article)  = False Then
				PushButton 75, y_pos-3, 280, 13, HOT_TOPIC_ARRAY(hot_topic_name_const, article), HOT_TOPIC_ARRAY(article_btn_const, article)		'The HT name is in a button
				If HOT_TOPIC_ARRAY(script_name_const, article) <> "" Then							'if there is a script listed with the article
					'display a button with the script name to run the script, plus a '?' button for instructions and a '+' button to add to favorites
					PushButton 365, y_pos-3, 140, 13, HOT_TOPIC_ARRAY(script_category_const, article) & " - " & HOT_TOPIC_ARRAY(script_name_const, article), HOT_TOPIC_ARRAY(run_script_btn, article)
					PushButton 510, y_pos-3, 15, 15, "?", HOT_TOPIC_ARRAY(instructions_btn_const, article)
					PushButton 530, y_pos-3, 15, 15, "+", HOT_TOPIC_ARRAY(add_to_favorites_btn, article)
				Else
					Text 365, y_pos, 140, 13, "No Specific Associated Script"						'if there is no associated script, displays text explaining this
				End If
				HOT_TOPIC_ARRAY(script_displayed, article)  = True		'setting that the information has already been displayed, so we don't duplicate
				y_pos = y_pos + 15										'move the location incrementor down for the next article

				'if it was found that this article has more than one entry in the array, it will show these additional scripts without repeating the HT article information
				If HOT_TOPIC_ARRAY(multiple_scripts_TF, article) = True Then
					For second_article = 0 to UBound(HOT_TOPIC_ARRAY, 2)		'looping through the array again
						'if the article url matches and the information was not already displayed, show it again
						If HOT_TOPIC_ARRAY(hot_topic_link_const, second_article) = HOT_TOPIC_ARRAY(hot_topic_link_const, article) and HOT_TOPIC_ARRAY(script_displayed, second_article)  = False Then
							'display a button with the script name to run the script, plus a '?' button for instructions and a '+' button to add to favorites
							PushButton 365, y_pos-3, 140, 13, HOT_TOPIC_ARRAY(script_category_const, second_article) & " - " & HOT_TOPIC_ARRAY(script_name_const, second_article), HOT_TOPIC_ARRAY(run_script_btn, second_article)
							PushButton 510, y_pos-3, 15, 15, "?", HOT_TOPIC_ARRAY(instructions_btn_const, second_article)
							PushButton 530, y_pos-3, 15, 15, "+", HOT_TOPIC_ARRAY(add_to_favorites_btn, second_article)
							HOT_TOPIC_ARRAY(script_displayed, second_article)  = True		'setting that the information has already been displayed, so we don't duplicate
							y_pos = y_pos + 15
						End if
					Next
				End If
			End If
		Next
		If y_pos >= 370 Then Exit For		'If we have too many HT articles, it will stop when we get to the end of what the dialog size limit is
	Next
	GroupBox 10, 35, 555, y_pos-35, "Hot Topics List"

	Text 10, dlg_len - 15, 140, 10, "Do you have another question or an idea?"
	PushButton 150, dlg_len-20, 95, 15, "Report to the BZST", report_to_BZST_btn
	OkButton 460, dlg_len-20, 50, 15
	CancelButton 515, dlg_len-20, 50, 15
EndDialog

'showing the
'This dialog has no password or error handling because it does not operate in MAXIS at all.
'This dialog essentially acts like a menu and either connects to websites or runs scripts.
Do
	dialog Dialog1
	cancel_without_confirmation

	'identifying the actions to take based on the button pressed.
	If ButtonPressed = bzst_hot_topics_page_btn Then Call open_URL_in_browser("https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/BlueZone-Scripts.aspx")
	If ButtonPressed = report_to_BZST_btn Then Call run_from_GitHub(script_repository & "utilities/report-to-the-bzst.vbs")

	'the rest of the buttons are associated with a button object in the array and we need to loop through that array to find the button
	For article = 0 to UBound(HOT_TOPIC_ARRAY, 2)
		If ButtonPressed = HOT_TOPIC_ARRAY(article_btn_const, article) Then Call open_URL_in_browser(HOT_TOPIC_ARRAY(hot_topic_link_const, article))		'open a webpage of the article
		If ButtonPressed = HOT_TOPIC_ARRAY(run_script_btn, article) Then Call run_from_GitHub(HOT_TOPIC_ARRAY(script_url_const, article))					'run a script from the dialog
		If ButtonPressed = HOT_TOPIC_ARRAY(instructions_btn_const, article) Then Call open_URL_in_browser(HOT_TOPIC_ARRAY(script_instructions_url_const, article))		'open the script instructions
		If ButtonPressed = HOT_TOPIC_ARRAY(add_to_favorites_btn, article) Then				'this will add the script to the list of favorites
			For i = 0 to ubound(script_array)
				If script_array(i).script_URL = HOT_TOPIC_ARRAY(script_url_const, article) Then							'finding the right script in the array from the CLOS
					If script_array(i).script_in_favorites = TRUE Then													'if this script is already IN favorites, we shouldn't add again
						MsgBox "The script " & script_array(i).category & "-" & script_array(i).script_name & " is already listed in favorites."
					Else
						new_favorite = script_array(i).category & "/" & script_array(i).script_name						'creating the string for the correct way to save a favorite script
						If all_favorites = "" Then																		'making a list of all the favorites, including the new one
							all_favorites = join(favorites_text_file_array, vbNewLine)
						End If
						all_favorites = all_favorites & vbNewLine & new_favorite

						SET updated_fav_scripts_fso = CreateObject("Scripting.FileSystemObject")						'creating the txt file that operates the favorites and saving it
						SET updated_fav_scripts_command = updated_fav_scripts_fso.CreateTextFile(favorites_text_file_location, 2)
						updated_fav_scripts_command.Write(all_favorites)
						updated_fav_scripts_command.Close

						'telling the user that the script was added to the favorites.
						MsgBox "The script " & script_array(i).category & "-" & script_array(i).script_name & " has been added to your list of favorites."
					End If
				End if
			Next
		End If
	Next
Loop until ButtonPressed = OK		'If the worker presses 'OK' the script will leave the dialog loop and stop running.

'Script ends
script_end_procedure("")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------05/04/2023
'--Tab orders reviewed & confirmed----------------------------------------------05/04/2023
'--Mandatory fields all present & Reviewed--------------------------------------N/A
'--All variables in dialog match mandatory fields-------------------------------N/A
'Review dialog names for content and content fit in dialog----------------------05/04/2023
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------N/A
'--MAXIS_background_check reviewed (if applicable)------------------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------N/A
'--Out-of-County handling reviewed----------------------------------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------05/04/2023
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------N/A						Left this as a minute, this script isn't really for time savings, but for information access
'--Incrementors reviewed (if necessary)-----------------------------------------N/A
'--Denomination reviewed -------------------------------------------------------N/A
'--Script name reviewed---------------------------------------------------------N/A
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------05/04/2023
'--comment Code-----------------------------------------------------------------N/A
'--Update Changelog for release/update------------------------------------------05/04/2023
'--Remove testing message boxes-------------------------------------------------05/04/2023
'--Remove testing code/unnecessary code-----------------------------------------05/04/2023
'--Review/update SharePoint instructions----------------------------------------05/04/2023
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------TODO
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------N/A
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------N/A
