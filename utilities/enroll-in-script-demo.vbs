'Required for statistical purposes==========================================================================================
name_of_script = "UTILITIES - Enroll in Script Demo.vbs"
start_time = timer
STATS_counter = 0                          'sets the stats counter at one
STATS_manualtime = 150                     'manual run time in seconds
STATS_denomination = "I"                   'C is for each CASE
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
call changelog_update("05/24/2019", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

function script_search(name_of_the_scripts)
	btn_counter = 5000
	For the_script = 0 to UBound(script_array, 1)
		script_array(the_script).script_btn_one = btn_counter
		btn_counter = btn_counter + 1

	Next
	Do
		dlg_len = 105
		If search_words <> "" Then search_words_array = split(search_words, " ")
		x_pos = 15
		y_pos = 65
		count = 1

		For the_script = 0 to UBound(script_array, 1)
			' script_array(the_script).script_checkbox_one = checked
			script_array(the_script).show_script = FALSE
			If search_category <> "" AND search_category <> "Select One..." AND search_category <> "SPECIALITY" Then
				If script_array(the_script).category = search_category OR (search_category = "NOTICES" AND script_array(the_script).category = "" ) Then
					If search_words <> "" Then
						For each the_word in search_words_array
							If InStr(script_array(the_script).script_name, the_word) <> 0 Then
								' MsgBox "1"
								script_array(the_script).show_script = TRUE
								count = count + 1
								If count = 5 Then
									dlg_len = dlg_len + 15
									count = 1
								End If
							End If
						Next
					Else
						' MsgBox "2"
						script_array(the_script).show_script = TRUE
						count = count + 1
						If count = 5 Then
							dlg_len = dlg_len + 15
							count = 1
						End If
					End If
				End If
			ElseIf search_category = "SPECIALITY" Then

			ElseIf search_words <> "" Then
				For each the_word in search_words_array
					If InStr(script_array(the_script).script_name, the_word) <> 0 Then
						' MsgBox "3"
						script_array(the_script).show_script = TRUE
						count = count + 1
						If count = 5 Then
							dlg_len = dlg_len + 15
							count = 1
						End If
					End If
				Next
			End If
			' script_array(the_script).script_checkbox_one = unchecked
		Next
		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 720, dlg_len, "Search for Scripts"
		  ButtonGroup ButtonPressed
			DropListBox 50, 25, 125, 45, "Select One..."+chr(9)+"NOTES"+chr(9)+"ADMIN"+chr(9)+"ACTIONS"+chr(9)+"BULK"+chr(9)+"NOTICES"+chr(9)+"UTILITIES"+chr(9)+"SPECIALTY", search_category
			EditBox 260, 25, 220, 15, search_words
			PushButton 490, 25, 75, 15, "SEARCH", search_go_btn
			Text 10, 10, 475, 10, "This dialog can display a number of scripts that we have available, so you can select the script(s) that apply to the information you are providing."
			Text 10, 30, 35, 10, "Category:"
			Text 190, 30, 70, 10, "Part of script name:"
			Text 10, 50, 160, 10, "Check the script(s) that you are reporting about:"
			For the_script = 0 to UBound(script_array, 1)
			  If script_array(the_script).show_script = TRUE Then
			  	  PushButton x_pos, y_pos, 20, 10, "ADD", script_array(the_script).script_btn_one
				  Text x_pos + 23, y_pos, 160, 10, script_array(the_script).category & " - " & script_array(the_script).script_name
				  x_pos = x_pos + 170
				  If x_pos = 695 Then
					  y_pos = y_pos + 15
					  x_pos = 15
				  End If
			  End If
			Next
			Text 15, y_pos + 15, 540, 10, name_of_the_scripts
				' OkButton 555, y_pos + 5, 50, 15
			PushButton 665, y_pos + 20, 50, 15, "DONE", done_btn
		EndDialog

		dialog Dialog1

		For the_script = 0 to UBound(script_array, 1)
			If ButtonPressed = script_array(the_script).script_btn_one Then name_of_the_scripts = name_of_the_scripts & ", " & script_array(the_script).category & " - " & script_array(the_script).script_name
		Next

	Loop until ButtonPressed = done_btn

	If left(name_of_the_scripts, 2) = ", " Then name_of_the_scripts = right(name_of_the_scripts, len(name_of_the_scripts)-2)
	' MsgBox name_of_the_scripts

	ButtonPressed = search_btn

end function
' class script_bowie
'
'     'Stuff the user indicates
' 	public script_name             	'The familiar name of the script (file name without file extension or category, and using familiar case)
' 	public description             	'The description of the script
' 	public button                  	'A variable to store the actual results of ButtonPressed (used by much of the script functionality)
' 	public SIR_instructions_button	'A variable to store the actual results of ButtonPressed (used by much of the script functionality)
'     public category               	'The script category (ACTIONS/BULK/etc)
' 	public workflows               	'The script workflows associated with this script (Changes Reported, Applications, etc)
'     public subcategory				'An array of all subcategories a script might exist in, such as "LTC" or "A-F"
' 	public release_date				'This allows the user to indicate when the script goes live (controls NEW!!! messaging)
'
'     'Details the menus will figure out (does not need to be explicitly declared)
'     public button_plus_increment	'Workflow scripts use a special increment for buttons (adding or subtracting from total times to run). This is the add button.
' 	public button_minus_increment	'Workflow scripts use a special increment for buttons (adding or subtracting from total times to run). This is the minus button.
' 	public total_times_to_run		'A variable for the total times the script should run
'
'     'Details the class itself figures out
' 	public property get script_URL
' 		If run_locally = true then
' 			script_repository = "C:\MAXIS-Scripts\"
' 			script_URL = script_repository & lcase(category) & "\" & lcase(replace(script_name, " ", "-") & ".vbs")
' 		Else
'         	If script_repository = "" then script_repository = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/"    'Assumes we're scriptwriters
'         	script_URL = script_repository & lcase(category) & "/" & replace(lcase(script_name) & ".vbs", " ", "-")
' 		End if
'     end property
'
'     'public property get SIR_instructions_URL 'The instructions URL in SIR
'     '    SIR_instructions_URL = "https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/Script%20Instructions%20Wiki/" & replace(ucase(script_name) & ".aspx", " ", "%20")
'     'end property
'
' end class
'
class script_demo

    public script_name
    public category
    public tags
    public instructions
    public demo_dates
    public demo_length
    public future_dates
	public demo_url
    public group_len

end class

'THESE FORST TWO ARE OLD FAKE ONES TO USE FOR EXPERIENCE SETTING'
script_num = 0
ReDim Preserve SCRIPT_DEMO_ARRAY(script_num)
Set SCRIPT_DEMO_ARRAY(script_num) = new script_demo
SCRIPT_DEMO_ARRAY(script_num).script_name   = "Client Contact"
SCRIPT_DEMO_ARRAY(script_num).category      = "NOTES"
SCRIPT_DEMO_ARRAY(script_num).tags          = ""
SCRIPT_DEMO_ARRAY(script_num).instructions  = "https://dept.hennepin.us/hsphd/sa/ews/BlueZone_Script_Instructions/NOTES/NOTES%20-%20CLIENT%20CONTACT.docx"
SCRIPT_DEMO_ARRAY(script_num).demo_dates    = ARRAY(#7/3/2020 3:00 PM#, #7/8/2020 9:00 AM#, #7/10/2020 10:00 AM#, #7/11/2020 2:00 PM#)
SCRIPT_DEMO_ARRAY(script_num).demo_length    = 45
SCRIPT_DEMO_ARRAY(script_num).future_dates  = FALSE
SCRIPT_DEMO_ARRAY(script_num).group_len     = 0
SCRIPT_DEMO_ARRAY(script_num).demo_url 		= ARRAY("www.google.com/one", "www.google.com/two", "www.google.com/three", "www.google.com/four")

script_num = script_num + 1
ReDim Preserve SCRIPT_DEMO_ARRAY(script_num)
Set SCRIPT_DEMO_ARRAY(script_num) = new script_demo
SCRIPT_DEMO_ARRAY(script_num).script_name   = "Application Received"
SCRIPT_DEMO_ARRAY(script_num).category      = "NOTES"
SCRIPT_DEMO_ARRAY(script_num).tags          = ""
SCRIPT_DEMO_ARRAY(script_num).instructions  = "https://dept.hennepin.us/hsphd/sa/ews/BlueZone_Script_Instructions/NOTES/NOTES%20-%20APPLICATION%20RECEIVED.docx"
SCRIPT_DEMO_ARRAY(script_num).demo_dates    = ARRAY(#7/10/2020 3:00 PM#, #7/15/2020 9:00 AM#, #7/17/2020 10:00 AM#, #7/18/2020 2:00 PM#)
SCRIPT_DEMO_ARRAY(script_num).demo_length    = 45
SCRIPT_DEMO_ARRAY(script_num).future_dates  = FALSE
SCRIPT_DEMO_ARRAY(script_num).group_len     = 0
SCRIPT_DEMO_ARRAY(script_num).demo_url 		= ARRAY("www.google.com/one", "www.google.com/two", "www.google.com/three", "www.google.com/four")

' 'TEMPLATE FOR ADDING MORE DEMOs
' script_num = script_num + 1
' ReDim Preserve SCRIPT_DEMO_ARRAY(script_num)
' Set SCRIPT_DEMO_ARRAY(script_num) = new script_demo
' SCRIPT_DEMO_ARRAY(script_num).script_name   = "SCRIPT_NAME_HERE"
' SCRIPT_DEMO_ARRAY(script_num).category      = "SCRIPT_CATEGORY"
' SCRIPT_DEMO_ARRAY(script_num).tags          = ""
' SCRIPT_DEMO_ARRAY(script_num).instructions  = "instructions_url_here"
' ' SCRIPT_DEMO_ARRAY(script_num).demo_dates    = ARRAY(#9/3/2020 3:00 PM#, #9/8/2020 9:00 AM#, #9/10/2020 10:00 AM#, #9/11/2020 2:00 PM#)
' SCRIPT_DEMO_ARRAY(script_num).demo_dates    = ARRAY(#mm/dd/yyyy h:mm am/pm#, #9/11/2020 2:00 PM#)
' SCRIPT_DEMO_ARRAY(script_num).demo_length    = time_in_minutes
' SCRIPT_DEMO_ARRAY(script_num).future_dates  = FALSE		'start as FALSE every time
' SCRIPT_DEMO_ARRAY(script_num).group_len     = 0			'start at 0 every time
' SCRIPT_DEMO_ARRAY(script_num).demo_url 		= ARRAY("team_meeting_demo_url", "team_meeting_url_next")		'these array items need to line up with the times.

' 'THESE ARE FOR TESTING IF YOU WANT TO TRY IT OUT'
' script_num = script_num + 1
' ReDim Preserve SCRIPT_DEMO_ARRAY(script_num)
' Set SCRIPT_DEMO_ARRAY(script_num) = new script_demo
' SCRIPT_DEMO_ARRAY(script_num).script_name   = "Counted ABAWD Months"
' SCRIPT_DEMO_ARRAY(script_num).category      = "ACTIONS"
' SCRIPT_DEMO_ARRAY(script_num).tags          = ""
' SCRIPT_DEMO_ARRAY(script_num).instructions  = "https://dept.hennepin.us/hsphd/sa/ews/BlueZone_Script_Instructions/ACTIONS/ACTIONS%20-%20COUNTED%20ABAWD%20MONTHS.docx"
' ' SCRIPT_DEMO_ARRAY(script_num).demo_dates    = ARRAY(#9/3/2020 3:00 PM#, #9/8/2020 9:00 AM#, #9/10/2020 10:00 AM#, #9/11/2020 2:00 PM#)
' SCRIPT_DEMO_ARRAY(script_num).demo_dates    = ARRAY(#9/3/2020 3:00 PM#, #9/11/2020 2:00 PM#)
' SCRIPT_DEMO_ARRAY(script_num).demo_length    = 45
' SCRIPT_DEMO_ARRAY(script_num).future_dates  = FALSE
' SCRIPT_DEMO_ARRAY(script_num).group_len     = 0
' SCRIPT_DEMO_ARRAY(script_num).demo_url 		= ARRAY("www.google.com/one", "www.google.com/two")
'
' script_num = script_num + 1
' ReDim Preserve SCRIPT_DEMO_ARRAY(script_num)
' Set SCRIPT_DEMO_ARRAY(script_num) = new script_demo
' SCRIPT_DEMO_ARRAY(script_num).script_name   = "Earned Income Budgeting"
' SCRIPT_DEMO_ARRAY(script_num).category      = "ACTIONS"
' SCRIPT_DEMO_ARRAY(script_num).tags          = ""
' SCRIPT_DEMO_ARRAY(script_num).instructions  = "https://dept.hennepin.us/hsphd/sa/ews/BlueZone_Script_Instructions/ACTIONS/ACTIONS%20-%20EARNED%20INCOME%20BUDGETING.docx"
' SCRIPT_DEMO_ARRAY(script_num).demo_dates    = ARRAY(#9/14/2020 3:00 PM#, #9/15/2020 8:30 AM#, #9/16/2020 10:00 AM#, #9/17/2020 2:00 PM#)
' SCRIPT_DEMO_ARRAY(script_num).demo_length    = 45
' SCRIPT_DEMO_ARRAY(script_num).future_dates  = FALSE
' SCRIPT_DEMO_ARRAY(script_num).group_len     = 0
' SCRIPT_DEMO_ARRAY(script_num).demo_url 		= ARRAY("www.google.com/one", "www.google.com/two", "www.google.com/three", "www.google.com/four")
'
' script_num = script_num + 1
' ReDim Preserve SCRIPT_DEMO_ARRAY(script_num)
' Set SCRIPT_DEMO_ARRAY(script_num) = new script_demo
' SCRIPT_DEMO_ARRAY(script_num).script_name   = "CAF Script"
' SCRIPT_DEMO_ARRAY(script_num).category      = "NOTES"
' SCRIPT_DEMO_ARRAY(script_num).tags          = ""
' SCRIPT_DEMO_ARRAY(script_num).instructions  = "https://dept.hennepin.us/hsphd/sa/ews/BlueZone_Script_Instructions/NOTES/NOTES%20-%20CAF.docx"
' SCRIPT_DEMO_ARRAY(script_num).demo_dates    = ARRAY(#10/5/2020 3:00 PM#, #10/6/2020 8:30 AM#, #10/7/2020 10:00 AM#)
' SCRIPT_DEMO_ARRAY(script_num).demo_length    = 45
' SCRIPT_DEMO_ARRAY(script_num).future_dates  = FALSE
' SCRIPT_DEMO_ARRAY(script_num).group_len     = 0
' SCRIPT_DEMO_ARRAY(script_num).demo_url 		= ARRAY("www.google.com/one", "www.google.com/two", "www.google.com/three")

'Find who is running - might use this in the future for tester specific demos or other things.
Set objNet = CreateObject("WScript.NetWork")                                    'getting the users windows ID
windows_user_ID = objNet.UserName
user_ID_for_validation = ucase(windows_user_ID)

For each tester in tester_array                                                 'Loop through all the testers in the array to see if the user is in the list of testers.
    If user_ID_for_validation = tester.tester_id_number Then
        worker_full_name            = tester.tester_full_name
        worker_first_name           = tester.tester_first_name
        worker_last_name            = tester.tester_last_name
        worker_email                = tester.tester_email
        worker_id_number            = tester.tester_id_number
        worker_x_number             = tester.tester_x_number
        worker_supervisor           = tester.tester_supervisor_name
        worker_supervisor_email     = tester.tester_supervisor_email
        worker_test_groups          = tester.tester_groups
        qi_staff = FALSE
        For each group in worker_test_groups                                 'looking at all of the groups this tester is a part of to see if QI or BZ
            If group = "QI" Then qi_staff = TRUE
            If group = "BZ" Then qi_staff = TRUE
        Next
    End If
Next
'If this did not find the user is a tester for QI the script will end as this is only for QI staff - access to the files and folders will be restricted and the script will fail
' If qi_staff = FALSE Then scsript_end_procedure_with_error_report("This script is for QI specific processes and only for QI staff. You are not listed as QI staff and running this script could cause errors in data reccording and QI processes. Please contact the BlueZone script team or pres 'Yes' below if you believe this to be in error.")
bzt_email = "HSPH.EWS.BlueZoneScripts@hennepin.us"
worker_name = worker_full_name
EMConnect ""

unique_scripts = 0
total_dates = 0

Dim CHECKBOX_ARRAY()
ReDim CHECKBOX_ARRAY(0)

For each scheduled_script in SCRIPT_DEMO_ARRAY
    no_future_dates = TRUE
    For each scheduled_date in scheduled_script.demo_dates
        ReDim Preserve CHECKBOX_ARRAY(checkbox_counter)
        'MsgBox "Scheduled date: " & scheduled_date & vbNewLine & "Diff: " & DateDiff("n", now, scheduled_date)
        If DateDiff("n", now, scheduled_date) > -1 Then
            no_future_dates = FALSE
            total_dates = total_dates + 1
            scheduled_script.future_dates  = TRUE
            scheduled_script.group_len = scheduled_script.group_len + 15
        End If
        checkbox_counter = checkbox_counter + 1
    Next
    If no-no_future_dates = FALSE Then unique_scripts = unique_scripts + 1
Next


Do
	Do
	    err_msg = ""
		' If dlg_len = 160 Then dlg_len = 185
		dlg_len = 205 + (unique_scripts * 25) + (total_dates * 15)
		y_pos = 155
		If dlg_len = 205 Then dlg_len = 185

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 391, dlg_len, "Select DEMOs to Enroll"
		  Text 115, 10, 145, 10, "Welcome to the BlueZone Script Roadshow!"
		  GroupBox 5, 25, 380, 90, "About Script Demos"
		  Text 15, 40, 350, 15, "As our project is constantly growing and changing, we want to show you how best to use the tools we create."
		  Text 15, 60, 360, 25, "Since we serve all populations and regions, and because our presence is mostly virtual, our trainings will be the same. The focus of our demos and information is on the tool, how it acts, and how you can use it. You don't need to see our faces, just MAXIS and the scripts. "
		  Text 15, 90, 360, 20, "Due to all of these reasons, our Demos and Trainings are scheduled as remote Skype meetings. These meetings can be found on our SharePoint site, all you have to do is click on them to join!"

		  checkbox_counter = 0
		  For each scheduled_script in SCRIPT_DEMO_ARRAY
		      If scheduled_script.future_dates = TRUE Then
		        GroupBox 10, y_pos, 375, 20 + scheduled_script.group_len, scheduled_script.category & " - " & scheduled_script.script_name
		        y_pos = y_pos + 15
		      End If
		      'For each scheduled_date in scheduled_script.demo_dates
		      For array_counter = 0 to UBound(scheduled_script.demo_dates)
		          scheduled_date = scheduled_script.demo_dates(array_counter)
		          If DateDiff("n", now, scheduled_date) > -1 Then

		              CheckBox 25, y_pos, 345, 10, FormatDateTime(scheduled_date, 1) & " at " & FormatDateTime(scheduled_date, 3) & " - " & scheduled_script.script_name & "(" & scheduled_script.demo_length & " minutes)", CHECKBOX_ARRAY(checkbox_counter)
		              y_pos = y_pos + 15

		          End If
		          checkbox_counter = checkbox_counter + 1
		      Next
		      If scheduled_script.future_dates = TRUE Then y_pos = y_pos + 10
		  Next
		  'y_pos = y_pos + 10


		  ' GroupBox 10, 175, 375, 60, "ACTIONS - Earned Income Budgeting"
		  ' CheckBox 25, 190, 345, 10, "Wednesday July 10th at 3:00 PM - Earned Income Budgeting (45 minutes)", checkBoxOne
		  ' CheckBox 25, 205, 345, 10, "Tuesday July 16th at 8:30 AM - Earned Income Budgeting (45 minutes)", Check2
		  ' CheckBox 25, 220, 345, 10, "Thursday July 25th at 10:00 AM - Earned Income Budgeting (45 minutes)", Check3
		  ' GroupBox 5, 245, 375, 60, "NOTES - CAF"
		  ' CheckBox 25, 260, 345, 10, "Wednesday August 7th at 3:00 PM - CAF Script (45 minutes)", Check4
		  ' CheckBox 25, 275, 345, 10, "Tuesday August 13th at 8:30 AM - CAF Script (45 minutes)", Check5
		  ' CheckBox 25, 290, 345, 10, "Thursday August 22nd at 10:00 AM - CAF Script (45 minutes)", Check6
		  If y_pos <> 155 Then
			  Text 140, 120, 85, 10, "Upcoming Script Demos"
			  ' Text 25, 130, 400, 20, "Check the box by any session to enroll in that Demo. This will schedule it in your Outlook and give us a notice that you will be joining."
			  Text 10, 130, 400, 10, "Check the box by any session to enroll in that Demo, adding it to Outlook and sending the enrollment to the BZST."

			  Text 175, 145, 100, 10, "Enter your Name for Enrollment:"
			  EditBox 285, 140, 100, 15, worker_name
			  ' y_pos = y_pos + 20
		  Else
			  Text 10, 120, 200, 10, "*** We do not have any demos scheduled at this time.***"
			  y_pos = 135
		  End If
		  Text 10, y_pos, 300, 10, "If you have any ideas of scripts you would like to see in a demo, please enter them here."
		  EditBox 10, y_pos+10, 330, 15, script_demo_ideas
		  ButtonGroup ButtonPressed
		    PushButton 345, y_pos+12, 40, 13, "SEARCH", search_btn
		    OkButton 280, y_pos+30, 50, 15
		    CancelButton 335, y_pos+30, 50, 15
		EndDialog


	    dialog Dialog1
	    cancel_without_confirmation

		If ButtonPressed = search_btn Then call script_search(script_demo_ideas)
	Loop until ButtonPressed = -1

    If err_msg <> "" Then MsgBox "Please Resolve to Continue:" & vbNewLine & err_msg

Loop until err_msg = ""

script_demo_ideas = trim(script_demo_ideas)
worker_name = trim(worker_name)

end_msg = "Success!"
If worker_name = "" Then worker_name = "THIS WORKER"
checkbox_counter = 0
For each scheduled_script in SCRIPT_DEMO_ARRAY
    For array_counter = 0 to UBound(scheduled_script.demo_dates)
        'MsgBox array_counter & vbNewLine & CHECKBOX_ARRAY(checkbox_counter)
        If CHECKBOX_ARRAY(checkbox_counter) = checked Then
        ' If scheduled_script.demo_checkbox(array_counter) = checked Then
            'MsgBox "EMAIL TO SEND" & vbNewLine & scheduled_script.category & " - " & scheduled_script.script_name & vbNewLine & scheduled_script.demo_dates(array_counter)

            'create_outlook_appointment(appt_date, appt_start_time, appt_end_time, appt_subject, appt_body, appt_location, appt_reminder, reminder_in_minutes, appt_category)

            ' body_text = "Join the BlueZone Script team remotely to see a script demo on " & scheduled_script.category & " - " & scheduled_script.script_name & vbCr & vbCr & "Go here and double click on the link to join the TEAMS meeting - " & scheduled_script.demo_url & vbCr & vbCr & "Instructions for this script can be found here - " & scheduled_script.instructions

            confirm_demo_schedule = MsgBox("You have selected to join the script demo for:" & vbCr & vbCr & scheduled_script.category & " - " & scheduled_script.script_name & vbCr & vbCr & "This demo will be held on " & WeekDayName(WeekDay(scheduled_script.demo_dates(array_counter))) & " " & FormatDateTime(scheduled_script.demo_dates(array_counter), 2) & " at " & FormatDateTime(scheduled_script.demo_dates(array_counter), 3) & vbCr & vbCr & "Do you wish to enroll in this Demo and schedule it?", vbQuestion + vbYesNo, "Confirm Enrollment in DEMO")

            If confirm_demo_schedule = vbYes Then
				STATS_counter = STATS_counter + 1
				If end_msg = "Success!" Then end_msg = "Success! The following demo(s) have been added to your calendar and your enrollment has been sent to the BlueZone Script team:" & vbNewLine
                ' Call create_outlook_appointment(FormatDateTime(scheduled_script.demo_dates(array_counter), 2), FormatDateTime(scheduled_script.demo_dates(array_counter), 3), FormatDateTime(DateAdd("n", scheduled_script.demo_length, scheduled_script.demo_dates(array_counter)), 3), "Script Demo - " & scheduled_script.script_name, body_text, "Microsoft Teams", TRUE, 60, "")

				'TESTING SOME THINGS OUT TO TRY TO CHANGE THE HYPERLINK TEXT'
				' Dim objInsp As Object, objDoc As Object, objSel As Object, strLinkone As String, strLinkTextone As String, strLinktwo As String, strLinkTexttwo As String

				strLinkone = scheduled_script.demo_url(array_counter)
				strLinkTextone = "Link to TEAMS Meeting"

				strLinktwo = scheduled_script.instructions
				strLinkTexttwo = "SCRIPT Instructions"



				'Assigning needed numbers as variables for readability
				olAppointmentItem = 1
				olRecursDaily = 0

				'Creating an Outlook object item
				Set objOutlook = CreateObject("Outlook.Application")
				Set objAppointment = objOutlook.CreateItem(olAppointmentItem)

				Set objInsp = objAppointment.GetInspector
				Set objDoc = objInsp.WordEditor
				Set objSel = objDoc.Windows(1).Selection


				'Assigning individual appointment options
				objAppointment.Start = FormatDateTime(scheduled_script.demo_dates(array_counter), 2) & " " & FormatDateTime(scheduled_script.demo_dates(array_counter), 3)		'Start date and time are carried over from parameters
				objAppointment.End = FormatDateTime(scheduled_script.demo_dates(array_counter), 2) & " " & FormatDateTime(DateAdd("n", scheduled_script.demo_length, scheduled_script.demo_dates(array_counter)), 3)			'End date and time are carried over from parameters
				objAppointment.AllDayEvent = False 								'Defaulting to false for this. Perhaps someday this can be true. Who knows.
				objAppointment.Subject = "Script Demo - " & scheduled_script.script_name							'Defining the subject from parameters
				' objAppointment.Body = body_text									'Defining the body from parameters

				' objAppointment.Body = "Join the BlueZone Script team remotely to see a script demo on " & scheduled_script.category & " - " & scheduled_script.script_name & vbCr & vbCr
				' objAppointment.Body = "Use this link to access the demo - "
				objAppointment.Body = vbCr & "Use this link above access the demo on " &  WeekDayName(WeekDay(scheduled_script.demo_dates(array_counter))) & " " & FormatDateTime(scheduled_script.demo_dates(array_counter), 2) & " at " & FormatDateTime(DateAdd("n", -5, scheduled_script.demo_dates(array_counter)), 3) & vbCr & vbCr & "Join the BlueZone Script team remotely to see a script demo on " & scheduled_script.category & " - " & scheduled_script.script_name & vbCr & vbCr & "Instructions for this script can be found here - " & scheduled_script.instructions & "."
				objDoc.Hyperlinks.Add objSel.Range, strLinkone, "", "", strLinkTextone, ""
				' objDoc.Hyperlinks.Add objSel.Range, strLinktwo, _
				' "", "", strLinkTexttwo, ""


				' body_text = "Join the BlueZone Script team remotely to see a script demo on " & scheduled_script.category & " - " & scheduled_script.script_name & vbCr & vbCr & "Go here and double click on the link to join the TEAMS meeting - " & "<a href " & chr(34) & scheduled_script.demo_url & chr(34) &">Join TEAMS Meeting</a>" & vbCr & vbCr & "Instructions for this script can be found here - " & scheduled_script.instructions

				objAppointment.Location = "Microsoft Teams"							'Defining the location from parameters
				' If appt_reminder = FALSE then									'If the reminder parameter is false, it skips the reminder, otherwise it sets it to match the number here.
				' 	objAppointment.ReminderSet = False
				' Else
				objAppointment.ReminderSet = True
				objAppointment.ReminderMinutesBeforeStart = 60
				' End if
				objAppointment.Categories = ""						'Defines a category
				objAppointment.Save


                email_text = worker_name & " has enrolled in a DEMO for " & scheduled_script.category & " - " & scheduled_script.script_name & vbCr & "On: " & scheduled_script.demo_dates(array_counter)
                Call create_outlook_email(bzt_email, "", "DEMO Enrollment", email_text, "", TRUE)
                end_msg = end_msg & vbNewLine & "* " & scheduled_script.category & " - " & scheduled_script.script_name & vbNewLine & "  On: " & scheduled_script.demo_dates(array_counter) & vbNewLine
            End If

        End If
        checkbox_counter = checkbox_counter + 1
    Next
Next

If script_demo_ideas <> "" Then
	STATS_counter = STATS_counter + 1
	email_text = "~ THIS IS AN AUTOMATED EMAIL FROM THE 'ENROLL IN SCRIPT DEMO' SCRIPT ~" & vbCr & vbCR
	email_text = email_text & "Idea for additional Script Demo(s):" & vbCr
	email_text = email_text & script_demo_ideas & vbCr & vbCr
	email_text = email_text & "Thank you." & vbCr & worker_name

	Call create_outlook_email(bzt_email, "", "Ideas for BZ Roadshow DEMOs", email_text, "", TRUE)

	end_msg = end_msg & vbNewLine & "Additional information sent to the BZST:" & vbNewLine & vbNewLine & "* Send ideas for more demos: " & script_demo_ideas
End If

script_end_procedure(end_msg)
