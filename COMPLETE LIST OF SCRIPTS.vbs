''LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
'IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
'	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
'		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
'			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
'		Else											'Everyone else should use the release branch.
'			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
'		End if
'		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
'		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
'		req.send													'Sends request
'		IF req.Status = 200 THEN									'200 means great success
'			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
'			Execute req.responseText								'Executes the script code
'		ELSE														'Error message
'			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
'                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
'                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
'                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
'            StopScript
'		END IF
'	ELSE
'		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
'		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
'		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
'		text_from_the_other_script = fso_command.ReadAll
'		fso_command.Close
'		Execute text_from_the_other_script
'	END IF
'END IF
''END FUNCTIONS LIBRARY BLOCK================================================================================================
'I AM A COMMENT TO MAKE A CHANGE'
class script_bowie

    'Stuff the user indicates
	public script_name             	'The familiar name of the script (file name without file extension or category, and using familiar case)
	' public description             	'The description of the script
	public button                  	'A variable to store the actual results of ButtonPressed (used by much of the script functionality)
	public SIR_instructions_button	'A variable to store the actual results of ButtonPressed (used by much of the script functionality)
    public fav_add_button
    public category               	'The script category (ACTIONS/BULK/etc)
	public workflows               	'The script workflows associated with this script (Changes Reported, Applications, etc)
    public tags                     'The tags
    public dlg_keys                 'codes'
    public subcategory				'An array of all subcategories a script might exist in, such as "LTC" or "A-F"
	public release_date				'This allows the user to indicate when the script goes live (controls NEW!!! messaging)
    public show_script              'This property is defined during the script run - it determines if the script meets the criteria for the selected tab
    public keywords                 'Future enhancemnt that will allow us to search for scripts by keyword
    public hot_topic_date           'If a script is in HOT TOPICS, adding a date here will be used to feature in favorites and resources
	public hot_topic_link			'If a script is in HOT TOPICS, here we can add a url to link directly to the article.'
    public retirement_date          'Adding a date here indicates the script should not be shown because it has been retired. We must leave it in the list for favorites
    public in_testing               'This can be set to TRUE if we have a new script_bowie	that is in testing
    public testing_category         'idetify what we are using to determine WHO is testing - use ONLY ALL, GROUP, REGION, POPULATION, or SCRIPT
    public testing_criteria         'ARRAY list which of the category is being used
	public script_checkbox_one
	public script_btn_one
	public used_for_elig
    ' public stats_denomination_type
    ' public stats_manual_time_listed
    ' public stats_increments
    ' public region_available
    ' public last_update_date

    'Details the menus will figure out (does not need to be explicitly declared)
    public button_plus_increment	'Workflow scripts use a special increment for buttons (adding or subtracting from total times to run). This is the add button.
	public button_minus_increment	'Workflow scripts use a special increment for buttons (adding or subtracting from total times to run). This is the minus button.
	public total_times_to_run		'A variable for the total times the script should run

    'Details the class itself figures out
	public property get script_URL
        header_exists = FALSE
        header = ""
        If InStr(script_name, " - ") <> 0 Then
            header_exists = TRUE
            header_len = InStr(script_name, " - ") - 1
            header = left(script_name, header_len)
            name_to_use = right(script_name, len(script_name) - (header_len+3))
        End If
		If run_locally = true then
			script_repository = "C:\MAXIS-Scripts\"
            If header_exists = TRUE Then
                script_URL = script_repository & lcase(category) & "\" & lcase(header) & "-" & lcase(replace(name_to_use, " ", "-") & ".vbs")
            Else
			    script_URL = script_repository & lcase(category) & "\" & lcase(replace(script_name, " ", "-") & ".vbs")
            End If
		Else
        	If script_repository = "" then script_repository = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/"    'Assumes we're scriptwriters
            If header_exists = TRUE Then
                script_URL = script_repository & lcase(category) & "/" & lcase(header) & "-" & replace(lcase(name_to_use) & ".vbs", " ", "-")
            Else
                script_URL = script_repository & lcase(category) & "/" & replace(lcase(script_name) & ".vbs", " ", "-")
            End If
		End if
    end property

    public property get SharePoint_instructions_URL 'The instructions URL in SIR
        ' SharePoint_instructions_URL = "https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/Script%20Instructions%20Wiki/" & replace(ucase(script_name) & ".aspx", " ", "%20")
		' "https://dept.hennepin.us/hsphd/sa/ews/BlueZone_Script_Instructions/"			'OLD URL
        ' "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/"		'NEW URL'
		If category = "NOTES" AND left(script_name, 5) = "LTC -" Then
            SharePoint_instructions_URL = "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/" & UCase(category) & "/LTC%20scripts/" & UCase(category) & "%20-%20" & replace(replace(ucase(script_name) & ".docx", " - ", "%20"), " ", "%20")
		ElseIf script_name = "Add WCOM" Then
			SharePoint_instructions_URL = "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTICES/NOTICES%20-%20ADD%20WCOM.docx"
		ElseIf left(script_name, 4) = "REPT" OR script_name = "DAIL Report" OR script_name = "EMPS" OR script_name = "LTC-GRH List Generator" Then
			SharePoint_instructions_URL = "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/BULK/BULK%20-%20REPT%20LISTS.docx"
		Else
            SharePoint_instructions_URL = "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/" & UCase(category) & "/" & UCase(category) & "%20-%20" & replace(ucase(script_name) & ".docx", " ", "%20")
        End If
    end property

    public property get script_in_favorites
        ' MsgBox favorites_exist
        if favorites_exist = FALSE Then
            script_in_favorites = FALSE
        else
            For Each favorite_script in favorites_text_file_array
                fav_cat = ""
                fav_call = ""
                favorite_script = trim(favorite_script)
                category_end = InStr(favorite_script, "/")
                If category_end <> 0 Then
                    fav_cat = left(favorite_script, (category_end - 1))
                    fav_call = right(favorite_script, (len(favorite_script) - category_end))
                End If
                If fav_cat = category and fav_call = script_name Then script_in_favorites = TRUE
            Next
            If script_in_favorites = "" Then script_in_favorites = FALSE
        end if
    end Property

	' public sub create_decription(description)
	public property get description

		For each key in dlg_keys
			If key = "Cn" Then maxis_actions = maxis_actions & "Creates CASE:NOTE, "
			If key = "Fi" Then maxis_actions = maxis_actions & "FIATs Eligibility, "
			If key = "Tk" Then maxis_actions = maxis_actions & "Creates TIKL, "
			If key = "Up" Then maxis_actions = maxis_actions & "Updates Panels, "
			If key = "Ev" Then maxis_actions = maxis_actions & "Evaluates, "
			If key = "Sm" Then maxis_notice = maxis_notice & "SPEC:MEMO, "
			If key = "Sw" Then maxis_notice = maxis_notice & "SPEC:WCOM, "

			If key = "Ex" Then office_actions = office_actions & "Uses Excel, "
			If key = "Oa" Then office_actions = office_actions & "Creates Outlook Appointment, "
			If key = "Oe" Then office_actions = office_actions & "Generates Outlook Email, "
			If key = "Wrd" Then office_actions = office_actions & "Uses Word, "
		Next

		If maxis_actions <> "" Then description = "Actions: "
		If right(maxis_actions, 2) = ", " Then maxis_actions = left(maxis_actions, len(maxis_actions)-2)
		If right(maxis_notice, 2) = ", " Then maxis_notice = left(maxis_notice, len(maxis_notice)-2)
		If right(office_actions, 2) = ", " Then office_actions = left(office_actions, len(office_actions)-2)

		If maxis_actions <> "" Then
			' description = description & "In MAXIS: " & maxis_actions
			description = description & maxis_actions
			If maxis_notice <> "" Then description = description & ", Generates " & maxis_notice
		Else
			If maxis_notice <> "" Then description = description & "Generates " & maxis_notice
		End If
		If office_actions <> "" Then description = description & " " & office_actions

		all_programs = FALSE
		all_Cash = FALSE
		family_Cash = FALSE

		all_the_tags = join(tags, " ")
		If InStr(all_the_tags, "SNAP") <> 0 AND InStr(all_the_tags, "MFIP") <> 0 AND InStr(all_the_tags, "DWP") <> 0 AND InStr(all_the_tags, "Health Care") <> 0 AND InStr(all_the_tags, "HS/GRH") <> 0 AND InStr(all_the_tags, "Adult Cash") <> 0 AND InStr(all_the_tags, "EMER") <> 0 Then
		   	all_programs = TRUE
		ELSE
			If InStr(all_the_tags, "MFIP") <> 0 AND InStr(all_the_tags, "DWP") <> 0 AND InStr(all_the_tags, "Adult Cash") <> 0 Then
				all_Cash = TRUE
			ElseIf InStr(all_the_tags, "MFIP") <> 0  AND InStr(all_the_tags, "DWP") <> 0  Then
			   	family_Cash = TRUE
			End If
		End If
		programs_helped = ""

		If all_programs = TRUE Then
			programs_helped = "All Programs"
		ElseIf all_Cash = TRUE Then
			programs_helped = "All Cash"
		ElseIf family_Cash = TRUE Then
			programs_helped = "Family Cash"
		End IF
		For each tag in tags
			If all_programs = TRUE Then
			ElseIf all_Cash = TRUE OR family_Cash = TRUE Then
				If tag = "SNAP" Then programs_helped = programs_helped & "/SNAP"
				If tag = "Health Care" Then programs_helped = programs_helped & "/Health Care"
				If tag = "HS/GRH" Then programs_helped = programs_helped & "/GRH"
				If tag = "EMER" Then programs_helped = programs_helped & "/EMER"
				If tag = "LTC" AND InStr(all_the_tags, "Health Care") = 0 Then programs_helped = programs_helped & "/Health Care"
			Else
				If tag = "SNAP" Then programs_helped = programs_helped & "/SNAP"
				If tag = "MFIP" Then programs_helped = programs_helped & "/MFIP"
				If tag = "DWP" Then programs_helped = programs_helped & "/DWP"
				If tag = "Health Care" Then programs_helped = programs_helped & "/Health Care"
				If tag = "HS/GRH" Then programs_helped = programs_helped & "/GRH"
				If tag = "Adult Cash" Then programs_helped = programs_helped & "/Adult Cash"
				If tag = "EMER" Then programs_helped = programs_helped & "/EMER"
				If tag = "LTC" AND InStr(all_the_tags, "Health Care") = 0 Then programs_helped = programs_helped & "/Health Care"
			End If
		Next
		If left(programs_helped, 1) = "/" Then programs_helped = right(programs_helped, len(programs_helped)-1)
		If programs_helped <> "" Then description = description & " --- Programs Supported: " & programs_helped
		' If programs_helped <> "" Then description = description & " --- Programs Supported: " & programs_helped & vbCr & vbCR & all_the_tags
		' If InStr(all_the_tags, "SNAP") AND InStr(all_the_tags, "MFIP") AND InStr(all_the_tags, "DWP") AND InStr(all_the_tags, "Health Care") AND InStr(all_the_tags, "HS/GRH") AND InStr(all_the_tags, "Adult") AND InStr(all_the_tags, "EMER") AND InStr(all_the_tags, "LTC") Then
		' MsgBox script_name & vbCr & vbCr & all_the_tags & vbCr &_
		' 	   "SNAP - " & InStr(all_the_tags, "SNAP") & vbCr &_
		' 	   "MFIP - " & InStr(all_the_tags, "MFIP") & vbCr &_
		' 	   "DWP - " & InStr(all_the_tags, "DWP") & vbCr &_
		' 	   "Health Care - " & InStr(all_the_tags, "Health Care") & vbCr &_
		' 	   "HS/GRH - " & InStr(all_the_tags, "HS/GRH") & vbCr &_
		' 	   "Adult Cash - " & InStr(all_the_tags, "Adult Cash") & vbCr &_
		' 	   "EMER - " & InStr(all_the_tags, "EMER") & vbCr &_
		' 	   "PROGRAMS HELPED - " & programs_helped
		If DateDiff("m", release_date, DateAdd("m", -2, date)) <= 0 then
			description = "NEW " & release_date & "!!! " & description
			If left(description, 3) <> "---" Then description = "--- " & description
		End if
		If in_testing = TRUE Then description = "IN TESTING - " & description
		if left(script_name, 4) = "REPT" Then description = description & " --- Reads details from MAXIS REPT Lists"
		If script_name = "Update Worker Signature" Then description = "Sets or updates the default worker signature for this user."
		If script_name = "EMPS" Then description = description & " --- EMPS Panel Information in a List"
		If script_name = "DAIL Report" Then description = description & " --- List of DAILs selected by Type"
		If script_name ="Delete DAIL Tasks" Then description = description & " --- USE WITH CAUTION! Deletes info from SQL Database."
		If script_name ="Open Interview PDF" Then description = description & " --- Opens a PDF generated from NOTES - Interview if not yet in ECF."

	end property


    public sub show_button(see_the_button)
        see_the_button = FALSE
        If in_testing = TRUE Then
            For each tester in tester_array
                If user_ID_for_validation = tester.tester_id_number Then
                    Select Case testing_category
                        Case "ALL"
                            see_the_button = TRUE
                        Case "GROUP" ' ADD OPTION FOR the_selection to be an array'
                            For each group in tester.tester_groups
                                For each selection in testing_criteria
                                    selection = trim(selection)
                                    If UCase(selection) = UCase(group) Then see_the_button = TRUE
                                    ' MsgBox "Group - " & group & vbNewLine & "Selection - " & selection & vbNewLine & "see the button - " & see_the_button
                                    selected_group = group
                                Next
                            Next
                            selected_group = selection
						Case "PROGRAM" ' ADD OPTION FOR the_selection to be an array'
							For each prog in tester.tester_programs
								For each selection in testing_criteria
									selection = trim(selection)
									If UCase(selection) = UCase(prog) Then see_the_button = TRUE
									' MsgBox "Group - " & group & vbNewLine & "Selection - " & selection & vbNewLine & "see the button - " & see_the_button
									selected_prog = prog
								Next
							Next
							selected_prog = selection
                        Case "REGION"
                            For each selection in testing_criteria
                                selection = trim(selection)
                                If UCase(selection) = UCase(tester.tester_region) Then see_the_button = TRUE
                            Next
                        Case "POPULATION"
                            For each selection in testing_criteria
                                selection = trim(selection)
                                If UCase(selection) = UCase(tester.tester_population) Then see_the_button = TRUE
                            Next
                        Case "SCRIPT"
                            For each each_script in tester.tester_scripts
                                script_file_name = script_name & ".vbs"
                                If script_file_name = each_script Then see_the_button = TRUE
                            Next
                    End Select
                    If tester.tester_population = "BZ" Then see_the_button = TRUE
                End If
            Next
        Else
            see_the_button = TRUE
        End If
    end sub
end class

'Needs to determine MyDocs directory before proceeding.
Set wshshell = CreateObject("WScript.Shell")
user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

favorites_text_file_location = user_myDocs_folder & "\scripts-favorites.txt"
hotkeys_text_file_location = user_myDocs_folder & "\scripts-hotkeys.txt"
'Opening the favorites text
Dim oTxtFile
With (CreateObject("Scripting.FileSystemObject"))
    favorites_exist = ""
	'>>> If the file exists, we will grab the list of the user's favorite scripts and run the favorites menu.
	If .FileExists(favorites_text_file_location) Then
        favorites_exist = TRUE
		Set fav_scripts = CreateObject("Scripting.FileSystemObject")
		Set fav_scripts_command = fav_scripts.OpenTextFile(favorites_text_file_location)
		fav_scripts_array = fav_scripts_command.ReadAll
		IF fav_scripts_array <> "" THEN favorites_text_file_array = fav_scripts_array
		fav_scripts_command.Close

        favorites_text_file_array = trim(favorites_text_file_array)
        favorites_text_file_array = split(favorites_text_file_array, vbNewLine)
	ELSE
		'>>> ...otherwise, if the file does not exist, the script will require the user to select their favorite scripts.
		favorites_exist = FALSE
	END IF
END WITH
'INSTRUCTIONS: simply add your new script_bowie	below. Scripts are listed in alphabetical order first by category, then by script name. Copy a block of code from above and paste your script info in. The function does the rest.

'INSTANCE TEMPLATE'
' script_num = script_num + 1						'Increment by one
' ReDim Preserve script_array(script_num)
' Set script_array(script_num) = new script_bowie
' script_array(script_num).script_name 			= ""																		'Script name
' ' script_array(script_num).description 			= ""
' script_array(script_num).category               = ""
' script_array(script_num).workflows              = ""
' script_array(script_num).tags                   = array("")
' script_array(script_num).dlg_keys               = array("")
' script_array(script_num).subcategory            = array("")
' script_array(script_num).release_date           = #10/01/2000#

'ACTIONS SCRIPTS=====================================================================================================================================
script_num = 0
ReDim Preserve script_array(script_num)
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 			= "DAIL Scrubber"																		'Script name
' script_array(script_num).description 			= "Runs the DAILs from DAIL."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).keywords               = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 			= "7th Sanction Identifier"																		'Script name
' script_array(script_num).description 			= "Pulls a list of active MFIP cases that may meet 7th sanction criteria into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("MFIP", "Reports")
script_array(script_num).dlg_keys               = array("Ex", "Ev")
script_array(script_num).subcategory            = array("ENHANCED LISTS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "12 Month Contact"																		'Script name
' script_array(script_num).description 			= "Sends a MEMO to the client of their reporting responsibilities (required for SNAP 2-yr certifications, per POLI/TEMP TE02.08.165)."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Sm")
script_array(script_num).subcategory            = array("SNAP")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 			= "ABAWD Exemption"																		'Script name
' script_array(script_num).description 			= "Updates FSET/ABAWD coding on STAT/WREG and case notes ABAWD exemptions."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("ABAWD", "Application", "Communication", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Up")
script_array(script_num).subcategory            = array("ABAWD")
script_array(script_num).release_date           = #09/25/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 			= "ABAWD FIATer"																		'Script name
' script_array(script_num).description 			= "FIATS SNAP eligibility, income, and deductions for HH members with more than 3 counted months on the ABAWD tracking record."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("ABAWD", "Application", "Communication", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Fi", "Up")
script_array(script_num).subcategory            = array("ABAWD")
script_array(script_num).release_date           = #01/17/2017#
script_array(script_num).retirement_date        = #06/03/2020#					'Script removed during the COVID-19 PEACETIME STATE OF EMERGENCY
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= True

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "ABAWD FSET Exemption Check"																		'Script name
' script_array(script_num).description 			= "Double checks a case to see if any possible ABAWD/FSET exemptions exist."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("ABAWD", "Application", "Communication", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Ev")
script_array(script_num).subcategory            = array("ABAWD")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)   'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name 			= "ABAWD Report"											'Script name
' script_array(script_num).description 			= "BULK script that gathers ABAWD/FSET codes for members on SNAP/MFIP active cases."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "Monthly Tasks", "SNAP", "MFIP")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).retirement_date        = #06/14/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "ABAWD Screening Tool"
' script_array(script_num).description 			= "A tool to walk through a screening to determine if client is ABAWD."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("ABAWD", "Application", "Communication", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Ev")
script_array(script_num).subcategory            = array("ABAWD")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "ABAWD Tracking Record"																		'Script name
' script_array(script_num).description 			= "Template for documenting details about the ABAWD actvity for the case."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("ABAWD", "Application", "Communication", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #09/25/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

'script added during the COVID-19 PEACETIME STATE OF EMERGENCY
script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "ABAWD Waived Approval"																		'Script name
' script_array(script_num).description 			= "Documenting approval of SNAP for a case with ABAWD Waived during the pandemic."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("ABAWD", "SNAP")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #09/25/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Add GRH Rate 2 to MMIS"
' script_array(script_num).description 			= "Adds new supplemental service rate (SSR) agreements to MMIS for GRH Rate 2 cases."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Communication", "HS/GRH", "Reviews")
script_array(script_num).dlg_keys               = array("Cn", "Up")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #08/13/2018#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Add WCOM"																		'Script name
' script_array(script_num).description 			= "All-in-one WCOM selection menu."
script_array(script_num).category               = ""
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("ABAWD", "Application", "Assets", "Communication", "Deductions", "Health Care", "LTC", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Exp", "Sw")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #09/27/2018#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "All Scripts"																		'Script name
' script_array(script_num).description 			= "Template for documenting details about an appeal, and the appeal process."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Support")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("TOOL")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #08/09/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Appeals"																		'Script name
' script_array(script_num).description 			= "Template for documenting details about an appeal, and the appeal process."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("SNAP", "MFIP", "Adult Cash", "Communication", "DWP", "EMER", "Health Care", "HS/GRH", "LTC")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #12/12/2016#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Application Check"																		'Script name
' script_array(script_num).description 			= "Template for documenting details and tracking pending cases."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Oa", "Tk")
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #12/12/2016#
script_array(script_num).hot_topic_date			= #1/19/2021#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/NOTES-%E2%80%93-APPLICATION-CHECK-enhancement.aspx"
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Application Received"																		'Script name
' script_array(script_num).description 			= "Template for documenting details about an application recevied."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Exp", "Oe", "Sm")
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Approved Programs"																		'Script name
' script_array(script_num).description 			= "Template for when you approve a client's programs."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "Communication", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Exp")
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= True

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Asset Reduction"
' ' script_array(script_num).description 			= "Template for documenting pending and resolving an asset reduction."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "Assets", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "Reviews")
script_array(script_num).dlg_keys               = array("Cn", "Tk")
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #01/19/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)   'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name 			= "Auto-Dialer Case Status"											'Script name
' script_array(script_num).description 			= "BULK script that gathers case status for cases with recerts for SNAP/MFIP the previous month."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "Monthly Tasks", "SNAP", "MFIP")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "AVS"													'Script name
' script_array(script_num).description		    = "Supports for AVS forms and AVS Submission/Results processes for MA-ABD Cases."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Health Care")
script_array(script_num).dlg_keys               = array("Cn", "Tk", "Wrd")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #04/27/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "AVS Report"													'Script name
' script_array(script_num).description		    = "BULK script that supports the AVS processing needs for active MA recipients."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "Health Care")
script_array(script_num).dlg_keys               = array("Ex", "Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array(script_num)	    'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	   'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "AVS Submitted"
' script_array(script_num).description		    = "Creates a case note and sets a 10-day TIKL to check status of AVS submission."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Health Care")
script_array(script_num).dlg_keys               = array("Cn", "Tk")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).retirement_date        = #06/25/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

' script_num = script_num + 1							   'Increment by one
' ReDim Preserve script_array(script_num)	    'Resets the array to add one more element to it
' Set script_array(script_num) = new script_bowie	   'Set this array element to be a new script. Script details below...
' script_array(script_num).script_name		    = "Basket Review"
' script_array(script_num).file_name			= "basket-review.vbs"
' ' script_array(script_num).description		    = "A script that creates a report of cases and pages pending on a list of baskets."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "BILS Updater"
' script_array(script_num).description 			= "Updates a BILS panel with reoccurring or actual BILS received."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Communication", "Deductions", "Health Care", "LTC", "Reviews")
script_array(script_num).dlg_keys               = array("Up")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 		    = "Budget Estimator"											'Script name
' script_array(script_num).description 			= "UTILITIES script that can be used to calculate an expected budget outside of MAXIS."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("QI", "SNAP", "MFIP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Ev")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).retirement_date        = #09/29/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)	'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name			= "BULK - Inactive Transfer"													'Script name
' script_array(script_num).description 			= "Script to transfer inactive cases via SPEC/XFER"
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "Monthly Tasks", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Up", "Ex", "Ev")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array(script_num)	    'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	   'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "BULK - REPT USER List"
' script_array(script_num).description		    = "Report to pull MAXIS USER detail into an Excel Spreadsheet."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Burial Assets"
' script_array(script_num).description 			= "Template for burial assets."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Assets", "Health Care", "LTC", "Reviews")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "CAF"
' script_array(script_num).description 			= "Template for when you're processing a CAF. Works for intake as well as recertification and reapplication.*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "Assets", "Deductions", "DWP", "EMER", "HS/GRH", "MFIP", "Reviews", "SNAP", "Health Care")
script_array(script_num).dlg_keys               = array("Cn", "Exp", "Tk", "Up")
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_date			= #08/31/2021#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/New-scripts-available-September-1-for-interview-and-Expedited-SNAP.aspx"
script_array(script_num).used_for_elig			= False

script_num = script_num + 1					'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Calculate Rate 2 Units"
' script_array(script_num).description 			= "Calculates the GRH Rate 2 total units to input into MMIS."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Calculators", "Utility", "HS/GRH")
script_array(script_num).dlg_keys               = array("Ev")
script_array(script_num).subcategory            = array("TOOL")
script_array(script_num).release_date           = #08/10/2018#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Case Discrepancy"
' script_array(script_num).description 			= "Template for case noting information about a case discrepancy."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "Assets", "Communication", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Tk")
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/24/2016#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 			= "CASE NOTE from List"																		'Script name
' script_array(script_num).description 			= "Creates the same case note on cases listed in REPT/ACTV, manually entered, or from an Excel spreadsheet of your choice."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Reports", "Utility")
script_array(script_num).dlg_keys               = array("Cn", "Ex")
script_array(script_num).subcategory            = array("BULK ACTIONS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Case Transfer"																		'Script name
' script_array(script_num).description 			= "Searches caseload(s) by selected parameters. Transfers a specified number of those cases to another worker. Creates list of these cases."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Reports", "Utility", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Ex", "Up", "Cn", "Sm")
script_array(script_num).subcategory            = array("BULK ACTIONS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Change Report Form Received"
' script_array(script_num).description 			= "Template for case noting information reported from a Change Report Form."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Assets", "Communication", "Deductions", "DWP", "EMER", "Health Care", "HS/GRH", "Income", "LTC", "MFIP", "SNAP")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Change Reported"
' script_array(script_num).description 			= "Template for case noting HHLD Comp or Baby Born being reported. **More changes to be added in the future**"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Assets", "Communication", "Deductions", "DWP", "EMER", "Health Care", "HS/GRH", "Income", "LTC", "MFIP", "SNAP")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Check EDRS"
' script_array(script_num).description 			= "Checks EDRS for HH members with disqualifications on a case."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "MFIP", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Ev")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Check SNAP for GA RCA"
' script_array(script_num).description 			= "Compares the amount of GA and RCA FIAT'd into SNAP and creates a list of the results."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Income", "Reports", "SNAP")
script_array(script_num).dlg_keys               = array("Ex", "Ev")
script_array(script_num).subcategory            = array("ENHANCED LISTS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Citizenship Identity Verified"
' script_array(script_num).description 			= "Template for documenting citizenship/identity status for a case."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "Communication", "DWP", "EMER", "HS/GRH", "MFIP", "SNAP", "Health Care")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Claim Referral Tracking"
' script_array(script_num).description 			= "Assists in tracking overpayments/potential overpayments on STAT/MISC and case note."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "Income", "MFIP", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Tk", "Up")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #09/25/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Client Contact"
' script_array(script_num).description 			= "Template for documenting client contact, either from or to a client."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Communication", "DWP", "EMER", "HS/GRH", "MFIP", "SNAP", "Health Care")
script_array(script_num).dlg_keys               = array("Cn", "Tk")
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)	'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name			= "Close MMIS Rate 2 in MMIS"													'Script name
' script_array(script_num).description 			= "Script to assist in closing SSR agreements in MMIS for GRH Rate 2 cases."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "Monthly Tasks", "HS/GRH")
script_array(script_num).dlg_keys               = array("Ex", "Up", "Ev")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Closed Programs"
' script_array(script_num).description 			= "Template for indicating which programs are closing, and when. Also case notes intake/REIN dates based on various selections."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "Communication", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Sw")
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= True

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "COLA Decimator"													'Script name
' script_array(script_num).description		    = "BULK script that deletes and case notes auto-approval COLA messages."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "Adult Cash", "HS/GRH", "MFIP", "SNAP")
script_array(script_num).dlg_keys               = array("Ex", "Up", "Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array(script_num)	    'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	   'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Complete Phone CAF"
' script_array(script_num).description		    = "Complete all of the CAF questions in a script to output to a PDF for transcribing into ECF CAF form."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Support", "Utility", "SNAP", "DWP", "MFIP", "Adult Cash", "HS/GRH", "EMER")
script_array(script_num).dlg_keys               = array("Cn", "Wrd")
script_array(script_num).subcategory            = array("TOOL")
script_array(script_num).release_date           = #11/24/2020#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array(script_num)	    'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	   'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Contact Knowledge Now"
' script_array(script_num).description		    = "Used to submit an email to the QI Knowledge Now team."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Support", "Utility", "SNAP", "MFIP", "DWP", "Adult Cash", "HS/GRH", "Health Care", "EMER")
script_array(script_num).dlg_keys               = array("Oe")
script_array(script_num).subcategory            = array("TOOL", "POLICY")
script_array(script_num).release_date           = #08/19/2020#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array(script_num)	    'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	   'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Copy Case Data for Training"
' script_array(script_num).description		    = "Copies data from a case to a spreadsheet to be run on the Training Case Generator."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("MFIP", "SNAP", "Adult Cash", "Health Care")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array(script_num)	    'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	   'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Copy Panels to Word"
' script_array(script_num).description		    = "Copies MAXIS panels to Word en masse for a case for easier review."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("SNAP", "MFIP", "Health Care", "Adult Cash")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Counted ABAWD Months"
' script_array(script_num).description 			= "Displays all markings on ABAWD tracking record and issuances for affected programs in Excel."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("ABAWD", "Application", "Communication", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Ex", "Ev")
script_array(script_num).subcategory            = array("ABAWD")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Create Fake CS DAILs as TIKLs"
' script_array(script_num).description 			= "Mocks up a DAIL message that looks like CSES Disb message but is made from TIKL to build and test CSES Scrubber."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "SNAP", "MFIP")
script_array(script_num).dlg_keys               = array("Tk")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #03/19/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1
ReDim Preserve script_array(script_num)
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 		    = "CS Good Cause"											'Script name
' script_array(script_num).description 			= "Completes updates to ABPS and case notes actions taken."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("MFIP", "DWP")
script_array(script_num).dlg_keys               = array("Up", "Cn", "Sm")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "CSR"
' script_array(script_num).description 			= "Template for the Combined Six-month Report (CSR).*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Assets", "Deductions", "Health Care", "HS/GRH", "Income", "LTC", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Tk")
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "DAIL CCD"													'Script name
' script_array(script_num).description		    = "BULK script that captures, case notes and deletes specific DAILS based on content, and collects them into an Excel spreadsheet."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Ex", "Up", "Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "DAIL Decimator"													'Script name
' script_array(script_num).description		    = "BULK script that deletes specific DAILS based on content, and collects them into an Excel spreadsheet."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Ex", "Up", "Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "DAIL Report"
' script_array(script_num).description 			= "Pulls a list of DAILS in DAIL/DAIL into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Reports")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("BULK LISTS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Deceased Client Summary"																		'Script name
' script_array(script_num).description 			= "Adds details about a deceased client to a CASE/NOTE."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP", "Adult Cash", "LTC")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #04/25/2016#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Delete DAIL Tasks"													'Script name
' script_array(script_num).description		    = "Script Function that will delete Task-based DAIL's from SQL Database. Use with caution."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Oe")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #02/11/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Denied Programs"																		'Script name
' script_array(script_num).description 			= "Template for indicating which programs you've denied, and when. Also case notes intake/REIN dates based on various selections."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "Communication", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Sw")
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= True

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array(script_num)	    'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	   'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Disaster Food Replacement"
' script_array(script_num).description		    = "Case note to help with replacing food destroyed in a disaster"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "SNAP", "MFIP")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #06/01/2020#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Documents Received"
' script_array(script_num).description 			= "Template for case noting information about documents received."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("ABAWD", "Adult Cash", "Application", "Assets", "Communication", "Deductions", "DWP", "EMER", "Health Care", "HS/GRH", "Income", "LTC", "MFIP", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Tk", "Up")
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Drug Felon"
' script_array(script_num).description 			= "Template for noting drug felon info."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "Communication", "DWP", "HS/GRH", "MFIP", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array(script_num)	    'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	   'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Drug Felon List"
' script_array(script_num).description		    = "Reviews the Drug Felon list from DHS to update these cases."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "MFIP", "DWP", "SNAP")
script_array(script_num).dlg_keys               = array("Ex", "Cn", "Up")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "DWP ES Referral"																		'Script name
' script_array(script_num).description 			= "Creates a case note, a manual referral in INFC/WF1M and sends a SPEC/MEMO to the client."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Communication", "DWP")
script_array(script_num).dlg_keys               = array("Cn", "Sm", "Up")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).retirement_date        = #06/04/2020#					'script removed during the COVID-19 PEACETIME STATE OF EMERGENCY
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Earned Income Budgeting"
' script_array(script_num).description 			= "Reviews income, Updates JOBS, CASE/NOTE for multiple Earned Income Panels on a single case."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "Communication", "DWP", "EMER", "Health Care", "HS/GRH", "Income", "LTC", "MFIP", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Exp", "Up")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #03/05/2019#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "EDRS DISQ Match Found"
' script_array(script_num).description 			= "Template for noting the action steps when a SNAP recipient has an eDRS DISQ per TE02.08.127."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "MFIP", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Tk")
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Eligibility Notifier"																		'Script name
' script_array(script_num).description 			= "Sends a MEMO informing client of possible program eligibility for SNAP, MA, MSP, MNsure or Cash."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Communication", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Sm")
script_array(script_num).subcategory            = array("HEALTH CARE", "SNAP", "Cash")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Emergency"
' script_array(script_num).description 			= "Template for EA/EGA applications.*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "EMPS"
' script_array(script_num).description 			= "Pulls a list of STAT/EMPS information into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("DWP", "MFIP", "Reports")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("BULK LISTS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "EMPS Updater"
' script_array(script_num).description 			= "Updates the EMPS panel, and case notes when for Child Under 12 Months Exemptions."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Communication", "DWP", "MFIP", "Reviews")
script_array(script_num).dlg_keys               = array("Cn", "Tk", "Up")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #11/03/2016#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Enroll in Script Demo"
' script_array(script_num).description 			= "Script to display any upcoming BlueZone Script Demos that are scheduled and allow you to sign up and add to your Outlook Schedule."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Support", "Utility")
script_array(script_num).dlg_keys               = array("Oe", "Oa")
script_array(script_num).subcategory            = array("TOOL", "POLICY")
script_array(script_num).release_date           = #08/10/2020#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Expedited Determination"
' script_array(script_num).description 			= "Template for noting detail about how expedited was determined for a case."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Assets", "Deductions", "Income", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Exp", "Ev")
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_date			= #08/31/2021#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/New-scripts-available-September-1-for-interview-and-Expedited-SNAP.aspx"
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Expedited Determination Report"
' script_array(script_num).description 			= "Template for noting detail about how expedited was determined for a case."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "BZ", "Monthly Tasks", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Exp", "Ev")
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_date			= #08/31/2021#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/New-scripts-available-September-1-for-interview-and-Expedited-SNAP.aspx"
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Expedited Review"													'Script name
' script_array(script_num).description		    = "BULK script to support reviewing and categorizing expedited SNAP cases in Hennepin County."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "SNAP")
script_array(script_num).dlg_keys               = array("Ev", "Ex", "Oe")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Expedited Screening"
' script_array(script_num).description 			= "Template for screening a client for expedited status."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Assets", "Deductions", "Income", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Exp", "Ev")
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "FIAT GA-RCA Into SNAP Budget"
' script_array(script_num).description 			= "FIATs GA or RCA income into SNAP budget for each month through cuurent month plus one."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Communication", "Adult Cash", "Income", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Fi")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #09/25/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= True

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Find MAEPD MEDI CEI"
' script_array(script_num).description 			= "Creates a list of cases and clients active on MA-EPD and Medicare Part B that are eligible for Part B reimbursement."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Health Care", "Reports")
script_array(script_num).dlg_keys               = array("Ex", "Ev")
script_array(script_num).subcategory            = array("ENHANCED LISTS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Find Panel Update Date"
' script_array(script_num).description 			= "Creates a list of cases from a caseload(s) showing when selected panels have been updated."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Reports", "Utility")
script_array(script_num).dlg_keys               = array("Ex", "Ev")
script_array(script_num).subcategory            = array("ENHANCED LISTS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).retirement_date		= #06/09/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1
ReDim Preserve script_array(script_num)
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 		    = "Find Q Flow Population"											'Script name
' script_array(script_num).description 			= "Utility to identfy Q Flow Populations by basket number."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("")
script_array(script_num).dlg_keys               = array("Ev")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #12/09/2020#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Fraud Info"
' script_array(script_num).description 			= "Template for noting fraud info."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Communication", "DWP", "EMER", "HS/GRH", "MFIP", "LTC", "SNAP", "Health Care")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "FSS Info"
' script_array(script_num).description 			= "Pulls a list of FSS identified info from EMPS and DISA into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("MFIP", "Reports")
script_array(script_num).dlg_keys               = array("Ex", "Ev")
script_array(script_num).subcategory            = array("ENHANCED LISTS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "FSS Status Change"
' script_array(script_num).description 			= "Updates STAT with information from a Status Update."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "MFIP", "Reviews")
script_array(script_num).dlg_keys               = array("Cn", "Tk", "Up")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #11/03/2016#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "FUBU"													'Script name
' script_array(script_num).description		    = "Get a sortable list of all of the scripts from the COMPLETE LIST OF SCRIPTS - the new one."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ")
script_array(script_num).dlg_keys               = array("Ex", "Ev")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).retirement_date		= #08/09/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "GA Basis of Eligibility"
' script_array(script_num).description 			= "Template to document the basis of eligibility and verification of the basis for GA recipients."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "Communication", "Reviews")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/20/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Get basket number"													'Script name
' script_array(script_num).description		    = "BULK script that will obtain the basket number and population."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ")
script_array(script_num).dlg_keys               = array("Ex", "Ev")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).retirement_date		= #06/09/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "GRH NON HRF POSTPAY"
' script_array(script_num).description 			= "Case note template for GRH post pay cases."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Assets", "Deductions", "HS/GRH", "Income", "Reviews")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "GRH Professional Need"
' script_array(script_num).description 			= "Pulls a list of active GRH cases and identified info into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("HS/GRH", "Reports")
script_array(script_num).dlg_keys               = array("Ex", "Ev")
script_array(script_num).subcategory            = array("ENHANCED LISTS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).retirement_date        = #06/14/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "HC Renewal"
' script_array(script_num).description 			= "Template for HC renewals.*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Assets", "Deductions", "Health Care", "Income", "Reviews")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "HCAPP"
' script_array(script_num).description 			= "Template for HCAPPs.*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Assets", "Application", "Deductions", "Health Care", "Income")
script_array(script_num).dlg_keys               = array("Cn", "Tk")
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Health Care Transition"
' script_array(script_num).description 			= "Template for the METS to MAXIS and MAXIS to METS transition process."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Assets", "Communication", "Deductions", "Health Care", "Income", "Reviews")
script_array(script_num).dlg_keys               = array("Cn", "Sm")
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Homeless Discrepancy"
' script_array(script_num).description 			= "Pulls a list of active SNAP/MFIP cases with identified info into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Reports")
script_array(script_num).dlg_keys               = array("Ex", "Ev")
script_array(script_num).subcategory            = array("ENHANCED LISTS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).retirement_date		= #06/09/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "HRF"
' script_array(script_num).description 			= "Template for HRFs (for GRH, use the ''GRH - HRF'' script).*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Assets", "Deductions", "HS/GRH", "Income", "LTC", "MFIP", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "IMIG - EMA"
' script_array(script_num).description 			= "Template for EMA applications."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Assets", "Deduction", "Health Care", "Income", "Reviews")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("IMIG")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "IMIG - Sponsor Income"
' script_array(script_num).description 			= "Template for the sponsor income deeming calculation (it will also help calculate it for you)."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "Communication", "DWP", "EMER", "Health Care", "HS/GRH", "Income", "LTC", "MFIP", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Ev")
script_array(script_num).subcategory            = array("IMIG")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "IMIG - STATUS"
' script_array(script_num).description 			= "Template for the SAVE system for verifying immigration status."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "Communication", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Up", "Oe", "Oa")
script_array(script_num).subcategory            = array("IMIG")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Individual Appointment Letter"													'Script name
' script_array(script_num).description		    = "Sends an appointment letter for a single case, with the same wording as On Demand Applications"
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("QI", "SNAP", "MFIP", "DWP", "Adult Cash", "HS/GRH")
script_array(script_num).dlg_keys               = array("Cn", "Sm")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Individual NOMI"													'Script name
' script_array(script_num).description		    = "Sends a NOMI for a single case, with the same wording as On Demand Applications"
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("QI", "SNAP", "MFIP", "DWP", "Adult Cash", "HS/GRH")
script_array(script_num).dlg_keys               = array("Cn", "Sm")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)	'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name			= "Individual Recertification Notices"													'Script name
' script_array(script_num).description 			= "NOTICES Script that will send ODW Recert Appointment Letter or NOMI on a single case."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "Monthly Tasks", "SNAP", "MFIP")
script_array(script_num).dlg_keys               = array("Cn", "Sm")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Insert MBI from MMIS"
' script_array(script_num).description 			= "Update STAT/MEDI with MBI number from RMCR in MMIS."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Navigation", "Utility")
script_array(script_num).dlg_keys               = array("Up")
script_array(script_num).subcategory            = array("MAXIS")
script_array(script_num).release_date           = #05/15/2020#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Interview"
' script_array(script_num).description 			= "Workflow for Interview process."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "Application", "Reviews", "SNAP", "MFIP", "DWP", "Adult Cash", "EMER", "HS/GRH")
script_array(script_num).dlg_keys               = array("Cn", "Sm", "Wrd")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #09/01/2021#								'Testing - #07/01/2021#
script_array(script_num).hot_topic_date         = #08/31/2021#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/New-scripts-available-September-1-for-interview-and-Expedited-SNAP.aspx"
script_array(script_num).retirement_date        = ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Interview Completed"
' script_array(script_num).description 			= "Template to case note an interview being completed but no stat panels updated."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "DWP", "EMER", "HS/GRH", "MFIP", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Oa")
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Interview No Show"
' script_array(script_num).description 			= "Template for case noting a client's no-showing their in-office or phone appointment."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "DWP", "EMER", "HS/GRH", "MFIP", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Exp")
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).retirement_date        = #05/12/2020#					'script removed during the COVID-19 PEACETIME STATE OF EMERGENCY
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)	'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name			= "Job Change Reported"													'Script name
' script_array(script_num).description 			= "Creates  or updates JOBS panel, CASE/NOTE and TIKL when a a change is reported about a JOB."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "Communication", "DWP", "EMER", "Health Care", "HS/GRH", "Income", "LTC", "MFIP", "Reviews", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Cn", "Up", "Tk")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name			= "Language Stats"													'Script name
' script_array(script_num).description 			= "Collects language statistics by language and region. Take approximately 10 hours to run."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "Monthly Tasks")
script_array(script_num).dlg_keys               = array("Ev", "Ex")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).retirement_date		= #06/09/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name			= "Lost ApplyMN"													'Script name
' script_array(script_num).description 			= "Sends an Email request to search for an ApplyMN that is not in ECF."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Support", "Utility")
script_array(script_num).dlg_keys               = array("Oe")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #08/11/2020#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "LTC - 5181"
' script_array(script_num).description 			= "Template for processing DHS-5181."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Communication", "Health Care", "LTC", "Reviews")
script_array(script_num).dlg_keys               = array("Cn", "Up")
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "LTC - Application Received"
' script_array(script_num).description 			= "Template for initial details of a LTC application.*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Assets", "Deductions", "LTC", "Income")
script_array(script_num).dlg_keys               = array("Cn", "Tk", "Up")
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "LTC - Asset Assessment"
' script_array(script_num).description 			= "Template for the LTC asset assessment. Will enter both person and case notes if desired."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Assets", "LTC")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "LTC - COLA Summary"
' script_array(script_num).description 			= "Template to summarize actions for the changes due to COLA.*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "Deductions", "Health Care", "Income", "LTC")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "LTC - Hospice Form Received"
' script_array(script_num).description 			= "Template for case noting entry or exit to Hospice.*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "Health Care", "LTC")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "LTC - Intake Approval"
' script_array(script_num).description 			= "Template for use when approving a LTC intake.*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Assets", "Communication", "Deductions", "LTC", "Income")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "LTC - MA Approval"
' script_array(script_num).description 			= "Template for approving LTC MA (can be used for changes, initial application, or recertification).*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Communication", "Deductions", "LTC", "Income", "Reviews")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "LTC - Renewal"
' script_array(script_num).description 			= "Template for LTC renewals.*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Assets", "Communication", "Deductions", "LTC", "Income", "Reviews")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "LTC - Transfer Penalty"
' script_array(script_num).description 			= "Template for noting a transfer penalty."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Assets", "Communication", "LTC")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "LTC Asset Transfer"
' script_array(script_num).description 			= "Sends a MEMO to a LTC client regarding asset transfers. "
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Assets", "Communication", "LTC", "Reviews")
script_array(script_num).dlg_keys               = array("Sm")
script_array(script_num).subcategory            = array("HEALTH CARE")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "LTC Spousal Allocation FIATer"
' script_array(script_num).description 			= "FIATs a spousal allocation across a budget period."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Deductions", "Income", "LTC", "Reviews")
script_array(script_num).dlg_keys               = array("Fi")
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= True

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "LTC ICF-DD Deduction FIATer"																			'Script name
' script_array(script_num).description 			= "FIATs earned income and deductions across a budget period."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Deductions", "Income", "LTC", "Reviews")
script_array(script_num).dlg_keys               = array("Fi")
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #05/23/2016#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= True

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "LTC-GRH List Generator"
' script_array(script_num).description 			= "Creates a list of FACIs, AREPs, and waiver types assigned to the various cases in a caseload(s)."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("HS/GRH", "Health Care", "LTC", "Reports")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("BULK LISTS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "MA FIATER for GRH MSA"
' script_array(script_num).description 			= "Script that will FIAT MA Eligibility to remove the ‘X’ method from the Health Care span."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Reviews", "Health Care", "HS/GRH", "Adult Cash")
script_array(script_num).dlg_keys               = array("Fi")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #04/02/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= True

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "MA Inmate Application WCOM"
' script_array(script_num).description 			= "Sends a WCOM on a MA notice for Inmate Applications"
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Communication", "Health Care")
script_array(script_num).dlg_keys               = array("Cn", "Sw")
script_array(script_num).subcategory            = array("HEALTH CARE")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "MA-EPD EI FIAT"
' script_array(script_num).description 			= "FIATs MA-EPD earned income (JOBS income) to be even across an entire budget period."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Health Care", "Income", "Reviews")
script_array(script_num).dlg_keys               = array("Fi")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= True

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "MA-EPD No Initial Premium"
' script_array(script_num).description 			= "Sends a WCOM on a denial for no initial MA-EPD premium."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Communication", "Health Care", "Reviews")
script_array(script_num).dlg_keys               = array("Sw")
script_array(script_num).subcategory            = array("HEALTH CARE")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)	'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name			= "MAXIS to METS Conversion"													'Script name
' script_array(script_num).description 			= "BULK script to collect case information for cases that may need to convert from MAXIS to METS."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "Monthly Tasks", "Health Care")
script_array(script_num).dlg_keys               = array("Ev", "Ex")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).retirement_date		= #06/09/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array(script_num)	   'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	  'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "MEMO from List"
' script_array(script_num).description		    = "Creates the same MEMO on cases listed in REPT/ACTV, manually entered, or from an Excel spreadsheet of your choice."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("")
script_array(script_num).dlg_keys               = array("Ex", "Sm")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "MEMO to Word"
' script_array(script_num).description 			= "Copies a MEMO or WCOM from MAXIS and formats it in a Word Document."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Communication", "DWP", "EMER", "HS/GRH", "MFIP", "LTC", "SNAP", "Health Care")
script_array(script_num).dlg_keys               = array("Sm", "Wrd")
script_array(script_num).subcategory            = array("WORD DOCS")
script_array(script_num).release_date           = #02/21/2018#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Method B WCOM"													'needs spaces to generate button width properly.
' script_array(script_num).description 			= "Makes detailed WCOM regarding spenddown vs. recipient amount for method B HC cases."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Communication", "Deductions", "Health Care", "Income", "LTC", "Reviews")
script_array(script_num).dlg_keys               = array("Sw")
script_array(script_num).subcategory            = array("HEALTH CARE")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "METS Retro Health Care"
' script_array(script_num).description 			= "Template and email support for when METS retro coverage has been requested."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Health Care")
script_array(script_num).dlg_keys               = array("Cn", "Oe")
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "MFIP Sanction"
' script_array(script_num).description 			= "Pulls a list of active MFIP cases with identified info into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("MFIP", "Reports")
script_array(script_num).dlg_keys               = array("Ex", "Ev")
script_array(script_num).subcategory            = array("ENHANCED LISTS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "MFIP Sanction And DWP Disqualification"
' script_array(script_num).description 			= "Template for MFIP sanctions and DWP disqualifications, both CS and ES."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "DWP", "MFIP", "Reviews")
script_array(script_num).dlg_keys               = array("Cn", "Sw", "Tk", "Up")
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name			= "MFIP Sanction FIATer"											'Script name
' script_array(script_num).description 			= "FIATs MFIP sanction actions for CS, ES and both types of sanctions."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("MFIP")
script_array(script_num).dlg_keys               = array("Fi", "Ev")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= True

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "MFIP to SNAP Transition"
' script_array(script_num).description 			= "Template for noting when closing MFIP and opening SNAP."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "Deductions", "Income", "MFIP", "SNAP")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "MSQ"
' script_array(script_num).description 			= "Template for noting Medical Service Questionaires (MSQ)."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "Health Care", "LTC")
script_array(script_num).dlg_keys               = array("Cn", "Up")
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "New Job Reported"
' script_array(script_num).description 			= "Creates a JOBS panel, CASE/NOTE and TIKL when a new job is reported. Use the DAIL scrubber for new hire DAILs."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "Communication", "DWP", "EMER", "Health Care", "HS/GRH", "Income", "MFIP", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Ti", "Up")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).retirement_date		= #09/25/2020#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "On Demand Notes"
' script_array(script_num).description 			= "Creates some case notes and assists with emails on reoccuring process issues."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("QI", "SNAP", "MFIP", "DWP", "Adult Cash", "HS/GRH")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #12/07/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		 'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	 'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "On Demand Waiver Applications"									'Script name
' script_array(script_num).description		    = "BULK script to collect information for cases that require an interview for the On Demand Waiver."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("QI", "SNAP", "MFIP", "DWP", "Adult Cash", "HS/GRH")
script_array(script_num).dlg_keys               = array("Ex", "Ev", "Sm", "Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)   'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "On Demand Waiver Recertifications"													'Script name
' script_array(script_num).description		    = "BULK script to send notices for cases at recertification that require an interview for the On Demand Waiver."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "Monthly Tasks")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).retirement_date		= #06/09/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Open Interview PDF"
' script_array(script_num).description 			= "Template for case noting information about sending a notice."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
' script_array(script_num).tags                   = array("Communication", "Application", "Reviews", "SNAP", "MFIP", "DWP", "Adult Cash", "EMER", "HS/GRH")
script_array(script_num).tags                   = array("Communication", "Application", "Reviews")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("TOOL")
script_array(script_num).release_date           = #09/01/2021#					'Testing - '#07/29/2021#
script_array(script_num).hot_topic_date			= #08/31/2021#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/New-scripts-available-September-1-for-interview-and-Expedited-SNAP.aspx"
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Other Benefits Referral"
' script_array(script_num).description 			= "Template for case noting information about sending a notice."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "Communication", "Health Care", "Income", "LTC", "MFIP", "Reviews")
script_array(script_num).dlg_keys               = array("Cn", "Tk")
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Out Of State"
' script_array(script_num).description 			= "Generates out of state inquiry (MS Word document) notice that can be used to fax."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Communication", "DWP", "EMER", "HS/GRH", "MFIP", "LTC", "SNAP", "Health Care")
script_array(script_num).dlg_keys               = array("Cn", "Wrd")
script_array(script_num).subcategory            = array("WORD DOCS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Overpayment"
' script_array(script_num).description 			= "Template for noting basic information about overpayments."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Communication", "DWP", "EMER", "Health Care", "HS/GRH", "Income", "LTC", "MFIP", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Oe", "Up")
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "PA Verif Request"
' script_array(script_num).description 			= "Creates a Word document with PA benefit totals for other agencies to determine client benefits."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
' script_array(script_num).tags                   = array("Adult Cash", "Communication", "MFIP, ""DWP", "EMER", "HS/GRH", "Health Care", "LTC", "SNAP")
script_array(script_num).tags                   = array("Adult Cash", "Communication", "MFIP", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Wrd", "Sm", "Sw")
script_array(script_num).subcategory            = array("WORD DOCS")
script_array(script_num).release_date           = #06/14/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)	'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name			= "Paperless IR"                                                       'Script name
' script_array(script_num).description 			= "Updates cases on a caseload(s) that require paperless IR processing. Does not approve cases."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "Monthly Tasks", "Health Care")
script_array(script_num).dlg_keys               = array("Ex", "Up", "Ev", "Tk")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "PF11 Actions"
' script_array(script_num).description 			= "PF11 actions for PMI merge, unactionable DAILS, duplicate case note, and MFIP spouse."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "MFIP", "Utility", "Health Care", "DWP", "HS/GRH", "SNAP", "Adult Cash", "EMER")
script_array(script_num).dlg_keys               = array("Cn", "Up", "Oa")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #07/01/2019#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "POLI TEMP List"
' script_array(script_num).description 			= "Creates a list of current POLI/TEMP topics, TEMP reference and revised date."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Reports")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("MAXIS", "POLICY")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "POLI TEMP to Word"
' script_array(script_num).description 			= "Creates a Word Document of a single POLI/TEMP reference, need the Table Number."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Utility")
script_array(script_num).dlg_keys               = array("Wrd")
script_array(script_num).subcategory            = array("MAXIS", "POLICY")
script_array(script_num).release_date           = #01/08/2019#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "PRISM Screen Finder"
' script_array(script_num).description 			= "Navigates to popular PRISM screens. The navigation window stays open until user closes it."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Navigation")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).retirement_date		= #06/09/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False


script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Proof of Relationship"
' script_array(script_num).description 			= "Template for documenting proof of relationship between a member 01 and someone else in the household."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Communication", "DWP", "MFIP", "Reviews")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1					'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "QI AVS request"
' script_array(script_num).description 			= "Creates an email requesting the QI team submit an AVS request."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Health Care", "Applications", "Reviews", "Utility")
script_array(script_num).dlg_keys               = array("Oe", "Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #03/06/2020#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		 'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	 'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "QC Results"									'Script name
' script_array(script_num).description		    = "Case note and WCOM script to support the DHS QC process."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("QI", "SNAP", "MFIP")
script_array(script_num).dlg_keys               = array("Cn", "Sw", "Oa")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "QI Renewal Accuracy"                                              'Script name
' script_array(script_num).description		    = "Template for documenting specific renewal information that has been reviewed by policy experts."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("QI", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Returned Mail Received"
' script_array(script_num).description 			= "Template for noting Returned Mail Received information."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Communication", "DWP", "Health Care", "HS/GRH", "MFIP", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Tk", "Up")
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Report to the BZST"
' script_array(script_num).description 			= "Send an idea, error report, or other kind of request to the blueZone Script Team"
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Support", "Utility")
script_array(script_num).dlg_keys               = array("Oe")
script_array(script_num).subcategory            = array("TOOL", "MAXIS", "POLICY")
script_array(script_num).release_date           = #08/01/2020#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-ACTV List"
' script_array(script_num).description 			= "Pulls a list of cases in REPT/ACTV into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Reports", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("BULK LISTS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-EOMC List"
' script_array(script_num).description 			= "Pulls a list of cases in REPT/EOMC into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Reports", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("BULK LISTS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-INAC List"
' script_array(script_num).description 			= "Pulls a list of cases in REPT/INAC into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Reports", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("BULK LISTS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-MAMS List"
' script_array(script_num).description 			= "Pulls a list of cases in REPT/MAMS into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Reports", "Health Care", "LTC")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("BULK LISTS")
script_array(script_num).release_date           = #06/27/2016#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-MFCM List"
' script_array(script_num).description 			= "Pulls a list of cases in REPT/MFCM into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Reports", "MFIP")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("BULK LISTS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-MONT List"
' script_array(script_num).description 			= "Pulls a list of cases in REPT/MONT into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Reports", "MFIP", "HS/GRH", "Health Care", "SNAP")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("BULK LISTS")
script_array(script_num).release_date           = #06/27/2016#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-MRSR List"
' script_array(script_num).description 			= "Pulls a list of cases in REPT/MRSR into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Reports", "MFIP", "HS/GRH", "Health Care", "SNAP")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("BULK LISTS")
script_array(script_num).release_date           = #06/27/2016#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-PND1 List"
' script_array(script_num).description 			= "Pulls a list of cases in REPT/PND1 into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Reports", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("BULK LISTS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-PND2 List"
' script_array(script_num).description 			= "Pulls a list of cases in REPT/PND2 into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Reports", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("BULK LISTS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-REVS List"
' script_array(script_num).description 			= "Pulls a list of cases in REPT/REVS into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Reports", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("BULK LISTS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-REVW List"
' script_array(script_num).description 			= "Pulls a list of cases in REPT/REVW into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Reports", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("BULK LISTS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-USER List"
' script_array(script_num).description 			= "Pulls a list of cases in REPT/USER into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("REPORTS")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("BULK LISTS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Request Access to PRIV Case"
' script_array(script_num).description 			= "Sends a request to QI Knowledge Now to grant access in MAXIS for a Privileged Case."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("")
script_array(script_num).dlg_keys               = array("Oe")
script_array(script_num).subcategory            = array("TOOL", "MAXIS")
script_array(script_num).release_date           = #08/19/2020#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)   'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	 'Set this array element to be a new script. Script details below...
script_array(script_num).script_name			= "Resolve HC EOMC in MMIS"													'Script name
' script_array(script_num).description 			= "BULK script that checks MMIS for all cases on EOMC for HC to ensure MMIS is set to close."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "Monthly Tasks", "Health Care")
script_array(script_num).dlg_keys               = array("Up", "Ex", "Ev")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Resources Notifier"
' script_array(script_num).description 			= "Sends a MEMO informing client of some possible outside resources."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Case notes", "MEMO", "Utility")
script_array(script_num).dlg_keys               = array("Cn", "Sm", "Wrd")
script_array(script_num).subcategory            = array("HEALTH CARE", "SNAP", "Cash")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)   'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name 			= "Review Report"											'Script name
' script_array(script_num).description 			= "BULK script that gathers reivew (ER, SR, etc.) cases and information."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "Monthly Tasks", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Ex", "Ev", "Cn", "Sm")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/20/2020#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Review Testers"										'Script name
' script_array(script_num).description		    = "Generates a list of all of the testers, which can be filtered and exported to Excel."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ")
script_array(script_num).dlg_keys               = array("Ex", "Ev")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)	'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name			= "REVW MONT Closures"													'Script name
' script_array(script_num).description 			= "Case notes all cases on REPT/REVW or REPT/MONT that are closing for missing or incomplete CAF/HRF/CSR/HC ER."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "Monthly Tasks")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).retirement_date		= #06/09/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Send CBO Manual Referrals"										'Script name
' script_array(script_num).description		    = "Sends manual referrals for a list of cases provided by Employment and Training."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "SNAP", "MFIP")
script_array(script_num).dlg_keys               = array("Ex", "Ev")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Send Email Correction"										'Script name
' script_array(script_num).description		    = "Send emails about process corrections for Expedited SNAP or On Demand."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("QI", "SNAP")
script_array(script_num).dlg_keys               = array("Oe", "Ex")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #07/21/2020#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Send SVES"
' script_array(script_num).description 			= "Sends a SVES/QURY."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Communication", "Deductions", "EMER", "Adult Cash", "Health Care", "HS/GRH", "Income", "LTC", "MFIP", "Reviews", "SNAP", "Utility", "DWP")
script_array(script_num).dlg_keys               = array("Cn", "Up")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Shelter Expense Verif Received"
' script_array(script_num).description 			= "Enter shelter expense/address information in a dialog and the script updates SHEL, HEST, and ADDR and case notes."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Communication", "Deductions", "DWP", "EMER", "HS/GRH", "LTC", "MFIP", "Reviews", "SNAP", "Utility")
script_array(script_num).dlg_keys               = array("Cn", "Up")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Significant Change"
' script_array(script_num).description 			= "Template for noting Significant Change information."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "Income", "MFIP", "Reviews")
script_array(script_num).dlg_keys               = array("Cn", "Tk", "Sm")
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "SMRT"
' script_array(script_num).description 			= "Template for case noting the SMRT process and information."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "Communication", "Health Care", "MFIP", "Reviews")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #01/19/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "SNAP E and T Letter"
' script_array(script_num).description 			= "Sends a SPEC/LETR informing client that they have an Employment and Training appointment."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("ABAWD", "Application", "Communication", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Sm", "Up")
script_array(script_num).subcategory            = array("SNAP")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Transfer Case"
' script_array(script_num).description 			= "SPEC/XFERs a case, and can send a client memo. For in-agency as well as out-of-county XFERs."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Communication", "Reviews", "Utility", "MFIP", "DWP", "SNAP", "HS/GRH", "Health Care", "Adult Cash", "EMER")
script_array(script_num).dlg_keys               = array("Cn", "Up", "Sm")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Task Based Assistor"													'Script name
' script_array(script_num).description		    = "Script that assists in the review for identified HSRs, and outputs them into an Excel."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Ex", "Ev")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #05/06/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Task Based DAIL Capture"													'Script name
' script_array(script_num).description		    = "BULK script that captures specified DAILS for identified populations, and outputs them into a SQL Database."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Ev")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #02/11/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)   'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name 			= "TIKL FROM LIST"											'Script name
' script_array(script_num).description 			= "BULK script that gathers ABAWD/FSET codes for members on SNAP/MFIP active cases."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "Monthly Tasks", "MFIP", "DWP", "SNAP", "HS/GRH", "Health Care", "Adult Cash", "EMER")
script_array(script_num).dlg_keys               = array("Ex", "Tk")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array(script_num)	    'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	   'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Training Case Creator"
' script_array(script_num).description		    = "Creates training case scenarios en masse and XFERs them to workers."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Health Care", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Up", "Fi", "Ex")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "UC Verification Request"
' script_array(script_num).description 			= "Creates an email of a houshold member to request Unemployment Compensation"
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Income", "Applications", "Reviews", "Utility", "SNAP", "MFIP", "DWP", "Adult Cash", "HS/GRH", "Health Care", "EMER")
script_array(script_num).dlg_keys               = array("Oe")
script_array(script_num).subcategory            = array("TOOL")
script_array(script_num).release_date           = #09/16/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)	'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name			= "UNEA Updater"										'Script name
' script_array(script_num).description 			= "BULK script that updates UNEA information and sends SPEC/MEMO for VA cases at ER."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "Monthly Tasks", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Up", "Cn", "Sm", "Ex")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array(script_num)	    'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	   'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Update Check Dates"
' script_array(script_num).description		    = "Updates the dates on JOBS and UNEA to the correct dates for the footer month."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Utility")
script_array(script_num).dlg_keys               = array("Up")
script_array(script_num).subcategory            = array("TOOL", "MAXIS")
script_array(script_num).release_date           = #07/17/2020#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Update Worker Signature"
' script_array(script_num).description 			= "Sets or updates the default worker signature for this user."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Utility")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("TOOL")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1					'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "VA Verification Request"
' script_array(script_num).description 			= "Creates an email requesting VA Benefits."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Income", "Applications", "Reviews", "Utility", "SNAP", "MFIP", "DWP", "Adult Cash", "HS/GRH", "Health Care", "EMER")
script_array(script_num).dlg_keys               = array("Oe")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #11/25/2020#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Vendor"
' script_array(script_num).description 			= "Template for documenting vendor inforamtion."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "DWP", "Income", "MFIP", "Reviews", "EMER", "HS/GRH")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #09/25/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Verifications Needed"
' script_array(script_num).description 			= "Template for when verifications are needed (enters each verification clearly)."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "Assets", "Communication", "Deductions", "DWP", "EMER", "Health Care", "HS/GRH", "Income", "LTC", "MFIP", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Tk")
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Verifications Still Needed"
' script_array(script_num).description 			= "Creates a Word document informing client of a list of verifications that are still required."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Communication", "DWP", "EMER", "HS/GRH", "MFIP", "LTC", "SNAP", "Health Care")
script_array(script_num).dlg_keys               = array("Cn", "Wrd")
script_array(script_num).subcategory            = array("WORD DOCS")
script_array(script_num).release_date			= #04/25/2016#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "View PNLP"
' script_array(script_num).description 			= "Set all the panels in STAT to 'V'iew in the PNLP order."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Utility", "SNAP", "MFIP", "DWP", "Adult Cash", "HS/GRH", "Health Care", "EMER")
script_array(script_num).dlg_keys               = array("Ev")
script_array(script_num).subcategory            = array("MAXIS")
script_array(script_num).release_date           = #04/17/2019#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Waived ER Interview Screening"										'Script name
' script_array(script_num).description		    = "Evaluate a case to determine if we can waive the ER Interview."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Utility")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/14/2020#
script_array(script_num).retirement_date		= #11/12/2020#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name			= "WF1 Case Status"													'Script name
' script_array(script_num).description 			= "Updates a list of cases from Excel with current case and ABAWD status information."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("SNAP")
script_array(script_num).dlg_keys               = array("Ex", "Ev")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Work Assignment Completed"
' script_array(script_num).description		    = "Reports information and details on the completion of QI Work Assignments"
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("QI", "SNAP")
script_array(script_num).dlg_keys               = array("Ex", "Oe")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #06/05/2020#
script_array(script_num).retirement_date		= #09/13/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Work Assignment from Excel"										'Script name
' script_array(script_num).description		    = "Takes work listed on a spreadsheet and splits it into assignment excel sheets for any number of workers."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False



'DAIL SCRIPTS=====================================================================================================================================

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "ABAWD FSET Exemption Check"																		'Script name
' script_array(script_num).description 			= "A tool to walk through a screening to determine if client is an ABAWD."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("ABAWD", "Communication", "SNAP")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Affiliated Case Lookup"																		'Script name
' script_array(script_num).description 			= "Navigates to CASE/NOTE for an affiliated case DAIL message."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Navigation")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "BNDX Scrubber"																		'Script name
' script_array(script_num).description 			= "Evaluates BNDX messages for discrepancies from UNEA."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "Income", "Utility")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Catch All"																		'Script name
' script_array(script_num).description 			= "Template case note to use when a DAIL messages is processed, and is not supported by another DAIL scrubber script."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication")
script_array(script_num).dlg_keys               = array("Cn", "Tk")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #05/01/2019#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Citizenship Verified"																		'Script name
' script_array(script_num).description 			= "Notes when a data-match verifies a client's citizenship."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "COLA Review and Approve"																		'Script name
' script_array(script_num).description 			= "Script to aid in the case noting of the HC approval completed after a COLA update."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "Deductions", "Health Care", "Income", "LTC")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "COLA SVES Response"																		'Script name
' script_array(script_num).description 			= "Gather's applicable client's SSN and navigates to TPQY."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "Deductions", "Health Care", "Income", "LTC")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

' script_num = script_num + 1						'Increment by one
' ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
' Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
' script_array(script_num).script_name 			= "CS Reported New Employer"																		'Script name
' ' script_array(script_num).description 			= ""
' script_array(script_num).category               = "DAIL"
' script_array(script_num).workflows              = ""
' script_array(script_num).tags                   = array("Communication", "DWP", "Income", "MFIP", "SNAP")
' script_array(script_num).dlg_keys               = array("")
' script_array(script_num).subcategory            = array("")
' script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "CSES Scrubber"																		'Script name
' script_array(script_num).description 			= "Checks PIC (SNAP), updates retro/pro (MFIP) for CSES messages."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "DWP", "Income", "MFIP", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Up")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #08/22/2016#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "DISA Message"																		'Script name
' script_array(script_num).description 			= "Processes DAIL: disability is ending in 60 days"
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "Adult Cash", "Health Care", "HS/GRH", "MFIP", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Tk")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "ES Referral Missing"																		'Script name
' script_array(script_num).description 			= "Processes PEPR Message: ES referral date needed."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "DWP", "MFIP")
script_array(script_num).dlg_keys               = array("Cn", "Up")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #09/30/2016#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Financial Orientation Missing"																		'Script name
' script_array(script_num).description 			= "PEPR: ES Referral date needed."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "DWP", "MFIP")
script_array(script_num).dlg_keys               = array("Cn", "Up")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #09/30/2016#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "FMED Deduction"																		'Script name
' script_array(script_num).description 			= "Sends a SPEC/MEMO informing of a possible FMED deduction."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Sm")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Incarceration"																		'Script name
' script_array(script_num).description 			= "Template to use when a client is incarcerated."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication")
script_array(script_num).dlg_keys               = array("Cn", "Tk", "Up")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #05/01/2019#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "LTC Remedial Care"																		'Script name
' script_array(script_num).description 			= "Updates the remedial care deduction on a client's BILS panel."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "Deductions", "LTC")
script_array(script_num).dlg_keys               = array("Up")
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Medi Check"																		'Script name
' script_array(script_num).description 			= "Script to support PEPR Message: Member has been disabled 2 years - Refer to Medicare."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "Health Care", "LTC")
script_array(script_num).dlg_keys               = array("Cn", "Tk")
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "New Hire NDNH"																		'Script name
' script_array(script_num).description 			= "Updates JOBS/case notes new HIRE message/TIKLs for proofs."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "DWP", "Adult Cash", "Health Care", "HS/GRH", "Income", "MFIP", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Tk", "Up")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "New Hire"																		'Script name
' script_array(script_num).description 			= "Updates JOBS/case notes new HIRE message/TIKLs for proofs."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "DWP", "Adult Cash", "Health Care", "HS/GRH", "Income", "MFIP", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Tk", "Up")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Overdue Baby"																		'Script name
' script_array(script_num).description 			= "Sends a MEMO informing client that they need to report information regarding the birth of their child, and/or pregnancy end date, within 10 days or their case may close."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "DWP", "Adult Cash", "Health Care", "MFIP", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Sm", "Tk")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #05/01/2019#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Paperless Dail"																		'Script name
' script_array(script_num).description 			= "Makes an approval case note for HC and LTC cases based on a DAIL scrubber message generated through the BULK - PAPERLESS IR script."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Health Care", "LTC", "Reviews")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #12/01/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Postponed Expedited SNAP Verifications"																		'Script name
' script_array(script_num).description 			= "Case notes verifications still needed for EXP SNAP closure due to postponed verifications not received."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Communication", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Exp")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "SDX Info Has Been Stored"																		'Script name
' script_array(script_num).description 			= "Jumps to SDXS for a related SDX info DAIL message."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Health Care", "HS/GRH", "Income", "LTC", "SNAP")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "TPQY Response"																		'Script name
' script_array(script_num).description 			= "Jumps to SVES/TPQY for the case which has received a response."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "Navigation", "Utility")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Wage Match Scrubber"																		'Script name
' script_array(script_num).description 			= "Script grabs quarterly earnings information from the match as well as earned income information"
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "DWP", "Adult Cash", "Health Care", "HS/GRH", "Income", "MFIP", "SNAP")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/04/2016#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

'NAV SCRIPTS=====================================================================================================================================

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "CASE-CURR"
' script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Navigation")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "CASE-PERS"
' script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Navigation")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #06/08/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "CASE-NOTE"
' script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Navigation")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Find MAXIS case in MMIS"
' script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Navigation")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Find MMIS PMI in MAXIS"
' script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Navigation")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "MMIS - GRH"
' script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Navigation")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "POLI-TEMP"
' script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Navigation")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "STAT-ADDR"
' script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Navigation")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "STAT-MEMB"
' script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Navigation")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "View PNLP"
' script_array(script_num).description 			= "Navigates to and sets all the panels in STAT to 'V'iew in the PNLP order."
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Navigation", "Utility")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False

' for test_thing = 0 to UBound(script_array)
' 	MsgBox script_array(test_thing).description
' next
