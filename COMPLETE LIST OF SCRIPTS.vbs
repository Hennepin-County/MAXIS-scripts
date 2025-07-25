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
	public usage_eval
	public script_checkbox_one
	public script_btn_one
	public used_for_elig
	public policy_references		'ARRAY of all policy references the script supports - use spaces between the 3 items the name should have underscores in place of spaces'
									'CM Section_Name XX.XX.XX
									'TE Section_Name XX.XX.XX
									'SHAREPOINT Site_Name URL
									'SIR Ref_Name URL
									'ONESOURCE Section_Name URL
									'EPM Section_Name URL
									'BULLETIN Section_Name URL
	public specialty_redirect
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

		url_category = category
		If url_category = "CA" Then url_category = "case-assignment"
		If url_category = "MHC" Then url_category = "managed-care"
		If run_locally = true then
			script_repository = "C:\MAXIS-Scripts\"
            If header_exists = TRUE Then
                script_URL = script_repository & lcase(url_category) & "\" & lcase(header) & "-" & lcase(replace(name_to_use, " ", "-") & ".vbs")
            Else
			    script_URL = script_repository & lcase(url_category) & "\" & lcase(replace(script_name, " ", "-") & ".vbs")
            End If
		Else
        	If script_repository = "" then
				script_repository = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/"    'Assumes we're scriptwriters
				IF on_the_desert_island = TRUE Then script_repository = desert_island_respository
			End If

            If header_exists = TRUE Then
                script_URL = script_repository & lcase(url_category) & "/" & lcase(header) & "-" & replace(lcase(name_to_use) & ".vbs", " ", "-")
            Else
                script_URL = script_repository & lcase(url_category) & "/" & replace(lcase(script_name) & ".vbs", " ", "-")
            End If
		End if
    end property

    public property get SharePoint_instructions_URL 'The instructions URL in SIR
        ' SharePoint_instructions_URL = "https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/Script%20Instructions%20Wiki/" & replace(ucase(script_name) & ".aspx", " ", "%20")
		' "https://dept.hennepin.us/hsphd/sa/ews/BlueZone_Script_Instructions/"			'OLD URL
        ' "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/"		'NEW URL'
		If script_name = "Add WCOM" Then
			SharePoint_instructions_URL = "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTICES/NOTICES%20-%20ADD%20WCOM.docx"
		ElseIf left(script_name, 4) = "REPT" OR script_name = "DAIL Report" OR script_name = "EMPS" OR script_name = "LTC-GRH List Generator" Then
			SharePoint_instructions_URL = "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/BULK/BULK%20-%20REPT%20LISTS.docx"
		ElseIf category = "CA" Then
            SharePoint_instructions_URL = "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/CASE%20ASSIGNMENT/" & UCase(category) & "%20-%20" & replace(ucase(script_name) & ".docx", " ", "%20")
		ElseIf category = "MHC" Then
            SharePoint_instructions_URL = "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/MMIS/" & UCase(category) & "%20-%20" & replace(ucase(script_name) & ".docx", " ", "%20")
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
		If script_name = "Delete DAIL Tasks" Then description = description & " --- USE WITH CAUTION! Deletes info from SQL Database."
		If script_name = "Open Interview PDF" Then description = description & " --- Opens a PDF generated from NOTES - Interview if not yet in ECF."
		If script_name = "Search CASE NOTE" Then description = description & " --- Searches all CASE:NOTEs for a particular case for word(s) or a phrase."
		If script_name = "Hot Topics" Then description = description & " --- Displays a list of BlueZone Script related Hot Topics with links to the articles and related scripts."
		If script_name = "XML File Cleanup" Then description = description & " --- Archives and removes aged MNBenefits xml files. Restricted to authorized users only."
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
' script_array(script_num).keywords               = array("")
' script_array(script_num).release_date           = #10/01/2000#
' script_array(script_num).hot_topic_link			= ""
' script_array(script_num).used_for_elig			= False
' script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'

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
script_array(script_num).hot_topic_date 		= #04/19/2022#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/sitepages/Maxis-IEVS-DAIL-Messages-Changes.aspx"
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "2 PM Return"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #06/20/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "311"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #06/20/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).retirement_date        = #02/28/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("CM LENGTH_OF_RECERTIFICATION_PERIODS 09.03", "TE SNAP_AGED/DISABLED_12_MONTH_CONTACTS 02.08.165", "SHAREPOINT SNAP_Recertification https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/SNAP_Recertification.aspx")
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).retirement_date        = #01/02/2024#
script_array(script_num).hot_topic_date         = #06/20/2023#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/TLRs.aspx"
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("CM Time-Limited_SNAP_Recipient 11.24", "CM Who_Is_Exempt_From_SNAP_Work_Registration 28.06.12")					'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).retirement_date		= ""
script_array(script_num).hot_topic_date         = #06/20/2023#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/TLRs.aspx"
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("CM Time-Limited_SNAP_Recipient 11.24", "CM Who_Is_Exempt_From_SNAP_Work_Registration 28.06.12")					'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STAR"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).retirement_date		= #01/12/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).retirement_date		= ""
script_array(script_num).hot_topic_date         = #06/20/2023#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/TLRs.aspx"
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "LIMITED"

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
script_array(script_num).retirement_date		= #06/19/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Accounting Refund"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Utility", "Communication", "MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #06/20/2017#
script_array(script_num).retirement_date		= #09/15/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "ACF Request Pend"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #06/20/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "ACF Used"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #06/20/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STAR"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Address Report"																		'Script name
' script_array(script_num).description 			= ""
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Monthly Tasks", "Utility", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #02/22/2024#
script_array(script_num).retirement_date        = #03/16/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "ADH Info and Hearing"																		'Script name
script_array(script_num).category               = "DEU"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("SNAP", "MFIP", "DWP", "HS/GRH", "Adult Cash", "EMER", "Health Care")
script_array(script_num).dlg_keys               = array("Cn", "Oe")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #07/07/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Adoption Assistance"																		'Script name
script_array(script_num).category               = "IV-E"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #11/25/2019#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "All Scripts"																		'Script name
' script_array(script_num).description 			= "Template for documenting details about an appeal, and the appeal process."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Support")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("TOP", "TOOL")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #08/09/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).retirement_date        = #06/03/2024#
' script_array(script_num).hot_topic_date			= #04/18/2023#
' script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/sitepages/power-pad-and-health-care-script-updates.aspx"
'script_array(script_num).hot_topic_date			= #11/15/2022#
'script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/sitepages/Notes-%e2%80%93-Application-Check-Updated-to-Support-Day-30-Processing.aspx"
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name			= "Application Inquiry"													'Script name
' script_array(script_num).description 			= "Sends an Email request to search for an ApplyMN that is not in ECF."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Support", "Utility")
script_array(script_num).dlg_keys               = array("Oe")
script_array(script_num).subcategory            = array("TOP", "REQUESTS")
script_array(script_num).release_date           = #04/01/2022#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

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
script_array(script_num).policy_references		= array("CM Emergency_Aid_Eligibility_-_SNAP/Expedited_Food 04.04","CM Applications 05","CM Application_-_Pending_Cases 05.09.12","TE Notice_of_Interview/Missed_Interview_(NOMI) 02.05.15", "SHAREPOINT Applications https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Applications.aspx")						'SEE Line 58 for format')						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "PRIMARY"
script_array(script_num).specialty_redirect		= "CA"

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
script_array(script_num).retirement_date        = #06/27/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= True
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "ATR Received"																		'Script name
script_array(script_num).category               = "DEU"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("SNAP", "MFIP", "DWP", "HS/GRH", "Adult Cash", "EMER", "Health Care")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #11/07/2017#
script_array(script_num).retirement_date        = #03/29/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)   'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name 			= "Auto-Dialer Case Status"											'Script name
' script_array(script_num).description 			= "BULK script that gathers case status for cases with recerts for SNAP/MFIP the previous month."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "SNAP", "MFIP")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).retirement_date        = #05/23/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "AVS"													'Script name
' script_array(script_num).description		    = "Supports for AVS forms and AVS Submission/Results processes for MA-ABD Cases."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Health Care", "Reviews")
script_array(script_num).dlg_keys               = array("Cn", "Tk", "Wrd")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #04/27/2021#
script_array(script_num).hot_topic_date			= #07/06/2021#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/sitepages/New-Script--NOTES---AVS.aspx"
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("SHAREPOINT AVS https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/AVS.aspx", "ONESOURCE Account_Validation_Service_(AVS) https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=onesource-17031", "EPM 2.3.1.2_MA-ABD_Account_Validation_Service_(AVS) http://hcopub.dhs.state.mn.us/epm/2_3_1_2.htm?rhhlterm=avs&rhsearch=avs")
script_array(script_num).usage_eval				= "PRIMARY"

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
script_array(script_num).retirement_date        = #12/29/2022#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

' script_num = script_num + 1							   'Increment by one
' ReDim Preserve script_array(script_num)	    'Resets the array to add one more element to it
' Set script_array(script_num) = new script_bowie	   'Set this array element to be a new script. Script details below...
' script_array(script_num).script_name		    = "Basket Review"
' script_array(script_num).file_name			= "basket-review.vbs"
' ' script_array(script_num).description		    = "A script that creates a report of cases and pages pending on a list of baskets."


script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Banked Months Updater"	'Script name
' script_array(script_num).description 			= ""
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("ABAWD", "Application", "Communication", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Ev", "Sm", "Cn")
script_array(script_num).subcategory            = array("ABAWD")
script_array(script_num).release_date           = #09/19/2023#
script_array(script_num).retirement_date		= "05/08/2025"
script_array(script_num).hot_topic_date         = ""
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("TE TLR_Banked_Months_-_System_Process 10.34.01", "TE Coding_the_WREG_Panel_for_SNAP 02.05.70", "BULLETIN #23-01-02_SNAP_Banked_Months_for_Time_Limited_Recipients_(TLR's) https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_FILE&RevisionSelectionMethod=LatestReleased&Rendition=Primary&allowInterrupt=1&noSaveAs=1&dDocName=mndhs-063946")
script_array(script_num).usage_eval				= "STAR"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STAR"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "BULK Match Cleared"																		'Script name
script_array(script_num).category               = "DEU"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("SNAP", "MFIP", "DWP", "HS/GRH", "Adult Cash", "EMER", "Health Care")
script_array(script_num).dlg_keys               = array("Cn", "Ex")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #03/20/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Bulk POLI TEMP to Word"
' script_array(script_num).description 			= "Creates a list of current POLI/TEMP topics, TEMP reference and revised date."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("QI", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Wrd", "Ex")
script_array(script_num).subcategory            = array()
script_array(script_num).release_date           = #10/04/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Bus Ticket Issued"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #08/01/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).hot_topic_date			= #07/11/2023#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/sitepages/NOTES%20%E2%80%93%20CAF%20Updates%20and%20Guidance.aspx"
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "PRIMARY"

script_num = script_num + 1					'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Calculate Rate 2 Units"
' script_array(script_num).description 			= "Calculates the GRH Rate 2 total units to input into MMIS."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Calculators", "Utility", "HS/GRH")
script_array(script_num).dlg_keys               = array("Ev")
script_array(script_num).subcategory            = array("TOP", "TOOL")
script_array(script_num).release_date           = #08/10/2018#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "LIMITED"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Case Sampling"																		'Script name
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("QI", "Utility", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #03/10/2025#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "LIMITED"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Case Transfer"																		'Script name
' script_array(script_num).description 			= "Searches caseload(s) by selected parameters. Transfers a specified number of those cases to another worker. Creates list of these cases."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Monthly Tasks", "Utility", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Ex", "Up", "Cn", "Sm")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "CES Screening Appt"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #09/23/2017#
script_array(script_num).retirement_date		= #03/12/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("TE COMPLETING_AN_INTER-AGENCY_CASE_TRANSFER 02.08.133", "SHAREPOINT Transfer_to_Another_County https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/To_Another_County.aspx")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).retirement_date        = #10/01/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("CM SNAP_Electronic_Disqualified_Recipient_System 25.24.08", "TE CASE_NOTE_I___INTRO/HH_COMP 02.08.093")
script_array(script_num).usage_eval				= "LIMITED"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).policy_references      = array("CM Mandatory_Verifications 10.18", "CM Citizenship_and_Immigration_Status 11.03", "TE Citizenship_&_Immig_Ver._For_MA_APPL 02.08.166", "SHAREPOINT Acceptable_Verifications https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Acceptable_Verification.aspx")                   'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("TE CLAIM_REFERRAL_TRACKING 02.09.47")
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).hot_topic_date         = #01/30/2024#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/sitepages/SNAP-Waived-Interview-Now-Handles-Return-Contacts.aspx?"
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("TE CASE_NOTE_I___INTRO/HH_COMP 02.08.093", "SHAREPOINT CASE_NOTE_GUIDELINES_AND_FORMAT https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Casenote_Guidelines_and_Format.aspx")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "PRIMARY"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Client Sheltered By Window A"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #09/23/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)	'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name			= "Close MMIS Rate 2 in MMIS"													'Script name
' script_array(script_num).description 			= "Script to assist in closing SSR agreements in MMIS for GRH Rate 2 cases."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "HS/GRH")
script_array(script_num).dlg_keys               = array("Ex", "Up", "Ev")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).retirement_date        = #06/27/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= True
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).retirement_date        = #08/13/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).retirement_date		= #3/2/2022#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array(script_num)	    'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	   'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Contact Knowledge Now"
' script_array(script_num).description		    = "Used to submit an email to the QI Knowledge Now team."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Support", "Utility", "SNAP", "MFIP", "DWP", "Adult Cash", "HS/GRH", "Health Care", "EMER")
script_array(script_num).dlg_keys               = array("Oe")
script_array(script_num).subcategory            = array("TOP", "REQUESTS", "POLICY")
script_array(script_num).release_date           = #08/19/2020#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array(script_num)	    'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	   'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Copy Case Note to Word"
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Support", "Utility", "SNAP", "MFIP", "DWP", "Adult Cash", "HS/GRH", "Health Care", "EMER")
script_array(script_num).dlg_keys               = array("Wrd")
script_array(script_num).subcategory            = array("TOP", "TOOL", "MAXIS")
script_array(script_num).release_date           = #09/12/2022#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array(script_num)	    'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	   'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Copy Panels to Word"
' script_array(script_num).description		    = "Copies MAXIS panels to Word en masse for a case for easier review."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Support", "Utility", "SNAP", "MFIP", "DWP", "Adult Cash", "HS/GRH", "Health Care", "EMER")
script_array(script_num).dlg_keys               = array("Wrd")
script_array(script_num).subcategory            = array("TOOL", "MAXIS")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

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
script_array(script_num).retirement_date		= #01/12/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "PRIMARY"

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)   'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name 			= "DAIL 12 Month Contact"											'Script name
' script_array(script_num).description 			= "BULK script that gathers ABAWD/FSET codes for members on SNAP/MFIP active cases."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Monthly Tasks", "SNAP")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #11/02/2021#
script_array(script_num).retirement_date        = ""
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("CM LENGTH_OF_RECERTIFICATION_PERIODS 09.03", "TE SNAP_AGED/DISABLED_12_MONTH_CONTACTS 02.08.165", "SHAREPOINT SNAP_Recertification https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/SNAP_Recertification.aspx")
script_array(script_num).usage_eval				= "STATS"

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "DAIL CCD"													'Script name
' script_array(script_num).description		    = "BULK script that captures, case notes and deletes specific DAILS based on content, and collects them into an Excel spreadsheet."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array(" DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "Monthly Tasks", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Ex", "Up", "Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "DAIL FMED Deduction"													'Script name
' script_array(script_num).description		    = ""
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Monthly Tasks", "SNAP", "Communication")
script_array(script_num).dlg_keys               = array("Ex", "Sm", "Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #09/07/2022#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "DAIL Unclear Information"
' script_array(script_num).description 			= "Evaluates HIRE and CSES messages in the DAIL to remove messages that fall under unclear information."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("SNAP")
script_array(script_num).dlg_keys               = array("Ex", "Ev", "Ex", "Up")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #08/06/2024#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/HIRE-CSES-dails.aspx"
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("CM Six-month_Reporting 07.03.02", "CM Income_of_Minor_Child/Caregiver_Under_20 17.15.15")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Data Access Test"
' script_array(script_num).description 			= "Pulls a list of DAILS in DAIL/DAIL into an Excel spreadsheet."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Utility")
script_array(script_num).dlg_keys               = array("Ev")
script_array(script_num).subcategory            = array("MAXIS")
script_array(script_num).release_date           = #06/05/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

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
script_array(script_num).retirement_date		= #12/05/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

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
script_array(script_num).retirement_date        = #06/27/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= True
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array(script_num)	    'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	   'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Display Benefits"
' script_array(script_num).description		    = "Case note to help with replacing food destroyed in a disaster"
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Support", "Utility", "SNAP", "MFIP", "DWP", "HS/GRH", "Adult Cash")
script_array(script_num).dlg_keys               = array("Ev")
script_array(script_num).subcategory            = array("TOOL", "MAXIS")
script_array(script_num).release_date           = #11/17/2020#
script_array(script_num).hot_topic_date			= #12/06/2022#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/sitepages/New-Script-%e2%80%93-Display-Benefits.aspx"
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Diversion Program Referral"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #07/05/2018#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Diversion Program Referral Results"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #07/05/2018#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "PRIMARY"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "EA Approved"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "EMER")
script_array(script_num).dlg_keys               = array("Cn", "Sm")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #07/05/2018#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "EA Extension"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #07/05/2018#
script_array(script_num).retirement_date 		= #02/22/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).hot_topic_date			= #08/26/2023#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/sitepages/Updates-to-ACTIONS-%E2%80%93-Earned-Income-Budgeting.aspx"
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("CM Determining_Gross_Income 17", "TE Student_Income 02.08.087", "TE AmeriCorps 02.05.34", "TE SNAP_Benefits_-_2nd_30_Days_Applicant_Delay 02.05.107", "TE CASE_NOTE_II___Assets_Income 02.08.094", "TE Determining_MFIP_Initial/Ongoing_Eligibility 02.05.80", "TE Budget_Cycle_Monthly_Reporting_Changes 02.08.032", "SHAREPOINT Income https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Income.aspx", "SHAREPOINT Earned_Income https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Earned_Income.aspx")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STAR"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "EBT Benefits Stolen"													'Script name
' script_array(script_num).description		    = "Documents benefits replacement request for stolen benefits."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "SNAP", "MFAP", "MFIP", "DWP")
script_array(script_num).dlg_keys               = array("Cn", "Tk", "Sm")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #06/18/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("TE EBT_Stolen_Benefits_-_Client_Notification 02.11.126", "TE EBT_Stolen_Benefits_-_Client_Reports 02.11.127", "TE EBT_Stolen_Benefits_-_Case_Note 02.11.128")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("CM SNAP_Electronic_Disqualified_Recipient_System 25.24.08", "TE CASE_NOTE_I___INTRO/HH_COMP 02.08.093")
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "ELIG Results to Word"
' script_array(script_num).description 			= "Creates a Word Document of a single POLI/TEMP reference, need the Table Number."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Support", "Utility", "SNAP", "MFIP", "DWP", "Adult Cash", "HS/GRH", "EMER")
script_array(script_num).dlg_keys               = array("Wrd")
script_array(script_num).subcategory            = array("TOOL", "MAXIS")
script_array(script_num).release_date           = #03/27/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Eligibility Summary"																		'Script name
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "Communication", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Sm", "Exp")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #07/26/2022#
script_array(script_num).hot_topic_date			= #07/11/2023#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/sitepages/NOTES%20%E2%80%93%20CAF%20Updates%20and%20Guidance.aspx"
script_array(script_num).used_for_elig			= True
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "PRIMARY"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Elig Review"																		'Script name
script_array(script_num).category               = "IV-E"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #11/25/2019#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "PRIMARY"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Emergency Exceeds Approval Limit"
' script_array(script_num).description 			= "Template for EA/EGA applications.*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/09/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Enroll in Script Demo"
' script_array(script_num).description 			= "Script to display any upcoming BlueZone Script Demos that are scheduled and allow you to sign up and add to your Outlook Schedule."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Support", "Utility")
script_array(script_num).dlg_keys               = array("Oe", "Oa")
script_array(script_num).subcategory            = array("REQUESTS", "POLICY")
script_array(script_num).release_date           = #08/10/2020#
script_array(script_num).retirement_date		= #01/12/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Enrollment"
script_array(script_num).category               = "MHC"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Health Care")
script_array(script_num).dlg_keys               = array("Cn", "Up")
script_array(script_num).subcategory            = array()
script_array(script_num).release_date           = #04/02/2019#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "PRIMARY"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Enrollment Note"
script_array(script_num).category               = "MHC"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Health Care")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array()
script_array(script_num).release_date           = #04/28/2018#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)   'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name 			= "Ex Parte Report"											'Script name
' script_array(script_num).description 			= "BULK script that gathers reivew (ER, SR, etc.) cases and information."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "Health Care")
script_array(script_num).dlg_keys               = array("Ex", "Ev", "Cn", "Up")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #05/25/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Enter Date of Death"																		'Script name
' script_array(script_num).description 			= "Update MAXIS panels with date of death for household member."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "DWP", "EMER", "Health Care", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Up")
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #12/05/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/New-Scripts-Available-September-1-for-Interview-and-Expedited-SNAP.aspx"
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("CM Emergency_Aid_Eligibility_-_SNAP/Expedited_Food 04.04", "TE Expedited_SNAP_W/_Postponed_Verifs 02.10.01", "TE WREG_Expedited_SNAP_Postponed_Verifs 02.05.70.01", "TE Expedited_FS_2nd_Month_Eligibility 02.10.79", "TE Asset_Coding_for_Expedited_SNAP 02.08.185", "SHAREPOINT Acceptable_Verifications https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Acceptable_Verification.aspx")
script_array(script_num).usage_eval				= "LIMITED"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Expedited Determination Report"
' script_array(script_num).description 			= "Template for noting detail about how expedited was determined for a case."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "BZ", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Exp", "Ev")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).retirement_date		= #07/20/2022#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("CM Emergency_Aid_Eligibility_-_SNAP/Expedited_Food 04.04")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "LIMITED"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STAR"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Find Hidden Excel"
' script_array(script_num).description 			= "Creates a list of cases and clients active on MA-EPD and Medicare Part B that are eligible for Part B reimbursement."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Utility")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #05/26/2022#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Find Hidden Excel"
' script_array(script_num).description 			= "Creates a list of cases and clients active on MA-EPD and Medicare Part B that are eligible for Part B reimbursement."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Utility")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("TOP", "TOOL")
script_array(script_num).release_date           = #05/26/2022#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Find MEMB in MMIS"
' script_array(script_num).description 			= ""
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Navigation", "Utility")
script_array(script_num).dlg_keys               = array("Health Care")
script_array(script_num).subcategory            = array("MAXIS", "TOOL")
script_array(script_num).release_date           = #08/19/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).retirement_date        = #02/28/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).retirement_date        = #4/15/2025#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).retirement_date        = #10/18/2022#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).retirement_date        = #05/11/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Health Care Budget Report"
' script_array(script_num).description 			= "Pulls a list of active SNAP/MFIP cases with identified info into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Reports", "Health Care")
script_array(script_num).dlg_keys               = array("Ex", "Ev")
script_array(script_num).subcategory            = array("ENHANCED LISTS")
script_array(script_num).release_date           = #01/31/2023#
script_array(script_num).retirement_date        = #05/23/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Health Care Evaluation"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Assets", "Application", "Deductions", "Health Care", "Income")
script_array(script_num).dlg_keys               = array("Cn", "Tk")
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #4/10/2023#
script_array(script_num).hot_topic_date         = #04/18/2023#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/sitepages/A-New-Script-Experience-for-Health-Care.aspx"
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "PRIMARY"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Health Care Pending Assignments"
' script_array(script_num).description 			= "Template for the METS to MAXIS and MAXIS to METS transition process."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Health Care")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("TOP", "TOOL")
script_array(script_num).release_date           = #04/10/2025#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Health Care Pending Assignments"
' script_array(script_num).description 			= "Template for the METS to MAXIS and MAXIS to METS transition process."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Health Care")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("TOP", "TOOL")
script_array(script_num).release_date           = #04/10/2025#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("SHAREPOINT METS_to_MAXIS_Transitions https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/METS_to_MAXIS_Transitions.aspx", "SHAREPOINT Use_MA_Transition_Communication_form_for_METS_to_MAXIS_Transition https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/Reminder--METS-to-MAXIS-Transition.aspx")
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Homelessness Verified"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #07/05/2018#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Hot Topics"
' script_array(script_num).description 			= "Update STAT/MEDI with MBI number from RMCR in MMIS."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Navigation", "Utility")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("TOP", "POLICY")
script_array(script_num).release_date           = #05/08/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "PRIMARY"

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
script_array(script_num).retirement_date        = #05/11/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Immigration Status"
' script_array(script_num).description 			= "Template for the SAVE system for verifying immigration status."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "Communication", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Up", "Oe", "Oa")
script_array(script_num).subcategory            = array("G-L")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)	'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name			= "Inactive Transfer"													'Script name
' script_array(script_num).description 			= "Script to transfer inactive cases via SPEC/XFER"
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "Monthly Tasks", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Up", "Ex", "Ev")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).retirement_date        = #01/23/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).retirement_date        = #03/06/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).retirement_date        = #03/06/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)	'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name			= "Individual Recertification Notices"													'Script name
' script_array(script_num).description 			= "NOTICES Script that will send ODW Recert Appointment Letter or NOMI on a single case."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "SNAP", "MFIP")
script_array(script_num).dlg_keys               = array("Cn", "Sm")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).retirement_date        = #03/06/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)	'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name			= "Individual On Demand Notices"													'Script name
' script_array(script_num).description 			= "NOTICES Script that will send ODW Recert Appointment Letter or NOMI on a single case."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "QI", "SNAP", "MFIP", "DWP", "Adult Cash", "HS/GRH")
script_array(script_num).dlg_keys               = array("Cn", "Sm")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #03/06/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "LIMITED"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

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
script_array(script_num).hot_topic_date         = #12/19/2023#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/sitepages/NOTES-%e2%80%93-Client-Contact-is-Changing.aspx?"
script_array(script_num).retirement_date        = ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "PRIMARY"

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
script_array(script_num).retirement_date        = #10/1/2021#
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Interview Team Cases Worklist"
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "DWP", "EMER", "HS/GRH", "MFIP", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #01/14/2025#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "IV-E"																		'Script name
script_array(script_num).category               = "IV-E"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #11/25/2019#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STAR"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).retirement_date		= #04/01/2022#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).retirement_date        = #05/11/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "LTC - Asset Assessment"
' script_array(script_num).description 			= "Template for the LTC asset assessment. Will enter both person and case notes if desired."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Assets", "LTC")
script_array(script_num).dlg_keys               = array("Cn", "Sm")
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).retirement_date        = #10/16/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).retirement_date        = #10/01/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).retirement_date        = #11/07/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).retirement_date        = #10/16/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).retirement_date        = #10/18/2022#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("EPM Lookback_Period_and_Transfer_Date http://hcopub.dhs.state.mn.us/epm/2_4_1_3_1.htm?rhhlterm=baseline%20date&rhsearch=baseline%20date", "EPM Transfer_Penalty http://hcopub.dhs.state.mn.us/epm/2_4_1_3_2.htm?rhhlterm=ltc&rhsearch=LTC", "TE Uncompensated_Asset_Income_Transfers_-_MA 02.14.27")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("EPM Community_Spouse_Asset_Allowance http://hcopub.dhs.state.mn.us/epm/2_4_2_1_2.htm?rhhlterm=12%20community%20spouse&rhsearch=12%20months%20community%20spouse", "ONESOURCE Processing_an_Asset_or_Income_Transfer_for_MA-LTC https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=ONESOURCE-170126")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STAR"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STAR"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STAR"

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
script_array(script_num).retirement_date		= #05/05/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STAR"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Match Cleared"																		'Script name
script_array(script_num).category               = "DEU"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("SNAP", "MFIP", "DWP", "HS/GRH", "Adult Cash", "EMER", "Health Care")
script_array(script_num).dlg_keys               = array("Cn", "Up")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #11/14/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references      = array("","","","","","","","", "")
script_array(script_num).policy_references(0)   = "CM INCOME_AND_ELIGIBILTY_VERIFICATION_SYSTEM 10.24"
script_array(script_num).policy_references(1)   = "TE IEVS_WAGE_MATCH___EARNER_DISCREPANCY 02.12.10"
script_array(script_num).policy_references(2)   = "TE QTIP_#64_IVES_MATCH_USING_BC_CLOSED_CASE 19.164"
script_array(script_num).policy_references(4)   = "TE IEVS_DAIL_MESSAGES 02.08.083"
script_array(script_num).policy_references(5)   = "TE ACCESSING_INFORMATION_ABOUT_IEVS_MATCHES 02.08.084"
script_array(script_num).policy_references(6)   = "SHAREPOINT IEVS_MATCHES https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/IEVS_Matches.aspx"
script_array(script_num).policy_references(7)   = "SHAREPOINT IEVS https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/IEVS.aspx"
script_array(script_num).policy_references(8)   = "SHAREPOINT TYPES_OF_IEVS_MATCHES https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Types_of_Matches.aspx"
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Mandatory Vendor Approved"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Communication", "MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Cn", "Sm")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #07/05/2018#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Mandatory Vendor MEMO"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Communication", "MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Sm")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #07/05/2018#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("TE HC_-_RETRO 09.18", "TE FIAT_-_HC 09.17.02","SHAREPOINT Retroactive_Medical_Assistance https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Retroactive_MA.aspx", "SHAREPOINT MAXIS_and_MMIS_Action_for_Retroactive_MA https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/MAXIS_and_MMIS_Action_for_Retroactive_MA.aspx", "SHAREPOINT Retro_Flow_Chart https://hennepin.sharepoint.com/teams/hs-es-manual/MNsure%20Documents/Forms/AllItems.aspx?id=%2Fteams%2Fhs%2Des%2Dmanual%2FMNsure%20Documents%2FRetro%20flow%20chart%2Epdf&parent=%2Fteams%2Fhs%2Des%2Dmanual%2FMNsure%20Documents", "EPM 1.2.5_MHCP_Retroactive_Eligibility http://hcopub.dhs.state.mn.us/epm/1_2_5.htm?rhhlterm=retroactive&rhsearch=retroactve")
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "MFIP Orientation"
' script_array(script_num).description 			= "Template and email support for when METS retro coverage has been requested."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "MFIP")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #09/10/2022#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "LIMITED"

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
script_array(script_num).retirement_date        = #02/28/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STAR"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "MHC Client Contact"
script_array(script_num).category               = "MHC"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Health Care")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array()
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "MIPPA"
' script_array(script_num).description 			= "Template for noting Medical Service Questionaires (MSQ)."
script_array(script_num).category               = "CA"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Health Care")
script_array(script_num).dlg_keys               = array("Cn", "Tk", "Up")
script_array(script_num).subcategory            = array()
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "MONT Report"
' script_array(script_num).description 			= "Template for noting Medical Service Questionaires (MSQ)."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "MFIP", "Adult Cash", "Income", "Monthly Tasks")
script_array(script_num).dlg_keys               = array("Cn", "Ex")
script_array(script_num).subcategory            = array()
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).retirement_date		= #01/12/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "New Job Reported"
' script_array(script_num).description 			= "Creates a JOBS panel, CASE/NOTE and TIKL when a new job is reported. Use the DAIL scrubber for new hire DAILs."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "Communication", "DWP", "EMER", "Health Care", "HS/GRH", "Income", "MFIP", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Tk", "Up")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).retirement_date		= #09/25/2020#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Non IV-E Medical"																		'Script name
script_array(script_num).category               = "IV-E"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #11/25/2019#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Northstar Kinship Assistance"																		'Script name
script_array(script_num).category               = "IV-E"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("")
script_array(script_num).dlg_keys               = array("Cn", "Tk")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #09/11/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("EPM Medical_Assistance_for_Children_Receiving_Northstar_Kinship_Assistance_(MA-NKA) https://hcopub.dhs.state.mn.us/epm/2_5_6_2.htm")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "NSPOW Checked"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #07/05/2018#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "On Demand Dashboard"
' script_array(script_num).description 			= "Creates some case notes and assists with emails on reoccuring process issues."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("QI", "SNAP", "MFIP", "DWP", "Adult Cash", "HS/GRH")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #12/21/2022#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		 'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	 'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "On Demand Waiver Applications"									'Script name
' script_array(script_num).description		    = "BULK script to collect information for cases that require an interview for the On Demand Waiver."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "SNAP", "MFIP", "DWP", "Adult Cash", "HS/GRH")
script_array(script_num).dlg_keys               = array("Ex", "Ev", "Sm", "Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/New-Scripts-Available-September-1-for-Interview-and-Expedited-SNAP.aspx"
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).retirement_date		= #10/25/2022#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1                     'Increment by one
ReDim Preserve script_array(script_num)         'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie 'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name            = "Overpayment"
' script_array(script_num).description          = "Template for noting basic information about overpayments."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Communication", "DWP", "EMER", "Health Care", "HS/GRH", "Income", "LTC", "MFIP", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Oe", "Up")
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link         = ""
script_array(script_num).used_for_elig          = False
script_array(script_num).policy_references      = array("","","","","","","","","","","","","","","")
script_array(script_num).policy_references(0)       = "CM Benefit_Adjustment_and_Recovery 25"
script_array(script_num).policy_references(1)       = "TE Claim_Entry_Initiates_Transaction 02.09.07"
script_array(script_num).policy_references(2)       = "TE CLAIM_REFERRAL_TRACKING 02.09.47"
script_array(script_num).policy_references(3)       = "TE CASE_AND_PERSON_CLRA_PANELS 02.09.02"
script_array(script_num).policy_references(4)       = "TE CORRECT_MISTAKES_ON_CLAIM_ALREADY_ENTERED 02.09.05"
script_array(script_num).policy_references(5)       = "TE DEMAND_LETTERS_FOR_OVERPAYMENTS 02.09.00"
script_array(script_num).policy_references(6)       = "TE CASE_NOTE_III:___CLAIMS/SYSTEMS/TRANSFERS 02.08.095"
script_array(script_num).policy_references(7)       = "TE MCE 02.09.41"
script_array(script_num).policy_references(8)       = "TE MCE_PAYMENTS_AND_CONTACT_INFORMATION 02.09.41.03"
script_array(script_num).policy_references(9)       = "SHAREPOINT CLAIMS_AND_UNDERPAYMENTS_POLICY https://hennepin.sharepoint.com/teams/hs-es-manual/sitepages/Claims_and_Underpayments_Policy.aspx"
script_array(script_num).policy_references(10)      = "SHAREPOINT CLAIMS_AND_UNDERPAYMENTS_PROCEDURE https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Claims_and_Underpayments_Procedure.aspx"
script_array(script_num).policy_references(11)      = "SHAREPOINT CLAIM_DEMAND_LETTER https://hennepin.sharepoint.com/teams/hs-es-manual/sitepages/Claims_Demand_Letter.aspx"
script_array(script_num).policy_references(12)      = "SHAREPOINT APPEAL_&_FRAUD_RELATED_CLAIMS https://hennepin.sharepoint.com/teams/hs-es-manual/sitepages/Appeal_and_Fraud_Related_Claims.aspx"
script_array(script_num).policy_references(13)      = "SHAREPOINT UNDERPAYMENTS,_ADJUSTMENTS,_AND_CLOSING https://hennepin.sharepoint.com/teams/hs-es-manual/sitepages/Underpayments,_Adjustments_and_Closing.aspx"
script_array(script_num).policy_references(14)      = "EPM 1.3.2.5_MHCP_OVERPAYMENTS http://hcopub.dhs.state.mn.us/epm/1_3_2_5.htm?rhhlterm=overpayments%20overpayment&rhsearch=overpayments"
script_array(script_num).usage_eval				= "STANDARD"
script_array(script_num).specialty_redirect 	= "DEU"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "P Note"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #11/17/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).hot_topic_date			= #06/29/2021#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/sitepages/NOTICES-%e2%80%93-PA-Verif-Request-has-Returned!.aspx"
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)	'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name			= "Paperless IR"                                                       'Script name
' script_array(script_num).description 			= "Updates cases on a caseload(s) that require paperless IR processing. Does not approve cases."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "Health Care")
script_array(script_num).dlg_keys               = array("Ex", "Up", "Ev", "Tk")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "PARIS Match Cleared"																		'Script name
script_array(script_num).category               = "DEU"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("SNAP", "MFIP", "DWP", "HS/GRH", "Adult Cash", "EMER", "Health Care")
script_array(script_num).dlg_keys               = array("Cn", "Up")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #05/17/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references      = array("","","","","","","","")
script_array(script_num).policy_references(0)       = "CM PUBLIC_ASSISTANCE_REPORTING_INFORMATION_SYTSEM_(PARIS)_INTERSTATE_MATCH_PROGRAM 10.24.30"
script_array(script_num).policy_references(1)       = "TE ACCESSING_AND_RESOLVING_PARIS_MATCHES 02.08.182"
script_array(script_num).policy_references(2)       = "TE PARIS_DAILS_AND_ALERTS 02.08.181"
script_array(script_num).policy_references(4)       = "TE PARIS_MATCH_TIMELINE_AND_DATA_SELECTION 02.08.180"
script_array(script_num).policy_references(5)      = "SHAREPOINT Public_Assistance_Reporting_Information_System_(PARIS) https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/PARIS.aspx"
script_array(script_num).policy_references(6)       = "ONESOURCE Process_PARIS_Matches https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=ONESOURCE-170206"
script_array(script_num).policy_references(7)      = "EPM 1.4_State_Residency https://hcopub.dhs.state.mn.us/epm/1_4.htm"
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "PARIS Match CC Claim Entered"																		'Script name
script_array(script_num).category               = "DEU"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("SNAP", "MFIP", "DWP", "HS/GRH", "Adult Cash", "EMER", "Health Care")
script_array(script_num).dlg_keys               = array("Cn", "Up", "Oe")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #12/11/2017#
script_array(script_num).retirement_date        = #03/29/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Partner Calls"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #06/19/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Permanency Review"																		'Script name
script_array(script_num).category               = "IV-E"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #12/18/2019#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Permanent Housing Found"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #06/19/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Person Search"
' script_array(script_num).description 			= "Template for noting Medical Service Questionaires (MSQ)."
script_array(script_num).category               = "CA"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array()
script_array(script_num).release_date           = #10/05/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Personal Needs"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "Income", "Deductions")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #11/14/2017#
script_array(script_num).retirement_date		= #09/15/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "POLI TEMP Monthly Updates"
' script_array(script_num).description 			= "Creates a list of current POLI/TEMP topics, TEMP reference and revised date."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Wrd")
script_array(script_num).subcategory            = array()
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"


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
script_array(script_num).release_date           = #07/15/2022#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""


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
script_array(script_num).policy_references		= array("CM Mandatory_Verifications-Cash 10.18.01")
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1					'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "QI AVS request"
' script_array(script_num).description 			= "Creates an email requesting the QI team submit an AVS request."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Health Care", "Applications", "Reviews", "Utility")
script_array(script_num).dlg_keys               = array("Oe", "Cn")
script_array(script_num).subcategory            = array("TOP", "REQUESTS")
script_array(script_num).release_date           = #03/06/2020#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "FLAG"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "REIMB Shelter ACCT"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #06/19/2017#
script_array(script_num).retirement_date		= #09/15/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Returned Mail"
' script_array(script_num).description 			= "Template for noting Returned Mail information."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Communication", "DWP", "Health Care", "HS/GRH", "MFIP", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Tk", "Up")
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Repair Ex Parte Phase 2"
' script_array(script_num).description 			= "Send an idea, error report, or other kind of request to the blueZone Script Team"
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #07/07/2023#
script_array(script_num).retirement_date		= #08/03/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Report a Duplicate UNEA Panel"
' script_array(script_num).description 			= "Send an idea, error report, or other kind of request to the blueZone Script Team"
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Support", "Utility")
script_array(script_num).dlg_keys               = array("Oe")
script_array(script_num).subcategory            = array("REQUESTS", "MAXIS", "POLICY")
script_array(script_num).release_date           = #10/25/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "LIMITED"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Report to the BZST"
' script_array(script_num).description 			= "Send an idea, error report, or other kind of request to the blueZone Script Team"
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Support", "Utility")
script_array(script_num).dlg_keys               = array("Oe")
script_array(script_num).subcategory            = array("REQUESTS", "MAXIS", "POLICY")
script_array(script_num).release_date           = #08/01/2020#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "IEVC List"																		'Script name
script_array(script_num).category               = "DEU"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("SNAP", "MFIP", "DWP", "HS/GRH", "Adult Cash", "EMER", "Health Care")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #11/28/2016#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "INTR List"																		'Script name
script_array(script_num).category               = "DEU"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("SNAP", "MFIP", "DWP", "HS/GRH", "Adult Cash", "EMER", "Health Care")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #11/28/2016#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).policy_references		= array("TE QTIP_#118_-_ASTERISK_ON_REPT_REVW 19.118", "TE HC_6-MONTH_RENEWALS 09.42","SHAREPOINT Medical_Assistance_(MA)_Renewals https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Medical-Assistance-(MA)-Renewals.aspx")
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Request Access to PRIV Case"
' script_array(script_num).description 			= "Sends a request to QI Knowledge Now to grant access in MAXIS for a Privileged Case."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("")
script_array(script_num).dlg_keys               = array("Oe")
script_array(script_num).subcategory            = array("TOP", "REQUESTS", "MAXIS")
script_array(script_num).release_date           = #08/19/2020#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)   'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	 'Set this array element to be a new script. Script details below...
script_array(script_num).script_name			= "Resolve HC EOMC in MMIS"													'Script name
' script_array(script_num).description 			= "BULK script that checks MMIS for all cases on EOMC for HC to ensure MMIS is set to close."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "Health Care")
script_array(script_num).dlg_keys               = array("Up", "Ex", "Ev")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "LIMITED"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Review QCR Reports"
' script_array(script_num).description 			= "Creates some case notes and assists with emails on reoccuring process issues."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("QI", "BZ", "SNAP", "MFIP", "DWP", "Adult Cash", "HS/GRH", "EMER", "Health Care")
script_array(script_num).dlg_keys               = array("Ex")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #09/19/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Revoucher"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Reviews", "MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #12/11/2016#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Search CASE NOTE"
' script_array(script_num).description 			= "Sends a SVES/QURY."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Utility", "Reviews")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("TOOL", "MAXIS")
script_array(script_num).release_date           = #11/17/2021#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).retirement_date        = #02/25/2022# 'Script removed until an agency wide process can be implemented
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Shelter Alternative"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #06/19/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Shelter Change Reported"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #11/28/2016#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).retirement_date		= #06/09/2025#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STAR"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Shelter Interview"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #06/19/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Sheriff Foreclosure"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #06/19/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references = array("CM Removing_or_Recalculating_income 08.06.15", "CM Glossary 02.61", "CM MFIP_Housing_Assistance_Grant 13.03.09", "CM Suspensions 22.18", "CM Opting_Out_of_MFIP_Cash_Portion 14.03.03.03", "CM When_to_Switch_Budget_cycles_SNAP 22.09.03", "TE SIGNIFICANT_CHANGE 02.13.11", "SHAREPOINT Significant_Change https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Budgeting_Significant_Change.aspx")
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "SMRT"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "Communication", "Health Care", "MFIP", "Reviews")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #01/19/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("SHAREPOINT State_Medical_Review_Team_(SMRT) https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/State_Medical_Review_Team_(SMRT).aspx")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"
script_array(script_num).specialty_redirect		= "MHC"

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
script_array(script_num).retirement_date		= #03/09/2022#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "SNAP Waived Interview"
' script_array(script_num).description 			= "Workflow for SNAP interview waiver."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Communication", "Application")
script_array(script_num).dlg_keys               = array("Cn", "Ev")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #01/02/2024#
script_array(script_num).hot_topic_date         = #01/30/2024#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/sitepages/SNAP-Waived-Interview-Now-Handles-Return-Contacts.aspx?"
script_array(script_num).retirement_date        = #10/31/24#
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("SHAREPOINT Processing_SNAP_Applications_with_Waived_Interviews https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/Processing-SNAP-Applications-with-Waived-Interviews.aspx")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Special EA"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #06/19/2017#
script_array(script_num).retirement_date 		= #02/22/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Sponsor Income"
' script_array(script_num).description 			= "Template for the sponsor income deeming calculation (it will also help calculate it for you)."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "Communication", "DWP", "EMER", "Health Care", "HS/GRH", "Income", "LTC", "MFIP", "Reviews", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Ev")
script_array(script_num).subcategory            = array("R-Z")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Subsequent Application"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Application", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Exp", "Oe", "Sm")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #1/31/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("CM Application_-_Pending_Cases 05.09.12")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "PRIMARY"
script_array(script_num).specialty_redirect		= "CA"

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
script_array(script_num).policy_references		= array("TE SPEC_XFER_For_Inter-Agency_Case_Transfers 02.08.134", "SHAREPOINT Transfer_To_Another_County https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/To_Another_County.aspx")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"
script_array(script_num).specialty_redirect		= "CA"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Trial Home Visit"																		'Script name
script_array(script_num).category               = "IV-E"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #11/25/2019#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).retirement_date		= #01/26/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Test Ex Parte Data Access"													'Script name
' script_array(script_num).description		    = "BULK script that captures specified DAILS for identified populations, and outputs them into a SQL Database."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("")
script_array(script_num).dlg_keys               = array("Ev")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #05/26/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Test Staff Info"													'Script name
' script_array(script_num).description		    = "BULK script that captures specified DAILS for identified populations, and outputs them into a SQL Database."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Utility")
script_array(script_num).dlg_keys               = array("Ev")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #05/01/2023#
script_array(script_num).retirement_date		= #05/19/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)   'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name 			= "TIKL FROM LIST"											'Script name
' script_array(script_num).description 			= "BULK script that gathers ABAWD/FSET codes for members on SNAP/MFIP active cases."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("QI", "MFIP", "DWP", "SNAP", "HS/GRH", "Health Care", "Adult Cash", "EMER")
script_array(script_num).dlg_keys               = array("Ex", "Tk")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 			= "TLR Report"																		'Script name
' script_array(script_num).description 			= "Updates FSET/ABAWD coding on STAT/WREG and case notes ABAWD exemptions."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ","ABAWD","SNAP")
script_array(script_num).dlg_keys               = array("Cn", "Up")
script_array(script_num).subcategory            = array("ABAWD")
script_array(script_num).release_date           = #06/17/2021#
script_array(script_num).hot_topic_date         = ""
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("CM Time-Limited_SNAP_Recipient 11.24", "CM Who_Is_Exempt_From_SNAP_Work_Registration 28.06.12")
script_array(script_num).usage_eval				= "STATS"


script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 			= "TLR Screening"																		'Script name
' script_array(script_num).description 			= "Updates FSET/ABAWD coding on STAT/WREG and case notes ABAWD exemptions."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("ABAWD", "Application", "Communication", "Reviews", "SNAP", "Renewals")
script_array(script_num).dlg_keys               = array("Cn", "Up")
script_array(script_num).subcategory            = array("ABAWD")
script_array(script_num).release_date           = #01/02/2024#
script_array(script_num).hot_topic_date         = #01/02/2024#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/SNAP-Interview-Waiver---New-Applications.aspx"
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("CM Time-Limited_SNAP_Recipient 11.24", "CM Who_Is_Exempt_From_SNAP_Work_Registration 28.06.12")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STAR"

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)   'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name 			= "Track Autoclose Overpayments"											'Script name
' script_array(script_num).description 			= "BULK script that gathers ABAWD/FSET codes for members on SNAP/MFIP active cases."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("QI", "MFIP", "SNAP")
script_array(script_num).dlg_keys               = array("Ex", "Wrd")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #03/29/2022#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""
script_array(script_num).retirement_date		= #6/1/2022#

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "UC Verification Request"
' script_array(script_num).description 			= "Creates an email of a houshold member to request Unemployment Compensation"
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Income", "Applications", "Reviews", "Utility", "SNAP", "MFIP", "DWP", "Adult Cash", "HS/GRH", "Health Care", "EMER")
script_array(script_num).dlg_keys               = array("Oe")
script_array(script_num).subcategory            = array("TOP", "REQUESTS")
script_array(script_num).release_date           = #09/16/2021#
script_array(script_num).hot_topic_date 		= #05/24/2022#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/sitepages/All-ES-Staff-Can-Email-hsph.es.deed-for-Unemployment-Verification.aspx"
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("CM Unearned_Income 17.12.03", "CM Applying_For_Other_Benefits 12.12", "EPM 2.2.3.4_Income_Methodology https://hcopub.dhs.state.mn.us/epm/2_2_3_4.htm","EPM 2.3.3.3.2.1_Countable_Income https://hcopub.dhs.state.mn.us/epm/2_3_3_3_2_1.htm", "SHAREPOINT Unemployment_Insurance https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Unemployment_Insurance.aspx")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

script_num = script_num + 1							'Increment by one
ReDim Preserve script_array(script_num)	'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script. Script details below...
script_array(script_num).script_name			= "UNEA Updater"										'Script name
' script_array(script_num).description 			= "BULK script that updates UNEA information and sends SPEC/MEMO for VA cases at ER."
script_array(script_num).category               = "ADMIN"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("BZ", "DWP", "EMER", "Health Care", "HS/GRH", "LTC", "MFIP", "SNAP", "Adult Cash")
script_array(script_num).dlg_keys               = array("Up", "Cn", "Sm", "Ex")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).retirement_date		= #07/25/2025#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array(script_num)	    'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	   'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "Update Check Dates"
' script_array(script_num).description		    = "Updates the dates on JOBS and UNEA to the correct dates for the footer month."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Utility")
script_array(script_num).dlg_keys               = array("Up")
script_array(script_num).subcategory            = array("TOP", "TOOL", "MAXIS")
script_array(script_num).release_date           = #07/17/2020#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Update Worker Signature"
' script_array(script_num).description 			= "Sets or updates the default worker signature for this user."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Utility")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("TOP", "TOOL")
script_array(script_num).release_date           = #10/01/2000#
script_array(script_num).hot_topic_date			= #10/13/2020#
script_array(script_num).hot_topic_link			= "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/sitepages/UTILITIES-%e2%80%93-Update-Worker-Signature.aspx"
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Utility Information"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Application", "MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #06/19/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1					'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "VA Verification Request"
' script_array(script_num).description 			= "Creates an email requesting VA Benefits."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Income", "Applications", "Reviews", "Utility", "SNAP", "MFIP", "DWP", "Adult Cash", "HS/GRH", "Health Care", "EMER")
script_array(script_num).dlg_keys               = array("Oe")
script_array(script_num).subcategory            = array("REQUESTS")
script_array(script_num).release_date           = #11/25/2020#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "LIMITED"

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
script_array(script_num).retirement_date		= #07/24/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Voucher Extended"
script_array(script_num).category               = "SHELTER"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Reviews", "MFIP", "EMER")
script_array(script_num).dlg_keys               = array("Cn")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #06/19/2017#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STATS"

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
script_array(script_num).retirement_date		= #11/20/2023#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "OPERATIONAL"



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
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("CM Time-Limited_SNAP_Recipient 11.24", "CM Who_Is_Exempt_From_SNAP_Work_Registration 28.06.12")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references      = array("CM Mandatory_Verifications 10.18", "CM Citizenship_and_Immigration_Status 11.03", "TE Citizenship_&_Immig_Ver._For_MA_APPL 02.08.166", "SHAREPOINT Acceptable_Verifications https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Acceptable_Verification.aspx")                   'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).retirement_date        = #09/09/2022#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1							   'Increment by one
ReDim Preserve script_array(script_num)	   'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	  'Set this array element to be a new script. Script details below...
script_array(script_num).script_name		    = "MEC2 Message"
' script_array(script_num).description		    = "Deletes non-actionable MEC2 messages and opens a dialog with policy links for actionable MEC2 messages."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #01/21/2025#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("TE New_HIRE_Matches 02.08.142")
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("TE New_HIRE_Matches 02.08.142")
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references      = array("CM Emergency_Aid_Eligibility_-_SNAP/Expedited_Food 04.04", "TE Expedited_SNAP_W/_Postponed_Verifs 02.10.01", "TE WREG_Expedited_SNAP_Postponed_Verifs 02.05.70.01")
script_array(script_num).usage_eval				= "STANDARD"


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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "SDX Match"																		'Script name
' script_array(script_num).description 			= "Opens a dialog with links to policy information for processing DAIL messages."
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Adult Cash", "Health Care", "HS/GRH", "Income", "LTC", "SNAP")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #11/06/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("CM Interim_Assistance_Agreements 12.12.03", "CM Interim_Assistance_Reimbursement_Interface 02.12.14")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= "STANDARD"

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "View INFC"																		'Script name
script_array(script_num).category               = "DEU"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Navigation")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/05/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'
script_array(script_num).usage_eval				= ""

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Find MEMB in MMIS"
' script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Navigation")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #08/19/2024#
script_array(script_num).hot_topic_link			= ""
script_array(script_num).used_for_elig			= False
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'

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
script_array(script_num).policy_references		= array("")						'SEE Line 58 for format'

script_num = script_num + 1                     'Increment by one
ReDim Preserve script_array(script_num)         'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie     'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name            = "XML File Cleanup"
' script_array(script_num).description          = "Navigates to and sets all the panels in STAT to 'V'iew in the PNLP order."
script_array(script_num).category               = "CA"
script_array(script_num).workflows              = ""
script_array(script_num).tags                   = array("Utility")
script_array(script_num).dlg_keys               = array("")
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #03/04/2024#
script_array(script_num).hot_topic_link         = ""
script_array(script_num).used_for_elig          = False
script_array(script_num).policy_references      = array("")                     'SEE Line 58 for format'
' for test_thing = 0 to UBound(script_array)
' 	MsgBox script_array(test_thing).description
' next
