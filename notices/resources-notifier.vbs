'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - RESOURCES NOTIFIER.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 195                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
'END OF stats block==============================================================================================

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
call changelog_update("01/09/2019", "Updated COPE resource to EMERGENCY MENTAL HEALTH SERVICES. Also Updated DHS MMIS Helpdesk text to DHS MMIS RECIPIENT HELPDESK. These changes align with the resources provided in the CONTACTS FOR HSR's resource.", "Ilse Ferris, Hennepin County")
call changelog_update("12/18/2018", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog Resources_MEMO_dialog, 0, 0, 206, 240, "Resources MEMO"
  EditBox 60, 5, 50, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    PushButton 150, 5, 50, 10, "Check All", check_all_button
    OkButton 95, 220, 50, 15
    CancelButton 150, 220, 50, 15
  CheckBox 15, 45, 140, 10, "Community Action Partnership - CAP", cap_checkbox
  CheckBox 15, 60, 115, 10, "DHS MMIS Recipient HelpDesk", MMIS_helpdesk_checkbox
  CheckBox 15, 75, 180, 10, "DHS MNSure Helpdesk   * NOT FOR MA CLIENTS", MNSURE_helpdesk_checkbox
  CheckBox 15, 90, 145, 10, "Disability Hub (Disability Linkage Line)", disability_hub_checkbox
  CheckBox 15, 105, 125, 10, "Emergency Mental Health Services", emer_mental_health_checkbox
  CheckBox 15, 120, 175, 10, "Emergency Food Shelf Network (The Food Group)", emer_food_network_checkbox
  CheckBox 15, 135, 50, 10, "Front Door", front_door_checkbox
  CheckBox 15, 150, 75, 10, "Senior Linkage Line", sr_linkage_line_checkbox
  CheckBox 15, 165, 130, 10, "United Way First Call for Help (211)", united_way_checkbox
  CheckBox 15, 180, 60, 10, "Xcel Energy", xcel_checkbox
  EditBox 80, 200, 120, 15, worker_signature
  Text 10, 30, 195, 10, "Check any to send detail about the service to a client:"
  Text 10, 205, 65, 10, "Worker signature:"
  Text 10, 10, 50, 10, "Case number:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""
'Searches for a case number
call MAXIS_case_number_finder(MAXIS_case_number)

'This Do...loop shows the appointment letter dialog, and contains logic to require most fields.
DO
	Do
		err_msg = ""
		Dialog Resources_MEMO_dialog
		If ButtonPressed = cancel then stopscript
        If cap_checkbox = unchecked AND emer_mental_health_checkbox = unchecked AND MMIS_helpdesk_checkbox = unchecked AND MNSURE_helpdesk_checkbox = unchecked AND disability_hub_checkbox = unchecked AND emer_food_network_checkbox = unchecked AND front_door_checkbox = unchecked AND sr_linkage_line_checkbox = unchecked AND united_way_checkbox = unchecked AND xcel_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "You must select at least one resource."
		If isnumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & "You must fill in a valid case number." & vbNewLine
		If worker_signature = "" then err_msg = err_msg & "You must sign your case note." & vbNewLine
        If ButtonPressed = check_all_button Then
            err_msg = "LOOP" & err_msg

            cap_checkbox = checked
            MMIS_helpdesk_checkbox = checked
            MNSURE_helpdesk_checkbox = checked
            disability_hub_checkbox = checked
            emer_food_network_checkbox = checked
            emer_mental_health_checkbox = checked
            front_door_checkbox = checked
            sr_linkage_line_checkbox = checked
            united_way_checkbox = checked
            xcel_checkbox = checked
        End If
		IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN msgbox err_msg
	Loop until err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

script_to_say = "Resource detail:" & vbNewLine

If cap_checkbox = checked Then
    script_to_say = script_to_say & vbNewLine & "CAP - Community Action Partnership (Includes Energy Assist)" & vbNewLine &_
        "Hours: Mon-Fri 8:00AM - 4:30PM Website: www.caphennepin.org" & vbNewLine &_
        "Locations: Minneapolis Urban League   Phone: 952-930-3541" & vbNewLine &_
        "           MN Council of Churches     Phone: 952-933-9639" & vbNewLine &_
        "           Sabathani Community Center Phone: 952-930-3541" & vbNewLine &_
        "           St. Louis Park             Phone: 952-933-9639" & vbNewLine &_
        "--   --   --   --   --   --   --   --   --   --   --"
End If
If MMIS_helpdesk_checkbox = checked Then
    script_to_say = script_to_say & vbNewLine & "MN Health Care Recipient Help Desk - 651-431-2670" & vbNewLine &_
    "--   --   --   --   --   --   --   --   --   --   --"
End If
If MNSURE_helpdesk_checkbox = checked Then
    script_to_say = script_to_say & vbNewLine & "MNSure Helpdesk - 1-855-366-7873 (1-855-3MNSURE)" & vbNewLine &_
    "--   --   --   --   --   --   --   --   --   --   --"
End If
If disability_hub_checkbox = checked Then
    script_to_say = script_to_say & vbNewLine & "Disability Hub (formerly Disability Linkage Line)" & vbNewLine &_
        "Phone: 1-866-333-2466 -Hrs: Mon - Fri 8:00AM - 5:00PM" & vbNewLine &_
        "Website: disabilityhubmn.org" & vbNewLine &_
        "--   --   --   --   --   --   --   --   --   --   --"
End If
If emer_food_network_checkbox = checked Then
    script_to_say = script_to_say & vbNewLine & "The Food Group (formerly Emergency Food Network)" & vbNewLine &_
        "Phone: 763-450-3860  - Website: thefoodgroupmn.org" & vbNewLine &_
        "--   --   --   --   --   --   --   --   --   --   --"
End If
If emer_mental_health_checkbox = checked Then
    script_to_say = script_to_say & vbNewLine & "Emergency Mental Health Services" & vbNewLine &_
        "Adults 18 and older (COPE): 612-596-1223" & vbNewLine &_
        "Children (Child Crisis Services): 612-348-2233" & vbNewLine &_
        "--   --   --   --   --   --   --   --   --   --   --"
End If
If front_door_checkbox = checked Then
    script_to_say = script_to_say & vbNewLine & "Hennepin County FRONT DOOR - 612-348-4111" & vbNewLine &_
    "--   --   --   --   --   --   --   --   --   --   --"
End If
If sr_linkage_line_checkbox = checked Then
    script_to_say = script_to_say & vbNewLine & "Senior Linkage Line" & vbNewLine &_
        "Phone: 1-800-333-2433  - Hours: Mon - Fri 8:00 AM - 4:30 PM" & vbNewLine &_
        "   Currently has extended hours Mon - Thur 4:30 PM - 6:30 PM" & vbNewLine &_
        "Website: metroaging.org" & vbNewLine &_
        "--   --   --   --   --   --   --   --   --   --   --"
End If
If united_way_checkbox = checked Then
    script_to_say = script_to_say & vbNewLine & "United Way First Call for Help (211)" & vbNewLine &_
        "Phone: 1-800-543-7709 OR 651- 291-0211  - Available 24 Hrs" & vbNewLine &_
        "Website: www.211unitedway.org" & vbNewLine &_
        "--   --   --   --   --   --   --   --   --   --   --"
End If
If xcel_checkbox = checked Then
    script_to_say = script_to_say & vbNewLine & "Xcel Energy - 1-800-331-5262" & vbNewLine &_
    "--   --   --   --   --   --   --   --   --   --   --"
End If

script_to_say = script_to_say & vbNewLine & "Relay any of the above information to the client verbally now." & vbNewLine &_
    "Then press OK and all of this detail will be added to a SPEC/MEMO so the client can have the information in writing."

MsgBox script_to_say

BeginDialog Resources_MEMO_dialog, 0, 0, 106, 80, "Dialog"
  DropListBox 15, 40, 80, 45, "SPEC/MEMO"+chr(9)+"Word Document", resource_method
  ButtonGroup ButtonPressed
    OkButton 45, 60, 50, 15
  Text 5, 10, 90, 20, "What is the best format for the Resource information?"
EndDialog

dialog Resources_MEMO_dialog

'Create a question - MEMO or Word Doc'
If resource_method = "SPEC/MEMO" Then
    Call start_a_new_spec_memo  ' start the memo writing process

    need_divider = FALSE
    'Writes the MEMO.
    call write_variable_in_SPEC_MEMO("  ----Outside Resources - current as of " & date & "----")
    If cap_checkbox = checked Then
        If need_divider = TRUE Then call write_variable_in_SPEC_MEMO("     --   --   --   --   --   --   --   --   --   --")
        call write_variable_in_SPEC_MEMO("* CAP - Community Action Partnership (Inc. Energy Assist)")
        call write_variable_in_SPEC_MEMO("Hours: Mon-Fri 8:00AM - 4:30PM Website: www.caphennepin.org")
        call write_variable_in_SPEC_MEMO("Locations: Minneapolis Urban League   Phone: 952-930-3541")
        call write_variable_in_SPEC_MEMO("           MN Council of Churches     Phone: 952-933-9639")
        call write_variable_in_SPEC_MEMO("           Sabathani Community Center Phone: 952-930-3541")
        call write_variable_in_SPEC_MEMO("           St. Louis Park             Phone: 952-933-9639")
        need_divider = FALSE
        'call write_variable_in_SPEC_MEMO("--   --   --   --   --   --   --   --   --   --   --")
    End If
    
    If MMIS_helpdesk_checkbox = checked Then
        If need_divider = TRUE Then call write_variable_in_SPEC_MEMO("     --   --   --   --   --   --   --   --   --   --")
        call write_variable_in_SPEC_MEMO("* MN Health Care Recipient Help Desk - 651-431-2670")
        need_divider = TRUE
        'call write_variable_in_SPEC_MEMO("--   --   --   --   --   --   --   --   --   --   --")
    End If
    If MNSURE_helpdesk_checkbox = checked Then
        If need_divider = TRUE Then call write_variable_in_SPEC_MEMO("     --   --   --   --   --   --   --   --   --   --")
        call write_variable_in_SPEC_MEMO("* MNSure Helpdesk - 1-855-366-7873 (1-855-3MNSURE)")
        need_divider = TRUE
        'call write_variable_in_SPEC_MEMO("--   --   --   --   --   --   --   --   --   --   --")
    End If
    If disability_hub_checkbox = checked Then
        If need_divider = TRUE Then call write_variable_in_SPEC_MEMO("     --   --   --   --   --   --   --   --   --   --")
        call write_variable_in_SPEC_MEMO("* Disability Hub (formerly Disability Linkage Line)")
        call write_variable_in_SPEC_MEMO("    Phone: 1-866-333-2466 -Hrs: Mon - Fri 8:00AM - 5:00PM")
        call write_variable_in_SPEC_MEMO("    Website: disabilityhubmn.org")
        need_divider = FALSE
        'call write_variable_in_SPEC_MEMO("--   --   --   --   --   --   --   --   --   --   --")
    End If
    If emer_food_network_checkbox = checked Then
        If need_divider = TRUE Then call write_variable_in_SPEC_MEMO("     --   --   --   --   --   --   --   --   --   --")
        call write_variable_in_SPEC_MEMO("* The Food Group (formerly Emergency Food Network)")
        call write_variable_in_SPEC_MEMO("     Phone: 763-450-3860  - Website: thefoodgroupmn.org")
        need_divider = FALSE
        'call write_variable_in_SPEC_MEMO("--   --   --   --   --   --   --   --   --   --   --")
    End If
    If emer_mental_health_checkbox = checked Then
        If need_divider = TRUE Then call write_variable_in_SPEC_MEMO("     --   --   --   --   --   --   --   --   --   --")
        call write_variable_in_SPEC_MEMO("* Emergency Mental Health Services")
        call write_variable_in_SPEC_MEMO("       Adults 18 and older (COPE): 612-596-1223")
        call write_variable_in_SPEC_MEMO("       Children (Child Crisis Services): 612-348-2233")
        need_divider = FALSE
    End If
    If front_door_checkbox = checked Then
        If need_divider = TRUE Then call write_variable_in_SPEC_MEMO("     --   --   --   --   --   --   --   --   --   --")
        call write_variable_in_SPEC_MEMO("* Hennepin County FRONT DOOR - 612-348-4111")
        need_divider = TRUE
        'call write_variable_in_SPEC_MEMO("--   --   --   --   --   --   --   --   --   --   --")
    End If
    If sr_linkage_line_checkbox = checked Then
        If need_divider = TRUE Then call write_variable_in_SPEC_MEMO("     --   --   --   --   --   --   --   --   --   --")
        call write_variable_in_SPEC_MEMO("* Senior Linkage Line - Hours: Mon - Fri 8:00AM - 4:30PM")
        call write_variable_in_SPEC_MEMO("   Phone: 1-800-333-2433")
        call write_variable_in_SPEC_MEMO("   Currently has extended hours Mon - Thur 4:30PM - 6:30PM")
        call write_variable_in_SPEC_MEMO("   Website: metroaging.org")
        need_divider = FALSE
        'call write_variable_in_SPEC_MEMO("--   --   --   --   --   --   --   --   --   --   --")
    End If
    If united_way_checkbox = checked Then
        If need_divider = TRUE Then call write_variable_in_SPEC_MEMO("     --   --   --   --   --   --   --   --   --   --")
        call write_variable_in_SPEC_MEMO("* United Way First Call for Help (211)")
        call write_variable_in_SPEC_MEMO("   Phone: 1-800-543-7709 OR 651- 291-0211 - 24 Hrs")
        call write_variable_in_SPEC_MEMO("   Website: www.211unitedway.org")
        need_divider = FALSE
        'call write_variable_in_SPEC_MEMO("--   --   --   --   --   --   --   --   --   --   --")
    End If
    If xcel_checkbox = checked Then
        If need_divider = TRUE Then call write_variable_in_SPEC_MEMO("     --   --   --   --   --   --   --   --   --   --")
        call write_variable_in_SPEC_MEMO("* Xcel Energy - 1-800-331-5262")
        'call write_variable_in_SPEC_MEMO("--   --   --   --   --   --   --   --   --   --   --")
    End If

    'Exits the MEMO
    PF4
End If

If resource_method = "Word Document" Then
    '****writing the word document
    Set objWord = CreateObject("Word.Application")
    Const wdDialogFilePrint = 88
    Const end_of_doc = 6
    objWord.Caption = "Outside Resource Information"
    objWord.Visible = True

    Set objDoc = objWord.Documents.Add()
    Set objSelection = objWord.Selection

    objSelection.PageSetup.LeftMargin = 50
    objSelection.PageSetup.RightMargin = 50
    objSelection.PageSetup.TopMargin = 30
    objSelection.PageSetup.BottomMargin = 25

    todays_date = date & ""
    objSelection.Font.Name = "Arial"
    objSelection.Font.Size = "14"
    objSelection.Font.Bold = TRUE
    objSelection.TypeText "Outside Resource Information - Current as of "
    objSelection.TypeText todays_date
    objSelection.TypeParagraph()
    objSelection.ParagraphFormat.SpaceAfter = 0
    'objSelection.TypeText "The following is contact information for other agencies/programs:"
    'objSelection.TypeParagraph()
    'objSelection.TypeParagraph()

    objSelection.Font.Size = "12"
    objSelection.Font.Bold = FALSE
    If cap_checkbox = checked Then
        objSelection.TypeText "* CAP - Community Action Partnership (Inc. Energy Assist)" & vbCr
        objSelection.TypeText "  Hours: Mon-Fri 8:00AM - 4:30PM Website: www.caphennepin.org" & vbCr
        objSelection.TypeText "  Locations: Minneapolis Urban League   Phone: 952-930-3541" & vbCr
        objSelection.TypeText "                    MN Council of Churches     Phone: 952-933-9639" & vbCr
        objSelection.TypeText "                    Sabathani Community Center Phone: 952-930-3541" & vbCr
        objSelection.TypeText "                    St. Louis Park             Phone: 952-933-9639" & vbCr
        'objSelection.TypeText "_____________________________" & chr(10)
        objSelection.TypeParagraph()
    End If
    If MMIS_helpdesk_checkbox = checked Then
        objSelection.TypeText "* MN Health Care Recipient Help Desk - 651-431-2670" & vbCr
        'objSelection.TypeText "_____________________________" & vbCr
        objSelection.TypeParagraph()
    End If
    If MNSURE_helpdesk_checkbox = checked Then
        objSelection.TypeText "* MNSure Helpdesk - 1-855-366-7873 (1-855-3MNSURE)" & vbCr
        'objSelection.TypeText "_____________________________" & vbCr
        objSelection.TypeParagraph()
    End If
    If disability_hub_checkbox = checked Then
        objSelection.TypeText "* Disability Hub (formerly Disability Linkage Line)" & vbCr
        objSelection.TypeText "    Phone: 1-866-333-2466 -Hrs: Mon - Fri 8:00AM - 5:00PM" & vbCr
        objSelection.TypeText "    Website: disabilityhubmn.org" & vbCr
        'objSelection.TypeText "_____________________________" & vbCr
        objSelection.TypeParagraph()
    End If
    If emer_food_network_checkbox = checked Then
        objSelection.TypeText "* The Food Group (formerly Emergency Food Network)" & vbCr
        objSelection.TypeText "     Phone: 763-450-3860  - Website: thefoodgroupmn.org" & vbCr
        'objSelection.TypeText "_____________________________" & vbCr
        objSelection.TypeParagraph()
    End If
    If emer_mental_health_checkbox = checked Then
        objSelection.TypeText "* Emergency Mental Health Services" & vbCr
        objSelection.TypeText "       Adults 18 and older (COPE): 612-596-1223" & vbCr
        objSelection.TypeText "       Children (Child Crisis Services): 612-348-2233" & vbCr
        'objSelection.TypeText "_____________________________" & vbCr
        objSelection.TypeParagraph()
    End If
    If front_door_checkbox = checked Then
        objSelection.TypeText "* Hennepin County FRONT DOOR - 612-348-4111" & vbCr
        'objSelection.TypeText "_____________________________" & vbCr
        objSelection.TypeParagraph()
    End If
    If sr_linkage_line_checkbox = checked Then
        objSelection.TypeText "* Senior Linkage Line" & vbCr
        objSelection.TypeText "   Phone: 1-800-333-2433 - Hours: Mon - Fri 8:00AM - 4:30PM" & vbCr
        objSelection.TypeText "   Currently has extended hours Mon - Thur 4:30PM - 6:30PM" & vbCr
        objSelection.TypeText "   Website: metroaging.org" & vbCr
        'objSelection.TypeText "_____________________________" & vbCr
        objSelection.TypeParagraph()
    End If
    If united_way_checkbox = checked Then
        objSelection.TypeText "* United Way First Call for Help (211)" & vbCr
        objSelection.TypeText "   Phone: 1-800-543-7709 OR 651- 291-0211 - 24 Hrs" & vbCr
        objSelection.TypeText "   Website: www.211unitedway.org" & vbCr
        'objSelection.TypeText "_____________________________" & vbCr
        objSelection.TypeParagraph()
    End If
    If xcel_checkbox = checked Then
        objSelection.TypeText "* Xcel Energy - 1-800-331-5262" & vbCr
        'objSelection.TypeText "_____________________________" & vbCr
        objSelection.TypeParagraph()
    End If


End If
'Navigates to CASE/NOTE and starts a blank one
start_a_blank_CASE_NOTE

'Writes the case note--------------------------------------------
call write_variable_in_CASE_NOTE("Outside resource information sent to client")
If resource_method = "SPEC/MEMO" Then Call write_variable_in_CASE_NOTE("* Information added to SPEC/MEMO to send in overnight batch.")
If resource_method = "Word Document" Then Call write_variable_in_CASE_NOTE("* Information added to Word Document for printing locally.")

IF cap_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Compunity Action Partnership - CAP (Energy Assistance)")
IF MMIS_helpdesk_checkbox = checked Then Call write_variable_in_CASE_NOTE("* DHS MHCP Recipient HelpDesk")
IF MNSURE_helpdesk_checkbox = checked Then Call write_variable_in_CASE_NOTE("* DHS MNSure HelpDesk")
IF disability_hub_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Disability Hub")
IF emer_food_network_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Emergency Food Network")
IF emer_mental_health_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Emergency Mental Health Services")
IF front_door_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Front Door")
IF sr_linkage_line_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Senior Linkage Line")
IF united_way_checkbox = checked Then Call write_variable_in_CASE_NOTE("* United Way - 211")
IF xcel_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Xcel Energy")
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")
