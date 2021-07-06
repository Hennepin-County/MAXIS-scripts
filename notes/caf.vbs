'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - CAF.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 1200                     'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
'END OF stats block=========================================================================================================

IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
    IF on_the_desert_island = TRUE Then
        FuncLib_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\desert-island\MASTER FUNCTIONS LIBRARY.vbs"
        Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
        Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
        text_from_the_other_script = fso_command.ReadAll
        fso_command.Close
        Execute text_from_the_other_script
    ELSEIF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
Call changelog_update("11/6/2020", "The script will now check the interview date entered on PROG or REVW to confirm the updated happened accurately when the script is tasked with updating PROG or REVW.##~## ##~##There may be a message that the update failed, this means this update must be completed manually.##~##", "Casey Love, Hennepin County")
Call changelog_update("11/6/2020", "Added new functionality to have the script update the 'REVW' panel with the Interview Date and CAF date for Recertification Cases.##~## As this is a new functionality, please let us know how it works for you!.##~##", "Casey Love, Hennepin County")
Call changelog_update("10/09/2020", "Updated the functionality to adjust review dates for some cases to not require an interview date for Adult Cash Recertification cases.##~##", "Casey Love, Hennepin County")
Call changelog_update("10/01/2020", "Updated Standard Utility Allowances for 10/2020.", "Ilse Ferris, Hennepin County")
call changelog_update("09/01/2020", "UPDATE IN TESTING##~## ##~##There is new functionality added to this script to assist in updating REVW for some cases. There is detail in SIR about cases that need to have the next ER date adjusted due to a previous postponement. ##~## ##~##We have not had time with this functionality to complete live testing so all script runs will be a part of the test. Please let us know if you have any issues running this script.", "Casey Love, Hennepin County")
call changelog_update("05/18/2020", "Additional handling to be sure when saving the interview date to PROG the script does not get stuck.", "Casey Love, Hennepin County")
call changelog_update("04/02/2020", "BUG FIX - The 'notes on child support' detail was not always pulling into the case note. Updated to ensure the information entered in this field will be entered in the NOTE.##~##", "Casey Love, Hennepin County")
call changelog_update("03/01/2020", "Updated TIKL functionality.", "Ilse Ferris")
Call changelog_update("01/29/2020", "When entering a expedited approval date or denial date on Dialog 8, the script will evaluate to be sure this date is not in the future. The script was requiring delay explanation when dates in the future were entered by accident, this additional handling will provide greater clarity of the updates needed.", "Casey Love, Hennepin County")
Call changelog_update("01/08/2020", "BUG FIX - When selecting CASH and SNAP for an MFIP Recertification, the script would error out and could not continue due to not being able to find the SNAP ER date on REVW. Updated the script to ignore that blank recert date.##~##", "Casey Love, Hennepin County")
Call changelog_update("11/27/2019", "Added handling to support 10 or more children on STAT/PARE panels.", "Ilse Ferris, Hennepin County")
Call changelog_update("11/22/2019", "Added a checkbox on the verifications dialog pop-up. This checkbox will add detail to the verifications case note that there are verifications that have been postponed.##~##", "Casey Love, Hennepin County")
Call changelog_update("11/22/2019", "Added handling for ID information and ID requirements for household members and AREP (if interviewed). This information is added to Dialog One.##~##This functionality mandates detail if the ID verification is 'Other' and is required.##~##", "Casey Love, Hennepin County")
Call changelog_update("11/14/2019", "BUG FIX - Dialog 4 had some fields overlapping each other sometimes, which made it difficult to read/update. Fixed the layout of Dialog  4 (CSES).##~##", "Casey Love, Hennepin County")
Call changelog_update("11/08/2019", "Added handling for the script to change a 4 digit footer year to a 2 digit footer year (2019 becomes 19) when entering recertification month and year by program. ##~##", "Casey Love, Hennepin County")
Call changelog_update("11/07/2019", "BUG FIX - Dialog 4 was sometimes too short. If there are a number of people with child support income, not all of the child support detail would be viewable as it would be taller than the computer screen. Updated the script so that Dialog 4 now has tabs if there are more than four members with child support income, so there are multiple pages of Dialog 4 (like Dialog 2 and 3).##~##", "Casey Love, Hennepin County")
Call changelog_update("10/16/2019", "BUG Fix - sometimes the script hit an error after leaving Dialog 8 - this should resolve that error. ##~## ##~## Added a NEW BUTTON that will display the Missing Fields Message (also called the 'Error Message') after clicking 'Done' on dialog 8 if the script needs updates. Look for the button 'Show Dialog Review Message' on each dialog after the message shows for the first time. ##~## This button will allow you to review the missing fields or updates that need to be made so that you do not have to try to remember them. The button only appears after the message was shown for the first time.##~##", "Casey Love, Hennepin County")
Call changelog_update("10/14/2019", "Added autofill functionality for TIME and SANC panels so the editboxes are filled if the panel is present.##~##", "Casey Love, Hennepin County")
call changelog_update("10/10/2019", "Updated 3 bugs/issues: ##~## ##~## - Sometimmes the list of clients on the 'Qualifying Quesitons Dialog' was not filled and was blank, this is now resolved and should always have a list of clients. ##~## - The script was 'forgetting' informmation typed into a ComboBox when a dialog appears for a subsequent time. This is now resolved. ##~## - Added headers to the mmissed fields/error message after Dialog 8 for more readability.", "Casey Love, Hennepin County")
Call changelog_update("10/01/2019", "CAF Functionality is enhanced for more complete and comprehensive documentation of CAF processing. This new functionality has been available for trial for the past 2 weeks. ##~## ##~## Live Skype Demos of this new functionality are availble this week and next. See Hot Topics for more details about the enhanceed functionality and the demo sessions. ##~##", "Casey Love, Hennepin County")
Call changelog_update("10/01/2019", "This script will be updated at the end of the day (10/1/2019) to the new CAF functionality. Additional details and resources can be found in Hot Topics or the BlueZone Script Team Sharepoint page.", "Casey Love, Hennepin County")
Call changelog_update("04/10/2019", "There was a bug that sometimes made the dialogs write over each other and be illegible, updated the script to keep this from happening.", "Casey Love, Hennepin County")
Call changelog_update("03/06/2019", "Added 2 new options to the Notes on Income button to support referencing CASE/NOTE made by Earned Income Budgeting.", "Casey Love, Hennepin County")
CALL changelog_update("10/17/2018", "Updated dialog box to reflect currecnt application process.", "MiKayla Handley, Hennepin County")
call changelog_update("05/05/2018", "Added autofill functionality for the DIET panel. NAT errors have been resolved.", "Ilse Ferris, Hennepin County")
call changelog_update("05/04/2018", "Removed autofill functionality for the DIET panel temporarily until MAXIS help desk can resolve NAT errors.", "Ilse Ferris, Hennepin County")
call changelog_update("01/11/2017", "Adding functionality to offer a TIKL for 12 month contact on 24 month SNAP renewals.", "Charles Potter, DHS")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS =================================================================================================================
'This function will message box the err_msg from this script, outlining it by dialog and adding the headers.
'This function should be added to the end of the dialogs after the review button and at the end of dialog 8 after the error message collection.
function display_errors(the_err_msg, execute_nav)
    If the_err_msg <> "" Then       'If the error message is blank - there is nothing to show.
        If left(the_err_msg, 3) = "~!~" Then the_err_msg = right(the_err_msg, len(the_err_msg) - 3)     'Trimming the message so we don't have a blank array item
        err_array = split(the_err_msg, "~!~")           'making the list of errors an array.

        error_message = ""                              'blanking out variables
        msg_header = ""
        for each message in err_array                   'going through each error message to order them and add headers'
            current_listing = left(message, 1)          'This is the dialog the error came from
            If current_listing <> msg_header Then                   'this is comparing to the dialog from the last message - if they don't match, we need a new header entered
                If current_listing = "1" Then tagline = ": Personal Information"        'Adding a specific tagline to the header for the errors
                If current_listing = "2" Then tagline = ": JOBS"
                If current_listing = "3" Then tagline = ": BUSI"
                If current_listing = "4" Then tagline = ": Child Support"
                If current_listing = "5" Then tagline = ": Unearned Income"
                If current_listing = "6" Then tagline = ": WREG, Expenses, Address"
                If current_listing = "7" Then tagline = ": Assets and Misc."
                If current_listing = "8" Then tagline = ": Interview Detail"
                error_message = error_message & vbNewLine & vbNewLine & "----- Dialog " & current_listing & tagline & " -------"    'This is the header verbiage being added to the message text.
            End If
            if msg_header = "" Then back_to_dialog = current_listing
            msg_header = current_listing        'setting for the next loop

            message = replace(message, "##~##", vbCR)       'This is notation used in the creation of the message to indicate where we want to have a new line.'

            error_message = error_message & vbNewLine & right(message, len(message) - 2)        'Adding the error information to the message list.
        Next

        'This is the display of all of the messages.
        view_errors = MsgBox("In order to complete the script and CASE/NOTE, additional details need to be added or refined. Please review and update." & vbNewLine & error_message, vbCritical, "Review detail required in Dialogs")

        'The function can be operated without moving to a different dialog or not. The only time this will be activated is at the end of dialog 8.
        If execute_nav = TRUE Then
            If back_to_dialog = "1" Then ButtonPressed = dlg_one_button         'This calls another function to go to the first dialog that had an error
            If back_to_dialog = "2" Then ButtonPressed = dlg_two_button
            If back_to_dialog = "3" Then ButtonPressed = dlg_three_button
            If back_to_dialog = "4" Then ButtonPressed = dlg_four_button
            If back_to_dialog = "5" Then ButtonPressed = dlg_five_button
            If back_to_dialog = "6" Then ButtonPressed = dlg_six_button
            If back_to_dialog = "7" Then ButtonPressed = dlg_seven_button
            If back_to_dialog = "8" Then ButtonPressed = dlg_eight_button

            Call assess_button_pressed          'this is where the navigation happens
        End If
    End If
End Function

Function HCRE_panel_bypass()
    'handling for cases that do not have a completed HCRE panel
    PF3		'exits PROG to prommpt HCRE if HCRE insn't complete
    Do
        EMReadscreen HCRE_panel_check, 4, 2, 50
        If HCRE_panel_check = "HCRE" then
            PF10	'exists edit mode in cases where HCRE isn't complete for a member
            PF3
        END IF
    Loop until HCRE_panel_check <> "HCRE"
End Function

'This function calls the dialog to determine and assess the household Composition
'This also determines the members that are including in gathering information.
function HH_comp_dialog(HH_member_array)
	CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.

    member_count = 0            'resetting these counts/variables
    adult_cash_count = 0
    child_cash_count = 0
    adult_snap_count = 0
    child_snap_count = 0
    adult_emer_count = 0
    child_emer_count = 0
	DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
		EMReadscreen ref_nbr, 2, 4, 33
        EMReadScreen access_denied_check, 13, 24, 2         'Sometimes MEMB gets this access denied issue and we have to work around it.
        If access_denied_check = "ACCESS DENIED" Then
            PF10
            last_name = "UNABLE TO FIND"
            first_name = " - Access Denied"
            mid_initial = ""
        Else
    		EMReadscreen last_name, 25, 6, 30
    		EMReadscreen first_name, 12, 6, 63
    		EMReadscreen mid_initial, 1, 6, 79
            EMReadScreen memb_age, 3, 8, 76
            memb_age = trim(memb_age)
            If memb_age = "" Then memb_age = 0
            memb_age = memb_age * 1
    		last_name = trim(replace(last_name, "_", ""))
    		first_name = trim(replace(first_name, "_", ""))
    		mid_initial = replace(mid_initial, "_", "")
            EMReadScreen id_verif_code, 2, 9, 68


            EMReadScreen rel_to_applcnt, 2, 10, 42              'reading the relationship from MEMB'
            If rel_to_applcnt = "02" Then relationship_detail = relationship_detail & "Memb " & ref_nbr & " is the Spouse of Memb 01.; "
            If rel_to_applcnt = "04" Then relationship_detail = relationship_detail & "Memb " & ref_nbr & " is the Parent of Memb 01.; "
            If rel_to_applcnt = "05" Then relationship_detail = relationship_detail & "Memb " & ref_nbr & " is the Sibling of Memb 01.; "
            If rel_to_applcnt = "12" Then relationship_detail = relationship_detail & "Memb " & ref_nbr & " is the Niece of Memb 01.; "
            If rel_to_applcnt = "13" Then relationship_detail = relationship_detail & "Memb " & ref_nbr & " is the Nephew of Memb 01.; "
            If rel_to_applcnt = "15" Then relationship_detail = relationship_detail & "Memb " & ref_nbr & " is the Grandparent of Memb 01.; "
            If rel_to_applcnt = "16" Then relationship_detail = relationship_detail & "Memb " & ref_nbr & " is the Grandchild of Memb 01.; "
        End If

        ReDim Preserve ALL_MEMBERS_ARRAY(clt_notes, member_count)       'resizing the array to add the next household member

        ALL_MEMBERS_ARRAY(memb_numb, member_count) = ref_nbr            'adding client information to the array
        ALL_MEMBERS_ARRAY(clt_name, member_count) = last_name & ", " & first_name & " " & mid_initial
        ALL_MEMBERS_ARRAY(full_clt, member_count) = ref_nbr & " - " & first_name & " " & last_name
        ALL_MEMBERS_ARRAY(clt_age, member_count) = memb_age

        If id_verif_code = "BC" Then ALL_MEMBERS_ARRAY(clt_id_verif, member_count) = "BC - Birth Certificate"
        If id_verif_code = "RE" Then ALL_MEMBERS_ARRAY(clt_id_verif, member_count) = "RE - Religious Record"
        If id_verif_code = "DL" Then ALL_MEMBERS_ARRAY(clt_id_verif, member_count) = "DL - Drivers License/ST ID"
        If id_verif_code = "DV" Then ALL_MEMBERS_ARRAY(clt_id_verif, member_count) = "DV - Divorce Decree"
        If id_verif_code = "AL" Then ALL_MEMBERS_ARRAY(clt_id_verif, member_count) = "AL - Alien Card"
        If id_verif_code = "AD" Then ALL_MEMBERS_ARRAY(clt_id_verif, member_count) = "AD - Arrival//Depart"
        If id_verif_code = "DR" Then ALL_MEMBERS_ARRAY(clt_id_verif, member_count) = "DR - Doctor Stmt"
        If id_verif_code = "PV" Then ALL_MEMBERS_ARRAY(clt_id_verif, member_count) = "PV - Passport/Visa"
        If id_verif_code = "OT" Then ALL_MEMBERS_ARRAY(clt_id_verif, member_count) = "OT - Other Document"
        If id_verif_code = "NO" Then ALL_MEMBERS_ARRAY(clt_id_verif, member_count) = "NO - No Verif Prvd"

        If cash_checkbox = checked Then             'If Cash is selected
            ALL_MEMBERS_ARRAY(include_cash_checkbox, member_count) = checked    'default to having the counted boxes checked for SNAP
            ALL_MEMBERS_ARRAY(count_cash_checkbox, member_count) = checked
            If memb_age > 18 then       'Adding to the cash count
                adult_cash_count = adult_cash_count + 1
            Else
                child_cash_count = child_cash_count + 1
            End If
        End If
        If SNAP_checkbox = checked Then             'If SNAP is selected
            ALL_MEMBERS_ARRAY(include_snap_checkbox, member_count) = checked    'default to having the counted boxes checked for SNAP
            ALL_MEMBERS_ARRAY(count_snap_checkbox, member_count) = checked
            If memb_age > 21 then       'adding to the snap household member count
                adult_snap_count = adult_snap_count + 1
            Else
                child_snap_count = child_snap_count + 1
            End If
        End If
        If EMER_checkbox = checked Then             'If EMER is selected
            ALL_MEMBERS_ARRAY(include_emer_checkbox, member_count) = checked    'default to having the counted boxes checked for EMER
            ALL_MEMBERS_ARRAY(count_emer_checkbox, member_count) = checked
            If memb_age > 18 then       'Adding to the EMER count
                adult_emer_count = adult_emer_count + 1
            Else
                child_emer_count = child_emer_count + 1
            End If
        End If

		client_string = ref_nbr & last_name & first_name & mid_initial            'creating an array of all of the clients
		client_array = client_array & client_string & "|"
		transmit      'Going to the next MEMB panel
		Emreadscreen edit_check, 7, 24, 2 'looking to see if we are at the last member
        member_count = member_count + 1
	LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

    Call navigate_to_MAXIS_screen("STAT", "PARE")       'Going to get relationship information from PARE
    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2) 'looping through each member
        EMWriteScreen ALL_MEMBERS_ARRAY(memb_numb, each_member), 20, 76     'Going to PARE for each member
        transmit

        EMReadScreen panel_check, 14, 24, 13        'Making sure there is a PARE panel to read from
        If panel_check <> "DOES NOT EXIST" Then
            pare_row = 8                            'start of information on PARE
            Do
                EMReadScreen child_ref_nbr, 2, pare_row, 24     'Reading child, relationship and verif
                EMReadScreen rela_type, 1, pare_row, 53
                EMReadScreen rela_verif, 2, pare_row, 71
                If child_ref_nbr = "__" then exit do

                If rela_type = "1" then relationship_type = "Parent"            'Changing the code for the relationship to the words are used instead of code.
                If rela_type = "2" then relationship_type = "Stepparent"
                If rela_type = "3" then relationship_type = "Grandparent"
                If rela_type = "4" then relationship_type = "Relative Caregiver"
                If rela_type = "5" then relationship_type = "Foster parent"
                If rela_type = "6" then relationship_type = "Caregiver"
                If rela_type = "7" then relationship_type = "Guardian"
                If rela_type = "8" then relationship_type = "Relative"

                If rela_verif = "BC" Then relationship_verif = "Birth Certificate"      'Change the code for verif to full words for readability
                If rela_verif = "AR" Then relationship_verif = "Adoption Records"
                If rela_verif = "LG" Then relationship_verif = "Legal Guardian"
                If rela_verif = "RE" Then relationship_verif = "Religious Records"
                If rela_verif = "HR" Then relationship_verif = "Hospital Records"
                If rela_verif = "RP" Then relationship_verif = "Recognition of Parantage"
                If rela_verif = "OT" Then relationship_verif = "Other"
                If rela_verif = "NO" Then relationship_verif = "NONE"

                'Here is where the relationship information is added to the field of the dialog
                If child_ref_nbr <> "__" Then relationship_detail = relationship_detail & "Memb " & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " is the " & relationship_type & " of Memb " & child_ref_nbr & " - Verif: " & relationship_verif & "; "
                pare_row = pare_row + 1 'going to the next rwo
                If pare_row = 18 then
                    PF20 'shift PF8
                    EmReadscreen last_screen, 21, 24, 2
                    If last_screen = "THIS IS THE LAST PAGE" then
                        exit do
                    Else
                        pare_row = 8
                    End if
                End if
            Loop
        End If
    Next

    client_array = TRIM(client_array)
    client_array = split(client_array, "|")
    If SNAP_checkbox = checked then call read_EATS_panel        'If SNAP, we need to read EATS. This is a local function.

    Do
        Do
            err_msg = ""
            adult_cash_count = adult_cash_count & ""            'Setting variables to be strings
            child_cash_count = child_cash_count & ""
            adult_snap_count = adult_snap_count & ""
            child_snap_count = child_snap_count & ""
            adult_emer_count = adult_emer_count & ""
            child_emer_count = child_emer_count & ""

            'Dialog of the Household Composition
            dlg_len = 115 + (15 * UBound(ALL_MEMBERS_ARRAY, 2))     'setting the size of the dialog based on the number of household members
            if dlg_len < 145 Then dlg_len = 145                     'This is the minimum height of the dialog
            BeginDialog Dialog1, 0, 0, 446, dlg_len, "HH Composition Dialog"
              Text 10, 10, 250, 10, "This dialog will clarify the household relationships and details for the case."
              Text 105, 25, 100, 10, "Included and Counted in Grant"
              x_pos = 110
              count_cash_pos = x_pos + 5
              If cash_checkbox = checked Then
                Text x_pos, 40, 20, 10, "Cash"
                x_pos = x_pos + 35
              End If
              count_snap_pos = x_pos + 5
              If SNAP_checkbox = checked Then
                Text x_pos, 40, 20, 10, "SNAP"
                x_pos = x_pos + 35
              End If
              count_emer_pos = x_pos + 5
              If EMER_checkbox = checked Then Text x_pos, 40, 20, 10, "EMER"
              Text 230, 25, 90, 10, "Income Counted - Deeming"
              x_pos = 230
              deem_cash_pos = x_pos + 5
              If cash_checkbox = checked Then
                Text x_pos, 40, 20, 10, "Cash"
                x_pos = x_pos + 35
              End If
              deem_snap_pos = x_pos + 5
              If SNAP_checkbox = checked Then
                Text x_pos, 40, 20, 10, "SNAP"
                x_pos = x_pos + 35
              End If
              deem_emer_pos = x_pos + 5
              If EMER_checkbox = checked Then Text x_pos, 40, 20, 10, "EMER"
              GroupBox 330, 5, 105, 120, "HH Count by program"
              Text 335, 15, 100, 20, "Enter the number of adults and children for each program"
              Text 370, 35, 20, 10, "Adults"
              Text 400, 35, 30, 10, "Children"
              hh_comp_pos = 45
              If cash_checkbox = checked Then
                  Text 345, hh_comp_pos + 5, 20, 10, "Cash"
                  EditBox 370, hh_comp_pos, 20, 15, adult_cash_count
                  EditBox 405, hh_comp_pos, 20, 15, child_cash_count
                  CheckBox 355, hh_comp_pos + 20, 75, 10, "Pregnant Caregiver", pregnant_caregiver_checkbox
                  hh_comp_pos = hh_comp_pos + 35
              End If
              If SNAP_checkbox = checked Then
                  Text 345, hh_comp_pos + 5, 20, 10, "SNAP"
                  EditBox 370, hh_comp_pos, 20, 15, adult_snap_count
                  EditBox 405, hh_comp_pos, 20, 15, child_snap_count
                  hh_comp_pos = hh_comp_pos + 20
              End If
              If EMER_checkbox = checked then
                  Text 345, hh_comp_pos, 25, 10, "EMER"
                  EditBox 370, hh_comp_pos, 20, 15, adult_emer_count
                  EditBox 405, hh_comp_pos, 20, 15, child_emer_count
              End If
              y_pos = 55
              For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                  Text 10, y_pos, 100, 10, ALL_MEMBERS_ARRAY(clt_name, each_member)
                  If cash_checkbox = checked Then CheckBox count_cash_pos, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(include_cash_checkbox, each_member)
                  If SNAP_checkbox = checked Then CheckBox count_snap_pos, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(include_snap_checkbox, each_member)
                  If EMER_checkbox = checked then CheckBox count_emer_pos, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(include_emer_checkbox, each_member)
                  If cash_checkbox = checked Then CheckBox deem_cash_pos, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(count_cash_checkbox, each_member)
                  If SNAP_checkbox = checked Then CheckBox deem_snap_pos, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(count_snap_checkbox, each_member)
                  If EMER_checkbox = checked then CheckBox deem_emer_pos, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(count_emer_checkbox, each_member)
                  y_pos = y_pos + 15
              Next
              if y_pos < 100 Then Y_pos = 100
              y_pos = y_pos + 5
              Text 10, y_pos + 5, 25, 10, "EATS:"
              EditBox 35, y_pos, 290, 15, EATS
              Text 10, y_pos + 25, 90, 10, "Household Relationships:"
              EditBox 105, y_pos + 20, 220, 15, relationship_detail
              ButtonGroup ButtonPressed
                OkButton 335, y_pos + 20, 50, 15
                CancelButton 390, y_pos + 20, 50, 15
            EndDialog

            dialog Dialog1
            cancel_without_confirmation

            If trim(adult_cash_count) = "" Then adult_cash_count = 0            ''
            If trim(child_cash_count) = "" Then child_cash_count = 0
            If IsNumeric(adult_cash_count) = False and IsNumeric(child_cash_count) = False Then err_msg = err_msg & vbNewLine & "* Enter a valid count for the number of adults and children for the Cash program."

            If trim(adult_snap_count) = "" Then adult_snap_count = 0
            If trim(child_snap_count) = "" Then child_snap_count = 0
            If IsNumeric(adult_snap_count) = False and IsNumeric(child_snap_count) = False Then err_msg = err_msg & vbNewLine & "* Enter a valid count for the number of adults and children for the SNAP program."
            If SNAP_checkbox = checked AND trim(EATS) = "" Then err_msg = err_msg & vbNewLine & "* Clarify who purchases and prepares together since SNAP is being considered."

            If trim(adult_emer_count) = "" Then adult_emer_count = 0
            If trim(child_emer_count) = "" Then child_emer_count = 0
            If IsNumeric(adult_emer_count) = False and IsNumeric(child_emer_count) = False Then err_msg = err_msg & vbNewLine & "* Enter a valid count for the number of adults and children for the EMER program."

            adult_cash_count = adult_cash_count * 1
            child_cash_count = child_cash_count * 1
            adult_snap_count = adult_snap_count * 1
            child_snap_count = child_snap_count * 1
            adult_emer_count = adult_emer_count * 1
            child_emer_count = child_emer_count * 1

            If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
        Loop until err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE

    HH_member_array = ""

    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        ALL_MEMBERS_ARRAY(gather_detail, each_member) = FALSE
        If ALL_MEMBERS_ARRAY(include_cash_checkbox, each_member) = checked Then
            HH_member_array = HH_member_array & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " "
            ALL_MEMBERS_ARRAY(gather_detail, each_member) = TRUE
        ElseIf ALL_MEMBERS_ARRAY(include_snap_checkbox, each_member) = checked Then
            HH_member_array = HH_member_array & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " "
            ALL_MEMBERS_ARRAY(gather_detail, each_member) = TRUE
        ElseIf ALL_MEMBERS_ARRAY(include_emer_checkbox, each_member) = checked Then
            HH_member_array = HH_member_array & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " "
            ALL_MEMBERS_ARRAY(gather_detail, each_member) = TRUE
        ElseIf ALL_MEMBERS_ARRAY(count_cash_checkbox, each_member) = checked Then
            HH_member_array = HH_member_array & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " "
            ALL_MEMBERS_ARRAY(gather_detail, each_member) = TRUE
        ElseIf ALL_MEMBERS_ARRAY(count_snap_checkbox, each_member) = checked Then
            HH_member_array = HH_member_array & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " "
            ALL_MEMBERS_ARRAY(gather_detail, each_member) = TRUE
        ElseIf ALL_MEMBERS_ARRAY(count_emer_checkbox, each_member) = checked Then
            HH_member_array = HH_member_array & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " "
            ALL_MEMBERS_ARRAY(gather_detail, each_member) = TRUE
        End If
    Next

	HH_member_array = TRIM(HH_member_array)							'Cleaning up array for ease of use.
    HH_member_array = REPLACE(HH_member_array, "  ", " ")
	HH_member_array = SPLIT(HH_member_array, " ")
    ' MsgBox "All members ubound - " & UBound(ALL_MEMBERS_ARRAY, 2)
end function

function read_ADDR_panel()
    Call navigate_to_MAXIS_screen("STAT", "ADDR")

    EMReadScreen line_one, 22, 6, 43
    EMReadScreen line_two, 22, 7, 43
    EMReadScreen city_line, 15, 8, 43
    EMReadScreen state_line, 2, 8, 66
    EMReadScreen zip_line, 7, 9, 43
    EMReadScreen county_line, 2, 9, 66
    EMReadScreen verif_line, 2, 9, 74
    EMReadScreen homeless_line, 1, 10, 43
    EMReadScreen reservation_line, 1, 10, 74
    EMReadScreen living_sit_line, 2, 11, 43

    addr_line_one = replace(line_one, "_", "")
    addr_line_two = replace(line_two, "_", "")
    city = replace(city_line, "_", "")
    state = state_line
    zip = replace(zip_line, "_", "")

    If county_line = "01" Then addr_county = "01 Aitkin"
    If county_line = "02" Then addr_county = "02 Anoka"
    If county_line = "03" Then addr_county = "03 Becker"
    If county_line = "04" Then addr_county = "04 Beltrami"
    If county_line = "05" Then addr_county = "05 Benton"
    If county_line = "06" Then addr_county = "06 Big Stone"
    If county_line = "07" Then addr_county = "07 Blue Earth"
    If county_line = "08" Then addr_county = "08 Brown"
    If county_line = "09" Then addr_county = "09 Carlton"
    If county_line = "10" Then addr_county = "10 Carver"
    If county_line = "11" Then addr_county = "11 Cass"
    If county_line = "12" Then addr_county = "12 Chippewa"
    If county_line = "13" Then addr_county = "13 Chisago"
    If county_line = "14" Then addr_county = "14 Clay"
    If county_line = "15" Then addr_county = "15 Clearwater"
    If county_line = "16" Then addr_county = "16 Cook"
    If county_line = "17" Then addr_county = "17 Cottonwood"
    If county_line = "18" Then addr_county = "18 Crow Wing"
    If county_line = "19" Then addr_county = "19 Dakota"
    If county_line = "20" Then addr_county = "20 Dodge"
    If county_line = "21" Then addr_county = "21 Douglas"
    If county_line = "22" Then addr_county = "22 Faribault"
    If county_line = "23" Then addr_county = "23 Fillmore"
    If county_line = "24" Then addr_county = "24 Freeborn"
    If county_line = "25" Then addr_county = "25 Goodhue"
    If county_line = "26" Then addr_county = "26 Grant"
    If county_line = "27" Then addr_county = "27 Hennepin"
    If county_line = "28" Then addr_county = "28 Houston"
    If county_line = "29" Then addr_county = "29 Hubbard"
    If county_line = "30" Then addr_county = "30 Isanti"
    If county_line = "31" Then addr_county = "31 Itasca"
    If county_line = "32" Then addr_county = "32 Jackson"
    If county_line = "33" Then addr_county = "33 Kanabec"
    If county_line = "34" Then addr_county = "34 Kandiyohi"
    If county_line = "35" Then addr_county = "35 Kittson"
    If county_line = "36" Then addr_county = "36 Koochiching"
    If county_line = "37" Then addr_county = "37 Lac Qui Parle"
    If county_line = "38" Then addr_county = "38 Lake"
    If county_line = "39" Then addr_county = "39 Lake Of Woods"
    If county_line = "40" Then addr_county = "40 Le Sueur"
    If county_line = "41" Then addr_county = "41 Lincoln"
    If county_line = "42" Then addr_county = "42 Lyon"
    If county_line = "43" Then addr_county = "43 Mcleod"
    If county_line = "44" Then addr_county = "44 Mahnomen"
    If county_line = "45" Then addr_county = "45 Marshall"
    If county_line = "46" Then addr_county = "46 Martin"
    If county_line = "47" Then addr_county = "47 Meeker"
    If county_line = "48" Then addr_county = "48 Mille Lacs"
    If county_line = "49" Then addr_county = "49 Morrison"
    If county_line = "50" Then addr_county = "50 Mower"
    If county_line = "51" Then addr_county = "51 Murray"
    If county_line = "52" Then addr_county = "52 Nicollet"
    If county_line = "53" Then addr_county = "53 Nobles"
    If county_line = "54" Then addr_county = "54 Norman"
    If county_line = "55" Then addr_county = "55 Olmsted"
    If county_line = "56" Then addr_county = "56 Otter Tail"
    If county_line = "57" Then addr_county = "57 Pennington"
    If county_line = "58" Then addr_county = "58 Pine"
    If county_line = "59" Then addr_county = "59 Pipestone"
    If county_line = "60" Then addr_county = "60 Polk"
    If county_line = "61" Then addr_county = "61 Pope"
    If county_line = "62" Then addr_county = "62 Ramsey"
    If county_line = "63" Then addr_county = "63 Red Lake"
    If county_line = "64" Then addr_county = "64 Redwood"
    If county_line = "65" Then addr_county = "65 Renville"
    If county_line = "66" Then addr_county = "66 Rice"
    If county_line = "67" Then addr_county = "67 Rock"
    If county_line = "68" Then addr_county = "68 Roseau"
    If county_line = "69" Then addr_county = "69 St. Louis"
    If county_line = "70" Then addr_county = "70 Scott"
    If county_line = "71" Then addr_county = "71 Sherburne"
    If county_line = "72" Then addr_county = "72 Sibley"
    If county_line = "73" Then addr_county = "73 Stearns"
    If county_line = "74" Then addr_county = "74 Steele"
    If county_line = "75" Then addr_county = "75 Stevens"
    If county_line = "76" Then addr_county = "76 Swift"
    If county_line = "77" Then addr_county = "77 Todd"
    If county_line = "78" Then addr_county = "78 Traverse"
    If county_line = "79" Then addr_county = "79 Wabasha"
    If county_line = "80" Then addr_county = "80 Wadena"
    If county_line = "81" Then addr_county = "81 Waseca"
    If county_line = "82" Then addr_county = "82 Washington"
    If county_line = "83" Then addr_county = "83 Watonwan"
    If county_line = "84" Then addr_county = "84 Wilkin"
    If county_line = "85" Then addr_county = "85 Winona"
    If county_line = "86" Then addr_county = "86 Wright"
    If county_line = "87" Then addr_county = "87 Yellow Medicine"
    If county_line = "89" Then addr_county = "89 Out-of-State"

    If homeless_line = "Y" Then homeless_yn = "Yes"
    If homeless_line = "N" Then homeless_yn = "No"
    If reservation_line = "Y" Then reservation_yn = "Yes"
    If reservation_line = "N" Then reservation_yn = "No"

    If verif_line = "SF" Then addr_verif = "SF - Shelter Form"
    If verif_line = "Co" Then addr_verif = "CO - Coltrl Stmt"
    If verif_line = "MO" Then addr_verif = "MO - Mortgage Papers"
    If verif_line = "TX" Then addr_verif = "TX - Prop Tax Stmt"
    If verif_line = "CD" Then addr_verif = "CD - Contrct for Deed"
    If verif_line = "UT" Then addr_verif = "UT - Utility Stmt"
    If verif_line = "DL" Then addr_verif = "DL - Driver Lic/State ID"
    If verif_line = "OT" Then addr_verif = "OT - Other Document"
    If verif_line = "NO" Then addr_verif = "NO - No Ver Prvd"
    If verif_line = "?_" Then addr_verif = "? - Delayed"
    If verif_line = "__" Then addr_verif = "Blank"


    If living_sit_line = "__" Then living_situation = "Blank"
    If living_sit_line = "01" Then living_situation = "01 - Own home, lease or roommate"
    If living_sit_line = "02" Then living_situation = "02 - Family/Friends - economic hardship"
    If living_sit_line = "03" Then living_situation = "03 -  servc prvdr- foster/group home"
    If living_sit_line = "04" Then living_situation = "04 - Hospital/Treatment/Detox/Nursing Home"
    If living_sit_line = "05" Then living_situation = "05 - Jail/Prison//Juvenile Det."
    If living_sit_line = "06" Then living_situation = "06 - Hotel/Motel"
    If living_sit_line = "07" Then living_situation = "07 - Emergency Shelter"
    If living_sit_line = "08" Then living_situation = "08 - Place not meant for Housing"
    If living_sit_line = "09" Then living_situation = "09 - Declined"
    If living_sit_line = "10" Then living_situation = "10 - Unknown"

    EMReadScreen addr_eff_date, 8, 4, 43
    EMReadScreen addr_future_date, 8, 4, 66
    EMReadScreen mail_line_one, 22, 13, 43
    EMReadScreen mail_line_two, 22, 14, 43
    EMReadScreen mail_city_line, 15, 15, 43
    EMReadScreen mail_state_line, 2, 16, 43
    EMReadScreen mail_zip_line, 7, 16, 52

    addr_eff_date = replace(addr_eff_date, " ", "/")
    addr_future_date = trim(addr_future_date)
    addr_future_date = replace(addr_future_date, " ", "/")
    mail_line_one = replace(mail_line_one, "_", "")
    mail_line_two = replace(mail_line_two, "_", "")
    mail_city_line = replace(mail_city_line, "_", "")
    mail_state_line = replace(mail_state_line, "_", "")
    mail_zip_line = replace(mail_zip_line, "_", "")

    notes_on_address = "Address effective: " & addr_eff_date & "."
    If mail_line_one <> "" Then
        If mail_line_two = "" Then notes_on_address = notes_on_address & " Mailing address: " & mail_line_one & " " & mail_city_line & ", " & mail_state_line & " " & mail_zip_line
        If mail_line_two <> "" Then notes_on_address = notes_on_address & " Mailing address: " & mail_line_one & " " & mail_line_two & " " & mail_city_line & ", " & mail_state_line & " " & mail_zip_line
    End If
    If addr_future_date <> "" Then notes_on_address = notes_on_address & "; ** Address will update effective " & addr_future_date & "."

end function

function read_BUSI_panel()
    EMReadScreen income_type, 2, 5, 37
    EMReadScreen retro_rpt_hrs, 3, 13, 60
    EMReadScreen prosp_rpt_hrs, 3, 13, 74
    EMReadScreen retro_min_wg_hrs, 3, 14, 60
    EMReadScreen prosp_min_wg_hrs, 3, 14, 74
    EMReadScreen self_emp_method, 2, 16, 53
    EMReadScreen method_date, 8, 16, 63
    EMReadScreen income_start, 8, 5, 55
    EMReadScreen income_end, 8, 5, 72

    If income_type = "01" Then ALL_BUSI_PANELS_ARRAY(busi_type, busi_count) = "01 - Farming"
    If income_type = "02" Then ALL_BUSI_PANELS_ARRAY(busi_type, busi_count) = "02 - Real Estate"
    If income_type = "03" Then ALL_BUSI_PANELS_ARRAY(busi_type, busi_count) = "03 - Home Product Sales"
    If income_type = "04" Then ALL_BUSI_PANELS_ARRAY(busi_type, busi_count) = "04 - Other Sales"
    If income_type = "05" Then ALL_BUSI_PANELS_ARRAY(busi_type, busi_count) = "05 - Personal Services"
    If income_type = "06" Then ALL_BUSI_PANELS_ARRAY(busi_type, busi_count) = "06 - Paper Route"
    If income_type = "07" Then ALL_BUSI_PANELS_ARRAY(busi_type, busi_count) = "07 - In Home Daycare"
    If income_type = "08" Then ALL_BUSI_PANELS_ARRAY(busi_type, busi_count) = "08 - Rental Income"
    If income_type = "09" Then ALL_BUSI_PANELS_ARRAY(busi_type, busi_count) = "09 - Other"
    ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, busi_count) = trim(retro_rpt_hrs)
    ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, busi_count) = trim(prosp_rpt_hrs)
    ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, busi_count) = trim(retro_min_wg_hrs)
    ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, busi_count) = trim(prosp_min_wg_hrs)
    If self_emp_method = "01" Then ALL_BUSI_PANELS_ARRAY(calc_method, busi_count) = "50% Gross Inc"
    If self_emp_method = "02" Then ALL_BUSI_PANELS_ARRAY(calc_method, busi_count) = "Tax Forms"
    ALL_BUSI_PANELS_ARRAY(mthd_date, busi_count) = replace(method_date, " ", "/")
    If ALL_BUSI_PANELS_ARRAY(mthd_date, busi_count) = "__/__/__" Then ALL_BUSI_PANELS_ARRAY(mthd_date, busi_count) = ""
    ALL_BUSI_PANELS_ARRAY(start_date, busi_count) = replace(income_start, " ", "/")
    ALL_BUSI_PANELS_ARRAY(end_date, busi_count) = replace(income_end, " ", "/")
    If ALL_BUSI_PANELS_ARRAY(end_date, busi_count) = "__/__/__" Then ALL_BUSI_PANELS_ARRAY(end_date, busi_count) = ""

    EMWriteScreen "X", 6, 26
    transmit

    EMReadScreen retro_cash_inc, 8, 9, 43
    EMReadScreen prosp_cash_inc, 8, 9, 59
    EMReadScreen cash_inc_verif, 1, 9, 73
    EMReadScreen retro_cash_exp, 8, 15, 43
    EMReadScreen prosp_cash_exp, 8, 15, 59
    EMReadScreen cash_exp_verif, 1, 15, 73
    ALL_BUSI_PANELS_ARRAY(income_ret_cash, busi_count) = trim(retro_cash_inc)
    If ALL_BUSI_PANELS_ARRAY(income_ret_cash, busi_count) = "________" Then ALL_BUSI_PANELS_ARRAY(income_ret_cash, busi_count) = "0"
    ALL_BUSI_PANELS_ARRAY(income_pro_cash, busi_count) = trim(prosp_cash_inc)
    If ALL_BUSI_PANELS_ARRAY(income_pro_cash, busi_count) = "________" Then ALL_BUSI_PANELS_ARRAY(income_pro_cash, busi_count) = "0"
    If cash_inc_verif = "1" Then ALL_BUSI_PANELS_ARRAY(cash_income_verif, busi_count) = "Income Tax Returns"
    If cash_inc_verif = "2" Then ALL_BUSI_PANELS_ARRAY(cash_income_verif, busi_count) = "Receipts of Sales/Purch"
    If cash_inc_verif = "3" Then ALL_BUSI_PANELS_ARRAY(cash_income_verif, busi_count) = "Busi Records/Ledger"
    If cash_inc_verif = "6" Then ALL_BUSI_PANELS_ARRAY(cash_income_verif, busi_count) = "Other Document"
    If cash_inc_verif = "N" Then ALL_BUSI_PANELS_ARRAY(cash_income_verif, busi_count) = "No Verif Provided"
    If cash_inc_verif = "?" Then ALL_BUSI_PANELS_ARRAY(cash_income_verif, busi_count) = "Delayed Verif"
    If cash_inc_verif = "_" Then ALL_BUSI_PANELS_ARRAY(cash_income_verif, busi_count) = "Blank"
    ALL_BUSI_PANELS_ARRAY(expense_ret_cash, busi_count) = trim(retro_cash_exp)
    If ALL_BUSI_PANELS_ARRAY(expense_ret_cash, busi_count) = "________" Then ALL_BUSI_PANELS_ARRAY(expense_ret_cash, busi_count) = "0"
    ALL_BUSI_PANELS_ARRAY(expense_pro_cash, busi_count) = trim(prosp_cash_exp)
    If ALL_BUSI_PANELS_ARRAY(expense_pro_cash, busi_count) = "________" Then ALL_BUSI_PANELS_ARRAY(expense_pro_cash, busi_count) = "0"
    If cash_exp_verif = "1" Then ALL_BUSI_PANELS_ARRAY(cash_expense_verif, busi_count) = "Income Tax Returns"
    If cash_exp_verif = "2" Then ALL_BUSI_PANELS_ARRAY(cash_expense_verif, busi_count) = "Receipts of Sales/Purch"
    If cash_exp_verif = "3" Then ALL_BUSI_PANELS_ARRAY(cash_expense_verif, busi_count) = "Busi Records/Ledger"
    If cash_exp_verif = "6" Then ALL_BUSI_PANELS_ARRAY(cash_expense_verif, busi_count) = "Other Document"
    If cash_exp_verif = "N" Then ALL_BUSI_PANELS_ARRAY(cash_expense_verif, busi_count) = "No Verif Provided"
    If cash_exp_verif = "?" Then ALL_BUSI_PANELS_ARRAY(cash_expense_verif, busi_count) = "Delayed Verif"
    If cash_exp_verif = "_" Then ALL_BUSI_PANELS_ARRAY(cash_expense_verif, busi_count) = "Blank"

    EMReadScreen prosp_ive_inc, 8, 10, 59
    EMReadScreen ive_inc_verif, 1, 10, 73
    EMReadScreen prosp_ive_exp, 8, 16, 59
    EMReadScreen ive_exp_verif, 1, 16, 73

    EMReadScreen retro_snap_inc, 8, 11, 43
    EMReadScreen prosp_snap_inc, 8, 11, 59
    EMReadScreen snap_inc_verif, 1, 11, 73
    EMReadScreen retro_snap_exp, 8, 17, 43
    EMReadScreen prosp_snap_exp, 8, 17, 59
    EMReadScreen snap_exp_verif, 1, 17, 73
    ALL_BUSI_PANELS_ARRAY(income_ret_snap, busi_count) = trim(retro_snap_inc)
    If ALL_BUSI_PANELS_ARRAY(income_ret_snap, busi_count) = "________" Then ALL_BUSI_PANELS_ARRAY(income_ret_snap, busi_count) = "0"
    ALL_BUSI_PANELS_ARRAY(income_pro_snap, busi_count) = trim(prosp_snap_inc)
    If ALL_BUSI_PANELS_ARRAY(income_pro_snap, busi_count) = "________" Then ALL_BUSI_PANELS_ARRAY(income_pro_snap, busi_count) = "0"
    If snap_inc_verif = "1" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, busi_count) = "Income Tax Returns"
    If snap_inc_verif = "2" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, busi_count) = "Receipts of Sales/Purch"
    If snap_inc_verif = "3" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, busi_count) = "Busi Records/Ledger"
    If snap_inc_verif = "4" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, busi_count) = "Pend Out State Verif"
    If snap_inc_verif = "6" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, busi_count) = "Other Document"
    If snap_inc_verif = "N" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, busi_count) = "No Verif Provided"
    If snap_inc_verif = "?" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, busi_count) = "Delayed Verif"
    If snap_inc_verif = "_" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, busi_count) = "Blank"

    ALL_BUSI_PANELS_ARRAY(expense_ret_snap, busi_count) = trim(retro_snap_exp)
    If ALL_BUSI_PANELS_ARRAY(expense_ret_snap, busi_count) = "________" Then ALL_BUSI_PANELS_ARRAY(expense_ret_snap, busi_count) = "0"
    ALL_BUSI_PANELS_ARRAY(expense_pro_snap, busi_count) = trim(prosp_snap_exp)
    If ALL_BUSI_PANELS_ARRAY(expense_pro_snap, busi_count) = "________" Then ALL_BUSI_PANELS_ARRAY(expense_pro_snap, busi_count) = "0"
    If snap_exp_verif = "1" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, busi_count) = "Income Tax Returns"
    If snap_exp_verif = "2" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, busi_count) = "Receipts of Sales/Purch"
    If snap_exp_verif = "3" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, busi_count) = "Busi Records/Ledger"
    If snap_exp_verif = "4" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, busi_count) = "Pend Out State Verif"
    If snap_exp_verif = "6" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, busi_count) = "Other Document"
    If snap_exp_verif = "N" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, busi_count) = "No Verif Provided"
    If snap_exp_verif = "?" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, busi_count) = "Delayed Verif"
    If snap_exp_verif = "_" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, busi_count) = "Blank"

    EMReadScreen prosp_hca_inc, 8, 12, 59
    EMReadScreen hca_inc_verif, 1, 12, 73
    EMReadScreen prosp_hca_exp, 8, 18, 59
    EMReadScreen hca_exp_verif, 1, 18, 73

    EMReadScreen prosp_hcb_inc, 8, 13, 59
    EMReadScreen hcb_inc_verif, 1, 13, 73
    EMReadScreen prosp_hcb_exp, 8, 19, 59
    EMReadScreen hcb_exp_verif, 1, 19, 73

    ALL_BUSI_PANELS_ARRAY(budget_explain, busi_count) = ""
    PF3

end function

function read_EATS_panel()
    call navigate_to_MAXIS_screen("stat", "eats")

    'Now it checks for the total number of panels. If there's 0 Of 0 it'll exit the function for you so as to save oodles of time.
    EMReadScreen panel_total_check, 6, 2, 73
    IF panel_total_check = "0 Of 0" THEN
        If UBound(ALL_MEMBERS_ARRAY, 2) = 0 Then EATS = "Single member case, EATS panel is not needed,"
        exit function		'Exits out if there's no panel info
    End If
    EMReadScreen all_eat_together, 1, 4, 72
    If all_eat_together = "Y" Then
        EATS = "All clients on this case purchase and prepare food together."
    Else
        EATS = "SNAP unit p/p sep from memb(s):"
        EMReadScreen group_one, 40, 13, 39
        EMReadScreen group_two, 40, 14, 39
        EMReadScreen group_three, 40, 15, 39
        EMReadScreen group_four, 40, 16, 39
        EMReadScreen group_five, 40, 17, 39

        group_one = replace(group_one, "__", "")
        group_two = replace(group_two, "__", "")
        group_three = replace(group_three, "__", "")
        group_four = replace(group_four, "__", "")
        group_five = replace(group_five, "__", "")

        group_one = trim(group_one)
        group_two = trim(group_two)
        group_three = trim(group_three)
        group_four = trim(group_four)
        group_five = trim(group_five)

        If group_one <> "" Then
            EMReadScreen group_one_no, 2, 13, 28
            group_one = replace(group_one, "  ", ", ")
            EATS = EATS & "Eating group " & group_one_no & " with memb(s) " & group_one
        End If
        If group_two <> "" Then
            EMReadScreen group_two_no, 2, 13, 28
            group_two = replace(group_two, "  ", ", ")
            EATS = EATS & "; Eating group " & group_two_no & " with memb(s) " & group_two
        End If
        If group_three <> "" Then
            EMReadScreen group_three_no, 2, 13, 28
            group_three = replace(group_three, "  ", ", ")
            EATS = EATS & "; Eating group " & group_three_no & " with memb(s) " & group_three
        End If
        If group_four <> "" Then
            EMReadScreen group_four_no, 2, 13, 28
            group_four = replace(group_four, "  ", ", ")
            EATS = EATS & "; Eating group " & group_four_no & " with memb(s) " & group_four
        End If
        If group_five <> "" Then
            EMReadScreen group_five_no, 2, 13, 28
            group_five = replace(group_five, "  ", ", ")
            EATS = EATS & "; Eating group " & group_five_no & " with memb(s) " & group_five
        End If

    End If
end function

function read_HEST_panel()
    hest_information = ""
    call navigate_to_MAXIS_screen("stat", "HEST")

    'Now it checks for the total number of panels. If there's 0 Of 0 it'll exit the function for you so as to save oodles of time.
    EMReadScreen panel_total_check, 6, 2, 73
    IF panel_total_check = "0 Of 0" THEN
        hest_information = "NONE - $0"
    ELSE
        EMReadScreen prosp_heat_air, 1, 13, 60
        EMReadScreen prosp_electric, 1, 14, 60
        EMReadScreen prosp_phone, 1, 15, 60
        CAF_datestamp = cdate(CAF_datestamp)

        If CAF_datestamp >= cdate("10/01/2020") then
            If prosp_heat_air = "Y" Then
                hest_information = "AC/Heat - Full $496"
            ElseIf prosp_electric = "Y" Then
                If prosp_phone = "Y" Then
                    hest_information = "Electric and Phone - $210"
                Else
                    hest_information = "Electric ONLY - $154"
                End If
            ElseIf prosp_phone = "Y" Then
                hest_information = "Phone ONLY - $56"
            End If
        Else
            If prosp_heat_air = "Y" Then
                hest_information = "AC/Heat - Full $490"
            ElseIf prosp_electric = "Y" Then
                If prosp_phone = "Y" Then
                    hest_information = "Electric and Phone - $192"
                Else
                    hest_information = "Electric ONLY - $143"
                End If
            ElseIf prosp_phone = "Y" Then
                hest_information = "Phone ONLY - $49"
            End If
        End If
        CAF_datestamp = CAF_datestamp & ""
    END IF
end function

function read_JOBS_panel()
    EMReadScreen JOBS_month, 5, 20, 55									'reads Footer month
    JOBS_month = replace(JOBS_month, " ", "/")					'Cleans up the read number by putting a / in place of the blank space between MM YY
    EMReadScreen JOBS_type, 30, 7, 42										'Reads up name of the employer and then cleans it up
    JOBS_type = replace(JOBS_type, "_", ""	)
    JOBS_type = trim(JOBS_type)
    JOBS_type = split(JOBS_type)
    For each JOBS_part in JOBS_type											'Correcting case on the name of the employer as it reads in all CAPS
        If JOBS_part <> "" then
            first_letter = ucase(left(JOBS_part, 1))
            other_letters = LCase(right(JOBS_part, len(JOBS_part) -1))
            new_JOBS_type = new_JOBS_type & first_letter & other_letters & " "
        End if
    Next
    ALL_JOBS_PANELS_ARRAY(employer_name, job_count) = new_JOBS_type
    EMReadScreen jobs_hourly_wage, 6, 6, 75   'reading hourly wage field
    ALL_JOBS_PANELS_ARRAY(hrly_wage, job_count) = replace(jobs_hourly_wage, "_", "")   'trimming any underscores

    ' Navigates to the FS PIC
    EMWriteScreen "x", 19, 38
    transmit
    EMReadScreen SNAP_JOBS_amt, 8, 17, 56
    ALL_JOBS_PANELS_ARRAY(pic_pay_date_income, job_count) = trim(SNAP_JOBS_amt)
    EMReadScreen jobs_SNAP_prospective_amt, 8, 18, 56
    ALL_JOBS_PANELS_ARRAY(pic_prosp_income, job_count) = trim(jobs_SNAP_prospective_amt)  'prospective amount from PIC screen
    EMReadScreen snap_pay_frequency, 1, 5, 64
    ALL_JOBS_PANELS_ARRAY(pic_pay_freq, job_count) = snap_pay_frequency
    EMReadScreen date_of_pic_calc, 8, 5, 34
    ALL_JOBS_PANELS_ARRAY(pic_calc_date, job_count) = replace(date_of_pic_calc, " ", "/")
    transmit
    'Navigats to GRH PIC
    EMReadscreen GRH_PIC_check, 3, 19, 73 	'This must check to see if the GRH PIC is there or not. If fun on months 06/16 and before it will cause an error if it pf3s on the home panel.
    IF GRH_PIC_check = "GRH" THEN
    	EMWriteScreen "x", 19, 71
    	transmit
    	EMReadScreen GRH_JOBS_pay_amt, 8, 16, 69
    	GRH_JOBS_pay_amt = trim(GRH_JOBS_pay_amt)
        EMReadScreen GRH_JOBS_total_amt, 8, 17, 69
        GRH_JOBS_total_amt = trim(GRH_JOBS_total_amt)
    	EMReadScreen GRH_pay_frequency, 1, 3, 63
    	EMReadScreen GRH_date_of_pic_calc, 8, 3, 30
    	GRH_date_of_pic_calc = replace(GRH_date_of_pic_calc, " ", "/")
        ALL_JOBS_PANELS_ARRAY(grh_calc_date, job_count) = GRH_date_of_pic_calc
        ALL_JOBS_PANELS_ARRAY(grh_pay_freq, job_count) = GRH_pay_frequency
        ALL_JOBS_PANELS_ARRAY(grh_pay_day_income, job_count) = GRH_JOBS_pay_amt
        ALL_JOBS_PANELS_ARRAY(grh_prosp_income, job_count) = GRH_JOBS_total_amt
    	PF3
    END IF
    '  Reads the information on the retro side of JOBS
    EMReadScreen retro_JOBS_amt, 8, 17, 38
    EMReadScreen retro_JOBS_hrs, 3, 18, 43
    ALL_JOBS_PANELS_ARRAY(job_retro_income, job_count) = trim(retro_JOBS_amt)
    ALL_JOBS_PANELS_ARRAY(retro_hours, job_count) = trim(retro_JOBS_hrs)

    '  Reads the information on the prospective side of JOBS
    EMReadScreen prospective_JOBS_amt, 8, 17, 67
    EMReadScreen prosp_JOBS_hrs, 3, 18, 72
    ALL_JOBS_PANELS_ARRAY(job_prosp_income, job_count) = trim(prospective_JOBS_amt)
    ALL_JOBS_PANELS_ARRAY(prosp_hours, job_count) = trim(prosp_JOBS_hrs)

    '  Reads the information about health care off of HC Income Estimator
    EMReadScreen pay_frequency, 1, 18, 35
    ALL_JOBS_PANELS_ARRAY(main_pay_freq, job_count) = pay_frequency
    EMReadScreen HC_income_est_check, 3, 19, 63 'reading to find the HC income estimator is moving 6/1/16, to account for if it only affects future months we are reading to find the HC inc EST
    IF HC_income_est_check = "Est" Then 'this is the old position
    	EMWriteScreen "x", 19, 54
    ELSE								'this is the new position
    	EMWriteScreen "x", 19, 48
    END IF
    transmit
    EMReadScreen HC_JOBS_amt, 8, 11, 63
    HC_JOBS_amt = trim(HC_JOBS_amt)
    transmit

    EMReadScreen JOBS_ver, 25, 6, 36
    ALL_JOBS_PANELS_ARRAY(verif_code, job_count) = trim(JOBS_ver)
    If ALL_JOBS_PANELS_ARRAY(verif_code, job_count) = "" Then
        EMReadScreen JOBS_ver, 1, 6, 34
        If JOBS_ver = "?" Then ALL_JOBS_PANELS_ARRAY(verif_code, job_count) = "Delayed"
    End If
    EMReadScreen JOBS_income_end_date, 8, 9, 49
    'This now cleans up the variables converting codes read from the panel into words for the final variable to be used in the output.
    If JOBS_income_end_date <> "__ __ __" then JOBS_income_end_date = replace(JOBS_income_end_date, " ", "/")
    If IsDate(JOBS_income_end_date) = True then ALL_JOBS_PANELS_ARRAY(budget_explain, job_count) = "Income ended " & JOBS_income_end_date & ".; "

    If ALL_JOBS_PANELS_ARRAY(main_pay_freq, job_count) = "4" Then ALL_JOBS_PANELS_ARRAY(main_pay_freq, job_count) = "Weekly"
    If ALL_JOBS_PANELS_ARRAY(main_pay_freq, job_count) = "3" Then ALL_JOBS_PANELS_ARRAY(main_pay_freq, job_count) = "Biweekly"
    If ALL_JOBS_PANELS_ARRAY(main_pay_freq, job_count) = "2" Then ALL_JOBS_PANELS_ARRAY(main_pay_freq, job_count) = "Semi-Monthly"
    If ALL_JOBS_PANELS_ARRAY(main_pay_freq, job_count) = "1" Then ALL_JOBS_PANELS_ARRAY(main_pay_freq, job_count) = "Monthly"

    If ALL_JOBS_PANELS_ARRAY(pic_pay_freq, job_count) = "4" Then ALL_JOBS_PANELS_ARRAY(pic_pay_freq, job_count) = "Weekly"
    If ALL_JOBS_PANELS_ARRAY(pic_pay_freq, job_count) = "3" Then ALL_JOBS_PANELS_ARRAY(pic_pay_freq, job_count) = "Biweekly"
    If ALL_JOBS_PANELS_ARRAY(pic_pay_freq, job_count) = "2" Then ALL_JOBS_PANELS_ARRAY(pic_pay_freq, job_count) = "Semi-Monthly"
    If ALL_JOBS_PANELS_ARRAY(pic_pay_freq, job_count) = "1" Then ALL_JOBS_PANELS_ARRAY(pic_pay_freq, job_count) = "Monthly"

    If ALL_JOBS_PANELS_ARRAY(grh_pay_freq, job_count) = "4" Then ALL_JOBS_PANELS_ARRAY(grh_pay_freq, job_count) = "Weekly"
    If ALL_JOBS_PANELS_ARRAY(grh_pay_freq, job_count) = "3" Then ALL_JOBS_PANELS_ARRAY(grh_pay_freq, job_count) = "Biweekly"
    If ALL_JOBS_PANELS_ARRAY(grh_pay_freq, job_count) = "2" Then ALL_JOBS_PANELS_ARRAY(grh_pay_freq, job_count) = "Semi-Monthly"
    If ALL_JOBS_PANELS_ARRAY(grh_pay_freq, job_count) = "1" Then ALL_JOBS_PANELS_ARRAY(grh_pay_freq, job_count) = "Monthly"

end function

function read_SANC_panel()

    call  navigate_to_MAXIS_screen("stat", "sanc")
    'Now it checks for the total number of panels. If there's 0 Of 0 it'll exit the function for you so as to save oodles of time.
    EMReadScreen panel_total_check, 6, 2, 73
    IF panel_total_check = "0 Of 0" THEN exit function		'Exits out if there's no panel info
    first_sanc_panel  = true

    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        If ALL_MEMBERS_ARRAY(gather_detail, each_member) = TRUE Then
            EMWriteScreen ALL_MEMBERS_ARRAY(memb_numb, each_member), 20, 76
            EMWriteScreen "01", 20, 79
            transmit
            EMReadScreen SANC_total, 1, 2, 78
            If SANC_total <> 0 then
                EMReadScreen memb_sanc_number, 1, 16, 43
                EMReadScreen case_sanc_number, 1, 17, 43
                EMReadScreen case_compliance_date, 8, 17, 72
                EMReadScreen closed_for_7_sanc, 5, 18, 43
                EMReadScreen closed_for_post_7_sanc, 5, 19, 43


                If closed_for_7_sanc = "     " Then
                    If first_sanc_panel = true Then notes_on_sanction = notes_on_sanction & "Total case sanctions: " & case_sanc_number & ".; "

                    notes_on_sanction = notes_on_sanction & "Memb " & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " has incurred " & memb_sanc_number & " sanctions.; "
                Else
                    closed_for_7_sanc = replace(closed_for_7_sanc, " ", "/")
                    case_compliance_date = replace(case_compliance_date, " ", "/")
                    If first_sanc_panel = true Then
                        notes_on_sanction = notes_on_sanction & "Total case sanctions: 7. Case was closed for 7th sanction " & closed_for_7_sanc & ".; "
                        ' MsgBox "Case compliance Date - " & case_compliance_date & vbNewLine & "Is Date - " & IsDate(case_compliance_date)
                        If IsDate(case_compliance_date) = True Then notes_on_sanction = notes_on_sanction & "Case came into commpliance after closure for sanction on " & case_compliance_date & ".; "
                        If closed_for_post_7_sanc <> "     " Then
                            closed_for_post_7_sanc = replace(closed_for_post_7_sanc, " ", "/")
                            notes_on_sanction = notes_on_sanction & "Case clossed for 2nd Post-7th saction " & closed_for_post_7_sanc & ".; "
                        End If
                    End If
                    If memb_sanc_number = "6" Then
                        notes_on_sanction = notes_on_sanction & "Memb " & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " has incurred 7 sanctions and was closed for 7th sanction " & closed_for_7_sanc & ".; "
                    Else
                        notes_on_sanction = notes_on_sanction & "Memb " & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " has incurred " & memb_sanc_number & " sanctions.; "
                    End If


                End If

                first_sanc_panel = false
            End If
        End If
    Next
end function

function read_SHEL_panel()

    call navigate_to_MAXIS_screen("stat", "shel")

    'Now it checks for the total number of panels. If there's 0 Of 0 it'll exit the function for you so as to save oodles of time.
    EMReadScreen panel_total_check, 6, 2, 73
    IF panel_total_check = "0 Of 0" THEN exit function		'Exits out if there's no panel info

    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        EMWriteScreen ALL_MEMBERS_ARRAY(memb_numb, each_member), 20, 76
        EMWriteScreen "01", 20, 79
        transmit
        EMReadScreen SHEL_total, 1, 2, 78
        If SHEL_total <> 0 then
            member_number_designation = "Member " & HH_member & "- "
            row = 11
            ALL_MEMBERS_ARRAY(shel_exists, each_member) = TRUE
            Do
                EMReadScreen SHEL_HUD_code, 1, 6, 46
                EMReadScreen SHEL_share_code, 1, 6, 64

                If SHEL_HUD_code = "Y" Then ALL_MEMBERS_ARRAY(shel_subsudized, each_member) = "Yes"
                If SHEL_HUD_code = "N" Then ALL_MEMBERS_ARRAY(shel_subsudized, each_member) = "No"
                If SHEL_share_code = "Y" Then ALL_MEMBERS_ARRAY(shel_shared, each_member) = "Yes"
                If SHEL_share_code = "N" Then ALL_MEMBERS_ARRAY(shel_shared, each_member) = "No"

                EmReadScreen SHEL_retro_amount, 8, row, 37
                EMReadScreen SHEL_prosp_amount, 8, row, 56
                If SHEL_retro_amount <> "________" OR SHEL_prosp_amount <> "________" then
                    EMReadScreen SHEL_retro_proof, 2, row, 48
                    EMReadScreen SHEL_prosp_proof, 2, row, 67

                    If SHEL_prosp_amount = "________" Then SHEL_prosp_amount = 0
                    SHEL_prosp_amount = trim(SHEL_prosp_amount)
                    SHEL_prosp_amount = SHEL_prosp_amount * 1

                    If SHEL_retro_amount = "________" Then SHEL_retro_amount = 0
                    SHEL_retro_amount = trim(SHEL_retro_amount)
                    SHEL_retro_amount = SHEL_retro_amount * 1

                    If SHEL_retro_proof = "__" Then SHEL_retro_proof = "Blank"
                    If SHEL_prosp_proof = "__" Then SHEL_prosp_proof = "Blank"

                    If SHEL_retro_proof = "SF" Then SHEL_retro_proof = "SF - Shelter Form"
                    If SHEL_prosp_proof = "SF" Then SHEL_prosp_proof = "SF - Shelter Form"
                    If SHEL_retro_proof = "LE" Then SHEL_retro_proof = "LE - Lease"
                    If SHEL_prosp_proof = "LE" Then SHEL_prosp_proof = "LE - Lease"
                    If SHEL_retro_proof = "RE" Then SHEL_retro_proof = "RE - Rent Receipt"
                    If SHEL_prosp_proof = "RE" Then SHEL_prosp_proof = "RE - Rent Receipt"
                    If SHEL_retro_proof = "BI" Then SHEL_retro_proof = "BI - Billing Stmt"
                    If SHEL_prosp_proof = "BI" Then SHEL_prosp_proof = "BI - Billing Stmt"
                    If SHEL_retro_proof = "MO" Then SHEL_retro_proof = "MO - Mort Pmt Book"
                    If SHEL_prosp_proof = "MO" Then SHEL_prosp_proof = "MO - Mort Pmt Book"
                    If SHEL_retro_proof = "CD" Then SHEL_retro_proof = "CD - Ctrct For Deed"
                    If SHEL_prosp_proof = "CD" Then SHEL_prosp_proof = "CD - Ctrct For Deed"
                    If SHEL_retro_proof = "TX" Then SHEL_retro_proof = "TX - Prop Tax Stmt"
                    If SHEL_prosp_proof = "TX" Then SHEL_prosp_proof = "TX - Prop Tax Stmt"
                    If SHEL_retro_proof = "OT" Then SHEL_retro_proof = "OT - Other Doc"
                    If SHEL_prosp_proof = "OT" Then SHEL_prosp_proof = "OT - Other Doc"
                    If SHEL_retro_proof = "NC" Then SHEL_retro_proof = "NC - Change - Neg Impact"
                    If SHEL_prosp_proof = "NC" Then SHEL_prosp_proof = "NC - Change - Neg Impact"
                    If SHEL_retro_proof = "PC" Then SHEL_retro_proof = "PC - Change - Pos Impact"
                    If SHEL_prosp_proof = "PC" Then SHEL_prosp_proof = "PC - Change - Pos Impact"
                    If SHEL_retro_proof = "NO" Then SHEL_retro_proof = "NO - No Verif"
                    If SHEL_prosp_proof = "NO" Then SHEL_prosp_proof = "NO - No Verif"
                    If SHEL_retro_proof = "?_" Then SHEL_retro_proof = "? - Delayed Verif"
                    If SHEL_prosp_proof = "?_" Then SHEL_prosp_proof = "? - Delayed Verif"

                    If row = 11 Then
                        ALL_MEMBERS_ARRAY(shel_retro_rent_amt, each_member) = SHEL_retro_amount
                        ALL_MEMBERS_ARRAY(shel_retro_rent_verif, each_member) = SHEL_retro_proof
                        ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, each_member) = SHEL_prosp_amount
                        ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member) = SHEL_prosp_proof
                    ElseIf row = 12 Then
                        ALL_MEMBERS_ARRAY(shel_retro_lot_amt, each_member) = SHEL_retro_amount
                        ALL_MEMBERS_ARRAY(shel_retro_lot_verif, each_member) = SHEL_retro_proof
                        ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, each_member) = SHEL_prosp_amount
                        ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member) = SHEL_prosp_proof
                    ElseIf row = 13 Then
                        ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, each_member) = SHEL_retro_amount
                        ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, each_member) = SHEL_retro_proof
                        ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, each_member) = SHEL_prosp_amount
                        ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member) = SHEL_prosp_proof
                    ElseIf row = 14 Then
                        ALL_MEMBERS_ARRAY(shel_retro_ins_amt, each_member) = SHEL_retro_amount
                        ALL_MEMBERS_ARRAY(shel_retro_ins_verif, each_member) = SHEL_retro_proof
                        ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, each_member) = SHEL_prosp_amount
                        ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member) = SHEL_prosp_proof
                    ElseIf row = 15 Then
                        ALL_MEMBERS_ARRAY(shel_retro_tax_amt, each_member) = SHEL_retro_amount
                        ALL_MEMBERS_ARRAY(shel_retro_tax_verif, each_member) = SHEL_retro_proof
                        ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, each_member) = SHEL_prosp_amount
                        ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member) = SHEL_prosp_proof
                    ElseIf row = 16 Then
                        ALL_MEMBERS_ARRAY(shel_retro_room_amt, each_member) = SHEL_retro_amount
                        ALL_MEMBERS_ARRAY(shel_retro_room_verif, each_member) = SHEL_retro_proof
                        ALL_MEMBERS_ARRAY(shel_prosp_room_amt, each_member) = SHEL_prosp_amount
                        ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member) = SHEL_prosp_proof
                    ElseIf row = 17 Then
                        ALL_MEMBERS_ARRAY(shel_retro_garage_amt, each_member) = SHEL_retro_amount
                        ALL_MEMBERS_ARRAY(shel_retro_garage_verif, each_member) = SHEL_retro_proof
                        ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, each_member) = SHEL_prosp_amount
                        ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member) = SHEL_prosp_proof
                    ElseIf row = 18 Then
                        ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, each_member) = SHEL_retro_amount
                        ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, each_member) = SHEL_retro_proof
                        ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, each_member) = SHEL_prosp_amount
                        ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member) = SHEL_prosp_proof
                    End If

                    'ADD Reading of panel and saving to the array here'
                End if
                row = row + 1
            Loop until row = 19
        Else
            ALL_MEMBERS_ARRAY(shel_exists, each_member) = False
        End if
        SHEL_expense = ""
    Next
end function

function read_TIME_panel()
    call  navigate_to_MAXIS_screen("stat", "time")
    'Now it checks for the total number of panels. If there's 0 Of 0 it'll exit the function for you so as to save oodles of time.
    EMReadScreen panel_total_check, 6, 2, 73
    IF panel_total_check = "0 Of 0" THEN exit function		'Exits out if there's no panel info

    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        If ALL_MEMBERS_ARRAY(gather_detail, each_member) = TRUE Then
            EMWriteScreen ALL_MEMBERS_ARRAY(memb_numb, each_member), 20, 76
            EMWriteScreen "01", 20, 79
            transmit
            EMReadScreen TIME_total, 1, 2, 78
            If TIME_total <> 0 then
                EMReadScreen fed_tanf_months, 3, 17, 31
                EMReadScreen state_tanf_months, 3, 17, 53
                EMReadScreen total_tanf_months, 3, 17, 69
                EMReadScreen banked_tanf_months, 3, 19, 16
                EMReadScreen memb_ext_code, 2, 19, 31
                EMReadScreen memb_ext_total, 3, 19, 69

                fed_tanf_months = trim(fed_tanf_months)
                state_tanf_months = trim(state_tanf_months)
                total_tanf_months = trim(total_tanf_months)
                banked_tanf_months = trim(banked_tanf_months)
                memb_ext_total = trim(memb_ext_total)

                used_tanf = total_tanf_months * 1
                tanf_left = 60 - total_tanf_months
                If tanf_left < 0 Then tanf_left = 0

                If memb_ext_code = "01" Then memb_ext_info = "Ill or Incapacitated for more than 30 days"
                If memb_ext_code = "02" Then memb_ext_info = "Care of someone who is Ill or Incapacitated"
                If memb_ext_code = "03" Then memb_ext_info = "Care of someone with Special Medical Criteria"
                If memb_ext_code = "05" Then memb_ext_info = "Unemployable"
                If memb_ext_code = "06" Then memb_ext_info = "Low IQ"
                If memb_ext_code = "07" Then memb_ext_info = "Learning Disabled"
                If memb_ext_code = "08" Then memb_ext_info = "Employed 30+ hours per week (1 caregiver HH)"
                If memb_ext_code = "09" Then memb_ext_info = "Employed 55+ hours per week (2 caregived HH)"
                If memb_ext_code = "10" Then memb_ext_info = "Family Violence"
                If memb_ext_code = "11" Then memb_ext_info = "Developmental Disabilities"
                If memb_ext_code = "12" Then memb_ext_info = "Mental Illness"
                If memb_ext_code = "NO" Then memb_ext_info = "NONE"
                If memb_ext_code = "AP" Then memb_ext_info = "Appeal in Process"
                If memb_ext_code = "__" Then memb_ext_info = ""

                notes_on_time = notes_on_time & "Memb " & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " has used a total of " & total_tanf_months & " TANF months (" & fed_tanf_months & " Federal and " & state_tanf_months & "State) and has " & tanf_left & " TANF months remaining.; "
                If banked_tanf_months <> "0" Then notes_on_time = notes_on_time & "Memb " & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " - " & banked_tanf_months & " TANF Banked Months.; "
            End If
        End If
    Next
end function

function read_UNEA_panel()
    call navigate_to_MAXIS_screen("stat", "unea")

    'Now it checks for the total number of panels. If there's 0 Of 0 it'll exit the function for you so as to save oodles of time.
    EMReadScreen panel_total_check, 6, 2, 73
    IF panel_total_check = "0 Of 0" THEN exit function		'Exits out if there's no panel info

    If variable_written_to <> "" then variable_written_to = variable_written_to & "; "
    unea_array_counter = 0
    For each HH_member in HH_member_array
        EMWriteScreen HH_member, 20, 76
        EMWriteScreen "01", 20, 79
        transmit
        EMReadScreen UNEA_total, 1, 2, 78

        ReDim Preserve UNEA_INCOME_ARRAY(budget_notes, unea_array_counter)
        UNEA_INCOME_ARRAY(UC_exists, unea_array_counter) = FALSE
        UNEA_INCOME_ARRAY(CS_exists, unea_array_counter) = FALSE
        UNEA_INCOME_ARRAY(SSA_exists, unea_array_counter) = FALSE
        UNEA_INCOME_ARRAY(memb_numb, unea_array_counter) = HH_member
        If UNEA_total <> 0 then
            Do
                EMReadScreen income_type, 2, 5, 37

                EMReadScreen panel_month, 5, 20, 55
                panel_month = replace(panel_month, " ", "/")
                UNEA_INCOME_ARRAY(UNEA_month, unea_array_counter) = panel_month

                EMReadScreen UNEA_ver, 1, 5, 65
                If UNEA_ver = "1" Then UNEA_ver = "Copy of Checks"
                If UNEA_ver = "2" Then UNEA_ver = "Award Letters"
                If UNEA_ver = "3" Then UNEA_ver = "System Initiated Verif"
                If UNEA_ver = "4" Then UNEA_ver = "Colateral Statement"
                If UNEA_ver = "5" Then UNEA_ver = "Pend Out State Verif"
                If UNEA_ver = "6" Then UNEA_ver = "Other Document"
                If UNEA_ver = "7" Then UNEA_ver = "Worker Initiated Verif"
                If UNEA_ver = "8" Then UNEA_ver = "RI Stubs"
                If UNEA_ver = "N" Then UNEA_ver = "No Verif"
                If UNEA_ver = "?" Then UNEA_ver = "Delayed"
                If UNEA_ver = "_" Then UNEA_ver = "Blank"

                EMReadScreen UNEA_income_end_date, 8, 7, 68
                If UNEA_income_end_date <> "__ __ __" then UNEA_income_end_date = replace(UNEA_income_end_date, " ", "/")
                EMReadScreen UNEA_income_start_date, 8, 7, 37
                If UNEA_income_start_date <> "__ __ __" then UNEA_income_start_date = replace(UNEA_income_start_date, " ", "/")

                EMReadScreen prosp_amt, 8, 18, 68
                prosp_amt = trim(prosp_amt)
                EMReadScreen retro_amt, 8, 18, 39
                retro_amt = trim(retro_amt)

                EMWriteScreen "x", 10, 26
                transmit
                EMReadScreen SNAP_UNEA_amt, 8, 18, 56
                SNAP_UNEA_amt = trim(SNAP_UNEA_amt)
                EMReadScreen snap_pay_frequency, 1, 5, 64
                EMReadScreen date_of_pic_calc, 8, 5, 34
                date_of_pic_calc = replace(date_of_pic_calc, " ", "/")
                transmit

                If prosp_amt = "" Then prosp_amt = 0
                prosp_amt = prosp_amt * 1
                If retro_amt = "" Then retro_amt = 0
                retro_amt = retro_amt * 1
                If SNAP_UNEA_amt = "" Then SNAP_UNEA_amt = 0
                SNAP_UNEA_amt = SNAP_UNEA_amt * 1

                IF snap_pay_frequency = "1" THEN snap_pay_frequency = "monthly"
                IF snap_pay_frequency = "2" THEN snap_pay_frequency = "semimonthly"
                IF snap_pay_frequency = "3" THEN snap_pay_frequency = "biweekly"
                IF snap_pay_frequency = "4" THEN snap_pay_frequency = "weekly"
                IF snap_pay_frequency = "5" THEN snap_pay_frequency = "non-monthly"

                variable_name_for_UNEA = variable_name_for_UNEA & "UNEA from " & trim(UNEA_type) & ", " & UNEA_month  & " amts:; "
                If SNAP_UNEA_amt <> 0 THEN variable_name_for_UNEA = variable_name_for_UNEA & "- PIC: $" & SNAP_UNEA_amt & "/" & snap_pay_frequency & ", calculated " & date_of_pic_calc & "; "
                If retro_UNEA_amt <> 0 THEN variable_name_for_UNEA = variable_name_for_UNEA & "- Retrospective: $" & retro_UNEA_amt & " total; "
                If prosp_UNEA_amt <> 0 THEN variable_name_for_UNEA = variable_name_for_UNEA & "- Prospective: $" & prosp_UNEA_amt & " total; "
                'Leaving out HC income estimator if footer month is not Current month + 1
                If UNEA_ver = "N" or UNEA_ver = "?" then variable_name_for_UNEA = variable_name_for_UNEA & "- No proof provided for this panel; "

                If income_type = "01" or income_type = "02" or income_type = "03" or income_type = "44" Then
                    If IsDate(UNEA_income_end_date) = TRUE Then
                        If income_type = "01" or income_type = "02" Then ssa_type_for_note = "RSDI"
                        If income_type = "03" Then ssa_type_for_note = "SSI"
                        If income_type = "44" then ssa_type_for_note = "Excess Calculation of"
                        notes_on_ssa_income = notes_on_ssa_income & ssa_type_for_note & " income for Memb " & HH_member & " ended on " & UNEA_income_end_date & ". Verification: " & UNEA_ver & "; "
                    Else
                        UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) & ", SSA Income"
                        UNEA_INCOME_ARRAY(SSA_exists, unea_array_counter) = TRUE

                        If income_type = "01" or income_type = "02" Then
                            UNEA_INCOME_ARRAY(UNEA_RSDI_amt, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_RSDI_amt, unea_array_counter) + prosp_amt
                            If income_type = "01" Then UNEA_INCOME_ARRAY(UNEA_RSDI_notes, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_RSDI_notes, unea_array_counter) & "RSDI is Disability Income.; "
                            If SNAP_checkbox = checked and prosp_amt <> SNAP_UNEA_amt Then UNEA_INCOME_ARRAY(UNEA_RSDI_notes, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_RSDI_notes, unea_array_counter) & "SNAP budgeted Inomce =  $" & SNAP_UNEA_amt & "; "
                           UNEA_INCOME_ARRAY(UNEA_RSDI_notes, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_RSDI_notes, unea_array_counter) & "Verif: " & UNEA_ver & "; "
                            If IsDate(UNEA_income_start_date) = TRUE Then
                                If DateDiff("m", UNEA_income_start_date, date) < 6 Then UNEA_INCOME_ARRAY(UNEA_RSDI_notes, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_RSDI_notes, unea_array_counter) & "Income started in the past 6 months on " & UNEA_income_start_date & "; "
                            End If
                            If IsDate(UNEA_income_end_date) = True then UNEA_INCOME_ARRAY(UNEA_RSDI_notes, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_RSDI_notes, unea_array_counter) & "Income ended " & UNEA_income_end_date & "; "
                        End If
                        If income_type = "03" Then
                            UNEA_INCOME_ARRAY(UNEA_SSI_amt, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_SSI_amt, unea_array_counter) + prosp_amt
                            If SNAP_checkbox = checked and prosp_amt <> SNAP_UNEA_amt Then UNEA_INCOME_ARRAY(UNEA_SSI_notes, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_SSI_notes, unea_array_counter) & "SNAP budgeted Inomce =  $" & SNAP_UNEA_amt & "; "
                            UNEA_INCOME_ARRAY(UNEA_SSI_notes, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_SSI_notes, unea_array_counter) & "Verif: " & UNEA_ver & "; "
                            If IsDate(UNEA_income_start_date) = TRUE Then
                                If DateDiff("m", UNEA_income_start_date, date) < 6 Then UNEA_INCOME_ARRAY(UNEA_SSI_notes, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_SSI_notes, unea_array_counter) & "Income started in the past 6 months on " & UNEA_income_start_date & "; "
                            End If
                            If IsDate(UNEA_income_end_date) = True then UNEA_INCOME_ARRAY(UNEA_SSI_notes, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_SSI_notes, unea_array_counter) & "Income ended " & UNEA_income_end_date & "; "
                        End If
                    End If
                ElseIf income_type = "14" Then
                    If IsDate(UNEA_income_end_date) = TRUE Then
                        other_uc_income_notes = other_uc_income_notes & "Unemployment Income for Memb " & HH_member & " ended on " & UNEA_income_end_date & ". Verification: " & UNEA_ver & ".; "
                    Else
                        UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) & ", Unemployment"
                        UNEA_INCOME_ARRAY(UC_exists, unea_array_counter) = TRUE

                        UNEA_INCOME_ARRAY(UNEA_UC_retro_amt, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_retro_amt, unea_array_counter) + retro_amt
                        UNEA_INCOME_ARRAY(UNEA_UC_prosp_amt, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_prosp_amt, unea_array_counter) + prosp_amt
                        UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, unea_array_counter) + SNAP_UNEA_amt

                        EMReadScreen pay_day, 8, 13, 68
                        pay_day = trim(pay_day)
                        If pay_day = "" Then pay_day = 0
                        pay_day = pay_day * 1

                        UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, unea_array_counter) + pay_day
                       UNEA_INCOME_ARRAY(UNEA_UC_notes, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_notes, unea_array_counter) & "Verif: " & UNEA_ver & "; "
                        If IsDate(UNEA_income_start_date) = TRUE Then
                            If DateDiff("m", UNEA_income_start_date, date) < 6 Then UNEA_INCOME_ARRAY(UNEA_UC_notes, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_notes, unea_array_counter) & "Income started in the past 6 months on " & UNEA_income_start_date & "; "
                        End If
                        UNEA_INCOME_ARRAY(UNEA_UC_start_date, unea_array_counter) = UNEA_income_start_date
                    End If
                ElseIf income_type = "08" or income_type = "36" or income_type = "39" or income_type = "43" Then
                    If IsDate(UNEA_income_end_date) = TRUE Then
                        If income_type = "08" Then cs_type_for_note = "Direct Child Support"
                        If income_type = "36" Then cs_type_for_note = "Disbursed Child Support"
                        If income_type = "39" Then cs_type_for_note = "Disbursed Child Support Arrears"
                        notes_on_cses = notes_on_cses & cs_type_for_note & " income for Memb " & HH_member & " ended on " & UNEA_income_end_date & ". Verification: " & UNEA_ver & ".; "
                    Else
                        UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) & ", Child Support"
                        UNEA_INCOME_ARRAY(CS_exists, unea_array_counter) = TRUE

                        If income_type = "08" Then
                            UNEA_INCOME_ARRAY(direct_CS_amt, unea_array_counter) = UNEA_INCOME_ARRAY(direct_CS_amt, unea_array_counter) + prosp_amt
                            UNEA_INCOME_ARRAY(direct_CS_notes, unea_array_counter) = UNEA_INCOME_ARRAY(direct_CS_notes, unea_array_counter) & "Verif: " & UNEA_ver & "; "
                            If IsDate(UNEA_income_start_date) = TRUE Then
                                If DateDiff("m", UNEA_income_start_date, date) < 6 Then UNEA_INCOME_ARRAY(direct_CS_notes, unea_array_counter) = UNEA_INCOME_ARRAY(direct_CS_notes, unea_array_counter) & "Income started in the past 6 months on " & UNEA_income_start_date & "; "
                            End If
                            If IsDate(UNEA_income_end_date) = True then UNEA_INCOME_ARRAY(direct_CS_notes, unea_array_counter) = UNEA_INCOME_ARRAY(direct_CS_notes, unea_array_counter) & "Income ended " & UNEA_income_end_date & "; "
                        ElseIf income_type = "36" Then
                            UNEA_INCOME_ARRAY(disb_CS_amt, unea_array_counter) = UNEA_INCOME_ARRAY(disb_CS_amt, unea_array_counter) + prosp_amt
                            UNEA_INCOME_ARRAY(disb_CS_prosp_budg, unea_array_counter) = UNEA_INCOME_ARRAY(disb_CS_prosp_budg, unea_array_counter) + SNAP_UNEA_amt
                            UNEA_INCOME_ARRAY(disb_CS_notes, unea_array_counter) = UNEA_INCOME_ARRAY(disb_CS_notes, unea_array_counter) & "Verif: " & UNEA_ver & "; "
                            If IsDate(UNEA_income_start_date) = TRUE Then
                                If DateDiff("m", UNEA_income_start_date, date) < 6 Then UNEA_INCOME_ARRAY(disb_CS_notes, unea_array_counter) = UNEA_INCOME_ARRAY(disb_CS_notes, unea_array_counter) & "Income started in the past 6 months on " & UNEA_income_start_date & "; "
                            End If
                            If IsDate(UNEA_income_end_date) = True then UNEA_INCOME_ARRAY(disb_CS_notes, unea_array_counter) = UNEA_INCOME_ARRAY(disb_CS_notes, unea_array_counter) & "Income ended " & UNEA_income_end_date & "; "
                        ElseIf income_type = "39" Then
                            UNEA_INCOME_ARRAY(disb_CS_arrears_amt, unea_array_counter) = UNEA_INCOME_ARRAY(disb_CS_arrears_amt, unea_array_counter) + prosp_amt
                            UNEA_INCOME_ARRAY(disb_CS_arrears_budg, unea_array_counter) = UNEA_INCOME_ARRAY(disb_CS_arrears_budg, unea_array_counter) + SNAP_UNEA_amt
                            UNEA_INCOME_ARRAY(disb_cs_arrears_notes, unea_array_counter) = UNEA_INCOME_ARRAY(disb_cs_arrears_notes, unea_array_counter) & "Verif: " & UNEA_ver & "; "
                            If IsDate(UNEA_income_start_date) = TRUE Then
                                If DateDiff("m", UNEA_income_start_date, date) < 6 Then UNEA_INCOME_ARRAY(disb_cs_arrears_notes, unea_array_counter) = UNEA_INCOME_ARRAY(disb_cs_arrears_notes, unea_array_counter) & "Income started in the past 6 months on " & UNEA_income_start_date & "; "
                            End If
                            If IsDate(UNEA_income_end_date) = True then UNEA_INCOME_ARRAY(disb_cs_arrears_notes, unea_array_counter) = UNEA_INCOME_ARRAY(disb_cs_arrears_notes, unea_array_counter) & "Income ended " & UNEA_income_end_date & "; "
                        End If
                    End If
                ElseIf income_type = "11" or income_type = "12" or income_type = "13" or income_type = "38" Then
                    If income_type = "11" Then income_detail = "Disability Benefit"
                    If income_type = "12" Then income_detail = "Pension"
                    If income_type = "13" Then income_detail = "Other"
                    If income_type = "38" Then income_detail = "Aid & Attendance"

                    UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) & ", VA - " & income_detail

                    notes_on_VA_income = notes_on_VA_income & "; Member " & HH_member & "unearned income from VA (" & income_detail & "), verif: " & UNEA_ver & ", " & UNEA_INCOME_ARRAY(UNEA_month, unea_array_counter) & " amts:; "
                    If SNAP_UNEA_amt <> "" THEN notes_on_VA_income = notes_on_VA_income & "- PIC: $" & SNAP_UNEA_amt & "/" & snap_pay_frequency & ", calculated " & date_of_pic_calc & "; "
                    If retro_UNEA_amt <> "" THEN notes_on_VA_income = notes_on_VA_income & "- Retrospective: $" & retro_UNEA_amt & " total; "
                    If prosp_UNEA_amt <> "" THEN notes_on_VA_income = notes_on_VA_income & "- Prospective: $" & prosp_UNEA_amt & " total; "
                    If IsDate(UNEA_income_start_date) = TRUE Then
                        If DateDiff("m", UNEA_income_start_date, date) < 6 Then notes_on_VA_income = notes_on_VA_income & "Income started in the past 6 months on " & UNEA_income_start_date & "; "
                    End If
                    If IsDate(UNEA_income_end_date) = True then notes_on_VA_income = notes_on_VA_income & "Income ended " & UNEA_income_end_date & "; "
                ElseIf income_type = "15" Then
                    UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) & ", Worker's Comp"

                    notes_on_WC_income = notes_on_WC_income & "; Member " & HH_member & "unearned income from Worker's Comp, verif: " & UNEA_ver & ", " & UNEA_INCOME_ARRAY(UNEA_month, unea_array_counter) & " amts:; "
                    If SNAP_UNEA_amt <> "" THEN notes_on_WC_income = notes_on_WC_income & "- PIC: $" & SNAP_UNEA_amt & "/" & snap_pay_frequency & ", calculated " & date_of_pic_calc & "; "
                    If retro_UNEA_amt <> "" THEN notes_on_WC_income = notes_on_WC_income & "- Retrospective: $" & retro_UNEA_amt & " total; "
                    If prosp_UNEA_amt <> "" THEN notes_on_WC_income = notes_on_WC_income & "- Prospective: $" & prosp_UNEA_amt & " total; "
                    If IsDate(UNEA_income_start_date) = TRUE Then
                        If DateDiff("m", UNEA_income_start_date, date) < 6 Then notes_on_WC_income = notes_on_WC_income & "Income started in the past 6 months on " & UNEA_income_start_date & "; "
                    End If
                    If IsDate(UNEA_income_end_date) = True then notes_on_WC_income = notes_on_WC_income & "Income ended " & UNEA_income_end_date & "; "

                Else
                    If income_type = "06" Then income_type = "Public Assistance not in MN"
                    If income_type = "19" or income_type = "21" Then income_type = "Foster Care"
                    If income_type = "20" or income_type = "22" Then income_type = "Foster Care (not req FS)"
                    If income_type = "16" Then income_type = "Railroad Retirement"
                    If income_type = "17" Then income_type = "Retirement"
                    If income_type = "35" or income_type = "37" or income_type = "40" Then income_type = "Spousal Support"
                    If income_type = "18" Then income_type = "Military Entitlement"
                    If income_type = "23" Then income_type = "Dividends"
                    If income_type = "24" Then income_type = "Interest"
                    If income_type = "25" Then income_type = "Prizes and Gifts"
                    If income_type = "26" Then income_type = "Strike Benefit"
                    If income_type = "27" Then income_type = "Contract for Deed"
                    If income_type = "28" Then income_type = "Illegal Income"
                    If income_type = "29" Then income_type = "Other Countable"
                    If income_type = "30" Then income_type = "Infreq Irreg"
                    If income_type = "31" Then income_type = "Other FS Only"
                    If income_type = "45" Then income_type = "County 88 Gaming"
                    If income_type = "47" Then income_type = "Tribal Income"
                    If income_type = "48" Then income_type = "Trust Income"
                    If income_type = "49" Then income_type = "Non-Recurring"

                    UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) & ", " & income_type

                    notes_on_other_UNEA = notes_on_other_UNEA & "; Member " & HH_member & "unearned income from " & income_type & ", verif: " & UNEA_ver & ", " & UNEA_INCOME_ARRAY(UNEA_month, unea_array_counter) & " amts:; "
                    If SNAP_UNEA_amt <> "" THEN notes_on_other_UNEA = notes_on_other_UNEA & "- PIC: $" & SNAP_UNEA_amt & "/" & snap_pay_frequency & ", calculated " & date_of_pic_calc & "; "
                    If retro_UNEA_amt <> "" THEN notes_on_other_UNEA = notes_on_other_UNEA & "- Retrospective: $" & retro_UNEA_amt & " total; "
                    If prosp_UNEA_amt <> "" THEN notes_on_other_UNEA = notes_on_other_UNEA & "- Prospective: $" & prosp_UNEA_amt & " total; "
                    If IsDate(UNEA_income_start_date) = TRUE Then
                        If DateDiff("m", UNEA_income_start_date, date) < 6 Then notes_on_other_UNEA = notes_on_other_UNEA & "Income started in the past 6 months on " & UNEA_income_start_date & "; "
                    End If
                    If IsDate(UNEA_income_end_date) = True then notes_on_other_UNEA = notes_on_other_UNEA & "Income ended " & UNEA_income_end_date & "; "

                End If

                EMReadScreen UNEA_panel_current, 1, 2, 73
                If cint(UNEA_panel_current) < cint(UNEA_total) then transmit
            Loop until cint(UNEA_panel_current) = cint(UNEA_total)
        End if
        If left(UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter), 2) = ", " Then UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) = right(UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter), len(UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter)) - 2)

        UNEA_INCOME_ARRAY(UNEA_prosp_amt, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_prosp_amt, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(UNEA_retro_amt, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_retro_amt, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(UNEA_SNAP_amt, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_SNAP_amt, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(UNEA_UC_retro_amt, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_retro_amt, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(UNEA_UC_prosp_amt, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_prosp_amt, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(direct_CS_amt, unea_array_counter) = UNEA_INCOME_ARRAY(direct_CS_amt, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(disb_CS_amt, unea_array_counter) = UNEA_INCOME_ARRAY(disb_CS_amt, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(disb_CS_arrears_amt, unea_array_counter) = UNEA_INCOME_ARRAY(disb_CS_arrears_amt, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(UNEA_RSDI_amt, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_RSDI_amt, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(UNEA_SSI_amt, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_SSI_amt, unea_array_counter) & ""

        unea_array_counter = unea_array_counter + 1
    Next

    ' "01 - RSDI, Disa"
    ' "02 - RSDI, No Disa"
    ' "03 - SSI"
    ' "06 - Non-MN PA"
    ' "11 - VA Disability Benefit"
    ' "12 - VA Pension"
    ' "13 - VA Other"
    ' "38 - VA Aid & Attendance"
    ' "14 - Unemployment Insurance"
    ' "15 - Worker's Comp"
    ' "16 - Railroad Retirement"
    ' "17 - Other Retirement"
    ' "18 - Military Entitlement"
    ' "19 - FC Child Requestiong FS"
    ' "20 - FC Child Not Req FS"
    ' "21 - FC Adult Requesting FS"
    ' "22 - FC Adult Not Req FS"
    ' "23 - Dividends"
    ' "24 - Interest"
    ' "25 - Cnt Gifts or Prizes"
    ' "26 - Strike Benefit"
    ' "27 - Contract for Deed"
    ' "28 - Illegal Income"
    ' "29 - Other Countable"
    ' "30 - Infrequent <30 Not Counted"
    ' "31 - Other FS Only"
    '
    ' "08 - Direct Child Support"
    ' "35 - Direct Spousal Support"
    ' "36 - Disbursed Child Support"
    ' "37 - Disbursed Spousal Sup"
    ' "39 - Disbursed CS Arrears"
    ' "40 - Disbursed Spsl Sup Arrears"
    ' "43 - Disbursed Excess CS"
    '
    ' "44 - MSA - Excess Inc for SSI"
    ' "45 - County 88 Gaming"
    ' "47 - Counted Tribal Income"
    ' "48 - Trust Income (CASH)"
    ' "49 - Non-Recurring Income > $60/ptr (CASH)"
end function

function read_WREG_panel()
    call navigate_to_MAXIS_screen("stat", "wreg")

    'Now it checks for the total number of panels. If there's 0 Of 0 it'll exit the function for you so as to save oodles of time.
    EMReadScreen panel_total_check, 6, 2, 73
    IF panel_total_check = "0 Of 0" THEN exit function		'Exits out if there's no panel info
    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        If ALL_MEMBERS_ARRAY(include_snap_checkbox, each_member) = checked Then
            ' MsgBox "Member number is " & ALL_MEMBERS_ARRAY(memb_numb, each_member)
            EMWriteScreen ALL_MEMBERS_ARRAY(memb_numb, each_member), 20, 76
            transmit
            EMReadScreen wreg_total, 1, 2, 78
            IF wreg_total <> "0" THEN
                ALL_MEMBERS_ARRAY(wreg_exists, each_member) = TRUE
                EmWriteScreen "x", 13, 57
                transmit
                bene_mo_col = (15 + (4*cint(MAXIS_footer_month)))
                bene_yr_row = 10
                abawd_counted_months = 0
                second_abawd_period = 0
                month_count = 0
                DO
                    'establishing variables for specific ABAWD counted month dates
                    If bene_mo_col = "19" then counted_date_month = "01"
                    If bene_mo_col = "23" then counted_date_month = "02"
                    If bene_mo_col = "27" then counted_date_month = "03"
                    If bene_mo_col = "31" then counted_date_month = "04"
                    If bene_mo_col = "35" then counted_date_month = "05"
                    If bene_mo_col = "39" then counted_date_month = "06"
                    If bene_mo_col = "43" then counted_date_month = "07"
                    If bene_mo_col = "47" then counted_date_month = "08"
                    If bene_mo_col = "51" then counted_date_month = "09"
                    If bene_mo_col = "55" then counted_date_month = "10"
                    If bene_mo_col = "59" then counted_date_month = "11"
                    If bene_mo_col = "63" then counted_date_month = "12"
                    'reading to see if a month is counted month or not
                    EMReadScreen is_counted_month, 1, bene_yr_row, bene_mo_col
                    'counting and checking for counted ABAWD months
                    IF is_counted_month = "X" or is_counted_month = "M" THEN
                        EMReadScreen counted_date_year, 2, bene_yr_row, 15			'reading counted year date
                        abawd_counted_months_string = counted_date_month & "/" & counted_date_year
                        abawd_info_list = abawd_info_list & ", " & abawd_counted_months_string			'adding variable to list to add to array
                        abawd_counted_months = abawd_counted_months + 1				'adding counted months
                    END IF

                    'declaring & splitting the abawd months array
                    If left(abawd_info_list, 1) = "," then abawd_info_list = right(abawd_info_list, len(abawd_info_list) - 1)
                    abawd_months_array = Split(abawd_info_list, ",")

                    'counting and checking for second set of ABAWD months
                    IF is_counted_month = "Y" or is_counted_month = "N" THEN
                        EMReadScreen counted_date_year, 2, bene_yr_row, 15			'reading counted year date
                        second_abawd_period = second_abawd_period + 1				'adding counted months
                        second_counted_months_string = counted_date_month & "/" & counted_date_year			'creating new variable for array
                        second_set_info_list = second_set_info_list & ", " & second_counted_months_string	'adding variable to list to add to array
                        ALL_MEMBERS_ARRAY(first_second_set, each_member) = second_counted_months_string
                    END IF

                    'declaring & splitting the second set of abawd months array
                    If left(second_set_info_list, 1) = "," then second_set_info_list = right(second_set_info_list, len(second_set_info_list) - 1)
                    second_months_array = Split(second_set_info_list,",")

                    bene_mo_col = bene_mo_col - 4
                    IF bene_mo_col = 15 THEN
                        bene_yr_row = bene_yr_row - 1
                        bene_mo_col = 63
                    END IF
                    month_count = month_count + 1
                LOOP until month_count = 36
                ALL_MEMBERS_ARRAY(numb_abawd_used, each_member) = abawd_counted_months & ""
                If ALL_MEMBERS_ARRAY(numb_abawd_used, each_member) = "0" Then ALL_MEMBERS_ARRAY(numb_abawd_used, each_member) = ""
                ALL_MEMBERS_ARRAY(list_abawd_mo, each_member) = abawd_info_list
                ALL_MEMBERS_ARRAY(list_second_set, each_member) = second_set_info_list
                PF3

                EmreadScreen read_WREG_status, 2, 8, 50
                If read_WREG_status = "03" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "03  Unfit for Employment"
                If read_WREG_status = "04" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "04  Responsible for Care of Another"
                If read_WREG_status = "05" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "05  Age 60+"
                If read_WREG_status = "06" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "06  Under Age 16"
                If read_WREG_status = "07" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "07  Age 16-17, live w/ parent"
                If read_WREG_status = "08" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "08  Care of Child <6"
                If read_WREG_status = "09" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "09  Employed 30+ hrs/wk"
                If read_WREG_status = "10" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "10  Matching Grant"
                If read_WREG_status = "11" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "11  Unemployment Insurance"
                If read_WREG_status = "12" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "12  Enrolled in School/Training"
                If read_WREG_status = "13" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "13  CD Program"
                If read_WREG_status = "14" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "14  Receiving MFIP"
                If read_WREG_status = "20" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "20  Pend/Receiving DWP"
                If read_WREG_status = "15" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "15  Age 16-17 not live w/ Parent"
                If read_WREG_status = "16" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "16  50-59 Years Old"
                If read_WREG_status = "21" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "21  Care child < 18"
                If read_WREG_status = "17" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "17  Receiving RCA or GA"
                If read_WREG_status = "30" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "30  FSET Participant"
                If read_WREG_status = "02" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "02  Fail FSET Coop"
                If read_WREG_status = "33" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "33  Non-coop being referred"
                If read_WREG_status = "__" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "__  Blank"

                EmreadScreen read_abawd_status, 2, 13, 50
                If read_abawd_status = "01" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "01  WREG Exempt"
                If read_abawd_status = "02" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "02  Under Age 18"
                If read_abawd_status = "03" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "03  Age 50+"
                If read_abawd_status = "04" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "04  Caregiver of Minor Child"
                If read_abawd_status = "05" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "05  Pregnant"
                If read_abawd_status = "06" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "06  Employed 20+ hrs/wk"
                If read_abawd_status = "07" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "07  Work Experience"
                If read_abawd_status = "08" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "08  Other E and T"
                If read_abawd_status = "09" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "09  Waivered Area"
                IF read_abawd_status = "10" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "10  ABAWD Counted"
                If read_abawd_status = "11" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "11  Second Set"
                If read_abawd_status = "12" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "12  RCA or GA Participant"
                If read_abawd_status = "13" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "13  ABAWD Banked Months"
                If read_abawd_status = "__" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "__  Blank"

                EMReadScreen read_counter, 1, 14, 50
                If read_counter = "_" Then read_counter = 0
                ALL_MEMBERS_ARRAY(numb_banked_mo, each_member) = read_counter
            End If
        END IF
    Next
End function

function split_string_into_parts(full_string, partial_one, partial_two, partial_three, length_one, length_two)
    If left(full_string, 1) = "*" Then full_string = right(full_string, len(full_string) - 1)
    full_string = trim(full_string)
    If right(full_string, 1) = ";" Then full_string = left(full_string, len(full_string) - 1)
    full_string = trim(full_string)

    If len(full_string) =< length_one Then
        partial_one = full_string
        exit function
    End If

    full_string = replace(full_string, "  ", " ")
    word_array = split(full_string, " ")
    level = 1

    For each word in word_array
        If level = 1 Then
            If len(partial_one & " " & word) > length_one Then level = 2
        ElseIf level = 2 Then
            If partial_three <> "NONE" Then
                If len(partial_two & " " & word) > length_two Then level = 3
            End If
        End If

        If level = 1 Then
            partial_one = partial_one & " " & word
        ElseIf level = 2 Then
            partial_two = partial_two & " " & word
        ElseIf level = 3 Then
            partial_three = partial_three & " " & word
        End If
        '
        ' If len(partial_one & " " & word) < length_one Then
        '     partial_one = partial_one & " " & word
        ' ElseIf partial_three = "NONE" Then
        '     partial_two = partial_two & " " & word
        ' Else
        '     If len(partial_two & " " & word) < length_two Then
        '         partial_two = partial_two & " " & word
        '     Else
        '
        '         partial_three = partial_three & " " & word
        '     End If
        ' End If
        ' MsgBox "Word - " & word & vbNewLine & partial_one & vbNewLine & partial_two & vbNewLine & partial_three
    Next
end function

function update_shel_notes()
    total_shelter_amount = 0
    full_shelter_details = ""
    shelter_details = ""
    shelter_details_two = ""
    shelter_details_three = ""

    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        If ALL_MEMBERS_ARRAY(gather_detail, each_member) = TRUE Then

            If ALL_MEMBERS_ARRAY(shel_exists, each_member) = TRUE Then
                full_shelter_details = full_shelter_details & "* M" & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " shelter expense(s): "
                If ALL_MEMBERS_ARRAY(shel_retro_rent_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_rent_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_rent_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member) <> "Select one" Then
                        If ALL_MEMBERS_ARRAY(shel_retro_rent_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, each_member) Then
                            full_shelter_details = full_shelter_details & "Rent $" & ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member)) - 4) & ".; "
                        Else
                            full_shelter_details = full_shelter_details & "change in Rent retro - $" & ALL_MEMBERS_ARRAY(shel_retro_rent_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_rent_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_rent_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member)) - 4) & ".; " & ". "
                        End If
                        total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, each_member)
                    Else
                        full_shelter_details = full_shelter_details & "Rent (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_rent_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_rent_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_rent_verif, each_member)) - 4) & ".; "
                    End If
                ElseIf ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member) <> "Select one" Then
                    full_shelter_details = full_shelter_details & "Rent (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member)) - 4) & ".; "
                    total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, each_member)
                End If

                If ALL_MEMBERS_ARRAY(shel_retro_lot_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_lot_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_lot_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member) <> "Select one" Then
                        If ALL_MEMBERS_ARRAY(shel_retro_lot_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, each_member) Then
                            full_shelter_details = full_shelter_details & "Lot Rent $" & ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member)) - 4) & ".; "
                        Else
                            full_shelter_details = full_shelter_details & "change in Lot Rent retro - $" & ALL_MEMBERS_ARRAY(shel_retro_lot_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_lot_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_lot_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member)) - 4) & ".; "
                        End If
                        total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, each_member)
                    Else
                        full_shelter_details = full_shelter_details & "Lot Rent (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_lot_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_lot_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_lot_verif, each_member)) - 4) & ".; "
                    End If
                ElseIf ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member) <> "Select one" Then
                    full_shelter_details = full_shelter_details & "Lot Rent (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member)) - 4) & ".; "
                    total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, each_member)
                End If

                If ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member) <> "Select one" Then
                        If ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, each_member) Then
                            full_shelter_details = full_shelter_details & "Mortgage $" & ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member)) - 4) & ".; "
                        Else
                            full_shelter_details = full_shelter_details & "change in Mortgage retro - $" & ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member)) - 4) & ".; "
                        End If
                        total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, each_member)
                    Else
                        full_shelter_details = full_shelter_details & "Mortgage (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, each_member)) - 4) & ".; "
                    End If
                ElseIf ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member) <> "Select one" Then
                    full_shelter_details = full_shelter_details & "Mortgage (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member)) - 4) & ".; "
                    total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, each_member)
                End If

                If ALL_MEMBERS_ARRAY(shel_retro_ins_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_ins_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_ins_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member) <> "Select one" Then
                        If ALL_MEMBERS_ARRAY(shel_retro_ins_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, each_member) Then
                            full_shelter_details = full_shelter_details & "Home Insurance $" & ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member)) - 4) & ".; "
                        Else
                            full_shelter_details = full_shelter_details & "change in Home Insurance retro - $" & ALL_MEMBERS_ARRAY(shel_retro_ins_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_ins_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_ins_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member)) - 4) & ".; "
                        End If
                        total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, each_member)
                    Else
                        full_shelter_details = full_shelter_details & "Home Insurance (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_ins_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_ins_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_ins_verif, each_member)) - 4) & ".; "
                    End If
                ElseIf ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member) <> "Select one" Then
                    full_shelter_details = full_shelter_details & "Home Insurance (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member)) - 4) & ".; "
                    total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, each_member)
                End If

                If ALL_MEMBERS_ARRAY(shel_retro_tax_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_tax_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_tax_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member) <> "Select one" Then
                        If ALL_MEMBERS_ARRAY(shel_retro_tax_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, each_member) Then
                            full_shelter_details = full_shelter_details & "Property Tax $" & ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member)) - 4) & ".; "
                        Else
                            full_shelter_details = full_shelter_details & "change in Property Tax retro - $" & ALL_MEMBERS_ARRAY(shel_retro_tax_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_tax_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_tax_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member)) - 4) & ".; "
                        End If
                        total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, each_member)
                    Else
                        full_shelter_details = full_shelter_details & "Property Tax (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_tax_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_tax_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_tax_verif, each_member)) - 4) & ".; "
                    End If
                ElseIf ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member) <> "Select one" Then
                    full_shelter_details = full_shelter_details & "Property Tax (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member)) - 4) & ".; "
                    total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, each_member)
                End If

                If ALL_MEMBERS_ARRAY(shel_retro_room_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_room_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_room_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member) <> "Select one" Then
                        If ALL_MEMBERS_ARRAY(shel_retro_room_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_room_amt, each_member) Then
                            full_shelter_details = full_shelter_details & "Room $" & ALL_MEMBERS_ARRAY(shel_prosp_room_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member)) - 4) & ".; "
                        Else
                            full_shelter_details = full_shelter_details & "change in Room retro - $" & ALL_MEMBERS_ARRAY(shel_retro_room_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_room_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_room_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_room_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member)) - 4) & ".; "
                        End If
                        total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_room_amt, each_member)
                    Else
                        full_shelter_details = full_shelter_details & "Room (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_room_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_room_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_room_verif, each_member)) - 4) & ".; "
                    End If
                ElseIf ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member) <> "Select one" Then
                    full_shelter_details = full_shelter_details & "Room (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_room_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member)) - 4) & ".; "
                    total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_room_amt, each_member)
                End If

                If ALL_MEMBERS_ARRAY(shel_retro_garage_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_garage_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_garage_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member) <> "Select one" Then
                        If ALL_MEMBERS_ARRAY(shel_retro_garage_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, each_member) Then
                            full_shelter_details = full_shelter_details & "Garage $" & ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member)) - 4) & ".; "
                        Else
                            full_shelter_details = full_shelter_details & "change in Garage retro - $" & ALL_MEMBERS_ARRAY(shel_retro_garage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_garage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_garage_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member)) - 4) & ".; "
                        End If
                        total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, each_member)
                    Else
                        full_shelter_details = full_shelter_details & "Garage (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_garage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_garage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_garage_verif, each_member)) - 4) & ".; "
                    End If
                ElseIf ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member) <> "Select one" Then
                    full_shelter_details = full_shelter_details & "Garage (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member)) - 4) & ".; "
                    total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, each_member)
                End If

                If ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member) <> "Select one" Then
                        If ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, each_member) Then
                            full_shelter_details = full_shelter_details & "Subsidy $" & ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member)) - 4) & ".; "
                        Else
                            full_shelter_details = full_shelter_details & "change in Subsidy retro - $" & ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member)) - 4) & ".; "
                        End If
                        'total_shelter_amount = total_shelter_amount - ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, each_member)
                    Else
                        full_shelter_details = full_shelter_details & "Subsidy (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, each_member)) - 4) & ".; "
                    End If
                ElseIf ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member) <> "Select one" Then
                    full_shelter_details = full_shelter_details & "Subsidy (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member)) - 4) & ".; "
                    'total_shelter_amount = total_shelter_amount - ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, each_member)
                End If
                If ALL_MEMBERS_ARRAY(shel_shared, each_member) = "Yes" Then full_shelter_details = full_shelter_details & "Shelter expense is SHARED. "
                If ALL_MEMBERS_ARRAY(shel_subsudized, each_member) = "Yes" Then full_shelter_details = full_shelter_details & "Shelter expense is subsidized. "
            End If
        End If
    Next

    total_shelter_amount = FormatCurrency(total_shelter_amount)

    ' MsgBOx "Length of full_shelter_details is " & len(full_shelter_details)
    Call split_string_into_parts(full_shelter_details, shelter_details, shelter_details_two, shelter_details_three, 85, 85)
    ' if left(full_shelter_details, 2) = "; " Then full_shelter_details = right(full_shelter_details, len(full_shelter_details) - 2)
    ' If len(full_shelter_details) > 85 Then
    '     shelter_details = left(full_shelter_details, 85)
    '     shelter_details_two = right(full_shelter_details, len(full_shelter_details) - 85)
    ' Else
    '     shelter_details = full_shelter_details
    ' End If
end function

function update_wreg_and_abawd_notes()
    notes_on_wreg = ""
    full_abawd_info = ""
    notes_on_abawd = ""
    notes_on_abawd_two = ""
    notes_on_abawd_three = ""
    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        ' MsgBox "Each member - " & each_member & vbNewLine & "M" & ALL_MEMBERS_ARRAY(memb_numb, each_member) & vbNewLine & "WREG info - " & ALL_MEMBERS_ARRAY(clt_wreg_status, each_member)
        If ALL_MEMBERS_ARRAY(include_snap_checkbox, each_member) = checked Then
            If trim(ALL_MEMBERS_ARRAY(clt_wreg_status, each_member)) <> "" Then
                notes_on_wreg = notes_on_wreg & "M" & ALL_MEMBERS_ARRAY(memb_numb, each_member) & ": WREG - " & right(ALL_MEMBERS_ARRAY(clt_wreg_status, each_member), len(ALL_MEMBERS_ARRAY(clt_wreg_status, each_member)) - 4) & " ABAWD - " & right(ALL_MEMBERS_ARRAY(clt_abawd_status, each_member), len(ALL_MEMBERS_ARRAY(clt_abawd_status, each_member)) - 4) & "; "
                clt_currently_is = ""
                full_abawd_info = full_abawd_info & "M" & ALL_MEMBERS_ARRAY(memb_numb, each_member)
                If left(ALL_MEMBERS_ARRAY(clt_wreg_status, each_member), 2) = "30" Then
                    If left(ALL_MEMBERS_ARRAY(clt_abawd_status, each_member), 2) = "10" Then clt_currently_is = "ABAWD"
                    If left(ALL_MEMBERS_ARRAY(clt_abawd_status, each_member), 2) = "11" Then clt_currently_is = "SECOND SET"
                    'If left(ALL_MEMBERS_ARRAY(clt_abawd_status, each_member), 2) = "13" Then clt_currently_is = "BANKED"
                End If
                If clt_currently_is <> "" Then
                    full_abawd_info = full_abawd_info & " currently using " & clt_currently_is & " months."
                End If
                If ALL_MEMBERS_ARRAY(pwe_checkbox, each_member) = checked Then full_abawd_info = full_abawd_info & "; M" & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " is the  SNAP PWE"
                If ALL_MEMBERS_ARRAY(numb_abawd_used, each_member) <> "" OR trim(ALL_MEMBERS_ARRAY(list_abawd_mo, each_member)) <> "" Then full_abawd_info = full_abawd_info & "; ABAWD months used: " & ALL_MEMBERS_ARRAY(numb_abawd_used, each_member) & " - " & ALL_MEMBERS_ARRAY(list_abawd_mo, each_member)
                If trim(ALL_MEMBERS_ARRAY(first_second_set, each_member)) <> "" Then full_abawd_info = full_abawd_info & "; 2nd Set used starting: " & ALL_MEMBERS_ARRAY(first_second_set, each_member)
                If trim(ALL_MEMBERS_ARRAY(explain_no_second, each_member)) <> "" Then full_abawd_info = full_abawd_info & "; 2nd Set not available due to: " & ALL_MEMBERS_ARRAY(explain_no_second, each_member)
                'If trim(ALL_MEMBERS_ARRAY(numb_banked_mo, each_member)) <> "" Then full_abawd_info = full_abawd_info & "; Banked months used: " & ALL_MEMBERS_ARRAY(numb_banked_mo, each_member)
                If trim(ALL_MEMBERS_ARRAY(clt_abawd_notes, each_member)) <> "" Then full_abawd_info = full_abawd_info & "; Notes: " & ALL_MEMBERS_ARRAY(clt_abawd_notes, each_member)
                full_abawd_info = full_abawd_info & "; "
            End If
        End If
    Next
    if right(notes_on_wreg, 2) = "; " Then notes_on_wreg = left(notes_on_wreg, len(notes_on_wreg) - 2)

    Call split_string_into_parts(full_abawd_info, notes_on_abawd, notes_on_abawd_two, notes_on_abawd_three, 135, 135)
    ' if right(full_abawd_info, 2) = "; " Then full_abawd_info = left(full_abawd_info, len(full_abawd_info) - 2)
    ' If len(full_abawd_info) > 400 Then
    '     notes_on_abawd = left(full_abawd_info, 400)
    '     notes_on_abawd_two = right(full_abawd_info, len(full_abawd_info) - 400)
    ' Else
    '     notes_on_abawd = full_abawd_info
    ' End If
end function

function verification_dialog()
    If ButtonPressed = verif_button Then
        If second_call <> TRUE Then
            income_source_list = "Select or Type Source"

            For each_job = 0 to UBound(ALL_JOBS_PANELS_ARRAY, 2)
                If ALL_JOBS_PANELS_ARRAY(employer_name, each_job) <> "" Then income_source_list = income_source_list+chr(9)+"JOB - " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job)
            Next
            For each_busi = 0 to UBound(ALL_BUSI_PANELS_ARRAY, 2)
                If ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) <> "" Then
                    If ALL_BUSI_PANELS_ARRAY(busi_desc, each_busi) <> "" Then
                        income_source_list = income_source_list+chr(9)+"Self Emp - " & ALL_BUSI_PANELS_ARRAY(busi_desc, each_busi)
                    Else
                        income_source_list = income_source_list+chr(9)+"Self Employment"
                    End If
                End If
            Next
            employment_source_list = income_source_list
            income_source_list = income_source_list+chr(9)+"Child Support"+chr(9)+"Social Security Income"+chr(9)+"Unemployment Income"+chr(9)+"VA Income"+chr(9)+"Pension"
            income_verif_time = "[Enter Time Frame]"
            bank_verif_time = "[Enter Time Frame]"
            second_call = TRUE
        End If

        Do
            verif_err_msg = ""

            BeginDialog Dialog1, 0, 0, 610, 395, "Select Verifications"
              Text 280, 10, 120, 10, "Date Verification Request Form Sent:"
              EditBox 400, 5, 50, 15, verif_req_form_sent_date

              Groupbox 5, 35, 555, 130, "Personal and Household Information"

              CheckBox 10, 50, 75, 10, "Verification of ID for ", id_verif_checkbox
              ComboBox 90, 45, 150, 45, verification_memb_list, id_verif_memb
              CheckBox 300, 50, 100, 10, "Social Security Number for ", ssn_checkbox
              ComboBox 405, 45, 150, 45, verification_memb_list, ssn_verif_memb

              CheckBox 10, 70, 70, 10, "US Citizenship for ", us_cit_status_checkbox
              ComboBox 85, 65, 150, 45, verification_memb_list, us_cit_verif_memb
              CheckBox 300, 70, 85, 10, "Immigration Status for", imig_status_checkbox
              ComboBox 390, 65, 150, 45, verification_memb_list, imig_verif_memb

              CheckBox 10, 90, 90, 10, "Proof of relationship for ", relationship_checkbox
              ComboBox 105, 85, 150, 45, verification_memb_list, relationship_one_verif_memb
              Text 260, 90, 90, 10, "and"
              ComboBox 280, 85, 150, 45, verification_memb_list, relationship_two_verif_memb

              CheckBox 10, 110, 85, 10, "Student Information for ", student_info_checkbox
              ComboBox 100, 105, 150, 45, verification_memb_list, student_verif_memb
              Text 255, 110, 10, 10, "at"
              EditBox 270, 105, 150, 15, student_verif_source

              CheckBox 10, 130, 85, 10, "Proof of Pregnancy for", preg_checkbox
              ComboBox 100, 125, 150, 45, verification_memb_list, preg_verif_memb

              CheckBox 10, 150, 115, 10, "Illness/Incapacity/Disability for", illness_disability_checkbox
              ComboBox 130, 145, 150, 45, verification_memb_list, disa_verif_memb
              Text 285, 150, 30, 10, "verifying:"
              EditBox 320, 145, 150, 15, disa_verif_type

              GroupBox 5, 165, 555, 50, "Income Information"

              CheckBox 10, 180, 45, 10, "Income for ", income_checkbox
              ComboBox 60, 175, 150, 45, verification_memb_list, income_verif_memb
              Text 215, 180, 15, 10, "from"
              ComboBox 235, 175, 150, 45, income_source_list, income_verif_source
              Text 390, 180, 10, 10, "for"
              EditBox 405, 175, 150, 15, income_verif_time

              CheckBox 10, 200, 85, 10, "Employment Status for ", employment_status_checkbox
              ComboBox 100, 195, 150, 45, verification_memb_list, emp_status_verif_memb
              Text 255, 200, 10, 10, "at"
              ComboBox 270, 195, 150, 45, employment_source_list, emp_status_verif_source

              GroupBox 5, 215, 555, 50, "Expense Information"

              CheckBox 10, 230, 105, 10, "Educational Funds/Costs for", educational_funds_cost_checkbox
              ComboBox 120, 225, 150, 45, verification_memb_list, stin_verif_memb

              CheckBox 10, 250, 65, 10, "Shelter Costs for ", shelter_checkbox
              ComboBox 80, 245, 150, 45, verification_memb_list, shelter_verif_memb
              checkBox 240, 250, 175, 10, "Check here if this verif is NOT MANDATORY", shelter_not_mandatory_checkbox

              GroupBox 5, 265, 600, 30, "Asset Information"

              CheckBox 10, 280, 70, 10, "Bank Account for", bank_account_checkbox
              ComboBox 80, 275, 150, 45, verification_memb_list, bank_verif_memb
              Text 235, 280, 45, 10, "account type"
              ComboBox 285, 275, 145, 45, "Select or Type"+chr(9)+"Checking"+chr(9)+"Savings"+chr(9)+"Certificate of Deposit (CD)"+chr(9)+"Stock"+chr(9)+"Money Market", bank_verif_type
              Text 435, 280, 10, 10, "for"
              EditBox 450, 275, 150, 15, bank_verif_time

              Text 5, 305, 20, 10, "Other:"
              EditBox 30, 300, 570, 15, other_verifs
              Checkbox 10, 320, 200, 10, "Check here to have verifs numbered in the CASE/NOTE.", number_verifs_checkbox
              Checkbox 220, 320, 200, 10, "Check here if there are verifs that have been postponed.", verifs_postponed_checkbox

              ButtonGroup ButtonPressed
                PushButton 485, 10, 50, 15, "FILL", fill_button
                PushButton 540, 10, 60, 15, "Return to Dialog", return_to_dialog_button
              Text 10, 340, 580, 50, verifs_needed
              Text 10, 10, 235, 10, "Check the boxes for any verification you want to add to the CASE/NOTE."
              Text 10, 20, 470, 10, "Note: After you press 'Fill' or 'Return to Dialog' the information from the boxes will fill in the Verification Field and the boxes will be 'unchecked'."
            EndDialog

            dialog Dialog1

            If ButtonPressed = 0 Then
                id_verif_checkbox = unchecked
                us_cit_status_checkbox = unchecked
                imig_status_checkbox = unchecked
                ssn_checkbox = unchecked
                relationship_checkbox = unchecked
                income_checkbox = unchecked
                employment_status_checkbox = unchecked
                student_info_checkbox = unchecked
                educational_funds_cost_checkbox = unchecked
                shelter_checkbox = unchecked
                bank_account_checkbox = unchecked
                preg_checkbox = unchecked
                illness_disability_checkbox = unchecked
            End If
            If ButtonPressed = -1 Then ButtonPressed = fill_button

            If id_verif_checkbox = checked AND (id_verif_memb = "Select or Type Member" OR trim(id_verif_memb) = "") Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member that needs ID verified."
            If us_cit_status_checkbox = checked AND (us_cit_verif_memb = "Select or Type Member" OR trim(us_cit_verif_memb) = "") Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member that needs citizenship verified."
            If imig_status_checkbox = checked AND (imig_verif_memb = "Select or Type Member" OR trim(imig_verif_memb) = "") Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member that needs immigration status verified."
            If ssn_checkbox = checked AND (ssn_verif_memb = "Select or Type Member" OR trim(ssn_verif_memb) = "") Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member for which we need social security number."
            If relationship_checkbox = checked Then
                If relationship_one_verif_memb = "Select or Type Member" OR trim(relationship_one_verif_memb) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the two household members whose relationship needs to be verified."
                If relationship_two_verif_memb = "Select or Type Member" OR trim(relationship_two_verif_memb) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the two household members whose relationship needs to be verified."
            End If
            If income_checkbox = checked Then
                If income_verif_memb = "Select or Type Member" OR trim(income_verif_memb) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member whose income needs to be verified."
                If trim(income_verif_source) = "" OR trim(income_verif_source) = "Select or Type Source" Then verif_err_msg = verif_err_msg & vbNewLine & "* Enter the source of income to be verified."
                If trim(income_verif_time) = "[Enter Time Frame]" Then verif_err_msg = verif_err_msg & vbNewLine & "* Enter the time frame of the income verification needed."
            End If
            If employment_status_checkbox = checked Then
                If trim(emp_status_verif_source) = "" OR trim(emp_status_verif_source) = "Select or Type Source" Then verif_err_msg = verif_err_msg & vbNewLine & "* Enter the source of the employment that needs status verified."
                If emp_status_verif_memb = "Select or Type Member" OR trim(emp_status_verif_memb) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member whose employment status needs to be verified."
            End If
            If student_info_checkbox = checked Then
                If trim(student_verif_source) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Enter the source of school information to be verified"
                If student_verif_memb = "Select or Type Member" OR trim(student_verif_memb) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member for which we need school verification."
            End If
            If educational_funds_cost_checkbox = checked AND (stin_verif_memb = "Select or Type Member" OR trim(stin_verif_memb) = "") Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member with educational funds and costs we need verified."
            If shelter_checkbox = checked AND (shelter_verif_memb = "Select or Type Member" OR trim(shelter_verif_memb) = "") Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member whose shelter expense we need verified."
            If bank_account_checkbox = checked Then
                If trim(bank_verif_type) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Enter the type of bank account to verify."
                If bank_verif_memb = "Select or Type Member" OR trim(bank_verif_memb) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member whose bank account we need verified."
                If trim(bank_verif_time) = "[Enter Time Frame]" Then verif_err_msg = verif_err_msg & vbNewLine & "* Enter the time frame of the bank account verification needed."
            End If
            If preg_checkbox = checked AND (preg_verif_memb = "Select or Type Member" OR trim(preg_verif_memb) = "") Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member whose pregnancy needs to be verified."
            If illness_disability_checkbox = checked Then
                If trim(disa_verif_type) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Enter the type (or details) of the illness/incapacity/disability that need to be verified."
                If disa_verif_memb = "Select or Type Member" OR trim(disa_verif_memb) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member whose illness/incapacity/disability needs to be verified."
            End If

            If verif_err_msg = "" Then
                If id_verif_checkbox = checked Then
                    If IsNumeric(left(id_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Identity for Memb " & id_verif_memb & ".; "
                    Else
                        verifs_needed = verifs_needed & "Identity for " & id_verif_memb & ".; "
                    End If
                    id_verif_checkbox = unchecked
                    id_verif_memb = ""
                End If
                If us_cit_status_checkbox = checked Then
                    If IsNumeric(left(us_cit_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "US Citizenship for Memb " & us_cit_verif_memb & ".; "
                    Else
                        verifs_needed = verifs_needed & "US Citizenship for " & us_cit_verif_memb & ".; "
                    End If
                    us_cit_status_checkbox = unchecked
                    us_cit_verif_memb = ""
                End If
                If imig_status_checkbox = checked Then
                    If IsNumeric(left(imig_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Immigration documentation for Memb " & imig_verif_memb & ".; "
                    Else
                        verifs_needed = verifs_needed & "Immigration documentation for " & imig_verif_memb & ".; "
                    End If
                    imig_status_checkbox = unchecked
                    imig_verif_memb = ""
                End If
                If ssn_checkbox = checked Then
                    If IsNumeric(left(ssn_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Social Security number for Memb " & ssn_verif_memb & ".; "
                    Else
                        verifs_needed = verifs_needed & "Social Security number for " & ssn_verif_memb & ".; "
                    End If
                    ssn_checkbox = unchecked
                    ssn_verif_memb = ""
                End If
                If relationship_checkbox = checked Then
                    If IsNumeric(left(relationship_one_verif_memb, 2)) = TRUE AND IsNumeric(left(relationship_two_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Relationship between Memb " & relationship_one_verif_memb & " and Memb " & relationship_two_verif_memb & ".; "
                    Else
                        verifs_needed = verifs_needed & "Relationship between " & relationship_one_verif_memb & " and " & relationship_two_verif_memb & ".; "
                    End If
                    relationship_checkbox = unchecked
                    relationship_one_verif_memb = ""
                    relationship_two_verif_memb = ""
                End If
                If income_checkbox = checked Then
                    If IsNumeric(left(income_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Income for Memb " & income_verif_memb & " at " & income_verif_source & " for " & income_verif_time & ".; "
                    Else
                        verifs_needed = verifs_needed & "Income for " & income_verif_memb & " at " & income_verif_source & " for " & income_verif_time & ".; "
                    End If
                    income_checkbox = unchecked
                    income_verif_source = ""
                    income_verif_memb = ""
                    income_verif_time = ""
                End If
                If employment_status_checkbox = checked Then
                    If IsNumeric(left(emp_status_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Employment Status for Memb " & emp_status_verif_memb & " from " & emp_status_verif_source & ".; "
                    Else
                        verifs_needed = verifs_needed & "Employment Status for " & emp_status_verif_memb & " from " & emp_status_verif_source & ".; "
                    End If
                    employment_status_checkbox = unchecked
                    emp_status_verif_memb = ""
                    emp_status_verif_source = ""
                End If
                If student_info_checkbox = checked Then
                    If IsNumeric(left(student_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Student information for Memb " & student_verif_memb & " at " & student_verif_source & ".; "
                    Else
                        verifs_needed = verifs_needed & "Student information for " & student_verif_memb & " at " & student_verif_source & ".; "
                    End If
                    student_info_checkbox = unchecked
                    student_verif_memb = ""
                    student_verif_source = ""
                End If
                If educational_funds_cost_checkbox = checked Then
                    If IsNumeric(left(stin_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Educational funds and costs for Memb " & stin_verif_memb & ".; "
                    Else
                        verifs_needed = verifs_needed & "Educational funds and costs for " & stin_verif_memb & ".; "
                    End If
                    educational_funds_cost_checkbox = unchecked
                    stin_verif_memb = ""
                End If
                If shelter_checkbox = checked Then
                    If IsNumeric(left(shelter_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Shelter costs for Memb " & shelter_verif_memb & ". "
                    Else
                        verifs_needed = verifs_needed & "Shelter costs for " & shelter_verif_memb & ". "
                    End If
                    If shelter_not_mandatory_checkbox = checked Then verifs_needed = verifs_needed & " THIS VERIFICATION IS NOT MANDATORY."
                    verifs_needed = verifs_needed & "; "
                    shelter_checkbox = unchecked
                    shelter_verif_memb = ""
                End If
                If bank_account_checkbox = checked Then
                    If IsNumeric(left(bank_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & bank_verif_type & " account for Memb " & bank_verif_memb & " for " & bank_verif_time & ".; "
                    Else
                        verifs_needed = verifs_needed & bank_verif_type & " account for " & bank_verif_memb & " for " & bank_verif_time & ".; "
                    End If
                    bank_account_checkbox = unchecked
                    bank_verif_type = ""
                    bank_verif_memb = ""
                    bank_verif_time = ""
                End If
                If preg_checkbox = checked Then
                    If IsNumeric(left(preg_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Pregnancy for Memb " & preg_verif_memb & ".; "
                    Else
                        verifs_needed = verifs_needed & "Pregnancy for " & preg_verif_memb & ".; "
                    End If
                    preg_checkbox = unchecked
                    preg_verif_memb = ""
                End If
                If illness_disability_checkbox = checked Then
                    If IsNumeric(left(disa_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Ill/Incap or Disability for Memb " & disa_verif_memb & " of " & disa_verif_type & ",; "
                    Else
                        verifs_needed = verifs_needed & "Ill/Incap or Disability for " & disa_verif_memb & " of " & disa_verif_type & ",; "
                    End If
                    illness_disability_checkbox = unchecked
                    disa_verif_memb = ""
                    disa_verif_type = ""
                End If
                other_verifs = trim(other_verifs)
                If other_verifs <> "" Then verifs_needed = verifs_needed & other_verifs & "; "
                other_verifs = ""
            Else
                MsgBox "Additional detail about verifications to note is needed:" & vbNewLine & verif_err_msg
            End If

            If ButtonPressed = fill_button Then verif_err_msg = "LOOP" & verif_err_msg
        Loop until verif_err_msg = ""
        ButtonPressed = verif_button
    End If

end function

'===========================================================================================================================

'DECLARATIONS ==============================================================================================================
'Constants
'JOBS Array and BUSI Array constants
const memb_numb             = 0
const panel_instance        = 1
const employer_name         = 2
const busi_type             = 2         'for BUSI Array
Const estimate_only         = 3
const verif_explain         = 4
const verif_code            = 5
const calc_method           = 5         'for BUSI Array
const info_month            = 6
const hrly_wage             = 7
const mthd_date             = 7          'for BUSI Array'
const main_pay_freq         = 8
const rept_retro_hrs        = 8          'for BUSI Array'
const job_retro_income      = 9
const rept_prosp_hrs        = 9          'for BUSI Array'
const job_prosp_income      = 10
const min_wg_retro_hrs      = 10         'for BUSI Array'
const retro_hours           = 11
const min_wg_prosp_hrs      = 11         'for BUSI Array'
const prosp_hours           = 12
const income_ret_cash       = 12         'for BUSI Array'
const pic_pay_date_income   = 13
const income_pro_cash       = 13         'for BUSI Array'
const pic_pay_freq          = 14
const cash_income_verif     = 14         'for BUSI Array'
const pic_prosp_income      = 15
const expense_ret_cash      = 15         'for BUSI Array'
const pic_calc_date         = 16
const expense_pro_cash      = 16         'for BUSI Array'
const EI_case_note          = 17
const cash_expense_verif    = 17         'for BUSI Array'
const grh_calc_date         = 18
const income_ret_snap       = 18         'for BUSI Array'
const grh_pay_freq          = 19
const income_pro_snap       = 19         'for BUSI Array'
const grh_pay_day_income    = 20
const snap_income_verif     = 20         'for BUSI Array'
const grh_prosp_income      = 21
const expense_ret_snap      = 21         'for BUSI Array'
const expense_pro_snap      = 22         'for BUSI Array'
const snap_expense_verif    = 23         'for BUSI Array'
const method_convo_checkbox = 24         'for BUSI Array'
const start_date            = 25
const end_date              = 26
const busi_desc             = 27         'for BUSI Array'
const busi_structure        = 28         'for BUSI Array'
const share_num             = 29         'for BUSI Array'
const share_denom           = 30         'for BUSI Array'
const partners_in_HH        = 31         'for BUSI Array'
const exp_not_allwd         = 32         'for BUSI Array'
const verif_checkbox        = 33
const verif_added           = 34
const budget_explain        = 35

'Member array constants
const clt_name                  = 1
const clt_age                   = 2
const full_clt                  = 3
const clt_id_verif              = 4
const include_cash_checkbox     = 5
const include_snap_checkbox     = 6
const include_emer_checkbox     = 7
const count_cash_checkbox       = 8
const count_snap_checkbox       = 9
const count_emer_checkbox       = 10
const clt_wreg_status           = 11
const clt_abawd_status          = 12
const pwe_checkbox              = 13
const numb_abawd_used           = 14
const list_abawd_mo             = 15
const first_second_set          = 16
const list_second_set           = 17
const explain_no_second         = 18
const numb_banked_mo            = 19
const clt_abawd_notes           = 20
const shel_exists               = 21
const shel_subsudized           = 22
const shel_shared               = 23
const shel_retro_rent_amt       = 24
const shel_retro_rent_verif     = 25
const shel_prosp_rent_amt       = 26
const shel_prosp_rent_verif     = 27
const shel_retro_lot_amt        = 28
const shel_retro_lot_verif      = 29
const shel_prosp_lot_amt        = 30
const shel_prosp_lot_verif      = 31
const shel_retro_mortgage_amt   = 32
const shel_retro_mortgage_verif = 33
const shel_prosp_mortgage_amt   = 34
const shel_prosp_mortgage_verif = 35
const shel_retro_ins_amt        = 36
const shel_retro_ins_verif      = 37
const shel_prosp_ins_amt        = 38
const shel_prosp_ins_verif      = 39
const shel_retro_tax_amt        = 40
const shel_retro_tax_verif      = 41
const shel_prosp_tax_amt        = 42
const shel_prosp_tax_verif      = 43
const shel_retro_room_amt       = 44
const shel_retro_room_verif     = 45
const shel_prosp_room_amt       = 46
const shel_prosp_room_verif     = 47
const shel_retro_garage_amt     = 48
const shel_retro_garage_verif   = 49
const shel_prosp_garage_amt     = 50
const shel_prosp_garage_verif   = 51
const shel_retro_subsidy_amt    = 52
const shel_retro_subsidy_verif  = 53
const shel_prosp_subsidy_amt    = 54
const shel_prosp_subsidy_verif  = 55
const wreg_exists               = 56
const shel_verif_checkbox       = 57
const shel_verif_added          = 58
const gather_detail             = 59
const id_detail                 = 60
const id_required               = 61
const clt_notes                 = 62

'FOR CS Array'
const UNEA_type                 = 2
const UNEA_month                = 3
const UNEA_verif                = 4
const UNEA_prosp_amt            = 5
const UNEA_retro_amt            = 6
const UNEA_SNAP_amt             = 7
const UNEA_pay_freq             = 8
const UNEA_pic_date_calc        = 9

const UNEA_UC_start_date        = 10
const UNEA_UC_weekly_gross      = 11
const UNEA_UC_counted_ded       = 12
const UNEA_UC_exclude_ded       = 13
const UNEA_UC_weekly_net        = 14
const UNEA_UC_monthly_snap      = 15
const UNEA_UC_retro_amt         = 16
const UNEA_UC_prosp_amt         = 17
const UNEA_UC_notes             = 18
const UNEA_UC_tikl_date         = 19
const UNEA_UC_account_balance   = 20

const direct_CS_amt             = 21
const disb_CS_amt               = 22
const disb_CS_arrears_amt       = 23
const direct_CS_notes           = 24
const disb_CS_notes             = 25
const disb_CS_arrears_notes     = 26
const disb_CS_months            = 27
const disb_CS_prosp_budg        = 28
const disb_CS_arrears_months    = 29
const disb_CS_arrears_budg      = 30

const UNEA_RSDI_amt             = 31
const UNEA_RSDI_notes           = 32
const UNEA_SSI_amt              = 33
const UNEA_SSI_notes            = 34

const UC_exists                 = 35
const CS_exists                 = 36
const SSA_exists                = 37
const calc_button               = 38

const budget_notes              = 39

'Arrays
Dim ALL_JOBS_PANELS_ARRAY()
ReDim ALL_JOBS_PANELS_ARRAY(budget_explain, 0)


Dim ALL_BUSI_PANELS_ARRAY()
ReDim ALL_BUSI_PANELS_ARRAY(budget_explain, 0)

Dim ALL_MEMBERS_ARRAY()
ReDim ALL_MEMBERS_ARRAY(clt_notes, 0)

Dim UNEA_INCOME_ARRAY()
ReDim UNEA_INCOME_ARRAY(budget_notes, 0)
manual_amount_used = FALSE

'variables
Dim EATS, row, col, total_shelter_amount, full_shelter_details, shelter_details, shelter_details_two, shelter_details_three, hest_information, addr_line_one, relationship_detail
Dim addr_line_two, city, state, zip, address_confirmation_checkbox, addr_county, homeless_yn, addr_verif, reservation_yn, living_situation, number_verifs_checkbox, verifs_postponed_checkbox
Dim notes_on_address, notes_on_wreg, full_abawd_info, notes_on_busi, notes_on_abawd, notes_on_abawd_two, notes_on_abawd_three, verifs_needed, verif_req_form_sent_date
Dim other_uc_income_notes, notes_on_ssa_income, notes_on_VA_income, notes_on_WC_income, notes_on_other_UNEA, notes_on_cses, verification_memb_list, notes_on_time, notes_on_sanction

HH_memb_row = 5 'This helps the navigation buttons work!
application_signed_checkbox = checked 'The script should default to having the application signed.
verifs_needed = "[Information here creates a SEPARATE CASE/NOTE.]"

member_count = 0
adult_cash_count = 0
child_cash_count = 0
adult_snap_count = 0
child_snap_count = 0
adult_emer_count = 0
child_emer_count = 0

'===========================================================================================================================

'FUNCTIONS =================================================================================================================
'===========================================================================================================================

'Specialty functionality
Set objNet = CreateObject("WScript.NetWork")
windows_user_ID = objNet.UserName
user_ID_for_special_function = ucase(windows_user_ID)

'SCRIPT ====================================================================================================================
EMConnect ""
get_county_code				'since there is a county specific checkbox, this makes the the county clear
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
Call remove_dash_from_droplist(county_list)
script_run_lowdown = ""

BeginDialog Dialog1, 0, 0, 281, 235, "CAF Script Case number dialog"
  EditBox 65, 50, 60, 15, MAXIS_case_number
  EditBox 210, 50, 15, 15, MAXIS_footer_month
  EditBox 230, 50, 15, 15, MAXIS_footer_year
  CheckBox 10, 85, 30, 10, "CASH", CASH_on_CAF_checkbox
  CheckBox 50, 85, 35, 10, "SNAP", SNAP_on_CAF_checkbox
  CheckBox 90, 85, 35, 10, "EMER", EMER_on_CAF_checkbox
  DropListBox 135, 85, 140, 15, "Select One:"+chr(9)+"CAF (DHS-5223)"+chr(9)+"HUF (DHS-8107)"+chr(9)+"SNAP App for Srs (DHS-5223F)"+chr(9)+"ApplyMN"+chr(9)+"Combined AR for Certain Pops (DHS-3727)"+chr(9)+"CAF Addendum (DHS-5223C)", CAF_form
  EditBox 40, 130, 220, 15, cash_other_req_detail
  EditBox 40, 150, 220, 15, snap_other_req_detail
  EditBox 40, 170, 220, 15, emer_other_req_detail
  CheckBox 10, 195, 150, 10, "HC REVW Form is also being processed.", HC_checkbox
  ButtonGroup ButtonPressed
    PushButton 35, 215, 15, 15, "!", tips_and_tricks_button
    OkButton 170, 215, 50, 15
    CancelButton 225, 215, 50, 15
    PushButton 10, 30, 105, 10, "NOTES - Interview Completed", interview_completed_button
  Text 10, 10, 265, 20, "This script works best when run AFTER all STAT panels have been updated. If STAT panels have not been updated but you need to case note the interview use "
  Text 10, 55, 50, 10, "Case number:"
  Text 140, 55, 65, 10, "Footer month/year: "
  GroupBox 5, 70, 125, 30, "Programs marked on CAF"
  Text 135, 75, 65, 10, "Actual CAF Form:"
  GroupBox 5, 105, 265, 85, "OTHER Program Requests (not marked on CAF)"
  Text 40, 120, 130, 10, "Explain how the program was requested."
  Text 15, 135, 20, 10, "Cash:"
  Text 15, 155, 20, 10, "SNAP:"
  Text 15, 175, 25, 10, "EMER:"
  Text 55, 220, 105, 10, "Look for me for Tips and Tricks!"
EndDialog

'initial dialog
Do
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation

		If buttonpressed = interview_completed_button Then
            confirm_run_another_script = MsgBox("You have selected the 'NOTES - Interview completed' option. This will stop the NOTES - CAF script and run the script NOTES - Interview Completed." & vbNewLine & vbNewLine &_
                                                "This option is best for when the STAT panels have not been updated when running the script. We recommend runing NOTES - CAF once STAT panels are updated to capture the correct case information in CASE/NOTE." & vbNewLine & vbNewLine &_
                                                "Would you like to continue to NOTES - Interview Completed?", vbQuestion + vbYesNo, "Stop CAF Script?")
            If confirm_run_another_script = vbYes Then Call run_from_GitHub(script_repository & "notes/interview-completed.vbs")
            If confirm_run_another_script = vbNo Then err_msg = "LOOP" & err_msg
        End If
        If ButtonPressed = tips_and_tricks_button Then
            tips_tricks_msg = MsgBox("*** Tips and Tricks ***" & vbNewLine & "--------------------" & vbNewLine & vbNewLine & "Once the script reads the case, updates in MAXIS will not be reflected in the dialogs or case notes. This is not new, but if you run the script and realize a panel is out of date, definitely update the panel while the script is running, just don't expect the script to know that it was updated. You must also change the information IN the dialog. Or you can cancel the script, update and rerun the script with the panels correct." & vbNewLine & vbNewLine &_
                                    "Footer month/year - Use the month with the most accurate information for the CAF being processed." & vbNewLine & "Typically: " & vbNewLine & " - Recertifications use the month of recert." & vbNewLine & " - Applications use the month of application." & vbNewLine & vbNewLine &_
                                    "CAF Form - Select the actual form that was received. If the form is CAF Addendum (DHS-5223C) the script will call special functionality to handle specifically for an addendum." & vbNewLine & vbNewLine &_
                                    "Programs Requested - Listing anything in the boxes for other program requests will have the script assume that program is requested. Do not write anything here if that particular program has not been requested." & vbNewLine & "** An example would be a CAF with SNAP requested and in the interview a client requests CASH." & vbNewLine & vbNewLine &_
                                    "*** REMINDER***" & vbNewLine & "This script works best when MAXIS has been updated because it creates special dialogs with details from the MAXIS STAT panels and the most detail and specifics will be captured from an updated case." & vbNewLine &_
                                    "** Due to the complexity of this script and the noting needs, this script can take some time to complete. Use 'Interview Completed' if STAT has not been updated OR a quick note needs to be made. Run CAF once the case is updated.", vbInformation, "Tips and Tricks")

            err_msg = "LOOP" & err_msg
        End If

        If CAF_form = "Select One:" then err_msg = err_msg & vbnewline & "* You must select the CAF form received."
        Call validate_MAXIS_case_number(err_msg, "*")
        If IsNumeric(MAXIS_footer_month) = FALSE OR len(MAXIS_footer_month) > 2 Then err_msg = err_msg & vbNewLine & "* Enter a valid Footer Month."
        If IsNumeric(MAXIS_footer_year) = FALSE OR len(MAXIS_footer_year) > 2 Then err_msg = err_msg & vbNewLine & "* Enter a valid Footer Year."
        If CASH_on_CAF_checkbox = unchecked AND SNAP_on_CAF_checkbox = unchecked AND EMER_on_CAF_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "* At least one program should be marked on the CAF."
        If CASH_on_CAF_checkbox = checked AND trim(cash_other_req_detail) <> "" Then err_msg = err_msg & vbNewLine & "* If CASH was marked on the CAF, then another way of requesting does not need to be indicated."
        If SNAP_on_CAF_checkbox = checked AND trim(snap_other_req_detail) <> "" Then err_msg = err_msg & vbNewLine & "* If SNAP was marked on the CAF, then another way of requesting does not need to be indicated."
        If EMER_on_CAF_checkbox = checked AND trim(emer_other_req_detail) <> "" Then err_msg = err_msg & vbNewLine & "* If Emergency was marked on the CAF, then another way of requesting does not need to be indicated."
        If CAF_form = "SNAP App for Srs (DHS-5223F)" Then
            If CASH_on_CAF_checkbox = checked or trim(cash_other_req_detail) <> "" Then
                err_msg = err_msg & vbNewLine & "* The SNAP Application for Seniors can only be used for SNAP, not cash programs."
                CASH_on_CAF_checkbox = unchecked
                cash_other_req_detail = ""
            End If
            If EMER_on_CAF_checkbox = checked or trim(emer_other_req_detail) <> "" Then
                err_msg = err_msg & vbNewLine & "* The SNAP Application for Seniors can only be used for SNAP, not emergenc programs."
                EMER_on_CAF_checkbox = unchecked
                emer_other_req_detail = ""
            End If
        End If

        IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

If CAF_form = "CAF Addendum (DHS-5223C)" Then
    Call run_from_GitHub(script_repository & "notes/caf-addendum.vbs")
End If

If CASH_on_CAF_checkbox = checked or trim(cash_other_req_detail) <> "" Then cash_checkbox = checked
If SNAP_on_CAF_checkbox = checked or trim(snap_other_req_detail) <> "" Then SNAP_checkbox = checked
If EMER_on_CAF_checkbox = checked or trim(emer_other_req_detail) <> "" Then EMER_checkbox = checked
'grh_checkbox = checked

Call back_to_SELF
continue_in_inquiry = ""
EMReadScreen MX_region, 12, 22, 48
MX_region = trim(MX_region)
If MX_region = "INQUIRY DB" Then
    continue_in_inquiry = MsgBox("It appears you are in INQUIRY. Income information cannot be saved to STAT and a CASE/NOTE cannot be created." & vbNewLine & vbNewLine & "Do you wish to continue?", vbQuestion + vbYesNo, "Continue in Inquiry?")
    If continue_in_inquiry = vbNo Then script_end_procedure("Script ended since it was started in Inquiry.")
End If

exp_det_case_note_found = False
interview_completed_case_note_found = False
verifications_requested_case_note_found = False
caf_qualifying_questions_case_note_found = False

MAXIS_footer_month = right("00" & MAXIS_footer_month, 2)
MAXIS_footer_year = right("00" & MAXIS_footer_year, 2)
call check_for_MAXIS(False)	'checking for an active MAXIS session
MAXIS_footer_month_confirmation	'function will check the MAXIS panel footer month/year vs. the footer month/year in the dialog, and will navigate to the dialog month/year if they do not match.

script_run_lowdown = script_run_lowdown & vbCr & "CAF Type: " & CAF_type
script_run_lowdown = script_run_lowdown & vbCr & "Footer month:; " & MAXIS_footer_month & "/" & MAXIS_footer_year

If CASH_on_CAF_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "CASH Checked"
If trim(cash_other_req_detail) <> "" Then script_run_lowdown = script_run_lowdown & vbCr & "CASH: " & cash_other_req_detail
If SNAP_on_CAF_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "SNAP Checked"
If trim(snap_other_req_detail) <> "" Then script_run_lowdown = script_run_lowdown & vbCr & "SNAP: " & snap_other_req_detail
If EMER_on_CAF_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "EMER Checked"
If trim(emer_other_req_detail) <> "" Then script_run_lowdown = script_run_lowdown & vbCr & "EMER: " & emer_other_req_detail

'GRABBING THE DATE RECEIVED AND THE HH MEMBERS---------------------------------------------------------------------------------------------------------------------------------------------------------------------
loop_start = timer
Do
    call navigate_to_MAXIS_screen("STAT", "SUMM")
    EMReadScreen SUMM_check, 4, 2, 46
    Call back_to_SELF
    If timer - loop_start > 300 Then script_end_procedure("Can't get in to STAT. The script has attempted for 5 mintutes to get into STAT and iit appears to be stuck. The script timed out.")
Loop until SUMM_check = "SUMM"

'Creating a custom dialog for determining who the HH members are
call HH_comp_dialog(HH_member_array)

'GRABBING THE INFO FOR THE CASE NOTE-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
If cash_checkbox = checked Then
    If trim(child_cash_count) = "" OR child_cash_count = 0 Then
        adult_cash = TRUE
        family_cash = FALSE
    Else
        adult_cash = FALSE
        family_cash = TRUE
    End If
    If child_cash_count = 1 AND adult_cash_count = 0 Then
        adult_cash = TRUE
        family_cash = FALSE
    End If
    If pregnant_caregiver_checkbox = checked Then
        adult_cash = FALSE
        family_cash = TRUE
    End If
Else
    adult_cash = FALSE
    family_cash = FALSE
End If

If cash_checkbox = checked OR snap_checkbox = checked OR hc_checkbox = checked Then
    Call navigate_to_MAXIS_screen("STAT", "REVW")
    EMReadScreen cash_revw_code, 1, 7, 40
    EMReadScreen snap_revw_code, 1, 7, 60
    EMReadScreen hc_revw_code, 1, 7, 73
    If cash_revw_code = "N" or cash_revw_code = "U" or cash_revw_code = "I" or cash_revw_code = "A" Then
        the_process_for_cash = "Recertification"
        cash_recert_mo = MAXIS_footer_month
        cash_recert_yr = MAXIS_footer_year
    End If
    If snap_revw_code = "N" or snap_revw_code = "U" or snap_revw_code = "I" or snap_revw_code = "A" Then
        the_process_for_snap = "Recertification"
        snap_recert_mo = MAXIS_footer_month
        snap_recert_yr = MAXIS_footer_year
    End If
    If hc_revw_code = "N" or hc_revw_code = "U" or hc_revw_code = "I" or hc_revw_code = "A" Then
        the_process_for_hc = "Recertification"
        hc_recert_mo = MAXIS_footer_month
        hc_recert_yr = MAXIS_footer_year
    End If

    Call navigate_to_MAXIS_screen("STAT", "PROG")
    EMReadScreen cash_prog_code_one, 4, 6, 74
    EMReadScreen cash_prog_code_two, 4, 6, 74
    EMReadScreen snap_prog_code, 4, 10, 74
    EMReadScreen hc_prog_code, 4, 12, 74
    If cash_prog_code_one = "PEND" OR cash_prog_code_two = "PEND" Then the_process_for_cash = "Application"
    If snap_prog_code = "PEND" Then the_process_for_snap = "Application"
    If hc_prog_code = "PEND" Then the_process_for_hc = "Application"
    If the_process_for_cash = "Recertification" AND the_process_for_snap = "" AND cash_checkbox = checked AND snap_checkbox = checked Then
        EMReadScreen cash_prog_one, 2, 6, 67
        EMReadScreen cash_prog_two, 2, 7, 67
        If cash_prog_one = "MF" OR cash_prog_two = "MF" Then
            the_process_for_snap = "Recertification"
            snap_recert_mo = MAXIS_footer_month
            snap_recert_yr = MAXIS_footer_year
        End If
    End If

    If adult_cash = TRUE Then type_of_cash = "Adult"
    If family_cash = TRUE Then type_of_cash = "Family"
    dlg_len = 50
    y_pos = 25
    If cash_checkbox = checked Then dlg_len = dlg_len + 20
    If snap_checkbox = checked Then dlg_len = dlg_len + 20
    If HC_checkbox = checked Then dlg_len = dlg_len + 20

    BeginDialog Dialog1, 0, 0, 205, dlg_len, "CAF Process"
      Text 10, 10, 35, 10, "Program"
      Text 80, 10, 50, 10, "CAF Process"
      Text 155, 10, 50, 10, "Recert MM/YY"
      If cash_checkbox = checked Then
          Text 10, y_pos + 5, 20, 10, "Cash"
          DropListBox 35, y_pos, 35, 45, "Family"+chr(9)+"Adult", type_of_cash
          DropListBox 80, y_pos, 65, 45, "Select One..."+chr(9)+"Application"+chr(9)+"Recertification", the_process_for_cash
          EditBox 155, y_pos, 20, 15, cash_recert_mo
          EditBox 180, y_pos, 20, 15, cash_recert_yr
          y_pos = y_pos + 20
      End If
      If snap_checkbox = checked Then
          Text 10, y_pos + 5, 20, 10, "SNAP"
          DropListBox 80, y_pos, 65, 45, "Select One..."+chr(9)+"Application"+chr(9)+"Recertification", the_process_for_snap
          EditBox 155, y_pos, 20, 15, snap_recert_mo
          EditBox 180, y_pos, 20, 15, snap_recert_yr
          y_pos = y_pos + 20
      End If
      If HC_checkbox = checked Then
          Text 10, y_pos + 5, 40, 10, "Health Care"
          DropListBox 80, y_pos, 65, 45, "Select One..."+chr(9)+"Application"+chr(9)+"Recertification", the_process_for_hc
          EditBox 155, y_pos, 20, 15, hc_recert_mo
          EditBox 180, y_pos, 20, 15, hc_recert_yr
          y_pos = y_pos + 20
      End If
      y_pos = y_pos + 5
      ButtonGroup ButtonPressed
        OkButton 150, y_pos, 50, 15
    EndDialog

    Do
    	DO
    		err_msg = ""
    		Dialog Dialog1
    		cancel_confirmation

            If len(cash_recert_yr) = 4 AND left(cash_recert_yr, 2) = "20" Then cash_recert_yr = right(cash_recert_yr, 2)
            If len(snap_recert_yr) = 4 AND left(snap_recert_yr, 2) = "20" Then snap_recert_yr = right(snap_recert_yr, 2)
            If len(hc_recert_yr) = 4 AND left(hc_recert_yr, 2) = "20" Then hc_recert_yr = right(hc_recert_yr, 2)
            If cash_checkbox = checked Then
                If the_process_for_cash = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select if the CASH program is at application or recertification."
                If the_process_for_cash = "Recertification" AND (len(cash_recert_mo) <> 2 or len(cash_recert_yr) <> 2) Then err_msg = err_msg & vbNewLine & "* For CASH at recertification, enter the footer month and year the of the recertification."
                If CAF_form = "HUF (DHS-8107)" AND the_process_for_cash = "Application" then err_msg = err_msg & vbNewLine & "* An application for Cash cannot be processed using the HUF (Household Update Form). If you have a CAF type document, restart the script and select that form type. Otherwise you should select 'Recertification' for Cash."
            End If
            If snap_checkbox = checked Then
                If the_process_for_snap = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select if the SNAP program is at application or recertification."
                If the_process_for_snap = "Recertification" AND (len(snap_recert_mo) <> 2 or len(snap_recert_yr) <> 2) Then err_msg = err_msg & vbNewLine & "* For SNAP at recertification, enter the footer month and year the of the recertification."
                If CAF_form = "HUF (DHS-8107)" AND the_process_for_snap = "Application" then err_msg = err_msg & vbNewLine & "* An application for SNAP cannot be processed using the HUF (Household Update Form). If you have a CAF type document, restart the script and select that form type. Otherwise you should select 'Recertification' for SNAP."
            End If
            If HC_checkbox = checked Then
                If the_process_for_hc = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select if the Health Care program is at application or recertification."
                If the_process_for_hc = "Recertification" AND (len(hc_recert_mo) <> 2 or len(hc_recert_yr) <> 2) Then err_msg = err_msg & vbNewLine & "* For HC at recertification, enter the footer month and year the of the recertification."
            End If


            IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** Please resolve to continue ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
    	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

    If the_process_for_cash = "Recertification" OR the_process_for_snap = "Recertification" Then CAF_type = "Recertification"
    If the_process_for_cash = "Application" OR the_process_for_snap = "Application" Then CAF_type = "Application"
    If type_of_cash = "Family" Then
        adult_cash = FALSE
        family_cash = TRUE
    End If
    If type_of_cash = "Adult" Then
        adult_cash = TRUE
        family_cash = FALSE
    End If
End If
If EMER_checkbox = checked Then CAF_type = "Application"

' If interview_required = TRUE Then
'     Interview_notice = MsgBox("Has an interview been completed?" &vbNewLine & vbNewLine & "                *~* WITHOUT AN INTERVIEW *~* " & vbNewLine & "             *~* A CAF CANNOT BE PROCESSED *~*" & vbNewLine & vbNewLine & "If you have not completed an interview, do not use this script as you cannot process a CAF without an interview.  There are a couple scripts that may be useful for noting review of a case with a CAF received before the interview has been completed:" & vbNewLine & "          NOTES - Application Check" & vbNewLine & "          NOTES - Client Contact" & vbNewLine & vbNewLine & "Press OK if you completed an interview (or are processing one of the two exceptions)." & vbNewLine & "Press Cancel if you have not interviewed yet.", vbExclamation + vbOkCancel, "Have you done an interview?")
'     If Interview_notice = vbCancel Then script_end_procedure_with_error_report("The script has been cancelled as no interview has been completed.")
' End If

If CAF_type = "Recertification" then                                                          'For recerts it goes to one area for the CAF datestamp. For other app types it goes to STAT/PROG.
	' call autofill_editbox_from_MAXIS(HH_member_array, "REVW", CAF_datestamp)
    Call navigate_to_MAXIS_screen("STAT", "REVW")
    EMReadScreen CAF_datestamp, 8, 13, 37                       'reading the prog date
    CAF_datestamp = replace(CAF_datestamp, " ", "/")
    If isdate(CAF_datestamp) = True then
      CAF_datestamp = cdate(CAF_datestamp) & ""
    Else
      CAF_datestamp = ""
    End if

    EMReadScreen interview_date, 8, 15, 37                       'reading the prog date
    interview_date = replace(interview_date, " ", "/")
    If isdate(interview_date) = True then
      interview_date = cdate(interview_date) & ""
    Else
      interview_date = ""
    End if

	IF SNAP_checkbox = checked THEN																															'checking for SNAP 24 month renewals.'
		EMWriteScreen "X", 05, 58																																	'opening the FS revw screen.
		transmit
		EMReadScreen SNAP_recert_date, 8, 9, 64
		PF3
		SNAP_recert_date = replace(SNAP_recert_date, " ", "/")
        If SNAP_recert_date <> "__/01/__" Then 																	'replacing the read blank spaces with / to make it a date
    		SNAP_recert_compare_date = dateadd("m", "12", MAXIS_footer_month & "/01/" & MAXIS_footer_year)		'making a dummy variable to compare with, by adding 12 months to the requested footer month/year.
    		IF datediff("d", SNAP_recert_compare_date, SNAP_recert_date) > 0 THEN											'If the read recert date is more than 0 days away from 12 months plus the MAXIS footer month/year then it is likely a 24 month period.'
    			SNAP_recert_is_likely_24_months = TRUE
    		ELSE
    			SNAP_recert_is_likely_24_months = FALSE																									'otherwise if we don't we set it as false
    		END IF
        Else
            SNAP_recert_is_likely_24_months = FALSE
        End If
	END IF
Else
	' call autofill_editbox_from_MAXIS(HH_member_array, "PROG", CAF_datestamp)
    Call navigate_to_MAXIS_screen("STAT", "PROG")

    row = 6
    Do
        EMReadScreen appl_prog_date, 8, row, 33
        If appl_prog_date <> "__ __ __" then appl_prog_date_array = appl_prog_date_array & replace(appl_prog_date, " ", "/") & " "

        EMReadScreen appl_intv_date, 8, row, 55
        If appl_intv_date <> "__ __ __" AND appl_intv_date <> "        " then appl_intv_date_array = appl_intv_date_array & replace(appl_intv_date, " ", "/") & " "

        row = row + 1
    Loop until row = 13
    appl_prog_date_array = split(appl_prog_date_array)
    CAF_datestamp = CDate(appl_prog_date_array(0))
    for i = 0 to ubound(appl_prog_date_array) - 1
        if CDate(appl_prog_date_array(i)) > CAF_datestamp then
            CAF_datestamp = CDate(appl_prog_date_array(i))
        End if
    next
    If isdate(CAF_datestamp) = True then
        CAF_datestamp = cdate(CAF_datestamp) & ""
    Else
        CAF_datestamp = ""
    End if

    If trim(appl_intv_date_array) <> "" Then
        appl_intv_date_array = split(appl_intv_date_array)
        If IsArray(appl_intv_date_array) = TRUE AND IsDate(appl_intv_date_array(0)) = TRUE Then
            interview_date = CDate(appl_intv_date_array(0))
            for i = 0 to ubound(appl_intv_date_array) - 1
                if CDate(appl_intv_date_array(i)) > interview_date then
                    interview_date = CDate(appl_intv_date_array(i))
                End if
            next
            If isdate(interview_date) = True then
                interview_date = cdate(interview_date) & ""
            Else
                interview_date = ""
            End if
        End If
    End If
End if
If IsDate(CAF_datestamp) = False Then
    BeginDialog Dialog1, 0, 0, 125, 45, "CAF Datestamp"
      EditBox 75, 5, 45, 15, CAF_datestamp
      ButtonGroup ButtonPressed
        OkButton 5, 25, 50, 15
        CancelButton 60, 25, 50, 15
      Text 10, 10, 60, 10, "CAF Datestamp:"
    EndDialog

    'Runs the first dialog - which confirms the case number
    Do
    	Do
    		err_msg = ""
    		dialog Dialog1
    		cancel_confirmation
            If IsDate(CAF_datestamp) = False Then err_msg = err_msg & vbNewLine & "* Please enter a valid date for the date the CAF was received."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
End If

'THIS IS HANDLING SPECIFICALLY AROUND THE ALLOWANCE TO WAIVE INTERVIEWS FOR RENEWALS IN EFFECT STARTING FOR 04/21 REVW
check_for_waived_interview = FALSE
interview_waived = FALSE
interview_required = FALSE

If the_process_for_snap = "Recertification" Then check_for_waived_interview = TRUE
If the_process_for_cash = "Recertification" AND type_of_cash = "Family" Then check_for_waived_interview = TRUE

If SNAP_checkbox = checked OR family_cash = TRUE OR CAF_type = "Application" then interview_required = TRUE

If check_for_waived_interview = TRUE AND interview_required = TRUE Then
    interview_is_being_waived = MsgBox("Renewals can be processed without an interview per DHS." & vbNewLine & vbNewLine & " --- Are you waiving the interview? ---" & vbNewLine & vbNewLine & "clicking 'YES' will will prevent the script from requiring Interview Detail.", vbquestion + vbYesNo, "")
    If interview_is_being_waived = vbYes Then
        interview_required = FALSE
        interview_waived = TRUE
    End If
End If

MAXIS_case_number = trim(MAXIS_case_number)

look_for_expedited_determination_case_note = False
look_for_expedited_determination_case_note = False
look_for_expedited_determination_case_note = False

If CAF_type = "Application" Then
    If SNAP_checkbox = checked Then look_for_expedited_determination_case_note = True
End If
Call Navigate_to_MAXIS_screen("CASE", "NOTE")

too_old_date = DateAdd("D", -1, CAF_datestamp)

note_row = 5
Do
    EMReadScreen note_date, 8, note_row, 6

    EMReadScreen note_title, 55, note_row, 25
    note_title = trim(note_title)

    If look_for_expedited_determination_case_note = True Then
        If left(note_title, 47) = "Expedited Determination: SNAP appears expedited" Then
            exp_det_case_note_found = TRUE
            snap_exp_yn = "Yes"
        End If
        If left(note_title, 55) = "Expedited Determination: SNAP does not appear expedited" Then
            exp_det_case_note_found = TRUE
            snap_exp_yn = "No"
        End If
        If left(note_title, 42) = "Expedited Determination: SNAP to be denied" Then
            exp_det_case_note_found = TRUE

            EMWriteScreen "X", note_row, 3  'Opens the note to read the denial date'
            transmit

            read_row = ""
            EMReadScreen find_denial_date_line, 22, 5, 3
            If find_denial_date_line = "* SNAP to be denied on" Then
                read_row = 5
            Else
                EMReadScreen find_denial_date_line, 22, 6, 3
                If find_denial_date_line = "* SNAP to be denied on" Then read_row = 6
            End If
            If read_row <> "" Then
                EMReadScreen note_denial_date, 10, row, 25
                note_denial_date = replace(note_denial_date, "", ".")
                note_denial_date = replace(note_denial_date, "", "S")
                note_denial_date = replace(note_denial_date, "", "i")
                note_denial_date = replace(note_denial_date, "", "n")
                note_denial_date = replace(note_denial_date, "", "c")
                note_denial_date = replace(note_denial_date, "", "e")
                note_denial_date = trim(note_denial_date)
                If IsDate(note_denial_date) = True Then snap_denial_date = note_denial_date
            End If

            PF3                             'closing the note
        End IF
    End If

    If left(note_title, 24) = "~ Interview Completed on" Then
        interview_completed_case_note_found = True
    End If
    If left(note_title, 23) = "VERIFICATIONS REQUESTED" Then
        verifications_requested_case_note_found = True
        verifs_needed = "PREVIOUS NOTE EXISTS"
    End If
    If left(note_title, 43) = "Qualifying Questions had an answer of 'YES'" Then
        caf_qualifying_questions_case_note_found = True
    End If

    if note_date = "        " then Exit Do
    ' if exp_det_case_note_found = TRUE then Exit Do

    note_row = note_row + 1
    if note_row = 19 then
        'MsgBox "Next Page" & vbNewLine & "Note Date:" & note_date
        note_row = 5
        PF8
        EMReadScreen check_for_last_page, 9, 24, 14
        If check_for_last_page = "LAST PAGE" Then Exit Do
    End If
    EMReadScreen next_note_date, 8, note_row, 6
    if next_note_date = "        " then Exit Do
Loop until DateDiff("d", too_old_date, next_note_date) <= 0

' 'TESTING CODE'
' MsgBox "Did we find CASE:NOTES?" & vbCr & vbCr &_
'        "EXP Determination - " & exp_det_case_note_found & vbCr &_
'        "Interview Completed - " & interview_completed_case_note_found & vbCr &_
'        "VERIFS Requested - " & verifications_requested_case_note_found & vbCr &_
'        "CAF Qual Questions - " & caf_qualifying_questions_case_note_found

If exp_det_case_note_found = TRUE Then script_run_lowdown = script_run_lowdown & vbCr & "Found Expedited Case Note"
If interview_completed_case_note_found = TRUE Then script_run_lowdown = script_run_lowdown & vbCr & "Found Interview Completed Note"
If verifications_requested_case_note_found = TRUE Then script_run_lowdown = script_run_lowdown & vbCr & "Found Verifs Requested Note"
If caf_qualifying_questions_case_note_found = TRUE Then script_run_lowdown = script_run_lowdown & vbCr & "Found CAF Qual Questions Note"

' call autofill_editbox_from_MAXIS(HH_member_array, "SHEL", SHEL_HEST)
call read_ADDR_panel
call read_SHEL_panel
call update_shel_notes
call read_HEST_panel
' call autofill_editbox_from_MAXIS(HH_member_array, "HEST", SHEL_HEST)

'Now it grabs the rest of the info, not dependent on which programs are selected.
call autofill_editbox_from_MAXIS(HH_member_array, "ABPS", ABPS)
call autofill_editbox_from_MAXIS(HH_member_array, "ACCI", ACCI)
call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", notes_on_acct)
call autofill_editbox_from_MAXIS(HH_member_array, "ACUT", notes_on_acut)
call autofill_editbox_from_MAXIS(HH_member_array, "AREP", AREP)
call autofill_editbox_from_MAXIS(HH_member_array, "BILS", BILS)
call autofill_editbox_from_MAXIS(HH_member_array, "CASH", notes_on_cash)
call autofill_editbox_from_MAXIS(HH_member_array, "CARS", notes_on_cars)
call autofill_editbox_from_MAXIS(HH_member_array, "COEX", notes_on_coex)
call autofill_editbox_from_MAXIS(HH_member_array, "DCEX", notes_on_dcex)
If cash_checkbox = checked Then call autofill_editbox_from_MAXIS(HH_member_array, "DIET", DIET)
call autofill_editbox_from_MAXIS(HH_member_array, "DISA", DISA)
call autofill_editbox_from_MAXIS(HH_member_array, "EMPS", EMPS)
call autofill_editbox_from_MAXIS(HH_member_array, "FACI", FACI)
call autofill_editbox_from_MAXIS(HH_member_array, "FMED", FMED)
call autofill_editbox_from_MAXIS(HH_member_array, "IMIG", IMIG)
call autofill_editbox_from_MAXIS(HH_member_array, "INSA", INSA)
'call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)

job_count = 0
call navigate_to_MAXIS_screen("STAT", "JOBS")
EMReadScreen panel_total_check, 6, 2, 73
IF panel_total_check <> "0 Of 0" Then
    For each HH_member in HH_member_array
        EMWriteScreen HH_member, 20, 76
        EMWriteScreen "01", 20, 79
        transmit
        EMReadScreen JOBS_total, 1, 2, 78
        If JOBS_total <> 0 then
            Do
                ReDim Preserve ALL_JOBS_PANELS_ARRAY(budget_explain, job_count)
                ALL_JOBS_PANELS_ARRAY(memb_numb, job_count) = HH_member
                ALL_JOBS_PANELS_ARRAY(info_month, job_count) = MAXIS_footer_month & "/" & MAXIS_footer_year
                call read_JOBS_panel

                EMReadScreen JOBS_panel_current, 1, 2, 73
                ALL_JOBS_PANELS_ARRAY(panel_instance, job_count) = "0" & JOBS_panel_current

                If cint(JOBS_panel_current) < cint(JOBS_total) then transmit
                job_count = job_count + 1
            Loop until cint(JOBS_panel_current) = cint(JOBS_total)
        End if
    Next
End If

If ALL_JOBS_PANELS_ARRAY(memb_numb, 0) <> "" Then
    For each_memb = 0 to UBound(ALL_JOBS_PANELS_ARRAY, 2)
        Call Navigate_to_MAXIS_screen("CASE", "NOTE")

        too_old_date = DateAdd("D", -7, CAF_datestamp)
        ALL_JOBS_PANELS_ARRAY(EI_case_note, each_memb) = FALSE

        note_row = 5
        Do
            EMReadScreen note_date, 8, note_row, 6

            EMReadScreen note_title, 55, note_row, 25
            note_title = trim(note_title)

            If left(note_title, 14) = "INCOME DETAIL:" Then
                member_reference = mid(note_title, 17, 2)
                len_emp_name = len(ALL_JOBS_PANELS_ARRAY(employer_name, each_memb))
                jobs_employer_name = mid(note_title, 29, len_emp_name)
                jobs_employer_name = UCase(jobs_employer_name)

                If member_reference = ALL_JOBS_PANELS_ARRAY(memb_numb, each_memb) AND jobs_employer_name = UCase(ALL_JOBS_PANELS_ARRAY(employer_name, each_memb)) Then
                    ALL_JOBS_PANELS_ARRAY(EI_case_note, each_memb) = TRUE
                End If
            End If

            if note_date = "        " then Exit Do
            if ALL_JOBS_PANELS_ARRAY(EI_case_note, each_memb) = TRUE = TRUE then Exit Do

            note_row = note_row + 1
            if note_row = 19 then
                'MsgBox "Next Page" & vbNewLine & "Note Date:" & note_date
                note_row = 5
                PF8
                EMReadScreen check_for_last_page, 9, 24, 14
                If check_for_last_page = "LAST PAGE" Then Exit Do
            End If
            EMReadScreen next_note_date, 8, note_row, 6
            if next_note_date = "        " then Exit Do
        Loop until DateDiff("d", too_old_date, next_note_date) <= 0
    Next
End If

busi_count = 0
call navigate_to_MAXIS_screen("STAT", "BUSI")
EMReadScreen panel_total_check, 6, 2, 73
IF panel_total_check <> "0 Of 0" Then
    For each HH_member in HH_member_array
        EMWriteScreen HH_member, 20, 76
        EMWriteScreen "01", 20, 79
        transmit
        EMReadScreen BUSI_total, 1, 2, 78
        If BUSI_total <> 0 then

            Do
                ReDim Preserve ALL_BUSI_PANELS_ARRAY(budget_explain, busi_count)
                ALL_BUSI_PANELS_ARRAY(memb_numb, busi_count) = HH_member
                ALL_BUSI_PANELS_ARRAY(info_month, busi_count) = MAXIS_footer_month & "/" & MAXIS_footer_year
                call read_BUSI_panel

                EMReadScreen BUSI_panel_current, 1, 2, 73
                ALL_BUSI_PANELS_ARRAY(panel_instance, busi_count) = "0" & BUSI_panel_current

                If cint(BUSI_panel_current) < cint(BUSI_total) then transmit
                busi_count = busi_count + 1
            Loop until cint(BUSI_panel_current) = cint(BUSI_total)

        End if
    Next
Else

End If

'FOR EACH JOB PANEL GO LOOK FOR A RECENT EI CASE NOTE'
call autofill_editbox_from_MAXIS(HH_member_array, "MEDI", INSA)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMI", cit_id)
call autofill_editbox_from_MAXIS(HH_member_array, "OTHR", other_assets)
call autofill_editbox_from_MAXIS(HH_member_array, "PBEN", case_changes)
call autofill_editbox_from_MAXIS(HH_member_array, "PREG", PREG)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "REST", notes_on_rest)
call autofill_editbox_from_MAXIS(HH_member_array, "SCHL", SCHL)
call autofill_editbox_from_MAXIS(HH_member_array, "SECU", other_assets)
call autofill_editbox_from_MAXIS(HH_member_array, "STWK", notes_on_jobs)
call read_TIME_panel
call read_SANC_panel
' call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)

Call read_UNEA_panel

call read_WREG_panel
call update_wreg_and_abawd_notes
'call autofill_editbox_from_MAXIS(HH_member_array, "WREG", notes_on_abawd)

'MAKING THE GATHERED INFORMATION LOOK BETTER FOR THE CASE NOTE
If cash_checkbox = checked then programs_applied_for = programs_applied_for & "CASH, "
If SNAP_checkbox = checked then programs_applied_for = programs_applied_for & "SNAP, "
If EMER_checkbox = checked then programs_applied_for = programs_applied_for & "emergency, "
programs_applied_for = trim(programs_applied_for)
if right(programs_applied_for, 1) = "," then programs_applied_for = left(programs_applied_for, len(programs_applied_for) - 1)

'SHOULD DEFAULT TO TIKLING FOR APPLICATIONS THAT AREN'T RECERTS.
If CAF_type = "Application" then TIKL_checkbox = checked

Call generate_client_list(interview_memb_list, "Select or Type")
Call generate_client_list(shel_memb_list, "Select")
Call generate_client_list(verification_memb_list, "Select or Type Member")
verification_memb_list = " "+chr(9)+verification_memb_list

Call navigate_to_MAXIS_screen("STAT", "AREP")
EMReadScreen version_numb, 1, 2, 73
If version_numb = "1" Then
    EMReadScreen arep_name, 37, 4, 32
    arep_name = replace(arep_name, "_", "")
    interview_memb_list = interview_memb_list+chr(9)+"AREP - " & arep_name
End If


' call verification_dialog
prev_err_msg = ""
notes_on_busi = ""

Do
    Do
        Do
            Do
                Do
                    Do
                        Do
                            Do
                                Do
                                    Do
                                        tab_button = False
                                        full_err_msg = ""
                                        err_array = ""
                                        If show_one = true Then
                                            dlg_len = 285
                                            For the_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                                              If ALL_MEMBERS_ARRAY(gather_detail, the_member) = TRUE Then
                                                  If ALL_MEMBERS_ARRAY(clt_age, the_member) > 17 OR ALL_MEMBERS_ARRAY(memb_numb, the_member) = "01" Then dlg_len = dlg_len + 20
                                              End If
                                            Next

                                            Dialog1 = ""
                                            BeginDialog Dialog1, 0, 0, 465, dlg_len, "CAF Dialog 1 - Personal Information"
                                              If interview_required = TRUE Then Text 5, 10, 300, 10,  "* CAF datestamp:                             * Interview type:"
                                              If interview_required = FALSE Then Text 5, 10, 300, 10, "* CAF datestamp:                             Interview type:"
                                              If interview_required = TRUE Then Text 5, 30, 300, 10,  "* Interview date:                               * How was application received?:"
                                              If interview_required = FALSE Then Text 5, 30, 300, 10, "  Interview date:                               * How was application received?:"
                                              If interview_required = TRUE Then Text 5, 50, 400, 10, "* Interview completed with:                                                                                     If AREP Intvw, ID Info:"
                                              If interview_required = FALSE Then Text 5, 50, 85, 10, "Interview completed with: "

                                              EditBox 60, 5, 50, 15, CAF_datestamp
                                              ComboBox 175, 5, 70, 15, "Select or Type"+chr(9)+"phone"+chr(9)+"office"+chr(9)+interview_type, interview_type
                                              CheckBox 255, 10, 65, 10, "Used Interpreter", Used_Interpreter_checkbox
                                              EditBox 60, 25, 50, 15, interview_date
                                              ComboBox 230, 25, 95, 15, "Select or Type"+chr(9)+"Fax"+chr(9)+"Mail"+chr(9)+"Office"+chr(9)+"Online"+chr(9)+how_app_rcvd, how_app_rcvd
                                              ComboBox 90, 45, 150, 45, interview_memb_list+chr(9)+interview_with, interview_with
                                              ButtonGroup ButtonPressed
                                                PushButton 240, 45, 15, 15, "!", tips_and_tricks_interview_button
                                              If interview_required = TRUE Then EditBox 335, 45, 125, 15, arep_id_info

                                              Text 5, 65, 450, 10, "Member Name                         ID Type                              Detail                                                                                   Required"
                                              y_pos = 80
                                              For the_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                                                If ALL_MEMBERS_ARRAY(gather_detail, the_member) = TRUE Then
                                                    ' MsgBox "Name: " & ALL_MEMBERS_ARRAY(clt_name, the_member) & vbNewLine & "Age: " & ALL_MEMBERS_ARRAY(clt_age, the_member)
                                                    If ALL_MEMBERS_ARRAY(clt_age, the_member) > 17 OR ALL_MEMBERS_ARRAY(memb_numb, the_member) = "01" Then
                                                        If ALL_MEMBERS_ARRAY(memb_numb, the_member) = "01" Then ALL_MEMBERS_ARRAY(id_required, the_member) = checked
                                                        Text 5, y_pos, 85, 10, ALL_MEMBERS_ARRAY(clt_name, the_member)
                                                        ComboBox 100, y_pos - 5, 80, 15, "Type or Select"+chr(9)+"BC - Birth Certificate"+chr(9)+"RE - Religious Record"+chr(9)+"DL - Drivers License/ST ID"+chr(9)+"DV - Divorce Decree"+chr(9)+"AL - Alien Card"+chr(9)+"AD - Arrival//Depart"+chr(9)+"DR - Doctor Stmt"+chr(9)+"PV - Passport/Visa"+chr(9)+"OT - Other Document"+chr(9)+"NO - No Verif Prvd", ALL_MEMBERS_ARRAY(clt_id_verif, the_member)
                                                        EditBox 185, y_pos - 5, 180, 15, ALL_MEMBERS_ARRAY(id_detail, the_member)
                                                        CheckBox 370, y_pos, 90, 10, "ID Verification Required", ALL_MEMBERS_ARRAY(id_required, the_member)
                                                        y_pos = y_pos + 20
                                                    End If
                                                End If
                                              Next
                                              Text 5, y_pos, 25, 10, "Citizen:"
                                              EditBox 35, y_pos -5, 425, 15, cit_id
                                              y_pos = y_pos + 20
                                              EditBox 35, y_pos - 5, 425, 15, IMIG
                                              y_pos = y_pos + 20
                                              EditBox 60, y_pos - 5, 120, 15, AREP
                                              EditBox 270, y_pos - 5, 190, 15, SCHL
                                              y_pos = y_pos + 20
                                              EditBox 60, y_pos - 5, 210, 15, DISA
                                              EditBox 310, y_pos - 5, 150, 15, FACI
                                              y_pos = y_pos + 20
                                              EditBox 35, y_pos - 5, 425, 15, PREG
                                              y_pos = y_pos + 20
                                              EditBox 35, y_pos - 5, 290, 15, ABPS
                                              If trim(ABPS) <> "" AND the_process_for_cash = "Application" Then
                                                Text 335, y_pos, 75, 10, "* Date CS Forms Sent:"
                                              Else
                                                Text 335, y_pos, 75,10, "Date CS Forms Sent:"
                                              End If
                                              EditBox 410, y_pos - 5, 35, 15, CS_forms_sent_date
                                              ButtonGroup ButtonPressed
                                                PushButton 445, y_pos - 5, 15, 15, "!", tips_and_tricks_cs_forms_button
                                              y_pos = y_pos + 20
                                              Text 5, y_pos, 30, 10, "Changes:"
                                              EditBox 40, y_pos - 5, 420, 15, case_changes
                                              y_pos = y_pos + 20 '210'
                                              EditBox 60, y_pos - 5, 385, 15, verifs_needed
                                              Text 10, y_pos + 50, 350, 10, "1 - Personal    |                    |                   |                   |                    |                   |                      |"
                                              ButtonGroup ButtonPressed
                                                PushButton 445, y_pos - 5, 15, 15, "!", tips_and_tricks_verifs_button
                                                PushButton 5, y_pos, 50, 10, "Verifs needed:", verif_button
                                                PushButton 60, y_pos + 50, 35, 10, "2 - JOBS", dlg_two_button
                                                PushButton 100, y_pos + 50, 35, 10, "3 - BUSI", dlg_three_button
                                                PushButton 140, y_pos + 50, 35, 10, "4 - CSES", dlg_four_button
                                                PushButton 180, y_pos + 50, 35, 10, "5 - UNEA", dlg_five_button
                                                PushButton 220, y_pos + 50, 35, 10, "6 - Other", dlg_six_button
                                                PushButton 260, y_pos + 50, 40, 10, "7 - Assets", dlg_seven_button
                                                PushButton 305, y_pos + 50, 50, 10, "8 - Interview", dlg_eight_button
                                                PushButton 370, y_pos + 45, 35, 15, "NEXT", go_to_next_page
                                                CancelButton 410, y_pos + 45, 50, 15
                                                PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
                                                PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
                                                PushButton 335, 25, 45, 10, "next panel", next_panel_button
                                                PushButton 395, 25, 45, 10, "next memb", next_memb_button
                                                PushButton 5, y_pos - 120, 20, 10, "IMIG:", IMIG_button
                                                PushButton 5, y_pos - 100, 25, 10, "AREP/", AREP_button
                                                PushButton 30, y_pos - 100, 25, 10, "ALTP:", ALTP_button
                                                PushButton 190, y_pos - 100, 25, 10, "SCHL/", SCHL_button
                                                PushButton 215, y_pos - 100, 25, 10, "STIN/", STIN_button
                                                PushButton 240, y_pos - 100, 25, 10, "STEC:", STEC_button
                                                PushButton 5, y_pos - 80, 25, 10, "DISA/", DISA_button
                                                PushButton 30, y_pos - 80, 25, 10, "PDED:", PDED_button
                                                PushButton 280, y_pos - 80, 25, 10, "FACI:", FACI_button
                                                PushButton 5, y_pos - 60, 25, 10, "PREG:", PREG_button
                                                PushButton 5, y_pos - 40, 25, 10, "ABPS:", ABPS_button
                                                PushButton 10, y_pos + 25, 20, 10, "DWP", ELIG_DWP_button
                                                PushButton 30, y_pos + 25, 15, 10, "FS", ELIG_FS_button
                                                PushButton 45, y_pos + 25, 15, 10, "GA", ELIG_GA_button
                                                PushButton 60, y_pos + 25, 15, 10, "HC", ELIG_HC_button
                                                PushButton 75, y_pos + 25, 20, 10, "MFIP", ELIG_MFIP_button
                                                PushButton 95, y_pos + 25, 20, 10, "MSA", ELIG_MSA_button
                                                PushButton 130, y_pos + 25, 25, 10, "ADDR", ADDR_button
                                                PushButton 155, y_pos + 25, 25, 10, "MEMB", MEMB_button
                                                PushButton 180, y_pos + 25, 25, 10, "MEMI", MEMI_button
                                                PushButton 205, y_pos + 25, 25, 10, "PROG", PROG_button
                                                PushButton 230, y_pos + 25, 25, 10, "REVW", REVW_button
                                                PushButton 255, y_pos + 25, 25, 10, "SANC", SANC_button
                                                PushButton 280, y_pos + 25, 25, 10, "TIME", TIME_button
                                                PushButton 305, y_pos + 25, 25, 10, "TYPE", TYPE_button
                                                If prev_err_msg <> "" Then PushButton 360, y_pos + 25, 100, 15, "Show Dialog Review Message", dlg_revw_button
                                                OkButton 600, y_pos + 300, 50, 15
                                              GroupBox 5, y_pos + 15, 115, 25, "ELIG panels:"
                                              GroupBox 125, y_pos + 15, 210, 25, "other STAT panels:"
                                              GroupBox 330, 5, 115, 35, "STAT-based navigation"
                                              GroupBox 5, y_pos + 40, 355, 25, "Dialog Tabs"
                                            EndDialog

                                            Dialog Dialog1
                                            cancel_confirmation
                                            MAXIS_dialog_navigation

                                            For the_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                                                If ALL_MEMBERS_ARRAY(id_required, the_member) = checked AND ALL_MEMBERS_ARRAY(clt_id_verif, the_member) = "NO - No Verif Prvd" Then
                                                    verif_text = "Identity for Memb " & ALL_MEMBERS_ARRAY(memb_numb, the_member)
                                                    If InStr(verifs_needed, verif_text) = 0 Then verifs_needed = verifs_needed & "Identity for Memb " & ALL_MEMBERS_ARRAY(memb_numb, the_member) & ".; "
                                                End If
                                            Next

                                            verification_dialog

                                            If ButtonPressed = tips_and_tricks_interview_button Then tips_msg = MsgBox("*** Interview Detail ***" & vbNewLine & vbNewLine & "In order to actually process a CAF for all situations except one, an interview mst be completed. The CAF cannot be processed without an interview. This is why interview information is mandatory." & vbNewLine & vbNewLine &_
                                                                                                                       "An adult cash program ONLY at recertification is the only situation where the interview is not required. Any SNAP or application processing requires an interview." & vbNewLine & vbNewLine &_
                                                                                                                       "If an interview has not been completed, use either Client Contact to indicate the attempt to reach a client for an interview or Application Check to note information about a pending case.", vbInformation, "Tips and Tricks")
                                            If ButtonPressed = tips_and_tricks_cs_forms_button Then tips_msg = MsgBox("*** Date CS Forms Sent ***" & vbNewLine & vbNewLine & "For a Family Cash application and if there is information in the ABPS field, the script will require a date entered here." & vbNewLine & vbNewLine &_
                                                                                                                      "For family cash cases that are being denied enter 'N/A' to have the script bypass this field. Otherwise the date is required here." & vbNewLine & vbNewLine &_
                                                                                                                      "This field can also be used if the forms are given, instead of sent.", vbInformation, "Tips and Tricks")
                                            If ButtonPressed = tips_and_tricks_verifs_button Then tips_msg = MsgBox("*** Verifications Needed ***" & vbNewLine & vbNewLine & "This portion of the script has special functionality. Anytime this field is in a dialog, it is preceeded by a button instead of text." & vbNewLine & "** Press the button to open a special dialog to select verifications." & vbNewLine & vbNewLine &_
                                                                                                                    "Detail about this field/functionality:" & vbNewLine & " - The text '[Information here creates a SEPARATE CASE?NOTE]' can either be deleted or left in place. The script will ignore that phrase when entering a case note. The phrase must be exactly as is for the script to ignore." & vbNewLine &_
                                                                                                                    " - Use a '; ' - semi-colon followed by a space - to have the script go to the next line for the case note - great for formatting the case note." & vbNewLine & " - You can always type directly into the field by the button - you are not required to use the prepared checkboxes on other dialogs." & vbNewLine & vbNewLine &_
                                                                                                                    "VERIFICATIONS ARE ENTERED IN A SEPARATE CASE/NOTE. Do not list other case information in this field. Use 'Other Notes' or fields specific to the information to add.", vbInformation, "Tips and Tricks")
                                            ' If ButtonPressed = tips_and_tricks_interview_button Then ButtonPressed = dlg_one_button
                                            ' If ButtonPressed = tips_and_tricks_cs_forms_button Then ButtonPressed = dlg_one_button
                                            ' If ButtonPressed = tips_and_tricks_verifs_button Then ButtonPressed = dlg_one_button
                                            If ButtonPressed = dlg_revw_button THen Call display_errors(prev_err_msg, FALSE)
                                            If ButtonPressed = -1 Then ButtonPressed = go_to_next_page

                                            Call assess_button_pressed
                                            If ButtonPressed = go_to_next_page Then pass_one = true
                                            If ButtonPressed = verif_button then
                                                pass_one = false
                                                show_one = true
                                            End If
                                            Dim Dialog1
                                        End If
                                    Loop Until pass_one = TRUE
                                    If show_two = true Then
                                        all_jobs = UBound(ALL_JOBS_PANELS_ARRAY, 2)
                                        all_jobs = all_jobs + 1
                                        jobs_pages = all_jobs/3
                                        If jobs_pages <> Int(jobs_pages) Then jobs_pages = Int(jobs_pages) + 1

                                        each_job = 0
                                        loop_start = 0
                                        job_limit = 2
                                        Do
                                            last_job_reviewed = FALSE

                                            dlg_len = 85
                                            jobs_grp_len = 80
                                            length_factor = 80
                                            If snap_checkbox = checked Then length_factor = length_factor + 20
                                            If grh_checkbox = checked Then length_factor = length_factor + 20
                                            If ALL_JOBS_PANELS_ARRAY(memb_numb, 0) = "" Then
                                                dlg_len = 100
                                            Else
                                                If UBound(ALL_JOBS_PANELS_ARRAY, 2) >= job_limit Then
                                                    dlg_len = 325
                                                    If snap_checkbox = checked Then dlg_len = dlg_len + 60
                                                    If grh_checkbox = checked Then dlg_len = dlg_len + 60
                                                    'jobs_grp_len = 315
                                                Else
                                                    dlg_len = length_factor * (UBound(ALL_JOBS_PANELS_ARRAY, 2) - loop_start + 1) + dlg_len
                                                    'jobs_grp_len = 100 * (UBound(ALL_JOBS_PANELS_ARRAY, 2) - loop_start + 1) + 15
                                                End If
                                            End If
                                            If snap_checkbox = checked Then jobs_grp_len = jobs_grp_len + 20
                                            If grh_checkbox = checked Then jobs_grp_len = jobs_grp_len + 20
                                            ' each_job = loop_start
                                            ' Do
                                            '     dlg_len = dlg_len + 100
                                            '     jobs_grp_len = jobs_grp_len + 100
                                            '     if each_job = job_limit Then Exit Do
                                            '     each_job = each_job + 1
                                            ' Loop until each_job = UBound(ALL_JOBS_PANELS_ARRAY, 2)
                                            y_pos = 5
                                            'MsgBox dlg_len
                                            Dialog1 = ""
                                            BeginDialog Dialog1, 0, 0, 705, dlg_len, "CAF Dialog 2 - JOBS Information"
                                              'GroupBox 5, 5, 595, jobs_grp_len, "Earned Income"
                                              If ALL_JOBS_PANELS_ARRAY(memb_numb, 0) = "" Then
                                                y_pos = y_pos + 5
                                                Text 10, y_pos, 590, 10, "There are no JOBS panels found on this case. The script could not pull JOBS details for a case note."
                                                Text 10, y_pos + 10, 590, 10, " ** If this case has income from job source(s) it is best to add the JOBS panels before running this script. **"
                                                Text 10, y_pos + 30, 50, 10, "JOBS Details:"
                                                EditBox 55, y_pos + 25, 545, 15, notes_on_jobs
                                                y_pos = y_pos + 50
                                              Else
                                                  each_job = loop_start
                                                  Do
                                                      GroupBox 5, y_pos, 695, jobs_grp_len, "Member " & ALL_JOBS_PANELS_ARRAY(memb_numb, each_job) & " - " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job)
                                                      Text 180, y_pos, 200, 10, "Verif: " & ALL_JOBS_PANELS_ARRAY(verif_code, each_job)
                                                      CheckBox 365, y_pos, 220, 10, "Check here if this income is not verified and is only an estimate.", ALL_JOBS_PANELS_ARRAY(estimate_only, each_job)
                                                      y_pos = y_pos + 20
                                                      IF ALL_JOBS_PANELS_ARRAY(EI_case_note, each_job) = TRUE Then
                                                        Text 15, y_pos, 690, 10, "Verification:                                                                                                                                                   EARNED INCOME BUDGETING CASE NOTE FOUND                                             to list of verifs needed."
                                                      Else
                                                        Text 15, y_pos, 690, 10, "Verification:                                                                                                                                                                                                                                                                                       to list of verifs needed."
                                                      End If
                                                      EditBox 65, y_pos - 5, 250, 15, ALL_JOBS_PANELS_ARRAY(verif_explain, each_job)
                                                      CheckBox 595, y_pos-10, 100, 10, "Check here to add this JOB", ALL_JOBS_PANELS_ARRAY(verif_checkbox, each_job)
                                                      y_pos = y_pos + 20
                                                      Text 15, y_pos, 600, 10, "Hourly Wage:                              Retro - Income:                              Hours:                                   Prosp - Income:                               Hours:                  Pay Freq:"
                                                      EditBox 65, y_pos - 5, 40, 15, ALL_JOBS_PANELS_ARRAY(hrly_wage, each_job)
                                                      EditBox 170, y_pos - 5, 45, 15, ALL_JOBS_PANELS_ARRAY(job_retro_income, each_job)
                                                      EditBox 250, y_pos - 5, 20, 15, ALL_JOBS_PANELS_ARRAY(retro_hours, each_job)
                                                      EditBox 370, y_pos - 5, 45, 15, ALL_JOBS_PANELS_ARRAY(job_prosp_income, each_job)
                                                      EditBox 450, y_pos - 5, 20, 15, ALL_JOBS_PANELS_ARRAY(prosp_hours, each_job)
                                                      ComboBox 520, y_pos - 5, 60, 45, "Type or select"+chr(9)+"Weekly"+chr(9)+"Biweekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly"+chr(9)+ALL_JOBS_PANELS_ARRAY(main_pay_freq, each_job), ALL_JOBS_PANELS_ARRAY(main_pay_freq, each_job)
                                                      y_pos = y_pos + 20
                                                      If snap_checkbox = checked Then
                                                          Text 15, y_pos, 600, 10, "SNAP PIC:   * Pay Date Amount:                                                                          * Prospective Amount:                                               Calculated:"
                                                          EditBox 125, y_pos - 5, 50, 15, ALL_JOBS_PANELS_ARRAY(pic_pay_date_income, each_job)
                                                          ComboBox 185, y_pos - 5, 60, 45, "Type or select"+chr(9)+"Weekly"+chr(9)+"Biweekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly"+chr(9)+ALL_JOBS_PANELS_ARRAY(pic_pay_freq, each_job), ALL_JOBS_PANELS_ARRAY(pic_pay_freq, each_job)
                                                          EditBox 340, y_pos - 5, 60, 15, ALL_JOBS_PANELS_ARRAY(pic_prosp_income, each_job)
                                                          EditBox 470, y_pos - 5, 50, 15, ALL_JOBS_PANELS_ARRAY(pic_calc_date, each_job)
                                                          y_pos = y_pos + 20
                                                      End If
                                                      If grh_checkbox = checked Then
                                                          Text 15, y_pos, 35, 10, "GRH PIC:"
                                                          Text 65, y_pos, 60, 10, "Pay Date Amount: "
                                                          EditBox 125, y_pos - 5, 50, 15, ALL_JOBS_PANELS_ARRAY(grh_pay_day_income, each_job)
                                                          ComboBox 185, y_pos - 5, 60, 45, "Type or select"+chr(9)+"Weekly"+chr(9)+"Biweekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly"+chr(9)+chr(9)+ALL_JOBS_PANELS_ARRAY(grh_pay_freq, each_job), ALL_JOBS_PANELS_ARRAY(grh_pay_freq, each_job)
                                                          Text 265, y_pos, 70, 10, "Prospective Amount:"
                                                          EditBox 340, y_pos - 5, 60, 15, ALL_JOBS_PANELS_ARRAY(grh_prosp_income, each_job)
                                                          Text 420, y_pos, 40, 10, "Calculated:"
                                                          EditBox 470, y_pos - 5, 50, 15, ALL_JOBS_PANELS_ARRAY(grh_calc_date, each_job)
                                                          y_pos = y_pos + 20
                                                      End If
                                                      If ALL_JOBS_PANELS_ARRAY(EI_case_note, each_job) = FALSE Then
                                                        Text 10, y_pos, 55, 10, "* Explain Budget:"
                                                      Else
                                                        Text 15, y_pos, 55, 10, "Explain Budget:"
                                                      End If
                                                      EditBox 70, y_pos - 5, 620, 15, ALL_JOBS_PANELS_ARRAY(budget_explain, each_job)
                                                      y_pos = y_pos + 25
                                                      if each_job = job_limit Then Exit Do
                                                      each_job = each_job + 1
                                                  Loop until each_job = UBound(ALL_JOBS_PANELS_ARRAY, 2) + 1
                                                  Text 10, y_pos + 5, 70, 40, "JOBS Details:                                              Other Earned Income:"
                                                  EditBox 65, y_pos, 620, 15, notes_on_jobs
                                                  Y_pos = y_pos + 20
                                                  If prev_err_msg <> "" Then
                                                    EditBox 85, y_pos, 510, 15, earned_income
                                                  Else
                                                    EditBox 85, y_pos, 615, 15, earned_income
                                                  End If
                                                  y_pos = y_pos + 25
                                              End If
                                              y_pos = y_pos + 5
                                              GroupBox 5, y_pos - 10, 355, 25, "Dialog Tabs"
                                              Text 10, y_pos, 300, 10, "                       |   2 - JOBS   |                   |                   |                    |                   |                      |"

                                              ButtonGroup ButtonPressed
                                                PushButton 685, y_pos - 50, 15, 15, "!", tips_and_tricks_jobs_button
                                                If prev_err_msg <> "" Then PushButton 600, y_pos - 30, 100, 15, "Show Dialog Review Message", dlg_revw_button
                                                PushButton 10, y_pos, 45, 10, "1 - Personal", dlg_one_button
                                                PushButton 100, y_pos, 35, 10, "3 - BUSI", dlg_three_button
                                                PushButton 140, y_pos, 35, 10, "4 - CSES", dlg_four_button
                                                PushButton 180, y_pos, 35, 10, "5 - UNEA", dlg_five_button
                                                PushButton 220, y_pos, 35, 10, "6 - Other", dlg_six_button
                                                PushButton 260, y_pos, 40, 10, "7 - Assets", dlg_seven_button
                                                PushButton 305, y_pos, 50, 10, "8 - Interview", dlg_eight_button

                                                If jobs_pages >= 2 Then
                                                    If jobs_pages = 2 Then
                                                        If loop_start = 0 Then
                                                            Text 420, y_pos, 15, 10, "1"
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                        ElseIf loop_start = 3 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            Text 440, y_pos, 15, 10, "2"
                                                        End If
                                                    ElseIf jobs_pages = 3 Then
                                                        If loop_start = 0 Then
                                                            Text 420, y_pos, 15, 10, "1"
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                        ElseIf loop_start = 3 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            Text 435, y_pos, 15, 10, "2"
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                        ElseIf loop_start = 6 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            Text 450, y_pos, 15, 10, "3"
                                                        End If
                                                    ElseIf jobs_pages = 4 Then
                                                        If loop_start = 0 Then
                                                            Text 420, y_pos, 15, 10, "1"
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                        ElseIf loop_start = 3 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            Text 435, y_pos, 15, 10, "2"
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                        ElseIf loop_start = 6 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            Text 450, y_pos, 15, 10, "3"
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                        ElseIf loop_start = 9 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            Text 465, y_pos, 15, 10, "4"
                                                        End If
                                                    ElseIf jobs_pages = 5 Then
                                                        If loop_start = 0 Then
                                                            Text 420, y_pos, 15, 10, "1"
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                            PushButton 480, y_pos, 15, 10, "5", jobs_page_five
                                                        ElseIf loop_start = 3 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            Text 435, y_pos, 15, 10, "2"
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                            PushButton 480, y_pos, 15, 10, "5", jobs_page_five
                                                        ElseIf loop_start = 6 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            Text 450, y_pos, 15, 10, "3"
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                            PushButton 480, y_pos, 15, 10, "5", jobs_page_five
                                                        ElseIf loop_start = 9 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            Text 465, y_pos, 15, 10, "4"
                                                            PushButton 480, y_pos, 15, 10, "5", jobs_page_five
                                                        ElseIf loop_start = 12 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                            Text 480, y_pos, 15, 10, "5"
                                                        End If
                                                    ElseIf jobs_pages = 6 Then
                                                        If loop_start = 0 Then
                                                            Text 420, y_pos, 15, 10, "1"
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                            PushButton 480, y_pos, 15, 10, "5", jobs_page_five
                                                            PushButton 495, y_pos, 15, 10, "6", jobs_page_six
                                                        ElseIf loop_start = 3 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            Text 435, y_pos, 15, 10, "2"
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                            PushButton 480, y_pos, 15, 10, "5", jobs_page_five
                                                            PushButton 495, y_pos, 15, 10, "6", jobs_page_six
                                                        ElseIf loop_start = 6 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            Text 450, y_pos, 15, 10, "3"
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                            PushButton 480, y_pos, 15, 10, "5", jobs_page_five
                                                            PushButton 495, y_pos, 15, 10, "6", jobs_page_six
                                                        ElseIf loop_start = 9 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            Text 465, y_pos, 15, 10, "4"
                                                            PushButton 480, y_pos, 15, 10, "5", jobs_page_five
                                                            PushButton 495, y_pos, 15, 10, "6", jobs_page_six
                                                        ElseIf loop_start = 12 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                            Text 480, y_pos, 15, 10, "5"
                                                            PushButton 495, y_pos, 15, 10, "6", jobs_page_six
                                                        ElseIf loop_start = 15 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                            PushButton 480, y_pos, 15, 10, "5", jobs_page_five
                                                            Text 495, y_pos, 15, 10, "6"
                                                        End If
                                                    End If
                                                End If

                                                PushButton 610, y_pos - 5, 35, 15, "NEXT", go_to_next_page
                                                CancelButton 650, y_pos - 5, 50, 15
                                                OkButton 750, 500, 50, 15
                                            EndDialog

                                            dialog Dialog1
                                            cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
                                            MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)

                                            If ButtonPressed = tips_and_tricks_jobs_button Then tips_msg = MsgBox("*** Entering JOBS Information ***" & vbNewLine & vbNewLine & "* If SNAP is checked, the SNAP specific information is ALWAYS required. We need more detail of earned income information in CASE/NOTE and these fields assist with that documentation." & vbNewLine & vbNewLine &_
                                                                                                                  "* The EXPLAIN BUDGET field is very important as it is where you can detail the conversation you had with the client about the income. This conversation is crucial to correct budgeting of JOBS income." & vbNewLine & vbNewLine &_
                                                                                                                  "* If you run the Earned Income Budgeting for a script prior to using the CAF script. The script will find the Earned Income CASE/NOTE and indicate it on this dialog. If that note is present you do NOT need to complete 'Explain Budget' or the SNAP Information as that has been well detailed by the Earned Income Script." & vbNewLine & vbNewLine &_
                                                                                                                  "* If you check the box at the top of the job information, indicating the information is only an estimate, additional detail in the 'Explain Budget' is not required. However, it is recommended to add additional detail if there was any conversation that occured or if there is specific detail that cannot be captured on JOBS." & vbNewLine & vbNewLine &_
                                                                                                                  "** WHAT TO DO IF A JOB HAS ENDED **" & vbNewLine & "This has come up a lot with all the required fields in the JOBS Dialog." & vbNewLine & vbNewLine &_
                                                                                                                  "* Income end date and STWK will be captured when the script gathers information. They will be listed in the fields in this dialog." & vbNewLine & vbNewLine &_
                                                                                                                  "* All the same fields are still mandatory. Since this JOBS panel exists in MAXIS, we need to address it in the case note. If ongoing income is 0, you can list 0 for the SNAP income. Explain budget can detail information about the job and what the changes are." & vbNewLine & vbNewLine &_
                                                                                                                  "* If this income is no longer budgeted - the panel can be removed. Review program specific information but typically once the job is out of the budget month and a STWK panel exists, the JOBS can be deleted. If the panel does not exist - then no detail would need to be entered about the job. (The panel must be deleted PRIOR to the script run.)" & vbNewLine & vbNewLine &_
                                                                                                                  "Generally, we have too little information about earned income in our CASE/NOTEs, this dialog guides you through adding sufficient detail about earned inocme and how it should be budgeted. The more information - the better, so use all applicable and available fields and explain IN FULL.", vbInformation, "Tips and Tricks")
                                            If ButtonPressed = dlg_revw_button THen Call display_errors(prev_err_msg, FALSE)
                                            If ButtonPressed = tips_and_tricks_jobs_button Then ButtonPressed = dlg_two_button
                                            If ButtonPressed = -1 Then ButtonPressed = go_to_next_page
                                            If each_job >= UBound(ALL_JOBS_PANELS_ARRAY, 2) Then last_job_reviewed = TRUE

                                            Call assess_button_pressed
                                            If tab_button = TRUE Then last_job_reviewed = TRUE
                                            If ButtonPressed = go_to_next_page AND last_job_reviewed = TRUE Then pass_two = true

                                            job_limit = job_limit + 3
                                            loop_start = loop_start + 3
                                            If ButtonPressed = jobs_page_one Then
                                                loop_start = 0
                                                job_limit = 2
                                            ElseIf ButtonPressed = jobs_page_two Then
                                                loop_start = 3
                                                job_limit = 5
                                            ElseIf ButtonPressed = jobs_page_three Then
                                                loop_start = 6
                                                job_limit = 8
                                            ElseIf ButtonPressed = jobs_page_four Then
                                                loop_start = 9
                                                job_limit = 11
                                            ElseIf ButtonPressed = jobs_page_five Then
                                                loop_start = 12
                                                job_limit = 14
                                            ElseIf ButtonPressed = jobs_page_six Then
                                                loop_start = 15
                                                job_limit = 17
                                            End If
                                        Loop until last_job_reviewed = TRUE

                                        For each_job = o to UBound(ALL_JOBS_PANELS_ARRAY, 2)
                                            If ALL_JOBS_PANELS_ARRAY(verif_checkbox, each_job) = checked Then
                                                If ALL_JOBS_PANELS_ARRAY(verif_added, each_job) <> TRUE Then verifs_needed = verifs_needed & "Income for Memb " & ALL_JOBS_PANELS_ARRAY(memb_numb, each_job) & " at " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job) & ".; "
                                                ALL_JOBS_PANELS_ARRAY(verif_added, each_job) = TRUE
                                            End If
                                        Next

                                    End If
                                Loop Until pass_two = true
                                If show_three = true Then
                                    all_busi = UBound(ALL_BUSI_PANELS_ARRAY, 2)
                                    all_busi = all_busi + 1
                                    busi_pages = all_busi
                                    If busi_pages <> Int(busi_pages) Then busi_pages = Int(busi_pages) + 1

                                    each_busi = 0
                                    loop_start = 0
                                    last_busi_reviewed = FALSE
                                    busi_limit = 0
                                    Do
                                        dlg_len = 65
                                        busi_grp_len = 145
                                        length_factor = 140
                                        If snap_checkbox = checked Then length_factor = length_factor + 60
                                        If cash_checkbox = checked OR EMER_checkbox = checked Then length_factor = length_factor + 40
                                        'NEED HANDLING FOR IF NO JOBS'
                                        If ALL_BUSI_PANELS_ARRAY(memb_numb, 0) = "" Then
                                            dlg_len = 80
                                        Else
                                            dlg_len = dlg_len + length_factor
                                            ' If UBound(ALL_busi_PANELS_ARRAY, 2) >= busi_limit Then
                                            '     dlg_len = dlg_len + 65
                                            '     If snap_checkbox = checked Then dlg_len = dlg_len + 60
                                            '     If cash_checkbox = checked OR EMER_checkbox = checked Then dlg_len = dlg_len + 60
                                            '     'busi_grp_len = 315
                                            ' Else
                                            '     dlg_len = length_factor * (UBound(ALL_BUSI_PANELS_ARRAY, 2) - loop_start + 1) + 65
                                            '     'busi_grp_len = 100 * (UBound(ALL_BUSI_PANELS_ARRAY, 2) - loop_start + 1) + 15
                                            ' End If
                                        End If
                                        If snap_checkbox = checked Then busi_grp_len = busi_grp_len + 60
                                        If cash_checkbox = checked OR EMER_checkbox = checked Then busi_grp_len = busi_grp_len + 40
                                        ' each_busi = loop_start
                                        ' Do
                                        '     dlg_len = dlg_len + 100
                                        '     busi_grp_len = busi_grp_len + 100
                                        '     if each_busi = busi_limit Then Exit Do
                                        '     each_busi = each_busi + 1
                                        ' Loop until each_busi = UBound(ALL_BUSI_PANELS_ARRAY, 2)
                                        y_pos = 5

                                        BeginDialog Dialog1, 0, 0, 546, dlg_len, "CAF Dialog 3 - BUSI"
                                          If ALL_BUSI_PANELS_ARRAY(memb_numb, 0) = "" Then
                                            Text 10, y_pos, 535, 10, "There are no BUSI panels found on this case. The script could not pull BUSI details for a case note."
                                            Text 10, y_pos + 10, 535, 10, " ** If this case has income from self employment it is best to add the BUSI panels before running this script. **"
                                            Text 10, y_pos + 30, 50, 10, "BUSI Details:"
                                            EditBox 65, y_pos + 25, 475, 15, notes_on_busi
                                            y_pos = u_pos + 50
                                          Else
                                              each_busi = loop_start
                                              Do
                                                  GroupBox 5, y_pos, 535, busi_grp_len, "Member " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & "    Type: " & ALL_BUSI_PANELS_ARRAY(busi_type, each_busi)
                                                  CheckBox 290, y_pos, 220, 10, "Check here if this income is not verified and is only an estimate.", ALL_BUSI_PANELS_ARRAY(estimate_only, each_busi)
                                                  y_pos = y_pos + 20
                                                  Text 15, y_pos, 60, 10, "BUSI Description:"
                                                  EditBox 80, y_pos - 5, 445, 15, ALL_BUSI_PANELS_ARRAY(busi_desc, each_busi)
                                                  y_pos = y_pos + 20
                                                  Text 15, y_pos, 55, 10, "BUSI Structure:"
                                                  ComboBox 75, y_pos - 5, 150, 45, "Select or Type"+chr(9)+"Sole Proprietor"+chr(9)+"Partnership"+chr(9)+"LLC"+chr(9)+"S Corp"+chr(9)+chr(9)+ALL_BUSI_PANELS_ARRAY(busi_structure, each_busi), ALL_BUSI_PANELS_ARRAY(busi_structure, each_busi)
                                                  Text 245, y_pos, 55, 10, "Ownership Share"
                                                  EditBox 305, y_pos - 5, 20, 15, ALL_BUSI_PANELS_ARRAY(share_num, each_busi)
                                                  Text 325, y_pos, 5, 10, "/"
                                                  EditBox 330, y_pos - 5, 20, 15, ALL_BUSI_PANELS_ARRAY(share_denom, each_busi)
                                                  Text 365, y_pos, 50, 10, "Partners in HH:"
                                                  EditBox 420, y_pos - 5, 105, 15, ALL_BUSI_PANELS_ARRAY(partners_in_HH, each_busi)
                                                  y_pos = y_pos + 20
                                                  Text 15, y_pos, 90, 10, "* Self Employment Method:"
                                                  DropListBox 105, y_pos - 5, 120, 45, "Select One"+chr(9)+"50% Gross Inc"+chr(9)+"Tax Forms", ALL_BUSI_PANELS_ARRAY(calc_method, each_busi)
                                                  Text 240, y_pos, 45, 10, "Choice Date:"
                                                  EditBox 290, y_pos - 5, 50, 15, ALL_BUSI_PANELS_ARRAY(mthd_date, each_busi)
                                                  CheckBox 350, y_pos, 185, 10, "Check here if SE Method was discussed with client", ALL_BUSI_PANELS_ARRAY(method_convo_checkbox, each_busi)
                                                  y_pos = y_pos + 20
                                                  Text 15, y_pos, 200, 10, "Reported Hours:     Retro-                     Prosp-"
                                                  EditBox 100, y_pos - 5, 25, 15, ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, each_busi)
                                                  EditBox 160, y_pos - 5, 25, 15, ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, each_busi)
                                                  Text 205, y_pos, 300, 10, "Minimum Wage Hours:      Retro-                    Prosp-                    Income Start Date:"
                                                  EditBox 315, y_pos - 5, 25, 15, ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, each_busi)
                                                  EditBox 375, y_pos - 5, 25, 15, ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, each_busi)
                                                  EditBox 470, y_pos - 5, 50, 15, ALL_BUSI_PANELS_ARRAY(start_date, each_busi)
                                                  y_pos = y_pos + 20
                                                  If SNAP_checkbox = checked Then
                                                      Text 15, y_pos, 200, 10, "SNAP:          Gross Income:      Retro-"
                                                      EditBox 140, y_pos - 5, 45, 15, ALL_BUSI_PANELS_ARRAY(income_ret_snap, each_busi)
                                                      Text 195, y_pos, 25, 10, "Prosp-"
                                                      EditBox 225, y_pos - 5, 45, 15, ALL_BUSI_PANELS_ARRAY(income_pro_snap, each_busi)
                                                      Text 295, y_pos, 100, 10, "Expenses:      Retro-"
                                                      EditBox 365, y_pos - 5, 45, 15, ALL_BUSI_PANELS_ARRAY(expense_ret_snap, each_busi)
                                                      Text 420, y_pos, 25, 10, "Prosp-"
                                                      EditBox 450, y_pos - 5, 45, 15, ALL_BUSI_PANELS_ARRAY(expense_pro_snap, each_busi)
                                                      y_pos = y_pos + 20
                                                      Text 50, y_pos, 85, 10, "* Expenses not allowed:"
                                                      EditBox 140, y_pos - 5, 355, 15, ALL_BUSI_PANELS_ARRAY(exp_not_allwd, each_busi)
                                                      y_pos = y_pos + 20
                                                      Text 115, y_pos, 25, 10, "Verif:"
                                                      ComboBox 140, y_pos - 5, 130, 45, "Select or Type"+chr(9)+"Income Tax Returns"+chr(9)+"Receipts of Sales/Purch"+chr(9)+"Busi Records/Ledger"+chr(9)+"Pend Out State Verif"+chr(9)+"Other Document"+chr(9)+"No Verif Provided"+chr(9)+"Delayed Verif"+chr(9)+"Blank"+chr(9)+ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi), ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi)
                                                      Text 340, y_pos, 25, 10, "Verif:"
                                                      ComboBox 365, y_pos - 5, 130, 45, "Select or Type"+chr(9)+"Income Tax Returns"+chr(9)+"Receipts of Sales/Purch"+chr(9)+"Busi Records/Ledger"+chr(9)+"Pend Out State Verif"+chr(9)+"Other Document"+chr(9)+"No Verif Provided"+chr(9)+"Delayed Verif"+chr(9)+"Blank"+chr(9)+ALL_BUSI_PANELS_ARRAY(snap_expense_verif, each_busi), ALL_BUSI_PANELS_ARRAY(snap_expense_verif, each_busi)
                                                      y_pos = y_pos + 20
                                                  End If
                                                  If cash_checkbox = checked OR EMER_checkbox = checked Then
                                                      Text 15, y_pos, 200, 10, "Cash/Emer:    Gross Income:      Retro-"
                                                      EditBox 140, y_pos - 5, 45, 15, ALL_BUSI_PANELS_ARRAY(income_ret_cash, each_busi)
                                                      Text 195, y_pos, 25, 10, "Prosp-"
                                                      EditBox 225, y_pos - 5, 45, 15, ALL_BUSI_PANELS_ARRAY(income_pro_cash, each_busi)
                                                      Text 295, y_pos, 100, 10, "Expenses:     Retro-"
                                                      EditBox 365, y_pos - 5, 45, 15, ALL_BUSI_PANELS_ARRAY(expense_ret_cash, each_busi)
                                                      Text 420, y_pos, 25, 10, "Prosp-"
                                                      EditBox 450, y_pos - 5, 45, 15, ALL_BUSI_PANELS_ARRAY(expense_pro_cash, each_busi)
                                                      y_pos = y_pos + 20
                                                      Text 115, y_pos, 25, 10, "Verif:"
                                                      ComboBox 140, y_pos - 5, 130, 45, "Select or Type"+chr(9)+"Income Tax Returns"+chr(9)+"Receipts of Sales/Purch"+chr(9)+"Busi Records/Ledger"+chr(9)+"Pend Out State Verif"+chr(9)+"Other Document"+chr(9)+"No Verif Provided"+chr(9)+"Delayed Verif"+chr(9)+"Blank"+chr(9)+ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi), ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi)
                                                      Text 340, y_pos, 25, 10, "Verif:"
                                                      ComboBox 365, y_pos - 5, 130, 45, "Select or Type"+chr(9)+"Income Tax Returns"+chr(9)+"Receipts of Sales/Purch"+chr(9)+"Busi Records/Ledger"+chr(9)+"Pend Out State Verif"+chr(9)+"Other Document"+chr(9)+"No Verif Provided"+chr(9)+"Delayed Verif"+chr(9)+"Blank"+chr(9)+ALL_BUSI_PANELS_ARRAY(cash_expense_verif, each_busi), ALL_BUSI_PANELS_ARRAY(cash_expense_verif, each_busi)
                                                      y_pos = y_pos + 20
                                                  End If
                                                  Text 15, y_pos, 65, 10, "Verification Detail:"
                                                  EditBox 80, y_pos - 5, 445, 15, ALL_BUSI_PANELS_ARRAY(verif_explain, each_busi)
                                                  y_pos = y_pos + 15
                                                  CheckBox 80, y_pos, 400, 10, "Check here if verification about this Self Employment is requested.", ALL_BUSI_PANELS_ARRAY(verif_checkbox, each_busi)
                                                  y_pos = y_pos + 15
                                                  Text 15, y_pos, 60, 10, "* Explain Budget:"
                                                  EditBox 80, y_pos - 5, 445, 15, ALL_BUSI_PANELS_ARRAY(budget_explain, each_busi)
                                                  y_pos = y_pos + 25
                                                  if each_busi = busi_limit Then Exit Do
                                                  each_busi = each_busi + 1
                                              Loop until each_busi = UBound(ALL_BUSI_PANELS_ARRAY, 2) + 1
                                              Text 10, y_pos, 50, 10, "BUSI Details:"
                                              If prev_err_msg <> "" Then
                                                EditBox 60, y_pos - 5, 360, 15, notes_on_busi
                                              Else
                                                EditBox 60, y_pos - 5, 465, 15, notes_on_busi
                                              End If
                                              y_pos = y_pos + 20
                                          End If
                                          y_pos = y_pos + 10
                                          GroupBox 5, y_pos - 10, 355, 25, "Dialog Tabs"
                                          Text 10, y_pos, 300, 10, "                       |                    |   3 - BUSI   |                   |                    |                   |                      |"
                                          ButtonGroup ButtonPressed
                                            If prev_err_msg <> "" Then PushButton 425, y_pos - 35, 100, 15, "Show Dialog Review Message", dlg_revw_button
                                            PushButton 10, y_pos, 45, 10, "1 - Personal", dlg_one_button
                                            PushButton 60, y_pos, 35, 10, "2 - JOBS", dlg_two_button
                                            PushButton 140, y_pos, 35, 10, "4 - CSES", dlg_four_button
                                            PushButton 180, y_pos, 35, 10, "5 - UNEA", dlg_five_button
                                            PushButton 220, y_pos, 35, 10, "6 - Other", dlg_six_button
                                            PushButton 260, y_pos, 40, 10, "7 - Assets", dlg_seven_button
                                            PushButton 305, y_pos, 50, 10, "8 - Interview", dlg_eight_button
                                            If busi_pages >= 2 Then
                                                If busi_pages = 2 Then
                                                    If loop_start = 0 Then
                                                        Text 365, y_pos, 15, 10, "1"
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                    ElseIf loop_start = 1 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        Text 380, y_pos, 15, 10, "2"
                                                    End If
                                                ElseIf busi_pages = 3 Then
                                                    If loop_start = 0 Then
                                                        Text 365, y_pos, 15, 10, "1"
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                    ElseIf loop_start = 3 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        Text 380, y_pos, 15, 10, "2"
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                    ElseIf loop_start = 6 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        Text 395, y_pos, 15, 10, "3"
                                                    End If
                                                ElseIf busi_pages = 4 Then
                                                    If loop_start = 0 Then
                                                        Text 365, y_pos, 15, 10, "1"
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                    ElseIf loop_start = 3 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        Text 380, y_pos, 15, 10, "2"
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                    ElseIf loop_start = 6 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        Text 395, y_pos, 15, 10, "3"
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                    ElseIf loop_start = 9 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        Text 410, y_pos, 15, 10, "4"
                                                    End If
                                                ElseIf busi_pages = 5 Then
                                                    If loop_start = 0 Then
                                                        Text 365, y_pos, 15, 10, "1"
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                        PushButton 420, y_pos, 15, 10, "5", busi_page_five
                                                    ElseIf loop_start = 3 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        Text 380, y_pos, 15, 10, "2"
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                        PushButton 420, y_pos, 15, 10, "5", busi_page_five
                                                    ElseIf loop_start = 6 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        Text 395, y_pos, 15, 10, "3"
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                        PushButton 420, y_pos, 15, 10, "5", busi_page_five
                                                    ElseIf loop_start = 9 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        Text 410, y_pos, 15, 10, "4"
                                                        PushButton 420, y_pos, 15, 10, "5", busi_page_five
                                                    ElseIf loop_start = 12 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                        Text 425, y_pos, 15, 10, "5"
                                                    End If
                                                ElseIf busi_pages = 6 Then
                                                    If loop_start = 0 Then
                                                        Text 365, y_pos, 15, 10, "1"
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                        PushButton 420, y_pos, 15, 10, "5", busi_page_five
                                                        PushButton 435, y_pos, 15, 10, "6", busi_page_six
                                                    ElseIf loop_start = 3 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        Text 380, y_pos, 15, 10, "2"
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                        PushButton 420, y_pos, 15, 10, "5", busi_page_five
                                                        PushButton 435, y_pos, 15, 10, "6", busi_page_six
                                                    ElseIf loop_start = 6 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        Text 395, y_pos, 15, 10, "3"
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                        PushButton 420, y_pos, 15, 10, "5", busi_page_five
                                                        PushButton 435, y_pos, 15, 10, "6", busi_page_six
                                                    ElseIf loop_start = 9 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        Text 410, y_pos, 15, 10, "4"
                                                        PushButton 420, y_pos, 15, 10, "5", busi_page_five
                                                        PushButton 435, y_pos, 15, 10, "6", busi_page_six
                                                    ElseIf loop_start = 12 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                        Text 425, y_pos, 15, 10, "5"
                                                        PushButton 435, y_pos, 15, 10, "6", busi_page_six
                                                    ElseIf loop_start = 15 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                        PushButton 420, y_pos, 15, 10, "5", busi_page_five
                                                        Text 440, y_pos, 15, 10, "6"
                                                    End If
                                                End If
                                            End If
                                            PushButton 450, y_pos - 5, 35, 15, "NEXT", go_to_next_page
                                            CancelButton 490, y_pos - 5, 50, 15
                                            PushButton 525, 5, 15, 15, "!", tips_and_tricks_busi_button
                                            OkButton 600, 500, 50, 15
                                        EndDialog


                                        dialog Dialog1
                                        cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
                                        MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)

                                        If ButtonPressed = tips_and_tricks_busi_button Then tips_msg = MsgBox("*** Self_Employment ***" & vbNewLine & vbNewLine & "There is a policy update around Self Employment for SNAP that was went into effect 08/2019. If you are unfamiliar, this dialog will have elements that seem incorrect. Review the new policy on SIR and in the Policy Manuals." & vbNewLine & vbNewLine &_
                                                                                                              "* Business Description - This is not a field in MAXIS and can be used to further identify the self employment in CASE/NOTE. This can assist in the next worker understanding more about this case situation, making budgeting information clear, and make it easier to find documentation of this business in the case file." & vbNewLine & vbNewLine &_
                                                                                                              "* Business Structure, Ownership share, and Partners in Household - these fields also hep with idetifying budgeting and the correct focumentation required and on file for the businees. These fields are not required, but very helpful in a complete documentation." & vbNewLine & vbNewLine &_
                                                                                                              "* SNAP BUSI Budget - The new policy requires that we review TAX forms if that is the verification receivved to identify any allowed tax deductions that are not allowed as a part of SNAP budgeting. This field 'Expenses not Allowed' is required, though if all are allowed, simply use this field to indicate the review was done and all are allowed." & vbNewLine & vbNewLine &_
                                                                                                              "Checking the box that says 'Check here if verification about this Self Employment is requested' will add a line to the 'Verifs Needed' about self employment for this HH Member. Use this instead of typing to pulling up the verification dialog.", vbInformation, "Tips and Tricks")
                                        If ButtonPressed = dlg_revw_button THen Call display_errors(prev_err_msg, FALSE)
                                        If ButtonPressed = tips_and_tricks_busi_button Then ButtonPressed = dlg_three_button
                                        If ButtonPressed = -1 Then ButtonPressed = go_to_next_page

                                        If each_busi >= UBound(ALL_BUSI_PANELS_ARRAY, 2) Then last_busi_reviewed = TRUE
                                        each_busi = loop_start
                                        Do
                                            'busi_err_msg'
                                            'IF THERE IS AN EI CASE NOTE - DON'T WORRY ABOUT MUCH ERR HANDLING
                                            if each_busi = busi_limit Then Exit Do
                                            each_busi = each_busi + 1
                                        Loop until each_busi = UBound(ALL_BUSI_PANELS_ARRAY, 2)

                                        Call assess_button_pressed
                                        If tab_button = TRUE Then last_busi_reviewed = TRUE
                                        If ButtonPressed = go_to_next_page AND last_busi_reviewed = TRUE Then pass_three = true

                                        busi_limit = busi_limit + 1
                                        loop_start = loop_start + 1

                                        If ButtonPressed = busi_page_one Then
                                            loop_start = 0
                                            job_limit = 0
                                        ElseIf ButtonPressed = busi_page_two Then
                                            loop_start = 1
                                            job_limit = 1
                                        ElseIf ButtonPressed = busi_page_three Then
                                            loop_start = 2
                                            job_limit = 2
                                        ElseIf ButtonPressed = busi_page_four Then
                                            loop_start = 3
                                            job_limit = 3
                                        ElseIf ButtonPressed = busi_page_five Then
                                            loop_start = 4
                                            job_limit = 4
                                        ElseIf ButtonPressed = busi_page_six Then
                                            loop_start = 5
                                            job_limit = 5
                                        End If

                                    Loop until last_busi_reviewed = TRUE

                                    For each_job = o to UBound(ALL_BUSI_PANELS_ARRAY, 2)
                                        If ALL_BUSI_PANELS_ARRAY(verif_checkbox, each_job) = checked Then
                                            If ALL_BUSI_PANELS_ARRAY(verif_added, each_job) <> TRUE Then verifs_needed = verifs_needed & "Self Employment Income for Memb " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_job) & ".; "
                                            ALL_BUSI_PANELS_ARRAY(verif_added, each_job) = TRUE
                                        End If
                                    Next

                                End If
                            Loop Until pass_three = true
                            If show_four = true Then
                                show_cses_detail = FALSE
                                group_len = 75
                                'If SNAP_checkbox = checked Then group_len = group_len + 40
                                group_wide = 465
                                If SNAP_checkbox = checked Then group_wide = 765
                                number_of_cs_members = 0
                                For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
                                    If UNEA_INCOME_ARRAY(CS_exists, each_unea_memb) = TRUE Then
                                        ' dlg_four_len = dlg_four_len + 70
                                        'If SNAP_checkbox = checked Then dlg_four_len = dlg_four_len + 40
                                        show_cses_detail = TRUE
                                        number_of_cs_members = number_of_cs_members + 1
                                    End If
                                Next
                                cs_pages = number_of_cs_members/4
                                If cs_pages <> Int(cs_pages) Then cs_pages = Int(cs_pages) + 1
                                If show_cses_detail = FALSE Then dlg_four_len = 100
                                If SNAP_checkbox = checked AND show_cses_detail = TRUE Then
                                    dlg_wide = 775
                                Else
                                    dlg_wide = 480
                                End If

                                loop_start = 0
                                last_cs_reviewed = FALSE
                                cs_limit = 4

                                Do
                                    dlg_four_len = 85
                                    cs_counter = 0
                                    For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
                                        If UNEA_INCOME_ARRAY(CS_exists, each_unea_memb) = TRUE Then
                                            If cs_counter >= loop_start Then dlg_four_len = dlg_four_len + 70
                                            ' MsgBox "Counter - " & cs_counter & vbNewLine & "Limit - " & cs_limit & vbNewLine & "Loop start - " & loop_start & vbNewLine & "Dlg len - " & dlg_four_len
                                            cs_counter = cs_counter + 1
                                        End If
                                        If cs_counter = cs_limit Then Exit For
                                    Next
                                    If show_cses_detail = FALSE Then dlg_four_len = 100
                                    y_pos = 5
                                    ' MsgBox "Number of CS members - " & number_of_cs_members
                                    BeginDialog Dialog1, 0, 0, dlg_wide, dlg_four_len, "Dialog 4 - CSES"
                                      If show_cses_detail = FALSE Then
                                          Text 10, y_pos, 445, 10, "There are no UNEA panels for Child Support (08, 36, 39) and the script could not pull child support detail information."
                                          Text 10, y_pos + 10, 445, 10, " ** If this case has income from child support it is best to add the UNEA panels before running this script. **"
                                          y_pos = y_pos + 30
                                      Else
                                          cs_counter = 0
                                          For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
                                              If UNEA_INCOME_ARRAY(CS_exists, each_unea_memb) = TRUE Then
                                                  If cs_counter >= loop_start Then
                                                      GroupBox 5, y_pos, group_wide, group_len, "Member " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb)
                                                      y_pos = y_pos + 15
                                                      Text 10, y_pos, 260, 10, "Direct Child Support:       Amt/Mo: $                        Notes:"
                                                      EditBox 125, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(direct_CS_amt, each_unea_memb)
                                                      If SNAP_checkbox = checked Then EditBox 195, y_pos - 5, 570, 15, UNEA_INCOME_ARRAY(direct_CS_notes, each_unea_memb)
                                                      If SNAP_checkbox = unchecked Then EditBox 195, y_pos - 5, 270, 15, UNEA_INCOME_ARRAY(direct_CS_notes, each_unea_memb)
                                                      y_pos = y_pos + 20
                                                      If SNAP_checkbox = checked Then
                                                        Text 10, y_pos, 600, 10, "Disb Child Support(36):   Amt/Mo: $                        Notes:                                                                                                        Months to Average:                            Prosp Budg Notes:"
                                                        EditBox 125, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(disb_CS_amt, each_unea_memb)
                                                        EditBox 195, y_pos - 5, 200, 15, UNEA_INCOME_ARRAY(disb_CS_notes, each_unea_memb)
                                                        EditBox 465, y_pos - 5, 50, 15, UNEA_INCOME_ARRAY(disb_CS_months, each_unea_memb)
                                                        EditBox 580, y_pos - 5, 185, 15, UNEA_INCOME_ARRAY(disb_CS_prosp_budg, each_unea_memb)
                                                      Else
                                                        Text 10, y_pos, 250, 10, "Disb Child Support(36):   Amt/Mo: $                        Notes:"
                                                        EditBox 125, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(disb_CS_amt, each_unea_memb)
                                                        EditBox 195, y_pos - 5, 270, 15, UNEA_INCOME_ARRAY(disb_CS_notes, each_unea_memb)
                                                      End If
                                                      y_pos = y_pos + 20

                                                      If SNAP_checkbox = checked Then
                                                        Text 10, y_pos, 600, 10, "Disb CS Arrears(39):        Amt/Mo: $                        Notes:                                                                                                        Months to Average:                            Prosp Budg Notes:"
                                                        EditBox 125, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(disb_CS_arrears_amt, each_unea_memb)
                                                        EditBox 195, y_pos - 5, 200, 15, UNEA_INCOME_ARRAY(disb_CS_arrears_notes, each_unea_memb)
                                                        EditBox 465, y_pos - 5, 50, 15, UNEA_INCOME_ARRAY(disb_CS_arrears_months, each_unea_memb)
                                                        EditBox 580, y_pos - 5, 185, 15, UNEA_INCOME_ARRAY(disb_CS_arrears_budg, each_unea_memb)
                                                      Else
                                                        Text 10, y_pos, 250, 10, "Disb CS Arrears(39):        Amt/Mo: $                        Notes:"
                                                        EditBox 125, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(disb_CS_arrears_amt, each_unea_memb)
                                                        EditBox 195, y_pos - 5, 270, 15, UNEA_INCOME_ARRAY(disb_CS_arrears_notes, each_unea_memb)
                                                      End If
                                                      y_pos = y_pos + 20
                                                  End If
                                                  cs_counter = cs_counter + 1
                                              End If
                                              If cs_counter = cs_limit Then Exit For

                                          Next
                                          y_pos = y_pos + 10
                                      End If
                                      Text 10, y_pos, 60, 10, "Other CSES Detail:"

                                      If SNAP_checkbox = checked AND show_cses_detail = TRUE Then
                                          If prev_err_msg <> "" Then
                                            EditBox 75, y_pos - 5, 580, 15, notes_on_cses
                                            ButtonGroup ButtonPressed
                                              PushButton 660, y_pos - 5, 100, 15, "Show Dialog Review Message", dlg_revw_button
                                          Else
                                            EditBox 75, y_pos - 5, 685, 15, notes_on_cses
                                          End If
                                          y_pos = y_pos + 20
                                          EditBox 60, y_pos - 5, 700, 15, verifs_needed
                                      Else
                                          If prev_err_msg <> "" Then
                                            EditBox 75, y_pos - 5, 290, 15, notes_on_cses
                                            ButtonGroup ButtonPressed
                                              PushButton 370, y_pos - 5, 100, 15, "Show Dialog Review Message", dlg_revw_button
                                          Else
                                            EditBox 75, y_pos - 5, 395, 15, notes_on_cses
                                          End If
                                          y_pos = y_pos + 20
                                          EditBox 60, y_pos - 5, 410, 15, verifs_needed
                                      End If

                                      y_pos = y_pos + 25
                                      GroupBox 10, y_pos - 10, 355, 25, "Dialog Tabs"
                                      Text 15, y_pos, 300, 10, "                       |                    |                  |   4 - CSES   |                    |                   |                      |"
                                      ButtonGroup ButtonPressed
                                        PushButton 5, y_pos - 25, 50, 10, "Verifs needed:", verif_button
                                        PushButton 15, y_pos, 45, 10, "1 - Personal", dlg_one_button
                                        PushButton 65, y_pos, 35, 10, "2 - JOBS", dlg_two_button
                                        PushButton 105, y_pos, 35, 10, "3 - BUSI", dlg_three_button
                                        PushButton 185, y_pos, 35, 10, "5 - UNEA", dlg_five_button
                                        PushButton 225, y_pos, 35, 10, "6 - Other", dlg_six_button
                                        PushButton 265, y_pos, 40, 10, "7 - Assets", dlg_seven_button
                                        PushButton 310, y_pos, 50, 10, "8 - Interview", dlg_eight_button

                                        If cs_pages >= 2 Then
                                            If cs_pages = 2 Then
                                                If loop_start = 0 Then
                                                    Text 375, y_pos, 15, 10, "1"
                                                    PushButton 385, y_pos, 15, 10, "2", cs_page_two
                                                ElseIf loop_start = 4 Then
                                                    PushButton 370, y_pos, 15, 10, "1", cs_page_one
                                                    Text 390, y_pos, 15, 10, "2"
                                                End If
                                            ElseIf cs_pages = 3 Then
                                                If loop_start = 0 Then
                                                    Text 365, y_pos, 15, 10, "1"
                                                    PushButton 375, y_pos, 15, 10, "2", cs_page_two
                                                    PushButton 390, y_pos, 15, 10, "3", cs_page_three
                                                ElseIf loop_start = 4 Then
                                                    PushButton 360, y_pos, 15, 10, "1", cs_page_one
                                                    Text 380, y_pos, 15, 10, "2"
                                                    PushButton 390, y_pos, 15, 10, "3", cs_page_three
                                                ElseIf loop_start = 8 Then
                                                    PushButton 360, y_pos, 15, 10, "1", cs_page_one
                                                    PushButton 375, y_pos, 15, 10, "2", cs_page_two
                                                    Text 395, y_pos, 15, 10, "3"
                                                End If
                                            End If
                                        End If

                                        If SNAP_checkbox = checked AND show_cses_detail = TRUE Then
                                            PushButton 700, y_pos - 5, 25, 15, "NEXT", go_to_next_page
                                            CancelButton 730, y_pos - 5, 30, 15
                                        Else
                                            PushButton 410, y_pos - 5, 25, 15, "NEXT", go_to_next_page
                                            CancelButton 440, y_pos - 5, 30, 15
                                        End If
                                        OkButton 700, 700, 50, 15
                                    EndDialog

                                    Dialog Dialog1
                                    cancel_confirmation
                                    verification_dialog
                                    'MsgBox ButtonPressed
                                    If ButtonPressed = dlg_revw_button THen Call display_errors(prev_err_msg, FALSE)
                                    If ButtonPressed = -1 Then ButtonPressed = go_to_next_page
                                    If ButtonPressed = verif_button Then ButtonPressed = dlg_four_button
                                    If cs_counter >= number_of_cs_members Then last_cs_reviewed = TRUE

                                    Call assess_button_pressed
                                    If tab_button = TRUE Then last_cs_reviewed = TRUE
                                    If ButtonPressed = go_to_next_page AND last_cs_reviewed = TRUE Then pass_four = true

                                    cs_limit = cs_limit + 4
                                    loop_start = loop_start + 4
                                    If ButtonPressed = cs_page_one Then
                                        loop_start = 0
                                        cs_limit = 4
                                    ElseIf ButtonPressed = cs_page_two Then
                                        loop_start = 4
                                        cs_limit = 8
                                    ElseIf ButtonPressed = cs_page_three Then
                                        loop_start = 9
                                        cs_limit = 12
                                    End If
                                Loop until last_cs_reviewed = TRUE
                            End If
                        Loop Until pass_four = true
                        If show_five = true Then
                            dlg_five_len = 190
                            ssa_group_len = 30
                            uc_group_len = 30
                            unea_income_found = FALSE
                            For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
                                If UNEA_INCOME_ARRAY(UC_exists, each_unea_memb) = TRUE Then
                                    dlg_five_len = dlg_five_len + 70
                                    uc_group_len = uc_group_len + 80
                                    UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb) = UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb) & ""
                                    unea_income_found = TRUE
                                End If
                                If UNEA_INCOME_ARRAY(SSA_exists, each_unea_memb) = TRUE Then
                                    dlg_five_len = dlg_five_len + 40
                                    ssa_group_len = ssa_group_len + 40
                                    unea_income_found = TRUE
                                End If
                            Next
                            If trim(notes_on_VA_income) <> "" Then unea_income_found = TRUE
                            If trim(notes_on_WC_income) <> "" Then unea_income_found = TRUE
                            If trim(notes_on_other_UNEA) <> "" Then unea_income_found = TRUE
                            If unea_income_found = FALSE Then dlg_five_len = dlg_five_len + 20

                            y_pos = 5
                            BeginDialog Dialog1, 0, 0, 466, dlg_five_len, "Dialog 5 - UNEA"
                              If unea_income_found = FALSE Then
                                  Text 10, y_pos, 445, 10, "There are no UNEA panels found and the script could not pull detail about SSA/WC/VA/UC or other UNEA income."
                                  Text 10, y_pos + 10, 445, 10, " ** If this case has income from SSI, RSDI, or Unemployment it is best to add the UNEA panels before running this script. **"
                                  y_pos = y_pos + 25
                              End If
                              GroupBox 5, y_pos, 455, ssa_group_len, "SSA Income"
                              y_pos = y_pos + 15
                              For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
                                  If UNEA_INCOME_ARRAY(SSA_exists, each_unea_memb) = TRUE Then
                                      Text 15, y_pos, 40, 10, "Member " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb)
                                      Text 60, y_pos, 55, 10, "RSDI: Amount: $"
                                      EditBox 120, y_pos - 5, 30, 15, UNEA_INCOME_ARRAY(UNEA_RSDI_amt, each_unea_memb)
                                      Text 155, y_pos, 30, 10, "* Notes:"
                                      EditBox 185, y_pos - 5, 270, 15, UNEA_INCOME_ARRAY(UNEA_RSDI_notes, each_unea_memb)
                                      y_pos = y_pos + 20
                                      Text 60, y_pos, 55, 10, "SSI: Amount: $"
                                      EditBox 120, y_pos - 5, 30, 15, UNEA_INCOME_ARRAY(UNEA_SSI_amt, each_unea_memb)
                                      Text 155, y_pos, 30, 10, "* Notes:"
                                      EditBox 185, y_pos - 5, 270, 15, UNEA_INCOME_ARRAY(UNEA_SSI_notes, each_unea_memb)
                                      y_pos = y_pos + 20
                                  End If
                              Next
                              Text 10, y_pos, 65, 10, "Other SSA Income:"
                              EditBox 80, y_pos - 5, 375, 15, notes_on_ssa_income
                              y_pos = y_pos + 25
                              Text 5, y_pos, 40, 10, "VA Income:"
                              EditBox 45, y_pos - 5, 415, 15, notes_on_VA_income
                              y_pos = y_pos + 20
                              Text 5, y_pos, 55, 10, "Worker's Comp:"
                              EditBox 60, y_pos - 5, 400, 15, notes_on_WC_income
                              y_pos = y_pos + 15
                              GroupBox 5, y_pos, 455, uc_group_len, "Unemployment Income"
                              y_pos = y_pos + 15
                              For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
                                  If UNEA_INCOME_ARRAY(UC_exists, each_unea_memb) = TRUE Then
                                      UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb) = UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb) & ""
                                      UNEA_INCOME_ARRAY(UNEA_UC_account_balance, each_unea_memb) = UNEA_INCOME_ARRAY(UNEA_UC_account_balance, each_unea_memb) & ""
                                      UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb) = UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb) & ""
                                      UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb) = UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb) & ""
                                      UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, each_unea_memb) = UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, each_unea_memb) & ""

                                      Text 15, y_pos, 40, 10, "Member " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb)
                                      Text 65, y_pos, 120, 10, "Unemployment Start Date: " & UNEA_INCOME_ARRAY(UNEA_UC_start_date, each_unea_memb)
                                      Text 195, y_pos, 95, 10, "* Budgeted Weekly Amount:"
                                      EditBox 290, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb)
                                      Text 345, y_pos, 70, 10, "UC Acct Bal:"
                                      EditBox 395, y_pos - 5, 50, 15, UNEA_INCOME_ARRAY(UNEA_UC_account_balance, each_unea_memb)
                                      y_pos = y_pos + 20
                                      Text 30, y_pos, 50, 10, "Weekly Gross:"
                                      EditBox 85, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb)
                                      Text 130, y_pos, 70, 10, "Allowed Deductions:"
                                      EditBox 200, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb)
                                      Text 245, y_pos, 75, 10, "Excluded Deductions:"
                                      EditBox 320, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, each_unea_memb)
                                      Text 375, y_pos - 5, 80, 15, "Enter a TIKL date to check if UC has ended:"
                                      y_pos = y_pos + 20
                                      Text 30, y_pos, 50, 10, "Retro Income:"
                                      EditBox 80, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(UNEA_UC_retro_amt, each_unea_memb)
                                      Text 130, y_pos, 50, 10, "Prosp Income:"
                                      EditBox 185, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(UNEA_UC_prosp_amt, each_unea_memb)
                                      If SNAP_checkbox = checked Then Text 250, y_pos, 65, 10, "* SNAP Prosp Amt: $"
                                      If SNAP_checkbox = unchecked Then Text 250, y_pos, 65, 10, "SNAP Prosp Amt: $"
                                      EditBox 315, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, each_unea_memb)
                                      ButtonGroup ButtonPressed
                                        PushButton 365, y_pos, 35, 10, "Calc", calc_button
                                      EditBox 405, y_pos - 5, 50, 15, UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb)
                                      y_pos = y_pos + 20
                                      Text 30, y_pos, 25, 10, "Notes:"
                                      EditBox 60, y_pos - 5, 395, 15, UNEA_INCOME_ARRAY(UNEA_UC_notes, each_unea_memb)
                                      y_pos = y_pos + 20
                                  End If
                              Next
                              Text 15, y_pos, 60, 10, "Other UC Income:"
                              EditBox 75, y_pos - 5, 380, 15, other_uc_income_notes
                              y_pos = y_pos + 25
                              Text 10, y_pos, 45, 10, "Other UNEA:"
                              If prev_err_msg <> "" Then
                                EditBox 55, y_pos - 5, 305, 15, notes_on_other_UNEA
                                ButtonGroup ButtonPressed
                                  PushButton 365, y_pos - 5, 100, 15, "Show Dialog Review Message", dlg_revw_button
                              Else
                                EditBox 55, y_pos - 5, 405, 15, notes_on_other_UNEA
                              End If
                              y_pos = y_pos + 20
                              ButtonGroup ButtonPressed
                                PushButton 5, y_pos, 50, 10, "Verifs needed:", verif_button
                              EditBox 60, y_pos - 5, 400, 15, verifs_needed
                              y_pos = y_pos + 25
                              GroupBox 10, y_pos - 10, 355, 25, "Dialog Tabs"
                              Text 15, y_pos, 300, 10, "                       |                    |                   |                   |  5 - UNEA   |                   |                      |"
                              ButtonGroup ButtonPressed
                                PushButton 15, y_pos, 45, 10, "1 - Personal", dlg_one_button
                                PushButton 65, y_pos, 35, 10, "2 - JOBS", dlg_two_button
                                PushButton 105, y_pos, 35, 10, "3 - BUSI", dlg_three_button
                                PushButton 145, y_pos, 35, 10, "4 - CSES", dlg_four_button
                                PushButton 225, y_pos, 35, 10, "6 - Other", dlg_six_button
                                PushButton 265, y_pos, 40, 10, "7 - Assets", dlg_seven_button
                                PushButton 310, y_pos, 50, 10, "8 - Interview", dlg_eight_button
                                PushButton 370, y_pos - 5, 35, 15, "NEXT", go_to_next_page
                                CancelButton 410, y_pos - 5, 50, 15
                                OkButton 600, 500, 50, 15
                            EndDialog

                            Dialog Dialog1
                            cancel_confirmation
                            MAXIS_dialog_navigation
                            verification_dialog

                            If ButtonPressed = dlg_revw_button THen Call display_errors(prev_err_msg, FALSE)
                            If ButtonPressed = -1 Then ButtonPressed = go_to_next_page

                            If ButtonPressed = calc_button Then
                                For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
                                    If UNEA_INCOME_ARRAY(UC_exists, each_unea_memb) = TRUE Then
                                        If IsNumeric(UNEA_INCOME_ARRAY(UNEA_UC_account_balance, each_unea_memb)) = TRUE Then
                                            If IsNumeric(UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb)) = TRUE Then
                                                weeks_of_UC_benefits = Int(UNEA_INCOME_ARRAY(UNEA_UC_account_balance, each_unea_memb)/UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb))
                                                'MsgBox weeks_of_UC_benefits
                                                UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb) = DateAdd("ww", weeks_of_UC_benefits, date)
                                            Else
                                                MsgBox "The scriupt cannot calculate the potential date of UC account balance depletion for Member " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " without the UC account balance and UC Weekly Gross income. Enter these amounts as numbers and the script will enter a date for the TIKL into the dialog. The TIKL date can also be entered or changed manually."
                                            End If
                                        End If
                                    End If
                                Next
                                ButtonPressed = dlg_five_button
                            End If
                            If ButtonPressed = verif_button then ButtonPressed = dlg_five_button

                            Call assess_button_pressed
                            If ButtonPressed = go_to_next_page Then pass_five = true
                        End If
                    Loop Until pass_five = true
                    If show_six = true Then
                        If left(total_shelter_amount, 1) <> "$" Then total_shelter_amount = "$" & total_shelter_amount

                        BeginDialog Dialog1, 0, 0, 556, 290, "CAF Dialog 6 - WREG, Expenses, Address"
                          EditBox 45, 50, 500, 15, notes_on_wreg
                          ButtonGroup ButtonPressed
                            PushButton 440, 30, 105, 15, "Update ABAWD and WREG", abawd_button
                            PushButton 235, 85, 50, 15, "Update SHEL", update_shel_button
                          DropListBox 45, 140, 100, 45, "Select ALLOWED HEST"+chr(9)+"AC/Heat - Full $496"+chr(9)+"AC/Heat - Full $490"+chr(9)+"Electric and Phone - $210"+chr(9)+"Electric and Phone - $192"+chr(9)+"Electric ONLY - $154"+chr(9)+"Electric ONLY - $143"+chr(9)+"Phone ONLY - $56"+chr(9)+"Phone ONLY - $49"+chr(9)+"NONE - $0", hest_information
                          EditBox 180, 140, 110, 15, notes_on_acut
                          EditBox 45, 160, 245, 15, notes_on_coex
                          EditBox 45, 180, 245, 15, notes_on_dcex
                          EditBox 45, 200, 245, 15, notes_on_other_deduction
                          EditBox 45, 220, 245, 15, expense_notes
                          CheckBox 320, 85, 125, 10, "Check here to confirm the address.", address_confirmation_checkbox
                          DropListBox 345, 150, 85, 45, county_list, addr_county
                          DropListBox 480, 150, 30, 45, "No"+chr(9)+"Yes", homeless_yn
                          DropListBox 335, 170, 95, 45, "SF - Shelter Form"+chr(9)+"CO - Coltrl Stmt"+chr(9)+"LE - Lease/Rent Doc"+chr(9)+"MO - Mortgage Papers"+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"CD - Contrct for Deed"+chr(9)+"UT - Utility Stmt"+chr(9)+"DL - Driver Lic/State ID"+chr(9)+"OT - Other Document"+chr(9)+"NO - No Ver Prvd"+chr(9)+"? - Delayed"+chr(9)+"Blank", addr_verif
                          DropListBox 480, 170, 30, 45, "No"+chr(9)+"Yes", reservation_yn
                          DropListBox 375, 190, 165, 45, " "+chr(9)+"01 - Own home, lease or roommate"+chr(9)+"02 - Family/Friends - economic hardship"+chr(9)+"03 -  servc prvdr- foster/group home"+chr(9)+"04 - Hospital/Treatment/Detox/Nursing Home"+chr(9)+"05 - Jail/Prison//Juvenile Det."+chr(9)+"06 - Hotel/Motel"+chr(9)+"07 - Emergency Shelter"+chr(9)+"08 - Place not meant for Housing"+chr(9)+"09 - Declined"+chr(9)+"10 - Unknown"+chr(9)+"Blank", living_situation
                          EditBox 315, 220, 230, 15, notes_on_address
                          EditBox 60, 245, 490, 15, verifs_needed
                          GroupBox 5, 5, 545, 65, "WREG and ABAWD Information"
                          Text 15, 15, 55, 10, "ABAWD Details:"
                          Text 75, 15, 470, 10, notes_on_abawd
                          Text 15, 25, 400, 10, notes_on_abawd_two
                          Text 15, 35, 400, 10, notes_on_abawd_three
                          GroupBox 5, 75, 290, 165, "Expenses and Deductions"
                          Text 15, 90, 50, 10, "Total Shelter:"
                          Text 70, 90, 155, 10, total_shelter_amount
                          Text 10, 105, 285, 10, shelter_details
                          Text 10, 115, 285, 10, shelter_details_two
                          Text 10, 125, 285, 10, shelter_details_three
                          Text 20, 205, 20, 10, "Other:"
                          Text 20, 225, 25, 10, "Notes:"
                          GroupBox 305, 75, 245, 165, "Address"
                          Text 350, 100, 175, 10, addr_line_one
                          If addr_line_two = "" Then
                            Text 350, 115, 175, 10, city & ", " & state & " " & zip
                          Else
                            Text 350, 115, 175, 10, addr_line_two
                            Text 350, 130, 175, 10, city & ", " & state & " " & zip
                          End If
                          Text 315, 155, 25, 10, "County:"
                          Text 440, 155, 35, 10, "Homeless:"
                          Text 315, 175, 20, 10, "Verif:"
                          Text 435, 175, 45, 10, "Reservation:"
                          Text 315, 195, 55, 10, "* Living Situation:"
                          Text 315, 210, 75, 10, "Notes on address:"
                          GroupBox 105, 265, 355, 25, "Dialog Tabs"
                          Text 110, 275, 300, 10, "                       |                    |                   |                    |                    |  6 - Other   |                      |"
                          ButtonGroup ButtonPressed
                            PushButton 5, 250, 50, 10, "Verifs needed:", verif_button
                            If prev_err_msg <> "" Then PushButton 5, 270, 100, 15, "Show Dialog Review Message", dlg_revw_button
                            PushButton 110, 275, 45, 10, "1 - Personal", dlg_one_button
                            PushButton 160, 275, 35, 10, "2 - JOBS", dlg_two_button
                            PushButton 200, 275, 35, 10, "3 - BUSI", dlg_three_button
                            PushButton 240, 275, 35, 10, "4 - CSES", dlg_four_button
                            PushButton 280, 275, 35, 10, "5 - UNEA", dlg_five_button
                            PushButton 360, 275, 40, 10, "7 - Assets", dlg_seven_button
                            PushButton 405, 275, 50, 10, "8 - Interview", dlg_eight_button
                            PushButton 460, 270, 35, 15, "NEXT", go_to_next_page
                            CancelButton 500, 270, 50, 15
                            If SNAP_checkbox = checked Then PushButton 10, 55, 30, 10, "* WREG", wreg_button
                            If SNAP_checkbox = unchecked Then PushButton 10, 55, 25, 10, "WREG", wreg_button
                            PushButton 315, 100, 25, 10, "ADDR", addr_button
                            PushButton 15, 145, 25, 10, "HEST", hest_button
                            PushButton 150, 145, 25, 10, "ACUT", acut_button
                            PushButton 15, 165, 25, 10, "COEX", coex_button
                            PushButton 15, 185, 25, 10, "DCEX", dcex_button
                            OkButton 600, 500, 50, 15
                        EndDialog

                        Dialog Dialog1			'Displays the second dialog
                        cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
                        MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
                        verification_dialog

                        If ButtonPressed = dlg_revw_button THen Call display_errors(prev_err_msg, FALSE)
                        If ButtonPressed = -1 Then ButtonPressed = go_to_next_page

                        If ButtonPressed = abawd_button Then
                            Do
                                abawd_err_msg = ""

                                notes_on_wreg = ""
                                notes_on_abawd = ""
                                notes_on_abawd_two = ""
                                notes_on_abawd_three = ""
                                dlg_len = 40
                                For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                                  If ALL_MEMBERS_ARRAY(include_snap_checkbox, each_member) = checked AND ALL_MEMBERS_ARRAY(wreg_exists, each_member) = TRUE Then
                                    dlg_len = dlg_len + 95
                                  End If
                                Next
                                y_pos = 10
                                BeginDialog Dialog1, 0, 0, 551, dlg_len, "ABAWD Detail"
                                  For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                                    If ALL_MEMBERS_ARRAY(include_snap_checkbox, each_member) = checked AND ALL_MEMBERS_ARRAY(wreg_exists, each_member) = TRUE Then
                                      GroupBox 5, y_pos, 540, 95, "Member " & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " - " & ALL_MEMBERS_ARRAY(clt_name, each_member)
                                      y_pos = y_pos + 20
                                      Text 15, y_pos, 70, 10, "FSET WREG Status:"
                                      DropListBox 90, y_pos - 5, 130, 45, " "+chr(9)+"03  Unfit for Employment"+chr(9)+"04  Responsible for Care of Another"+chr(9)+"05  Age 60+"+chr(9)+"06  Under Age 16"+chr(9)+"07  Age 16-17, live w/ parent"+chr(9)+"08  Care of Child <6"+chr(9)+"09  Employed 30+ hrs/wk"+chr(9)+"10  Matching Grant"+chr(9)+"11  Unemployment Insurance"+chr(9)+"12  Enrolled in School/Training"+chr(9)+"13  CD Program"+chr(9)+"14  Receiving MFIP"+chr(9)+"20  Pend/Receiving DWP"+chr(9)+"15  Age 16-17 not live w/ Parent"+chr(9)+"16  50-59 Years Old"+chr(9)+"21  Care child < 18"+chr(9)+"17  Receiving RCA or GA"+chr(9)+"30  FSET Participant"+chr(9)+"02  Fail FSET Coop"+chr(9)+"33  Non-coop being referred"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(clt_wreg_status, each_member)
                                      Text 230, y_pos, 55, 10, "ABAWD Status:"
                                      DropListBox 285, y_pos - 5, 110, 45, " "+chr(9)+"01  WREG Exempt"+chr(9)+"02  Under Age 18"+chr(9)+"03  Age 50+"+chr(9)+"04  Caregiver of Minor Child"+chr(9)+"05  Pregnant"+chr(9)+"06  Employed 20+ hrs/wk"+chr(9)+"07  Work Experience"+chr(9)+"08  Other E and T"+chr(9)+"09  Waivered Area"+chr(9)+"10  ABAWD Counted"+chr(9)+"11  Second Set"+chr(9)+"12  RCA or GA Participant"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(clt_abawd_status, each_member)
                                      CheckBox 405, y_pos - 5, 130, 10, "Check here if this person is the PWE", ALL_MEMBERS_ARRAY(pwe_checkbox, each_member)
                                      y_pos = y_pos + 20
                                      Text 15, y_pos, 145, 10, "Number of ABAWD months used in past 36:"
                                      EditBox 160, y_pos - 5, 25, 15, ALL_MEMBERS_ARRAY(numb_abawd_used, each_member)
                                      Text 200, y_pos, 95, 10, "List all ABAWD months used:"
                                      EditBox 300, y_pos - 5, 135, 15, ALL_MEMBERS_ARRAY(list_abawd_mo, each_member)
                                      y_pos = y_pos + 20
                                      Text 15, y_pos, 135, 10, "If used, list the first month of Second Set:"
                                      EditBox 155, y_pos - 5, 40, 15, ALL_MEMBERS_ARRAY(first_second_set, each_member)
                                      Text 205, y_pos, 130, 10, "If NOT Eligible for Second Set, Explain:"
                                      EditBox 335, y_pos - 5, 200, 15, ALL_MEMBERS_ARRAY(explain_no_second, each_member)
                                      y_pos = y_pos + 20
                                      'Text 15, y_pos, 115, 10, "Number of BANKED months used:"
                                      'EditBox 130, y_pos - 5, 25, 15, ALL_MEMBERS_ARRAY(numb_banked_mo, each_member)
                                      Text 15, y_pos, 45, 10, "Other Notes:"
                                      EditBox 60, y_pos - 5, 475, 15, ALL_MEMBERS_ARRAY(clt_abawd_notes, each_member)

                                      y_pos = y_pos + 15
                                    End If
                                  Next
                                  y_pos = y_pos + 10
                                  ButtonGroup ButtonPressed
                                    PushButton 455, y_pos, 90, 15, "Return to Main Dialog", return_button
                                    OkButton 600, 500, 50, 15
                                EndDialog

                                Dialog Dialog1

                                If ButtonPressed = -1 Then ButtonPressed = return_button
                                If ButtonPressed = 0 Then ButtonPressed = return_button

                                call update_wreg_and_abawd_notes
                                If ButtonPressed = return_button Then ButtonPressed = dlg_six_button

                            Loop until abawd_err_msg = ""
                        End If

                        If ButtonPressed = update_shel_button Then
                            shel_client = ""
                            For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                                If ALL_MEMBERS_ARRAY(shel_exists, each_member) = TRUE Then
                                    shel_client = each_member
                                    Exit For
                                End If
                            Next
                            If shel_client <> "" Then clt_shel_is_for = ALL_MEMBERS_ARRAY(full_clt, shel_client)
                            'ADD an IF here to determine the right HH member or if one is not yet selected AND preselect the one that has a SHEL'
                            Do
                                shel_err_msg = ""

                                If clt_SHEL_is_for = "Select" Then
                                    dlg_len = 30
                                Else
                                    dlg_len = 250
                                    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                                        If clt_shel_is_for = ALL_MEMBERS_ARRAY(full_clt, each_member) Then
                                            shel_client = each_member
                                            ALL_MEMBERS_ARRAY(shel_exists, each_member) = TRUE
                                        End If
                                    Next
                                End If
                                if shel_client = "" Then
                                    shel_client = 0
                                    clt_shel_is_for = ALL_MEMBERS_ARRAY(full_clt, shel_client)
                                End If

                                If ALL_MEMBERS_ARRAY(shel_subsudized, shel_client) = "" Then ALL_MEMBERS_ARRAY(shel_subsudized, shel_client) = "No"
                                If ALL_MEMBERS_ARRAY(shel_shared, shel_client) = "" Then ALL_MEMBERS_ARRAY(shel_shared, shel_client) = "No"
                                shel_verif_needed_checkbox = unchecked
                                If manual_total_shelter = "" Then manual_total_shelter = total_shelter_amount & ""
                                If manual_amount_used = FALSE Then manual_total_shelter = total_shelter_amount & ""
                                start_total_shel = manual_total_shelter

                                BeginDialog Dialog1, 0, 0, 340, dlg_len, "SHEL Detail Dialog"
                                  DropListBox 60, 10, 125, 45, shel_memb_list, clt_SHEL_is_for
                                  Text 5, 15, 55, 10, "SHEL for Memb"
                                  ButtonGroup ButtonPressed
                                    PushButton 190, 10, 40, 10, "Load", load_button
                                  Text 235, 10, 55, 10, "Total Shelter:"
                                  EditBox 290, 5, 40, 15, manual_total_shelter
                                  If clt_shel_is_for <> "Select" Then
                                      'ALL_MEMBERS_ARRAY
                                      ALL_MEMBERS_ARRAY(shel_retro_rent_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_rent_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_retro_lot_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_lot_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_retro_ins_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_ins_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_retro_tax_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_tax_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_retro_room_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_room_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_prosp_room_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_room_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_retro_garage_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_garage_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, shel_client) & ""

                                      DropListBox 85, 30, 30, 45, "Yes"+chr(9)+"No", ALL_MEMBERS_ARRAY(shel_subsudized, shel_client)
                                      DropListBox 175, 30, 30, 45, "Yes"+chr(9)+"No", ALL_MEMBERS_ARRAY(shel_shared, shel_client)
                                      EditBox 45, 60, 35, 15, ALL_MEMBERS_ARRAY(shel_retro_rent_amt, shel_client)
                                      DropListBox 85, 60, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_retro_rent_verif, shel_client)
                                      EditBox 195, 60, 35, 15, ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, shel_client)
                                      DropListBox 235, 60, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, shel_client)
                                      EditBox 45, 80, 35, 15, ALL_MEMBERS_ARRAY(shel_retro_lot_amt, shel_client)
                                      DropListBox 85, 80, 100, 45, "Select one"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_retro_lot_verif, shel_client)
                                      EditBox 195, 80, 35, 15, ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, shel_client)
                                      DropListBox 235, 80, 100, 45, "Select one"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, shel_client)
                                      EditBox 45, 100, 35, 15, ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, shel_client)
                                      DropListBox 85, 100, 100, 45, "Select one"+chr(9)+"MO - Mort Pmt Book"+chr(9)+"CD - Ctrct For Deed"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, shel_client)
                                      EditBox 195, 100, 35, 15, ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, shel_client)
                                      DropListBox 235, 100, 100, 45, "Select one"+chr(9)+"MO - Mort Pmt Book"+chr(9)+"CD - Ctrct For Deed"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, shel_client)
                                      EditBox 45, 120, 35, 15, ALL_MEMBERS_ARRAY(shel_retro_ins_amt, shel_client)
                                      DropListBox 85, 120, 100, 45, "Select one"+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_retro_ins_verif, shel_client)
                                      EditBox 195, 120, 35, 15, ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, shel_client)
                                      DropListBox 235, 120, 100, 45, "Select one"+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, shel_client)
                                      EditBox 45, 140, 35, 15, ALL_MEMBERS_ARRAY(shel_retro_tax_amt, shel_client)
                                      DropListBox 85, 140, 100, 45, "Select one"+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_retro_tax_verif, shel_client)
                                      EditBox 195, 140, 35, 15, ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, shel_client)
                                      DropListBox 235, 140, 100, 45, "Select one"+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, shel_client)
                                      EditBox 45, 160, 35, 15, ALL_MEMBERS_ARRAY(shel_retro_room_amt, shel_client)
                                      DropListBox 85, 160, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_retro_room_verif, shel_client)
                                      EditBox 195, 160, 35, 15, ALL_MEMBERS_ARRAY(shel_prosp_room_amt, shel_client)
                                      DropListBox 235, 160, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_prosp_room_verif, shel_client)
                                      EditBox 45, 180, 35, 15, ALL_MEMBERS_ARRAY(shel_retro_garage_amt, shel_client)
                                      DropListBox 85, 180, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_retro_garage_verif, shel_client)
                                      EditBox 195, 180, 35, 15, ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, shel_client)
                                      DropListBox 235, 180, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, shel_client)
                                      EditBox 45, 200, 35, 15, ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, shel_client)
                                      DropListBox 85, 200, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"OT - Other Doc"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, shel_client)
                                      EditBox 195, 200, 35, 15, ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, shel_client)
                                      DropListBox 235, 200, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"OT - Other Doc"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, shel_client)
                                      CheckBox 45, 220, 150, 10, "Check here if verification is requested.", ALL_MEMBERS_ARRAY(shel_verif_checkbox, shel_client)
                                      CheckBox 45, 235, 185, 10, "Check here if this verification is NOT MANDATORY.", not_mand_checkbox
                                      ButtonGroup ButtonPressed
                                        PushButton 245, 230, 90, 15, "Return to Main Dialog", return_button
                                        OkButton 600, 500, 50, 15
                                      Text 15, 35, 60, 10, "HUD Subsidized:"
                                      Text 140, 35, 30, 10, "Shared:"
                                      Text 45, 50, 50, 10, "Retrospective"
                                      Text 195, 50, 50, 10, "Prospective"
                                      Text 20, 65, 20, 10, "Rent:"
                                      Text 10, 85, 30, 10, "Lot Rent:"
                                      Text 5, 105, 35, 10, "Mortgage:"
                                      Text 5, 125, 35, 10, "Insurance:"
                                      Text 15, 145, 25, 10, "Taxes:"
                                      Text 15, 165, 25, 10, "Room:"
                                      Text 10, 185, 30, 10, "Garage:"
                                      Text 10, 205, 30, 10, "Subsidy:"
                                  End If
                                EndDialog

                                dialog Dialog1

                                If IsNumeric(manual_total_shelter) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Total Shelter costs must be a number."
                                If ALL_MEMBERS_ARRAY(shel_retro_rent_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_rent_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Rent Expense."
                                If ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective Rent Expense."
                                If ALL_MEMBERS_ARRAY(shel_retro_lot_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_lot_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Lot Rent Expense."
                                If ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective Lot Rent Expense."
                                If ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Morgage Expense."
                                If ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective ortgage Expense."
                                If ALL_MEMBERS_ARRAY(shel_retro_ins_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_ins_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Insurance Expense."
                                If ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective Insurance Expense."
                                If ALL_MEMBERS_ARRAY(shel_retro_tax_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_tax_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Tax Expense."
                                If ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective Tax Expense."
                                If ALL_MEMBERS_ARRAY(shel_retro_room_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_room_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Room Expense."
                                If ALL_MEMBERS_ARRAY(shel_prosp_room_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_room_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective Room Expense."
                                If ALL_MEMBERS_ARRAY(shel_retro_garage_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_garage_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Garage Expense."
                                If ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective Garage Expense."
                                If ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Subsidy Amount."
                                If ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective Subsidy Amount."

                                If ButtonPressed = load_button Then shel_err_msg = "LOOP" & shel_err_msg

                                If left(shel_err_msg, 4) <> "LOOP" AND shel_err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & shel_err_msg

                                If ALL_MEMBERS_ARRAY(shel_retro_rent_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_retro_rent_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_rent_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_retro_lot_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_retro_lot_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_lot_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_retro_ins_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_retro_ins_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_ins_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_retro_tax_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_retro_tax_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_tax_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_retro_room_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_retro_room_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_room_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_prosp_room_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_prosp_room_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_room_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_retro_garage_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_retro_garage_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_garage_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, shel_client) * 1

                                call update_shel_notes
                                If ALL_MEMBERS_ARRAY(shel_verif_checkbox, shel_client) = checked Then
                                    If ALL_MEMBERS_ARRAY(shel_verif_added, shel_client) <> TRUE Then
                                        verifs_needed = verifs_needed & "Shelter costs for Memb " & ALL_MEMBERS_ARRAY(full_clt, shel_client) & ". "
                                        If not_mand_checkbox = checked Then verifs_needed = verifs_needed & " THIS VERIFICATION IS NOT MANDATORY."
                                        verifs_needed = verifs_needed & "; "
                                    End If
                                    ALL_MEMBERS_ARRAY(shel_verif_added, shel_client) = TRUE
                                End If

                                If ButtonPressed = -1 Then ButtonPressed = return_button
                                If ButtonPressed = 0 Then ButtonPressed = return_button

                                If ButtonPressed = return_button Then ButtonPressed = dlg_six_button
                                If manual_total_shelter <> start_total_shel Then
                                    manual_amount_used = TRUE
                                    total_shelter_amount = manual_total_shelter
                                End If
                                If manual_amount_used = TRUE Then total_shelter_amount = manual_total_shelter
                                total_shelter_amount = total_shelter_amount * 1
                            Loop until shel_err_msg = ""
                        End If
                        If ButtonPressed = verif_button then ButtonPressed = dlg_six_button

                        Call assess_button_pressed
                        If ButtonPressed = go_to_next_page Then pass_six = true
                    End If
                Loop Until pass_six = true
                If show_seven = true Then
                    app_month_assets = app_month_assets & ""

                    BeginDialog Dialog1, 0, 0, 561, 340, "CAF Dialog 7 - Asset and Miscellaneous Info"
                      EditBox 435, 20, 115, 15, app_month_assets
                      EditBox 45, 40, 395, 15, notes_on_acct
                      EditBox 475, 40, 75, 15, notes_on_cash
                      CheckBox 45, 60, 350, 10, "Check here to confirm NO account panels and all income was reviewed for direct deposit payments.", confirm_no_account_panel_checkbox
                      EditBox 45, 80, 235, 15, notes_on_cars
                      EditBox 315, 80, 235, 15, notes_on_rest
                      EditBox 115, 100, 435, 15, notes_on_other_assets
                      EditBox 40, 130, 275, 15, MEDI
                      EditBox 360, 130, 195, 15, DIET
                      EditBox 40, 150, 515, 15, FMED
                      EditBox 40, 170, 515, 15, DISQ
                      EditBox 40, 205, 510, 15, notes_on_time
                      EditBox 60, 225, 490, 15, notes_on_sanction
                      EditBox 50, 245, 500, 15, EMPS
                      ButtonGroup ButtonPressed
                        PushButton 25, 265, 15, 15, "!", tips_and_tricks_emps_button
                      CheckBox 50, 270, 180, 10, "Sent MFIP financial orientation DVD to participant(s).", MFIP_DVD_checkbox
                      EditBox 60, 290, 495, 15, verifs_needed
                      GroupBox 105, 310, 355, 25, "Dialog Tabs"
                      Text 110, 320, 300, 10, "                       |                    |                   |                  |                    |                    |   7 - Assets    |"
                      ButtonGroup ButtonPressed
                        PushButton 15, 45, 25, 10, "ACCT", acct_button
                        PushButton 445, 45, 25, 10, "CASH", cash_button
                        PushButton 10, 135, 25, 10, "MEDI:", MEDI_button
                        PushButton 325, 135, 25, 10, "DIET:", DIET_button
                        PushButton 10, 155, 25, 10, "FMED:", FMED_button
                        PushButton 15, 85, 25, 10, "CARS", cars_button
                        If InStr(shelter_details, "Mortgage") <> 0 Then PushButton 285, 85, 25, 10, "* REST", rest_button
                        If InStr(shelter_details, "Mortgage") = 0 Then PushButton 285, 85, 25, 10, "REST", rest_button
                        PushButton 15, 105, 25, 10, "SECU", secu_button
                        PushButton 40, 105, 25, 10, "TRAN", tran_button
                        PushButton 65, 105, 45, 10, "other assets", other_asset_button
                        PushButton 10, 175, 25, 10, "DISQ:", disq_button
                        If family_cash = TRUE Then PushButton 15, 250, 30, 10, "* EMPS:", emps_button
                        If family_cash = FALSE Then PushButton 20, 250, 25, 10, "EMPS:", emps_button
                        PushButton 5, 295, 50, 10, "Verifs needed:", verif_button
                        If prev_err_msg <> "" Then PushButton 450, 265, 100, 15, "Show Dialog Review Message", dlg_revw_button
                        PushButton 110, 320, 45, 10, "1 - Personal", dlg_one_button
                        PushButton 160, 320, 35, 10, "2 - JOBS", dlg_two_button
                        PushButton 200, 320, 35, 10, "3 - BUSI", dlg_three_button
                        PushButton 240, 320, 35, 10, "4 - CSES", dlg_four_button
                        PushButton 280, 320, 35, 10, "5 - UNEA", dlg_five_button
                        PushButton 320, 320, 35, 10, "6 - Other", dlg_six_button
                        PushButton 405, 320, 50, 10, "8 - Interview", dlg_eight_button
                        PushButton 465, 315, 35, 15, "NEXT", go_to_next_page
                        CancelButton 505, 315, 50, 15
                        OkButton 600, 500, 50, 15
                      GroupBox 10, 10, 545, 115, "Assets"
                      If the_process_for_snap = "Application" Then
                        Text 310, 25, 110, 10, "* Total Liquid Assets in App Month:"
                      Else
                        Text 310, 25, 110, 10, "Total Liquid Assets in App Month:"
                      End If
                      GroupBox 10, 190, 545, 95, "MFIP/DWP"
                      If family_cash = TRUE Then
                          Text 15, 210, 25, 10, "* Time:"
                          Text 15, 230, 35, 10, "* Sanction:"
                      Else
                          Text 20, 210, 20, 10, "Time:"
                          Text 20, 230, 30, 10, "Sanction:"
                      End If
                    EndDialog

                    Dialog Dialog1
                    cancel_confirmation
                    MAXIS_dialog_navigation
                    verification_dialog

                    If ButtonPressed = tips_and_tricks_emps_button Then tips_msg = MsgBox("*** TIME, Sanction, and EMPS ***" & vbNewLine & "Why are these now required?" & vbNewLine & vbNewLine &_
                                                                                          "Information about TIME (TANF months used), SANC (Details about MFIP Sanctions), and EMPS (MFIP Employment Services) are now required for any case that is Family Cash when running the CAF. These elements are paramount to the MFIP program and should be addressed at least once per year. Review of these peices of a case can go here." & vbNewLine & vbNewLine &_
                                                                                          "What if it is a new case?" & vbNewLine & "* This is a great place to indicate that there is no history of time or sanctions used, that the client reports no benefits in another state, or that you are waiting on detail from another state. This is also a good place to identify EMPS requirement was explained or DWP overview scheduled/completed." & vbNewLine & vbNewLine &_
                                                                                          "This is a relative caregiver case, why is it needed here?" & vbNewLine & "* Since these function differently for these cases, you may not be detailing time used. Detailing that it is specifically NOT being used is extremely helpful to the new HSR or reviewer that works on this case. Add detail about how these typically mandatory elements do NOT apply in this case." & vbNewLine & vbNewLine &_
                                                                                          "The script will try to autofill this information but additional detail is helpful as always.", vbInformation, "Tips and Tricks")

                    If ButtonPressed = dlg_revw_button THen Call display_errors(prev_err_msg, FALSE)
                    If ButtonPressed = tips_and_tricks_emps_button Then ButtonPressed = dlg_seven_button
                    If ButtonPressed = -1 Then ButtonPressed = go_to_next_page
                    If ButtonPressed = verif_button then ButtonPressed = dlg_seven_button

                    Call assess_button_pressed
                    If ButtonPressed = go_to_next_page Then pass_seven = true

                    If IsNumeric(app_month_assets) = TRUE Then app_month_assets = app_month_assets * 1
                End If
            Loop Until pass_seven = true
            If show_eight = true Then
                If app_month_expenses = "" Then
                    app_month_expenses = 0
                    total_shelter_amount = replace(total_shelter_amount, "$", "")
                    total_shelter_amount = total_shelter_amount * 1
                    app_month_expenses = total_shelter_amount
                    slide = 3
                    Do
                        hest_amount = right(hest_information, slide)
                        If IsNumeric(hest_amount) = True Then Exit Do
                        slide = slide - 1
                    Loop until slide = 0
                    if IsNumeric(hest_amount) = TRUE Then app_month_expenses = app_month_expenses + hest_amount

                    app_month_expenses = app_month_expenses & ""
                End If
                app_month_income = app_month_income & ""
                app_month_assets = app_month_assets & ""
                app_month_expenses = app_month_expenses & ""

                BeginDialog Dialog1, 0, 0, 500, 370, "CAF Dialog 8 - Interview Info"
                  EditBox 60, 10, 20, 15, next_er_month
                  EditBox 85, 10, 20, 15, next_er_year
                  ComboBox 330, 10, 165, 15, "Select or Type"+chr(9)+"incomplete"+chr(9)+"approved"+chr(9)+CAF_status, CAF_status
                  EditBox 60, 30, 435, 15, actions_taken
                  DropListBox 135, 60, 30, 45, "?"+chr(9)+"Yes"+chr(9)+"No", snap_exp_yn
                  ButtonGroup ButtonPressed
                    PushButton 165, 60, 15, 15, "!", tips_and_tricks_xfs_button
                  EditBox 270, 60, 40, 15, app_month_income '210'
                  EditBox 350, 60, 40, 15, app_month_assets '290'
                  EditBox 445, 60, 40, 15, app_month_expenses '385'
                  EditBox 90, 80, 35, 15, exp_snap_approval_date
                  EditBox 195, 80, 295, 15, exp_snap_delays
                  EditBox 90, 100, 35, 15, snap_denial_date
                  EditBox 195, 100, 295, 15, snap_denial_explain
                  CheckBox 20, 155, 80, 10, "Application signed?", application_signed_checkbox
                  CheckBox 20, 170, 50, 10, "eDRS sent?", eDRS_sent_checkbox
                  CheckBox 20, 185, 65, 10, "Updated MMIS?", updated_MMIS_checkbox
                  CheckBox 20, 200, 95, 10, "Workforce referral made?", WF1_checkbox
                  CheckBox 125, 155, 85, 10, "Sent forms to AREP?", Sent_arep_checkbox
                  CheckBox 125, 170, 80, 10, "Intake packet given?", intake_packet_checkbox
                  CheckBox 125, 185, 70, 10, "IAAs/OMB given?", IAA_checkbox
                  CheckBox 220, 155, 115, 10, "Informed client of recert period?", recert_period_checkbox
                  CheckBox 220, 170, 130, 10, "Rights and Responsibilities explained?", R_R_checkbox
                  CheckBox 220, 185, 150, 10, "Client Requests to participate with E and T", E_and_T_checkbox
                  CheckBox 220, 200, 125, 10, "Eligibility Requirements Explained?", elig_req_explained_checkbox
                  CheckBox 220, 215, 160, 10, "Benefits and Payment Information Explained?", benefit_payment_explained_checkbox
                  EditBox 55, 240, 440, 15, other_notes
                  EditBox 60, 260, 435, 15, verifs_needed
                  CheckBox 15, 295, 240, 10, "Check here to update PND2 to show client delay (pending cases only).", client_delay_checkbox
                  CheckBox 15, 310, 200, 10, "Check here to create a TIKL to deny at the 30 day mark.", TIKL_checkbox
                  CheckBox 15, 325, 265, 10, "Check here to send a TIKL (10 days from now) to update PND2 for Client Delay.", client_delay_TIKL_checkbox
                  EditBox 295, 295, 50, 15, verif_req_form_sent_date
                  EditBox 295, 325, 150, 15, worker_signature
                  GroupBox 5, 345, 345, 25, "Dialog Tabs"
                  Text 10, 355, 335, 10, "                       |                    |                   |                   |                    |                    |                     | 8 - Interview"

                  ButtonGroup ButtonPressed
                    PushButton 5, 265, 50, 10, "Verifs needed:", verif_button
                    PushButton 10, 355, 45, 10, "1 - Personal", dlg_one_button
                    PushButton 60, 355, 35, 10, "2 - JOBS", dlg_two_button
                    PushButton 100, 355, 35, 10, "3 - BUSI", dlg_three_button
                    PushButton 140, 355, 35, 10, "4 - CSES", dlg_four_button
                    PushButton 180, 355, 35, 10, "5 - UNEA", dlg_five_button
                    PushButton 220, 355, 35, 10, "6 - Other", dlg_six_button
                    PushButton 260, 355, 40, 10, "7 - Assets", dlg_seven_button
                    PushButton 405, 350, 35, 15, "Done", finish_dlgs_button
                    CancelButton 445, 350, 50, 15
                    OkButton 650, 500, 50, 15
                  Text 5, 15, 55, 10, "Next ER REVW:"
                  Text 280, 15, 50, 10, "* CAF status:"
                  Text 5, 35, 55, 10, "* Actions taken:"
                  ' GroupBox 5, 50, 490, 70, "SNAP Expedited"
                  If the_process_for_snap = "Application" AND exp_det_case_note_found = FALSE Then
                    GroupBox 5, 50, 490, 70, "*** SNAP Expedited"
                  Else
                    GroupBox 5, 50, 490, 70, "SNAP Expedited"
                  End If
                  '     Text 15, 65, 120, 10, "* Is this SNAP Application Expedited?"
                  '     Text 15, 85, 75, 10, "* EXP Approval Date:"
                  '     Text 195, 65, 75, 10, "* App Month - Income:" '135'
                  '     Text 320, 65, 30, 10, "* Assets:" '260'
                  '     Text 405, 65, 40, 10, "* Expenses:" '345'
                  ' Else
                  Text 15, 65, 120, 10, "Is this SNAP Application Expedited?"
                  Text 20, 85, 65, 10, "EXP Approval Date:"
                  Text 195, 65, 70, 10, "App Month - Income:" '135'
                  Text 320, 65, 25, 10, "Assets:" '260'
                  Text 405, 65, 40, 10, "Expenses:" '345'
                  ' End If
                  Text 135, 50, 90, 10, "CAF Date: " & CAF_datestamp
                  If exp_det_case_note_found = TRUE Then Text 260, 50, 180, 10, "EXPEDITED DETERMINATION CASE/NOTE FOUND"
                  Text 135, 85, 55, 10, "Explain Delays:"
                  Text 15, 105, 75, 10, "SNAP Denial Date:"
                  Text 135, 105, 55, 10, "Explain denial:"
                  GroupBox 5, 130, 490, 105, "Common elements workers should case note:"
                  GroupBox 15, 140, 100, 90, "Application Processing"
                  GroupBox 120, 140, 90, 90, "Form Actions"
                  GroupBox 215, 140, 175, 90, "Interview"
                  Text 5, 245, 50, 10, "Other notes:"
                  GroupBox 5, 280, 280, 60, "Actions the script can do:"
                  Text 295, 285, 120, 10, "Date Verification Request Form Sent:"
                  Text 295, 315, 60, 10, "Worker signature:"
                EndDialog

                Dialog Dialog1
                cancel_confirmation
                MAXIS_dialog_navigation
                verification_dialog

                If ButtonPressed = tips_and_tricks_xfs_button Then tips_msg = MsgBox("*** Expedited SNAP ***" & vbNewLine & "Anytime the CAF script is run for SNAP at application, expedited information is required. Since the interview is complete, you have enough information to make an EXPEDITED DETERMINATION (different from the screening already completed)." & vbNewLine & vbNewLine &_
                                                                                     "The only time this information is NOT required is if you have run the separate script 'Expedited Determination' - this script will find the case note from that script run and allow you to skip this part of the dialog. The Expedited Determination script has more autofill specific to this process and more detail explained in the note." & vbNewLine & vbNewLine &_
                                                                                     "* App Month Income - Enter the total amount of income received in the month of application here. This income does NOT need to be verified. The script does not caclulate this for you." & vbNewLine & vbNewLine &_
                                                                                     "* App Month Assets - Enter the total LIQUID assets the client has available to them in the month of application. This field is also available on the Asset dialog and will carry over. The script does not calculate this for you." & vbNewLine & vbNewLine &_
                                                                                     "* App Month Expenses - Enter the shelter expense paid/responsible in the month of application plus the standard utility for which the client can claim. If these are completed correctly in the previous dialog, the script will calculate this for you when it first displays. You can change it." & vbNewLine & vbNewLine &_
                                                                                     "THIS INFORMATION SHOULD NOT BE FROM CAF1, but from the client's report and conversation complted during the interview along with any documentation we do have on file - though most are not required." & vbNewLine & vbNewLine &_
                                                                                     "Based on these amounts, enter if the client is expedited or not using the dorpdown with 'Yes' or 'No'. No other consideration should be made to determine the client's eligibility for Expedited. Answering Yes here does not mean you have approved it BUT it does mean the client is eligible for expedited processing." & vbNewLine & vbNewLine &_
                                                                                     "In most situations, the case should be approved if determined to be expedited. If the approval is done or will be done shortly, enter the date of approval." & vbNewLine & vbNewLine &_
                                                                                     "If the approval took more than the expedited processing time (7 days) then explain the delay - this may very well be that no interview had been completed." & vbNewLine & vbNewLine &_
                                                                                     "If the approval cannot be made - leave the date of approval blank and detail what is preventing the approval. Very few things prevent the approval of Expedited SNAP. If you are unsure, check the HSR Manual or contact Knowledge Now.", vbInformation, "Tips and Tricks")

                If ButtonPressed = tips_and_tricks_xfs_button Then ButtonPressed = dlg_eight_button
                If ButtonPressed = -1 Then ButtonPressed = finish_dlgs_button
                If ButtonPressed = verif_button then ButtonPressed = dlg_eight_button

                Call assess_button_pressed

                If ButtonPressed = finish_dlgs_button Then
                    'DIALOG 1
                    'New error message formatting for ease of reading.
                    If IsDate(CAF_datestamp) = FALSE Then full_err_msg = full_err_msg & "~!~" & "1^* CAF DATESTAMP ##~##   - Enter a valid date for the CAF datestamp.##~##"
                    If interview_required = TRUE Then
                        If interview_type = "Select or Type" OR trim(interview_type) = "" Then full_err_msg = full_err_msg & "~!~1^* INTERVIEW TYPE ##~##   - This case requires and interview to process the CAF - enter the interview type.##~##"
                        If IsDate(interview_date) = False Then full_err_msg = full_err_msg & "~!~1^* INTERVIEW DATE ##~##   - This case requires and interview to process the CAF - enter the interview date.##~##"
                        If interview_with = "Select or Type" OR trim(interview_with) = "" Then full_err_msg = full_err_msg & "~!~1^* INTERVIEW COMPLETED WITH ##~##   - This case requires and interview to process the CAF - indicate who the interview was completed with.##~##"
                    End If
                    For the_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                      If ALL_MEMBERS_ARRAY(gather_detail, the_member) = TRUE Then
                          ' MsgBox "Name: " & ALL_MEMBERS_ARRAY(clt_name, the_member) & vbNewLine & "Age: " & ALL_MEMBERS_ARRAY(clt_age, the_member)
                          If ALL_MEMBERS_ARRAY(clt_age, the_member) > 17 OR ALL_MEMBERS_ARRAY(memb_numb, the_member) = "01" Then
                              ALL_MEMBERS_ARRAY(id_detail, the_member) = trim(ALL_MEMBERS_ARRAY(id_detail, the_member))
                              If ALL_MEMBERS_ARRAY(clt_id_verif, the_member) = "OT - Other Document" AND ALL_MEMBERS_ARRAY(id_detail, the_member) = "" Then full_err_msg = full_err_msg & "~!~1^* DETAIL (ID Verif for " & ALL_MEMBERS_ARRAY(clt_name, the_member) & ") ##~##   - Any ID type of OT (Other) needs explanation of what is used for ID verification."
                          End If
                      End If
                    Next
                    If the_process_for_cash = "Application" AND trim(ABPS) <> "" Then
                        If trim(CS_forms_sent_date) <> "N/A" AND IsDate(CS_forms_sent_date) = False AND cash_checkbox = checked Then full_err_msg = full_err_msg & "~!~" & "1^* DATE CS FORMS SENT ##~##   - Enter a valid date for the day that child support forms were sent or given to the client. This is required for Cash cases at application with absent parents.##~##"
                    End If

                    'DIALOG 2
                    For each_job = 0 to UBound(ALL_JOBS_PANELS_ARRAY, 2)
                        If ALL_JOBS_PANELS_ARRAY(employer_name, each_job) <> "" THen
                            IF ALL_JOBS_PANELS_ARRAY(EI_case_note, each_job) = FALSE Then
                                If ALL_JOBS_PANELS_ARRAY(estimate_only, each_job) = unchecked Then
                                    ALL_JOBS_PANELS_ARRAY(budget_explain, each_job) = trim(ALL_JOBS_PANELS_ARRAY(budget_explain, each_job))
                                    If ALL_JOBS_PANELS_ARRAY(budget_explain, each_job) = "" Then
                                        full_err_msg = full_err_msg & "~!~" & "2^* EXPLAIN BUDGET for " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job) & " ##~##   - Additional detail about how the job - " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job) & " - was budgeted is required. Complete the 'Explain Budget' field for this job."
                                    ElseIf len(ALL_JOBS_PANELS_ARRAY(budget_explain, each_job)) < 20 Then
                                        full_err_msg = full_err_msg & "~!~" & "2^* EXPLAIN BUDGET for " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job) & " ##~##   - Budget detail for job - " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job) & " - should be longer. Budget cannot be sufficiently explained in a short note."
                                    End If
                                End If
                                If SNAP_checkbox = checked Then
                                    If IsNumeric(ALL_JOBS_PANELS_ARRAY(pic_pay_date_income, each_job)) = FALSE Then full_err_msg = full_err_msg & "~!~" & "2^* SNAP PIC - PAY DATE AMOUNT for " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job) & " ##~##   - For a SNAP case the average pay date amount must be entered as a number. Update the 'Pay Date Amount' for job - " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job) & "."
                                    If ALL_JOBS_PANELS_ARRAY(pic_pay_freq, each_job) = "Type or select" Then full_err_msg = full_err_msg & "~!~" & "2^* SNAP PIC - PAY FREQUENCY for " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job) & " ##~##   - The pay frequency for SNAP pay date amount needs to be identified to correctly note the income. Update the frequency after 'Pay Date Amount' for the job - " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job) & "."
                                    If IsNumeric(ALL_JOBS_PANELS_ARRAY(pic_prosp_income, each_job)) = False Then full_err_msg = full_err_msg & "~!~" & "2^* SNAP PIC - PROSPECTIVE AMOUNT for " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job) & " ##~##   - For SNAP cases, the monthly prospective amount needs to be entered as a number in the 'Prospective Amount' field for jobw - " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job) & "."
                                End If
                            End If
                        End If
                    Next

                    'DIALOG 3
                    If ALL_BUSI_PANELS_ARRAY(memb_numb, 0) <> "" Then
                        For each_busi = 0 to UBound(ALL_BUSI_PANELS_ARRAY, 2)
                            ALL_BUSI_PANELS_ARRAY(budget_explain, each_busi) = trim(ALL_BUSI_PANELS_ARRAY(budget_explain, each_busi))
                            If ALL_BUSI_PANELS_ARRAY(estimate_only, each_busi) = unchecked Then
                                If ALL_BUSI_PANELS_ARRAY(budget_explain, each_busi) = "" Then
                                    full_err_msg = full_err_msg & "~!~3^* EXPLAIN BUDGET for BUSI " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & " ##~##   - Additional detail about how BUSI " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & " was budgeted is required. Complete the 'Explain Budget' field for this self employment."
                                ElseIf len(ALL_BUSI_PANELS_ARRAY(budget_explain, each_busi)) < 20 Then
                                    full_err_msg = full_err_msg & "~!~3^* EXPLAIN BUDGET for BUSI " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & " ##~##   - Additional detail about how BUSI " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & " was budgeted should be longer - the note is too short so sufficiently explain how the income was budgeted."
                                End If
                            End If
                            If ALL_BUSI_PANELS_ARRAY(calc_method, each_busi) = "Select One" Then full_err_msg = full_err_msg & "~!~3^* SELF EMPLOYMENT METHOD for BUSI " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & " ##~##   - Indicate which calculation method will be used for BUSI " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & "."
                            If ALL_BUSI_PANELS_ARRAY(calc_method, each_busi) = "Tax Forms" Then
                                If SNAP_checkbox = checked Then
                                    If ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi) = "Income Tax Returns" AND trim(ALL_BUSI_PANELS_ARRAY(exp_not_allwd, each_busi)) = "" Then full_err_msg = full_err_msg & "~!~3^* EXPENSES NOT ALLOWED for BUSI " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & " ##~##   - Since the calculation method is 'Tax Forms' and this is a SNAP case with Tax Forms verifying, indicate what (if any) expenses on taxes have been excluded."
                                    If ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi) = "Pend Out State Verif" OR ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi) = "No Verif Provided" OR ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi) = "Delayed Verif" Then
                                    Else
                                        If ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi) <> "Income Tax Returns" Then full_err_msg = full_err_msg & "~!~3^* SNAP INCOME VERIFICATION for BUSI " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & " ##~##   - Verification of income for SNAP should be 'Income Tax Returns' when calculation method is 'Tax Forms'"
                                        If ALL_BUSI_PANELS_ARRAY(snap_expense_verif, each_busi) <> "Income Tax Returns" Then full_err_msg = full_err_msg & "~!~3^* SNAP EXPENSE VERIFICATION for BUSI " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & " ##~##   - Verification of expenses for SNAP should be 'Income Tax Returns' when calculation method is 'Tax Forms'"
                                    End If
                                End If
                                If cash_checkbox = checked or EMER_checkbox = checked Then
                                    If ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi) = "Pend Out State Verif" OR ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi) = "No Verif Provided" OR ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi) = "Delayed Verif" Then
                                    Else
                                        If ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi) <> "Income Tax Returns" Then full_err_msg = full_err_msg & "~!~3^* CASH INCOMME VERIFICATION for BUSI " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & " ##~##   - Verification of income for Cash/EMER should be 'Income Tax Returns' when calculation method is 'Tax Forms'"
                                        If ALL_BUSI_PANELS_ARRAY(cash_expense_verif, each_busi) <> "Income Tax Returns" Then full_err_msg = full_err_msg & "~!~3^* CASH EXPENSE VERIFICATION for BUSI " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & " ##~##   - Verification of expenses for Cash/EMER should be 'Income Tax Returns' when calculation method is 'Tax Forms'"
                                    End If
                                End If
                            End If
                        Next
                    End If

                    'DIALOG 4

                    'DIALOG 5
                    For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
                        If UNEA_INCOME_ARRAY(SSA_exists, each_unea_memb) = TRUE Then
                            If trim(UNEA_INCOME_ARRAY(UNEA_RSDI_amt, each_unea_memb)) <> "" AND trim(UNEA_INCOME_ARRAY(UNEA_RSDI_notes, each_unea_memb)) = "" Then full_err_msg = full_err_msg & "~!~5^* RSDI NOTES for MEMB " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " ##~##   - Explain details about RSDI Income and Budgeting."
                            If trim(UNEA_INCOME_ARRAY(UNEA_SSI_amt, each_unea_memb)) <> "" AND trim(UNEA_INCOME_ARRAY(UNEA_SSI_notes, each_unea_memb)) = "" Then full_err_msg = full_err_msg & "~!~5^* SSI NOTES for MEMB " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " ##~##   - Explain details about SSI Income and Budgeting."
                        End If
                    Next
                    For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
                        If UNEA_INCOME_ARRAY(UC_exists, each_unea_memb) = TRUE Then
                            If SNAP_checkbox = checked and IsNumeric(UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, each_unea_memb)) = False Then full_err_msg = full_err_msg & "~!~5^* UC SNAP PROSP AMOUNT for MEMB " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " ##~##   - Indicate the prospective amount of UC income that will be budgeted for SNAP."
                            If UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb) <> "" Then
                                If IsDate(UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb)) = False Then
                                    full_err_msg = full_err_msg & "~!~5^* UC TIKL DATE for MEMB " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " ##~##   - In order to set a TIKL, a valid date needs to be entered in the box for the UC TIKL."
                                Else
                                    If DateDiff("d", date, UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb)) < 0 Then full_err_msg = full_err_msg & "~!~5^* UC TIKL DATE for MEMB " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " ##~##   - To set a TIKL for the end of UC income, the TIKL date must be in the future."
                                End If
                            End If
                            If trim(UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb)) <> "" Then
                                If IsNumeric(UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb)) = False Then full_err_msg = full_err_msg & "~!~5^* UC WEEKLY GROSS for MEMB " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " ##~##   - The UC Gross weekly amount needs to be entered as a number."
                                If UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb) <> UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb) Then
                                    If IsNumeric(UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb)) = FALSE or IsNumeric(UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb)) = FALSE or IsNumeric(UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb)) = FALSE Then
                                        If IsNumeric(UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb)) = FALSE Then full_err_msg = full_err_msg & "~!~5^* UC BUDGETED WEEKLY AMOUNT for MEMB " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " ##~##   - Enter the UC weekly Net Amount as a number."
                                        If IsNumeric(UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb)) = FALSE Then full_err_msg = full_err_msg & "~!~5^* UC WEEKLY ALLOWED DEDUCTIONS for MEMB " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " ##~##   - Enter the weekly allowed deductions for UC as a number."
                                    Else
                                        calculated_net_weekly = UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb) - UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb)
                                        UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb) = UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb) * 1
                                        'MsgBox "Calc Net Weekly - " & calculated_net_weekly & vbCR & "Entered Net Weekly - " & UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb)
                                        If calculated_net_weekly <> UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb) Then full_err_msg = full_err_msg & "~!~5^* UC WEEKLY GROSS, BUDGETED AMOUNT, ALLOWED DEDUCTIONS for MEMB " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " ##~##   - Review your UC weekly gross, net and counted deductions. The net amount is not equal to the gross amount less counted deductions. ##~## Weekly Gross ($" & UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb) & ") - Allowed Deductions ($" & UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb) & ") = $" & calculated_net_weekly & " ##~## Weekly Budgeted Amount $" & UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb) &  " ##~## Difference between gross and allowed deductions should equal the weekly budgeted ammount."
                                    End If
                                End If
                            End If
                        End If
                    Next

                    'DIALOG 6
                    If SNAP_checkbox = checked and trim(notes_on_wreg) = "" Then full_err_msg = full_err_msg & "~!~6^* WREG Notes ##~##   - Update WREG detail as this is a SNAP case."
                    If living_situation = "Blank" or living_situation = "  " Then full_err_msg = full_err_msg & "~!~6^* LIVING SITUATION ##~##   - Living situation needs to be entered for each case. 'Blank' is not valid."
                    'We are not erroring for if ADDR verification is 'NO' or '?' - if we get additional policy information that this is necessary - add it here

                    'DIALOG 7
                    ' If SNAP_checkbox = checked and CAF_type = "Application" Then
                    '     If trim(app_month_assets) = "" OR IsNumeric(app_month_assets) = FALSE AND exp_det_case_note_found = FALSE Then full_err_msg = full_err_msg & "~!~7^* Indicate the total of liquid assets in the application month."
                    ' End If

                    If family_cash = TRUE and trim(notes_on_time) = "" Then full_err_msg = full_err_msg & "~!~7^* TIME ##~##   - For a family cash case, detail on TIME needs to be added."
                    If family_cash = TRUE and trim(notes_on_sanction) = "" Then full_err_msg = full_err_msg & "~!~7^* SANCTION ##~##   - This is a family cash case, sanction detail needs to be added."
                    If family_cash = TRUE and trim(EMPS) = "" Then full_err_msg = full_err_msg & "~!~7^* EMMPS ##~##   - EMPS detail needs to be added for a family cash case. "
                    If cash_checkbox = unchecked AND trim(DIET) <> "" Then full_err_msg = full_err_msg & "~!~7^* DIET ##~##   - DIET information should not be entered into a non-cash case."
                    If InStr(shelter_details, "Mortgage") AND trim(notes_on_rest) = "" Then full_err_msg = full_err_msg & "~!~7^* REST ##~##   - SHEL indicates that Mortgage is being paid, but no information has been added to REST. Update Shelter information or add detail to REST."

                    'DIALOG 8
                    If CAF_status = "Select or Type" Then full_err_msg = full_err_msg & "~!~8^* CAF STATUS ##~##   - Indicate the CAF Status."
                    If the_process_for_snap = "Application" AND exp_det_case_note_found = FALSE Then
                        If trim(snap_denial_date) <> "" AND IsDate(snap_denial_date) = FALSE Then
                            full_err_msg = full_err_msg & "~!~8^* SNAP DENIAL DATE ##~##   - This is a a SNAP case at application. You entered something in the SNAP denial date but it does not appear to be a date. Please list the date that SNAP will be denied if SNAP is being denied."
                        ElseIf IsDate(snap_denial_date) = TRUE Then
                            If DateDiff("d", date, snap_denial_date) > 0 Then full_err_msg = full_err_msg & "~!~8^* SNAP DENIAL DATE ##~##   - The denial date is listed as a future date. Review the date entered in the SNAP denial date field."
                        ElseIf trim(snap_denial_date) = "" Then
                            If snap_exp_yn = "?" Then
                                full_err_msg = full_err_msg & "~!~8^* IS THIS SNAP APPLICATION EXPEDITED ##~##   - This is a a SNAP case at application. Indicate if this case has been determined to be expedited SNAP or not."
                            Else
                                If IsNumeric(app_month_income) = FALSE Then full_err_msg = full_err_msg & "~!~8^* APP MONTH - INCOME ##~##   - Enter the income for the application month as a number."
                                If IsNumeric(app_month_assets) = FALSE Then full_err_msg = full_err_msg & "~!~8^* APP MONTH - ASSETS ##~##   - Enter the liquid assets for the application month as a number."
                                If IsNumeric(app_month_expenses) = FALSE Then full_err_msg = full_err_msg & "~!~8^* APP MONTH - EXPENSES ##~##   - Enter the expenses (shelter and utilities) for the application month as a number."

                                case_should_be_xfs = FALSE
                                If IsNumeric(app_month_income) = TRUE AND IsNumeric(app_month_assets) = TRUE AND IsNumeric(app_month_expenses) = TRUE Then
                                    If app_month_assets <=100 AND app_month_income < 150 Then
                                        case_should_be_xfs = TRUE
                                        ' MsgBox "low resources"
                                    End If
                                    app_month_assets = app_month_assets * 1
                                    app_month_income = app_month_income * 1
                                    app_month_expenses = app_month_expenses * 1
                                    app_month_resources = app_month_assets + app_month_income
                                    If app_month_resources < app_month_expenses Then
                                        case_should_be_xfs = TRUE
                                        ' MsgBox "insufficient resources" & vbCR & "Resources - " & app_month_resources & vbCR & "Expenses - " & app_month_expenses
                                    End If

                                    If snap_exp_yn = "Yes" and case_should_be_xfs = FALSE Then full_err_msg = full_err_msg & "~!~8^* SNAP EXPEDITED ##~##   - This is indicated as Expedited, though based on app month details it appears to be NOT Expedited. ##~## App Month: Income - $" & app_month_income & ". Assets - $" & app_month_assets & ". Expenses - $" & app_month_expenses & "."
                                    If snap_exp_yn = "No" AND case_should_be_xfs = TRUE Then full_err_msg = full_err_msg & "~!~8^* SNAP EXPEDITED ##~##   - This is indicated as NOT Expedited, though based on app month details it appears to be EXPEDITED. ##~## App Month: Income - $" & app_month_income & ". Assets - $" & app_month_assets & ". Expenses - $" & app_month_expenses & "."
                                End If
                                If snap_exp_yn = "Yes" Then
                                    If IsDate(exp_snap_approval_date) = TRUE Then
                                        If DateDiff("d", date, exp_snap_approval_date) > 0 Then
                                            full_err_msg = full_err_msg & "~!~8^* EXP APPROVAL DATE ##~##   - The date listed in the expedited approval date is a future date. Please review the date listed and reenter if necessary."
                                        ElseIf DateDiff("d", CAF_datestamp, exp_snap_approval_date) > 7 AND trim(exp_snap_delays) = "" Then
                                            full_err_msg = full_err_msg & "~!~8^* EXPLAIN DELAYS ##~##   - Since Expedited SNAP is not approved within 7 days of the date of application, pease explain the reason for the delay."
                                        End If
                                    Else
                                        If trim(exp_snap_delays) = "" Then full_err_msg = full_err_msg & "~!~8^* EXPLAIN DELAYS ##~##   - Since the Expedited SNAP does not have an approval date yet, either explain the reason for the delay or indicate the date of Expedited SNAP Approval."
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If trim(actions_taken) = "" Then full_err_msg = full_err_msg & "~!~8^* ACTIONS TAKEN ##~##   - Indicate what actions were taken when processing this CAF."
                    prev_err_msg = full_err_msg
                End If

                Call display_errors(full_err_msg, TRUE)
                If full_err_msg = "" and ButtonPressed = finish_dlgs_button Then pass_eight = true
                If ButtonPressed = finish_dlgs_button Then ButtonPressed = -1
            End If
            ' MsgBox "Button - " & ButtonPressed & vbNewLine & "Pass Eight - " & pass_eight
        Loop until pass_eight = true
        ' MsgBox "Now we call proceed confirmation"
        CALL proceed_confirmation(case_note_confirm)			'Checks to make sure that we're ready to case note.
    Loop until case_note_confirm = TRUE							'Loops until we affirm that we're ready to case note.
    ' MsgBox "Now We call check for password"
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = False

Call back_to_SELF
If continue_in_inquiry = "" Then
    Do
        Call back_to_SELF
        EMReadScreen MX_region, 12, 22, 48
        MX_region = trim(MX_region)
        If MX_region = "INQUIRY DB" Then

            BeginDialog dialog1, 0, 0, 266, 120, "Still in Inquiry"
              ButtonGroup ButtonPressed
                PushButton 165, 80, 95, 15, "Stop the Script Run (ESC)", stop_script_button
                PushButton 140, 100, 120, 15, "Continue - I have switched (Enter)", continue_script
              Text 10, 10, 110, 20, "It appears you are now running in INQUIRY on this session."
              Text 10, 40, 105, 20, "The script cannot update or CASE/NOTE in INQUIRY."
              Text 10, 65, 255, 10, "Switch to Production now to ensure the note is entered and continue the script."
            EndDialog

            Do
                dialog dialog1
                If ButtonPressed = stop_script_button Then ButtonPressed = 0
                If ButtonPressed = 0 Then script_end_procedure("Script ended since it was started in Inquiry.")
                If ButtonPressed = -1 Then ButtonPressed = continue_script

                Call check_for_password(are_we_passworded_out)
            Loop until are_we_passworded_out = FALSE

        Else
            ButtonPressed = continue_script
        End If
    Loop until ButtonPressed = continue_script AND MX_region <> "INQUIRY DB"
End If

If trim(CS_forms_sent_date) = "N/A" Then CS_forms_sent_date = ""

'Go to ADDR to update living situation
Call navigate_to_MAXIS_screen("STAT", "ADDR")
EMReadScreen panel_living_sit, 2, 11, 43
If living_situation = "Blank" or living_situation = "  " Then
    dialog_liv_sit_code = "__"
Else
    dialog_liv_sit_code = left(living_situation, 2)
End If

If dialog_liv_sit_code <> panel_living_sit OR dialog_liv_sit_code = "__" Then
    PF9
    EMWriteScreen dialog_liv_sit_code, 11, 43
    transmit
End If


'This code will update the interview date in PROG.
If CAF_type = "Application" Then        'Interview date is not on PROG for recertifications or addendums
    If SNAP_checkbox = checked OR cash_checkbox = checked Then          'Interviews are only required for Cash and SNAP
        intv_date_needed = FALSE
        Call navigate_to_MAXIS_screen("STAT", "PROG")                   'Going to STAT to check to see if there is already an interview indicated.

        If SNAP_checkbox = checked Then                                 'If the script is being run for a SNAP interview
            EMReadScreen entered_intv_date, 8, 10, 55                   'REading what is entered in the SNAP interview
            'MsgBox "SNAP interview date - " & entered_intv_date
            If entered_intv_date = "__ __ __" Then intv_date_needed = TRUE  'If this is blank - the script needs to prompt worker to update it
        End If

        If cash_checkbox = checked THen                             'If the script is bring run for a Cash interview
            EMReadScreen cash_one_app, 8, 6, 33                     'First the script needs to identify if it is cash 1 or cash 2 that has the application information
            EMReadScreen cash_two_app, 8, 7, 33
            EMReadScreen grh_cash_app, 8, 9, 33

            cash_one_app = replace(cash_one_app, " ", "/")          'Turning this in to a date format
            cash_two_app = replace(cash_two_app, " ", "/")
            grh_cash_app = replace(grh_cash_app, " ", "/")

            If cash_one_app <> "__/__/__" Then      'Error handling - VB doesn't like date comparisons with non-dates
                If IsDate(cash_one_app) = TRUE Then
                    if DateDiff("d", cash_one_app, CAF_datestamp) = 0 then prog_row = 6     'If date of application on PROG matches script date of applicaton
                End If
            End If
            If cash_two_app <> "__/__/__" Then
                If IsDate(cash_two_app) = TRUE Then
                    if DateDiff("d", cash_two_app, CAF_datestamp) = 0 then prog_row = 7
                End If
            End If

            If grh_cash_app <> "__/__/__" Then
                If IsDate(grh_cash_app) = TRUE THen
                    if DateDiff("d", grh_cash_app, CAF_datestamp) = 0 then prog_row = 9
                End If
            End If

            EMReadScreen entered_intv_date, 8, prog_row, 55                     'Reading the right interview date with row defined above
            'MsgBox "Cash interview date - " & entered_intv_date
            If entered_intv_date = "__ __ __" Then intv_date_needed = TRUE      'If this is blank - script needs to prompt worker to have it updated
        End If

        If intv_date_needed = TRUE Then         'If previous code has determined that PROG needs to be updated
            If the_process_for_snap = "Application" Then prog_update_SNAP_checkbox = checked     'Auto checking based on the programs the script is being run for.
            If the_process_for_cash = "Application" Then prog_update_cash_checkbox = checked

            'Dialog code
            BeginDialog Dialog1, 0, 0, 231, 130, "Update PROG?"
              OptionGroup RadioGroup1
                RadioButton 10, 10, 155, 10, "YES! Update PROG with the Interview Date", confirm_update_prog
                RadioButton 10, 60, 90, 10, "No, do not update PROG", do_not_update_prog
              EditBox 165, 5, 50, 15, interview_date
              CheckBox 25, 25, 30, 10, "SNAP", prog_update_SNAP_checkbox
              CheckBox 25, 40, 30, 10, "CASH", prog_update_cash_checkbox
              Text 20, 75, 200, 10, "Reason PROG should not be updated with the Interview Date:"
              EditBox 20, 90, 195, 15, no_update_reason
              ButtonGroup ButtonPressed
                OkButton 175, 110, 50, 15
            EndDialog

            'Running the dialog
            Do
                Do
                    err_msg = ""
                    Dialog Dialog1
                    'Requiring a reason for not updating PROG and making sure if confirm is updated that a program is selected.
                    If do_not_update_prog = 1 AND no_update_reason = "" Then err_msg = err_msg & vbNewLine & "* If PROG is not to be updated, please explain why PROG should not be updated."
                    IF confirm_update_prog = 1 AND prog_update_SNAP_checkbox = unchecked AND prog_update_cash_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "* Select either CASH or SNAP to have updated on PROG."

                    If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
                Loop until err_msg = ""
                Call check_for_password(are_we_passworded_out)
            Loop until are_we_passworded_out = FALSE

            If confirm_update_prog = 1 Then     'If the dialog selects to have PROG updated
                CALL back_to_SELF               'Need to do this because we need to go to the footer month of the application and we may be in a different month

                keep_footer_month = MAXIS_footer_month      'Saving the footer month and year that was determined earlier in the script. It needs t obe changed for nav functions to work correctly
                keep_footer_year = MAXIS_footer_year

                app_month = DatePart("m", CAF_datestamp)    'Setting the footer month and year to the app month.
                app_year = DatePart("yyyy", CAF_datestamp)

                MAXIS_footer_month = right("00" & app_month, 2)
                MAXIS_footer_year = right(app_year, 2)

                Call back_to_SELF
                CALL navigate_to_MAXIS_screen ("STAT", "PROG")  'Now we can navigate to PROG in the application footer month and year
                PF9                                             'Edit

                intv_mo = DatePart("m", interview_date)     'Setting the date parts to individual variables for ease of writing
                intv_day = DatePart("d", interview_date)
                intv_yr = DatePart("yyyy", interview_date)

                intv_mo = right("00"&intv_mo, 2)            'formatting variables in to 2 digit strings - because MAXIS
                intv_day = right("00"&intv_day, 2)
                intv_yr = right(intv_yr, 2)
                intv_date_to_check = intv_mo & " " & intv_day & " " & intv_yr

                If prog_update_SNAP_checkbox = checked Then     'If it was selected to SNAP interview to be updated
                    programs_w_interview = "SNAP"               'Setting a variable for case noting

                    EMWriteScreen intv_mo, 10, 55               'SNAP is easy because there is only one area for interview - the variables go there
                    EMWriteScreen intv_day, 10, 58
                    EMWriteScreen intv_yr, 10, 61
                End If

                If prog_update_cash_checkbox = checked Then     'If it was selected to update for Cash
                    If programs_w_interview = "" Then programs_w_interview = "CASH"     'variable for the case note
                    If programs_w_interview <> "" Then programs_w_interview = "SNAP and CASH"
                    EMReadScreen cash_one_app, 8, 6, 33     'Reading app dates of both cash lines
                    EMReadScreen cash_two_app, 8, 7, 33
                    EMReadScreen grh_cash_app, 8, 9, 33

                    cash_one_app = replace(cash_one_app, " ", "/")      'Formatting as dates
                    cash_two_app = replace(cash_two_app, " ", "/")
                    grh_cash_app = replace(grh_cash_app, " ", "/")

                    If cash_one_app <> "__/__/__" Then              'Comparing them to the date of application to determine which row to use
                        If IsDate(cash_one_app) = TRUE Then
                            if DateDiff("d", cash_one_app, CAF_datestamp) = 0 then prog_row = 6
                        End If
                    End If
                    If cash_two_app <> "__/__/__" Then
                        If IsDate(cash_two_app) = TRUE Then
                            if DateDiff("d", cash_two_app, CAF_datestamp) = 0 then prog_row = 7
                        End If
                    End If

                    If grh_cash_app <> "__/__/__" Then
                        If IsDate(grh_cash_app) = TRUE Then
                            if DateDiff("d", grh_cash_app, CAF_datestamp) = 0 then prog_row = 9
                        End If
                    End If

                    EMWriteScreen intv_mo, prog_row, 55     'Writing the interview date in
                    EMWriteScreen intv_day, prog_row, 58
                    EMWriteScreen intv_yr, prog_row, 61
                End If

                transmit                                    'Saving the panel

                Call HCRE_panel_bypass
                Call back_to_SELF
                Call MAXIS_background_check

                MAXIS_footer_month = keep_footer_month      'resetting the footer month and year so the rest of the script uses the worker identified footer month and year.
                MAXIS_footer_year = keep_footer_year
            End If
        ENd If

        If intv_date_needed = TRUE and confirm_update_prog = 1 Then         'If previous code has determined that PROG needs to be updated
            snap_intv_date_updated = FALSE
            cash_intv_date_updated = FALSE
            show_prog_update_failure = FALSE
            Call back_to_SELF
            CALL navigate_to_MAXIS_screen ("STAT", "PROG")  'Now we can navigate to PROG in the application footer month and year
            If prog_update_SNAP_checkbox = checked Then
                EMReadScreen new_snap_intv_date, 8, 10, 55
                If new_snap_intv_date = intv_date_to_check Then snap_intv_date_updated = TRUE
                If snap_intv_date_updated = FALSE Then show_prog_update_failure = TRUE
            End If
            If prog_update_cash_checkbox = checked Then
                EMReadScreen new_cash_intv_date, 8, prog_row, 55
                If new_cash_intv_date = intv_date_to_check Then cash_intv_date_updated = TRUE
                If cash_intv_date_updated = FALSE Then show_prog_update_failure = TRUE
            End If

            If show_prog_update_failure = TRUE Then
                fail_msg = "You have requested the script update PROG for "
                If prog_update_SNAP_checkbox = checked AND prog_update_cash_checkbox = checked Then
                    fail_msg = fail_msg & "Cash and SNAP "
                ElseIf prog_update_SNAP_checkbox = checked Then
                    fail_msg = fail_msg & "SNAP "
                ElseIf prog_update_cash_checkbox = checked Then
                    fail_msg = fail_msg & "Cash "
                End If

                fail_msg = fail_msg & "to enter the interview date on PROG." & vbCr & vbCr & "The script was unable to update PROG completely." & vbCr

                If prog_update_SNAP_checkbox = checked Then
                    fail_msg = fail_msg & " - The SNAP Interview Date was not entered." & vbCr
                ElseIf prog_update_cash_checkbox = checked Then
                    fail_msg = fail_msg & " - The Cash Interview Date was not entered." & vbCr
                End If
                fail_msg = fail_msg & vbCr & "The PROG panel will need to be updated manually with the interview information."

                MsgBox fail_msg
            End If
        End If
    End If
End If

If the_process_for_cash = "Recertification" OR the_process_for_snap = "Recertification" Then
    If interview_required = TRUE or interview_waived = TRUE Then
        revw_panel_update_needed = FALSE
        Call Navigate_to_MAXIS_screen("STAT", "REVW")
        EMReadScreen STAT_REVW_caf_date, 8, 13, 37
        EMReadScreen STAT_REVW_intvw_date, 8, 15, 37
        If STAT_REVW_caf_date = "__ __ __" Then revw_panel_update_needed = TRUE
        If STAT_REVW_intvw_date = "__ __ __" Then revw_panel_update_needed = TRUE

        If revw_panel_update_needed = TRUE Then
            If interview_waived = TRUE AND trim(interview_date) = "" Then interview_date = date
            EMReadScreen cash_stat_revw_status, 1, 7, 40
            EMReadScreen snap_stat_revw_status, 1, 7, 60

            BeginDialog Dialog1, 0, 0, 241, 165, "Update REVW"
              OptionGroup RadioGroup1
                RadioButton 10, 10, 185, 10, "YES! Update REVW with the Interview Date/CAF Date", confirm_update_revw
                RadioButton 10, 95, 100, 10, "No, do not update REVW", do_not_update_revw
              EditBox 70, 25, 45, 15, interview_date
              EditBox 70, 45, 45, 15, caf_datestamp
              EditBox 20, 125, 215, 15, no_update_reason
              ButtonGroup ButtonPressed
                OkButton 185, 145, 50, 15
              Text 20, 30, 50, 10, "Interview Date:"
              If interview_is_being_waived = vbYes Then Text 125, 25, 105, 35, "THIS INTERVIEW WAS WAIVED. Today's date will be used."
              Text 35, 50, 35, 10, "CAF Date:"
              Text 20, 70, 175, 20, "If the REVW Status has not been updated already, it will be changed to an 'I' when the dates are entered."
              Text 20, 110, 220, 10, "Reason REVW should not be updated with the Interview/CAF Date:"
            EndDialog

            'Running the dialog
            Do
                Do
                    err_msg = ""
                    Dialog Dialog1
                    'Requiring a reason for not updating PROG and making sure if confirm is updated that a program is selected.
                    If do_not_update_revw = 1 AND no_update_reason = "" Then err_msg = err_msg & vbNewLine & "* If REVW is not to be updated, please explain why REVW should not be updated."
                    IF confirm_update_revw = 1 Then
                        If IsDate(interview_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Check the Interview Date as it appears invalid."
                        If IsDate(caf_datestamp) = FALSE Then err_msg = err_msg & vbNewLine & "* Check the CAF Date as it appears invalid."
                    End If

                    If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
                Loop until err_msg = ""
                Call check_for_password(are_we_passworded_out)
            Loop until are_we_passworded_out = FALSE

            IF confirm_update_revw = 1 Then
                Call Navigate_to_MAXIS_screen("STAT", "REVW")
                PF9
                Call create_mainframe_friendly_date(CAF_datestamp, 13, 37, "YY")
                Call create_mainframe_friendly_date(interview_date, 15, 37, "YY")

                If cash_stat_revw_status = "N" Then EMWriteScreen "I", 7, 40
                If snap_stat_revw_status = "N" Then EMWriteScreen "I", 7, 60

                attempt_count = 1
                Do
                    transmit
                    EMReadScreen actually_saved, 7, 24, 2
                    attempt_count = attempt_count + 1
                    If attempt_count = 20 Then
                        PF10
                        revw_panel_updated = FALSE
                        Exit Do
                    End If
                Loop until actually_saved = "ENTER A"

                revw_intv_date_updated = FALSE
                Call back_to_SELF
                Call Navigate_to_MAXIS_screen("STAT", "REVW")

                EMReadScreen updated_intv_date, 8, 15, 37
                If IsDate(updated_intv_date) = TRUE Then
                    updated_intv_date = DateAdd("d", 0, updated_intv_date)
                    If updated_intv_date = interview_date Then revw_intv_date_updated = TRUE

                    fail_msg = "You have requested the script update REVW with the interview date." & vbCr & vbCr & "The script was unable to update REVW completely." & vbCr & vbCr & "The REVW panel will need to be updated manually with the interview information."
                    If revw_intv_date_updated = FALSE Then MsgBox fail_msg
                End If
            End If

            If interview_is_being_waived = vbYes AND interview_date = date Then interview_date = ""
        End If
    End If
End If

Do
    Do
        qual_err_msg = ""

        BeginDialog Dialog1, 0, 0, 451, 205, "CAF Qualifying Questions"
          DropListBox 220, 40, 25, 45, "No"+chr(9)+"Yes", qual_question_one
          ComboBox 340, 40, 105, 45, verification_memb_list, qual_memb_one
          DropListBox 220, 80, 25, 45, "No"+chr(9)+"Yes", qual_question_two
          ComboBox 340, 80, 105, 45, verification_memb_list, qual_memb_two
          DropListBox 220, 110, 25, 45, "No"+chr(9)+"Yes", qual_question_three
          ComboBox 340, 110, 105, 45, verification_memb_list, qual_memb_there
          DropListBox 220, 140, 25, 45, "No"+chr(9)+"Yes", qual_question_four
          ComboBox 340, 140, 105, 45, verification_memb_list, qual_memb_four
          DropListBox 220, 160, 25, 45, "No"+chr(9)+"Yes", qual_question_five
          ComboBox 340, 160, 105, 45, verification_memb_list, qual_memb_five
          ButtonGroup ButtonPressed
            OkButton 340, 185, 50, 15
            CancelButton 395, 185, 50, 15
          Text 10, 10, 395, 15, "Qualifying Questions are listed at the end of the CAF form and are completed by the client. Indicate the answers to those questions here. If any are 'Yes' then indicate which household member to which the question refers."
          Text 10, 40, 200, 40, "Has a court or any other civil or administrative process in Minnesota or any other state found anyone in the household guilty or has anyone been disqualified from receiving public assistance for breaking any of the rules listed in the CAF?"
          Text 10, 80, 195, 30, "Has anyone in the household been convicted of making fraudulent statements about their place of residence to get cash or SNAP benefits from more than one state?"
          Text 10, 110, 195, 30, "Is anyone in your householdhiding or running from the law to avoid prosecution being taken into custody, or to avoid going to jail for a felony?"
          Text 10, 140, 195, 20, "Has anyone in your household been convicted of a drug felony in the past 10 years?"
          Text 10, 160, 195, 20, "Is anyone in your household currently violating a condition of parole, probation or supervised release?"
          Text 260, 40, 70, 10, "Household Member:"
          Text 260, 80, 70, 10, "Household Member:"
          Text 260, 110, 70, 10, "Household Member:"
          Text 260, 140, 70, 10, "Household Member:"
          Text 260, 160, 70, 10, "Household Member:"
        EndDialog

        dialog Dialog1
        cancel_confirmation

        If qual_question_one = "Yes" AND (trim(qual_memb_one) = "" OR qual_memb_one = "Select or Type") Then qual_err_msg = qual_err_msg & vbNewLine & "* For Quesion 1, yes is indicated however no member is listed - please enter the member that this question applies to."
        If qual_question_two = "Yes" AND (trim(qual_memb_two) = "" OR qual_memb_two = "Select or Type") Then qual_err_msg = qual_err_msg & vbNewLine & "* For Quesion 2, yes is indicated however no member is listed - please enter the member that this question applies to."
        If qual_question_three = "Yes" AND (trim(qual_memb_three) = "" OR qual_memb_three = "Select or Type") Then qual_err_msg = qual_err_msg & vbNewLine & "* For Quesion 3, yes is indicated however no member is listed - please enter the member that this question applies to."
        If qual_question_four = "Yes" AND (trim(qual_memb_four) = "" OR qual_memb_four = "Select or Type") Then qual_err_msg = qual_err_msg & vbNewLine & "* For Quesion 4, yes is indicated however no member is listed - please enter the member that this question applies to."
        If qual_question_five = "Yes" AND (trim(qual_memb_five) = "" OR qual_memb_five = "Select or Type") Then qual_err_msg = qual_err_msg & vbNewLine & "* For Quesion 5, yes is indicated however no member is listed - please enter the member that this question applies to."

        If qual_err_msg <> "" Then MsgBox "Please Resolve to Continue:" & vbNewLine & qual_err_msg
    Loop until qual_err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

qual_questions_yes = FALSE
If qual_question_one = "Yes" Then qual_questions_yes = TRUE
If qual_question_two = "Yes" Then qual_questions_yes = TRUE
If qual_question_three = "Yes" Then qual_questions_yes = TRUE
If qual_question_four = "Yes" Then qual_questions_yes = TRUE
If qual_question_five = "Yes" Then qual_questions_yes = TRUE

'Now, the client_delay_checkbox business. It'll update client delay if the box is checked and it isn't a recert.
If client_delay_checkbox = checked and CAF_type <> "Recertification" then
	call navigate_to_MAXIS_screen("rept", "pnd2")

    limit_reached = FALSE
    row = 1
    col = 1
    EMSearch "The REPT:PND2 Display Limit Has Been Reached.", row, col
    If row <> 0 Then
        transmit
        limit_reached = TRUE
    End If

    If limit_reached = TRUE Then
        PND2_row = 7
        Do
            EMReadScreen PND2_case_number, 8, PND2_row, 5
            if trim(PND2_case_number) = MAXIS_case_number Then Exit Do
            PND2_row = PND2_row + 1
        Loop until PND2_row = 18
    Else
        EMGetCursor PND2_row, PND2_col
    End If

    If PND2_row = 18 Then
        client_delay_checkbox = unchecked
        MsgBox "The scriipt could not navigate to REPT/PND2 due to a MAXIS display limit. This case will not be updated for client delay. Please email to BlueZone Script Team with the case number and report that the Display Limit on REPT/PND2 was reached."
    End If

	for i = 0 to 1 'This is put in a for...next statement so that it will check for "additional app" situations, where the case could be on multiple lines in REPT/PND2. It exits after one if it can't find an additional app.
		EMReadScreen PND2_SNAP_status_check, 1, PND2_row, 62
		If PND2_SNAP_status_check = "P" then EMWriteScreen "C", PND2_row, 62
		EMReadScreen PND2_HC_status_check, 1, PND2_row, 65
		If PND2_HC_status_check = "P" then
			EMWriteScreen "x", PND2_row, 3
			transmit
			person_delay_row = 7
			Do
				EMReadScreen person_delay_check, 1, person_delay_row, 39
				If person_delay_check <> " " then EMWriteScreen "c", person_delay_row, 39
				person_delay_row = person_delay_row + 2
			Loop until person_delay_check = " " or person_delay_row > 20
			PF3
		End if
		EMReadScreen additional_app_check, 14, PND2_row + 1, 17
		If additional_app_check <> "ADDITIONAL APP" then exit for
		PND2_row = PND2_row + 1
	next
	PF3
	EMReadScreen PND2_check, 4, 2, 52
	If PND2_check = "PND2" then
		MsgBox "PND2 might not have been updated for client delay. There may have been a MAXIS error. Check this manually after case noting."
		PF10
		client_delay_checkbox = unchecked		'Probably unnecessary except that it changes the case note parameters
	End if
ElseIf client_delay_checkbox = checked and CAF_type = "Recertification" then
    client_delay_checkbox = unchecked		'Probably unnecessary except that it changes the case note parameters
End if

'Going to TIKL. Now using the write TIKL function
If TIKL_checkbox = checked and CAF_type <> "Recertification" then
	If DateDiff ("d", CAF_datestamp, date) > 30 Then 'Error handling to prevent script from attempting to write a TIKL in the past
		MsgBox "Cannot set TIKL as CAF Date is over 30 days old and TIKL would be in the past. You must manually track."
        TIKL_checkbox = unchecked
	Else
        If cash_checkbox = checked then TIKL_msg_one = TIKL_msg_one & "Cash/"
        If SNAP_checkbox = checked then TIKL_msg_one = TIKL_msg_one & "SNAP/"
        If EMER_checkbox = checked then TIKL_msg_one = TIKL_msg_one & "EMER/"
        TIKL_msg_one = Left(TIKL_msg_one, (len(TIKL_msg_one) - 1))
        TIKL_msg_one = TIKL_msg_one & " has been pending for 30 days. Evaluate for possible denial."
		'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
        Call create_TIKL(TIKL_msg_one, 30, CAF_datestamp, False, TIKL_note_text)
	End If
ElseIf TIKL_checkbox = checked and CAF_type = "Recertification" then
    TIKL_checkbox = unchecked
End if
If client_delay_TIKL_checkbox = checked then
    'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
    Call create_TIKL(">>>UPDATE PND2 FOR CLIENT DELAY IF APPROPRIATE<<<", 10, date, False, TIKL_note_text)
End if

For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
    If UNEA_INCOME_ARRAY(UC_exists, each_unea_memb) = TRUE Then
        If IsDate(UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb)) = TRUE Then
            'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
            tikl_msg = "Review UC Income for Member " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " as it may have ended or be near ending."
            Call create_TIKL(TIKL_msg, 10, UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb), False, TIKL_note_text)
        End If
    End If
Next
'--------------------END OF TIKL BUSINESS

If HC_checkbox = checked Then
    call autofill_editbox_from_MAXIS(HH_member_array, "HCRE-retro", retro_request)
    If the_process_for_hc = "Application" Then call autofill_editbox_from_MAXIS(HH_member_array, "HCRE", HC_datestamp)
    If the_process_for_hc = "Recertification" Then call autofill_editbox_from_MAXIS(HH_member_array, "REVW", HC_datestamp)
    call autofill_editbox_from_MAXIS(HH_member_array, "ACCI", hc_acci_info)
    call autofill_editbox_from_MAXIS(HH_member_array, "BILS", hc_bils_info)
    call autofill_editbox_from_MAXIS(HH_member_array, "FACI", hc_faci_info)
    call autofill_editbox_from_MAXIS(HH_member_array, "INSA", hc_insa_info)
    If CAF_form = "Combined AR for Certain Pops (DHS-3727)" Then
        HC_document_received = "DHS-3727 (Combined AR for Certain Pops)"
        HC_datestamp = CAF_datestamp & ""
    End If
    hc_medi_info = MEDI
    hc_faci_info = FACI

    BeginDialog Dialog1, 0, 0, 481, 295, "HC Detail"
      ComboBox 80, 5, 150, 15, "Select or Type"+chr(9)+"DHS-2128 (LTC Renewal)"+chr(9)+"DHS-3417B (Req. to Apply...)"+chr(9)+"DHS-3418 (HC Renewal)"+chr(9)+"DHS-3531 (LTC Application)"+chr(9)+"DHS-3876 (Certain Pops App)"+chr(9)+"DHS-6696 (MNsure HC App)"+chr(9)+"DHS-3727 (Combined AR for Certain Pops)"+chr(9)+HC_document_received, HC_document_received
      EditBox 80, 20, 50, 15, HC_datestamp
      ComboBox 360, 5, 115, 15, "Select of Type"+chr(9)+"incomplete"+chr(9)+"approved"+chr(9)+"denied"+chr(9)+HC_form_status, HC_form_status
      CheckBox 310, 25, 80, 10, "Application signed?", HC_application_signed_check
      CheckBox 405, 25, 65, 10, "MMIS updated?", MMIS_updated_check
      EditBox 65, 40, 165, 15, retro_request
      EditBox 290, 40, 185, 15, hc_hh_comp
      EditBox 35, 60, 440, 15, hc_medi_info
      EditBox 35, 80, 440, 15, hc_insa_info
      EditBox 35, 100, 440, 15, hc_acci_info
      EditBox 35, 120, 440, 15, hc_bils_info
      EditBox 35, 140, 440, 15, hc_faci_info
      EditBox 55, 160, 420, 15, waiver_ltc_info
      EditBox 55, 180, 420, 15, spenddown_info
      CheckBox 55, 200, 245, 10, "Check here to have the script create a TIKL to deny at the 45 day mark.", hc_tikl_checkbox
      EditBox 55, 215, 420, 15, hc_other_notes
      EditBox 55, 235, 420, 15, hc_verifs_needed
      EditBox 55, 255, 420, 15, hc_actions_taken
      ButtonGroup ButtonPressed
        OkButton 370, 275, 50, 15
        CancelButton 425, 275, 50, 15
        PushButton 5, 65, 25, 10, "MEDI:", MEDI_button
        PushButton 5, 85, 25, 10, "INSA", INSA_button
        PushButton 5, 105, 25, 10, "ACCI:", ACCI_button
        PushButton 5, 125, 25, 10, "BILS:", BILS_button
        PushButton 5, 145, 25, 10, "FACI:", FACI_button
      Text 10, 10, 70, 10, "HC Form Received:"
      Text 30, 25, 45, 10, "Date Stamp:"
      Text 300, 10, 60, 10, "HC Form status:"
      Text 10, 45, 50, 10, "Retro Request:"
      Text 240, 45, 50, 10, "HC HH Comp:"
      Text 10, 165, 45, 10, "Waiver/LTC:"
      Text 10, 185, 40, 10, "Spenddown:"
      Text 10, 220, 40, 10, "Other notes:"
      Text 5, 240, 50, 10, "Verifs needed:"
      Text 5, 260, 50, 10, "Actions taken:"
    EndDialog

    Do
        Do
            hc_err_msg = ""

            dialog dialog1

            cancel_confirmation
            MAXIS_dialog_navigation

            If IsDate(HC_datestamp) = False Then hc_err_msg = hc_err_msg & vbNewLine & "* Enter the date of the form as a valid date."
            If trim(HC_form_status) = "" OR trim(HC_form_status) = "Select or Type" Then hc_err_msg = hc_err_msg & vbNewLine & "* Indicate (by selection or typing manually) the status of the health care form being processed."
            If trim(HC_document_received) = "" or trim(HC_document_received) = "Select or Type" Then hc_err_msg = hc_err_msg & vbNewLine & "* Indicate (by selection or typing manually) what form is being processed for health care."

            If hc_err_msg <> "" Then MsgBox "Please Resolve to Continue:" & vbNewLine & hc_err_msg

        Loop until hc_err_msg = ""
        call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE

    If hc_tikl_checkbox = checked Then
        If DateDiff ("d", HC_datestamp, date) > 45 Then 'Error handling to prevent script from attempting to write a TIKL in the past
            MsgBox "Cannot set TIKL as HC Form Date is over 45 days old and TIKL would be in the past. You must manually track."
            hc_tikl_checkbox = unchecked
        Else
            'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
            Call create_TIKL("HC pending 45 days. Evaluate for possible denial. If any members are elderly/disabled, allow an additional 15 days and reTIKL out.", 45, HC_datestamp, False, TIKL_note_text)
        End If
    End If
End If

'Adding a colon to the beginning of the CAF status variable if it isn't blank (simplifies writing the header of the case note)
If CAF_status <> "" then CAF_status = ": " & CAF_status

'Adding footer month to the recertification case notes
If CAF_type = "Recertification" then CAF_type = MAXIS_footer_month & "/" & MAXIS_footer_year & " recert"
progs_list = ""
If cash_checkbox = checked Then progs_list = progs_list & ", Cash"
If SNAP_checkbox = checked Then progs_list = progs_list & ", SNAP"
If EMER_checkbox = checked Then progs_list = progs_list & ", EMER"
If left(progs_list, 1) = "," Then progs_list = right(progs_list, len(progs_list) - 2)

prog_and_type_list = ""
If cash_checkbox = checked Then
    If the_process_for_cash = "Application" Then prog_and_type_list = prog_and_type_list & ", Cash App"
    If the_process_for_cash = "Recertification" Then prog_and_type_list = prog_and_type_list & ", " & cash_recert_mo & "/" & cash_recert_yr & " Cash Recert"
End If
If snap_checkbox = checked Then
    If the_process_for_snap = "Application" Then prog_and_type_list = prog_and_type_list & ", SNAP App"
    If the_process_for_snap = "Recertification" Then prog_and_type_list = prog_and_type_list & ", " & snap_recert_mo & "/" & snap_recert_yr & " SNAP Recert"
End If
If EMER_checkbox = checked Then prog_and_type_list = prog_and_type_list & ", EMER App"
If left(prog_and_type_list, 1) = "," Then prog_and_type_list = right(prog_and_type_list, len(prog_and_type_list) - 2)

If SNAP_checkbox = checked Then
    adult_snap_count = adult_snap_count * 1
    child_snap_count = child_snap_count * 1
    total_snap_count = adult_snap_count + child_snap_count
    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        If ALL_MEMBERS_ARRAY(include_snap_checkbox, each_member) = checked Then
            included_snap_members = included_snap_members & ", M" & ALL_MEMBERS_ARRAY(memb_numb, each_member)
        Else
            If ALL_MEMBERS_ARRAY(count_snap_checkbox, each_member) = checked Then counted_snap_members = counted_snap_members & ", M" & ALL_MEMBERS_ARRAY(memb_numb, each_member)
        End If
    Next
    If included_snap_members <> "" Then included_snap_members = right(included_snap_members, len(included_snap_members) - 2)
    If counted_snap_members <> "" Then counted_snap_members = right(counted_snap_members, len(counted_snap_members) - 2)
End If
If cash_checkbox = checked Then
    adult_cash_count = adult_cash_count * 1
    child_cash_count = child_cash_count * 1
    total_cash_count = adult_cash_count + child_cash_count
    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        If ALL_MEMBERS_ARRAY(include_cash_checkbox, each_member) = checked Then
            included_cash_members = included_cash_members & ", M" & ALL_MEMBERS_ARRAY(memb_numb, each_member)
        Else
            If ALL_MEMBERS_ARRAY(count_cash_checkbox, each_member) = checked Then counted_cash_members = counted_cash_members & ", M" & ALL_MEMBERS_ARRAY(memb_numb, each_member)
        End If
    Next
    If included_cash_members <> "" Then included_cash_members = right(included_cash_members, len(included_cash_members) - 2)
    If counted_cash_members <> "" Then counted_cash_members = right(counted_cash_members, len(counted_cash_members) - 2)
End If
If EMER_checkbox = checked Then
    adult_emer_count = adult_emer_count * 1
    child_emer_count = child_emer_count * 1
    total_emer_count = adult_emer_count + child_emer_count
    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        If ALL_MEMBERS_ARRAY(include_emer_checkbox, each_member) = checked Then
            included_emer_members = included_emer_members & ", M" & ALL_MEMBERS_ARRAY(memb_numb, each_member)
        Else
            If ALL_MEMBERS_ARRAY(count_emer_checkbox, each_member) = checked Then counted_emer_members = counted_emer_members & ", M" & ALL_MEMBERS_ARRAY(memb_numb, each_member)
        End If
    Next
    If included_emer_members <> "" Then included_emer_members = right(included_emer_members, len(included_emer_members) - 2)
    If counted_emer_members <> "" Then counted_emer_members = right(counted_emer_members, len(counted_emer_members) - 2)
End If

'Determining if there are
'Income
case_has_income = FALSE

If ALL_JOBS_PANELS_ARRAY(memb_numb, 0) <> "" Then case_has_income = TRUE
If trim(notes_on_jobs) <> "" Then case_has_income = TRUE
If trim(earned_income) <> "" Then case_has_income = TRUE
If ALL_BUSI_PANELS_ARRAY(memb_numb, 0) <> "" Then case_has_income = TRUE
If trim(notes_on_busi) <> "" Then case_has_income = TRUE
For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
    If UNEA_INCOME_ARRAY(CS_exists, each_unea_memb) = TRUE Then case_has_income = TRUE
    If UNEA_INCOME_ARRAY(SSA_exists, each_unea_memb) = TRUE Then case_has_income = TRUE
    If UNEA_INCOME_ARRAY(UC_exists, each_unea_memb) = TRUE Then case_has_income = TRUE
Next
If trim(notes_on_cses) <> "" Then case_has_income = TRUE
If trim(notes_on_ssa_income) <> "" Then case_has_income = TRUE
If trim(notes_on_VA_income) <> "" Then case_has_income = TRUE
If trim(notes_on_WC_income) <> "" Then case_has_income = TRUE
If trim(other_uc_income_notes) <> "" Then case_has_income = TRUE
If trim(notes_on_other_UNEA) <> "" Then case_has_income = TRUE

'Personal
case_has_personal = FALSE

If trim(cit_id) <> "" Then case_has_personal = TRUE
If trim(IMIG) <> "" Then case_has_personal = TRUE
If trim(SCHL) <> "" Then case_has_personal = TRUE
If trim(DISA) <> "" Then case_has_personal = TRUE
If trim(FACI) <> "" Then case_has_personal = TRUE
If trim(PREG) <> "" Then case_has_personal = TRUE
If trim(ABPS) <> "" Then case_has_personal = TRUE
If CS_forms_sent_date <> "" Then case_has_personal = TRUE
If trim(AREP) <> "" Then case_has_personal = TRUE
If address_confirmation_checkbox = checked Then case_has_personal = TRUE
If homeless_yn = "Yes" Then case_has_personal = TRUE
If trim(addr_county) <> "" Then case_has_personal = TRUE
If trim(living_situation) <> "" Then case_has_personal = TRUE
If trim(notes_on_address) <> "" Then case_has_personal = TRUE
If trim(DISQ) <> "" Then case_has_personal = TRUE
If trim(notes_on_wreg) <> "" Then case_has_personal = TRUE
all_abawd_notes = notes_on_abawd & notes_on_abawd_two & notes_on_abawd_three
If trim(all_abawd_notes) <> "" Then case_has_personal = TRUE
If trim(notes_on_time) <> "" Then case_has_personal = TRUE
If trim(notes_on_sanction) <> "" Then case_has_personal = TRUE
If trim(EMPS) <> "" Then case_has_personal = TRUE
If MFIP_DVD_checkbox = checked Then case_has_personal = TRUE
If trim(MEDI) <> "" Then case_has_personal = TRUE
If trim(DIET) <> "" Then case_has_personal = TRUE
If trim(case_changes) <> "" Then case_has_personal = TRUE

'Resources
case_has_resources = FALSE

If confirm_no_account_panel_checkbox = checked Then case_has_resources = TRUE
If trim(notes_on_acct) <> "" Then case_has_resources = TRUE
If trim(notes_on_cash) <> "" Then case_has_resources = TRUE
If trim(notes_on_cars) <> "" Then case_has_resources = TRUE
If trim(notes_on_rest) <> "" Then case_has_resources = TRUE
If trim(notes_on_other_assets) <> "" Then case_has_resources = TRUE

'Expenses
case_has_expenses = FALSE

If trim(total_shelter_amount) <> "" Then case_has_expenses = TRUE
If trim(full_shelter_details) <> "" Then case_has_expenses = TRUE
If trim(notes_on_acut) <> "" Then case_has_expenses = TRUE
If hest_information <> "Select ALLOWED HEST" Then case_has_expenses = TRUE
If trim(notes_on_coex) <> "" Then case_has_expenses = TRUE
If trim(notes_on_dcex) <> "" Then case_has_expenses = TRUE
If trim(notes_on_other_deduction) <> "" Then case_has_expenses = TRUE
If trim(expense_notes) <> "" Then case_has_expenses = TRUE
If trim(FMED) <> "" Then case_has_expenses = TRUE

'THE CASE NOTES-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Expedited Determination Case Note
'Navigates to case note, and checks to make sure we aren't in inquiry.
If HC_checkbox = checked Then

    hc_note_header = HC_datestamp & " " & HC_document_received & ": " & HC_form_status

    Call start_a_blank_CASE_NOTE
    Call write_variable_in_CASE_NOTE(hc_note_header)

    Call write_bullet_and_variable_in_CASE_NOTE("Actions Taken", hc_actions_taken)

    If HC_application_signed_check = checked Then Call write_variable_in_CASE_NOTE("* HC form was signed.")

    Call write_bullet_and_variable_in_CASE_NOTE("HC form received", HC_datestamp)
    Call write_bullet_and_variable_in_CASE_NOTE("Retro Request", retro_request)
    Call write_bullet_and_variable_in_CASE_NOTE("HC HH Comp", hc_hh_comp)
    Call write_bullet_and_variable_in_CASE_NOTE("Spenddown", spenddown_info)

    'INCOME
    If case_has_income = TRUE Then
        Call write_variable_in_CASE_NOTE("===== INCOME =====")
    Else
        Call write_variable_in_CASE_NOTE("== No Income detail Listed for this case. ==")
    End If
    'JOBS
    If ALL_JOBS_PANELS_ARRAY(memb_numb, 0) <> "" Then
        ' Call write_variable_with_indent(variable_name)
        Call write_variable_in_CASE_NOTE("--- JOBS Income ---")
        For each_job = 0 to UBound(ALL_JOBS_PANELS_ARRAY, 2)
            Call write_variable_in_CASE_NOTE("Member " & ALL_JOBS_PANELS_ARRAY(memb_numb, each_job) & " at " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job))
            If ALL_JOBS_PANELS_ARRAY(estimate_only, each_job) = checked Then Call write_variable_in_CASE_NOTE("* This job has not been verified and this is only an estimate.")
            IF ALL_JOBS_PANELS_ARRAY(EI_case_note, each_job) = TRUE Then call write_variable_in_CASE_NOTE("* BUDGET DETAIL ABOUT THIS JOB IN PREVIOUS CASE NOTE.")
            If ALL_JOBS_PANELS_ARRAY(verif_code, each_job) = "Delayed" Then
                Call write_variable_in_CASE_NOTE("* Verification of this job has been delayed for review or approval of Expedited SNAP.")
            ElseIf ALL_JOBS_PANELS_ARRAY(estimate_only, each_job) = unchecked Then
                Call write_variable_in_CASE_NOTE("* Verification - " & ALL_JOBS_PANELS_ARRAY(verif_code, each_job))
            End If
            Call write_bullet_and_variable_in_CASE_NOTE("Verification", ALL_JOBS_PANELS_ARRAY(verif_explain, each_job))
            If ALL_JOBS_PANELS_ARRAY(job_retro_income, each_job) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Retro Income: $" & ALL_JOBS_PANELS_ARRAY(job_retro_income, each_job) & " - " & ALL_JOBS_PANELS_ARRAY(retro_hours, each_job) & " hours.")
            If ALL_JOBS_PANELS_ARRAY(job_prosp_income, each_job) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Prospective Income: $" & ALL_JOBS_PANELS_ARRAY(job_prosp_income, each_job) & " - " & ALL_JOBS_PANELS_ARRAY(prosp_hours, each_job) & " hours.")
            If ALL_JOBS_PANELS_ARRAY(budget_explain, each_job) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("About Budget: " & ALL_JOBS_PANELS_ARRAY(budget_explain, each_job))
        Next
    End If
    Call write_bullet_and_variable_in_CASE_NOTE("JOBS", notes_on_jobs)
    Call write_bullet_and_variable_in_CASE_NOTE("Other Earned Income", earned_income)

    'BUSI
    If ALL_BUSI_PANELS_ARRAY(memb_numb, 0) <> "" Then
        Call write_variable_in_CASE_NOTE("--- BUSI Income ---")
        For each_busi = 0 to UBound(ALL_BUSI_PANELS_ARRAY, 2)
            busi_det_msg = "Self Employment for Memb " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " - BUSI type:" & right(ALL_BUSI_PANELS_ARRAY(busi_type, each_busi), len(ALL_BUSI_PANELS_ARRAY(busi_type, each_busi)) - 4) & "."
            Call write_variable_in_CASE_NOTE(busi_det_msg)

            If ALL_BUSI_PANELS_ARRAY(busi_desc, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Description: " & ALL_BUSI_PANELS_ARRAY(busi_desc, each_busi) & ".")
            If ALL_BUSI_PANELS_ARRAY(busi_structure, each_busi) <> "Select or Type" and ALL_BUSI_PANELS_ARRAY(busi_structure, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Business structure: " & ALL_BUSI_PANELS_ARRAY(busi_structure, each_busi) & ".")
            If ALL_BUSI_PANELS_ARRAY(share_denom, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Clt owns " & ALL_BUSI_PANELS_ARRAY(share_num, each_busi) & "/" & ALL_BUSI_PANELS_ARRAY(share_denom, each_busi) & " of the business.")
            If ALL_BUSI_PANELS_ARRAY(partners_in_HH, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Business also owned by Memb(s) " & ALL_BUSI_PANELS_ARRAY(partners_in_HH, each_busi) & ".")

            se_method_det_msg = "* Self Employment Budgeting method selected: " & ALL_BUSI_PANELS_ARRAY(calc_method, each_busi) & "."
            Call write_variable_in_CASE_NOTE(se_method_det_msg)
            If ALL_BUSI_PANELS_ARRAY(mthd_date, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Method selected on: " & ALL_BUSI_PANELS_ARRAY(mthd_date, each_busi) & ".")
            If ALL_BUSI_PANELS_ARRAY(method_convo_checkbox, each_busi) = checked Then Call write_variable_with_indent_in_CASE_NOTE("The self employment method selected was discussed with the client.")

            If cash_checkbox = checked OR EMER_checkbox = checked Then
                Call write_variable_in_CASE_NOTE("* Cash Income and Expense Detail:")
                cash_income_det = ""
                cash_expense_det = ""

                If ALL_BUSI_PANELS_ARRAY(income_ret_cash, each_busi) <> "" Then
                    cash_income_det = "RETRO Income $" & ALL_BUSI_PANELS_ARRAY(income_ret_cash, each_busi) & " - "
                    cash_expense_det = "RETRO Expenses $" & ALL_BUSI_PANELS_ARRAY(expense_ret_cash, each_busi) & " - "
                End If
                If ALL_BUSI_PANELS_ARRAY(income_pro_cash, each_busi) <> "" Then
                    cash_income_det = cash_income_det & "PROSP Income $" & ALL_BUSI_PANELS_ARRAY(income_pro_cash, each_busi) & " - "
                    cash_expense_det = cash_expense_det & "PROSP Expenses $" & ALL_BUSI_PANELS_ARRAY(expense_pro_cash, each_busi) & " - "
                End If
                If ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi) <> "Select or Type" and ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi) <> "" Then cash_income_det = cash_income_det & "Verification: " & ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi)
                If ALL_BUSI_PANELS_ARRAY(cash_expense_verif, each_busi) <> "Select or Type" and ALL_BUSI_PANELS_ARRAY(cash_expense_verif, each_busi) <> "" Then cash_expense_det = cash_expense_det & "Verification: " & ALL_BUSI_PANELS_ARRAY(cash_expense_verif, each_busi)

                Call write_variable_with_indent_in_CASE_NOTE(cash_income_det)
                Call write_variable_with_indent_in_CASE_NOTE(cash_expense_det)
            End If
            If SNAP_checkbox = checked Then
                Call write_variable_in_CASE_NOTE("* SNAP Income and Expense Detail:")
                snap_income_det = ""
                snap_expense_det = ""

                If ALL_BUSI_PANELS_ARRAY(income_ret_snap, each_busi) <> "" Then
                    snap_income_det = "RETRO Income $" & ALL_BUSI_PANELS_ARRAY(income_ret_snap, each_busi) & " - "
                    snap_expense_det = "RETRO Expenses $" & ALL_BUSI_PANELS_ARRAY(expense_ret_snap, each_busi) & " - "
                End If
                If ALL_BUSI_PANELS_ARRAY(income_pro_snap, each_busi) <> "" Then
                    snap_income_det = snap_income_det & "PROSP Income $" & ALL_BUSI_PANELS_ARRAY(income_pro_snap, each_busi) & " - "
                    snap_expense_det = snap_expense_det & "PROSP Expenses $" & ALL_BUSI_PANELS_ARRAY(expense_pro_snap, each_busi) & " - "
                End If
                If ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi) <> "Select or Type" and ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi) <> "" Then snap_income_det = snap_income_det & "Verification: " & ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi)
                If ALL_BUSI_PANELS_ARRAY(snap_expense_verif, each_busi) <> "Select or Type" and ALL_BUSI_PANELS_ARRAY(snap_expense_verif, each_busi) <> "" Then snap_expense_det = snap_expense_det & "Verification: " & ALL_BUSI_PANELS_ARRAY(snap_expense_verif, each_busi)

                Call write_variable_with_indent_in_CASE_NOTE(snap_income_det)
                Call write_variable_with_indent_in_CASE_NOTE(snap_expense_det)
                If ALL_BUSI_PANELS_ARRAY(exp_not_allwd, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Expenses from taxes not allowed: " & ALL_BUSI_PANELS_ARRAY(exp_not_allwd, each_busi))
            End If
            rept_hours_det_msg = ""
            min_wg_hours_det_msg = ""
            If ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, each_busi) = "___" Then ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, each_busi) = ""
            If ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, each_busi) = "___" Then ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, each_busi) = ""
            If ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, each_busi) = "___" Then ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, each_busi) = ""
            If ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, each_busi) = "___" Then ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, each_busi) = ""

            If ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, each_busi) <> "" OR ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, each_busi) <> "" Then
                rept_hours_det_msg = rept_hours_det_msg & "Clt reported monthly work hours of: "
                If ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, each_busi) <> "" Then rept_hours_det_msg = rept_hours_det_msg & ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, each_busi) & " retrospecive work and "
                If ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, each_busi) <> "" Then rept_hours_det_msg = rept_hours_det_msg & ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, each_busi) & " prospoective work hrs"
                rept_hours_det_msg = rept_hours_det_msg & ". "
            End If
            If ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, each_busi) <> "" OR ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, each_busi) <> "" Then
                min_wg_hours_det_msg = min_wg_hours_det_msg & "Work earnings indicate Minumun Wage Hours of: "
                If ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, each_busi) <> "" Then min_wg_hours_det_msg = min_wg_hours_det_msg & ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, each_busi) & " retrospective and "
                If ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, each_busi) <> "" Then min_wg_hours_det_msg = min_wg_hours_det_msg & ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, each_busi) & " prospective"
                min_wg_hours_det_msg = min_wg_hours_det_msg & ". "
            End If
            If rept_hours_det_msg <> "" Then Call write_variable_in_CASE_NOTE("* " & rept_hours_det_msg)
            If min_wg_hours_det_msg <> "" Then Call write_variable_in_CASE_NOTE("* " & min_wg_hours_det_msg)
            If ALL_BUSI_PANELS_ARRAY(verif_explain, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Verif Detail: " & ALL_BUSI_PANELS_ARRAY(verif_explain, each_busi))
            If ALL_BUSI_PANELS_ARRAY(budget_explain, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Budget Detail: " & ALL_BUSI_PANELS_ARRAY(budget_explain, each_busi))
        Next
    End If
    Call write_bullet_and_variable_in_CASE_NOTE("BUSI", notes_on_busi)

    'CSES
    If show_cses_detail = TRUE Then
        Call write_variable_in_CASE_NOTE("--- Child Support Income ---")
        For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
            If UNEA_INCOME_ARRAY(CS_exists, each_unea_memb) = TRUE Then
                total_cs = 0
                If IsNumeric(UNEA_INCOME_ARRAY(direct_CS_amt, each_unea_memb)) = TRUE Then total_cs = total_cs + UNEA_INCOME_ARRAY(direct_CS_amt, each_unea_memb)
                If IsNumeric(UNEA_INCOME_ARRAY(disb_CS_amt, each_unea_memb)) = TRUE Then total_cs = total_cs + UNEA_INCOME_ARRAY(disb_CS_amt, each_unea_memb)
                If IsNumeric(UNEA_INCOME_ARRAY(disb_CS_arrears_amt, each_unea_memb)) = TRUE Then total_cs = total_cs + UNEA_INCOME_ARRAY(disb_CS_arrears_amt, each_unea_memb)

                Call write_variable_in_CASE_NOTE("* Total child support income for Member " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & ": $" & total_cs)
                If UNEA_INCOME_ARRAY(disb_CS_amt, each_unea_memb) <> "" Then
                    cs_disb_inc_det = "Disbursed child support: $" & UNEA_INCOME_ARRAY(disb_CS_amt, each_unea_memb)
                    If trim(UNEA_INCOME_ARRAY(disb_CS_notes, each_unea_memb)) <> "" Then cs_disb_inc_det = cs_disb_inc_det & ". Notes: " & UNEA_INCOME_ARRAY(disb_CS_notes, each_unea_memb)

                    Call write_variable_with_indent_in_CASE_NOTE(cs_disb_inc_det)
                    If trim(UNEA_INCOME_ARRAY(disb_CS_months, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Income was determined using " & UNEA_INCOME_ARRAY(disb_CS_months, each_unea_memb) & " month(s) of disbursement income.")
                    If trim(UNEA_INCOME_ARRAY(disb_CS_prosp_budg, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("SNAP prospective budget details: " & UNEA_INCOME_ARRAY(disb_CS_prosp_budg, each_unea_memb))
                End If

                If UNEA_INCOME_ARRAY(disb_CS_arrears_amt, each_unea_memb) <> "" Then
                    cs_arrears_inc_det = "Disbursed child support arrears: $" & UNEA_INCOME_ARRAY(disb_CS_arrears_amt, each_unea_memb)
                    If trim(UNEA_INCOME_ARRAY(disb_CS_arrears_notes, each_unea_memb)) <> "" Then cs_arrears_inc_det = cs_arrears_inc_det & ". Notes: " & UNEA_INCOME_ARRAY(disb_CS_arrears_notes, each_unea_memb)

                    Call write_variable_with_indent_in_CASE_NOTE(cs_arrears_inc_det)
                    If trim(UNEA_INCOME_ARRAY(disb_CS_arrears_months, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Income was determined using " & UNEA_INCOME_ARRAY(disb_CS_arrears_months, each_unea_memb) & " month(s) of disbursement income.")
                    If trim(UNEA_INCOME_ARRAY(disb_CS_arrears_budg, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("SNAP prospective budget details: " & UNEA_INCOME_ARRAY(disb_CS_arrears_budg, each_unea_memb))
                End If

                If UNEA_INCOME_ARRAY(direct_CS_amt, each_unea_memb) <> "" Then
                    direct_cs_inc_det = "Direct child support: $" & UNEA_INCOME_ARRAY(direct_CS_amt, each_unea_memb)
                    If trim(UNEA_INCOME_ARRAY(direct_CS_notes, each_unea_memb)) <> "" Then direct_cs_inc_det = direct_cs_inc_det & ". Notes: " & UNEA_INCOME_ARRAY(direct_CS_notes, each_unea_memb)

                    Call write_variable_with_indent_in_CASE_NOTE(direct_cs_inc_det)
                End if
            End If
        Next
    End If
    Call write_bullet_and_variable_in_CASE_NOTE("Other Child Support Income", notes_on_cses)

    'UNEA
    For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
        If UNEA_INCOME_ARRAY(SSA_exists, each_unea_memb) = TRUE Then
            rsdi_income_det = ""
            If trim(UNEA_INCOME_ARRAY(UNEA_RSDI_amt, each_unea_memb)) <> "" Then rsdi_income_det = rsdi_income_det & "RSDI: $" & UNEA_INCOME_ARRAY(UNEA_RSDI_amt, each_unea_memb)
            If trim(UNEA_INCOME_ARRAY(UNEA_RSDI_notes, each_unea_memb)) <> "" Then rsdi_income_det = rsdi_income_det & ". Notes: " & UNEA_INCOME_ARRAY(UNEA_RSDI_notes, each_unea_memb)

            ssi_income_det = ""
            If trim(UNEA_INCOME_ARRAY(UNEA_SSI_amt, each_unea_memb)) <> "" Then ssi_income_det = ssi_income_det & "SSI: $" & UNEA_INCOME_ARRAY(UNEA_SSI_amt, each_unea_memb)
            If trim(UNEA_INCOME_ARRAY(UNEA_SSI_notes, each_unea_memb)) <> "" Then ssi_income_det = ssi_income_det & ". Notes: " & UNEA_INCOME_ARRAY(UNEA_SSI_notes, each_unea_memb)

            Call write_variable_in_CASE_NOTE("* Member " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " SSA income:")
            If rsdi_income_det <> "" Then Call write_variable_with_indent_in_CASE_NOTE(rsdi_income_det)
            If ssi_income_det <> "" Then Call write_variable_with_indent_in_CASE_NOTE(ssi_income_det)
        End If
    Next
    Call write_bullet_and_variable_in_CASE_NOTE("Other SSA Income", notes_on_ssa_income)
    Call write_bullet_and_variable_in_CASE_NOTE("VA Income", notes_on_VA_income)
    Call write_bullet_and_variable_in_CASE_NOTE("Workers Comp Income", notes_on_WC_income)

    For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
        If UNEA_INCOME_ARRAY(UC_exists, each_unea_memb) = TRUE Then
            uc_income_det_one = ""
            uc_income_det_two = ""
            If trim(UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb)) <> "" Then
                uc_income_det_one = uc_income_det_one & "Budgeted UC weekly amount: $" & UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb) & ".; "
                uc_income_det_one = uc_income_det_one & "UC weekly gross income: $" & UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb) & ". "
                If trim(UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb)) <> "" Then uc_income_det_one = uc_income_det_one & "Deduction amount allowed: $" & UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb) & ". "
                If trim(UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, each_unea_memb)) <> "" Then uc_income_det_one = uc_income_det_one & "Deduction amount excluded: $" & UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, each_unea_memb) & ". "
            Else
                uc_income_det_one = uc_income_det_one & "Budgeted UC weekly amount: $" & UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb) & ".; "
                If trim(UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb)) <> "" Then uc_income_det_one = uc_income_det_one & "Deduction amount allowed: $" & UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb) & ". "
                If trim(UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, each_unea_memb)) <> "" Then uc_income_det_one = uc_income_det_one & "Deduction amount excluded: $" & UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, each_unea_memb) & ". "
            End If
            If trim(UNEA_INCOME_ARRAY(UNEA_UC_account_balance, each_unea_memb)) <> "" Then uc_income_det_one = uc_income_det_one & "Current UC account balance: $" & UNEA_INCOME_ARRAY(UNEA_UC_account_balance, each_unea_memb) & ". "
            If trim(UNEA_INCOME_ARRAY(UNEA_UC_retro_amt, each_unea_memb)) <> "" Then uc_income_det_two = uc_income_det_two & "UC Retro Income: $" & UNEA_INCOME_ARRAY(UNEA_UC_retro_amt, each_unea_memb) & ". "
            If trim(UNEA_INCOME_ARRAY(UNEA_UC_prosp_amt, each_unea_memb)) <> "" Then uc_income_det_two = uc_income_det_two & "UC Prosp Income: $" & UNEA_INCOME_ARRAY(UNEA_UC_prosp_amt, each_unea_memb) & ". "
            If trim(UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, each_unea_memb)) <> "" Then uc_income_det_two = uc_income_det_two & "UC SNAP budgeted Income: $" & UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, each_unea_memb) & ". "

            Call write_variable_in_CASE_NOTE("* Member " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " Unemployment Income:")
            If trim(UNEA_INCOME_ARRAY(UNEA_UC_start_date, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("UC Income started on: " & UNEA_INCOME_ARRAY(UNEA_UC_start_date, each_unea_memb) & ". ")
            If uc_income_det_one <> "" Then Call write_variable_with_indent_in_CASE_NOTE(uc_income_det_one)
            If uc_income_det_two <> "" Then Call write_variable_with_indent_in_CASE_NOTE(uc_income_det_two)
            If IsDate(UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb)) = TRUE Then Call write_variable_with_indent_in_CASE_NOTE("TIKL set to check for end of UC on: " & UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb))
            If trim(UNEA_INCOME_ARRAY(UNEA_UC_notes, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Notes: " & UNEA_INCOME_ARRAY(UNEA_UC_notes, each_unea_memb))
        End If
    Next
    Call write_bullet_and_variable_in_CASE_NOTE("Other UC Income", other_uc_income_notes)
    Call write_bullet_and_variable_in_CASE_NOTE("Unearned Income", notes_on_other_UNEA)

    If case_has_personal = TRUE Then
        If trim(cit_id) <> "" Then case_has_hc_personal = TRUE
        If trim(IMIG) <> "" Then case_has_hc_personal = TRUE
        If trim(hc_acci_info) <> "" Then case_has_hc_personal = TRUE
        If trim(hc_faci_info) <> "" Then case_has_hc_personal = TRUE
        If trim(waiver_ltc_info) <> "" Then case_has_hc_personal = TRUE
        If trim(hc_medi_info) <> "" Then case_has_hc_personal = TRUE
        If trim(hc_insa_info) <> "" Then case_has_hc_personal = TRUE
        If trim(DISA) <> "" Then case_has_hc_personal = TRUE
        If trim(PREG) <> "" Then case_has_hc_personal = TRUE
        If trim(ABPS) <> "" Then case_has_hc_personal = TRUE
        If trim(AREP) <> "" Then case_has_hc_personal = TRUE
        If trim(DISQ) <> "" Then case_has_hc_personal = TRUE
    End If
    If case_has_hc_personal = TRUE Then Call write_variable_in_CASE_NOTE("===== PERSONAL =====")

    Call write_bullet_and_variable_in_CASE_NOTE("Citizenship/ID", cit_id)
    Call write_bullet_and_variable_in_CASE_NOTE("IMIG", IMIG)
    Call write_bullet_and_variable_in_CASE_NOTE("Changes", case_changes)
    Call write_bullet_and_variable_in_CASE_NOTE("Accident", hc_acci_info)

    Call write_bullet_and_variable_in_CASE_NOTE("Facility", hc_faci_info)
    Call write_bullet_and_variable_in_CASE_NOTE("Waiver/LTC", waiver_ltc_info)

    Call write_bullet_and_variable_in_CASE_NOTE("Medicare", hc_medi_info)
    Call write_bullet_and_variable_in_CASE_NOTE("Other Insurance", hc_insa_info)

    Call write_bullet_and_variable_in_CASE_NOTE("DISA", DISA)
    Call write_bullet_and_variable_in_CASE_NOTE("PREG", PREG)
    Call write_bullet_and_variable_in_CASE_NOTE("Absent Parent", ABPS)
    If CS_forms_sent_date <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Child Support Forms given/sent to client on " & CS_forms_sent_date)
    Call write_bullet_and_variable_in_CASE_NOTE("AREP", AREP)

    'DISQ
    Call write_bullet_and_variable_in_CASE_NOTE("DISQ", DISQ)

    If case_has_expenses = TRUE Then
        If trim(notes_on_coex) <> "" Then case_has_hc_expenses = TRUE
        If trim(notes_on_dcex) <> "" Then case_has_hc_expenses = TRUE
        If trim(notes_on_other_deduction) <> "" Then case_has_hc_expenses = TRUE
        If trim(hc_bils_info) <> "" Then case_has_hc_expenses = TRUE
        If trim(expense_notes) <> "" Then case_has_hc_expenses = TRUE
    End If
    If case_has_hc_expenses = TRUE Then
        Call write_variable_in_CASE_NOTE("===== EXPENSES =====")
    Else
        Call write_variable_in_CASE_NOTE("== No expense detail for this case ==")
    End If

    'Expenses
    Call write_bullet_and_variable_in_CASE_NOTE("Court Ordered Expenses", notes_on_coex)
    Call write_bullet_and_variable_in_CASE_NOTE("Dependent Care Expenses", notes_on_dcex)
    Call write_bullet_and_variable_in_CASE_NOTE("Other Expenses", notes_on_other_deduction)
    Call write_bullet_and_variable_in_CASE_NOTE("Medical Bills", hc_bils_info)
    Call write_bullet_and_variable_in_CASE_NOTE("Expense Detail", expense_notes)

    If case_has_resources = TRUE Then
        Call write_variable_in_CASE_NOTE("===== RESOURCES =====")
    Else
        Call write_variable_in_CASE_NOTE("== No resource/asset detail for this case ==")
    End If
    'Assets
    If confirm_no_account_panel_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Income sources have been reviewed for direct deposit/associated accounts and none were found.")
    Call write_bullet_and_variable_in_CASE_NOTE("Accounts", notes_on_acct)
    Call write_bullet_and_variable_in_CASE_NOTE("Cash", notes_on_cash)
    Call write_bullet_and_variable_in_CASE_NOTE("Cars", notes_on_cars)
    Call write_bullet_and_variable_in_CASE_NOTE("Real Estate", notes_on_rest)
    Call write_bullet_and_variable_in_CASE_NOTE("Other Assets", notes_on_other_assets)

    Call write_variable_in_CASE_NOTE("=====================")
    Call write_bullet_and_variable_in_CASE_NOTE("Notes", hc_other_notes)
    Call write_bullet_and_variable_in_CASE_NOTE("Verifications Needed", hc_verifs_needed)

    If MMIS_updated_check = checked Then Call write_variable_in_CASE_NOTE("* MMIS Updated")
    If hc_tikl_checkbox = checked Then Call write_variable_in_CASE_NOTE("* TIKL set for 45 days from application date.")
    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)

    PF3

End If

If the_process_for_snap = "Application" AND exp_det_case_note_found = False Then
    Call start_a_blank_CASE_NOTE

    If IsDate(snap_denial_date) = TRUE Then
        case_note_header_text = "Expedited Determination: SNAP to be denied"
    Else
        IF snap_exp_yn = "Yes" then
        	case_note_header_text = "Expedited Determination: SNAP appears expedited"
        ELSEIF snap_exp_yn = "No" then
        	case_note_header_text = "Expedited Determination: SNAP does not appear expedited"
        END IF
    End If
    Call write_variable_in_CASE_NOTE(case_note_header_text)
    If interview_date <> "" Then Call write_variable_in_case_note ("* Interview completed on: " & interview_date & " and full Expedited Determination Done")
    If IsDate(snap_denial_date) = TRUE Then
        Call write_variable_in_CASE_NOTE("* SNAP to be denied on " & snap_denial_date & ". Since case is not SNAP eligible, case cannot receive Expedited issuance.")
        If snap_exp_yn = "Yes" Then
            Call write_variable_with_indent_in_CASE_NOTE("Case is determined to meet criteria based upon income alone.")
            Call write_variable_with_indent_in_CASE_NOTE("Expedited approval requires case to be otherwise eligble for SNAP and this does not meet this criteria.")
        ElseIf snap_exp_yn = "No" Then
            Call write_variable_with_indent_in_CASE_NOTE("Expedited SNAP cannot be approved as case does not meet all criteria")
        End If
        Call write_bullet_and_variable_in_CASE_NOTE("Explanation of Denial", snap_denial_explain)
    Else
        IF snap_exp_yn = "Yes" Then
            If trim(exp_snap_approval_date) <> "" Then
                Call write_variable_in_case_note ("* Case is determined to meet criteria and Expedited SNAP can be approved.")
            Else
                Call write_variable_in_case_note ("* Case is determined to meet expedited SNAP criteria, approval not yet completed.")
            End If
        End If
        IF snap_exp_yn = "No" Then Call write_variable_in_case_note ("* Expedited SNAP cannot be approved as case does not meet all criteria")
        If snap_exp_yn = "Yes" Then
            If IsDate(exp_snap_approval_date) = TRUE Then Call write_variable_in_CASE_NOTE("* SNAP EXP approved on " & exp_snap_approval_date & " - " & DateDiff("d", CAF_datestamp, exp_snap_approval_date) & " days after the date of application.")
            Call write_bullet_and_variable_in_CASE_NOTE("Reason for delay", exp_snap_delays)
        End If
    End If
    If trim(app_month_income) <> "" AND trim(app_month_assets) <> "" AND trim(app_month_expenses) <> "" Then
        Call write_variable_in_CASE_NOTE("* Expedited Determination is based on information from application month:")
        Call write_variable_with_indent_in_CASE_NOTE("Income: $" & app_month_income)
        Call write_variable_with_indent_in_CASE_NOTE("Assets: $" & app_month_assets)
        Call write_variable_with_indent_in_CASE_NOTE("Expenses (Shelter & Utilities): $" & app_month_expenses)
    End If

    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)

    PF3
End If
interview_note = FALSE

'Interview Incomfation Detail Case Note
'Navigates to case note, and checks to make sure we aren't in inquiry.
' If SNAP_checkbox = checked OR family_cash = TRUE OR CAF_type = "Application" then
If interview_waived = TRUE Then
    Call start_a_blank_CASE_NOTE

    CALL write_variable_in_CASE_NOTE("Interview for the Renewal was WAIVED")
    CALL write_variable_in_CASE_NOTE("---")
    If the_process_for_snap = "Recertification" Then CALL write_variable_in_CASE_NOTE("Interview for the " & snap_recert_mo & "/" & snap_recert_yr & " SNAP RENEWAL has been waived.")
    If the_process_for_cash = "Recertification" AND family_cash = TRUE Then CALL write_variable_in_CASE_NOTE("Interview for the " & cash_recert_mo & "/" &  cash_recert_yr & " MFIP RENEWAL has been waived.")
    CALL write_variable_in_CASE_NOTE("---")
    CALL write_variable_in_CASE_NOTE("***A Waiver has been granted allowing:")
    CALL write_variable_in_CASE_NOTE("  - Processing of SNAP Annual Renewals to match the process of Six-Month")
    CALL write_variable_in_CASE_NOTE("    Renewals, which do not require interviews.")
    CALL write_variable_in_CASE_NOTE("  - Waiving of Interviews for MFIP cases in some situations.")
    CALL write_variable_in_CASE_NOTE("-- THIS WAIVER CANNOT APPLY TO NEW PROGRAM APPLICATIONS --")
    CALL write_variable_in_CASE_NOTE("Details can be seen in a SIR announcement on 4/14/2021")
    CALL write_variable_in_CASE_NOTE("---")
    CALL write_variable_in_CASE_NOTE(worker_signature)
End If

If interview_required = TRUE AND interview_completed_case_note_found = False Then
    interview_note = TRUE
    Call start_a_blank_CASE_NOTE

    CALL write_variable_in_CASE_NOTE("~ Interview Completed on " & interview_date & " ~")
    Call write_bullet_and_variable_in_CASE_NOTE("Form Datestamp", caf_datestamp)
    Call write_variable_in_CASE_NOTE("* Interview details:")

    Call write_variable_with_indent_in_CASE_NOTE("Conducted via " & interview_type)
    Call write_variable_with_indent_in_CASE_NOTE("Interview was completed with " & interview_with)
    If Used_Interpreter_checkbox = checked Then Call write_variable_with_indent_in_CASE_NOTE("Used interpreter to complete the interview.")
    Call write_variable_with_indent_in_CASE_NOTE("Interview completed on " & interview_date)

    Call write_bullet_and_variable_in_CASE_NOTE("AREP ID Info", arep_id_info)

    If confirm_update_prog = 1 Then
        If snap_intv_date_updated = TRUE AND cash_intv_date_updated = TRUE Then
            prog_updated_for_programs = "SNAP and Cash"
        ElseIf snap_intv_date_updated = TRUE Then
            prog_updated_for_programs = "SNAP"
        ElseIf cash_intv_date_updated = TRUE Then
            prog_updated_for_programs = "Cash"
        End If
        If snap_intv_date_updated = TRUE OR cash_intv_date_updated = TRUE Then CALL write_variable_in_CASE_NOTE("* Interview date entered on PROG for " & prog_updated_for_programs)
    End If
    If do_not_update_prog = 1 Then CALL write_bullet_and_variable_in_CASE_NOTE("PROG WAS NOT UPDATED WITH INTERVIEW DATE, because", no_update_reason)

    Call write_variable_in_CASE_NOTE("----- Programs requested " & progs_list & " -----")

    If CASH_on_CAF_checkbox = checked Then CAF_progs = CAF_progs & ", Cash"
    If SNAP_on_CAF_checkbox = checked Then CAF_progs = CAF_progs & ", SNAP"
    If EMER_on_CAF_checkbox = checked Then CAF_progs = CAF_progs & ", EMER"
    If CAF_progs <> "" Then
        CAF_progs = right(CAF_progs, len(CAF_progs) - 2)
    Else
        CAF_progs = "None"
    End If
    Call write_bullet_and_variable_in_CASE_NOTE("Programs requested in writing on the form", CAF_progs)
    Call write_bullet_and_variable_in_CASE_NOTE("Cash requested", cash_other_req_detail)
    Call write_bullet_and_variable_in_CASE_NOTE("SNAP requested", snap_other_req_detail)
    Call write_bullet_and_variable_in_CASE_NOTE("EMER requested", emer_other_req_detail)
    If family_cash = True Then Call write_variable_in_CASE_NOTE("* Cash request is for FAMILY programs.")
    If adult_cash = True Then Call write_variable_in_CASE_NOTE("* Cash request is for ADULT programs.")

    Call write_variable_in_CASE_NOTE("---")

    If SNAP_checkbox = checked Then
        Call write_variable_in_CASE_NOTE("* SNAP unit consists of " & total_snap_count & " people - " & adult_snap_count & " adults and " & child_snap_count & " children.")
        If included_snap_members <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Members on SNAP grant: " & included_snap_members)
        If counted_snap_members <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Members with income counted ONLY for SNAP: " & counted_snap_members)
        If EATS <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Information on EATS: " & EATS)
    End If
    If cash_checkbox = checked Then
        Call write_variable_in_CASE_NOTE("* CASH unit consists of " & total_cash_count & " people - " & adult_cash_count & " adults and " & child_cash_count & " children.")
        If pregnant_caregiver_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Pregnant Caregiver on Grant.")
        If included_cash_members <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Members on CASH grant: " & included_cash_members)
        If counted_cash_members <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Members with income counted ONLY for CASH: " & counted_cash_members)
    End If
    If EMER_checkbox = checked Then
        Call write_variable_in_CASE_NOTE("* EMER unit consists of " & total_emer_count & " people - " & adult_emer_count & " adults and " & child_emer_count & " children.")
        Call write_variable_with_indent_in_CASE_NOTE("Members on EMER grant: " & included_emer_members)
        If counted_emer_members <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Members with income counted ONLY for EMER: " & counted_emer_members)
    End If
    Call write_bullet_and_variable_in_CASE_NOTE("Relationships", relationship_detail)

    Call write_variable_in_CASE_NOTE("---")

    IF recert_period_checkbox = checked THEN call write_variable_in_CASE_NOTE("* Informed client of recert period.")
    IF R_R_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Rights and Responsibilities explained to client.")
    IF E_and_T_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Client requests to participate with E&T.")
    If elig_req_explained_checkbox Then CALL write_variable_in_CASE_NOTE("* Explained eligbility requirements to client.")
    If benefit_payment_explained_checkbox Then CALL write_variable_in_CASE_NOTE("* Benefits and Payment information explained to client")

    ' Call write_variable_with_indent_in_CASE_NOTE

    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)

    PF3
End If

'Verification NOTE
verifs_needed = replace(verifs_needed, "[Information here creates a SEPARATE CASE/NOTE.]", "")
If trim(verifs_needed) <> "" AND verifications_requested_case_note_found = False Then

    verif_counter = 1
    verifs_needed = trim(verifs_needed)
    If right(verifs_needed, 1) = ";" Then verifs_needed = left(verifs_needed, len(verifs_needed) - 1)
    If left(verifs_needed, 1) = ";" Then verifs_needed = right(verifs_needed, len(verifs_needed) - 1)
    If InStr(verifs_needed, ";") <> 0 Then
        verifs_array = split(verifs_needed, ";")
    Else
        verifs_array = array(verifs_needed)
    End If

    Call start_a_blank_CASE_NOTE

    Call write_variable_in_CASE_NOTE("VERIFICATIONS REQUESTED")

    Call write_bullet_and_variable_in_CASE_NOTE("Verif request form sent on", verif_req_form_sent_date)

    Call write_variable_in_CASE_NOTE("---")

    Call write_variable_in_CASE_NOTE("List of all verifications requested:")
    For each verif_item in verifs_array
        verif_item = trim(verif_item)
        If number_verifs_checkbox = checked Then verif_item = verif_counter & ". " & verif_item
        verif_counter = verif_counter + 1
        Call write_variable_with_indent_in_CASE_NOTE(verif_item)
    Next

    If verifs_postponed_checkbox = checked Then
        Call write_variable_in_CASE_NOTE("---")
        Call write_variable_in_CASE_NOTE("There may be verifications that are postponed to allow for the approval of Expedited SNAP.")
    End If
    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)

    PF3
End If

If qual_questions_yes = TRUE AND caf_qualifying_questions_case_note_found = False Then
    Call start_a_blank_CASE_NOTE

    Call write_variable_in_CASE_NOTE("Qualifying Questions had an answer of 'YES' for at least one question")
    If qual_question_one = "Yes" Then Call write_bullet_and_variable_in_CASE_NOTE("Fraud/DISQ for IPV (program violation)", qual_memb_one)
    If qual_question_two = "Yes" Then Call write_bullet_and_variable_in_CASE_NOTE("SNAP in more than One State", qual_memb_two)
    If qual_question_three = "Yes" Then Call write_bullet_and_variable_in_CASE_NOTE("Fleeing Felon", qual_memb_three)
    If qual_question_four = "Yes" Then Call write_bullet_and_variable_in_CASE_NOTE("Drug Felony", qual_memb_four)
    If qual_question_five = "Yes" Then Call write_bullet_and_variable_in_CASE_NOTE("Parole/Probation Violation", qual_memb_five)
    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)

End If

'MAIN CAF Information NOTE
'Navigates to case note, and checks to make sure we aren't in inquiry.
Call start_a_blank_CASE_NOTE

If CAF_form = "HUF (DHS-8107)" Then
    CALL write_variable_in_CASE_NOTE(CAF_datestamp & " HUF for " & prog_and_type_list & CAF_status)
Else
    CALL write_variable_in_CASE_NOTE(CAF_datestamp & " CAF for " & prog_and_type_list & CAF_status)
End If
Call write_bullet_and_variable_in_CASE_NOTE("Form Received", CAF_form)
'Programs requested
If interview_note = FALSE Then
    If CASH_on_CAF_checkbox = checked Then CAF_progs = CAF_progs & ", Cash"
    If SNAP_on_CAF_checkbox = checked Then CAF_progs = CAF_progs & ", SNAP"
    If EMER_on_CAF_checkbox = checked Then CAF_progs = CAF_progs & ", EMER"
    If CAF_progs <> "" Then
        CAF_progs = right(CAF_progs, len(CAF_progs) - 2)
    Else
        CAF_progs = "None"
    End If
    Call write_bullet_and_variable_in_CASE_NOTE("Programs requested in writing on the form", CAF_progs)
    Call write_bullet_and_variable_in_CASE_NOTE("Cash requested", cash_other_req_detail)
    Call write_bullet_and_variable_in_CASE_NOTE("SNAP requested", snap_other_req_detail)
    Call write_bullet_and_variable_in_CASE_NOTE("EMER requested", emer_other_req_detail)
    If family_cash = True Then Call write_variable_in_CASE_NOTE("* Cash request is for FAMILY programs.")
    If adult_cash = True Then Call write_variable_in_CASE_NOTE("* Cash request is for ADULT programs.")
Else
    Call write_bullet_and_variable_in_CASE_NOTE("Programs information is for", progs_list)
    If family_cash = True Then Call write_variable_in_CASE_NOTE("* Cash request is for FAMILY programs.")
    If adult_cash = True Then Call write_variable_in_CASE_NOTE("* Cash request is for ADULT programs.")
End If

'Household and personal information
If interview_note = FALSE Then
    If SNAP_checkbox = checked Then
        Call write_variable_in_CASE_NOTE("* SNAP unit consists of " & total_snap_count & " people - " & adult_snap_count & " adults and " & child_snap_count & " children.")
        If included_snap_members <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Members on SNAP grant: " & included_snap_members)
        If counted_snap_members <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Members with income counted ONLY for SNAP: " & counted_snap_members)
        If EATS <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Information on EATS: " & EATS)
    End If
    If cash_checkbox = checked Then
        Call write_variable_in_CASE_NOTE("* CASH unit consists of " & total_cash_count & " people - " & adult_cash_count & " adults and " & child_cash_count & " children.")
        If pregnant_caregiver_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Pregnant Caregiver on Grant.")
        If included_cash_members <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Members on CASH grant: " & included_cash_members)
        If counted_cash_members <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Members with income counted ONLY for CASH: " & counted_cash_members)
    End If
    If EMER_checkbox = checked Then
        Call write_variable_in_CASE_NOTE("* EMER unit consists of " & total_emer_count & " people - " & adult_emer_count & " adults and " & child_emer_count & " children.")
        Call write_variable_with_indent_in_CASE_NOTE("Members on EMER grant: " & included_emer_members)
        Call write_variable_with_indent_in_CASE_NOTE("Members with income counted ONLY for EMER: " & counted_emer_members)
    End If
    Call write_bullet_and_variable_in_CASE_NOTE("Relationships", relationship_detail)
End If

first_member = TRUE
For the_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
    If ALL_MEMBERS_ARRAY(gather_detail, the_member) = TRUE Then
        If ALL_MEMBERS_ARRAY(id_required, the_member) = checked Then
            If first_member = TRUE Then
                Call write_variable_in_CASE_NOTE("===== ID REQUIREMENT =====")
                first_member = FALSE
            End If
            Call write_variable_in_CASE_NOTE("* Identity of Memb " & ALL_MEMBERS_ARRAY(memb_numb, the_member) & " verified by: " & right(ALL_MEMBERS_ARRAY(clt_id_verif, the_member), len(ALL_MEMBERS_ARRAY(clt_id_verif, the_member)) - 5) & " and is required.")
            If trim(ALL_MEMBERS_ARRAY(id_detail, the_member)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Details: " & trim(ALL_MEMBERS_ARRAY(id_detail, the_member)))
        End If
    End If
Next

'INCOME
If case_has_income = TRUE Then
    Call write_variable_in_CASE_NOTE("===== INCOME =====")
Else
    Call write_variable_in_CASE_NOTE("== No Income detail Listed for this case. ==")
End If
'JOBS
If ALL_JOBS_PANELS_ARRAY(memb_numb, 0) <> "" Then
    ' Call write_variable_with_indent(variable_name)
    Call write_variable_in_CASE_NOTE("--- JOBS Income ---")
    For each_job = 0 to UBound(ALL_JOBS_PANELS_ARRAY, 2)
        Call write_variable_in_CASE_NOTE("Member " & ALL_JOBS_PANELS_ARRAY(memb_numb, each_job) & " at " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job))
        If ALL_JOBS_PANELS_ARRAY(estimate_only, each_job) = checked Then Call write_variable_in_CASE_NOTE("* This job has not been verified and this is only an estimate.")
        IF ALL_JOBS_PANELS_ARRAY(EI_case_note, each_job) = TRUE Then call write_variable_in_CASE_NOTE("* BUDGET DETAIL ABOUT THIS JOB IN PREVIOUS CASE NOTE.")
        If ALL_JOBS_PANELS_ARRAY(verif_code, each_job) = "Delayed" Then
            Call write_variable_in_CASE_NOTE("* Verification of this job has been delayed for review or approval of Expedited SNAP.")
        ElseIf ALL_JOBS_PANELS_ARRAY(estimate_only, each_job) = unchecked Then
            Call write_variable_in_CASE_NOTE("* Verification - " & ALL_JOBS_PANELS_ARRAY(verif_code, each_job))
        End If
        Call write_bullet_and_variable_in_CASE_NOTE("Verification", ALL_JOBS_PANELS_ARRAY(verif_explain, each_job))
        If ALL_JOBS_PANELS_ARRAY(job_retro_income, each_job) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Retro Income: $" & ALL_JOBS_PANELS_ARRAY(job_retro_income, each_job) & " - " & ALL_JOBS_PANELS_ARRAY(retro_hours, each_job) & " hours.")
        If ALL_JOBS_PANELS_ARRAY(job_prosp_income, each_job) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Prospective Income: $" & ALL_JOBS_PANELS_ARRAY(job_prosp_income, each_job) & " - " & ALL_JOBS_PANELS_ARRAY(prosp_hours, each_job) & " hours.")
        If snap_checkbox = checked Then Call write_variable_with_indent_in_CASE_NOTE("SNAP Budget Detail: Monthly budgeted amount - $" & ALL_JOBS_PANELS_ARRAY(pic_prosp_income, each_job) & " based on $" & ALL_JOBS_PANELS_ARRAY(pic_pay_date_income, each_job) & " paid " & ALL_JOBS_PANELS_ARRAY(pic_pay_freq, each_job) & ". Calculated on " & ALL_JOBS_PANELS_ARRAY(pic_calc_date, each_job))
        If ALL_JOBS_PANELS_ARRAY(budget_explain, each_job) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("About Budget: " & ALL_JOBS_PANELS_ARRAY(budget_explain, each_job))
    Next
End If
Call write_bullet_and_variable_in_CASE_NOTE("JOBS", notes_on_jobs)
Call write_bullet_and_variable_in_CASE_NOTE("Other Earned Income", earned_income)

'BUSI
If ALL_BUSI_PANELS_ARRAY(memb_numb, 0) <> "" Then
    Call write_variable_in_CASE_NOTE("--- BUSI Income ---")
    For each_busi = 0 to UBound(ALL_BUSI_PANELS_ARRAY, 2)
        busi_det_msg = "Self Employment for Memb " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " - BUSI type:" & right(ALL_BUSI_PANELS_ARRAY(busi_type, each_busi), len(ALL_BUSI_PANELS_ARRAY(busi_type, each_busi)) - 4) & "."
        Call write_variable_in_CASE_NOTE(busi_det_msg)

        If ALL_BUSI_PANELS_ARRAY(busi_desc, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Description: " & ALL_BUSI_PANELS_ARRAY(busi_desc, each_busi) & ".")
        If ALL_BUSI_PANELS_ARRAY(busi_structure, each_busi) <> "Select or Type" and ALL_BUSI_PANELS_ARRAY(busi_structure, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Business structure: " & ALL_BUSI_PANELS_ARRAY(busi_structure, each_busi) & ".")
        If ALL_BUSI_PANELS_ARRAY(share_denom, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Clt owns " & ALL_BUSI_PANELS_ARRAY(share_num, each_busi) & "/" & ALL_BUSI_PANELS_ARRAY(share_denom, each_busi) & " of the business.")
        If ALL_BUSI_PANELS_ARRAY(partners_in_HH, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Business also owned by Memb(s) " & ALL_BUSI_PANELS_ARRAY(partners_in_HH, each_busi) & ".")

        se_method_det_msg = "* Self Employment Budgeting method selected: " & ALL_BUSI_PANELS_ARRAY(calc_method, each_busi) & "."
        Call write_variable_in_CASE_NOTE(se_method_det_msg)
        If ALL_BUSI_PANELS_ARRAY(mthd_date, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Method selected on: " & ALL_BUSI_PANELS_ARRAY(mthd_date, each_busi) & ".")
        If ALL_BUSI_PANELS_ARRAY(method_convo_checkbox, each_busi) = checked Then Call write_variable_with_indent_in_CASE_NOTE("The self employment method selected was discussed with the client.")

        If cash_checkbox = checked OR EMER_checkbox = checked Then
            Call write_variable_in_CASE_NOTE("* Cash Income and Expense Detail:")
            cash_income_det = ""
            cash_expense_det = ""

            If ALL_BUSI_PANELS_ARRAY(income_ret_cash, each_busi) <> "" Then
                cash_income_det = "RETRO Income $" & ALL_BUSI_PANELS_ARRAY(income_ret_cash, each_busi) & " - "
                cash_expense_det = "RETRO Expenses $" & ALL_BUSI_PANELS_ARRAY(expense_ret_cash, each_busi) & " - "
            End If
            If ALL_BUSI_PANELS_ARRAY(income_pro_cash, each_busi) <> "" Then
                cash_income_det = cash_income_det & "PROSP Income $" & ALL_BUSI_PANELS_ARRAY(income_pro_cash, each_busi) & " - "
                cash_expense_det = cash_expense_det & "PROSP Expenses $" & ALL_BUSI_PANELS_ARRAY(expense_pro_cash, each_busi) & " - "
            End If
            If ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi) <> "Select or Type" and ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi) <> "" Then cash_income_det = cash_income_det & "Verification: " & ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi)
            If ALL_BUSI_PANELS_ARRAY(cash_expense_verif, each_busi) <> "Select or Type" and ALL_BUSI_PANELS_ARRAY(cash_expense_verif, each_busi) <> "" Then cash_expense_det = cash_expense_det & "Verification: " & ALL_BUSI_PANELS_ARRAY(cash_expense_verif, each_busi)

            Call write_variable_with_indent_in_CASE_NOTE(cash_income_det)
            Call write_variable_with_indent_in_CASE_NOTE(cash_expense_det)
        End If
        If SNAP_checkbox = checked Then
            Call write_variable_in_CASE_NOTE("* SNAP Income and Expense Detail:")
            snap_income_det = ""
            snap_expense_det = ""

            If ALL_BUSI_PANELS_ARRAY(income_ret_snap, each_busi) <> "" Then
                snap_income_det = "RETRO Income $" & ALL_BUSI_PANELS_ARRAY(income_ret_snap, each_busi) & " - "
                snap_expense_det = "RETRO Expenses $" & ALL_BUSI_PANELS_ARRAY(expense_ret_snap, each_busi) & " - "
            End If
            If ALL_BUSI_PANELS_ARRAY(income_pro_snap, each_busi) <> "" Then
                snap_income_det = snap_income_det & "PROSP Income $" & ALL_BUSI_PANELS_ARRAY(income_pro_snap, each_busi) & " - "
                snap_expense_det = snap_expense_det & "PROSP Expenses $" & ALL_BUSI_PANELS_ARRAY(expense_pro_snap, each_busi) & " - "
            End If
            If ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi) <> "Select or Type" and ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi) <> "" Then snap_income_det = snap_income_det & "Verification: " & ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi)
            If ALL_BUSI_PANELS_ARRAY(snap_expense_verif, each_busi) <> "Select or Type" and ALL_BUSI_PANELS_ARRAY(snap_expense_verif, each_busi) <> "" Then snap_expense_det = snap_expense_det & "Verification: " & ALL_BUSI_PANELS_ARRAY(snap_expense_verif, each_busi)

            Call write_variable_with_indent_in_CASE_NOTE(snap_income_det)
            Call write_variable_with_indent_in_CASE_NOTE(snap_expense_det)
            If ALL_BUSI_PANELS_ARRAY(exp_not_allwd, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Expenses from taxes not allowed: " & ALL_BUSI_PANELS_ARRAY(exp_not_allwd, each_busi))
        End If
        rept_hours_det_msg = ""
        min_wg_hours_det_msg = ""
        If ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, each_busi) = "___" Then ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, each_busi) = ""
        If ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, each_busi) = "___" Then ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, each_busi) = ""
        If ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, each_busi) = "___" Then ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, each_busi) = ""
        If ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, each_busi) = "___" Then ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, each_busi) = ""

        If ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, each_busi) <> "" OR ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, each_busi) <> "" Then
            rept_hours_det_msg = rept_hours_det_msg & "Clt reported monthly work hours of: "
            If ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, each_busi) <> "" Then rept_hours_det_msg = rept_hours_det_msg & ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, each_busi) & " retrospecive work and "
            If ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, each_busi) <> "" Then rept_hours_det_msg = rept_hours_det_msg & ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, each_busi) & " prospoective work hrs"
            rept_hours_det_msg = rept_hours_det_msg & ". "
        End If
        If ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, each_busi) <> "" OR ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, each_busi) <> "" Then
            min_wg_hours_det_msg = min_wg_hours_det_msg & "Work earnings indicate Minumun Wage Hours of: "
            If ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, each_busi) <> "" Then min_wg_hours_det_msg = min_wg_hours_det_msg & ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, each_busi) & " retrospective and "
            If ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, each_busi) <> "" Then min_wg_hours_det_msg = min_wg_hours_det_msg & ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, each_busi) & " prospective"
            min_wg_hours_det_msg = min_wg_hours_det_msg & ". "
        End If
        If rept_hours_det_msg <> "" Then Call write_variable_in_CASE_NOTE("* " & rept_hours_det_msg)
        If min_wg_hours_det_msg <> "" Then Call write_variable_in_CASE_NOTE("* " & min_wg_hours_det_msg)
        If ALL_BUSI_PANELS_ARRAY(verif_explain, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Verif Detail: " & ALL_BUSI_PANELS_ARRAY(verif_explain, each_busi))
        If ALL_BUSI_PANELS_ARRAY(budget_explain, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Budget Detail: " & ALL_BUSI_PANELS_ARRAY(budget_explain, each_busi))
    Next
End If
Call write_bullet_and_variable_in_CASE_NOTE("BUSI", notes_on_busi)

'CSES
If show_cses_detail = TRUE Then
    Call write_variable_in_CASE_NOTE("--- Child Support Income ---")
    For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
        If UNEA_INCOME_ARRAY(CS_exists, each_unea_memb) = TRUE Then
            total_cs = 0
            If IsNumeric(UNEA_INCOME_ARRAY(direct_CS_amt, each_unea_memb)) = TRUE Then total_cs = total_cs + UNEA_INCOME_ARRAY(direct_CS_amt, each_unea_memb)
            If IsNumeric(UNEA_INCOME_ARRAY(disb_CS_amt, each_unea_memb)) = TRUE Then total_cs = total_cs + UNEA_INCOME_ARRAY(disb_CS_amt, each_unea_memb)
            If IsNumeric(UNEA_INCOME_ARRAY(disb_CS_arrears_amt, each_unea_memb)) = TRUE Then total_cs = total_cs + UNEA_INCOME_ARRAY(disb_CS_arrears_amt, each_unea_memb)

            Call write_variable_in_CASE_NOTE("* Total child support income for Member " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & ": $" & total_cs)
            If UNEA_INCOME_ARRAY(disb_CS_amt, each_unea_memb) <> "" Then
                cs_disb_inc_det = "Disbursed child support: $" & UNEA_INCOME_ARRAY(disb_CS_amt, each_unea_memb)
                If trim(UNEA_INCOME_ARRAY(disb_CS_notes, each_unea_memb)) <> "" Then cs_disb_inc_det = cs_disb_inc_det & ". Notes: " & UNEA_INCOME_ARRAY(disb_CS_notes, each_unea_memb)

                Call write_variable_with_indent_in_CASE_NOTE(cs_disb_inc_det)
                If trim(UNEA_INCOME_ARRAY(disb_CS_months, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Income was determined using " & UNEA_INCOME_ARRAY(disb_CS_months, each_unea_memb) & " month(s) of disbursement income.")
                If trim(UNEA_INCOME_ARRAY(disb_CS_prosp_budg, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("SNAP prospective budget details: " & UNEA_INCOME_ARRAY(disb_CS_prosp_budg, each_unea_memb))
            End If

            If UNEA_INCOME_ARRAY(disb_CS_arrears_amt, each_unea_memb) <> "" Then
                cs_arrears_inc_det = "Disbursed child support arrears: $" & UNEA_INCOME_ARRAY(disb_CS_arrears_amt, each_unea_memb)
                If trim(UNEA_INCOME_ARRAY(disb_CS_arrears_notes, each_unea_memb)) <> "" Then cs_arrears_inc_det = cs_arrears_inc_det & ". Notes: " & UNEA_INCOME_ARRAY(disb_CS_arrears_notes, each_unea_memb)

                Call write_variable_with_indent_in_CASE_NOTE(cs_arrears_inc_det)
                If trim(UNEA_INCOME_ARRAY(disb_CS_arrears_months, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Income was determined using " & UNEA_INCOME_ARRAY(disb_CS_arrears_months, each_unea_memb) & " month(s) of disbursement income.")
                If trim(UNEA_INCOME_ARRAY(disb_CS_arrears_budg, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("SNAP prospective budget details: " & UNEA_INCOME_ARRAY(disb_CS_arrears_budg, each_unea_memb))
            End If

            If UNEA_INCOME_ARRAY(direct_CS_amt, each_unea_memb) <> "" Then
                direct_cs_inc_det = "Direct child support: $" & UNEA_INCOME_ARRAY(direct_CS_amt, each_unea_memb)
                If trim(UNEA_INCOME_ARRAY(direct_CS_notes, each_unea_memb)) <> "" Then direct_cs_inc_det = direct_cs_inc_det & ". Notes: " & UNEA_INCOME_ARRAY(direct_CS_notes, each_unea_memb)

                Call write_variable_with_indent_in_CASE_NOTE(direct_cs_inc_det)
            End if
        End If
    Next
End If
Call write_bullet_and_variable_in_CASE_NOTE("Other Child Support Income", notes_on_cses)

'UNEA
For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
    If UNEA_INCOME_ARRAY(SSA_exists, each_unea_memb) = TRUE Then
        rsdi_income_det = ""
        If trim(UNEA_INCOME_ARRAY(UNEA_RSDI_amt, each_unea_memb)) <> "" Then rsdi_income_det = rsdi_income_det & "RSDI: $" & UNEA_INCOME_ARRAY(UNEA_RSDI_amt, each_unea_memb)
        If trim(UNEA_INCOME_ARRAY(UNEA_RSDI_notes, each_unea_memb)) <> "" Then rsdi_income_det = rsdi_income_det & ". Notes: " & UNEA_INCOME_ARRAY(UNEA_RSDI_notes, each_unea_memb)

        ssi_income_det = ""
        If trim(UNEA_INCOME_ARRAY(UNEA_SSI_amt, each_unea_memb)) <> "" Then ssi_income_det = ssi_income_det & "SSI: $" & UNEA_INCOME_ARRAY(UNEA_SSI_amt, each_unea_memb)
        If trim(UNEA_INCOME_ARRAY(UNEA_SSI_notes, each_unea_memb)) <> "" Then ssi_income_det = ssi_income_det & ". Notes: " & UNEA_INCOME_ARRAY(UNEA_SSI_notes, each_unea_memb)

        Call write_variable_in_CASE_NOTE("* Member " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " SSA income:")
        If rsdi_income_det <> "" Then Call write_variable_with_indent_in_CASE_NOTE(rsdi_income_det)
        If ssi_income_det <> "" Then Call write_variable_with_indent_in_CASE_NOTE(ssi_income_det)
    End If
Next
Call write_bullet_and_variable_in_CASE_NOTE("Other SSA Income", notes_on_ssa_income)
Call write_bullet_and_variable_in_CASE_NOTE("VA Income", notes_on_VA_income)
Call write_bullet_and_variable_in_CASE_NOTE("Workers Comp Income", notes_on_WC_income)

For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
    If UNEA_INCOME_ARRAY(UC_exists, each_unea_memb) = TRUE Then
        uc_income_det_one = ""
        uc_income_det_two = ""
        If trim(UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb)) <> "" Then
            uc_income_det_one = uc_income_det_one & "Budgeted UC weekly amount: $" & UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb) & ".; "
            uc_income_det_one = uc_income_det_one & "UC weekly gross income: $" & UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb) & ". "
            If trim(UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb)) <> "" Then uc_income_det_one = uc_income_det_one & "Deduction amount allowed: $" & UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb) & ". "
            If trim(UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, each_unea_memb)) <> "" Then uc_income_det_one = uc_income_det_one & "Deduction amount excluded: $" & UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, each_unea_memb) & ". "
        Else
            uc_income_det_one = uc_income_det_one & "Budgeted UC weekly amount: $" & UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb) & ".; "
            If trim(UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb)) <> "" Then uc_income_det_one = uc_income_det_one & "Deduction amount allowed: $" & UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb) & ". "
            If trim(UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, each_unea_memb)) <> "" Then uc_income_det_one = uc_income_det_one & "Deduction amount excluded: $" & UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, each_unea_memb) & ". "
        End If
        If trim(UNEA_INCOME_ARRAY(UNEA_UC_account_balance, each_unea_memb)) <> "" Then uc_income_det_one = uc_income_det_one & "Current UC account balance: $" & UNEA_INCOME_ARRAY(UNEA_UC_account_balance, each_unea_memb) & ". "
        If trim(UNEA_INCOME_ARRAY(UNEA_UC_retro_amt, each_unea_memb)) <> "" Then uc_income_det_two = uc_income_det_two & "UC Retro Income: $" & UNEA_INCOME_ARRAY(UNEA_UC_retro_amt, each_unea_memb) & ". "
        If trim(UNEA_INCOME_ARRAY(UNEA_UC_prosp_amt, each_unea_memb)) <> "" Then uc_income_det_two = uc_income_det_two & "UC Prosp Income: $" & UNEA_INCOME_ARRAY(UNEA_UC_prosp_amt, each_unea_memb) & ". "
        If trim(UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, each_unea_memb)) <> "" Then uc_income_det_two = uc_income_det_two & "UC SNAP budgeted Income: $" & UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, each_unea_memb) & ". "

        Call write_variable_in_CASE_NOTE("* Member " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " Unemployment Income:")
        If trim(UNEA_INCOME_ARRAY(UNEA_UC_start_date, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("UC Income started on: " & UNEA_INCOME_ARRAY(UNEA_UC_start_date, each_unea_memb) & ". ")
        If uc_income_det_one <> "" Then Call write_variable_with_indent_in_CASE_NOTE(uc_income_det_one)
        If uc_income_det_two <> "" Then Call write_variable_with_indent_in_CASE_NOTE(uc_income_det_two)
        If IsDate(UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb)) = TRUE Then Call write_variable_with_indent_in_CASE_NOTE("TIKL set to check for end of UC on: " & UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb))
        If trim(UNEA_INCOME_ARRAY(UNEA_UC_notes, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Notes: " & UNEA_INCOME_ARRAY(UNEA_UC_notes, each_unea_memb))
    End If
Next
Call write_bullet_and_variable_in_CASE_NOTE("Other UC Income", other_uc_income_notes)
Call write_bullet_and_variable_in_CASE_NOTE("Unearned Income", notes_on_other_UNEA)

If case_has_personal = TRUE Then Call write_variable_in_CASE_NOTE("===== PERSONAL =====")

Call write_bullet_and_variable_in_CASE_NOTE("Citizenship/ID", cit_id)
Call write_bullet_and_variable_in_CASE_NOTE("IMIG", IMIG)
Call write_bullet_and_variable_in_CASE_NOTE("School", SCHL)
Call write_bullet_and_variable_in_CASE_NOTE("Changes", case_changes)
Call write_bullet_and_variable_in_CASE_NOTE("DISA", DISA)
Call write_bullet_and_variable_in_CASE_NOTE("FACI", FACI)
Call write_bullet_and_variable_in_CASE_NOTE("PREG", PREG)
Call write_bullet_and_variable_in_CASE_NOTE("Absent Parent", ABPS)
If CS_forms_sent_date <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Child Support Forms given/sent to client on " & CS_forms_sent_date)
Call write_bullet_and_variable_in_CASE_NOTE("AREP", AREP)

'Address Detail
If address_confirmation_checkbox = checked Then Call write_variable_in_CASE_NOTE("* The address on ADDR was reviewed and is correct.")
If homeless_yn = "Yes" Then Call write_variable_in_CASE_NOTE("* Household is homeless.")
Call write_variable_in_CASE_NOTE("* Client reports living in county " & addr_county)
Call write_bullet_and_variable_in_CASE_NOTE("Living Situation", living_situation)
Call write_bullet_and_variable_in_CASE_NOTE("Address Detail", notes_on_address)

'DISQ
Call write_bullet_and_variable_in_CASE_NOTE("DISQ", DISQ)

'WREG and ABAWD
Call write_bullet_and_variable_in_CASE_NOTE("WREG", notes_on_wreg)
all_abawd_notes = notes_on_abawd & notes_on_abawd_two & notes_on_abawd_three
Call write_bullet_and_variable_in_CASE_NOTE("ABAWD", all_abawd_notes)

Call write_bullet_and_variable_in_CASE_NOTE("Medicare", MEDI)
Call write_bullet_and_variable_in_CASE_NOTE("Diet", DIET)

'MFIP-DWP information
Call write_bullet_and_variable_in_CASE_NOTE("Time Tracking (MFIP)", notes_on_time)
Call write_bullet_and_variable_in_CASE_NOTE("MFIP Sanction", notes_on_sanction)
Call write_bullet_and_variable_in_CASE_NOTE("MF/DWP Employment Services", EMPS)
If MFIP_DVD_checkbox = checked Then Call write_variable_in_CASE_NOTE("* MFIP financial orientation DVD sent to participant(s).")

If case_has_expenses = TRUE Then
    Call write_variable_in_CASE_NOTE("===== EXPENSES =====")
Else
    Call write_variable_in_CASE_NOTE("== No expense detail for this case ==")
End If
'SHEL
Call write_bullet_and_variable_in_CASE_NOTE("Shelter Expense", "$" & total_shelter_amount)

If InStr(full_shelter_details, "*") <> 0 Then
    shelter_detail_array = split(full_shelter_details, "*")
Else
    shelter_detail_array = array(full_shelter_details)
End If
If full_shelter_details <> "" Then
    For each shel_info in shelter_detail_array
        shel_info = trim(shel_info)
        Call write_variable_with_indent_in_CASE_NOTE(shel_info)
    Next
End If
'HEST/ACUT
Call write_bullet_and_variable_in_CASE_NOTE("Actual Utility Expenses", notes_on_acut)
If hest_information <> "Select ALLOWED HEST" Then Call write_variable_in_CASE_NOTE("* Standard Utility expenses: " & hest_information)

'Expenses
Call write_bullet_and_variable_in_CASE_NOTE("Court Ordered Expenses", notes_on_coex)
Call write_bullet_and_variable_in_CASE_NOTE("Dependent Care Expenses", notes_on_dcex)
Call write_bullet_and_variable_in_CASE_NOTE("Other Expenses", notes_on_other_deduction)
Call write_bullet_and_variable_in_CASE_NOTE("Expense Detail", expense_notes)
Call write_bullet_and_variable_in_CASE_NOTE("FS Medical Expenses", FMED)

If case_has_resources = TRUE Then
    Call write_variable_in_CASE_NOTE("===== RESOURCES =====")
Else
    Call write_variable_in_CASE_NOTE("== No resource/asset detail for this case ==")
End If
'Assets
If confirm_no_account_panel_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Income sources have been reviewed for direct deposit/associated accounts and none were found.")
Call write_bullet_and_variable_in_CASE_NOTE("Accounts", notes_on_acct)
Call write_bullet_and_variable_in_CASE_NOTE("Cash", notes_on_cash)
Call write_bullet_and_variable_in_CASE_NOTE("Cars", notes_on_cars)
Call write_bullet_and_variable_in_CASE_NOTE("Real Estate", notes_on_rest)
Call write_bullet_and_variable_in_CASE_NOTE("Other Assets", notes_on_other_assets)

Call write_variable_in_CASE_NOTE("===== Case Information =====")
'Next review
If trim(next_er_month) <> "" Then Call write_bullet_and_variable_in_CASE_NOTE("Next ER", next_er_month & "/" & next_er_year)

IF application_signed_checkbox = checked THEN
	CALL write_variable_in_CASE_NOTE("* Application was signed.")
Else
	CALL write_variable_in_CASE_NOTE("* Application was not signed.")
END IF
IF eDRS_sent_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* eDRS sent.")
IF updated_MMIS_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Updated MMIS.")
IF WF1_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Workforce referral made.")

IF Sent_arep_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Sent form(s) to AREP.")
IF intake_packet_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Client received intake packet.")
IF IAA_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* IAAs/OMB given to client.")

IF client_delay_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* PND2 updated to show client delay.")
If TIKL_checkbox Then CALL write_variable_in_CASE_NOTE("* TIKL set to take action on " & DateAdd("d", 30, CAF_datestamp))
If client_delay_TIKL_checkbox Then CALL write_variable_in_CASE_NOTE("* TIKL set to update PND2 for Client Delay on " & DateAdd("d", 10, CAF_datestamp))

If qual_questions_yes = FALSE Then Call write_variable_in_CASE_NOTE("* All Qualifying Questions answered 'No'.")
Call write_bullet_and_variable_in_CASE_NOTE("Notes", other_notes)
If trim(verifs_needed) <> "" Then Call write_variable_in_CASE_NOTE("** VERIFICATIONS REQUESTED - See previous case note for detail")
' IF move_verifs_needed = False THEN CALL write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)			'IF global variable move_verifs_needed = False (on FUNCTIONS FILE), it'll case note at the bottom.
CALL write_bullet_and_variable_in_CASE_NOTE("Actions taken", actions_taken)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

IF SNAP_recert_is_likely_24_months = TRUE THEN					'if we determined on stat/revw that the next SNAP recert date isn't 12 months beyond the entered footer month/year
	TIKL_for_24_month = msgbox("Your SNAP recertification date is listed as " & SNAP_recert_date & " on STAT/REVW. Do you want set a TIKL on " & dateadd("m", "-1", SNAP_recert_compare_date) & " for 12 month contact?" & vbCR & vbCR & "NOTE: Clicking yes will navigate away from CASE/NOTE saving your case note.", VBYesNo)
	IF TIKL_for_24_month = vbYes THEN 												'if the select YES then we TIKL using our custom functions.
		'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
        Call create_TIKL("If SNAP is open, review to see if 12 month contact letter is needed. DAIL scrubber can send 12 Month Contact Letter if used on this TIKL.", 0, dateadd("m", "-1", SNAP_recert_compare_date), False, TIKL_note_text)
	END IF
END IF

end_msg = "Success! " & CAF_form & " has been successfully noted. Please remember to run the Approved Programs, Closed Programs, or Denied Programs scripts if  results have been APP'd."
If do_not_update_prog = 1 Then end_msg = end_msg & vbNewLine & vbNewLine & "It was selected that PROG would NOT be updated because " & no_update_reason
If interview_waived = TRUE Then end_msg = "INTERVIEW WAIVED" & vbCR & vbCr & end_msg

script_end_procedure_with_error_report(end_msg)
