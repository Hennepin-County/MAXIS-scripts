
'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - CAF.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 720                     'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
'END OF stats block=========================================================================================================
run_locally = true
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

'FUNCTIONS==================================================================================================================
Function Generate_Client_List(list_for_dropdown, initial_text)

	memb_row = 5
    list_for_dropdown = initital_text

	Call navigate_to_MAXIS_screen ("STAT", "MEMB")
	Do
		EMReadScreen ref_numb, 2, memb_row, 3
		If ref_numb = "  " Then Exit Do
		EMWriteScreen ref_numb, 20, 76
		transmit
		EMReadScreen first_name, 12, 6, 63
		EMReadScreen last_name, 25, 6, 30
		client_info = client_info & "~" & ref_numb & " - " & replace(first_name, "_", "") & " " & replace(last_name, "_", "")
		memb_row = memb_row + 1
	Loop until memb_row = 20

	client_info = right(client_info, len(client_info) - 1)
	client_list_array = split(client_info, "~")

	For each person in client_list_array
		list_for_dropdown = list_for_dropdown & chr(9) & person
	Next

End Function

function read_EATS_panel()
    call navigate_to_MAXIS_screen("stat", "eats")

    'Now it checks for the total number of panels. If there's 0 Of 0 it'll exit the function for you so as to save oodles of time.
    EMReadScreen panel_total_check, 6, 2, 73
    IF panel_total_check = "0 Of 0" THEN exit function		'Exits out if there's no panel info
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

function HH_comp_dialog(HH_member_array)
'--- This function creates an array of all household members in a MAXIS case, and allows users to select which members to seek/add information to add to edit boxes in dialogs.
'~~~~~ HH_member_array: should be HH_member_array for function to work
'===== Keywords: MAXIS, member, array, dialog
	CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.

    member_count = 0
    adult_cash_count = 0
    child_cash_count = 0
    adult_snap_count = 0
    child_snap_count = 0
    adult_emer_count = 0
    child_emer_count = 0
	DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
		EMReadscreen ref_nbr, 2, 4, 33
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

        ReDim Preserve ALL_MEMBERS_ARRAY(clt_notes, member_count)

        ALL_MEMBERS_ARRAY(memb_numb, member_count) = ref_nbr
        ALL_MEMBERS_ARRAY(clt_name, member_count) = last_name & ", " & first_name & " " & mid_initial
        ALL_MEMBERS_ARRAY(full_clt, member_count) = ref_nbr & " - " & first_name & " " & last_name

        If cash_checkbox = checked Then
            ALL_MEMBERS_ARRAY(include_cash_checkbox, member_count) = checked
            ALL_MEMBERS_ARRAY(count_cash_checkbox, member_count) = checked
            If memb_age > 18 then
                adult_cash_count = adult_cash_count + 1
            Else
                child_cash_count = child_cash_count + 1
            End If
        End If
        If SNAP_checkbox = checked Then
            ALL_MEMBERS_ARRAY(include_snap_checkbox, member_count) = checked
            ALL_MEMBERS_ARRAY(count_snap_checkbox, member_count) = checked
            If memb_age > 21 then
                adult_snap_count = adult_snap_count + 1
            Else
                child_snap_count = child_snap_count + 1
            End If
        End If
        If EMER_checkbox = checked Then
            ALL_MEMBERS_ARRAY(include_emer_checkbox, member_count) = checked
            ALL_MEMBERS_ARRAY(count_emer_checkbox, member_count) = checked
            If memb_age > 18 then
                adult_emer_count = adult_emer_count + 1
            Else
                child_emer_count = child_emer_count + 1
            End If
        End If

		client_string = ref_nbr & last_name & first_name & mid_initial
		client_array = client_array & client_string & "|"
		transmit
		Emreadscreen edit_check, 7, 24, 2
        member_count = member_count + 1
	LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

    Call navigate_to_MAXIS_screen("STAT", "PARE")
    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        EMWriteScreen ALL_MEMBERS_ARRAY(memb_numb, each_member), 20, 76
        transmit

        EMReadScreen panel_check, 14, 24, 13
        If panel_check <> "DOES NOT EXIST" Then
            pare_row = 8
            Do
                EMReadScreen child_ref_nbr, 2, pare_row, 24
                EMReadScreen rela_type, 1, pare_row, 53
                EMReadScreen rela_verif, 2, pare_row, 71

                If rela_type = "1" then relationship_type = "Parent"
                If rela_type = "2" then relationship_type = "Stepparent"
                If rela_type = "3" then relationship_type = "Grandparent"
                If rela_type = "4" then relationship_type = "Relative Caregiver"
                If rela_type = "5" then relationship_type = "Foster parent"
                If rela_type = "6" then relationship_type = "Caregiver"
                If rela_type = "7" then relationship_type = "Guardian"
                If rela_type = "8" then relationship_type = "Relative"

                If rela_verif = "BC" Then relationship_verif = "Birth Certificate"
                If rela_verif = "AR" Then relationship_verif = "Adoption Records"
                If rela_verif = "LG" Then relationship_verif = "Legal Guardian"
                If rela_verif = "RE" Then relationship_verif = "Religious Records"
                If rela_verif = "HR" Then relationship_verif = "Hospital Records"
                If rela_verif = "RP" Then relationship_verif = "Recognition of Parantage"
                If rela_verif = "OT" Then relationship_verif = "Other"
                If rela_verif = "NO" Then relationship_verif = "NONE"

                If child_ref_nbr <> "__" Then relationship_detail = relationship_detail & "Memb " & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " is the " & relationship_type & " of Memb " & child_ref_nbr & " - Verif: " & relationship_verif & "; "
                pare_row = pare_row + 1
            Loop Until child_ref_nbr = "__"
        End If
    Next

    client_array = TRIM(client_array)
    client_array = split(client_array, "|")
    If SNAP_checkbox = checked then call read_EATS_panel

	' client_array = TRIM(client_array)
	' test_array = split(client_array, "|")
	' total_clients = Ubound(test_array)			'setting the upper bound for how many spaces to use from the array
    '
	' DIM all_client_array()
	' ReDim all_clients_array(total_clients, 1)
    '
	' FOR x = 0 to total_clients				'using a dummy array to build in the autofilled check boxes into the array used for the dialog.
	' 	Interim_array = split(client_array, "|")
	' 	all_clients_array(x, 0) = Interim_array(x)
	' 	all_clients_array(x, 1) = 1
	' NEXT

    ' 'THIS IS THE BASE FOR THE DIALOG'
    ' BeginDialog Dialog1, 0, 0, 446, 130, "HH Composition Dialog"
    '   Text 10, 10, 250, 10, "This dialog will clarify the household relationships and details for the case."
    '   Text 105, 25, 100, 10, "Included and Counted  in Grant"
    '   Text 110, 40, 20, 10, "Cash"
    '   Text 145, 40, 20, 10, "SNAP"
    '   Text 180, 40, 20, 10, "EMER"
    '   Text 230, 25, 90, 10, "Income Counted - Deeming"
    '   Text 230, 40, 20, 10, "Cash"
    '   Text 265, 40, 20, 10, "SNAP"
    '   Text 300, 40, 20, 10, "EMER"
    '   GroupBox 330, 5, 110, 100, "HH Count by program"
    '   Text 335, 15, 100, 20, "Enter the number of adults and children for each program"
    '   Text 370, 35, 20, 10, "Adults"
    '   Text 400, 35, 30, 10, "Children"
    '   Text 345, 50, 20, 10, "Cash"
    '   EditBox 370, 45, 20, 15, adult_cash_count
    '   EditBox 405, 45, 20, 15, child_cash_count
    '   Text 345, 70, 20, 10, "SNAP"
    '   EditBox 370, 65, 20, 15, adult_snap_count
    '   EditBox 405, 65, 20, 15, child_snap_count
    '   Text 345, 90, 25, 10, "EMER"
    '   EditBox 370, 85, 20, 15, adult_emer_count
    '   EditBox 405, 85, 20, 15, child_emer_count
    '   Text 10, 55, 100, 10, "CLIENT NAME"
    '   CheckBox 115, 55, 10, 10, "", include_checkbox
    '   CheckBox 150, 55, 10, 10, "", include_snap_checkbox
    '   CheckBox 185, 55, 10, 10, "", include_emer_checkbox
    '   CheckBox 235, 55, 10, 10, "", count_cash_checkbox
    '   CheckBox 270, 55, 10, 10, "", count_snap_checkbox
    '   CheckBox 305, 55, 10, 10, "", count_emer_checkbox
    '   Text 10, 95, 25, 10, "EATS:"
    '   EditBox 35, 90, 290, 15, eats_detail
    '   Text 10, 115, 90, 10, "Household Relationships:"
    '   EditBox 105, 110, 220, 15, relationship_detail
    '   ButtonGroup ButtonPressed
    '     OkButton 335, 110, 50, 15
    '     CancelButton 390, 110, 50, 15
    '   Text 10, 70, 100, 10, "CLIENT NAME"
    '   CheckBox 115, 70, 10, 10, "", Check7
    '   CheckBox 150, 70, 10, 10, "", Check8
    '   CheckBox 185, 70, 10, 10, "", Check9
    '   CheckBox 235, 70, 10, 10, "", Check10
    '   CheckBox 270, 70, 10, 10, "", Check11
    '   CheckBox 305, 70, 10, 10, "", Check12
    ' EndDialog


    Do
        Do
            err_msg = ""
            adult_cash_count = adult_cash_count & ""
            child_cash_count = child_cash_count & ""
            adult_snap_count = adult_snap_count & ""
            child_snap_count = child_snap_count & ""
            adult_emer_count = adult_emer_count & ""
            child_emer_count = child_emer_count & ""

            dlg_len = 115 + (15 * UBound(ALL_MEMBERS_ARRAY, 2))
            if dlg_len < 130 Then dlg_len = 130
            BeginDialog Dialog1, 0, 0, 446, dlg_len, "HH Composition Dialog"
              Text 10, 10, 250, 10, "This dialog will clarify the household relationships and details for the case."
              Text 105, 25, 100, 10, "Included and Counted  in Grant"
              Text 110, 40, 20, 10, "Cash"
              Text 145, 40, 20, 10, "SNAP"
              Text 180, 40, 20, 10, "EMER"
              Text 230, 25, 90, 10, "Income Counted - Deeming"
              Text 230, 40, 20, 10, "Cash"
              Text 265, 40, 20, 10, "SNAP"
              Text 300, 40, 20, 10, "EMER"
              GroupBox 330, 5, 110, 100, "HH Count by program"
              Text 335, 15, 100, 20, "Enter the number of adults and children for each program"
              Text 370, 35, 20, 10, "Adults"
              Text 400, 35, 30, 10, "Children"
              Text 345, 50, 20, 10, "Cash"
              EditBox 370, 45, 20, 15, adult_cash_count
              EditBox 405, 45, 20, 15, child_cash_count
              Text 345, 70, 20, 10, "SNAP"
              EditBox 370, 65, 20, 15, adult_snap_count
              EditBox 405, 65, 20, 15, child_snap_count
              Text 345, 90, 25, 10, "EMER"
              EditBox 370, 85, 20, 15, adult_emer_count
              EditBox 405, 85, 20, 15, child_emer_count
              y_pos = 55
              For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                  Text 10, y_pos, 100, 10, ALL_MEMBERS_ARRAY(clt_name, each_member)
                  CheckBox 115, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(include_cash_checkbox, each_member)
                  CheckBox 150, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(include_snap_checkbox, each_member)
                  CheckBox 185, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(include_emer_checkbox, each_member)
                  CheckBox 235, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(count_cash_checkbox, each_member)
                  CheckBox 270, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(count_snap_checkbox, each_member)
                  CheckBox 305, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(count_emer_checkbox, each_member)
                  y_pos = y_pos + 15
              Next
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

            If cash_checkbox = checked Then
                If trim(adult_cash_count) = "" Then adult_cash_count = 0
                If trim(child_cash_count) = "" Then child_cash_count = 0
                If IsNumeric(adult_cash_count) = False and IsNumeric(child_cash_count) = False Then err_msg = err_msg & vbNewLine & "* Enter a valid count for the number of adults and children for the Cash program."
            End If
            If SNAP_checkbox = checked then
                If trim(adult_snap_count) = "" Then adult_snap_count = 0
                If trim(child_snap_count) = "" Then child_snap_count = 0
                If IsNumeric(adult_snap_count) = False and IsNumeric(child_snap_count) = False Then err_msg = err_msg & vbNewLine & "* Enter a valid count for the number of adults and children for the SNAP program."
                If trim(EATS) = "" Then err_msg = err_msg & vbNewLine & "* Clarify who purchases and prepares together since SNAP is being considered."
            End If
            If EMER_checkbox = checked Then
                If trim(adult_emer_count) = "" Then adult_emer_count = 0
                If trim(child_emer_count) = "" Then child_emer_count = 0
                If IsNumeric(adult_emer_count) = False and IsNumeric(child_emer_count) = False Then err_msg = err_msg & vbNewLine & "* Enter a valid count for the number of adults and children for the EMER program."
            End If
            adult_cash_count = adult_cash_count * 1
            child_cash_count = child_cash_count * 1
            adult_snap_count = adult_snap_count * 1
            child_snap_count = child_snap_count * 1
            adult_emer_count = adult_emer_count * 1
            child_emer_count = child_emer_count * 1

        Loop until err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE

    HH_member_array = ""

    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        If ALL_MEMBERS_ARRAY(include_cash_checkbox, each_member) = checked Then
            HH_member_array = HH_member_array & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " "
        ElseIf ALL_MEMBERS_ARRAY(include_snap_checkbox, each_member) = checked Then
            HH_member_array = HH_member_array & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " "
        ElseIf ALL_MEMBERS_ARRAY(include_emer_checkbox, each_member) = checked Then
            HH_member_array = HH_member_array & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " "
        ElseIf ALL_MEMBERS_ARRAY(count_cash_checkbox, each_member) = checked Then
            HH_member_array = HH_member_array & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " "
        ElseIf ALL_MEMBERS_ARRAY(count_snap_checkbox, each_member) = checked Then
            HH_member_array = HH_member_array & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " "
        ElseIf ALL_MEMBERS_ARRAY(count_emer_checkbox, each_member) = checked Then
            HH_member_array = HH_member_array & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " "
        End If
    Next
	' BEGINDIALOG HH_memb_dialog, 0, 0, 241, (35 + (total_clients * 15)), "HH Member Dialog"   'Creates the dynamic dialog. The height will change based on the number of clients it finds.
	' 	Text 10, 5, 105, 10, "Household members to look at:"
	' 	FOR i = 0 to total_clients										'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
	' 		IF all_clients_array(i, 0) <> "" THEN checkbox 10, (20 + (i * 15)), 160, 10, all_clients_array(i, 0), all_clients_array(i, 1)  'Ignores and blank scanned in persons/strings to avoid a blank checkbox
	' 	NEXT
	' 	ButtonGroup ButtonPressed
	' 	OkButton 185, 10, 50, 15
	' 	CancelButton 185, 30, 50, 15
	' ENDDIALOG
	' 												'runs the dialog that has been dynamically created. Streamlined with new functions.
	' Dialog HH_memb_dialog
	' If buttonpressed = 0 then stopscript
	' check_for_maxis(True)
    '
	' HH_member_array = ""
    '
	' FOR i = 0 to total_clients
	' 	IF all_clients_array(i, 0) <> "" THEN 						'creates the final array to be used by other scripts.
	' 		IF all_clients_array(i, 1) = 1 THEN						'if the person/string has been checked on the dialog then the reference number portion (left 2) will be added to new HH_member_array
	' 			'msgbox all_clients_
	' 			HH_member_array = HH_member_array & left(all_clients_array(i, 0), 2) & " "
	' 		END IF
	' 	END IF
	' NEXT

	HH_member_array = TRIM(HH_member_array)							'Cleaning up array for ease of use.
    HH_member_array = REPLACE(HH_member_array, "  ", " ")
	HH_member_array = SPLIT(HH_member_array, " ")
    ' MsgBox "All members ubound - " & UBound(ALL_MEMBERS_ARRAY, 2)
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
                        EMReadScreen counted_date_year, 2, bene_yr_row, 14			'reading counted year date
                        abawd_counted_months_string = counted_date_month & "/" & counted_date_year
                        abawd_info_list = abawd_info_list & ", " & abawd_counted_months_string			'adding variable to list to add to array
                        abawd_counted_months = abawd_counted_months + 1				'adding counted months
                    END IF

                    'declaring & splitting the abawd months array
                    If left(abawd_info_list, 1) = "," then abawd_info_list = right(abawd_info_list, len(abawd_info_list) - 1)
                    abawd_months_array = Split(abawd_info_list, ",")

                    'counting and checking for second set of ABAWD months
                    IF is_counted_month = "Y" or is_counted_month = "N" THEN
                        EMReadScreen counted_date_year, 2, bene_yr_row, 14			'reading counted year date
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
                ALL_MEMBERS_ARRAY(numb_abawd_used, each_member) = abawd_counted_months
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
                ALL_MEMBERS_ARRAY(numb_banked_mo, each_member) = read_counter * 1

                variable_written_to = variable_written_to & "Member " & HH_member & "- " & WREG_status & ", " & abawd_status & "; "
            End If
        END IF
    Next
End function

function update_wreg_and_abawd_notes()
    notes_on_wreg = ""
    full_abawd_info = ""
    notes_on_abawd = ""
    notes_on_abawd_two = ""
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
                    If left(ALL_MEMBERS_ARRAY(clt_abawd_status, each_member), 2) = "13" Then clt_currently_is = "BANKED"
                End If
                If clt_currently_is <> "" Then
                    full_abawd_info = full_abawd_info & " currently using " & clt_currently_is & " months."
                End If
                If ALL_MEMBERS_ARRAY(numb_abawd_used, each_member) <> "" OR trim(ALL_MEMBERS_ARRAY(list_abawd_mo, each_member)) <> "" Then full_abawd_info = full_abawd_info & " ABAWD months used: " & ALL_MEMBERS_ARRAY(numb_abawd_used, each_member) & " - " & ALL_MEMBERS_ARRAY(list_abawd_mo, each_member)
                If trim(ALL_MEMBERS_ARRAY(first_second_set, each_member)) <> "" Then full_abawd_info = full_abawd_info & " 2nd Set used starting: " & ALL_MEMBERS_ARRAY(first_second_set, each_member)
                If trim(ALL_MEMBERS_ARRAY(explain_no_second, each_member)) <> "" Then full_abawd_info = full_abawd_info & " 2nd Set not available due to: " & ALL_MEMBERS_ARRAY(explain_no_second, each_member)
                If trim(ALL_MEMBERS_ARRAY(numb_banked_mo, each_member)) <> "" Then full_abawd_info = full_abawd_info & " Banked months used: " & ALL_MEMBERS_ARRAY(numb_banked_mo, each_member)
                If trim(ALL_MEMBERS_ARRAY(clt_abawd_notes, each_member)) <> "" Then full_abawd_info = full_abawd_info & " Notes: " & ALL_MEMBERS_ARRAY(clt_abawd_notes, each_member)
                full_abawd_info = full_abawd_info & "; "
            End If
        End If
    Next
    if right(notes_on_wreg, 2) = "; " Then notes_on_wreg = left(notes_on_wreg, len(notes_on_wreg) - 2)
    if right(full_abawd_info, 2) = "; " Then full_abawd_info = left(full_abawd_info, len(full_abawd_info) - 2)
    If len(full_abawd_info) > 400 Then
        notes_on_abawd = left(full_abawd_info, 400)
        notes_on_abawd_two = right(full_abawd_info, len(full_abawd_info) - 400)
    Else
        notes_on_abawd = full_abawd_info
    End If
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

function update_shel_notes()
    total_shelter_amount = 0
    full_shelter_details = ""
    shelter_details = ""
    shelter_details_two = ""

    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        If ALL_MEMBERS_ARRAY(shel_exists, each_member) = TRUE Then
            full_shelter_details = full_shelter_details & "; M" & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " shelter expense(s): "
            If ALL_MEMBERS_ARRAY(shel_retro_rent_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_rent_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_rent_verif, each_member) <> "Select one" Then
                If ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_retro_rent_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, each_member) Then
                        full_shelter_details = full_shelter_details & "Rent $" & ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member)) - 4) & ". "
                    Else
                        full_shelter_details = full_shelter_details & "change in Rent retro - $" & ALL_MEMBERS_ARRAY(shel_retro_rent_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_rent_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_rent_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member)) - 4) & ". "
                    End If
                    total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, each_member)
                Else
                    full_shelter_details = full_shelter_details & "Rent (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_rent_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_rent_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_rent_verif, each_member)) - 4) & ". "
                End If
            ElseIf ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member) <> "Select one" Then
                full_shelter_details = full_shelter_details & "Rent (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member)) - 4) & ". "
                total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, each_member)
            End If

            If ALL_MEMBERS_ARRAY(shel_retro_lot_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_lot_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_lot_verif, each_member) <> "Select one" Then
                If ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_retro_lot_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, each_member) Then
                        full_shelter_details = full_shelter_details & "Lot Rent $" & ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member)) - 4) & ". "
                    Else
                        full_shelter_details = full_shelter_details & "change in Lot Rent retro - $" & ALL_MEMBERS_ARRAY(shel_retro_lot_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_lot_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_lot_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member)) - 4) & ". "
                    End If
                    total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, each_member)
                Else
                    full_shelter_details = full_shelter_details & "Lot Rent (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_lot_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_lot_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_lot_verif, each_member)) - 4) & ". "
                End If
            ElseIf ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member) <> "Select one" Then
                full_shelter_details = full_shelter_details & "Lot Rent (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member)) - 4) & ". "
                total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, each_member)
            End If

            If ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, each_member) <> "Select one" Then
                If ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, each_member) Then
                        full_shelter_details = full_shelter_details & "Mortgage $" & ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member)) - 4) & ". "
                    Else
                        full_shelter_details = full_shelter_details & "change in Mortgage retro - $" & ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member)) - 4) & ". "
                    End If
                    total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, each_member)
                Else
                    full_shelter_details = full_shelter_details & "Mortgage (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, each_member)) - 4) & ". "
                End If
            ElseIf ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member) <> "Select one" Then
                full_shelter_details = full_shelter_details & "Mortgage (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member)) - 4) & ". "
                total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, each_member)
            End If

            If ALL_MEMBERS_ARRAY(shel_retro_ins_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_ins_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_ins_verif, each_member) <> "Select one" Then
                If ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_retro_ins_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, each_member) Then
                        full_shelter_details = full_shelter_details & "Home Insurance $" & ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member)) - 4) & ". "
                    Else
                        full_shelter_details = full_shelter_details & "change in Home Insurance retro - $" & ALL_MEMBERS_ARRAY(shel_retro_ins_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_ins_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_ins_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member)) - 4) & ". "
                    End If
                    total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, each_member)
                Else
                    full_shelter_details = full_shelter_details & "Home Insurance (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_ins_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_ins_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_ins_verif, each_member)) - 4) & ". "
                End If
            ElseIf ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member) <> "Select one" Then
                full_shelter_details = full_shelter_details & "Home Insurance (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member)) - 4) & ". "
                total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, each_member)
            End If

            If ALL_MEMBERS_ARRAY(shel_retro_tax_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_tax_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_tax_verif, each_member) <> "Select one" Then
                If ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_retro_tax_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, each_member) Then
                        full_shelter_details = full_shelter_details & "Property Tax $" & ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member)) - 4) & ". "
                    Else
                        full_shelter_details = full_shelter_details & "change in Property Tax retro - $" & ALL_MEMBERS_ARRAY(shel_retro_tax_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_tax_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_tax_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member)) - 4) & ". "
                    End If
                    total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, each_member)
                Else
                    full_shelter_details = full_shelter_details & "Property Tax (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_tax_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_tax_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_tax_verif, each_member)) - 4) & ". "
                End If
            ElseIf ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member) <> "Select one" Then
                full_shelter_details = full_shelter_details & "Property Tax (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member)) - 4) & ". "
                total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, each_member)
            End If

            If ALL_MEMBERS_ARRAY(shel_retro_room_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_room_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_room_verif, each_member) <> "Select one" Then
                If ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_retro_room_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_room_amt, each_member) Then
                        full_shelter_details = full_shelter_details & "Room $" & ALL_MEMBERS_ARRAY(shel_prosp_room_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member)) - 4) & ". "
                    Else
                        full_shelter_details = full_shelter_details & "change in Room retro - $" & ALL_MEMBERS_ARRAY(shel_retro_room_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_room_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_room_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_room_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member)) - 4) & ". "
                    End If
                    total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_room_amt, each_member)
                Else
                    full_shelter_details = full_shelter_details & "Room (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_room_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_room_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_room_verif, each_member)) - 4) & ". "
                End If
            ElseIf ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member) <> "Select one" Then
                full_shelter_details = full_shelter_details & "Room (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_room_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member)) - 4) & ". "
                total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_room_amt, each_member)
            End If

            If ALL_MEMBERS_ARRAY(shel_retro_garage_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_garage_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_garage_verif, each_member) <> "Select one" Then
                If ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_retro_garage_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, each_member) Then
                        full_shelter_details = full_shelter_details & "Garage $" & ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member)) - 4) & ". "
                    Else
                        full_shelter_details = full_shelter_details & "change in Garage retro - $" & ALL_MEMBERS_ARRAY(shel_retro_garage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_garage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_garage_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member)) - 4) & ". "
                    End If
                    total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, each_member)
                Else
                    full_shelter_details = full_shelter_details & "Garage (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_garage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_garage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_garage_verif, each_member)) - 4) & ". "
                End If
            ElseIf ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member) <> "Select one" Then
                full_shelter_details = full_shelter_details & "Garage (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member)) - 4) & ". "
                total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, each_member)
            End If

            If ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, each_member) <> "Select one" Then
                If ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, each_member) Then
                        full_shelter_details = full_shelter_details & "Subsidy $" & ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member)) - 4) & ". "
                    Else
                        full_shelter_details = full_shelter_details & "change in Subsidy retro - $" & ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member)) - 4) & ". "
                    End If
                    total_shelter_amount = total_shelter_amount - ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, each_member)
                Else
                    full_shelter_details = full_shelter_details & "Subsidy (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, each_member)) - 4) & ". "
                End If
            ElseIf ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member) <> "Select one" Then
                full_shelter_details = full_shelter_details & "Subsidy (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member)) - 4) & ". "
                total_shelter_amount = total_shelter_amount - ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, each_member)
            End If
            If ALL_MEMBERS_ARRAY(shel_shared, each_member) = "Yes" Then full_shelter_details = full_shelter_details & "Shelter expense is SHARED. "
            If ALL_MEMBERS_ARRAY(shel_subsudized, each_member) = "Yes" Then full_shelter_details = full_shelter_details & "Shelter expense is subsidized. "
        End If
    Next

    total_shelter_amount = FormatCurrency(total_shelter_amount)

    ' MsgBOx "Length of full_shelter_details is " & len(full_shelter_details)
    if left(full_shelter_details, 2) = "; " Then full_shelter_details = right(full_shelter_details, len(full_shelter_details) - 2)
    If len(full_shelter_details) > 85 Then
        shelter_details = left(full_shelter_details, 85)
        shelter_details_two = right(full_shelter_details, len(full_shelter_details) - 85)
    Else
        shelter_details = full_shelter_details
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

        If prosp_heat_air = "Y" Then
            hest_information = "AC/Heat - Full $493"
        ElseIf prosp_electric = "Y" Then
            If prosp_phone = "Y" Then
                hest_information = "Electric and Phone - $173"
            Else
                hest_information = "Electric ONLY - $126"
            End If
        ElseIf prosp_phone = "Y" Then
            hest_information = "Phone ONLY - $47"
        End If
    END IF
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
    If living_sit_line = "01" Then living_situation = "01 - Own home, lease or roomate"
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

    addr_eff_date = replace(addr_eff_date, "' ", "/")
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

function read_JOBS_panel()
'--- This function adds STAT/JOBS data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ JOBS_variable: the variable used by the editbox you wish to autofill.
'===== Keywords: MAXIS, autofill, JOBS
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

    If income_type = "01" Then ALL_BUSI_PANELS_ARRAY(busi_type, job_count) = "01 - Farming"
    If income_type = "02" Then ALL_BUSI_PANELS_ARRAY(busi_type, job_count) = "02 - Real Estate"
    If income_type = "03" Then ALL_BUSI_PANELS_ARRAY(busi_type, job_count) = "03 - Home Product Sales"
    If income_type = "04" Then ALL_BUSI_PANELS_ARRAY(busi_type, job_count) = "04 - Other Sales"
    If income_type = "05" Then ALL_BUSI_PANELS_ARRAY(busi_type, job_count) = "05 - Personal Services"
    If income_type = "06" Then ALL_BUSI_PANELS_ARRAY(busi_type, job_count) = "06 - Paper Route"
    If income_type = "07" Then ALL_BUSI_PANELS_ARRAY(busi_type, job_count) = "07 - In Home Daycare"
    If income_type = "08" Then ALL_BUSI_PANELS_ARRAY(busi_type, job_count) = "08 - Rental Income"
    If income_type = "09" Then ALL_BUSI_PANELS_ARRAY(busi_type, job_count) = "09 - Other"
    ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, job_count) = trim(retro_rpt_hrs)
    ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, job_count) = trim(prosp_rpt_hrs)
    ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, job_count) = trim(retro_min_wg_hrs)
    ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, job_count) = trim(prosp_min_wg_hrs)
    If self_emp_method = "01" Then ALL_BUSI_PANELS_ARRAY(calc_method, job_count) = "50% Gross Inc"
    If self_emp_method = "02" Then ALL_BUSI_PANELS_ARRAY(calc_method, job_count) = "Tax Forms"
    ALL_BUSI_PANELS_ARRAY(mthd_date, job_count) = replace(method_date, " ", "/")
    If ALL_BUSI_PANELS_ARRAY(mthd_date, job_count) = "__/__/__" Then ALL_BUSI_PANELS_ARRAY(mthd_date, job_count) = ""
    ALL_BUSI_PANELS_ARRAY(start_date, job_count) = replace(income_start, " ", "/")
    ALL_BUSI_PANELS_ARRAY(end_date, job_count) = replace(income_end, " ", "/")
    If ALL_BUSI_PANELS_ARRAY(end_date, job_count) = "__/__/__" Then ALL_BUSI_PANELS_ARRAY(end_date, job_count) = ""

    EMWriteScreen "X", 6, 26
    transmit

    EMReadScreen retro_cash_inc, 8, 9, 43
    EMReadScreen prosp_cash_inc, 8, 9, 59
    EMReadScreen cash_inc_verif, 1, 9, 73
    EMReadScreen retro_cash_exp, 8, 15, 43
    EMReadScreen prosp_cash_exp, 8, 15, 59
    EMReadScreen cash_exp_verif, 1, 15, 73
    ALL_BUSI_PANELS_ARRAY(income_ret_cash, job_count) = trim(retro_cash_inc)
    If ALL_BUSI_PANELS_ARRAY(income_ret_cash, job_count) = "________" Then ALL_BUSI_PANELS_ARRAY(income_ret_cash, job_count) = "0"
    ALL_BUSI_PANELS_ARRAY(income_pro_cash, job_count) = trim(prosp_cash_inc)
    If ALL_BUSI_PANELS_ARRAY(income_pro_cash, job_count) = "________" Then ALL_BUSI_PANELS_ARRAY(income_pro_cash, job_count) = "0"
    If cash_inc_verif = "1" Then ALL_BUSI_PANELS_ARRAY(cash_income_verif, job_count) = "Income Tax Returns"
    If cash_inc_verif = "2" Then ALL_BUSI_PANELS_ARRAY(cash_income_verif, job_count) = "Receipts of Sales/Purch"
    If cash_inc_verif = "3" Then ALL_BUSI_PANELS_ARRAY(cash_income_verif, job_count) = "Busi Records/Ledger"
    If cash_inc_verif = "6" Then ALL_BUSI_PANELS_ARRAY(cash_income_verif, job_count) = "Other Document"
    If cash_inc_verif = "N" Then ALL_BUSI_PANELS_ARRAY(cash_income_verif, job_count) = "No Verif Provided"
    If cash_inc_verif = "?" Then ALL_BUSI_PANELS_ARRAY(cash_income_verif, job_count) = "Delayed Verif"
    If cash_inc_verif = "_" Then ALL_BUSI_PANELS_ARRAY(cash_income_verif, job_count) = "Blank"
    ALL_BUSI_PANELS_ARRAY(expense_ret_cash, job_count) = trim(retro_cash_exp)
    If ALL_BUSI_PANELS_ARRAY(expense_ret_cash, job_count) = "________" Then ALL_BUSI_PANELS_ARRAY(expense_ret_cash, job_count) = "0"
    ALL_BUSI_PANELS_ARRAY(expense_pro_cash, job_count) = trim(prosp_cash_exp)
    If ALL_BUSI_PANELS_ARRAY(expense_pro_cash, job_count) = "________" Then ALL_BUSI_PANELS_ARRAY(expense_pro_cash, job_count) = "0"
    If cash_exp_verif = "1" Then ALL_BUSI_PANELS_ARRAY(cash_expense_verif, job_count) = "Income Tax Returns"
    If cash_exp_verif = "2" Then ALL_BUSI_PANELS_ARRAY(cash_expense_verif, job_count) = "Receipts of Sales/Purch"
    If cash_exp_verif = "3" Then ALL_BUSI_PANELS_ARRAY(cash_expense_verif, job_count) = "Busi Records/Ledger"
    If cash_exp_verif = "6" Then ALL_BUSI_PANELS_ARRAY(cash_expense_verif, job_count) = "Other Document"
    If cash_exp_verif = "N" Then ALL_BUSI_PANELS_ARRAY(cash_expense_verif, job_count) = "No Verif Provided"
    If cash_exp_verif = "?" Then ALL_BUSI_PANELS_ARRAY(cash_expense_verif, job_count) = "Delayed Verif"
    If cash_exp_verif = "_" Then ALL_BUSI_PANELS_ARRAY(cash_expense_verif, job_count) = "Blank"

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
    ALL_BUSI_PANELS_ARRAY(income_ret_snap, job_count) = trim(retro_snap_inc)
    If ALL_BUSI_PANELS_ARRAY(income_ret_snap, job_count) = "________" Then ALL_BUSI_PANELS_ARRAY(income_ret_snap, job_count) = "0"
    ALL_BUSI_PANELS_ARRAY(income_pro_snap, job_count) = trim(prosp_snap_inc)
    If ALL_BUSI_PANELS_ARRAY(income_pro_snap, job_count) = "________" Then ALL_BUSI_PANELS_ARRAY(income_pro_snap, job_count) = "0"
    If snap_inc_verif = "1" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, job_count) = "Income Tax Returns"
    If snap_inc_verif = "2" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, job_count) = "Receipts of Sales/Purch"
    If snap_inc_verif = "3" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, job_count) = "Busi Records/Ledger"
    If snap_inc_verif = "4" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, job_count) = "Pend Out State Verif"
    If snap_inc_verif = "6" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, job_count) = "Other Document"
    If snap_inc_verif = "N" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, job_count) = "No Verif Provided"
    If snap_inc_verif = "?" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, job_count) = "Delayed Verif"
    If snap_inc_verif = "_" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, job_count) = "Blank"

    ALL_BUSI_PANELS_ARRAY(expense_ret_snap, job_count) = trim(retro_snap_exp)
    If ALL_BUSI_PANELS_ARRAY(expense_ret_snap, job_count) = "________" Then ALL_BUSI_PANELS_ARRAY(expense_ret_snap, job_count) = "0"
    ALL_BUSI_PANELS_ARRAY(expense_pro_snap, job_count) = trim(prosp_snap_exp)
    If ALL_BUSI_PANELS_ARRAY(expense_pro_snap, job_count) = "________" Then ALL_BUSI_PANELS_ARRAY(expense_pro_snap, job_count) = "0"
    If snap_exp_verif = "1" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, job_count) = "Income Tax Returns"
    If snap_exp_verif = "2" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, job_count) = "Receipts of Sales/Purch"
    If snap_exp_verif = "3" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, job_count) = "Busi Records/Ledger"
    If snap_exp_verif = "4" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, job_count) = "Pend Out State Verif"
    If snap_exp_verif = "6" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, job_count) = "Other Document"
    If snap_exp_verif = "N" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, job_count) = "No Verif Provided"
    If snap_exp_verif = "?" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, job_count) = "Delayed Verif"
    If snap_exp_verif = "_" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, job_count) = "Blank"

    EMReadScreen prosp_hca_inc, 8, 12, 59
    EMReadScreen hca_inc_verif, 1, 12, 73
    EMReadScreen prosp_hca_exp, 8, 18, 59
    EMReadScreen hca_exp_verif, 1, 18, 73

    EMReadScreen prosp_hcb_inc, 8, 13, 59
    EMReadScreen hcb_inc_verif, 1, 13, 73
    EMReadScreen prosp_hcb_exp, 8, 19, 59
    EMReadScreen hcb_exp_verif, 1, 19, 73

    PF3

    EMReadScreen
    EMReadScreen
end function

'===========================================================================================================================

'JOBS ARRAY CONSTANTS
const memb_numb     = 0
const panel_instance  = 1
const employer_name = 2
const busi_type     = 2         'for BUSI Array
Const estimate_only = 3
const verif_explain = 4
const verif_code    = 5
const calc_method   = 5         'for BUSI Array
const info_month    = 6
const hrly_wage     = 7
const mthd_date     = 7          'for BUSI Array'
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
const start_date            = 24
const end_date              = 25
const budget_explain        = 26

Dim ALL_JOBS_PANELS_ARRAY()
ReDim ALL_JOBS_PANELS_ARRAY(budget_explain, 0)


Dim ALL_BUSI_PANELS_ARRAY()
ReDim ALL_BUSI_PANELS_ARRAY(budget_explain, 0)

const clt_name                  = 1
const clt_age                   = 2
const full_clt                  = 3
const include_cash_checkbox     = 4
const include_snap_checkbox     = 5
const include_emer_checkbox     = 6
const count_cash_checkbox       = 7
const count_snap_checkbox       = 8
const count_emer_checkbox       = 9
const clt_wreg_status           = 10
const clt_abawd_status          = 11
const pwe_checkbox              = 12
const numb_abawd_used           = 13
const list_abawd_mo             = 14
const first_second_set          = 15
const list_second_set           = 16
const explain_no_second         = 17
const numb_banked_mo            = 18
const clt_abawd_notes           = 19
const shel_exists               = 20
const shel_subsudized           = 21
const shel_shared               = 22
const shel_retro_rent_amt       = 23
const shel_retro_rent_verif     = 24
const shel_prosp_rent_amt       = 25
const shel_prosp_rent_verif     = 26
const shel_retro_lot_amt        = 27
const shel_retro_lot_verif      = 28
const shel_prosp_lot_amt        = 29
const shel_prosp_lot_verif      = 30
const shel_retro_mortgage_amt   = 31
const shel_retro_mortgage_verif = 32
const shel_prosp_mortgage_amt   = 33
const shel_prosp_mortgage_verif = 34
const shel_retro_ins_amt        = 35
const shel_retro_ins_verif      = 36
const shel_prosp_ins_amt        = 37
const shel_prosp_ins_verif      = 38
const shel_retro_tax_amt        = 39
const shel_retro_tax_verif      = 40
const shel_prosp_tax_amt        = 41
const shel_prosp_tax_verif      = 42
const shel_retro_room_amt       = 43
const shel_retro_room_verif     = 44
const shel_prosp_room_amt       = 45
const shel_prosp_room_verif     = 46
const shel_retro_garage_amt     = 47
const shel_retro_garage_verif   = 48
const shel_prosp_garage_amt     = 49
const shel_prosp_garage_verif   = 50
const shel_retro_subsidy_amt    = 51
const shel_retro_subsidy_verif  = 52
const shel_prosp_subsidy_amt    = 53
const shel_prosp_subsidy_verif  = 54
const clt_notes                 = 55

Dim ALL_MEMBERS_ARRAY()
ReDim ALL_MEMBERS_ARRAY(clt_notes, 0)

EATS = ""

county_list = "01 Aitkin"
county_list = county_list+chr(9)+"02 Anoka"
county_list = county_list+chr(9)+"03 Becker"
county_list = county_list+chr(9)+"04 Beltrami"
county_list = county_list+chr(9)+"05 Benton"
county_list = county_list+chr(9)+"06 Big Stone"
county_list = county_list+chr(9)+"07 Blue Earth"
county_list = county_list+chr(9)+"08 Brown"
county_list = county_list+chr(9)+"09 Carlton"
county_list = county_list+chr(9)+"10 Carver"
county_list = county_list+chr(9)+"11 Cass"
county_list = county_list+chr(9)+"12 Chippewa"
county_list = county_list+chr(9)+"13 Chisago"
county_list = county_list+chr(9)+"14 Clay"
county_list = county_list+chr(9)+"15 Clearwater"
county_list = county_list+chr(9)+"16 Cook"
county_list = county_list+chr(9)+"17 Cottonwood"
county_list = county_list+chr(9)+"18 Crow Wing"
county_list = county_list+chr(9)+"19 Dakota"
county_list = county_list+chr(9)+"20 Dodge"
county_list = county_list+chr(9)+"21 Douglas"
county_list = county_list+chr(9)+"22 Faribault"
county_list = county_list+chr(9)+"23 Fillmore"
county_list = county_list+chr(9)+"24 Freeborn"
county_list = county_list+chr(9)+"25 Goodhue"
county_list = county_list+chr(9)+"26 Grant"
county_list = county_list+chr(9)+"27 Hennepin"
county_list = county_list+chr(9)+"28 Houston"
county_list = county_list+chr(9)+"29 Hubbard"
county_list = county_list+chr(9)+"30 Isanti"
county_list = county_list+chr(9)+"31 Itasca"
county_list = county_list+chr(9)+"32 Jackson"
county_list = county_list+chr(9)+"33 Kanabec"
county_list = county_list+chr(9)+"34 Kandiyohi"
county_list = county_list+chr(9)+"35 Kittson"
county_list = county_list+chr(9)+"36 Koochiching"
county_list = county_list+chr(9)+"37 Lac Qui Parle"
county_list = county_list+chr(9)+"38 Lake"
county_list = county_list+chr(9)+"39 Lake Of Woods"
county_list = county_list+chr(9)+"40 Le Sueur"
county_list = county_list+chr(9)+"41 Lincoln"
county_list = county_list+chr(9)+"42 Lyon"
county_list = county_list+chr(9)+"43 Mcleod"
county_list = county_list+chr(9)+"44 Mahnomen"
county_list = county_list+chr(9)+"45 Marshall"
county_list = county_list+chr(9)+"46 Martin"
county_list = county_list+chr(9)+"47 Meeker"
county_list = county_list+chr(9)+"48 Mille Lacs"
county_list = county_list+chr(9)+"49 Morrison"
county_list = county_list+chr(9)+"50 Mower"
county_list = county_list+chr(9)+"51 Murray"
county_list = county_list+chr(9)+"52 Nicollet"
county_list = county_list+chr(9)+"53 Nobles"
county_list = county_list+chr(9)+"54 Norman"
county_list = county_list+chr(9)+"55 Olmsted"
county_list = county_list+chr(9)+"56 Otter Tail"
county_list = county_list+chr(9)+"57 Pennington"
county_list = county_list+chr(9)+"58 Pine"
county_list = county_list+chr(9)+"59 Pipestone"
county_list = county_list+chr(9)+"60 Polk"
county_list = county_list+chr(9)+"61 Pope"
county_list = county_list+chr(9)+"62 Ramsey"
county_list = county_list+chr(9)+"63 Red Lake"
county_list = county_list+chr(9)+"64 Redwood"
county_list = county_list+chr(9)+"65 Renville"
county_list = county_list+chr(9)+"66 Rice"
county_list = county_list+chr(9)+"67 Rock"
county_list = county_list+chr(9)+"68 Roseau"
county_list = county_list+chr(9)+"69 St. Louis"
county_list = county_list+chr(9)+"70 Scott"
county_list = county_list+chr(9)+"71 Sherburne"
county_list = county_list+chr(9)+"72 Sibley"
county_list = county_list+chr(9)+"73 Stearns"
county_list = county_list+chr(9)+"74 Steele"
county_list = county_list+chr(9)+"75 Stevens"
county_list = county_list+chr(9)+"76 Swift"
county_list = county_list+chr(9)+"77 Todd"
county_list = county_list+chr(9)+"78 Traverse"
county_list = county_list+chr(9)+"79 Wabasha"
county_list = county_list+chr(9)+"80 Wadena"
county_list = county_list+chr(9)+"81 Waseca"
county_list = county_list+chr(9)+"82 Washington"
county_list = county_list+chr(9)+"83 Watonwan"
county_list = county_list+chr(9)+"84 Wilkin"
county_list = county_list+chr(9)+"85 Winona"
county_list = county_list+chr(9)+"86 Wright"
county_list = county_list+chr(9)+"87 Yellow Medicine"
county_list = county_list+chr(9)+"89 Out-of-State"

'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
HH_memb_row = 5 'This helps the navigation buttons work!
Dim row
Dim col
application_signed_checkbox = checked 'The script should default to having the application signed.

'GRABBING THE CASE NUMBER, THE MEMB NUMBERS, AND THE FOOTER MONTH------------------------------------------------------------------------------------------------------------------------------------------------
EMConnect ""
get_county_code				'since there is a county specific checkbox, this makes the the county clear
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

BeginDialog Dialog1, 0, 0, 281, 215, "Case number dialog"
  EditBox 65, 50, 60, 15, MAXIS_case_number
  EditBox 210, 50, 15, 15, MAXIS_footer_month
  EditBox 230, 50, 15, 15, MAXIS_footer_year
  CheckBox 10, 85, 30, 10, "CASH", CASH_on_CAF_checkbox
  CheckBox 50, 85, 35, 10, "SNAP", SNAP_on_CAF_checkbox
  CheckBox 90, 85, 35, 10, "EMER", EMER_on_CAF_checkbox
  DropListBox 185, 80, 75, 15, "Select One:"+chr(9)+"Application"+chr(9)+"Recertification"+chr(9)+"Addendum", CAF_type
  EditBox 40, 130, 220, 15, cash_other_req_detail
  EditBox 40, 150, 220, 15, snap_other_req_detail
  EditBox 40, 170, 220, 15, emer_other_req_detail
  ButtonGroup ButtonPressed
    OkButton 170, 195, 50, 15
    CancelButton 225, 195, 50, 15
    PushButton 10, 30, 105, 10, "NOTES - Interview Completed", interview_completed_button
  Text 10, 10, 265, 20, "This script works best when run AFTER all STAT panels have been updated. If STAT panels have not been updated but you need to case note the interview use "
  Text 10, 55, 50, 10, "Case number:"
  Text 140, 55, 65, 10, "Footer month/year: "
  GroupBox 5, 70, 125, 30, "Programs marked on CAF"
  Text 145, 85, 35, 10, "CAF type:"
  GroupBox 5, 105, 265, 85, "OTHER Program Requests (not marked on CAF)"
  Text 40, 120, 130, 10, "Explain how the program was reuested."
  Text 15, 135, 20, 10, "Cash:"
  Text 15, 155, 20, 10, "SNAP:"
  Text 15, 175, 25, 10, "EMER:"
EndDialog

'initial dialog
Do
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation

		If buttonpressed = interview_completed_button Then
            confirm_run_another_script = MsgBox("You have selected the 'NOTES - Interview completed' option. This will stop the NOTES - CAF script and run the script NOTES - Interview Completed." & vbNewLine & vbNewLine &_
                                                "This option is best for when the STAT panels have not been updated when running the script. We recommend runing NOTES - CAF once STAT panels are updated to capture the correct case information in CASE/NOTE." & vbNewLine & vbNewLine &_
                                                "Would you like to continue to NOTES - Interview Completed?", vbQuestion + vbYesNo, "Stop CAF Script?")
            If confirm_run_another_script = vbYes Then Call run_from_GitHub(script_repository & "notes/interview-completed.vbs")
            If confirm_run_another_script = vbNo Then err_msg = "LOOP" & err_msg
        End If

        If CAF_type = "Select One:" then err_msg = err_msg & vbnewline & "* You must select the CAF type."
        Call validate_MAXIS_case_number(err_msg, "*")
        If IsNumeric(MAXIS_footer_month) = FALSE OR len(MAXIS_footer_month) > 2 Then err_msg = err_msg & vbNewLine & "* Enter a valid Footer Month."
        If IsNumeric(MAXIS_footer_year) = FALSE OR len(MAXIS_footer_year) > 2 Then err_msg = err_msg & vbNewLine & "* Enter a valid Footer Year."
        If CASH_on_CAF_checkbox = unchecked AND SNAP_on_CAF_checkbox = unchecked AND EMER_on_CAF_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "* At least one program should be marked on the CAF."
        If CASH_on_CAF_checkbox = checked AND trim(cash_other_req_detail) <> "" Then err_msg = err_msg & vbNewLine & "* If CASH was marked on the CAF, then another way of requesting does not need to be indicated."
        If SNAP_on_CAF_checkbox = checked AND trim(snap_other_req_detail) <> "" Then err_msg = err_msg & vbNewLine & "* If CASH was marked on the CAF, then another way of requesting does not need to be indicated."
        If EMER_on_CAF_checkbox = checked AND trim(emer_other_req_detail) <> "" Then err_msg = err_msg & vbNewLine & "* If CASH was marked on the CAF, then another way of requesting does not need to be indicated."

        IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

If CASH_on_CAF_checkbox = checked or trim(cash_other_req_detail) <> "" Then cash_checkbox = checked
If SNAP_on_CAF_checkbox = checked or trim(snap_other_req_detail) <> "" Then SNAP_checkbox = checked
If EMER_on_CAF_checkbox = checked or trim(emer_other_req_detail) <> "" Then EMER_checkbox = checked
'grh_checkbox = checked

MAXIS_footer_month = right("00" & MAXIS_footer_month, 2)
MAXIS_footer_year = right("00" & MAXIS_footer_year, 2)
call check_for_MAXIS(False)	'checking for an active MAXIS session
MAXIS_footer_month_confirmation	'function will check the MAXIS panel footer month/year vs. the footer month/year in the dialog, and will navigate to the dialog month/year if they do not match.

'GRABBING THE DATE RECEIVED AND THE HH MEMBERS---------------------------------------------------------------------------------------------------------------------------------------------------------------------
call navigate_to_MAXIS_screen("stat", "hcre")
EMReadScreen STAT_check, 4, 20, 21
If STAT_check <> "STAT" then script_end_procedure("Can't get in to STAT. This case may be in background. Wait a few seconds and try again. If the case is not in background contact an alpha user for your agency.")

'Creating a custom dialog for determining who the HH members are
call HH_comp_dialog(HH_member_array)

'GRABBING THE INFO FOR THE CASE NOTE-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
If CAF_type = "Recertification" then                                                          'For recerts it goes to one area for the CAF datestamp. For other app types it goes to STAT/PROG.
	call autofill_editbox_from_MAXIS(HH_member_array, "REVW", CAF_datestamp)
	IF SNAP_checkbox = checked THEN																															'checking for SNAP 24 month renewals.'
		EMWriteScreen "X", 05, 58																																	'opening the FS revw screen.
		transmit
		EMReadScreen SNAP_recert_date, 8, 9, 64
		PF3
		SNAP_recert_date = replace(SNAP_recert_date, " ", "/")																		'replacing the read blank spaces with / to make it a date
		SNAP_recert_compare_date = dateadd("m", "12", MAXIS_footer_month & "/01/" & MAXIS_footer_year)		'making a dummy variable to compare with, by adding 12 months to the requested footer month/year.
		IF datediff("d", SNAP_recert_compare_date, SNAP_recert_date) > 0 THEN											'If the read recert date is more than 0 days away from 12 months plus the MAXIS footer month/year then it is likely a 24 month period.'
			SNAP_recert_is_likely_24_months = TRUE
		ELSE
			SNAP_recert_is_likely_24_months = FALSE																									'otherwise if we don't we set it as false
		END IF
	END IF
Else
	call autofill_editbox_from_MAXIS(HH_member_array, "PROG", CAF_datestamp)
End if
If HC_checkbox = checked and CAF_type <> "Recertification" then call autofill_editbox_from_MAXIS(HH_member_array, "HCRE-retro", retro_request)     'Grabbing retro info for HC cases that aren't recertifying
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)                                                                        'Grabbing HH comp info from MEMB.
If SNAP_checkbox = checked then call autofill_editbox_from_MAXIS(HH_member_array, "EATS", HH_comp)                                                 'Grabbing EATS info for SNAP cases, puts on HH_comp variable
'Removing semicolons from HH_comp variable, it is not needed.
HH_comp = replace(HH_comp, "; ", "")

'I put these sections in here, just because SHEL should come before HEST, it just looks cleaner.
total_shelter_amount = ""
full_shelter_details = ""
shelter_details = ""
shelter_details_two = ""
hest_information = ""

addr_line_one = ""
addr_line_two = ""
city = ""
state = ""
zip = ""
address_confirmation_checkbox = ""
addr_county = ""
homeless_yn = ""
addr_verif = ""
reservation_yn = ""
living_situation = ""
notes_on_address = ""

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
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "CASH", notes_on_cash)
call autofill_editbox_from_MAXIS(HH_member_array, "CARS", notes_on_cars)
call autofill_editbox_from_MAXIS(HH_member_array, "COEX", notes_on_coex)
call autofill_editbox_from_MAXIS(HH_member_array, "DCEX", notes_on_dcex)
call autofill_editbox_from_MAXIS(HH_member_array, "DIET", DIET)
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
Else

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
                ALL_BUSI_PANELS_ARRAY(memb_numb, job_count) = HH_member
                ALL_BUSI_PANELS_ARRAY(info_month, busi_count) = MAXIS_footer_month & "/" & MAXIS_footer_year
                call read_BUSI_panel

                EMReadScreen BUSI_panel_current, 1, 2, 73
                ALL_JOBS_PANELS_ARRAY(panel_instance, job_count) = "0" & BUSI_panel_current

                If cint(BUSI_panel_current) < cint(BUSI_total) then transmit
                job_count = job_count + 1
            Loop until cint(BUSI_panel_current) = cint(BUSI_total)

        End if
    Next
Else

End If

'FOR EACH JOB PANEL GO LOOK FOR A RECENT EI CASE NOTE'
call autofill_editbox_from_MAXIS(HH_member_array, "MEDI", INSA)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMI", cit_id)
call autofill_editbox_from_MAXIS(HH_member_array, "OTHR", other_assets)
call autofill_editbox_from_MAXIS(HH_member_array, "PBEN", income_changes)
call autofill_editbox_from_MAXIS(HH_member_array, "PREG", PREG)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "REST", notes_on_rest)
call autofill_editbox_from_MAXIS(HH_member_array, "SCHL", SCHL)
call autofill_editbox_from_MAXIS(HH_member_array, "SECU", other_assets)
call autofill_editbox_from_MAXIS(HH_member_array, "STWK", income_changes)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)
notes_on_wreg = ""
full_abawd_info = ""
call read_WREG_panel
call update_wreg_and_abawd_notes
'call autofill_editbox_from_MAXIS(HH_member_array, "WREG", notes_on_abawd)

'MAKING THE GATHERED INFORMATION LOOK BETTER FOR THE CASE NOTE
If cash_checkbox = checked then programs_applied_for = programs_applied_for & "CASH, "
If HC_checkbox = checked then programs_applied_for = programs_applied_for & "HC, "
If SNAP_checkbox = checked then programs_applied_for = programs_applied_for & "SNAP, "
If EMER_checkbox = checked then programs_applied_for = programs_applied_for & "emergency, "
programs_applied_for = trim(programs_applied_for)
if right(programs_applied_for, 1) = "," then programs_applied_for = left(programs_applied_for, len(programs_applied_for) - 1)

'SHOULD DEFAULT TO TIKLING FOR APPLICATIONS THAT AREN'T RECERTS.
If CAF_type <> "Recertification" then TIKL_checkbox = checked

Call generage_client_list(inteview_memb_list, "Select or Type")
Call generage_client_list(shel_memb_list, "Select")

'CASE NOTE DIALOG--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Do
	Do
		Do
			Do

                BeginDialog Dialog1, 0, 0, 451, 285, "CAF Dialog 1 - Personal Information"
                  EditBox 60, 5, 50, 15, CAF_datestamp
                  ComboBox 175, 5, 70, 15, "phone"+chr(9)+"office", interview_type
                  CheckBox 255, 5, 65, 10, "Used Interpreter", Used_Interpreter_checkbox
                  EditBox 60, 25, 50, 15, interview_date
                  ComboBox 230, 25, 95, 15, "Select One:"+chr(9)+"Fax"+chr(9)+"Mail"+chr(9)+"Office"+chr(9)+"Online", how_app_rcvd
                  ComboBox 90, 45, 150, 45, inteview_memb_list, interview_with
                  EditBox 35, 65, 410, 15, cit_id
                  EditBox 35, 85, 410, 15, IMIG
                  EditBox 60, 105, 120, 15, AREP
                  EditBox 270, 105, 175, 15, SCHL
                  EditBox 60, 125, 210, 15, DISA
                  EditBox 310, 125, 135, 15, FACI
                  EditBox 35, 155, 410, 15, PREG
                  EditBox 35, 175, 410, 15, ABPS
                  'Add editbox for date GC Form Sent to clt - check with Melissa Flores
                  EditBox 35, 195, 410, 15, EMPS
                  CheckBox 35, 215, 180, 10, "Sent MFIP financial orientation DVD to participant(s).", MFIP_DVD_checkbox
                  EditBox 55, 230, 390, 15, verifs_needed
                  ButtonGroup ButtonPressed
                    PushButton 340, 265, 50, 15, "NEXT", next_to_page_02_button
                    CancelButton 395, 265, 50, 15
                    PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
                    PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
                    PushButton 335, 25, 45, 10, "next panel", next_panel_button
                    PushButton 395, 25, 45, 10, "next memb", next_memb_button
                    PushButton 5, 90, 20, 10, "IMIG:", IMIG_button
                    PushButton 5, 110, 25, 10, "AREP/", AREP_button
                    PushButton 30, 110, 25, 10, "ALTP:", ALTP_button
                    PushButton 190, 110, 25, 10, "SCHL/", SCHL_button
                    PushButton 215, 110, 25, 10, "STIN/", STIN_button
                    PushButton 240, 110, 25, 10, "STEC:", STEC_button
                    PushButton 5, 130, 25, 10, "DISA/", DISA_button
                    PushButton 30, 130, 25, 10, "PDED:", PDED_button
                    PushButton 280, 130, 25, 10, "FACI:", FACI_button
                    PushButton 5, 160, 25, 10, "PREG:", PREG_button
                    PushButton 5, 180, 25, 10, "ABPS:", ABPS_button
                    PushButton 5, 200, 25, 10, "EMPS", EMPS_button
                    PushButton 10, 265, 20, 10, "DWP", ELIG_DWP_button
                    PushButton 30, 265, 15, 10, "FS", ELIG_FS_button
                    PushButton 45, 265, 15, 10, "GA", ELIG_GA_button
                    PushButton 60, 265, 15, 10, "HC", ELIG_HC_button
                    PushButton 75, 265, 20, 10, "MFIP", ELIG_MFIP_button
                    PushButton 95, 265, 20, 10, "MSA", ELIG_MSA_button
                    PushButton 130, 265, 25, 10, "ADDR", ADDR_button
                    PushButton 155, 265, 25, 10, "MEMB", MEMB_button
                    PushButton 180, 265, 25, 10, "MEMI", MEMI_button
                    PushButton 205, 265, 25, 10, "PROG", PROG_button
                    PushButton 230, 265, 25, 10, "REVW", REVW_button
                    PushButton 255, 265, 25, 10, "SANC", SANC_button
                    PushButton 280, 265, 25, 10, "TIME", TIME_button
                    PushButton 305, 265, 25, 10, "TYPE", TYPE_button
                  Text 5, 70, 25, 10, "CIT/ID:"
                  Text 5, 235, 50, 10, "Verifs needed:"
                  GroupBox 5, 255, 115, 25, "ELIG panels:"
                  GroupBox 125, 255, 210, 25, "other STAT panels:"
                  GroupBox 330, 5, 115, 35, "STAT-based navigation"
                  Text 5, 10, 55, 10, "CAF datestamp:"
                  Text 5, 30, 55, 10, "Interview date:"
                  Text 120, 10, 50, 10, "Interview type:"
                  Text 120, 30, 110, 10, "How was application received?:"
                  Text 5, 50, 85, 10, "Interview completed with:"
                EndDialog

				err_msg = ""
				Dialog Dialog1			'Displays the first dialog
				cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
				MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
				If CAF_datestamp = "" or len(CAF_datestamp) > 10 THEN err_msg = "Please enter a valid application datestamp."
				If err_msg <> "" THEN Msgbox err_msg
			Loop until ButtonPressed = next_to_page_02_button and err_msg = ""

            ' 'THIS IS THE BASE FOR AN INCOME SPECIFIC DIALOG'
            ' BeginDialog Dialog1, 0, 0, 606, 145, "CAF Income Information"
            '   GroupBox 5, 5, 595, 115, "Earned Income"
            '   Text 15, 20, 170, 10, "Member 01 - EMPLOYER"
            '   CheckBox 375, 20, 220, 10, "Check here if this income is not verified and is only an estimate.", no_verif_estimate_only_checkbox
            '   Text 35, 40, 40, 10, "Verification:"
            '   EditBox 85, 35, 240, 15, verif_explanation
            '   Text 340, 40, 75, 10, "Footer Month: XX/XX"
            '   Text 425, 40, 175, 10, "EARNED INCOME BUDGETING CASE NOTE FOUND"
            '   Text 35, 60, 45, 10, "Hourly Wage:"
            '   EditBox 85, 55, 50, 15, hourly_wage
            '   Text 150, 60, 75, 10, "Retrospective Income:"
            '   EditBox 230, 55, 125, 15, retro_income
            '   Text 360, 60, 70, 10, "Prospective Income:"
            '   EditBox 430, 55, 165, 15, prosp_income
            '   Text 35, 80, 35, 10, "SNAP PIC:"
            '   Text 75, 80, 60, 10, "Pay Date Amount: "
            '   EditBox 135, 75, 50, 15, pay_date_amount
            '   ComboBox 195, 75, 60, 45, "Type or select"+chr(9)+"Weekly"+chr(9)+"Biweekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly", pay_date_frequency
            '   Text 285, 80, 70, 10, "Prospective Amount:"
            '   EditBox 360, 75, 60, 15, pic_prosp_income
            '   Text 440, 80, 40, 10, "Calculated:"
            '   EditBox 490, 75, 50, 15, calculated_date
            '   Text 35, 100, 55, 10, "Explain Budget:"
            '   EditBox 90, 95, 505, 15, budget_explanation
            '   ButtonGroup ButtonPressed
            '     PushButton 430, 130, 60, 10, "previous page", previous_to_page_01_button
            '     PushButton 495, 125, 50, 15, "NEXT", next_to_page_03_button
            '     CancelButton 550, 125, 50, 15
            '     PushButton 15, 130, 25, 10, "BUSI", BUSI_button
            '     PushButton 45, 130, 25, 10, "JOBS", JOBS_button
            '     PushButton 75, 130, 25, 10, "PBEN", PBEN_button
            '     PushButton 105, 130, 25, 10, "RBIC", RBIC_button
            '     PushButton 135, 130, 25, 10, "UNEA", UNEA_button
            '     PushButton 175, 130, 45, 10, "prev. panel", prev_panel_button
            '     PushButton 225, 130, 45, 10, "next panel", next_panel_button
            '     PushButton 275, 130, 45, 10, "prev. memb", prev_memb_button
            '     PushButton 325, 130, 45, 10, "next memb", next_memb_button
            ' EndDialog
            '
            ' BeginDialog Dialog1, 0, 0, 606, 245, "CAF Income Information"
            '   GroupBox 5, 5, 595, 215, "Earned Income"
            '   Text 15, 20, 170, 10, "Member 01 - EMPLOYER"
            '   CheckBox 375, 20, 220, 10, "Check here if this income is not verified and is only an estimate.", no_verif_estimate_only_checkbox
            '   Text 35, 40, 40, 10, "Verification:"
            '   EditBox 85, 35, 240, 15, verif_explanation
            '   Text 340, 40, 75, 10, "Footer Month: XX/XX"
            '   Text 425, 40, 175, 10, "EARNED INCOME BUDGETING CASE NOTE FOUND"
            '   Text 35, 60, 45, 10, "Hourly Wage:"
            '   EditBox 85, 55, 50, 15, hourly_wage
            '   Text 150, 60, 75, 10, "Retrospective Income:"
            '   EditBox 230, 55, 125, 15, retro_income
            '   Text 360, 60, 70, 10, "Prospective Income:"
            '   EditBox 430, 55, 165, 15, prosp_income
            '   Text 35, 80, 35, 10, "SNAP PIC:"
            '   Text 75, 80, 60, 10, "Pay Date Amount: "
            '   EditBox 135, 75, 50, 15, pay_date_amount
            '   ComboBox 195, 75, 60, 45, "Type or select"+chr(9)+"Weekly"+chr(9)+"Biweekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly", pay_date_frequency
            '   Text 285, 80, 70, 10, "Prospective Amount:"
            '   EditBox 360, 75, 60, 15, pic_prosp_income
            '   Text 440, 80, 40, 10, "Calculated:"
            '   EditBox 490, 75, 50, 15, calculated_date
            '   Text 35, 100, 55, 10, "Explain Budget:"
            '   EditBox 90, 95, 505, 15, budget_explanation
            '   ButtonGroup ButtonPressed
            '     PushButton 430, 230, 60, 10, "previous page", previous_to_page_01_button
            '     PushButton 495, 225, 50, 15, "NEXT", next_to_page_03_button
            '     CancelButton 550, 225, 50, 15
            '     PushButton 15, 230, 25, 10, "BUSI", BUSI_button
            '     PushButton 45, 230, 25, 10, "JOBS", JOBS_button
            '     PushButton 75, 230, 25, 10, "PBEN", PBEN_button
            '     PushButton 105, 230, 25, 10, "RBIC", RBIC_button
            '     PushButton 135, 230, 25, 10, "UNEA", UNEA_button
            '     PushButton 175, 230, 45, 10, "prev. panel", prev_panel_button
            '     PushButton 225, 230, 45, 10, "next panel", next_panel_button
            '     PushButton 275, 230, 45, 10, "prev. memb", prev_memb_button
            '     PushButton 325, 230, 45, 10, "next memb", next_memb_button
            '   Text 15, 125, 170, 10, "Member 01 - EMPLOYER"
            '   CheckBox 375, 125, 220, 10, "Check here if this income is not verified and is only an estimate.", Check2
            '   Text 35, 145, 40, 10, "Verification:"
            '   EditBox 85, 140, 240, 15, Edit9
            '   Text 340, 145, 75, 10, "Footer Month: XX/XX"
            '   Text 425, 145, 175, 10, "EARNED INCOME BUDGETING CASE NOTE FOUND"
            '   Text 35, 165, 45, 10, "Hourly Wage:"
            '   EditBox 85, 160, 50, 15, Edit10
            '   Text 150, 165, 75, 10, "Retrospective Income:"
            '   EditBox 230, 160, 125, 15, Edit11
            '   Text 360, 165, 70, 10, "Prospective Income:"
            '   EditBox 430, 160, 165, 15, Edit12
            '   Text 35, 185, 35, 10, "SNAP PIC:"
            '   Text 75, 185, 60, 10, "Pay Date Amount: "
            '   EditBox 135, 180, 50, 15, Edit13
            '   ComboBox 195, 180, 60, 45, "Type or select"+chr(9)+"Weekly"+chr(9)+"Biweekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly", Combo2
            '   Text 285, 185, 70, 10, "Prospective Amount:"
            '   EditBox 360, 180, 60, 15, Edit14
            '   Text 440, 185, 40, 10, "Calculated:"
            '   EditBox 490, 180, 50, 15, Edit15
            '   Text 35, 205, 55, 10, "Explain Budget:"
            '   EditBox 90, 200, 505, 15, Edit16
            ' EndDialog


            Do
                each_job = 0
                loop_start = 0
                last_job_reviewed = FALSE
                job_limit = 2
                Do
                    Do
                        jobs_err_msg = ""

                        dlg_len = 45
                        jobs_grp_len = 80
                        length_factor = 80
                        If snap_checkbox = checked Then length_factor = length_factor + 20
                        If grh_checkbox = checked Then length_factor = length_factor + 20
                        'NEED HANDLING FOR IF NO JOBS'
                        If UBound(ALL_JOBS_PANELS_ARRAY, 2) >= job_limit Then
                            dlg_len = 285
                            If snap_checkbox = checked Then dlg_len = dlg_len + 60
                            If grh_checkbox = checked Then dlg_len = dlg_len + 60
                            'jobs_grp_len = 315
                        Else
                            dlg_len = length_factor * (UBound(ALL_JOBS_PANELS_ARRAY, 2) - loop_start + 1) + 45
                            'jobs_grp_len = 100 * (UBound(ALL_JOBS_PANELS_ARRAY, 2) - loop_start + 1) + 15
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
                        y_pos = 10
                        'MsgBox dlg_len

                        BeginDialog Dialog1, 0, 0, 606, dlg_len, "CAF Dialog 2 - JOBS Information"
                          'GroupBox 5, 5, 595, jobs_grp_len, "Earned Income"
                          each_job = loop_start
                          Do
                              GroupBox 5, y_pos, 595, jobs_grp_len, "Member " & ALL_JOBS_PANELS_ARRAY(memb_numb, each_job) & " - " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job)
                              Text 160, y_pos, 200, 10, "Verif: " & ALL_JOBS_PANELS_ARRAY(verif_code, each_job)
                              CheckBox 365, y_pos, 220, 10, "Check here if this income is not verified and is only an estimate.", ALL_JOBS_PANELS_ARRAY(estimate_only, each_job)
                              y_pos = y_pos + 20
                              Text 15, y_pos, 40, 10, "Verification:"
                              EditBox 65, y_pos - 5, 260, 15, ALL_JOBS_PANELS_ARRAY(verif_explain, each_job)
                              Text 340, y_pos, 75, 10, "Footer Month: " & ALL_JOBS_PANELS_ARRAY(info_month, each_job)
                              IF ALL_JOBS_PANELS_ARRAY(EI_case_note, each_job) = TRUE Then Text 420, y_pos, 175, 10, "EARNED INCOME BUDGETING CASE NOTE FOUND"
                              ' Text 420, y_pos, 175, 10, "EARNED INCOME BUDGETING CASE NOTE FOUND"
                              y_pos = y_pos + 20
                              Text 15, y_pos, 45, 10, "Hourly Wage:"
                              EditBox 65, y_pos - 5, 40, 15, ALL_JOBS_PANELS_ARRAY(hrly_wage, each_job)
                              Text 115, y_pos, 55, 10, "Retro - Income:"
                              EditBox 170, y_pos - 5, 45, 15, ALL_JOBS_PANELS_ARRAY(job_retro_income, each_job)
                              Text 225, y_pos, 25, 10, "Hours:"
                              EditBox 250, y_pos - 5, 20, 15, ALL_JOBS_PANELS_ARRAY(retro_hours, each_job)
                              Text 305, y_pos, 55, 10, "Prosp - Income:"
                              EditBox 365, y_pos - 5, 45, 15, ALL_JOBS_PANELS_ARRAY(job_prosp_income, each_job)
                              Text 415, y_pos, 25, 10, "Hours:"
                              EditBox 440, y_pos - 5, 20, 15, ALL_JOBS_PANELS_ARRAY(prosp_hours, each_job)
                              Text 480, y_pos, 40, 10, "Pay Freq:"
                              ComboBox 525, y_pos - 5, 60, 45, "Type or select"+chr(9)+"Weekly"+chr(9)+"Biweekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly", ALL_JOBS_PANELS_ARRAY(main_pay_freq, each_job)
                              y_pos = y_pos + 20
                              If snap_checkbox = checked Then
                                  Text 15, y_pos, 35, 10, "SNAP PIC:"
                                  Text 65, y_pos, 60, 10, "Pay Date Amount: "
                                  EditBox 125, y_pos - 5, 50, 15, ALL_JOBS_PANELS_ARRAY(pic_pay_date_income, each_job)
                                  ComboBox 185, y_pos - 5, 60, 45, "Type or select"+chr(9)+"Weekly"+chr(9)+"Biweekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly", ALL_JOBS_PANELS_ARRAY(pic_pay_freq, each_job)
                                  Text 265, y_pos, 70, 10, "Prospective Amount:"
                                  EditBox 340, y_pos - 5, 60, 15, ALL_JOBS_PANELS_ARRAY(pic_prosp_income, each_job)
                                  Text 420, y_pos, 40, 10, "Calculated:"
                                  EditBox 470, y_pos - 5, 50, 15, ALL_JOBS_PANELS_ARRAY(pic_calc_date, each_job)
                                  y_pos = y_pos + 20
                              End If
                              If grh_checkbox = checked Then
                                  Text 15, y_pos, 35, 10, "GRH PIC:"
                                  Text 65, y_pos, 60, 10, "Pay Date Amount: "
                                  EditBox 125, y_pos - 5, 50, 15, ALL_JOBS_PANELS_ARRAY(grh_pay_day_income, each_job)
                                  ComboBox 185, y_pos - 5, 60, 45, "Type or select"+chr(9)+"Weekly"+chr(9)+"Biweekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly", ALL_JOBS_PANELS_ARRAY(grh_pay_freq, each_job)
                                  Text 265, y_pos, 70, 10, "Prospective Amount:"
                                  EditBox 340, y_pos - 5, 60, 15, ALL_JOBS_PANELS_ARRAY(grh_prosp_income, each_job)
                                  Text 420, y_pos, 40, 10, "Calculated:"
                                  EditBox 470, y_pos - 5, 50, 15, ALL_JOBS_PANELS_ARRAY(grh_calc_date, each_job)
                                  y_pos = y_pos + 20
                              End If
                              Text 15, y_pos, 55, 10, "Explain Budget:"
                              EditBox 70, y_pos - 5, 515, 15, ALL_JOBS_PANELS_ARRAY(budget_explain, each_job)
                              y_pos = y_pos + 25
                              if each_job = job_limit Then Exit Do
                              each_job = each_job + 1
                          Loop until each_job = UBound(ALL_JOBS_PANELS_ARRAY, 2) + 1
                          y_pos = y_pos + 5
                          ButtonGroup ButtonPressed
                            PushButton 430, y_pos, 60, 10, "previous page", previous_to_page_01_button
                            PushButton 495, y_pos - 5, 50, 15, "NEXT", next_to_page_03_button
                            CancelButton 550, y_pos - 5, 50, 15
                            PushButton 15, y_pos, 25, 10, "BUSI", BUSI_button
                            PushButton 45, y_pos, 25, 10, "JOBS", JOBS_button
                            PushButton 75, y_pos, 25, 10, "PBEN", PBEN_button
                            PushButton 105, y_pos, 25, 10, "RBIC", RBIC_button
                            PushButton 135, y_pos, 25, 10, "UNEA", UNEA_button
                            PushButton 175, y_pos, 45, 10, "prev. panel", prev_panel_button
                            PushButton 225, y_pos, 45, 10, "next panel", next_panel_button
                            PushButton 275, y_pos, 45, 10, "prev. memb", prev_memb_button
                            PushButton 325, y_pos, 45, 10, "next memb", next_memb_button
                        EndDialog

                        dialog Dialog1
                        cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
                        MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)

                        each_job = loop_start
                        Do
                            'jobs_err_msg'
                            'IF THERE IS AN EI CASE NOTE - DON'T WORRY ABOUT MUCH ERR HANDLING
                            if each_job = job_limit Then Exit Do
                            each_job = each_job + 1
                        Loop until each_job = UBound(ALL_JOBS_PANELS_ARRAY, 2)
                        If each_job = UBound(ALL_JOBS_PANELS_ARRAY, 2) Then last_job_reviewed = TRUE
                    Loop until jobs_err_msg = ""
                    job_limit = job_limit + 3
                    loop_start = loop_start + 3
                Loop until last_job_reviewed = TRUE

            Loop until ButtonPressed = next_to_page_03_button and err_msg = ""
            'CAF Dialog 3 - Self Employment
            'CAF Dialog 4 - Child Support Income
            'CAF Dialog 5 - Unearned Income and RBIC

			Do
				Do      'THIS NEEDS TO BE WORKED ON'
                    BeginDialog Dialog1, 0, 0, 556, 360, "CAF Dialog 6 - WREG, Expenses, Assets"
                      ButtonGroup ButtonPressed
                        PushButton 480, 30, 65, 15, "Update ABAWD", abawd_button
                      EditBox 40, 50, 505, 15, notes_on_wreg
                      ButtonGroup ButtonPressed
                        PushButton 235, 90, 50, 15, "Update SHEL", update_shel_button
                      DropListBox 45, 140, 100, 45, "Select ALLOWED HEST"+chr(9)+"AC/Heat - Full $493"+chr(9)+"Electric and Phone - $173"+chr(9)+"Electric ONLY - $126"+chr(9)+"Phone ONLY - $47"+chr(9)+"NONE - $0", hest_information
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
                      DropListBox 375, 190, 165, 45, " "+chr(9)+"01 - Own home, lease or roomate"+chr(9)+"02 - Family/Friends - economic hardship"+chr(9)+"03 -  servc prvdr- foster/group home"+chr(9)+"04 - Hospital/Treatment/Detox/Nursing Home"+chr(9)+"05 - Jail/Prison//Juvenile Det."+chr(9)+"06 - Hotel/Motel"+chr(9)+"07 - Emergency Shelter"+chr(9)+"08 - Place not meant for Housing"+chr(9)+"09 - Declined"+chr(9)+"10 - Unknown"+chr(9)+"Blank", living_situation
                      EditBox 315, 220, 230, 15, notes_on_address
                      EditBox 35, 255, 405, 15, notes_on_acct
                      EditBox 470, 255, 75, 15, notes_on_cash
                      EditBox 35, 275, 240, 15, notes_on_cars
                      EditBox 305, 275, 240, 15, notes_on_rest
                      EditBox 110, 295, 435, 15, notes_on_other_assets
                      EditBox 55, 320, 495, 15, verifs_needed
                      GroupBox 5, 5, 545, 65, "WREG and ABAWD Information"
                      Text 15, 20, 55, 10, "ABAWD Details:"
                      Text 75, 20, 470, 10, notes_on_abawd
                      Text 15, 35, 330, 10, notes_on_abawd_two
                      GroupBox 5, 75, 290, 165, "Expenses and Deductions"
                      Text 15, 95, 50, 10, "Total Shelter:"
                      Text 70, 95, 155, 10, total_shelter_amount
                      Text 15, 110, 275, 10, shelter_details
                      Text 15, 125, 275, 10, shelter_details_two
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
                      Text 315, 195, 55, 10, "Living Situation:"
                      Text 315, 210, 75, 10, "Notes on address:"
                      GroupBox 5, 245, 545, 70, "Assets"
                      Text 5, 325, 50, 10, "Verifs needed:"
                      ButtonGroup ButtonPressed
                        PushButton 380, 345, 60, 10, "previous page", previous_to_page_02_button
                        PushButton 445, 340, 50, 15, "NEXT", next_to_page_04_button
                        CancelButton 500, 340, 50, 15
                        PushButton 10, 55, 25, 10, "WREG", wreg_button
                        PushButton 315, 100, 25, 10, "ADDR", addr_button
                        PushButton 15, 145, 25, 10, "HEST", hest_button
                        PushButton 150, 145, 25, 10, "ACUT", acut_button
                        PushButton 15, 165, 25, 10, "COEX", coex_button
                        PushButton 15, 185, 25, 10, "DCEX", dcex_button
                        PushButton 10, 260, 25, 10, "ACCT", acct_button
                        PushButton 445, 260, 25, 10, "CASH", cash_button
                        PushButton 10, 280, 25, 10, "CARS", cars_button
                        PushButton 280, 280, 25, 10, "REST", rest_button
                        PushButton 10, 300, 25, 10, "SECU", secu_button
                        PushButton 35, 300, 25, 10, "TRAN", tran_button
                        PushButton 60, 300, 45, 10, "other assets", other_asset_button
                    EndDialog

					err_msg = ""
					income_note_error_msg = ""
					Dialog Dialog1			'Displays the second dialog
					cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
					MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
					If ButtonPressed = abawd_button Then
                        Do
                            abawd_err_msg = ""

                            notes_on_wreg = ""
                            notes_on_abawd = ""
                            notes_on_abawd_two = ""
                            dlg_len = 40
                            For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                              If ALL_MEMBERS_ARRAY(include_snap_checkbox, each_member) = checked Then
                                dlg_len = dlg_len + 95
                              End If
                            Next
                            y_pos = 10
                            BeginDialog Dialog1, 0, 0, 551, dlg_len, "ABAWD Detail"
                              For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                                If ALL_MEMBERS_ARRAY(include_snap_checkbox, each_member) = checked Then
                                  GroupBox 5, y_pos, 540, 95, "Member " & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " - " & ALL_MEMBERS_ARRAY(clt_name, each_member)
                                  y_pos = y_pos + 20
                                  Text 15, y_pos, 70, 10, "FSET WREG Status:"
                                  DropListBox 90, y_pos - 5, 130, 45, " "+chr(9)+"03  Unfit for Employment"+chr(9)+"04  Responsible for Care of Another"+chr(9)+"05  Age 60+"+chr(9)+"06  Under Age 16"+chr(9)+"07  Age 16-17, live w/ parent"+chr(9)+"08  Care of Child <6"+chr(9)+"09  Employed 30+ hrs/wk"+chr(9)+"10  Matching Grant"+chr(9)+"11  Unemployment Insurance"+chr(9)+"12  Enrolled in School/Training"+chr(9)+"13  CD Program"+chr(9)+"14  Receiving MFIP"+chr(9)+"20  Pend/Receiving DWP"+chr(9)+"15  Age 16-17 not live w/ Parent"+chr(9)+"16  50-59 Years Old"+chr(9)+"21  Care child < 18"+chr(9)+"17  Receiving RCA or GA"+chr(9)+"30  FSET Participant"+chr(9)+"02  Fail FSET Coop"+chr(9)+"33  Non-coop being referred"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(clt_wreg_status, each_member)
                                  Text 230, y_pos, 55, 10, "ABAWD Status:"
                                  DropListBox 285, y_pos - 5, 110, 45, " "+chr(9)+"01  WREG Exempt"+chr(9)+"02  Under Age 18"+chr(9)+"03  Age 50+"+chr(9)+"04  Caregiver of Minor Child"+chr(9)+"05  Pregnant"+chr(9)+"06  Employed 20+ hrs/wk"+chr(9)+"07  Work Experience"+chr(9)+"08  Other E and T"+chr(9)+"09  Waivered Area"+chr(9)+"10  ABAWD Counted"+chr(9)+"11  Second Set"+chr(9)+"12  RCA or GA Participant"+chr(9)+"13  ABAWD Banked Months"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(clt_abawd_status, each_member)
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
                                  Text 15, y_pos, 115, 10, "Number of BANKED months used:"
                                  EditBox 130, y_pos - 5, 25, 15, ALL_MEMBERS_ARRAY(numb_banked_mo, each_member)
                                  Text 170, y_pos, 45, 10, "Other Notes:"
                                  EditBox 220, y_pos - 5, 315, 15, ALL_MEMBERS_ARRAY(clt_abawd_notes, each_member)

                                  y_pos = y_pos + 15
                                End If
                              Next
                              y_pos = y_pos + 10
                              ButtonGroup ButtonPressed
                                PushButton 455, y_pos, 90, 15, "Return to Main Dialog", return_button
                            EndDialog

    						Dialog Dialog1

                            call update_wreg_and_abawd_notes

                        Loop until abawd_err_msg = ""
					End If

                    If ButtonPressed = update_shel_button Then
                        shel_client = ""
                        For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                            If ALL_MEMBERS_ARRAY(shel_exists, each_member) = TRUE Then
                                shel_client = each_member
                            End If
                        Next
                        If shel_client <> "" Then clt_shel_is_for = ALL_MEMBERS_ARRAY(full_clt, shel_client)
                        'ADD an IF here to determine the right HH member or if one is not yet selected AND preselect the one that has a SHEL'
                        Do
                            shel_err_msg = ""

                            If clt_SHEL_is_for = "Select" Then
                                dlg_len = 30
                            Else
                                dlg_len = 240
                                For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                                    If clt_shel_is_for = ALL_MEMBERS_ARRAY(full_clt, each_member) Then
                                        shel_client = each_member
                                        ALL_MEMBERS_ARRAY(shel_exists, each_member) = TRUE
                                    End If
                                Next
                            End If

                            BeginDialog Dialog1, 0, 0, 340, dlg_len, "SHEL Detail Dialog"
                              DropListBox 60, 10, 125, 45, shel_memb_list, clt_SHEL_is_for
                              Text 5, 15, 55, 10, "SHEL for Memb"
                              ButtonGroup ButtonPressed
                                PushButton 200, 10, 40, 10, "Load", load_button
                              If clt_shel_is_for <> "Select" Then
                                  ALL_MEMBERS_ARRAY
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
                                  ButtonGroup ButtonPressed
                                    PushButton 245, 220, 90, 15, "Return to Main Dialog", return_button
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

                            If ALL_MEMBERS_ARRAY(shel_retro_rent_amt, each_member) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_rent_amt, each_member)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Rent Expense."
                            If ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, each_member) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, each_member)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective Rent Expense."
                            If ALL_MEMBERS_ARRAY(shel_retro_lot_amt, each_member) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_lot_amt, each_member)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Lot Rent Expense."
                            If ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, each_member) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, each_member)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective Lot Rent Expense."
                            If ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, each_member) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, each_member)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Morgage Expense."
                            If ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, each_member) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, each_member)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective ortgage Expense."
                            If ALL_MEMBERS_ARRAY(shel_retro_ins_amt, each_member) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_ins_amt, each_member)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Insurance Expense."
                            If ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, each_member) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, each_member)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective Insurance Expense."
                            If ALL_MEMBERS_ARRAY(shel_retro_tax_amt, each_member) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_tax_amt, each_member)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Tax Expense."
                            If ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, each_member) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, each_member)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective Tax Expense."
                            If ALL_MEMBERS_ARRAY(shel_retro_room_amt, each_member) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_room_amt, each_member)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Room Expense."
                            If ALL_MEMBERS_ARRAY(shel_prosp_room_amt, each_member) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_room_amt, each_member)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective Room Expense."
                            If ALL_MEMBERS_ARRAY(shel_retro_garage_amt, each_member) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_garage_amt, each_member)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Garage Expense."
                            If ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, each_member) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, each_member)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective Garage Expense."
                            If ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, each_member) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, each_member)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Subsidy Amount."
                            If ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, each_member) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, each_member)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective Subsidy Amount."

                            If button_pressed = load_button Then shel_err_msg = "LOOP" & shel_err_msg

                            If left(shel_err_msg, 4) <> "LOOP" AND shel_err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & shel_err_msg

                            ALL_MEMBERS_ARRAY(shel_retro_rent_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_rent_amt, shel_client) * 1
                            ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, shel_client) * 1
                            ALL_MEMBERS_ARRAY(shel_retro_lot_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_lot_amt, shel_client) * 1
                            ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, shel_client) * 1
                            ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, shel_client) * 1
                            ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, shel_client) * 1
                            ALL_MEMBERS_ARRAY(shel_retro_ins_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_ins_amt, shel_client) * 1
                            ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, shel_client) * 1
                            ALL_MEMBERS_ARRAY(shel_retro_tax_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_tax_amt, shel_client) * 1
                            ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, shel_client) * 1
                            ALL_MEMBERS_ARRAY(shel_retro_room_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_room_amt, shel_client) * 1
                            ALL_MEMBERS_ARRAY(shel_prosp_room_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_room_amt, shel_client) * 1
                            ALL_MEMBERS_ARRAY(shel_retro_garage_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_garage_amt, shel_client) * 1
                            ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, shel_client) * 1
                            ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, shel_client) * 1
                            ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, shel_client) * 1

                            call update_shel_notes
                        Loop until shel_err_msg = ""
                    End If
					IF (earned_income <> "" AND trim(notes_on_income) = "") OR (unearned_income <> "" AND notes_on_income = "") THEN income_note_error_msg = True
					If err_msg <> "" THEN Msgbox err_msg
				Loop until ButtonPressed = (next_to_page_04_button AND err_msg = "") or (ButtonPressed = previous_to_page_02_button AND err_msg = "")		'If you press either the next or previous button, this loop ends
				If ButtonPressed = previous_to_page_01_button then exit do		'If the button was previous, it exits this do loop and is caught in the next one, which sends you back to Dialog 1 because of the "If ButtonPressed = previous_to_page_01_button then exit do" later on
				Do
BeginDialog Dialog1, 0, 0, 451, 405, "CAF Dialog 7 - Interview Info"
  EditBox 60, 45, 385, 15, INSA
  EditBox 35, 65, 410, 15, ACCI
  EditBox 35, 85, 175, 15, DIET
  EditBox 245, 85, 200, 15, BILS
  EditBox 35, 105, 285, 15, FMED
  EditBox 390, 105, 55, 15, retro_request
  EditBox 180, 130, 265, 15, reason_expedited_wasnt_processed
  EditBox 100, 150, 345, 15, FIAT_reasons
  CheckBox 15, 190, 80, 10, "Application signed?", application_signed_checkbox
  CheckBox 15, 205, 65, 10, "Appt letter sent?", appt_letter_sent_checkbox
  CheckBox 15, 220, 150, 10, "Client willing to participate with E and T", E_and_T_checkbox
  CheckBox 15, 235, 70, 10, "EBT referral sent?", EBT_referral_checkbox
  CheckBox 115, 190, 50, 10, "eDRS sent?", eDRS_sent_checkbox
  CheckBox 115, 205, 50, 10, "Expedited?", expedited_checkbox
  CheckBox 115, 235, 70, 10, "IAAs/OMB given?", IAA_checkbox
  CheckBox 200, 190, 115, 10, "Informed client of recert period?", recert_period_checkbox
  CheckBox 200, 205, 80, 10, "Intake packet given?", intake_packet_checkbox
  CheckBox 200, 220, 105, 10, "Managed care packet sent?", managed_care_packet_checkbox
  CheckBox 200, 235, 105, 10, "Managed care referral made?", managed_care_referral_checkbox
  CheckBox 345, 190, 65, 10, "R/R explained?", R_R_checkbox
  CheckBox 345, 205, 85, 10, "Sent forms to AREP?", Sent_arep_checkbox
  CheckBox 345, 220, 65, 10, "Updated MMIS?", updated_MMIS_checkbox
  CheckBox 345, 235, 95, 10, "Workforce referral made?", WF1_checkbox
  EditBox 55, 260, 230, 15, other_notes
  EditBox 55, 280, 390, 15, verifs_needed
  EditBox 55, 300, 390, 15, actions_taken
  ComboBox 330, 260, 115, 15, " "+chr(9)+"incomplete"+chr(9)+"approved", CAF_status
  CheckBox 15, 335, 240, 10, "Check here to update PND2 to show client delay (pending cases only).", client_delay_checkbox
  CheckBox 15, 350, 200, 10, "Check here to create a TIKL to deny at the 30/45 day mark.", TIKL_checkbox
  CheckBox 15, 365, 265, 10, "Check here to send a TIKL (10 days from now) to update PND2 for Client Delay.", client_delay_TIKL_checkbox
  EditBox 395, 345, 50, 15, worker_signature
  ButtonGroup ButtonPressed
    PushButton 290, 370, 45, 10, "prev. page", previous_to_page_03_button
    OkButton 340, 365, 50, 15
    CancelButton 395, 365, 50, 15
    PushButton 10, 15, 20, 10, "DWP", ELIG_DWP_button
    PushButton 30, 15, 15, 10, "FS", ELIG_FS_button
    PushButton 45, 15, 15, 10, "GA", ELIG_GA_button
    PushButton 60, 15, 15, 10, "HC", ELIG_HC_button
    PushButton 75, 15, 20, 10, "MFIP", ELIG_MFIP_button
    PushButton 95, 15, 20, 10, "MSA", ELIG_MSA_button
    PushButton 115, 15, 15, 10, "WB", ELIG_WB_button
    PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 335, 25, 45, 10, "next panel", next_panel_button
    PushButton 395, 25, 45, 10, "next memb", next_memb_button
  GroupBox 5, 5, 130, 25, "ELIG panels:"
  GroupBox 330, 5, 115, 35, "STAT-based navigation"
  ButtonGroup ButtonPressed
    PushButton 5, 50, 25, 10, "INSA/", INSA_button
    PushButton 30, 50, 25, 10, "MEDI:", MEDI_button
    PushButton 5, 70, 25, 10, "ACCI:", ACCI_button
    PushButton 5, 90, 25, 10, "DIET:", DIET_button
    PushButton 5, 110, 25, 10, "FMED:", FMED_button
  Text 5, 135, 170, 10, "Reason expedited wasn't processed (if applicable):"
  Text 5, 155, 95, 10, "FIAT reasons (if applicable):"
  GroupBox 5, 175, 440, 75, "Common elements workers should case note:"
  Text 5, 265, 50, 10, "Other notes:"
  Text 290, 265, 40, 10, "CAF status:"
  Text 5, 285, 50, 10, "Verifs needed:"
  Text 5, 305, 50, 10, "Actions taken:"
  GroupBox 5, 320, 280, 60, "Actions the script can do:"
  Text 330, 350, 60, 10, "Worker signature:"
  ButtonGroup ButtonPressed
    PushButton 325, 110, 60, 10, "Retro Req. date:", HCRE_button
    PushButton 215, 90, 25, 10, "BILS:", BILS_button
EndDialog

					err_msg = ""
					Dialog Dialog1			'Displays the third dialog
					cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
					MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
					If ButtonPressed = previous_to_page_02_button then exit do		'Exits this do...loop here if you press previous. The second ""loop until ButtonPressed = -1" gets caught, and it loops back to the "Do" after "Loop until ButtonPressed = next_to_page_02_button"
					If actions_taken = "" THEN err_msg = err_msg & vbCr & "Please complete actions taken section."    'creating err_msg if required items are missing
					If income_note_error_msg = True THEN err_msg = err_msg & VbCr & "Income for this case was found in MAXIS. Please complete the 'notes on income and budget' field."
					If worker_signature = "" THEN err_msg = err_msg & vbCr & "Please enter a worker signature."
					If CAF_status = " " THEN err_msg = err_msg & vbCr & "Please select a CAF Status."
					If err_msg <> "" THEN Msgbox err_msg
				Loop until (ButtonPressed = -1 and err_msg = "") or (ButtonPressed = previous_to_page_03_button and err_msg = "")		'If OK or PREV, it exits the loop here, which is weird because the above also causes it to exit
			Loop until ButtonPressed = -1	'Because this is in here a second time, it triggers a return to the "Dialog CAF_dialog_02" line, where all those "DOs" start again!!!!!
			If ButtonPressed = previous_to_page_01_button then exit do 	'This exits this particular loop again for prev button on page 2, which sends you back to page 1!!
		Loop until err_msg = ""		'Loops all of that until those four sections are finished. Let's move that over to those particular pages. Folks would be less angry that way I bet.
		CALL proceed_confirmation(case_note_confirm)			'Checks to make sure that we're ready to case note.
	Loop until case_note_confirm = TRUE							'Loops until we affirm that we're ready to case note.
	call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false

Call check_for_maxis(FALSE)  'allows for looping to check for maxis after worker has complete dialog box so as not to lose a giant CAF case note if they get timed out while writing.

'This code will update the interview date in PROG.
If CAF_type <> "Recertification" AND CAF_type <> "Addendum" Then        'Interview date is not on PROG for recertifications or addendums
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
                if DateDiff("d", cash_one_app, CAF_datestamp) = 0 then prog_row = 6     'If date of application on PROG matches script date of applicaton
            End If
            If cash_two_app <> "__/__/__" Then
                if DateDiff("d", cash_two_app, CAF_datestamp) = 0 then prog_row = 7
            End If

            If grh_cash_app <> "__/__/__" Then
                if DateDiff("d", grh_cash_app, CAF_datestamp) = 0 then prog_row = 9
            End If

            EMReadScreen entered_intv_date, 8, prog_row, 55                     'Reading the right interview date with row defined above
            'MsgBox "Cash interview date - " & entered_intv_date
            If entered_intv_date = "__ __ __" Then intv_date_needed = TRUE      'If this is blank - script needs to prompt worker to have it updated
        End If

        If intv_date_needed = TRUE Then         'If previous code has determined that PROG needs to be updated
            If SNAP_checkbox = checked Then prog_update_SNAP_checkbox = checked     'Auto checking based on the programs the script is being run for.
            If cash_checkbox = checked Then prog_update_cash_checkbox = checked

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
                err_msg = ""
                Dialog Dialog1
                'Requiring a reason for not updating PROG and making sure if confirm is updated that a program is selected.
                If do_not_update_prog = 1 AND no_update_reason = "" Then err_msg = err_msg & vbNewLine & "* If PROG is not to be updated, please explain why PROG should not be updated."
                IF confirm_update_prog = 1 AND prog_update_SNAP_checkbox = unchecked AND prog_update_cash_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "* Select either CASH or SNAP to have updated on PROG."

                If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
            Loop until err_msg = ""

            If confirm_update_prog = 1 Then     'If the dialog selects to have PROG updated
                CALL back_to_SELF               'Need to do this because we need to go to the footer month of the application and we may be in a different month

                keep_footer_month = MAXIS_footer_month      'Saving the footer month and year that was determined earlier in the script. It needs t obe changed for nav functions to work correctly
                keep_footer_year = MAXIS_footer_year

                app_month = DatePart("m", CAF_datestamp)    'Setting the footer month and year to the app month.
                app_year = DatePart("yyyy", CAF_datestamp)

                MAXIS_footer_month = right("00" & app_month, 2)
                MAXIS_footer_year = right(app_year, 2)

                CALL navigate_to_MAXIS_screen ("STAT", "PROG")  'Now we can navigate to PROG in the application footer month and year
                PF9                                             'Edit

                intv_mo = DatePart("m", interview_date)     'Setting the date parts to individual variables for ease of writing
                intv_day = DatePart("d", interview_date)
                intv_yr = DatePart("yyyy", interview_date)

                intv_mo = right("00"&intv_mo, 2)            'formatting variables in to 2 digit strings - because MAXIS
                intv_day = right("00"&intv_day, 2)
                intv_yr = right(intv_yr, 2)

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
                        if DateDiff("d", cash_one_app, CAF_datestamp) = 0 then prog_row = 6
                    End If
                    If cash_two_app <> "__/__/__" Then
                        if DateDiff("d", cash_two_app, CAF_datestamp) = 0 then prog_row = 7
                    End If

                    If grh_cash_app <> "__/__/__" Then
                        if DateDiff("d", grh_cash_app, CAF_datestamp) = 0 then prog_row = 9
                    End If

                    EMWriteScreen intv_mo, prog_row, 55     'Writing the interview date in
                    EMWriteScreen intv_day, prog_row, 58
                    EMWriteScreen intv_yr, prog_row, 61
                End If

                transmit                                    'Saving the panel

                MAXIS_footer_month = keep_footer_month      'resetting the footer month and year so the rest of the script uses the worker identified footer month and year.
                MAXIS_footer_year = keep_footer_year
            End If
        ENd If
    End If
End If
'MsgBox "PROG Stuff done"
'MsgBox confirm_update_prog

'Now, the client_delay_checkbox business. It'll update client delay if the box is checked and it isn't a recert.
If client_delay_checkbox = checked and CAF_type <> "Recertification" then
	call navigate_to_MAXIS_screen("rept", "pnd2")
	EMGetCursor PND2_row, PND2_col
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
End if

'Going to TIKL. Now using the write TIKL function
If TIKL_checkbox = checked and CAF_type <> "Recertification" then
	If cash_checkbox = checked or EMER_checkbox = checked or SNAP_checkbox = checked then
		If DateDiff ("d", CAF_datestamp, date) > 30 Then 'Error handling to prevent script from attempting to write a TIKL in the past
			MsgBox "Cannot set TIKL as CAF Date is over 30 days old and TIKL would be in the past. You must manually track."
		Else
			call navigate_to_MAXIS_screen("dail", "writ")
			call create_MAXIS_friendly_date(CAF_datestamp, 30, 5, 18)
			If cash_checkbox = checked then TIKL_msg_one = TIKL_msg_one & "Cash/"
			If SNAP_checkbox = checked then TIKL_msg_one = TIKL_msg_one & "SNAP/"
			If EMER_checkbox = checked then TIKL_msg_one = TIKL_msg_one & "EMER/"
			TIKL_msg_one = Left(TIKL_msg_one, (len(TIKL_msg_one) - 1))
			TIKL_msg_one = TIKL_msg_one & " has been pending for 30 days. Evaluate for possible denial."
			Call write_variable_in_TIKL (TIKL_msg_one)
			PF3
		End If
	End if
	If HC_checkbox = checked then
		If DateDiff ("d", CAF_datestamp, date) > 45 Then 'Error handling to prevent script from attempting to write a TIKL in the past
			MsgBox "Cannot set TIKL as CAF Date is over 45 days old and TIKL would be in the past. You must manually track."
		Else
			call navigate_to_MAXIS_screen("dail", "writ")
			call create_MAXIS_friendly_date(CAF_datestamp, 45, 5, 18)
			Call write_variable_in_TIKL ("HC pending 45 days. Evaluate for possible denial. If any members are elderly/disabled, allow an additional 15 days and reTIKL out.")
			PF3
		End If
	End if
End if
If client_delay_TIKL_checkbox = checked then
	call navigate_to_MAXIS_screen("dail", "writ")
	call create_MAXIS_friendly_date(date, 10, 5, 18)
	Call write_variable_in_TIKL (">>>UPDATE PND2 FOR CLIENT DELAY IF APPROPRIATE<<<")
	PF3
End if
'----Here's the new bit to TIKL to APPL the CAF for CAF_datestamp if the CL fails to complete the CASH/SNAP reinstate and then TIKL again for DateAdd("D", 30, CAF_datestamp) to evaluate for possible denial.
'----IF the DatePart("M", CAF_datestamp) = MAXIS_footer_month (DatePart("M", CAF_datestamp) is converted to footer_comparo_month for the sake of comparison) and the CAF_status <> "Approved" and CAF_type is a recertification AND cash or snap is checked, then
'---------the script generates a TIKL.
footer_comparison_month = DatePart("M", CAF_datestamp)
IF len(footer_comparison_month) <> 2 THEN footer_comparison_month = "0" & footer_comparison_month
IF CAF_type = "Recertification" AND MAXIS_footer_month = footer_comparison_month AND CAF_status <> "approved" AND (cash_checkbox = checked OR SNAP_checkbox = checked) THEN
	CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
	start_of_next_month = DatePart("M", DateAdd("M", 1, CAF_datestamp)) & "/01/" & DatePart("YYYY", DateAdd("M", 1, CAF_datestamp))
	denial_consider_date = DateAdd("D", 30, CAF_datestamp)
	CALL create_MAXIS_friendly_date(start_of_next_month, 0, 5, 18)
	EMWriteScreen ("IF CLIENT HAS NOT COMPLETED RECERT, APPL CAF FOR " & CAF_datestamp), 9, 3
	EMWriteScreen ("AND TIKL FOR " & denial_consider_date & " TO EVALUATE FOR POSSIBLE DENIAL."), 10, 3
	transmit
	PF3
END IF

IF tikl_for_ui THEN
	Call navigate_to_MAXIS_screen ("DAIL", "WRIT")
	two_weeks_from_now = DateAdd("d", 14, date)
	call create_MAXIS_friendly_date(two_weeks_from_now, 10, 5, 18)
	call write_variable_in_TIKL ("Review client's application for Unemployment and request an update if needed.")
	PF3
END IF
'--------------------END OF TIKL BUSINESS

'Navigates to case note, and checks to make sure we aren't in inquiry.
Call start_a_blank_CASE_NOTE

'Adding a colon to the beginning of the CAF status variable if it isn't blank (simplifies writing the header of the case note)
If CAF_status <> "" then CAF_status = ": " & CAF_status

'Adding footer month to the recertification case notes
If CAF_type = "Recertification" then CAF_type = MAXIS_footer_month & "/" & MAXIS_footer_year & " recert"

'THE CASE NOTE-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
CALL write_variable_in_CASE_NOTE("***" & CAF_type & CAF_status & "***")
IF move_verifs_needed = TRUE THEN
	CALL write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)			'IF global variable move_verifs_needed = True (on FUNCTIONS FILE), it'll case note at the top.
	CALL write_variable_in_CASE_NOTE("------------------------------")
End if
CALL write_bullet_and_variable_in_CASE_NOTE("CAF datestamp", CAF_datestamp)
If Used_Interpreter_checkbox = checked then
	CALL write_variable_in_CASE_NOTE("* Interview type: " & interview_type & " w/ interpreter")
Else
	CALL write_bullet_and_variable_in_CASE_NOTE("Interview type", interview_type)
End if
CALL write_bullet_and_variable_in_CASE_NOTE("Interview date", interview_date)
'If intv_date_needed = FALSE THen CALL write_variable_in_CASE_NOTE("* Interview date entered on PROG by worker prior to script run.")
If confirm_update_prog = 1 Then CALL write_variable_in_CASE_NOTE("* Interview date entered on PROG for " & programs_w_interview)
If do_not_update_prog = 1 Then CALL write_bullet_and_variable_in_CASE_NOTE("PROG WAS NOT UPDATED WITH INTERVIEW DATE, because", no_update_reason)
CALL write_bullet_and_variable_in_CASE_NOTE("HC document received", HC_document_received)
CALL write_bullet_and_variable_in_CASE_NOTE("HC datestamp", HC_datestamp)
CALL write_bullet_and_variable_in_CASE_NOTE("Programs applied for", programs_applied_for)
CALL write_bullet_and_variable_in_CASE_NOTE("How CAF was received", how_app_rcvd)
CALL write_bullet_and_variable_in_CASE_NOTE("HH comp/EATS", HH_comp)
CALL write_bullet_and_variable_in_CASE_NOTE("Cit/ID", cit_id)
CALL write_bullet_and_variable_in_CASE_NOTE("IMIG", IMIG)
CALL write_bullet_and_variable_in_CASE_NOTE("AREP", AREP)
CALL write_bullet_and_variable_in_CASE_NOTE("FACI", FACI)
CALL write_bullet_and_variable_in_CASE_NOTE("SCHL/STIN/STEC", SCHL)
CALL write_bullet_and_variable_in_CASE_NOTE("DISA", DISA)
CALL write_bullet_and_variable_in_CASE_NOTE("PREG", PREG)
Call write_bullet_and_variable_in_CASE_NOTE("EMPS", EMPS)
If MFIP_DVD_checkbox = 1 then Call write_variable_in_CASE_NOTE("* MFIP financial orientation DVD sent to participant(s).")
CALL write_bullet_and_variable_in_CASE_NOTE("ABPS", ABPS)
CALL write_bullet_and_variable_in_CASE_NOTE("Earned inc.", earned_income)
CALL write_bullet_and_variable_in_CASE_NOTE("UNEA", unearned_income)
CALL write_bullet_and_variable_in_CASE_NOTE("Notes on income and budget", notes_on_income)
CALL write_bullet_and_variable_in_CASE_NOTE("STWK/inc. changes", income_changes)
CALL write_bullet_and_variable_in_CASE_NOTE("ABAWD Notes/WREG", notes_on_abawd)
CALL write_bullet_and_variable_in_CASE_NOTE("Is any work temporary", is_any_work_temporary)
CALL write_bullet_and_variable_in_CASE_NOTE("SHEL/HEST", SHEL_HEST)
CALL write_bullet_and_variable_in_CASE_NOTE("COEX/DCEX", COEX_DCEX)
CALL write_bullet_and_variable_in_CASE_NOTE("CASH/ACCTs", CASH_ACCTs)
CALL write_bullet_and_variable_in_CASE_NOTE("Other assets", other_assets)
CALL write_bullet_and_variable_in_CASE_NOTE("INSA", INSA)
CALL write_bullet_and_variable_in_CASE_NOTE("ACCI", ACCI)
CALL write_bullet_and_variable_in_CASE_NOTE("DIET", DIET)
CALL write_bullet_and_variable_in_CASE_NOTE("BILS", BILS)
CALL write_bullet_and_variable_in_CASE_NOTE("FMED", FMED)
CALL write_bullet_and_variable_in_CASE_NOTE("Retro Request (IF applicable)", retro_request)
IF application_signed_checkbox = checked THEN
	CALL write_variable_in_CASE_NOTE("* Application was signed.")
Else
	CALL write_variable_in_CASE_NOTE("* Application was not signed.")
END IF
' IF appt_letter_sent_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Appointment letter was sent before interview.")
' IF EBT_referral_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* EBT referral made for client.")
IF eDRS_sent_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* eDRS sent.")
' IF expedited_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Expedited SNAP.")
' CALL write_bullet_and_variable_in_CASE_NOTE("Reason expedited wasn't processed", reason_expedited_wasnt_processed)		'This is strategically placed next to expedited checkbox entry.
IF IAA_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* IAAs/OMB given to client.")
IF intake_packet_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Client received intake packet.")
' IF managed_care_packet_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Client received managed care packet.")
' IF managed_care_referral_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Managed care referral made.")
IF R_R_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* R/R explained to client.")
IF updated_MMIS_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Updated MMIS.")
IF WF1_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Workforce referral made.")
IF Sent_arep_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Sent form(s) to AREP.")
IF E_and_T_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Client is willing to participate with E&T.")
IF recert_period_checkbox = checked THEN call write_variable_in_CASE_NOTE("* Informed client of recert period.")
IF client_delay_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* PND2 updated to show client delay.")
CALL write_bullet_and_variable_in_CASE_NOTE("FIAT reasons", FIAT_reasons)
CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
IF move_verifs_needed = False THEN CALL write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)			'IF global variable move_verifs_needed = False (on FUNCTIONS FILE), it'll case note at the bottom.
CALL write_bullet_and_variable_in_CASE_NOTE("Actions taken", actions_taken)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

IF SNAP_recert_is_likely_24_months = TRUE THEN					'if we determined on stat/revw that the next SNAP recert date isn't 12 months beyond the entered footer month/year
	TIKL_for_24_month = msgbox("Your SNAP recertification date is listed as " & SNAP_recert_date & " on STAT/REVW. Do you want set a TIKL on " & dateadd("m", "-1", SNAP_recert_compare_date) & " for 12 month contact?" & vbCR & vbCR & "NOTE: Clicking yes will navigate away from CASE/NOTE saving your case note.", VBYesNo)
	IF TIKL_for_24_month = vbYes THEN 												'if the select YES then we TIKL using our custom functions.
		CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
		CALL create_MAXIS_friendly_date(dateadd("m", "-1", SNAP_recert_compare_date), 0, 5, 18)
		CALL write_variable_in_TIKL("If SNAP is open, review to see if 12 month contact letter is needed. DAIL scrubber can send 12 Month Contact Letter if used on this TIKL.")
	END IF
END IF

end_msg = "Success! CAF has been successfully noted. Please remember to run the Approved Programs, Closed Programs, or Denied Programs scripts if  results have been APP'd."
If do_not_update_prog = 1 Then end_msg = end_msg & vbNewLine & vbNewLine & "It was selected that PROG would NOT be updated because " & no_update_reason
script_end_procedure(end_msg)
