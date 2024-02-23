'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - PROOF OF RELATIONSHIP.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 85           'manual run time in seconds
STATS_denomination = "M"        'M is for household member
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
call changelog_update("11/04/2019", "Updated the script to better review if complete information has been provided for documenting relationship in the dialog. Adjusted handling to prevent blank notes.", "Casey Love, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT------------------------------------------------------------------------------------------------------
'Connects to BLUEZONE
EMConnect ""
'Grabs the MAXIS case number
CALL MAXIS_case_number_finder(MAXIS_case_number)
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 130, 55, "Case Number"
  EditBox 60, 5, 60, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 15, 30, 50, 15
    CancelButton 70, 30, 50, 15
  Text 10, 10, 50, 10, "Case Number: "
EndDialog
DO
	Do
    	err_msg = ""
    	Dialog Dialog1											'Running the case number dialog
    	cancel_confirmation							'Cancels the script if cancel button is pressed
    	If MAXIS_case_number = "" Then err_msg = "Please enter a case number."	'Case number must be entered or script will error out
    	If err_msg <> "" Then MsgBox err_msg								'Tells worker to enter a case number
    Loop until err_msg = ""
	CALL check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = False

'Generate dialog of all of the HH Membs and then create two arrays - 1) All HH Membs to gather relationship details (HH_Member_array), and 2) only selected HH Membs to display in dropdowns (selected_HH_Member_array)
CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.
EMWriteScreen "01", 20, 76						''make sure to start at Memb 01
transmit

DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
	EMReadscreen ref_nbr, 3, 4, 33
	EMReadScreen access_denied_check, 13, 24, 2
	'MsgBox access_denied_check
	If access_denied_check = "ACCESS DENIED" Then
		PF10
		EMWaitReady 0, 0
		last_name = "UNABLE TO FIND"
		first_name = " - Access Denied"
		mid_initial = ""
	Else
		EMReadscreen last_name, 25, 6, 30
		EMReadscreen first_name, 12, 6, 63
		EMReadscreen mid_initial, 1, 6, 79
		last_name = trim(replace(last_name, "_", "")) & " "
		first_name = trim(replace(first_name, "_", "")) & " "
		mid_initial = replace(mid_initial, "_", "")
	End If
	client_string = ref_nbr & last_name & first_name & mid_initial
	client_array = client_array & client_string & "|"
	transmit
	Emreadscreen edit_check, 7, 24, 2
LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

client_array = TRIM(client_array)
test_array = split(client_array, "|")
total_clients = Ubound(test_array)			'setting the upper bound for how many spaces to use from the array

DIM all_client_array()
ReDim all_clients_array(total_clients, 1)

FOR x = 0 to total_clients				'using a dummy array to build in the autofilled check boxes into the array used for the dialog.
	Interim_array = split(client_array, "|")
	all_clients_array(x, 0) = Interim_array(x)
	all_clients_array(x, 1) = 1
NEXT

BEGINDIALOG HH_memb_dialog, 0, 0, 241, (35 + (total_clients * 15)), "HH Member Dialog"   'Creates the dynamic dialog. The height will change based on the number of clients it finds.
	Text 10, 5, 105, 10, "Household members to look at:"
	FOR i = 0 to total_clients										'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
		IF all_clients_array(i, 0) <> "" THEN checkbox 10, (20 + (i * 15)), 160, 10, all_clients_array(i, 0), all_clients_array(i, 1)  'Ignores and blank scanned in persons/strings to avoid a blank checkbox
	NEXT
	ButtonGroup ButtonPressed
	OkButton 185, 10, 50, 15
	CancelButton 185, 30, 50, 15
ENDDIALOG

'runs the dialog that has been dynamically created. Streamlined with new functions.
Dialog HH_memb_dialog
Cancel_without_confirmation
check_for_maxis(True)

HH_member_array = ""

FOR i = 0 to total_clients
	IF all_clients_array(i, 0) <> "" THEN HH_member_array = HH_member_array & left(all_clients_array(i, 0), 2) & " "
NEXT

HH_member_array = TRIM(HH_member_array)							'Cleaning up array for ease of use.
HH_member_array = SPLIT(HH_member_array, " ")

selected_HH_member_array = "*"

FOR i = 0 to total_clients
	IF all_clients_array(i, 0) <> "" THEN 						'creates the final array to be used by other scripts.
		IF all_clients_array(i, 1) = 1 THEN						'if the person/string has been checked on the dialog then the reference number portion (left 2) will be added to new HH_member_array
			'msgbox all_clients_
			selected_HH_member_array = selected_HH_member_array & left(all_clients_array(i, 0), 2) & "*"
		END IF
	END IF
NEXT

selected_HH_member_array = TRIM(selected_HH_member_array)							'Cleaning up array for ease of use.

Call convert_array_to_droplist_items(HH_Member_Array, hh_member_dropdown)

Dim Pare_Line_Array ()				'Defines the array that will store multiple relationships
ReDim Pare_Line_Array (5, 0)		'Redefines the array to make it multi-dimensional and dynamic

'Setting constants so that the array coordinates are easier to read
Const received_for = 0
Const relationship_type = 1
Const other_relationship_list = 2
Const relationship_to = 3
Const documents_received = 4
Const new_checkbox = 5

array_item = 0 						'Setting the initial array item so that we can increment and add new array coordinates

'Goind to PARE to pull any child, grandchild, fosterchild, stepchild relationships that are already coded into MAXIS
For each member in HH_Member_Array																	'Looks at PARE for every houshold member
	STATS_counter = STATS_counter + 1																'Statistics gathering by HH member
	line_to_read = 8 																				'Setting the initial MAXIS line to read so we can increment
	Call Navigate_to_MAXIS_Screen ("STAT", "PARE")													'Go to PARE
	EMWriteScreen member, 20, 76																	'Go to PARE for the household member
	transmit
	Do
		EMReadScreen child_ref_num, 2, line_to_read, 24 											'Reads the reference number space on the current line
		If child_ref_num = "__" Then Exit Do 														'Once the line is blank, the script will stop reading more lines on this PARE panel as there is no more relevant information available
		ReDim Preserve Pare_Line_Array(5, array_item)												'Redefines the Array to resize based on the looping
		EMReadScreen relationship_code, 1, line_to_read, 53 										'Gets the type of relationship listed for this line on PARE
		EMReadScreen verif_code, 2, line_to_read, 71												'Gets the type of proof that is coded on file for this relationship on PARE
		Pare_Line_Array(relationship_to, array_item) = member 										'Adds the member reference number to the array of the client whose PARE panel it is (the parent) to the array
		Pare_Line_Array(received_for, array_item) = trim(child_ref_num)								'Adds the member reference number of the child listed on the current line of the PARE panel to the array

		Select Case relationship_code																'This is a logic function that will compare the type of relationship to a set of options to define known relationships and adds to the array
		Case "1"
			Pare_Line_Array(relationship_type, array_item) = "Is the Child of"
		Case "2"
			Pare_Line_Array(relationship_type, array_item) = "Is the Step Child of"
		Case "3"
			Pare_Line_Array(relationship_type, array_item) = "Is the Grandchild of"
		Case "5"
			Pare_Line_Array(relationship_type, array_item) = "Is the Foster Child of"
		Case Else 																					'There are some relationships that are not specific on PARE but are on MEMB - noting these for further investigation
			Pare_Line_Array(relationship_type, array_item) = "Needed"
		End Select

		Select Case verif_code																		'This is a logic function that will compart the verification code to a set of options to define known verifs and adds to the array
		Case "BC"
			Pare_Line_Array(documents_received, array_item) = "Birth Certificate"
		Case "RP"
			Pare_Line_Array(documents_received, array_item) = "Recognition of Parentge"
		Case "AR"
			Pare_Line_Array(documents_received, array_item) = "Adoption Records"
		Case "HR"
			Pare_Line_Array(documents_received, array_item) = "Hospital Record"
		Case "NO"
			Pare_Line_Array(documents_received, array_item) = "NONE ON FILE"
		Case Else
			Pare_Line_Array(documents_received, array_item) = ""
		End Select
		array_item = array_item + 1																	'Incremebing the array and the maxis row
		line_to_read = line_to_read + 1
		If line_to_read = 18 Then 																	'PARE only holds 10 lines per page - if there are more, PARE needs to PF8 to get to the next list
			PF8
			line_to_read = 8 																		'Resets the maxis row to start back at the top if it had to PF8
		End If
	Loop until child_ref_num = "__"
next

'The script will now go to STAT/MEMB to gather certain relationships of other HH members to M01
Call navigate_to_MAXIS_screen ("STAT", "MEMB")
For each member in HH_Member_Array
	EMWriteScreen member, 20, 76																	'Enters each reference number in the household member array to view each members MEMB panel
	transmit
	EMReadScreen rel_to_applicant, 2, 10, 42 														'Reads the relationship code on the current MEMB panel
	Select Case rel_to_applicant																	'Logic function to define a relationship for certain types of relationship codes
	Case "02"																						'Spouse
		ReDim Preserve Pare_Line_Array(5, array_item)
		Pare_Line_Array (received_for,      array_item) = member
		Pare_Line_Array (relationship_type, array_item) = "Is the Spouse of"
		Pare_Line_Array (relationship_to,   array_item) = "01"										'Always M01 because the only relationship defined on MEMB is in relation to M01
		array_item = array_item + 1
	Case "04"																						'Parent
		ReDim Preserve Pare_Line_Array(5, array_item)
		Pare_Line_Array (received_for,      array_item) = member
		Pare_Line_Array (relationship_type, array_item) = "Is the Parent of"
		Pare_Line_Array (relationship_to,   array_item) = "01"
		array_item = array_item + 1
	Case "18"																						'Legal Guardian
		ReDim Preserve Pare_Line_Array(5, array_item)
		Pare_Line_Array (received_for,      array_item) = member
		Pare_Line_Array (relationship_type, array_item) = "Is the Guardian of"
		Pare_Line_Array (relationship_to,   array_item) = "01"
		array_item = array_item + 1
	Case "24"																						'Not Related
		ReDim Preserve Pare_Line_Array(5, array_item)
		Pare_Line_Array (received_for,      array_item) = member
		Pare_Line_Array (relationship_type, array_item) = "Is Unrelated to"
		Pare_Line_Array (relationship_to,   array_item) = "01"
		array_item = array_item + 1
	End Select
Next

rel_to_applicant = ""																				'Blanking out a variable to prevent problems

For pare_item = 0 to UBound(Pare_Line_Array,2) 														'checks through the array
	If Pare_Line_Array(relationship_type, pare_item) = "Needed" AND Pare_Line_Array(relationship_to, pare_item) = "01" Then 	'Finds the relationship types that were coded above as needing further investigation for ONLY relationships to M01
		Call navigate_to_MAXIS_screen ("STAT", "MEMB") 												'Goes to STAT/MEMB
		EMWriteScreen Pare_Line_Array(received_for, pare_item), 20, 76								'Navigates to the member panel of the client who is related to M01 in a way that needs definition
		Transmit
		EMReadScreen rel_to_applicant, 2, 10, 42													'Reads the relationship of this client to M01
		Select Case rel_to_applicant																'Logic function to define relationship and adds it to the array
		Case "05"																					'Sibling
			Pare_Line_Array(relationship_type, pare_item) = "Is the Sibling of"
		Case "12"																					'Neice
			Pare_Line_Array(relationship_type, pare_item) = "Is the Niece of"
		Case "13"																					'Nephew
			Pare_Line_Array(relationship_type, pare_item) = "Is the Nephew of"
		Case Else
			Pare_Line_Array(relationship_type, pare_item) = ""										'If the relationship is other that the 3 defined here, it blanks 'needed' out of the array for the worker to enter manually
		End Select
	End If
Next

' ReDim Preserve Pare_Line_Array(5, array_item)														'Adds one blank array item to the array for a manual entry since the dialog is dynamically created based on the number of array items there are. If the script could not find all the relationships, the worker needs to be able manually add one.
' array_item = array_item + 1
array_item = UBound(Pare_Line_Array, 2)
selected_hh_memb_relationship_count = 0

For pare_item = 0 to UBound(Pare_Line_Array, 2)
	If Instr(selected_HH_member_array, Pare_Line_Array(received_for, pare_item)) OR Instr(selected_HH_member_array, Pare_Line_Array(relationship_to, pare_item)) Then
		selected_hh_memb_relationship_count = selected_hh_memb_relationship_count + 1
	End If
Next

add_field_count = 0

DO
    Do
		dialog_count = 0
		err_msg = ""
		
		'-------------------------------------------------------------------------------------------------DIALOG
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 570, (95 + (20 * ((selected_hh_memb_relationship_count + add_field_count)))), "Proof of Relationship"
			For pare_item = 0 to UBound(Pare_Line_Array, 2)
				If Instr(selected_HH_member_array, Pare_Line_Array(received_for, pare_item)) OR Instr(selected_HH_member_array, Pare_Line_Array(relationship_to, pare_item)) Then
					DropListBox 5, (20 + (dialog_count * 20)), 70, 15, hh_member_dropdown, Pare_Line_Array(received_for, pare_item)
					DropListBox 85, (20 + (dialog_count * 20)), 110, 15, "Select one..."+chr(9)+"Is Another Relative of"+chr(9)+"Is the Child of"+chr(9)+"Is the Foster Child of"+chr(9)+"Is the Grandchild of"+chr(9)+"Is the Guardian of"+chr(9)+"Is the Nephew of"+chr(9)+"Is the Niece of"+chr(9)+"Is the Parent of"+chr(9)+"Is the Sibling of"+chr(9)+"Is the Spouse of"+chr(9)+"Is the Step Child of"+chr(9)+"Is Unrelated to"+chr(9)+"Other", Pare_Line_Array(relationship_type, pare_item)
					EditBox 205, (20 + (dialog_count * 20)), 85, 15, Pare_Line_Array(other_relationship_list, pare_item)
					DropListBox 295, (20 + (dialog_count * 20)), 70, 15, hh_member_dropdown, Pare_Line_Array(relationship_to, pare_item)
					EditBox 375, (20 + (dialog_count * 20)), 105, 15, Pare_Line_Array(documents_received, pare_item)
					CheckBox 490, (25 + (dialog_count * 20)), 75, 10, "New/Updated Proof", Pare_Line_Array(new_checkbox, pare_item)
					dialog_count = dialog_count + 1
				ElseIf add_field_count > 0 and pare_item = UBound(Pare_Line_Array, 2) Then
					DropListBox 5, (20 + (dialog_count * 20)), 70, 15, hh_member_dropdown, Pare_Line_Array(received_for, pare_item)
					DropListBox 85, (20 + (dialog_count * 20)), 110, 15, "Select one..."+chr(9)+"Is Another Relative of"+chr(9)+"Is the Child of"+chr(9)+"Is the Foster Child of"+chr(9)+"Is the Grandchild of"+chr(9)+"Is the Guardian of"+chr(9)+"Is the Nephew of"+chr(9)+"Is the Niece of"+chr(9)+"Is the Parent of"+chr(9)+"Is the Sibling of"+chr(9)+"Is the Spouse of"+chr(9)+"Is the Step Child of"+chr(9)+"Is Unrelated to"+chr(9)+"Other", Pare_Line_Array(relationship_type, pare_item)
					EditBox 205, (20 + (dialog_count * 20)), 85, 15, Pare_Line_Array(other_relationship_list, pare_item)
					DropListBox 295, (20 + (dialog_count * 20)), 70, 15, hh_member_dropdown, Pare_Line_Array(relationship_to, pare_item)
					EditBox 375, (20 + (dialog_count * 20)), 105, 15, Pare_Line_Array(documents_received, pare_item)
					CheckBox 490, (25 + (dialog_count * 20)), 75, 10, "New/Updated Proof", Pare_Line_Array(new_checkbox, pare_item)
					dialog_count = dialog_count + 1
				End If
			Next
		ButtonGroup ButtonPressed
			PushButton 365, (35 + (selected_hh_memb_relationship_count + add_field_count) * 20), 200, 15, "Click Here to add another Relationship Line", add_another_button
		EditBox 90, (55 + ((selected_hh_memb_relationship_count + add_field_count) * 20)), 170, 15, other_verifs_needed
		EditBox 310, (55 + ((selected_hh_memb_relationship_count + add_field_count) * 20)), 235, 15, other_notes
		CheckBox 70, (80 + ((selected_hh_memb_relationship_count + add_field_count) * 20)), 30, 10, "PARE", pare_checkbox
		CheckBox 105, (80 + ((selected_hh_memb_relationship_count + add_field_count) * 20)), 35, 10, "MEMB", memb_checkbox
		CheckBox 145, (80 + ((selected_hh_memb_relationship_count + add_field_count) * 20)), 30, 10, "ABPS", abps_checkbox
		CheckBox 180, (80 + ((selected_hh_memb_relationship_count + add_field_count) * 20)), 25, 10, "SIBL", sibl_checkbox
		CheckBox 210, (80 + ((selected_hh_memb_relationship_count + add_field_count) * 20)), 30, 10, "Other:", other_checkbox
		EditBox 245, (75 + ((selected_hh_memb_relationship_count + add_field_count) * 20)), 30, 15, other_option
		EditBox 340, (75 + ((selected_hh_memb_relationship_count + add_field_count) * 20)), 105, 15, worker_signature
		ButtonGroup ButtonPressed
			OkButton 455, (75 + ((selected_hh_memb_relationship_count + add_field_count) * 20)), 50, 15
			CancelButton 510, (75 + ((selected_hh_memb_relationship_count + add_field_count) * 20)), 50, 15
		Text 5, 10, 75, 10, "Member received for: "
		Text 85, 10, 50, 10, "Relationship: "
		Text 205, 10, 90, 10, "If Other, list relationship:"
		Text 295, 10, 50, 10, "Related to:"
		Text 375, 10, 75, 10, "Document(s) received: "
		Text 5, (60 + ((selected_hh_memb_relationship_count + add_field_count) * 20)), 80, 10, "Other verifs still needed:"
		Text 265, (60 + ((selected_hh_memb_relationship_count + add_field_count) * 20)), 45, 10, "Other Notes:"
		Text 5, (80 + ((selected_hh_memb_relationship_count + add_field_count) * 20)), 60, 10, "Panel(s) updated: "
		Text 280, (80 + ((selected_hh_memb_relationship_count + add_field_count) * 20)), 60, 10, "Worker signature:"
		' Text 5, (35 + (array_item * 20)), 575, 10, "  * This last line is available for entry of relationship proof that was not documented in STAT. If the Relationship is left as 'SELECT ONE...' this line will not case note."
		EndDialog

    	
        'Dialog Box to list members and documentation received.
        'This dialog is here instead of the beginning because the dynamic thing only works if the array items are set before the dialog is defined
    	Dialog Dialog1
    	cancel_confirmation
    	For relationship = 0 to UBound (Pare_Line_Array,2)
            If Pare_Line_Array(relationship_type, relationship) = "Select one..." Then err_msg = err_msg & vbNewLine & "* Type of relationship between " & Pare_Line_Array(received_for, relationship) & " and " & Pare_Line_Array(relationship_to, relationship) & " must be explained."
    		If Pare_Line_Array(relationship_type, relationship) = "Other" AND Pare_Line_Array (other_relationship_list, relationship) = "" Then err_msg = err_msg & vbCr & _
    		  "You must define the *other* relationship between Memb " & Pare_Line_Array(received_for, relationship) & " and Memb " & Pare_Line_Array(relationship_to, relationship)			'Requires other relationship to be explained
            If Pare_Line_Array(relationship_type, relationship) = "Is Another Relative of" AND Pare_Line_Array (other_relationship_list, relationship) = "" Then err_msg = err_msg & vbCr & _
    		  "You must define the *other* relationship between Memb " & Pare_Line_Array(received_for, relationship) & " and Memb " & Pare_Line_Array(relationship_to, relationship)			'Requires other relationship to be explained
            If Pare_Line_Array(received_for, relationship) = Pare_Line_Array(relationship_to, relationship) Then err_msg = err_msg & vbNewLine & "* The 'Member Received For' and 'Related to' are the same member, you do not need to indicate a relationship to self. For Memb " & Pare_Line_Array(received_for, relationship) & "."
            If Pare_Line_Array(new_checkbox,relationship) = checked Then new_proof_exists = TRUE
            If Pare_Line_Array(new_checkbox,relationship) = unchecked Then old_proof_exists = TRUE
    	Next
    	IF other_checkbox = 1 AND other_option = "" THEN err_msg = err_msg & vbCr & "You must list a panel if Other is selected."			'Requires Other panel to be specified
    	IF worker_signature = "" THEN err_msg = err_msg & vbCr & "You must enter a worker signature."										'Requires worker signgnature
        If ButtonPressed = add_another_button Then
            err_msg = "LOOP" & err_msg
			add_field_count = add_field_count + 1
            array_item = array_item + 1
            ReDim Preserve Pare_Line_Array(5, array_item)
        End If
        If err_msg <> "" AND left(err_msg, 4) <> "LOOP" Then MsgBox "Please resolve to continue" & vbCr & vbCr & err_msg													'Displays the error message to the worker
    Loop until err_msg = ""
CALL check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = False

'Statements needed for the check boxes for panels updated, defined further in case notes below.
IF memb_checkbox =  1 THEN memb  = "MEMB/"
IF pare_checkbox =  1 THEN pare  = "PARE/"
IF abps_checkbox =  1 THEN abps  = "ABPS/"
IF sibl_checkbox =  1 THEN sibl  = "SIBL/"
IF other_checkbox = 1 THEN other = "Other: "
If memb = "" AND pare = "" AND abps = "" AND sibl = "" AND other = "" Then
	panels_updated = FALSE
Else
	panels_updated = TRUE
End If

STATS_counter = STATS_counter - 1 														'Remove one instance of the stats counter since it starts at 1

'Information for the case note.
start_a_blank_case_note
Call write_variable_in_CASE_NOTE("Documentation Received: Proof of Relationship")		'Case note heading
If new_proof_exists = TRUE Then 														'Seperates the new/updated proofs to the top of the case note
	Call write_variable_in_CASE_NOTE("New Relationships Verified:")						'Subheading for the new proofs
	For pare_item = 0 to UBound(Pare_Line_Array,2)
		If Pare_Line_Array(new_checkbox,pare_item) = checked AND Pare_Line_Array(relationship_type, pare_item) <> "Select one..." Then		'Listing all the items in the array with new/updated seleceted
			If Pare_Line_Array(other_relationship_list,pare_item) <> "" Then 			'If other relationship type is listed, the formate of the line is a little different
				Call write_variable_in_CASE_NOTE("* Relationship: Memb " & Pare_Line_Array(received_for,pare_item)& " - " & Pare_Line_Array(relationship_type,pare_item) & " - Memb " & Pare_Line_Array(relationship_to,pare_item) & ": " & Pare_Line_Array(other_relationship_list,pare_item) & ". Doc Rec'vd: " & Pare_Line_Array(documents_received,pare_item))
			Else
				Call write_variable_in_CASE_NOTE("* Relationship: Memb " & Pare_Line_Array(received_for,pare_item)& " - " & Pare_Line_Array(relationship_type,pare_item) & " - Memb " & Pare_Line_Array(relationship_to,pare_item) & ". Doc Rec'vd: " & Pare_Line_Array(documents_received,pare_item))
			End If
		End If
	Next
    Call write_variable_in_CASE_NOTE("---")
End If
If old_proof_exists = TRUE Then
    If new_proof_exists = TRUE Then 														'Subheading for the relationships that are not new/updated - this is slightly different depending on if any new proofs were listed
    	Call write_variable_in_CASE_NOTE("Relationships already known/verfied: ")
    Else
    	Call write_variable_in_CASE_NOTE("Household Relationships known/documented: ")
    End If
    For pare_item = 0 to UBound(Pare_Line_Array,2)											'Lists all the relationshps that are not marked as new/updated AND have an actual relationship type selected. - not 'select one'
    	If Pare_Line_Array(new_checkbox,pare_item) = unchecked AND Pare_Line_Array(relationship_type, pare_item) <> "Select one..." Then
    		If Pare_Line_Array(other_relationship_list,pare_item) <> "" Then
    			Call write_variable_in_CASE_NOTE("* Relationship: Memb " & Pare_Line_Array(received_for,pare_item)& " & Memb " & Pare_Line_Array(relationship_to,pare_item) & " are " & Pare_Line_Array(other_relationship_list,pare_item) & ". Doc Rec'vd: " & Pare_Line_Array(documents_received,pare_item))
    		Else
    			Call write_variable_in_CASE_NOTE("* Relationship: Memb " & Pare_Line_Array(received_for,pare_item)& " - " & Pare_Line_Array(relationship_type,pare_item) & " - Memb " & Pare_Line_Array(relationship_to,pare_item) & ". Doc Rec'vd: " & Pare_Line_Array(documents_received,pare_item))
    		End If
    	End If
    Next
    Call write_variable_in_CASE_NOTE("---")
End If
Call write_bullet_and_variable_in_CASE_NOTE("Verifs Needed", other_verifs_needed)		'Adding other verifs to the case note
Call write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)					'Adding other notes to the case nore
IF panels_updated = TRUE Then Call write_bullet_and_variable_in_CASE_NOTE("Panel(s) updated", memb & pare & abps & sibl & other & other_option)	'Adding list of panels updated
If other_verifs_needed <> "" OR other_notes <> "" OR panels_updated = TRUE Then Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure_with_error_report("Success! Relationship detail documented.")
