BeginDialog change_reported_dialog, 0, 0, 136, 105, "Change Reported"
  EditBox 70, 5, 35, 15, MAXIS_case_number
  EditBox 70, 25, 15, 15, MAXIS_footer_month
  EditBox 90, 25, 15, 15, MAXIS_footer_year
  DropListBox 20, 65, 85, 15, "Select One:"+chr(9)+"Address "+chr(9)+"Baby Born"+chr(9)+"HHLD Comp"+chr(9)+"Income "+chr(9)+"Shelter Cost "+chr(9)+"Other(please specify)", nature_change
  ButtonGroup ButtonPressed
    OkButton 45, 85, 40, 15
    CancelButton 90, 85, 40, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 30, 65, 10, "Footer month/year: "
  Text 5, 50, 130, 10, "Please select the nature of the change:"
EndDialog

BeginDialog HHLD_Comp_Change_Dialog, 0, 0, 161, 200, "Household Comp Change"
  EditBox 80, 5, 20, 15, HH_member
  EditBox 80, 25, 35, 15, date_reported
  EditBox 80, 45, 35, 15, effective_date
  CheckBox 15, 75, 90, 10, "Verifications sent to ECF", Verif_checkbox
  CheckBox 15, 85, 80, 10, "Updated STAT panels", STAT_checkbox
  CheckBox 15, 95, 80, 10, "Approved new results", APP_checkbox
  CheckBox 15, 105, 80, 10, "Notified other agency", notify_checkbox
  EditBox 50, 125, 100, 15, additional_notes
  EditBox 50, 145, 100, 15, worker_signature
  CheckBox 5, 165, 125, 10, "Check if the change is temporary", temporary_change_checkbox
  ButtonGroup ButtonPressed
    OkButton 65, 180, 40, 15
    CancelButton 110, 180, 40, 15
  Text 5, 10, 75, 10, "Member # HH change:"
  Text 30, 50, 50, 10, "Effective date:"
  Text 5, 130, 45, 10, "Other Notes:"
  GroupBox 5, 65, 145, 55, "Action Taken"
  Text 30, 30, 50, 10, "Date reported:"
  Text 5, 150, 40, 10, "Worker Sig:"
EndDialog



'this creates the client array for dropdown list
CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.
DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
	EMReadscreen ref_nbr, 2, 4, 33
	EMReadscreen last_name_array, 25, 6, 30								'took out clients last name apparently may be too much characters within the form restrictions.
	EMReadscreen first_name_array, 12, 6, 63
	last_name_array = replace(last_name_array, "_", "")
	last_name_array = Lcase(last_name_array)
	last_name_array = UCase(Left(last_name_array, 1)) &  Mid(last_name_array, 2)     	'took out clients last name apparently may be too much characters within the form restrictions.
	first_name_array = replace(first_name_array, "_", "") '& " "
	first_name_array = Lcase(first_name_array)
	first_name_array = UCase(Left(first_name_array, 1)) &  Mid(first_name_array, 2)
	client_string =  "MEMB " & ref_nbr & " - " & first_name_array & " " & last_name_array
	client_array = client_array & client_string & "|"
	transmit
	Emreadscreen edit_check, 7, 24, 2
LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.
client_array = TRIM(client_array)
test_array = split(client_array, "|")
total_clients = Ubound(test_array)			'setting the upper bound for how many spaces to use from the array
DIM all_client_array()
ReDim all_clients_array(total_clients, 1)
FOR clt_x = 0 to total_clients				'using a dummy array to build list into the array used for the dialog.
	Interim_array = split(client_array, "|")
	all_clients_array(clt_x, 0) = Interim_array(clt_x)
	all_clients_array(clt_x, 1) = 1
NEXT
HH_member_array = ""
FOR i = 0 to total_clients
	IF all_clients_array(i, 0) <> "" THEN 						'creates the final array to be used by other scripts.
		IF all_clients_array(i, 1) = 1 THEN						'if the person/string has been checked on the dialog then the reference number portion (left 2) will be added to new HH_member_array
			HH_member_array = chr(9) & HH_member_array & chr(9) & all_clients_array(i, 0)
		END IF
	END IF
NEXT
'removes all of the first 'chr(9)'
HH_member_array_dialog = Right(HH_member_array, len(HH_member_array) - total_clients)

Row = 15
DO
    EMReadScreen child_ref_number, 2, row, 35
msgbox  child_ref_number_I
EMReadScreen child_ref_number_II, 2, 16, 35
msgbox  child_ref_number_II
EMReadScreen child_ref_number_III, 2, 17, 35
msgbox  child_ref_number_III
IF child_ref_number_III <> "" THEN
    PF18 ' shift PF8 look into the function lib PF19 is shift f8' Pf20 is shift f8'
    PF18
    EMReadScreen child_ref_number_IV, 2, 15, 35
    msgbox  child_ref_number_IV
    EMReadScreen child_ref_number_V, 2, 16, 35
    EMReadScreen child_ref_number_VI, 2, 17, 35
    TRANSMIT
    msgbox "where am i checking ref number"
END IF

start_a_blank_case_note
'writes case note for Baby Born
IF nature_change = "Baby Born" THEN
	CALL write_variable_in_Case_Note("--CHANGE REPORTED - Client reports birth of baby--")
	CALL write_bullet_and_variable_in_Case_Note("Child's's name", babys_name)
	If baby_gender = "Select One:" then									'gender will be listed as unknown if not updated'
		CALL write_bullet_and_variable_in_Case_Note("Gender", "unknown")
	Else
		CALL write_bullet_and_variable_in_Case_Note("Gender", baby_gender)
	End If
	CALL write_bullet_and_variable_in_Case_Note("Date of birth", date_of_birth)
	father_HH = " - not reported in the same household"
	If parent_in_household = "Yes" Then father_HH = " - reported in the same household."
	If fathers_name = "" then fathers_name = "Unknown or not provided"
	CALL write_bullet_and_variable_in_Case_Note("Mother's name", mothers_name)
	CALL write_bullet_and_variable_in_Case_Note("Mother's employer", mothers_employer)
	CALL write_bullet_and_variable_in_Case_Note("Father's name", fathers_name & father_HH)
	CALL write_bullet_and_variable_in_Case_Note("Father's employer", fathers_employer)
	IF other_health_insurance = "Yes" THEN CALL write_bullet_and_variable_in_Case_Note("OHI", OHI_Source)
	IF MHC_plan_checkbox = CHECKED THEN CALL write_variable_in_CASE_NOTE("* Newborns MHC plan updated to match the mothers.")
	CALL write_bullet_and_variable_in_Case_Note("Other Notes", other_notes)
END IF

'writes case note for HHLD Comp Change
IF nature_change = "HHLD Comp Change" THEN
	CALL write_variable_in_case_note("--CHANGE REPORTED - HH Comp Change--")
	CALL write_bullet_and_variable_in_Case_Note("Unit member HH Member", HH_Member)
	CALL write_bullet_and_variable_in_Case_Note("Date Reported/Addendum", date_reported)
	CALL write_bullet_and_variable_in_Case_Note("Date Effective", effective_date)
	IF Temporary_Change_Checkbox = CHECKED THEN CALL write_variable_in_Case_Note("***Change is temporary***")
	IF Temporary_Change_Checkbox = UNCHECKED THEN CALL write_variable_in_Case_Note("***Change is NOT temporary***")
END IF


BeginDialog baby_born_dialog, 0, 0, 186, 265, "BABY BORN"
  EditBox 55, 5, 115, 15, babys_name
  EditBox 55, 25, 40, 15, date_of_birth
  DropListBox 130, 25, 40, 15, "Select One:"+chr(9)+"Male"+chr(9)+"Female", baby_gender
  DropListBox 100, 45, 40, 15, "Select One:"+chr(9)+"Yes"+chr(9)+"No", parent_in_household
  DropListBox 85, 75, 80, 15, "Select One:" & (HH_member_array_dialog), mothers_name
  EditBox 85, 95, 80, 15, mothers_employer
  EditBox 80, 130, 85, 15, fathers_name
  EditBox 80, 150, 85, 15, fathers_employer
  CheckBox 10, 170, 165, 10, "Newborns MHC plan updated to mother's carrier", MHC_plan_checkbox
  DropListBox 140, 185, 40, 15, "Select One:"+chr(9)+"Yes"+chr(9)+"No", other_health_insurance
  EditBox 110, 205, 70, 15, OHI_source
  EditBox 50, 225, 130, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 95, 245, 40, 15
    CancelButton 140, 245, 40, 15
  Text 5, 30, 45, 10, "Date of birth:"
  Text 100, 30, 25, 10, "Gender:"
  Text 5, 50, 95, 10, "Other parent in household?"
  Text 15, 135, 50, 10, "Fathers Name:"
  Text 5, 10, 50, 10, "Child's name:"
  Text 15, 155, 65, 10, "Father's Employer:"
  Text 5, 230, 45, 10, "Other Notes:"
  Text 5, 210, 105, 10, "If yes to OHI, source of the OHI:"
  Text 55, 190, 80, 10, "Other Health Insurance?"
  Text 15, 80, 65, 10, "Mother of Newborn: "
  Text 15, 100, 65, 10, "Mother's Employer: "
  GroupBox 5, 120, 175, 50, "Father's Information"
  GroupBox 5, 65, 175, 50, "Mother's Information"
EndDialog
