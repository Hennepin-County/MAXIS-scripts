'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - CAF Addendum.vbs"
start_time = timer
STATS_counter = 0                          'sets the stats counter at one
STATS_manualtime = 540                     'manual run time in seconds
STATS_denomination = "M"                   'C is for each CASE
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
Call changelog_update("10/22/2020", "Removed the default date for the addendum date as the functionality was pulling incorrect data and there is not a reliable way to know which date is accurate. Since this is a form date, it should be manually entered.", "Casey Love, Hennepin County")
Call changelog_update("12/21/2019", "Added 'NB' to the list of former states.", "Casey Love, Hennepin County")
Call changelog_update("09/25/2019", "Bug Fix - Verifs Needed was creating possible multiple case notes and noting when nothing was added. Also a typo in the case note wording.", "Casey Love, Hennepin County")
call changelog_update("09/12/2019", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
function reset_variables()
    new_member_to_note = "New Member not in MAXIS"
    new_memb_ref_numb = ""
    new_memb_first_name = ""
    new_memb_middle_name = ""
    new_memb_last_name = ""
    new_memb_suffix = "Select or Type"
    new_memb_sex = ""
    new_memb_dob = ""
    new_memb_marital_status = ""
    new_memb_ssn = ""
    new_memb_race_n = unchecked
    new_memb_race_a = unchecked
    new_memb_race_b = unchecked
    new_memb_race_p = unchecked
    new_memb_race_w = unchecked
    new_memb_hispanic = "No"
    new_memb_move_to_MN = ""
    new_memb_last_grade = ""
    new_memb_app_snap = unchecked
    new_memb_app_cash = unchecked
    new_memb_app_emer = unchecked
    new_memb_app_none = unchecked
    new_memb_p_p_tog = "Yes"
    new_memb_relationship = "Select or Type"
    citizen_yn = "Yes"
    disa_yn = "No"
    unable_to_work_yn = "No"
    school_yn = "No"
    assets_yn = "No"
    unea_yn = "No"
    earned_yn = "No"
    expenses_yn = "No"

    us_entry_date = ""
    nationality = ""
    immig_status_dropdown = ""
    imig_verif_checkbox = unchecked
    sponsor_name = ""
    sponsor_address = ""
    sponsor_phone = ""
    sponsor_verif_checkbox = unchecked
    city_moved_from = ""
    state_moved_from = "Select One..."
    to_date = ""
    from_date = ""
    medical_problem = ""
    doctors_info = ""
    disa_verif_checkbox = unchecked
    reason_unable_to_work = ""
    parent_one_name = ""
    parent_absent_or_not_one = "Parent in hone?"
    parent_one_address = ""
    custody_share_one = unchecked
    parent_two_name = ""
    parent_absent_or_not_two = "Parent in home?"
    parent_two_address = ""
    custody_share_two = unchecked
    school_name = ""
    school_address = ""
    school_verif_checkbox = unchecked
    asset_type_one = ""
    asset_value_one = ""
    asset_owed_one = ""
    asset_one_verif_checkbox = unchecked
    asset_type_two = ""
    asset_value_two = ""
    asset_owed_two = ""
    asset_two_verif_checkbox = unchecked
    unea_type  = ""
    unea_amount = ""
    unea_frequency = "Select or Type"
    unea_verif_checkbox = unchecked
    employer_name_one = ""
    hrs_per_wk_one = ""
    employer_amount_one = ""
    employer_frequency_one = ""
    employer_verif_one_checkbox = unchecked
    employer_name_two = ""
    hrs_per_wk_two = ""
    employer_amount_two = ""
    employer_frequency_two = ""
    employer_verif_two_checkbox = unchecked
    expense_type = ""
    expense_amount = ""
    expense_verif_checkbox = unchecked
end function


state_list = "Select One..."
state_list = state_list+chr(9)+"NB MN Newborn"
state_list = state_list+chr(9)+"FC Foreign Country"
state_list = state_list+chr(9)+"UN Unknown"
state_list = state_list+chr(9)+"AL Alabama"
state_list = state_list+chr(9)+"AK Alaska"
state_list = state_list+chr(9)+"AZ Arizona"
state_list = state_list+chr(9)+"AR Arkansas"
state_list = state_list+chr(9)+"CA California"
state_list = state_list+chr(9)+"CO Colorado"
state_list = state_list+chr(9)+"CT Connecticut"
state_list = state_list+chr(9)+"DE Delaware"
state_list = state_list+chr(9)+"DC District Of Columbia"
state_list = state_list+chr(9)+"FL Florida"
state_list = state_list+chr(9)+"GA Georgia"
state_list = state_list+chr(9)+"HI Hawaii"
state_list = state_list+chr(9)+"ID Idaho"
state_list = state_list+chr(9)+"IL Illnois"
state_list = state_list+chr(9)+"IN Indiana"
state_list = state_list+chr(9)+"IA Iowa"
state_list = state_list+chr(9)+"KS Kansas"
state_list = state_list+chr(9)+"KY Kentucky"
state_list = state_list+chr(9)+"LA Louisiana"
state_list = state_list+chr(9)+"ME Maine"
state_list = state_list+chr(9)+"MD Maryland"
state_list = state_list+chr(9)+"MA Massachusetts"
state_list = state_list+chr(9)+"MI Michigan"
state_list = state_list+chr(9)+"MS Mississippi"
state_list = state_list+chr(9)+"MO Missouri"
state_list = state_list+chr(9)+"MT Montana"
state_list = state_list+chr(9)+"NE Nebraska"
state_list = state_list+chr(9)+"NV Nevada"
state_list = state_list+chr(9)+"NH New Hampshire"
state_list = state_list+chr(9)+"NJ New Jersey"
state_list = state_list+chr(9)+"NM New Mexico"
state_list = state_list+chr(9)+"NY New York"
state_list = state_list+chr(9)+"NC North Carolina"
state_list = state_list+chr(9)+"ND North Dakota"
state_list = state_list+chr(9)+"OH Ohio"
state_list = state_list+chr(9)+"OK Oklahoma"
state_list = state_list+chr(9)+"OR Oregon"
state_list = state_list+chr(9)+"PA Pennsylvania"
state_list = state_list+chr(9)+"RI Rhode Island"
state_list = state_list+chr(9)+"SC South Carolina"
state_list = state_list+chr(9)+"SD South Dakota"
state_list = state_list+chr(9)+"TN Tennessee"
state_list = state_list+chr(9)+"TX Texas"
state_list = state_list+chr(9)+"UT Utah"
state_list = state_list+chr(9)+"VT Vermont"
state_list = state_list+chr(9)+"VA Virginia"
state_list = state_list+chr(9)+"WA Washington"
state_list = state_list+chr(9)+"WV West Virginia"
state_list = state_list+chr(9)+"WI Wisconsin"
state_list = state_list+chr(9)+"WY Wyoming"
state_list = state_list+chr(9)+"PR Puerto Rico"
state_list = state_list+chr(9)+"VI Virgin Islands"

state_array = split(state_list, chr(9))

Dim new_member_to_note, new_memb_ref_numb, new_memb_first_name, new_memb_middle_name, new_memb_last_name, new_memb_suffix, new_memb_sex, new_memb_dob, new_memb_marital_status, new_memb_ssn, new_memb_race_n
Dim new_memb_race_a, new_memb_race_b, new_memb_race_p, new_memb_race_w, new_memb_hispanic, new_memb_move_to_MN, new_memb_last_grade, new_memb_app_snap, new_memb_app_cash, new_memb_app_emer, new_memb_app_none
Dim new_memb_p_p_tog, new_memb_relationship, citizen_yn, disa_yn, unable_to_work_yn, school_yn, assets_yn, unea_yn, earned_yn, expenses_yn, us_entry_date, nationality, immig_status_dropdown, imig_verif_checkbox
Dim sponsor_name, sponsor_address, sponsor_phone, sponsor_verif_checkbox, city_moved_from, state_moved_from, to_date, from_date, medical_problem, doctors_info, disa_verif_checkbox, reason_unable_to_work
Dim parent_one_name, parent_absent_or_not_one, parent_one_address, custody_share_one, parent_two_name, parent_absent_or_not_two, parent_two_address, custody_share_two, school_name, school_address, school_verif_checkbox
Dim asset_type_one, asset_value_one, asset_owed_one, asset_one_verif_checkbox, asset_type_two, asset_value_two, asset_owed_two, asset_two_verif_checkbox, unea_type, unea_amount, unea_frequency, unea_verif_checkbox
Dim employer_name_one, hrs_per_wk_one, employer_amount_one, employer_frequency_one, employer_verif_one_checkbox, employer_name_two, hrs_per_wk_two, employer_amount_two, employer_frequency_two, employer_verif_two_checkbox
Dim expense_type, expense_amount, expense_verif_checkbox

'The SCRIPT ================================================================================================================
EMConnect ""
get_county_code				'since there is a county specific checkbox, this makes the the county clear
If MAXIS_case_number = "" Then          'This is sometimes run from another script (NOTES - CAF) in that case the case number and footer month and year are already known.
    Call MAXIS_case_number_finder(MAXIS_case_number)
    Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
End If
script_run_lowdown = ""

'This code is commented out because it needs review if we are rewriting.
'The purpose here is to autofill the date the addendum was received as it is entered on the ADME panel. However this code goes to the first ADME panel in the case and has no functionality
'to determine if it is the most recent, or newest member.
'Future stat this could be adjusted to read ALL the ADME panels and find the most recent date to autofil the date field, currently we do not have time to build this functionality.
' If MAXIS_case_number <> "" Then
'     Call Navigate_to_MAXIS_screen("STAT", "ADME")
'     EMReadScreen cash_adme_date, 8, 12, 38
'     EMReadScreen snap_adme_date, 8, 16, 38
'
'     cash_adme_date = replace(cash_adme_date, " ", "/")
'     snap_adme_date = replace(snap_adme_date, " ", "/")
'
'     If IsDate(snap_adme_date) = TRUE Then addendum_date = snap_adme_date
'     If IsDate(cash_adme_date) = TRUE Then addendum_date = cash_adme_date
' End If
how_many_new_members = "1"

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 266, 130, "Case Number Dialog"
  EditBox 60, 35, 50, 15, MAXIS_case_number
  EditBox 180, 35, 50, 15, addendum_date
  EditBox 135, 55, 35, 15, how_many_new_members
  CheckBox 20, 75, 130, 10, "Check here if client signed Addendum.", addendum_signed
  EditBox 85, 90, 170, 15, worker_signature
  ButtonGroup ButtonPressed
	OkButton 155, 110, 50, 15
	CancelButton 210, 110, 50, 15
  Text 10, 10, 250, 20, "A CAF Addendum (DHS-5223C) was received to add a new person to a case. This script will create a CASE/NOTE after processing the addendum."
  Text 10, 40, 50, 10, "Case Number:"
  Text 120, 40, 55, 10, "Addendum Date:"
  Text 10, 60, 120, 10, "Number of New Members Reported:"
  Text 10, 95, 70, 10, "Worker Signature:"
EndDialog
Do
    Do
	    err_msg = ""
        dialog Dialog1
        cancel_without_confirmation
        Call validate_MAXIS_case_number(err_msg, "*")
        If IsDate(addendum_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the date the Addendum Form was received."
        If IsNumeric(how_many_new_members) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the number of new household members listed on the CAF Addendum."
        If err_msg <> "" Then MsgBox("Please resolve to continue:" & vbNewLine & err_msg)
    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

MAXIS_footer_month = DatePart("m", addendum_date)
MAXIS_footer_month = right("0" & MAXIS_footer_month, 2)
MAXIS_footer_year = DatePart("yyyy", addendum_date)
MAXIS_footer_year = right(MAXIS_footer_year, 2)

how_many_new_members = how_many_new_members * 1
Call generate_client_list(all_clients_list, "New Member not in MAXIS")
Call generate_client_list(full_clients_list, "Select or Type")
client_array = split(full_clients_list, chr(9))

For the_member = 1 to how_many_new_members
    reset_variables
	'-------------------------------------------------------------------------------------------------DIALOG
	Dialog1 = "" 'Blanking out previous dialog detail
	'Dialog to ask if there is a member number already added for this member.'
	BeginDialog Dialog1, 0, 0, 191, 85, "Select New HH Member"
	  DropListBox 10, 45, 155, 45, all_clients_list, new_member_to_note
	  ButtonGroup ButtonPressed
		OkButton 135, 65, 50, 15
	  Text 10, 10, 160, 25, "If the household member has already been added to MAXIS, select the member's reference number to have the script fill SOME information."
	EndDialog
    Do
	    dialog Dialog1
        cancel_confirmation
        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE

    If new_member_to_note <> "New Member not in MAXIS" Then
        new_memb_ref_numb = left(new_member_to_note, 2)
        Call navigate_to_MAXIS_screen("STAT", "MEMB")
        EMWriteScreen new_memb_ref_numb, 20, 76
        transmit
        EMReadScreen new_memb_first_name, 12, 6, 63
        EMReadScreen new_memb_middle_name, 1, 6, 79
        EMReadScreen new_memb_last_name, 25, 6, 30
        EMReadScreen new_memb_sex, 1, 9, 42
        EMReadScreen new_memb_dob, 10, 8, 42
        EMReadScreen new_memb_ssn, 11, 7, 42
        EMReadScreen new_memb_race_entry, 35, 17, 42
        EMReadScreen new_memb_hispanic, 1, 16, 68
        EMReadScreen new_memb_relationship, 2, 10, 42
        new_memb_first_name = replace(new_memb_first_name, "_", "")
        new_memb_middle_name = replace(new_memb_middle_name, "_", "")
        new_memb_last_name = replace(new_memb_last_name, "_", "")
        If new_memb_sex = "M" Then new_memb_sex = "Male"
        If new_memb_sex = "F" Then new_memb_sex = "Female"
        new_memb_dob = replace(new_memb_dob, " ", "/")
        new_memb_ssn = replace(new_memb_ssn, " ", "-")
        new_memb_race_entry = trim(new_memb_race_entry)
        If new_memb_race_entry = "Asian" Then new_memb_race_a = checked
        If new_memb_race_entry = "Amer Indn Or Alaskan Native" Then new_memb_race_n = checked
        If new_memb_race_entry = "Black Or African Amer" Then new_memb_race_b = checked
        If new_memb_race_entry = "Pacific Is Or Native Hawaii" Then new_memb_race_p = checked
        If new_memb_race_entry = "White" Then new_memb_race_w = checked
        If new_memb_hispanic = "N" Then new_memb_hispanic = "No"
        If new_memb_hispanic = "Y" Then new_memb_hispanic = "Yes"
        If new_memb_relationship = "02" Then new_memb_relationship = "02 - Spouse"
        If new_memb_relationship = "03" Then new_memb_relationship = "03 - Child"
        If new_memb_relationship = "04" Then new_memb_relationship = "04 - Parent"
        If new_memb_relationship = "05" Then new_memb_relationship = "05 - Sibling"
        If new_memb_relationship = "06" Then new_memb_relationship = "06 - Step Sibling"
        If new_memb_relationship = "08" Then new_memb_relationship = "08 - Step Child"
        If new_memb_relationship = "09" Then new_memb_relationship = "09 - Step Parent"
        If new_memb_relationship = "10" Then new_memb_relationship = "10 - Aunt"
        If new_memb_relationship = "11" Then new_memb_relationship = "11 - Uncle"
        If new_memb_relationship = "12" Then new_memb_relationship = "12 - Niece"
        If new_memb_relationship = "13" Then new_memb_relationship = "13 - Nephew"
        If new_memb_relationship = "14" Then new_memb_relationship = "14 - Cousin"
        If new_memb_relationship = "15" Then new_memb_relationship = "15 - Grandparent"
        If new_memb_relationship = "16" Then new_memb_relationship = "16 - Grandchild"
        If new_memb_relationship = "17" Then new_memb_relationship = "17 - Other Relative"
        If new_memb_relationship = "18" Then new_memb_relationship = "18 - Legal Guardian"
        If new_memb_relationship = "24" Then new_memb_relationship = "24 - Not Related"
        If new_memb_relationship = "25" Then new_memb_relationship = "25 - Live-In Attendant"
        If new_memb_relationship = "27" Then new_memb_relationship = "27 - Unknown/Not Indc"

        If new_memb_race_entry = "" Then
        End If

        Call navigate_to_MAXIS_screen("STAT", "MEMI")
        EMWriteScreen new_memb_ref_numb, 20, 76
        transmit
        EMReadScreen new_memb_marital_status, 1, 7, 40
        EMReadScreen new_memb_move_to_MN, 8, 15, 49
        EMReadScreen new_memb_last_grade, 2, 10, 49

        If new_memb_marital_status = "N" Then new_memb_marital_status = "N - Never maried"
        If new_memb_marital_status = "M" Then new_memb_marital_status = "M - Married, living w/ spouse"
        If new_memb_marital_status = "S" Then new_memb_marital_status = "S - Seperated (married, living apart)"
        If new_memb_marital_status = "L" Then new_memb_marital_status = "L - Legally seperated"
        If new_memb_marital_status = "D" Then new_memb_marital_status = "D - Divorced"
        If new_memb_marital_status = "W" Then new_memb_marital_status = "W - Widowed"
        new_memb_move_to_MN = replace(new_memb_move_to_MN, " ", "/")
        If new_memb_move_to_MN = "__/__/__" Then new_memb_move_to_MN = ""
        If new_memb_last_grade = "00" Then new_memb_last_grade = "00  Pre 1st or Never"
        If new_memb_last_grade = "01" Then new_memb_last_grade = "01 Grade 1"
        If new_memb_last_grade = "02" Then new_memb_last_grade = "02 Grade 2"
        If new_memb_last_grade = "03" Then new_memb_last_grade = "03 Grade 3"
        If new_memb_last_grade = "04" Then new_memb_last_grade = "04 Grade 4"
        If new_memb_last_grade = "05" Then new_memb_last_grade = "05 Grade 5"
        If new_memb_last_grade = "06" Then new_memb_last_grade = "06 Grade 6"
        If new_memb_last_grade = "07" Then new_memb_last_grade = "07 Grade 7"
        If new_memb_last_grade = "08" Then new_memb_last_grade = "08 Grade 8"
        If new_memb_last_grade = "09" Then new_memb_last_grade = "09 Grade 9"
        If new_memb_last_grade = "10" Then new_memb_last_grade = "10 Grade 10"
        If new_memb_last_grade = "11" Then new_memb_last_grade = "11 Grade 11"
        If new_memb_last_grade = "12" Then new_memb_last_grade = "12 HS or GED"
        If new_memb_last_grade = "13" Then new_memb_last_grade = "13 Some Post Sec"
        If new_memb_last_grade = "14" Then new_memb_last_grade = "14 HS + certificate"
        If new_memb_last_grade = "15" Then new_memb_last_grade = "15 Four Yr Degree"
        If new_memb_last_grade = "16" Then new_memb_last_grade = "16 Grad Degree"
    End If
	'-------------------------------------------------------------------------------------------------DIALOG
	Dialog1 = "" 'Blanking out previous dialog detail
	' If new_memb_ref_numb = "" Then info_dlg_title = "CAF Addendum New Member"
	' If new_memb_ref_numb <> "" Then info_dlg_title = "CAF Addendum New Member - Memb " & new_memb_ref_numb
	If new_memb_ref_numb = "" Then BeginDialog Dialog1, 0, 0, 541, 295, "CAF Addendum New Member"
	If new_memb_ref_numb <> "" Then BeginDialog Dialog1, 0, 0, 541, 295, "CAF Addendum New Member - Memb " & new_memb_ref_numb
	  EditBox 45, 30, 95, 15, new_memb_first_name
	  EditBox 185, 30, 50, 15, new_memb_middle_name
	  EditBox 280, 30, 120, 15, new_memb_last_name
	  ComboBox 450, 30, 70, 45, "Select or Type"+chr(9)+"Junior"+chr(9)+"Senior"+chr(9)+"I"+chr(9)+"II"+chr(9)+"III"+chr(9)+"IV", new_memb_suffix
	  DropListBox 45, 50, 40, 45, "Select"+chr(9)+"Male"+chr(9)+"Female", new_memb_sex
	  EditBox 125, 50, 55, 15, new_memb_dob
	  DropListBox 250, 50, 120, 45, "Select"+chr(9)+"N - Never maried"+chr(9)+"M - Married, living w/ spouse"+chr(9)+"S - Seperated (married, living apart)"+chr(9)+"L - Legally seperated"+chr(9)+"D - Divorced"+chr(9)+"W - Widowed", new_memb_marital_status
	  EditBox 410, 50, 80, 15, new_memb_ssn
	  CheckBox 50, 75, 20, 10, "N", new_memb_race_n
	  CheckBox 70, 75, 20, 10, "A", new_memb_race_a
	  CheckBox 90, 75, 20, 10, "B", new_memb_race_b
	  CheckBox 110, 75, 20, 10, "P", new_memb_race_p
	  CheckBox 130, 75, 20, 10, "W", new_memb_race_w
	  DropListBox 200, 70, 30, 45, "No"+chr(9)+"Yes", new_memb_hispanic
	  EditBox 315, 70, 40, 15, new_memb_move_to_MN
	  DropListBox 445, 70, 75, 45, "Unknown"+chr(9)+"00  Pre 1st or Never"+chr(9)+"01 Grade 1"+chr(9)+"02 Grade 2"+chr(9)+"03 Grade 3"+chr(9)+"04 Grade 4"+chr(9)+"05 Grade 5"+chr(9)+"06 Grade 6"+chr(9)+"07 Grade 7"+chr(9)+"08 Grade 8"+chr(9)+"09 Grade 9"+chr(9)+"10 Grade 10"+chr(9)+"11 Grade 11"+chr(9)+"12 HS or GED"+chr(9)+"13 Some Post Sec"+chr(9)+"14 HS + certificate"+chr(9)+"15 Four Yr Degree"+chr(9)+"16 Grad Degree", new_memb_last_grade
	  CheckBox 100, 95, 30, 10, "SNAP", new_memb_app_snap
	  CheckBox 130, 95, 30, 10, "CASH", new_memb_app_cash
	  CheckBox 165, 95, 30, 10, "EMER", new_memb_app_emer
	  CheckBox 205, 95, 30, 10, "NONE", new_memb_app_none
	  DropListBox 305, 90, 30, 45, "Yes"+chr(9)+"No", new_memb_p_p_tog
	  ComboBox 410, 90, 110, 45, "Select or Type "+chr(9)+"02 - Spouse"+chr(9)+"03 - Child"+chr(9)+"04 - Parent"+chr(9)+"05 - Sibling"+chr(9)+"06 - Step Sibling"+chr(9)+"08 - Step Child"+chr(9)+"09 - Step Parent"+chr(9)+"10 - Aunt"+chr(9)+"11 - Uncle"+chr(9)+"12 - Neice"+chr(9)+"13 - Nephew"+chr(9)+"14 - Cousin"+chr(9)+"15 - Grandparent"+chr(9)+"16 - Grandchild"+chr(9)+"17 - Other Relative"+chr(9)+"18 - Legal Guardian"+chr(9)+"24 - Not Related"+chr(9)+"25 - Live-In Attendant"+chr(9)+"27 -     Unknown/Not Indc", new_memb_relationship
	  DropListBox 120, 125, 30, 45, "Yes"+chr(9)+"No", citizen_yn
	  DropListBox 180, 145, 30, 45, "No"+chr(9)+"Yes", disa_yn
	  DropListBox 230, 165, 30, 45, "No"+chr(9)+"Yes", unable_to_work_yn
	  DropListBox 105, 185, 30, 45, "No"+chr(9)+"Yes", school_yn
	  DropListBox 130, 205, 30, 45, "No"+chr(9)+"Yes", assets_yn
	  DropListBox 165, 225, 30, 45, "No"+chr(9)+"Yes", unea_yn
	  DropListBox 170, 245, 30, 45, "No"+chr(9)+"Yes", earned_yn
	  DropListBox 140, 265, 30, 45, "No"+chr(9)+"Yes", expenses_yn
	  DropListBox 490, 125, 30, 45, "No"+chr(9)+"Yes", qual_question_one
	  DropListBox 490, 165, 30, 45, "No"+chr(9)+"Yes", qual_question_two
	  DropListBox 490, 195, 30, 45, "No"+chr(9)+"Yes", qual_question_three
	  DropListBox 490, 225, 30, 45, "No"+chr(9)+"Yes", qual_question_four
	  DropListBox 490, 245, 30, 45, "No"+chr(9)+"Yes", qual_question_five
	  ButtonGroup ButtonPressed
		OkButton 430, 275, 50, 15
		CancelButton 485, 275, 50, 15
	  GroupBox 10, 10, 520, 100, "New Member Personal Information"
	  Text 20, 35, 20, 10, "First:"
	  Text 155, 35, 25, 10, "Middle:"
	  Text 255, 35, 20, 10, "Last:"
	  Text 20, 20, 65, 10, "New Member Name"
	  Text 415, 35, 25, 10, "Suffix:"
	  Text 20, 55, 15, 10, "Sex:"
	  Text 100, 55, 20, 10, "DOB:"
	  Text 195, 55, 50, 10, "Marital Status:"
	  Text 385, 55, 20, 10, "SSN:"
	  Text 20, 75, 25, 10, "Race:"
	  Text 165, 75, 35, 10, "Hispanic:"
	  Text 245, 75, 65, 10, "Date moved to MN:"
	  Text 370, 75, 75, 10, "Last grade completed:"
	  Text 20, 95, 80, 10, "Progams Applying For:"
	  Text 250, 95, 50, 10, "P/P Tpgether:"
	  Text 340, 95, 65, 10, "Relationship to 01:"
	  GroupBox 10, 115, 260, 175, "Addendum Questions"
	  Text 20, 130, 95, 10, "Is this person a US Citizen?"
	  Text 20, 150, 150, 10, "Is this person blind/disabled/ill/incapacitated?"
	  Text 20, 170, 205, 10, "Is this person unable to work for a reason other than diability?"
	  Text 20, 190, 80, 10, "Is this person in school?"
	  Text 20, 210, 105, 10, "Does this person have assets?"
	  Text 20, 230, 140, 10, "Does this person have unearned income?"
	  Text 20, 250, 145, 10, "Is this person employed or self-employed?"
	  Text 20, 270, 115, 10, "Does this person have expenses?"
	  GroupBox 275, 115, 255, 155, "Qualification Questions"
	  Text 285, 125, 200, 40, "Has a court or any other civil or administrative process in Minnesota or any other state found anyone in the household guilty or has anyone been disqualified from receiving public assistance for breaking any of the rules listed in the CAF?"
	  Text 285, 165, 195, 30, "Has anyone in the household been convicted of making fraudulent statements about their place of residence to get cash or SNAP benefits from more than one state?"
	  Text 285, 195, 195, 30, "Is anyone in your householdhiding or running from the law to avoid prosecution being taken into custody, or to avoid going to jail for a felony?"
	  Text 285, 225, 195, 20, "Has anyone in your household been convicted of a drug felony in the past 10 years?"
	  Text 285, 245, 195, 20, "Is anyone in your household currently violating a condition of parole, probation or supervised release?"
	EndDialog

    Do
        Do
		    err_msg = ""
            dialog Dialog1
            cancel_confirmation
            If trim(new_memb_first_name) = "" Then err_msg = err_msg & vbNewLine & "* Enter the first name of the new member."
            If trim(new_memb_last_name) = "" Then err_msg = err_msg & vbNewLine & "* Enter the last name of the new member."
            If new_memb_sex = "Select" Then err_msg = err_msg & vbNewLine & "* Indicate if new member is male or female."
            If IsDate(new_memb_dob) = False Then err_msg = err_msg & vbNewLine & "* Enter the new member's date of birth as a valid date."
            If new_memb_marital_status = "Select" Then err_msg = err_msg & vbNewLine & "* Select the new member's marital status."
            If IsDate(new_memb_move_to_MN) = False AND trim(new_memb_move_to_MN) <> "" Then err_msg = err_msg & vbNewLine & "* If the most recent date moved to Minnesota is known, enter it as a valid date."
            If trim(new_memb_relationship) = "" OR new_memb_relationship = "Select or Type" Then err_msg = err_msg & vbNewLine & "* Indicate how this new member is related to Memb 01."
            new_memb_ssn = trim(new_memb_ssn)
            If new_memb_ssn <> "" Then
                If len(new_memb_ssn) <> 9 AND len(new_memb_ssn) <> 11 Then err_msg = err_msg & vbNewLine & "* Enter only the SSN in the SSN field if provided. Any detail about the case of SSN should be added in 'Other Notes'."
            End If
            If err_msg <> "" Then MsgBox "Please Resolve to Continue:" & vbNewLine & err_msg
        Loop until err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE

    new_memb_age = ""
    new_memb_age = DateDiff("yyyy", new_memb_dob, date)
    If IsDate(new_memb_move_to_MN) = TRUE Then move_to_MN = DateDiff("m", new_memb_move_to_MN, date)
    new_memb_marital_status = right(new_memb_marital_status, len(new_memb_marital_status) - 4)
    multiple_races = False
    If new_memb_race_a = checked AND new_memb_race_n = checked Then multiple_races = TRUE
    If new_memb_race_a = checked AND new_memb_race_b = checked Then multiple_races = TRUE
    If new_memb_race_a = checked AND new_memb_race_p = checked Then multiple_races = TRUE
    If new_memb_race_a = checked AND new_memb_race_w = checked Then multiple_races = TRUE
    If new_memb_race_n = checked AND new_memb_race_b = checked Then multiple_races = TRUE
    If new_memb_race_n = checked AND new_memb_race_p = checked Then multiple_races = TRUE
    If new_memb_race_n = checked AND new_memb_race_w = checked Then multiple_races = TRUE
    If new_memb_race_b = checked AND new_memb_race_p = checked Then multiple_races = TRUE
    If new_memb_race_b = checked AND new_memb_race_w = checked Then multiple_races = TRUE
    If new_memb_race_p = checked AND new_memb_race_w = checked Then multiple_races = TRUE

    new_memb_race = ""
    If multiple_races = TRUE Then
        new_memb_race = "Multiple Races - ("
        If new_memb_race_a = checked Then new_memb_race = new_memb_race & "Asian, "
        If new_memb_race_b = checked Then new_memb_race = new_memb_race & "Black or African American, "
        If new_memb_race_p = checked Then new_memb_race = new_memb_race & "Pacific Islander, "
        If new_memb_race_n = checked Then new_memb_race = new_memb_race & "American Indian, "
        If new_memb_race_w = checked Then new_memb_race = new_memb_race & "White, "
        new_memb_race = left(new_memb_race, len(new_memb_race) - 2)
        new_memb_race = new_memb_race & ")"
    Else
        If new_memb_race_a = checked Then new_memb_race = "Asian"
        If new_memb_race_b = checked Then new_memb_race = "Black or African American"
        If new_memb_race_p = checked Then new_memb_race = "Pacific Islander"
        If new_memb_race_n = checked Then new_memb_race = "American Indian"
        If new_memb_race_w = checked Then new_memb_race = "White"
    End If
    If new_memb_race <> "" AND new_memb_hispanic = "Yes" Then new_memb_race = new_memb_race & " - Hispanic"

    dlg_one_len = 30
    dlg_two_len = 30
    show_detail_dialog_one = FALSE
    show_detail_dialog_two = FALSE

    If citizen_yn = "No" Then
        show_detail_dialog_one = TRUE
        dlg_one_len = dlg_one_len + 60
        If new_memb_ref_numb <> "" Then
            Call navigate_to_MAXIS_screen("STAT", "IMIG")
            EMWriteScreen new_memb_ref_numb, 20, 76
            transmit
            EMReadScreen versions, 1, 2, 78
            If versions <> "0" Then
                EMReadScreen us_entry_date, 10, 7, 45
                EMReadScreen nationality, 2, 10, 45
                EMReadScreen immig_status_dropdown, 30, 6, 45
                us_entry_date = replace(us_entry_date, " ", "/")
                If us_entry_date = "__/__/____" Then us_entry_date = ""
                If nationality = "AF" Then nationality = "Afghanistan"
                If nationality = "BK" Then nationality = "Bosnia"
                If nationality = "CB" Then nationality = "Cambodia"
                If nationality = "CH" Then nationality = "China, Mainland"
                If nationality = "CU" Then nationality = "Cuba"
                If nationality = "ES" Then nationality = "El Salvador"
                If nationality = "ER" Then nationality = "Eritrea"
                If nationality = "ET" Then nationality = "Ethiopia"
                If nationality = "GT" Then nationality = "Guatemala"
                If nationality = "HA" Then nationality = "Haiti"
                If nationality = "HO" Then nationality = "Honduras"
                If nationality = "IR" Then nationality = "Iran"
                If nationality = "IZ" Then nationality = "Iraq"
                If nationality = "LI" Then nationality = "Liberia"
                If nationality = "MC" Then nationality = "Micronesia"
                If nationality = "MI" Then nationality = "Marshall Islands"
                If nationality = "MX" Then nationality = "Mexico"
                If nationality = "WA" Then nationality = "Namibia"
                If nationality = "PK" Then nationality = "Pakistan"
                If nationality = "RP" Then nationality = "Philippines"
                If nationality = "PL" Then nationality = "Poland"
                If nationality = "RO" Then nationality = "Romania"
                If nationality = "RS" Then nationality = "Russia"
                If nationality = "SO" Then nationality = "Somalia"
                If nationality = "SF" Then nationality = "South Africa"
                If nationality = "TH" Then nationality = "Thailand"
                If nationality = "VM" Then nationality = "Vietnam"
                If nationality = "OT" Then nationality = ""
                If nationality = "AA" Then nationality = "Amerasian"
                If nationality = "EH" Then nationality = "Ethnic Chinese"
                If nationality = "EL" Then nationality = "Ethnic Lao"
                If nationality = "HG" Then nationality = "Hmong"
                If nationality = "KD" Then nationality = "Kurd"
                If nationality = "SJ" Then nationality = "Soviet Jew"
                If nationality = "TT" Then nationality = "Tinh"
                immig_status_dropdown = trim(immig_status_dropdown)
            End If

            Call navigate_to_MAXIS_screen("STAT", "SPON")
            EMWriteScreen new_memb_ref_numb, 20, 76
            transmit
            EMReadScreen versions, 1, 2, 78
            If versions <> "0" Then
                EMReadScreen sponsor_name, 20, 8, 38
                EMReadScreen sponsor_street, 18, 9, 38
                EMReadScreen sponsor_city, 18, 10, 38
                EMReadScreen sponsor_state, 2, 10, 62
                EMReadScreen sponsor_zip, 5, 10, 71
                EMReadScreen sponsor_phone_one, 3, 11, 40
                EMReadScreen sponsor_phone_two, 3, 11, 46
                EMReadScreen sponsor_phone_three, 3, 11, 50

                sponsor_name = replace(sponsor_name, "_", "")
                sponsor_address = sponsor_street & " " & sponsor_city & ", " & sponsor_state & " " & sponsor_zip
                sponsor_phone = "(" & sponsor_phone_one & ")" & sponsor_phone_two & "-" & sponsor_phone_three
            End If
        End If
    End If
    If move_to_MN < 12 AND move_to_MN <> "" Then
        show_detail_dialog_one = TRUE
        dlg_one_len = dlg_one_len + 40
        If new_memb_ref_numb <> "" Then
            Call navigate_to_MAXIS_screen("STAT", "MEMI")
            EMWriteScreen new_memb_ref_numb, 20, 76
            transmit
            EMReadScreen versions, 1, 2, 78
            If versions <> "0" Then
                EMReadScreen former_state, 2, 15, 78
                for each listed_state in state_array
                    If left(listed_state, 2) = former_state Then state_moved_from = listed_state
                Next
            End If
        End If
    End If
    If disa_yn = "Yes" Then
        show_detail_dialog_one = TRUE
        dlg_one_len = dlg_one_len + 40
    End If
    If unable_to_work_yn = "Yes" Then
        show_detail_dialog_one = TRUE
        dlg_one_len = dlg_one_len + 40
    End If
    If new_memb_age < 19 Then
        show_detail_dialog_one = TRUE
        dlg_one_len = dlg_one_len + 70
        If new_memb_ref_numb <> "" Then
            Call navigate_to_MAXIS_screen("STAT", "PARE")
            stat_row = 5
            Do
                EMReadScreen the_ref_numb, 2, stat_row, 3
                If the_ref_numb <> "  " Then
                    EMWriteScreen the_ref_numb, 20, 76
                    transmit
                    EMReadScreen pare_versions, 1, 2, 78
                    If pare_versions = "1" Then
                        pare_row = 8
                        Do
                            EMReadScreen child_ref_numb, 2, pare_row, 24
                            If child_ref_numb = new_memb_ref_numb Then
                                this_person = ""
                                For each clt_memb in client_array
                                    If left(clt_memb, 2) = the_ref_numb Then this_person = clt_memb
                                Next
                                If parent_one_name = "" Then
                                    parent_one_name = this_person
                                    parent_absent_or_not_one = "This parent lives in this home."
                                Else
                                    parent_two_name = this_person
                                    parent_absent_or_not_two = "This parent lives in this home."
                                End If
                            End If
                            pare_row = pare_row + 1
                        Loop until child_ref_numb = "__"
                    End If
                End If
                stat_row = stat_row + 1
            Loop until the_ref_numb = "  "

            Call navigate_to_MAXIS_screen("STAT", "ABPS")

            EMReadScreen versions, 1, 2, 78
            If versions <> "0" Then
                Do
                    abps_row = 15
                    Do
                        EMReadScreen child_ref_numb, 2, abps_row, 35
                        If child_ref_numb = new_memb_ref_numb Then
                            EMReadScreen abps_first_name, 12, 10, 63
                            EMReadScreen abps_last_name, 24, 10, 30
                            abps_first_name = replace(abps_first_name, "_", "")
                            abps_last_name = replace(abps_last_name, "_", "")
                            If abps_last_name <> "" Then
                                If parent_one_name = "" Then
                                    parent_one_name = abps_first_name & " " & abps_last_name
                                    parent_absent_or_not_one = "Absent Parent."
                                Else
                                    parent_two_name = abps_first_name & " " & abps_last_name
                                    parent_absent_or_not_two = "Absent Parent."
                                End If
                            End If
                        End If
                        abps_row = abps_row + 1
                    Loop until child_ref_numb = "__"

                    transmit
                    EMReadScreen last_page, 7, 24, 2
                Loop until last_page = "ENTER A"
            End If

        End If
    End If
    If school_yn = "Yes" Then
        show_detail_dialog_one = TRUE
        dlg_one_len = dlg_one_len + 40
    End If

    If assets_yn = "Yes" Then
        show_detail_dialog_two = TRUE
        dlg_two_len = dlg_two_len + 60
        If new_memb_ref_numb <> "" Then
            Call navigate_to_MAXIS_screen("STAT", "ACCT")
            EMWriteScreen new_memb_ref_numb, 20, 76
            transmit
            EMReadScreen versions, 1, 2, 78
            If versions <> "0" Then
                Do
                    EMReadScreen acct_pnl_type, 2, 6, 44
                    EMReadScreen acct_pnl_loc, 20, 8, 44
                    EMReadScreen acct_pnl_value, 8, 10, 46
                    If acct_pnl_type = "SV" Then acct_pnl_type = "Savings Acct"
                    If acct_pnl_type = "CK" Then acct_pnl_type = "Checking Acct"
                    If acct_pnl_type = "CE" Then acct_pnl_type = "Certificate of Deposit"
                    If acct_pnl_type = "MM" Then acct_pnl_type = "Money Market Acct"
                    If acct_pnl_type = "DC" Then acct_pnl_type = "Debit Card"
                    If acct_pnl_type = "KO" Then acct_pnl_type = "Keogh Account"
                    If acct_pnl_type = "FT" Then acct_pnl_type = "Federal Thrift SV"
                    If acct_pnl_type = "RA" Then acct_pnl_type = "Ret Annuities"
                    If acct_pnl_type = "IR" Then acct_pnl_type = "Indiv Ret Acct"
                    If acct_pnl_type = "RH" Then acct_pnl_type = "Roth IRA"
                    If acct_pnl_type = "RT" Then acct_pnl_type = "Other Ret Fund"
                    If len(acct_pnl_type) = 2 Then acct_pnl_type = ""
                    acct_pnl_loc = replace(acct_pnl_loc, "_", "")
                    if acct_pnl_type <> "" AND acct_pnl_loc <> "" Then acct_pnl_type = acct_pnl_type & " at " & acct_pnl_loc
                    acct_pnl_value = trim(acct_pnl_value)

                    If asset_type_one = "" Then
                        asset_type_one = acct_pnl_type
                        asset_value_one = acct_pnl_value
                    ElseIf asset_type_two = "" Then
                        asset_type_two = acct_pnl_type
                        asset_value_two = acct_pnl_value
                    End If
                    transmit
                    EMReadScreen last_panel, 7, 24, 2
                Loop until last_panel = "ENTER A"
            End If

            Call navigate_to_MAXIS_screen("STAT", "CARS")
            EMWriteScreen new_memb_ref_numb, 20, 76
            transmit
            EMReadScreen versions, 1, 2, 78
            If versions <> "0" Then
                Do
                    EMReadScreen cars_year, 4, 8, 31
                    EMReadScreen cars_make, 15, 8, 43
                    EMReadScreen cars_model, 15, 8, 66
                    EMReadScreen cars_pnl_value, 8, 9, 45
                    EMReadScreen cars_pnl_owed, 8, 12, 45

                    cars_make = replace(cars_make, "_", "")
                    cars_model = replace(cars_model, "_", "")
                    cars_pnl_value = trim(cars_pnl_value)
                    If cars_pnl_value = "" Then
                        EMReadScreen cars_pnl_value, 8, 9, 62
                        cars_pnl_value = trim(cars_pnl_value)
                    End If
                    cars_pnl_owed = trim(cars_pnl_owed)
                    cars_info = cars_year & " " & cars_make & " " & cars_model
                    If asset_type_one = "" Then
                        asset_type_one = cars_info
                        asset_value_one = cars_pnl_value
                        asset_owed_one = cars_pnl_owed
                    ElseIf asset_type_two = "" Then
                        asset_type_two = cars_info
                        asset_value_two = cars_pnl_value
                        asset_owed_two = cars_pnl_owed
                    End If
                    transmit
                    EMReadScreen last_panel, 7, 24, 2
                Loop until last_panel = "ENTER A"
            End If

            Call navigate_to_MAXIS_screen("STAT", "SECU")
            EMWriteScreen new_memb_ref_numb, 20, 76
            transmit
            EMReadScreen versions, 1, 2, 78
            If versions <> "0" Then
                Do
                    EMReadscreen secu_pnl_type, 2, 6, 50
                    EMReadScreen secu_pnl_loc, 20, 8, 50
                    EMReadScreen secu_pnl_value, 8, 10, 52
                    If secu_pnl_type = "LI" Then secu_pnl_type = "Life Insurance"
                    If secu_pnl_type = "ST" Then secu_pnl_type = "Stocks"
                    If secu_pnl_type = "BO" Then secu_pnl_type = "Bonds"
                    If secu_pnl_type = "CD" Then secu_pnl_type = "Ctrct for Deed"
                    If secu_pnl_type = "MO" Then secu_pnl_type = "Mortgage Note"
                    If secu_pnl_type = "AN" Then secu_pnl_type = "Annuity"
                    If secu_pnl_type = "OT" Then secu_pnl_type = ""
                    secu_pnl_value = trim(secu_pnl_value)
                    secu_pnl_loc = replace(secu_pnl_loc, "_", "")
                    if secu_pnl_type <> "" AND secu_pnl_loc <> "" Then secu_pnl_type = secu_pnl_type & " at " & secu_pnl_loc
                    If asset_type_one = "" Then
                        asset_type_one = secu_pnl_type
                        asset_value_one = secu_pnl_value
                    ElseIf asset_type_two = "" Then
                        asset_type_two = secu_pnl_type
                        asset_value_two = secu_pnl_value
                    End If
                    transmit
                    EMReadScreen last_panel, 7, 24, 2
                Loop until last_panel = "ENTER A"
            End If

            Call navigate_to_MAXIS_screen("STAT", "REST")
            EMWriteScreen new_memb_ref_numb, 20, 76
            transmit
            Do
                EMReadScreen rest_type, 15, 6, 41
                EMReadScreen rest_value, 10, 8, 41
                EMReadScreen rest_owed, 10, 9, 41
                rest_type = trim(rest_type)
                rest_value = trim(rest_value)
                rest_owed = trim(rest_owed)
                If asset_type_one = "" Then
                    asset_type_one = rest_type
                    asset_value_one = rest_value
                    asset_owed_one = rest_owed
                ElseIf asset_type_two = "" Then
                    asset_type_two = rest_type
                    asset_value_two = rest_value
                    asset_owed_two = rest_owed
                End If
                transmit
                EMReadScreen last_panel, 7, 24, 2
            Loop until last_panel = "ENTER A"
            EMReadScreen versions, 1, 2, 78
            If versions <> "0" Then
            End If
        End If
    End If
    If unea_yn = "Yes" Then
        show_detail_dialog_two = TRUE
        dlg_two_len = dlg_two_len + 40
        If new_memb_ref_numb <> "" Then
            Call navigate_to_MAXIS_screen("STAT", "UNEA")
            EMWriteScreen new_memb_ref_numb, 20, 76
            transmit
            EMReadScreen versions, 1, 2, 78
            If versions <> "0" Then
                Do
                    EMReadScreen unea_pnl_type, 20, 5, 40
                    EMReadScreen unea_pnl_total, 8, 18, 68
                    unea_pnl_type = trim(unea_pnl_type)
                    unea_pnl_total = trim(unea_pnl_total)
                    If unea_type = "" Then
                        unea_type = unea_pnl_type
                        unea_amount = unea_pnl_total
                        unea_frequency = "Monthly"
                    End If
                    transmit
                    EMReadScreen last_panel, 7, 24, 2
                Loop until last_panel = "ENTER A"
            End If
        End If
    End If
    If earned_yn = "Yes" Then
        show_detail_dialog_two = TRUE
        dlg_two_len = dlg_two_len + 60
        If new_memb_ref_numb <> "" Then
            Call navigate_to_MAXIS_screen("STAT", "JOBS")
            EMWriteScreen new_memb_ref_numb, 20, 76
            transmit
            EMReadScreen versions, 1, 2, 78
            If versions <> "0" Then
                Do
                    EMReadScreen jobs_pnl_name, 30, 7, 42
                    EMReadScreen jobs_pnl_freq, 1, 18, 35
                    EMReadScreen jobs_pnl_hrs, 3, 18, 72
                    EMReadscreen jobs_pnl_total, 8, 17, 67
                    EMReadScreen jobs_chk_one, 2, 12, 54
                    EMReadScreen jobs_chk_two, 2, 13, 54
                    EMReadScreen jobs_chk_three, 2, 14, 54
                    EMReadScreen jobs_chk_four, 2, 15, 54
                    EMReadScreen jobs_chk_five, 2, 16, 54
                    number_of_checks = 1
                    If jobs_chk_two <> "__" then number_of_checks = 2
                    If jobs_chk_three <> "__" then number_of_checks = 3
                    If jobs_chk_four <> "__" then number_of_checks = 4
                    If jobs_chk_five <> "__" then number_of_checks = 5
                    jobs_pnl_name = replace(jobs_pnl_name, "_", "")
                    weekly_hours = ""
                    If jobs_pnl_freq = "4" Then
                        weekly_hours = int(jobs_pnl_hrs/number_of_checks)
                    ElseIf jobs_pnl_freq = "3" Then
                        If number_of_checks = 3 Then
                            weekly_hours = int(jobs_pnl_hrs/6)
                        ElseIf number_of_checks = 2 Then
                            weekly_hours = int(jobs_pnl_hrs/4)
                        Else
                            weekly_hours = int(jobs_pnl_hrs/4.3)
                        End If
                    ElseIf jobs_pnl_freq = "2" OR jobs_pnl_freq = "1" Then
                        weekly_hours = int(jobs_pnl_hrs/4.3)
                    End If
                    weekly_hours = weekly_hours & ""
                    jobs_pnl_total = trim(jobs_pnl_total)
                    If jobs_pnl_total = "" Then jobs_pnl_total = 0
                    jobs_pnl_total = jobs_pnl_total * 1
                    jobs_pnl_pay = jobs_pnl_total/number_of_checks
                    jobs_pnl_pay = jobs_pnl_pay & ""
                    If jobs_pnl_freq = "1" Then jobs_pnl_freq = "Monthly"
                    If jobs_pnl_freq = "2" Then jobs_pnl_freq = "Semi-Monthly"
                    If jobs_pnl_freq = "3" Then jobs_pnl_freq = "Biweekly"
                    If jobs_pnl_freq = "4" Then jobs_pnl_freq = "Weekly"
                    If employer_name_one = "" Then
                        employer_name_one = jobs_pnl_name
                        hrs_per_wk_one = weekly_hours
                        employer_amount_one = jobs_pnl_pay
                        employer_frequency_one = jobs_pnl_freq
                    ElseIf employer_name_two Then
                        employer_name_two = jobs_pnl_name
                        hrs_per_wk_two = weekly_hours
                        employer_amount_two = jobs_pnl_pay
                        employer_frequency_two = jobs_pnl_freq
                    End If
                    transmit
                    EMReadScreen last_panel, 7, 24, 2
                Loop until last_panel = "ENTER A"
            End If

            Call navigate_to_MAXIS_screen("STAT", "BUSI")
            EMWriteScreen new_memb_ref_numb, 20, 76
            transmit
            EMReadScreen versions, 1, 2, 78
            If versions <> "0" Then
                Do
                    EMReadScreen busi_pnl_type, 2, 5, 37
                    EMReadScreen busi_cash_total, 8, 8, 69
                    EMReadScreen busi_snap_total, 8, 10, 69
                    EMReadScreen busi_pnl_hrs, 3, 13, 74
                    If busi_pnl_type = "01" Then busi_pnl_type = "Farming"
                    If busi_pnl_type = "02" Then busi_pnl_type = "Real Estate"
                    If busi_pnl_type = "03" Then busi_pnl_type = "Home Product Sales"
                    If busi_pnl_type = "04" Then busi_pnl_type = "Sales"
                    If busi_pnl_type = "05" Then busi_pnl_type = "Personal Services"
                    If busi_pnl_type = "06" Then busi_pnl_type = "Paper Route"
                    If busi_pnl_type = "07" Then busi_pnl_type = "In-Home Daycare"
                    If busi_pnl_type = "08" Then busi_pnl_type = "Rental Income"
                    If busi_pnl_type = "09" Then busi_pnl_type = ""
                    busi_pnl_hrs = busi_pnl_hrs/4.3
                    busi_pnl_hrs = int(busi_pnl_hrs)
                    busi_pnl_hrs = busi_pnl_hrs & ""
                    busi_cash_total = trim(busi_cash_total)
                    busi_snap_total = trim(busi_snap_total)
                    If busi_cash_total <> "0.00" Then
                        busi_pnl_total = busi_cash_total
                    ElseIf busi_snap_total <> "0.00" Then
                        busi_pnl_total = busi_snap_total
                    End If
                    If employer_name_one = "" Then
                        employer_name_one = "Self-Emp in " & busi_pnl_type
                        hrs_per_wk_one = busi_pnl_hrs
                        employer_amount_one = busi_pnl_total
                        employer_frequency_one = "Monthly"
                    ElseIf employer_name_two Then
                        employer_name_two = "Self-Emp in " & busi_pnl_type
                        hrs_per_wk_two = busi_pnl_hrs
                        employer_amount_two = busi_pnl_total
                        employer_frequency_two = "Monthly"
                    End If
                    transmit
                    EMReadScreen last_panel, 7, 24, 2
                Loop until last_panel = "ENTER A"
            End If
        End If
    End If
    If expenses_yn = "Yes" Then
        show_detail_dialog_two = TRUE
        dlg_two_len = dlg_two_len + 40
        If new_memb_ref_numb <> "" Then
            Call navigate_to_MAXIS_screen("STAT", "DCEX")
            EMWriteScreen new_memb_ref_numb, 20, 76
            transmit
            EMReadScreen versions, 1, 2, 78
            If versions <> "0" Then
            End If

            Call navigate_to_MAXIS_screen("STAT", "COEX")
            EMWriteScreen new_memb_ref_numb, 20, 76
            transmit
            EMReadScreen versions, 1, 2, 78
            If versions <> "0" Then
            End If

            Call navigate_to_MAXIS_screen("STAT", "SHEL")
            EMWriteScreen new_memb_ref_numb, 20, 76
            transmit
            EMReadScreen versions, 1, 2, 78
            If versions <> "0" Then
            End If
        End If
    End If

    new_memb_full_name = ""
    verif_name = ""
    If trim(new_memb_middle_name) = "" Then new_memb_full_name = new_memb_first_name & " " & new_memb_last_name
    If len(trim(new_memb_middle_name)) = 1 Then new_memb_full_name = new_memb_first_name & " " & new_memb_middle_name & ". " & new_memb_last_name
    If len(trim(new_memb_middle_name)) > 1 Then new_memb_full_name = new_memb_first_name & " " & new_memb_middle_name & " " & new_memb_last_name
    If new_memb_ref_numb <> "" Then
        verif_name = "Memb " & new_memb_ref_numb & " - " & new_memb_full_name
    Else
        verif_name = new_memb_full_name
    End If

    If show_detail_dialog_one = TRUE Then
        Do
            Do
                y_pos = 5
                If new_memb_ref_numb = "" Then dialog_title = "CAF Addendum Question Detail"
                If new_memb_ref_numb <> "" Then dialog_title = "CAF Addendum Question Detail for Memb " & new_memb_ref_numb
                BeginDialog Dialog1, 0, 0, 661, dlg_one_len, dialog_title
                  If citizen_yn = "No" Then
                      GroupBox 10, y_pos, 645, 55, "Addendum Question 1 - US Citizen/US National - No"
                      y_pos = y_pos + 20
                      Text 15, y_pos, 65, 10, "Date of Entry to US:"
                      EditBox 85, y_pos - 5, 50, 15, us_entry_date
                      Text 150, y_pos, 40, 10, "Nationality:"
                      EditBox 190, y_pos - 5, 75, 15, nationality
                      Text 275, y_pos, 65, 10, "Immigration Status:"
                      DropListBox 345, y_pos - 5, 110, 45, "Select One:"+chr(9)+"21 Refugee"+chr(9)+"22 Asylee"+chr(9)+"23 Deport/Remove Withheld"+chr(9)+"24 LPR"+chr(9)+"25 Paroled For 1 Year Or More"+chr(9)+"26 Conditional Entry < 4/80"+chr(9)+"27 Non-immigrant"+chr(9)+"28 Undocumented"+chr(9)+"50 Other Lawfully Residing"+chr(9)+"US Citizen", immig_status_dropdown
                      CheckBox 510, y_pos, 140, 10, "Requested Immigration Documentation", imig_verif_checkbox
                      y_pos = y_pos + 20
                      Text 15, y_pos, 55, 10, "Sponsor - Name:"
                      EditBox 75, y_pos - 5, 100, 15, sponsor_name
                      Text 185, y_pos, 30, 10, "Address:"
                      EditBox 220, y_pos - 5, 155, 15, sponsor_address
                      Text 385, y_pos, 25, 10, "Phone:"
                      EditBox 415, y_pos - 5, 70, 15, sponsor_phone
                      CheckBox 510, y_pos, 135, 10, "Requested Sponsor Information", sponsor_verif_checkbox
                      y_pos = y_pos + 20
                  End If

                  If move_to_MN < 12 AND move_to_MN <> "" Then
                      GroupBox 10, y_pos, 645, 35, "Addendum Question 2 - Move to MN - Less than 12 months"
                      y_pos = y_pos + 20
                      Text 20, y_pos, 180, 10, "Member Moved to MN on - " & new_memb_move_to_MN & ".   Moved From:"
                      Text 210, y_pos, 20, 10, "City:"
                      EditBox 230, y_pos - 5, 105, 15, city_moved_from
                      Text 350, y_pos, 20, 10, "State:"
                      DropListBox 375, y_pos - 5, 80, 45, state_list, state_moved_from
                      Text 470, y_pos, 55, 10, "Dates (to/from):"
                      EditBox 530, y_pos - 5, 40, 15, to_date
                      Text 575, y_pos, 5, 10, "/"
                      EditBox 580, y_pos - 5, 40, 15, from_date
                      y_pos = y_pos + 20
                  End If

                  If disa_yn = "Yes" Then
                      GroupBox 10, y_pos, 645, 35, "Addendum Question 3 - Blind/Disabled/Ill/Incapacitated - Yes"
                      y_pos = y_pos + 20
                      Text 20, y_pos, 60, 10, "Medical Problem:"
                      EditBox 85, y_pos - 5, 205, 15, medical_problem
                      Text 300, y_pos, 45, 10, "Doctor's Info:"
                      EditBox 350, y_pos - 5, 140, 15, doctors_info
                      CheckBox 510, y_pos, 120, 10, "Requested Disability Verification", disa_verif_checkbox
                      y_pos = y_pos + 20
                  End If

                  If unable_to_work_yn = "Yes" Then
                      GroupBox 10, y_pos, 645, 35, "Addendum Question 4 - Other Unable to Work (not disabled) - Yes"
                      y_pos = y_pos + 20
                      Text 20, y_pos, 85, 10, "Reason Unable to Work:"
                      EditBox 105, y_pos - 5, 540, 15, reason_unable_to_work
                      y_pos = y_pos + 20
                  End If

                  If new_memb_age < 19 Then
                      GroupBox 10, y_pos, 645, 65, "Addendum Question 5 - Under age 19"
                      Text 315, y_pos + 5, 335, 10, "Since new member is under 19, parent and possible absent parent information needs documentation."
                      y_pos = y_pos + 25
                      Text 20, y_pos, 55, 10, "Parent 1 - Name:"
                      ComboBox 80, y_pos - 5, 110, 45, full_clients_list, parent_one_name
                      DropListBox 200, y_pos - 5, 105, 45, "Parent in home?"+chr(9)+"This parent lives in this home."+chr(9)+"Absent Parent.", parent_absent_or_not_one
                      Text 315, y_pos, 60, 10, "Parent's address:"
                      EditBox 375, y_pos - 5, 130, 15, parent_one_address
                      CheckBox 515, y_pos, 135, 10, "Check here if parent shares custody.", custody_share_one
                      y_pos = y_pos + 20
                      Text 20, y_pos, 55, 10, "Parent 2 - Name:"
                      ComboBox 80, y_pos - 5, 110, 45, full_clients_list, parent_two_name
                      DropListBox 200, y_pos - 5, 105, 45, "Parent in home?"+chr(9)+"This parent lives in this home."+chr(9)+"Absent Parent.", parent_absent_or_not_two
                      Text 315, y_pos, 60, 10, "Parent's address:"
                      EditBox 375, y_pos - 5, 130, 15, parent_two_address
                      CheckBox 515, y_pos, 135, 10, "Check here if parent shares custody.", custody_share_two
                      y_pos = y_pos + 25
                  End If

                  If school_yn = "Yes" Then
                      GroupBox 10, y_pos, 645, 35, "Addendum Question 6 - In School - Yes"
                      y_pos = y_pos + 20
                      Text 20, y_pos, 55, 10, "Name of School:"
                      EditBox 80, y_pos - 5, 155, 15, school_name
                      Text 250, y_pos, 55, 10, "School Address:"
                      EditBox 310, y_pos - 5, 195, 15, school_address
                      CheckBox 510, y_pos, 140, 10, "Requested School Information", school_verif_checkbox
                      y_pos = y_pos + 20
                  End If

                  ' y_pos = y_pos + 10
                  EditBox 65, y_pos, 470, 15, verifs_needed
                  ButtonGroup ButtonPressed
                    PushButton 10, y_pos + 5, 50, 10, "Verifs Needed:", verif_button
                    OkButton 550, y_pos, 50, 15
                    CancelButton 605, y_pos, 50, 15
                EndDialog

                err_msg = ""

                dialog Dialog1
                cancel_confirmation

                If imig_verif_checkbox = checked Then verifs_needed = verifs_needed & "Immigration documentation for " & verif_memb & ".; "
                If sponsor_verif_checkbox = checked Then verifs_needed = verifs_needed & "Sponsor information of the sponsor of " & verif_memb & " - " & sponsor_name & ".; "
                If disa_verif_checkbox = checked Then verifs_needed = verifs_needed & "Ill//Incap or Diasability for " & verif_memb & " of " & medical_problem & ".; "
                If school_verif_checkbox = checked Then verifs_needed = verifs_needed & "Student Information for " & verif_memb & " at " & school_name & ".; "

                verification_dialog

                If ButtonPressed = verif_button Then err_msg = "LOOP" & err_msg
                If err_msg <> "" and left(err_msg, 4) <> "LOOP" Then MsgBox "Please Resolve to Continue:" & vbNewLine & err_msg
            Loop until err_msg = ""
            Call check_for_password(are_we_passworded_out)
        Loop until are_we_passworded_out = FALSE

    End If
    If show_detail_dialog_two = TRUE Then
        Do
            Do
                y_pos = 5
                If new_memb_ref_numb = "" Then dialog_two_titledialog_two_title = "CAF Addendum Question Detail"
                If new_memb_ref_numb <> "" Then dialog_two_title = "CAF Addendum Question Detail for Memb " & new_memb_ref_numb
				'-------------------------------------------------------------------------------------------------DIALOG
				Dialog1 = "" 'Blanking out previous dialog detail
                BeginDialog Dialog1, 0, 0, 661, dlg_two_len, dialog_two_title
                  If assets_yn = "Yes" Then
                      GroupBox 10, y_pos, 645, 55, "Addendum Question 7 - Assets - Yes"
                      y_pos = y_pos + 20
                      Text 20, y_pos, 50, 10, "Type of Asset:"
                      EditBox 70, y_pos - 5, 245, 15, asset_type_one
                      Text 330, y_pos, 30, 10, "Value: $"
                      EditBox 370, y_pos - 5, 50, 15, asset_value_one
                      Text 430, y_pos, 55, 10, "Amount Owed: $"
                      EditBox 490, y_pos - 5, 50, 15, asset_owed_one
                      CheckBox 560, y_pos, 85, 10, "Requested Verification", asset_one_verif_checkbox
                      y_pos = y_pos + 20
                      Text 20, y_pos, 50, 10, "Type of Asset:"
                      EditBox 70, y_pos - 5, 245, 15, asset_type_two
                      Text 330, y_pos, 30, 10, "Value: $"
                      EditBox 370, y_pos - 5, 50, 15, asset_value_two
                      Text 430, y_pos, 55, 10, "Amount Owed: $"
                      EditBox 490, y_pos - 5, 50, 15, asset_owed_two
                      CheckBox 560, y_pos, 85, 10, "Requested Verification", asset_two_verif_checkbox
                      y_pos = y_pos + 20
                  End If

                  If unea_yn = "Yes" Then
                      GroupBox 10, y_pos, 645, 35, "Addendum Question 8 - Unearned Income - Yes"
                      y_pos = y_pos + 20
                      Text 20, y_pos, 50, 10, "Type of UNEA:"
                      EditBox 70, y_pos - 5, 245, 15, unea_type
                      Text 330, y_pos, 35, 10, "Amount: $"
                      EditBox 370, y_pos - 5, 50, 15, unea_amount
                      Text 430, y_pos, 65, 10, "How Often Recvd:"
                      ComboBox 490, y_pos - 5, 60, 45, "Select or Type"+chr(9)+"Monthly"+chr(9)+"Semi-Monthly"+chr(9)+"Biweekly"+chr(9)+"Weekly", unea_frequency
                      CheckBox 560, y_pos, 85, 10, "Requested Verification", unea_verif_checkbox
                      y_pos = y_pos + 20
                  End If

                  If earned_yn = "Yes" Then
                      GroupBox 10, y_pos, 645, 55, "Addendum Question 9 - Employed/Self-Employed - Yes"
                      y_pos = y_pos + 20
                      Text 20, y_pos, 35, 10, "Employer:"
                      EditBox 55, y_pos - 5, 180, 15, employer_name_one
                      Text 250, y_pos, 30, 10, "Hrs/Wk:"
                      EditBox 280, y_pos - 5, 30, 15, hrs_per_wk_one
                      Text 330, y_pos, 35, 10, "Amount: $"
                      EditBox 365, y_pos - 5, 50, 15, employer_amount_one
                      Text 430, y_pos, 65, 10, "How Often Recvd:"
                      ComboBox 490, y_pos - 5, 60, 45, "Select or Type"+chr(9)+"Monthly"+chr(9)+"Semi-Monthly"+chr(9)+"Biweekly"+chr(9)+"Weekly", employer_frequency_one
                      CheckBox 560, y_pos, 85, 10, "Requested Verification", employer_verif_one_checkbox
                      y_pos = y_pos + 20
                      Text 20, y_pos, 35, 10, "Employer:"
                      EditBox 55, y_pos - 5, 180, 15, employer_name_two
                      Text 250, y_pos, 30, 10, "Hrs/Wk:"
                      EditBox 280, y_pos - 5, 30, 15, hrs_per_wk_two
                      Text 330, y_pos, 35, 10, "Amount: $"
                      EditBox 365, y_pos - 5, 50, 15, employer_amount_two
                      Text 430, y_pos, 65, 10, "How Often Recvd:"
                      ComboBox 490, y_pos - 5, 60, 45, "Select or Type"+chr(9)+"Monthly"+chr(9)+"Semi-Monthly"+chr(9)+"Biweekly"+chr(9)+"Weekly", employer_frequency_two
                      CheckBox 560, y_pos, 85, 10, "Requested Verification", employer_verif_two_checkbox
                      y_pos = y_pos + 20
                  End If

                  If expenses_yn = "Yes" Then
                      GroupBox 10, y_pos, 645, 35, "Addendum Question 10 - Expenses - Yes"
                      y_pos = y_pos + 20
                      Text 20, y_pos, 50, 10, "Expense Type:"
                      EditBox 75, y_pos - 5, 340, 15, expense_type
                      Text 425, y_pos, 65, 10, "Monthly Amount: $"
                      EditBox 490, y_pos - 5, 50, 15, expense_amount
                      CheckBox 560, y_pos, 85, 10, "Requested Verification", expense_verif_checkbox
                      y_pos = y_pos + 20
                  End If

                  y_pos = y_pos + 10
                  EditBox 65, y_pos - 5, 470, 15, verifs_needed
                  ButtonGroup ButtonPressed
                    PushButton 10, y_pos, 50, 10, "Verifs Needed:", verif_button
                    OkButton 550, y_pos - 5, 50, 15
                    CancelButton 605, y_pos - 5, 50, 15
                EndDialog

                err_msg = ""

                dialog Dialog1
                cancel_confirmation

                If asset_one_verif_checkbox = checked Then verifs_needed = verifs_needed & asset_type_one & " for " & verif_memb & ".; "
                If asset_two_verif_checkbox = checked Then verifs_needed = verifs_needed & asset_type_two & " for " & verif_memb & ".; "
                If unea_verif_checkbox = checked Then verifs_needed = verifs_needed & "Income for " & verif_memb & " from " & unea_type & ".; "
                If employer_verif_one_checkbox = checked Then verifs_needed = verifs_needed & "Income for " & verif_memb & " at " & employer_name_one & ".; "
                If employer_verif_two_checkbox = checked Then verifs_needed = verifs_needed & "Income for " & verif_memb & " at " & employer_name_two & ".; "
                If expense_verif_checkbox = checked Then verifs_needed = verifs_needed & expense_type & " expense of " & verif_memb & ".; "

                verification_dialog

                If ButtonPressed = verif_button Then err_msg = "LOOP" & err_msg
                If err_msg <> "" and left(err_msg, 4) <> "LOOP" Then MsgBox "Please Resolve to Continue:" & vbNewLine & err_msg
            Loop until err_msg = ""
            Call check_for_password(are_we_passworded_out)
        Loop until are_we_passworded_out = FALSE
    End If

    ' 'Verification NOTE
    ' verifs_needed = replace(verifs_needed, "[Information here creates a SEPARATE CASE/NOTE.]", "")
    ' If trim(verifs_needed) <> "" Then
    '
    '     verif_counter = 1
    '     verifs_needed = trim(verifs_needed)
    '     If right(verifs_needed, 1) = ";" Then verifs_needed = left(verifs_needed, len(verifs_needed) - 1)
    '     If left(verifs_needed, 1) = ";" Then verifs_needed = right(verifs_needed, len(verifs_needed) - 1)
    '     If InStr(verifs_needed, ";") <> 0 Then
    '         verifs_array = split(verifs_needed, ";")
    '     Else
    '         verifs_array = array(verifs_needed)
    '     End If
    '
    '     Call start_a_blank_CASE_NOTE
    '
    '     Call write_variable_in_CASE_NOTE("VERIFICATIONS REQUESTED")
    '
    '     Call write_bullet_and_variable_in_CASE_NOTE("Verif request form sent on", verif_req_form_sent_date)
    '
    '     Call write_variable_in_CASE_NOTE("---")
    '
    '     Call write_variable_in_CASE_NOTE("List of all verifications requested:")
    '     For each verif_item in verifs_array
    '         verif_item = trim(verif_item)
    '         If number_verifs_checkbox = checked Then verif_item = verif_counter & ". " & verif_item
    '         verif_counter = verif_counter + 1
    '         Call write_variable_with_indent_in_CASE_NOTE(verif_item)
    '     Next
    '
    '     Call write_variable_in_CASE_NOTE("---")
    '     Call write_variable_in_CASE_NOTE(worker_signature)
    '
    '     PF3
    ' End If

    'Main case note
    Call start_a_blank_CASE_NOTE

    Call write_variable_in_CASE_NOTE("CAF Addendum to add " & new_memb_full_name)

    Call write_variable_in_CASE_NOTE("* CAF Addendum received on " & addendum_date)
    If addendum_signed = checked Then Call write_variable_in_CASE_NOTE("* Addendum was signed.")
    If new_memb_ref_numb <> "" Then Call write_variable_in_CASE_NOTE("* Member added to MAXIS as Memb " & new_memb_ref_numb)
    Call write_variable_in_CASE_NOTE("--New Member Personal Information------")
    Call write_bullet_and_variable_in_CASE_NOTE("Name", new_memb_full_name)
    Call write_variable_in_CASE_NOTE("* Date of Birth: " & new_memb_dob & "     Age: " & new_memb_age)
    Call write_bullet_and_variable_in_CASE_NOTE("Sex", new_memb_sex)
    Call write_variable_with_indent_in_CASE_NOTE("Marital Status: " & new_memb_marital_status)
    If new_memb_ssn <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Social Security Number provided")
    If new_memb_race <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Race: " & new_memb_race)
    If new_memb_last_grade <> "Unknown" Then Call write_variable_with_indent_in_CASE_NOTE("Last Grade Completed: " & new_memb_last_grade)
    If new_memb_relationship <> "Select or Type" AND trim(new_memb_relationship) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Relationship to Memb 01: " & new_memb_relationship)
    Call write_variable_with_indent_in_CASE_NOTE("Purchases and Prepares food with Household: " & new_memb_p_p_tog)
    Call write_variable_in_CASE_NOTE("--Answeres to CAF Addendum Questions--")
    Call write_bullet_and_variable_in_CASE_NOTE("Q1. Person a US Citizen", citizen_yn)
    spon_info = ""
    If citizen_yn = "No" Then
        If trim(us_entry_date) <> "" Then write_variable_with_indent_in_CASE_NOTE("Entered US on " & us_entry_date & ". ")
        If trim(nationality) <> "" Then write_variable_with_indent_in_CASE_NOTE("Notionality - " & nationality & ". ")
        If immig_status_dropdown <> "Select One:" Then write_variable_with_indent_in_CASE_NOTE("Immigration status: " & right(immig_status_dropdown, len(immig_status_dropdown) - 3))
        If trim(sponsor_name) <> "" Then
            spon_info = "Sponsor: " & sponsor_name & ". "
            spon_info = spon_info & "Sponsor is at: " & sponsor_address & ". "
            spon_info = spon_info & "Phone: " & sponsor_phone & ". "
            Call write_variable_with_indent_in_CASE_NOTE(spon_info)
        End If
    End If

    If move_to_MN < 12 AND move_to_MN <> "" Then
        Call write_variable_in_CASE_NOTE("* Q2. Person moved to MN in past 12 Months: Yes")
        Call write_variable_with_indent_in_CASE_NOTE("Member Moved to MN on - " & new_memb_move_to_MN & ".")
        If trim(state_moved_from) <> "NB MN Newborn" Then
            Call write_variable_with_indent_in_CASE_NOTE("Member is Newborn in MN.")
            If trim(city_moved_from) <> "" Then
                If trim(state_moved_from) <> "Select One..." Then
                    Call write_variable_with_indent_in_CASE_NOTE("Moved from: " & trim(city_moved_from) & ", " & right(state_moved_from, len(state_moved_from) - 3 & "."))
                Else
                    Call write_variable_with_indent_in_CASE_NOTE("Moved from: " & trim(city_moved_from) & ".")
                End If
            Else
                If trim(state_moved_from) <> "Select One..." Then Call write_variable_with_indent_in_CASE_NOTE("Moved from: " & right(state_moved_from, len(state_moved_from) - 3 & "."))
            End If
        End If
        date_lived = ""
        If trim(to_date) <> "" or trim(from_date) <> "" Then dates_lived = "Lived here "
        If trim(to_date) <> "" Then dates_lived = dates_lived & trim(to_date)
        If trim(to_date) <> "" or trim(from_date) <> "" Then dates_lived = dates_lived & " - "
        If trim(from_date) <> "" Then dates_lived = dates_lived & trim(from_date)
        If dates_lived <> "" Then Call write_variable_with_indent_in_CASE_NOTE(date_lived)
    ElseIf move_to_MN > 12 Then
        Call write_variable_in_CASE_NOTE("* Q2. Person moved to MN in past 12 Months: No")
    End If

    Call write_bullet_and_variable_in_CASE_NOTE("Person disabled/ill/incapacitated", disa_yn)
    If disa_yn = "Yes" Then
        If trim(medical_problem) <> "" Then call write_variable_with_indent_in_CASE_NOTE("Medical problem: " & medical_problem & ".")
        If trim(doctors_info) <> "" Then call write_variable_with_indent_in_CASE_NOTE("Doctor's Information: " & doctors_info & ".")
    End If

    Call write_bullet_and_variable_in_CASE_NOTE("Person Unable to work for Another Reason", unable_to_work_yn)
    If unable_to_work_yn = "Yes" Then
        If trim(reason_unable_to_work) <> "" Then call write_variable_with_indent_in_CASE_NOTE("Reason unable to work: " & reason_unable_to_work)
    End If
    If new_memb_age < 19 Then
        Call write_variable_in_CASE_NOTE("* Person under 19: Yes")
        If parent_one_name <> "Select or Type" AND trim(parent_one_name) <> "" Then
            call write_variable_with_indent_in_CASE_NOTE("Parent of new member: " & parent_one_name & ". " & parent_absent_or_not_one)
            call write_variable_with_indent_in_CASE_NOTE("Address of this parent: " & parent_one_address)
            If custody_share_one = checked Then call write_variable_with_indent_in_CASE_NOTE("This parent shares custody of this child.")
        End If
        If parent_two_name <> "Select or Type" AND trim(parent_two_name) <> "" Then
            call write_variable_with_indent_in_CASE_NOTE("Parent of new member: " & parent_two_name & ". " & parent_absent_or_not_two)
            call write_variable_with_indent_in_CASE_NOTE("Address of this parent: " & parent_two_address)
            If custody_share_two = checked Then call write_variable_with_indent_in_CASE_NOTE("This parent shares custody of this child.")
        End If
    Else
        Call write_variable_in_CASE_NOTE("* Person under 19: No")
    End If

    Call write_bullet_and_variable_in_CASE_NOTE("Person in School", school_yn)
    school_info = ""
    If school_yn = "Yes" Then
        If trim(school_name) <> "" then school_info = "Name of school attending: " & school_name & ". "
        If trim(school_address) <> "" Then school_info = school_info & "School address: " & school_address & ". "
    End If

    asset_one_info = ""
    asset_two_info = ""
    Call write_bullet_and_variable_in_CASE_NOTE("Person have Assets", assets_yn)
    If assets_yn = "Yes" Then
        If trim(asset_type_one) <> "" Then
            asset_one_info = "Asset type: " & asset_type_one & ". "
            If trim(asset_value_one) <> "" Then asset_one_info = asset_one_info & "Value: $" & asset_one_info & ". "
            If trim(asset_owed_one) <> "" Then asset_one_info = asset_one_info & "Still owed: $" & asset_owed_one & ". "
            Call write_variable_with_indent_in_CASE_NOTE(asset_one_info)
        End If
        If trim(asset_ype_two) <> "" Then
            asset_two_info = "Asset type: " & asset_type_two & ". "
            If trim(asset_value_two) <> "" Then asset_two_info = asset_two_info & "Value: $" & asset_value_two & ". "
            If trim(asset_owed_two) <> "" Then asset_two_info = asset_two_info & "Still owed: $" & asset_owed_two & ". "
            Call write_variable_with_indent_in_CASE_NOTE(asset_two_info)
        End If
    End If

    Call write_bullet_and_variable_in_CASE_NOTE("Person has Unearned Income", unea_yn)
    unea_info = ""
    If unea_yn = "Yes" Then
        If trim(unea_type) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Unearned Income Type: " & unea_type)
        If trim(unea_amount) <> "" Then unea_info = "Amount of Unearned Income: $" & unea_amount & "."
        If trim(unea_frequency) <> "" AND unea_frequency <> "Select or Type" Then unea_info = unea_info & "Received " & unea_frequency & "."
    End If

    Call write_bullet_and_variable_in_CASE_NOTE("Person has Earned Income", earned_yn)
    earned_info_one = ""
    earned_info_two = ""
    If earned_yn = "Yes" Then
        If trim(employer_name_one) <> "" Then
            earned_info_one = "Work at " & employer_name_one & ". "
            If trim(hrs_per_wk_one) <> "" Then earned_info_one = earned_info_one & "; Work is " & hrs_per_wk_one & ". "
            If trim(employer_amount_one) <> "" Then earned_info_one = earned_info_one & "; Earning: $" & employer_amount_one & ". "
            If trim(employer_frequency_one) <> "" and employer_frequency_one <> "Select or Type" Then earned_info_one = earned_info_one & "Paid " & employer_frequency_one & "."
            Call write_variable_with_indent_in_CASE_NOTE(earned_info_one)
        End If
        If trim(employer_name_two) <> "" Then
            earned_info_two = "Work at " & employer_name_two & ". "
            If trim(hrs_per_wk_two) <> "" Then earned_info_two = earned_info_two & "; Work is " & hrs_per_wk_two & ". "
            If trim(employer_amount_two) <> "" Then earned_info_two = earned_info_two & "; Earning: $" & employer_amount_two & ". "
            If trim(employer_frequency_two) <> "" and employer_frequency_two <> "Select or Type" Then earned_info_two = earned_info_two & "Paid " & employer_frequency_two & "."
            call write_variable_with_indent_in_CASE_NOTE(earned_info_two)
        End If
    End If

    Call write_bullet_and_variable_in_CASE_NOTE("Person has Expenses", expenses_yn)
    expense_info = ""
    If expenses_yn = "Yes" Then
        If trim(expense_type) <> "" Then expense_info = "Expense Type: " & expense_type & ". "
        If trim(expense_amount) <> "" Then expense_info = expense_info & "Expense Amount: $" & expense_amount & ". "
        If expense_info <> "" Then Call write_variable_with_indent_in_CASE_NOTE(expense_info)
    End If

    If qual_question_one = "No" and qual_question_two = "No" and qual_question_three = "No" and qual_question_four = "No" and qual_question_five = "No" Then
        Call write_variable_in_CASE_NOTE("* All CAF Qualifying Questions answered 'No'.")
    End If

    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)


    end_msg = end_msg & vbNewLine & "Household Member - " & new_memb_full_name & " added."
    STATS_counter = STATS_counter + 1
Next

'Verification NOTE
verifs_needed = replace(verifs_needed, "[Information here creates a SEPARATE CASE/NOTE.]", "")
If trim(verifs_needed) <> "" Then

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

    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)

    PF3
End If


If qual_question_one = "Yes" or qual_question_two = "Yes" or qual_question_three = "Yes" or qual_question_four = "Yes" or qual_question_five = "Yes" Then

    Do
        Do
            err_msg = ""
            dlg_len = 55

            If qual_question_one = "Yes" Then dlg_len = dlg_len + 40
            If qual_question_two = "Yes" Then dlg_len = dlg_len + 30
            If qual_question_three = "Yes" Then dlg_len = dlg_len + 30
            If qual_question_four = "Yes" Then dlg_len = dlg_len + 20
            If qual_question_five = "Yes" Then dlg_len = dlg_len + 20


            y_pos = 30
			'-------------------------------------------------------------------------------------------------DIALOG
			Dialog1 = "" 'Blanking out previous dialog detail
            BeginDialog Dialog1, 0, 0, 416, dlg_len, "Qualification Questions Dialog"
              Text 10, 10, 395, 15, "At least one qualification question was answered with 'Yes'. Enter the Household Member that was indicated on the form. "
              If qual_question_one = "Yes" Then
                  Text 10, y_pos, 200, 40, "Has a court or any other civil or administrative process in Minnesota or any other state found anyone in the household guilty or has anyone been disqualified from receiving public assistance for breaking any of the rules listed in the CAF?"
                  Text 225, y_pos, 70, 10, "Household Member:"
                  ComboBox 305, y_pos, 105, 45, full_clients_list, qual_question_memb_one
                  y_pos = y_pos + 40
              End If
              If qual_question_two = "Yes" Then
                  Text 10, y_pos, 195, 30, "Has anyone in the household been convicted of making fraudulent statements about their place of residence to get cash or SNAP benefits from more than one state?"
                  Text 225, y_pos, 70, 10, "Household Member:"
                  ComboBox 305, y_pos, 105, 45, full_clients_list, qual_question_memb_two
                  y_pos = y_pos + 30
              End If
              If qual_question_three = "Yes" Then
                  Text 10, y_pos, 195, 30, "Is anyone in your householdhiding or running from the law to avoid prosecution being taken into custody, or to avoid going to jail for a felony?"
                  Text 225, y_pos, 70, 10, "Household Member:"
                  ComboBox 305, y_pos, 105, 45, full_clients_list, qual_question_memb_three
                  y_pos = y_pos + 30
              End If
              If qual_question_four = "Yes" Then
                  Text 10, y_pos, 195, 20, "Has anyone in your household been convicted of a drug felony in the past 10 years?"
                  Text 225, y_pos, 70, 10, "Household Member:"
                  ComboBox 305, y_pos, 105, 45, full_clients_list, qual_question_memb_four
                  y_pos = y_pos + 20
              End If
              If qual_question_five = "Yes" Then
                  Text 10, y_pos, 195, 20, "Is anyone in your household currently violating a condition of parole, probation or supervised release?"
                  Text 225, y_pos, 70, 10, "Household Member:"
                  ComboBox 305, y_pos, 105, 45, full_clients_list, qual_question_memb_five
                  y_pos = y_pos + 20
              End If
              y_pos = y_pos + 5
              ButtonGroup ButtonPressed
                OkButton 305, y_pos, 50, 15
                CancelButton 360, y_pos, 50, 15
            EndDialog

            dialog Dialog1
            cancel_confirmation

            If qual_question_one = "Yes" AND (qual_question_memb_one = "Select or Type" OR trim(qual_question_memb_one) = "") Then err_msg = err_msg & vbNewLine & "Since Question One was ansered 'yes' explain to whom the answer applies."
            If qual_question_two = "Yes" AND (qual_question_memb_two = "Select or Type" OR trim(qual_question_memb_two) = "") Then err_msg = err_msg & vbNewLine & "Since Question Two was ansered 'yes' explain to whom the answer applies."
            If qual_question_three = "Yes" AND (qual_question_memb_three = "Select or Type" OR trim(qual_question_memb_three) = "") Then err_msg = err_msg & vbNewLine & "Since Question Three was ansered 'yes' explain to whom the answer applies."
            If qual_question_four = "Yes" AND (qual_question_memb_four = "Select or Type" OR trim(qual_question_memb_four) = "") Then err_msg = err_msg & vbNewLine & "Since Question Four was ansered 'yes' explain to whom the answer applies."
            If qual_question_five = "Yes" AND (qual_question_memb_five = "Select or Type" OR trim(qual_question_memb_five) = "") Then err_msg = err_msg & vbNewLine & "Since Question Five was ansered 'yes' explain to whom the answer applies."

            If err_msg <> "" Then MsgBox "Please Resolve to Continue:" & vbNewLine & err_msg

        Loop until err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE

    Call start_a_blank_CASE_NOTE

    Call write_variable_in_CASE_NOTE("CAF Qualifying Questions had an answer of 'YES' for at least one question")
    If qual_question_one = "Yes" Then Call write_bullet_and_variable_in_CASE_NOTE("Fraud/DISQ for IPV (program violation)", qual_question_memb_one)
    If qual_question_two = "Yes" Then Call write_bullet_and_variable_in_CASE_NOTE("SNAP in more than One State", qual_question_memb_two)
    If qual_question_three = "Yes" Then Call write_bullet_and_variable_in_CASE_NOTE("Fleeing Felon", qual_question_memb_three)
    If qual_question_four = "Yes" Then Call write_bullet_and_variable_in_CASE_NOTE("Drug Felony", qual_question_memb_four)
    If qual_question_five = "Yes" Then Call write_bullet_and_variable_in_CASE_NOTE("Parole/Probation Violation", qual_question_memb_five)
    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)

End If

Call script_end_procedure_with_error_report(end_msg)
