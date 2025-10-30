'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "DAIL - COLA REVIEW AND APPROVE.vbs"
start_time = timer
functionality_time = timer
STATS_counter = 1              'sets the stats counter at 1
STATS_manualtime = 90          'manual run time in seconds
STATS_denomination = "C"       'C is for each case

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
call changelog_update("12/13/2023", "COLA Review and Approve functionality has been updated to support actions taken from many COLA DAIL messages we receive. Here are some of the functionality supports added to the new release of this script ##~## - Ability to create CASE/NOTE of a UNEA Income Summary.##~## - Review SVES/TPQY for any member on the case. ##~## - Redirect to NOTES - Eligibility Summary to document completed approvals you have completed in the same day.##~##", "Casey Love, Hennepin County")
call changelog_update("05/31/2019", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'New function that will make sure we are at DAIL, then enter the code to nagivate directly from the dail. This can help us keep the tie to the DAIL
function nav_in_DAIL(dail_nav_letter)
	EMReadScreen on_dail_check, 27, 2, 26						'Checking if we are on DAIL - so we can get there
	Do while on_dail_check <> "WORKERS DAILY REPORT (DAIL)"
		PF3														'back up one level
		EMReadScreen SELF_check, 4, 2, 50						'see if we are on SELF
		If SELF_check = "SELF" Then Call navigate_to_MAXIS_screen ("DAIL", "DAIL")	'if we are on SELF - navigate to DAIL/DAIL
		EMReadScreen on_dail_check, 27, 2, 26					'read to see if we made it to DAIL
	Loop

	'now we need to find the message that we started with
	dail_row = 6
	Do
		EMReadScreen line_message, 60, dail_row, 20				'read the top message for this case
		line_message = trim(line_message)						'trim the message
		If line_message = full_message Then						'If the message matches the message from the start of the dail scrubber run, then we can enter the info
			EMWriteScreen dail_nav_letter, dail_row, 3
			transmit
			Exit Do
		End If
		dail_row = dail_row + 1
	Loop until dail_row = 20
end function

'CASE/NOTE is in a Function because it is in a loop and looks cleaner in the later script run.
function cola_summary_note()
	functionality_info = "UNEA Information: "
	Call nav_in_DAIL("N")

	call start_a_blank_CASE_NOTE
	Call write_variable_in_CASE_NOTE ("===COLA INCOME SUMMARY for " & MAXIS_footer_month & "/" & MAXIS_footer_year & "===")
	Call write_variable_in_CASE_NOTE("--- UNEARNED INCOME ---")
	For unea_info = 0 to UBound(UNEA_FROM_DAIL_RUN, 2)
		If UNEA_FROM_DAIL_RUN(unea_checkbox_const, unea_info) = unchecked Then
			functionality_info = functionality_info &  " ~|~ " & "MEMB " & UNEA_FROM_DAIL_RUN(dail_memb_ref_const, unea_info) & " - " & UNEA_FROM_DAIL_RUN(dail_unea_type_code_const, unea_info) & " - " & UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_info)
			Call write_variable_in_CASE_NOTE("* MEMB " & UNEA_FROM_DAIL_RUN(dail_memb_ref_const, unea_info) & " - " & UNEA_FROM_DAIL_RUN(dail_memb_name_const, unea_info) & " - UNEA " & UNEA_FROM_DAIL_RUN(dail_memb_ref_const, unea_info) & " " & UNEA_FROM_DAIL_RUN(dail_unea_instance_const, unea_info))
			Call write_variable_in_CASE_NOTE("  Unearned Income Type: " & UNEA_FROM_DAIL_RUN(dail_unea_type_code_const, unea_info) & " - " & UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_info))
			Call write_variable_in_CASE_NOTE("  Prospective Amount: $ " & UNEA_FROM_DAIL_RUN(dail_prosp_total_const, unea_info) & "  -  COLA Disregard $ " & UNEA_FROM_DAIL_RUN(dail_cola_disregard_const, unea_info))
		End If
	Next
	functionality_info = functionality_info & " ========= "
	Call write_bullet_and_variable_in_CASE_NOTE("UNEA Notes", UNEA_notes)
	If MEDI_exists = True Then
		functionality_info = functionality_info & "MEDI Information"
		Call write_variable_in_CASE_NOTE("--- MEDICARE PART B ---")
		For medi_count = 0 to UBound(HH_member_array)
			functionality_info = functionality_info & " ~|~ " & "MEMB " & HH_member_array(medi_count)
			If MEDI_PART_B_ARRAY(medi_count) <> "" Then Call write_variable_in_CASE_NOTE("* MEMB " & HH_member_array(medi_count) & " - " & COLA_NAME_ARRAY(medi_count) & ": Medicare - Part B $ " & MEDI_PART_B_ARRAY(medi_count))
		Next
		Call write_bullet_and_variable_in_CASE_NOTE("MEDI Notes", MEDI_notes)
	End If
	If trim(other_notes) <> "" Then
		Call write_variable_in_CASE_NOTE("--- OTHER ---")
		Call write_bullet_and_variable_in_CASE_NOTE("Notes", other_notes)
	End If
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)
	cola_summary_note_entered = True
	STATS_manualtime = STATS_manualtime + 90

	PF3			'back to DAIL
	PF3
end function

'Declaring some information for an array about UNEA information
const dail_memb_ref_const			= 0
const dail_memb_name_const			= 1
const dail_unea_instance_const		= 2
const dail_unea_type_code_const		= 3
const dail_unea_type_info_const		= 4
const dail_cola_disregard_const		= 5
const dail_prosp_total_const		= 6
Const unea_checkbox_const			= 7
const dail_last_const				= 8

DIM UNEA_FROM_DAIL_RUN()
ReDIM UNEA_FROM_DAIL_RUN(dail_last_const, 0)
unea_count = 0

'SCRIPT======================================================================================================================
EMConnect ""					'connect to MAXIS

run_elig_summ_btn = 1001		'define the buttons we are going to use
cola_summary_note_btn = 1002
finish_script_btn = 1003

MEDI_exists = False					'default some booleans to identify what we are dealing with to false
cola_summary_note_entered = False

EMReadScreen cola_footer_month, 2, 6, 11		'Find the footer month/year from the
EMReadScreen cola_footer_year, 2, 6, 14

'Setting the footer month to either CM or CM+1
If cola_footer_month = CM_plus_1_mo AND cola_footer_year = CM_plus_1_yr Then
    MAXIS_footer_month = CM_plus_1_mo
    MAXIS_footer_year = CM_plus_1_yr
Else
    MAXIS_footer_month = CM_mo
    MAXIS_footer_year = CM_yr
End If

Call nav_in_DAIL("H")		'go to CASE/CURR

'Reading Case information
Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
PF3             'Back to DAIL

Call nav_in_DAIL("S")			'GOING to STAT/MEMB
EMWriteScreen "MEMB", 20, 71
transmit

'We need an array of the member names and some other information about members
Call HH_member_custom_dialog(HH_member_array)				'Creating a custom dialog for determining who the HH members are
number_of_persons = UBound(HH_member_array)
Dim MEDI_PART_B_ARRAY()
ReDim MEDI_PART_B_ARRAY(number_of_persons)
Dim COLA_NAME_ARRAY()
ReDim COLA_NAME_ARRAY(number_of_persons)
Dim SSN_ARRAY()
ReDim SSN_ARRAY(number_of_persons)
Dim PERS_BUTTON_ARRAY()
ReDim PERS_BUTTON_ARRAY(number_of_persons)
medi_count = 0												'set the counter for reading person information.

'GRABBING THE INFO FOR THE CASE NOTE-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'DETERMINES THE UNEARNED INCOME RECEIVED BY THE CLIENT
For each HH_member in HH_member_array						'for each person that was selected
	call navigate_to_MAXIS_screen("STAT", "MEMB")			'Go to STAT/MEMB
	EMWriteScreen HH_member, 20, 76							'Enter the member's number to get to the right panel
	transmit
	EMReadScreen first_name, 12, 6, 63						'read a bunch of information about the person
	EMReadScreen last_name, 25, 6, 30
	memb_name = replace(first_name, "_", "") & " " & replace(last_name, "_", "")
	EmReadscreen client_SSN, 11, 7, 42
	client_SSN = replace(client_SSN, " ", "")
	COLA_NAME_ARRAY(medi_count) = memb_name
	SSN_ARRAY(medi_count) = client_SSN
	PERS_BUTTON_ARRAY(medi_count) = 500 + medi_count

	call navigate_to_MAXIS_screen("STAT", "UNEA")			'Now let's check UNEA and grab all the income information
	EMWriteScreen HH_member, 20, 76
	EMWriteScreen "01", 20, 79
	transmit
	EMReadScreen UNEA_total, 1, 2, 78
	If UNEA_total <> 0 then
		Do
			ReDIM Preserve UNEA_FROM_DAIL_RUN(dail_last_const, unea_count)		'resizing the UNEA array

			'gathering information from the panel
			EMReadScreen UNEA_panel_current, 1, 2, 73
			EMReadScreen inc_type, 2, 5, 37
			EMReadScreen cola_disregard, 8, 10, 67
			EMReadScreen unea_amount, 8, 18, 68

			UNEA_FROM_DAIL_RUN(dail_memb_ref_const, unea_count) = HH_member
			UNEA_FROM_DAIL_RUN(dail_memb_name_const, unea_count) = memb_name
			UNEA_FROM_DAIL_RUN(dail_unea_instance_const, unea_count) = "0" & UNEA_panel_current
			UNEA_FROM_DAIL_RUN(dail_unea_type_code_const, unea_count) = inc_type

			If inc_type = "01" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "RSDI, Disability"
			If inc_type = "02" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "RSDI, No Disability"
			If inc_type = "03" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "SSI"
			If inc_type = "06" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Non-MN Public Assistance"
			If inc_type = "11" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "VA Disability Benefit"
			If inc_type = "12" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "VA Pension"
			If inc_type = "13" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "VA other"
			If inc_type = "38" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "VA Aid & Attendance"
			If inc_type = "14" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Unemployment Insurance"
			If inc_type = "15" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Worker's Comp"
			If inc_type = "16" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Railroad Retirement"
			If inc_type = "17" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Other Retirement"
			If inc_type = "18" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Military Entitlement"
			If inc_type = "19" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Foster Care Child Requesting SNAP"
			If inc_type = "20" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Foster Care Child NOT Requesting SNAP"
			If inc_type = "21" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Foster Care Adult Requesting SNAP"
			If inc_type = "22" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Foster Care Adult NOT Requesting SNAP"
			If inc_type = "23" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Dividends"
			If inc_type = "24" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Interest"
			If inc_type = "25" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Counted Gifts or Prizes"
			If inc_type = "26" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Strike Benefit"
			If inc_type = "27" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Contract for Deed"
			If inc_type = "28" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Illegal Income"
			If inc_type = "29" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Other Countable Income"
			If inc_type = "30" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Infrequent, <30, Not Counted"
			If inc_type = "31" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Other SNAP Only"
			If inc_type = "08" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Direct Child Support"
			If inc_type = "35" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Direct Spousal Support"
			If inc_type = "36" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Disbursed Child Support"
			If inc_type = "37" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Disbursed Spousal Support"
			If inc_type = "39" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Disbursed Child Support Arrears"
			If inc_type = "40" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Disbursed Spousal Support Arrears"
			If inc_type = "43" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Disbursed Excess Child Support"
			If inc_type = "44" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "MSA - Excess Income for SSI"
			If inc_type = "45" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "County 88 Child Support"
			If inc_type = "46" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "County 88 Gaming"
			If inc_type = "47" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Tribal Income"
			If inc_type = "48" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Trust income"
			If inc_type = "49" Then UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_count) = "Non-Recurring Income > $60 per Quarter"


			cola_disregard = replace(cola_disregard, "_", "")
			If cola_disregard = "" Then cola_disregard = "NONE"
			UNEA_FROM_DAIL_RUN(dail_cola_disregard_const, unea_count) = trim(cola_disregard)
			UNEA_FROM_DAIL_RUN(dail_prosp_total_const, unea_count) = trim(unea_amount)

			unea_count = unea_count + 1		'increasing the UNEA count

			transmit
			EMReadScreen warn_msg, 60, 24, 2
			warn_msg = trim(warn_msg)
		Loop Until warn_msg = "ENTER A VALID COMMAND OR PF-KEY"
	End if

	'here we go check MEDI to see if this exists for this household member
	call navigate_to_MAXIS_screen("STAT", "MEDI")
	EMWriteScreen HH_member, 20, 76
	transmit
    EMReadScreen MEDI_total, 1, 2, 78
    If MEDI_total <> 0 then
		MEDI_exists = True
        EMReadScreen medicare_part_B, 8, 7, 73
		MEDI_PART_B_ARRAY(medi_count) = trim(medicare_part_B)
    End if
	medi_count = medi_count + 1
Next
PF3             							'Back to DAIL

'Make sure we are actually on DAIL
EMReadScreen on_dail, 4, 2, 48
If on_dail <> "DAIL" Then
	Call back_to_SELF
	EMWriteScreen "DAIL", 16, 43
	EMWriteScreen "DAIL", 21, 70
	transmit

	Call write_value_and_transmit(MAXIS_case_number, 20, 38)
End If

'If this case has ANY UNEA income, we will display a dialog that will let a NOTE be entered about this income
If unea_count <> 0 Then
	'first we check CASE/NOTE to see if a note already exists
	Call nav_in_DAIL("N")							'navigate to CASE/NOTE

	recent_cola_inocme_summary_found = False
	too_old_date = DateAdd("d", -30, date) 			'we only need to look back 30 days

	note_row = 5
	Do
		EMReadScreen note_date, 8, note_row, 6                  'reading the note date

		EMReadScreen note_title, 55, note_row, 25               'reading the note header
		note_title = trim(note_title)

		If left(note_title, 26) = "===COLA INCOME SUMMARY for" Then		'this is the note header we are looking for
			recent_cola_inocme_summary_found = True
		End if

		if note_date = "        " then Exit Do                                      'if we are at the end of the list of notes - we can't read any more

		note_row = note_row + 1									'go to the next note to review it
		if note_row = 19 then
			note_row = 5
			PF8
			EMReadScreen check_for_last_page, 9, 24, 14
			If check_for_last_page = "LAST PAGE" Then Exit Do
		End If
		EMReadScreen next_note_date, 8, note_row, 6
		if next_note_date = "        " then Exit Do
	Loop until DateDiff("d", too_old_date, next_note_date) <= 0
	PF3             'Back to DAIL

	'Here is the dialog - creating the correct size
	dlg_len = 185
	dlg_len = dlg_len + (UBound(UNEA_FROM_DAIL_RUN, 2)+1) * 20
	If MEDI_exists = True Then
		dlg_len = dlg_len + 30
		For medi_count = 0 to UBound(HH_member_array)
			If MEDI_PART_B_ARRAY(medi_count) <> "" Then dlg_len = dlg_len + 10
		Next
	End If

	Do
		Do
			'create the dialog
			BeginDialog Dialog1, 0, 0, 550, dlg_len, "COLA Income Information"
				Text 10, 10, 100, 10, "* * * SCRIPT OPERATION * * *"
				If cola_summary_note_entered = True Then
					Text 20, 20, 7, 10, "---"
					Text 25, 20, 75, 20, "COLA Summary CASE/NOTE created."
				End If
				GroupBox 125, 10, 305, 45,  "Complete different tasks to support COLA processing:"
				Text 130, 20, 280, 10,  " --- View SVES by person   -   Select person's name on right."
				Text 130, 30, 280, 10,  " --- Create Approval CASE/NOTE   -   Select 'Run - NOTES- Eligibility Summary'."
				Text 130, 40, 295, 10,  " --- Create COLA Updates CASE/NOTE   -   Select 'Create COLA Income Summary NOTE'."

				Text 10, 55, 200, 10, "Active Programs: " & list_active_programs

				Text 10, 65, 300, 10, "COLA Income Summary for " & MAXIS_footer_month & "/" & MAXIS_footer_year & "  ---  This case has the following unearned income:"
				y_pos = 80
				For unea_info = 0 to UBound(UNEA_FROM_DAIL_RUN, 2)
					Text 25, y_pos, 155, 10, "MEMB " & UNEA_FROM_DAIL_RUN(dail_memb_ref_const, unea_info) & " - " & UNEA_FROM_DAIL_RUN(dail_memb_name_const, unea_info)
					Text 185, y_pos, 200, 10, "UNEA " & UNEA_FROM_DAIL_RUN(dail_memb_ref_const, unea_info) & " " & UNEA_FROM_DAIL_RUN(dail_unea_instance_const, unea_info) & "                      " & UNEA_FROM_DAIL_RUN(dail_unea_type_code_const, unea_info) & " - " & UNEA_FROM_DAIL_RUN(dail_unea_type_info_const, unea_info)
					Text 60, y_pos+10, 115, 10, "Prospective Amount $ " & UNEA_FROM_DAIL_RUN(dail_prosp_total_const, unea_info)
					Text 185, y_pos+10, 135, 10, "COLA Disregard $ " & UNEA_FROM_DAIL_RUN(dail_cola_disregard_const, unea_info)
					CheckBox 385, y_pos, 200, 10, "Check here to Exclude from CASE/NOTE", UNEA_FROM_DAIL_RUN(unea_checkbox_const, unea_info)
					y_pos = y_pos + 20
				Next
				y_pos = y_pos + 10
				Text 25, y_pos, 50, 10, "UNEA Notes:"
				EditBox 70, y_pos-5, 360, 15, UNEA_notes
				y_pos = y_pos + 15


				If MEDI_exists = True Then
					Text 10, y_pos, 95, 10, "This case has MEDICARE:"
					y_pos = y_pos + 15

					For medi_count = 0 to UBound(HH_member_array)
						If MEDI_PART_B_ARRAY(medi_count) <> "" Then
							Text 25, y_pos, 155, 10, "MEMB " & HH_member_array(medi_count) & " - " & COLA_NAME_ARRAY(medi_count)
							Text 185, y_pos, 160, 10, "MEDICARE - Part B $ " & MEDI_PART_B_ARRAY(medi_count)
							y_pos = y_pos + 10
						End If
					Next
					y_pos = y_pos + 5

					Text 25, y_pos, 50, 10, "MEDI Notes:"
					EditBox 70, y_pos-5, 360, 15, MEDI_notes
					y_pos = y_pos + 10

				End If
				y_pos = y_pos + 10
				Text 10, y_pos, 75, 10, "Additional Notes"
				EditBox 10, y_pos+10, 530, 15, other_notes
				y_pos = y_pos + 30

				Text 325, y_pos+5, 65, 10, "Worker Signature"
				EditBox 390, y_pos, 150, 15, worker_signature

				Text 10, y_pos+25, 100, 10, "What would you like to do?"
				ButtonGroup ButtonPressed
					PushButton 110, y_pos+20, 150, 15, "Run NOTES - Eligibliity Summary", run_elig_summ_btn
					PushButton 275, y_pos+20, 150, 15, "Create COLA Income Summary NOTE", cola_summary_note_btn
					PushButton 440, y_pos+20, 100, 15, "Finish - no additional actions", finish_script_btn
					GroupBox 440, 10, 105, 65, "Check SVES"
					y_pos = 20
					For i = 0 to UBound(COLA_NAME_ARRAY)
						PushButton 445, y_pos, 95, 13, COLA_NAME_ARRAY(i), PERS_BUTTON_ARRAY(i)
						y_pos = y_pos + 15
					Next
			EndDialog

			err_msg = ""

			dialog Dialog1
			If ButtonPressed = finish_script_btn Then call script_end_procedure("")
			cancel_without_confirmation

			err_msg = "LOOP"										'this is a little weird, we are just always looping

			'need to check password first
			CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
		Loop Until are_we_passworded_out = false					'loops until user passwords back in

		'Here we figure out what button was pushed
		For i = 0 to UBound(COLA_NAME_ARRAY)							'looking for a MEMBER button - to go to SVES for that member
			If ButtonPressed = PERS_BUTTON_ARRAY(i) Then
				Call nav_in_DAIL("I")
				EMWriteScreen SSN_ARRAY(i), 3, 63
				EMWaitReady 0,0
				Call write_value_and_transmit("SVES", 20, 71)
				EMWaitReady 0,0
				Call write_value_and_transmit("TPQY", 20, 70)
				STATS_manualtime = STATS_manualtime + 30
				call collect_script_usage_data(name_of_script & " - INFC-SVES Viewed ", "SVES for " & COLA_NAME_ARRAY(i), functionality_time)		'record script FUNCTIONALITY usage in SQL
				functionality_time = timer

				'checking for NON-DISCLOSURE AGREEMENT REQUIRED FOR ACCESS TO IEVS FUNCTIONS'
				EMReadScreen agreement_check, 9, 2, 24
				IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")
			End If
		Next

		If ButtonPressed = cola_summary_note_btn Then
			functionality_info = ""			'this variable will be used to capture information to pass to the usage log
			Call cola_summary_note			'Button to create the CASE/NOTE
			call collect_script_usage_data(name_of_script & " - COLA Summary Note Created", functionality_info, functionality_time)		'record script FUNCTIONALITY usage in SQL
			functionality_time = timer
		End If
		If ButtonPressed = run_elig_summ_btn Then
			Call collect_script_usage_data(name_of_script & " - REDIRECT to Elig Summary", "", functionality_time)		'record script FUNCTIONALITY usage in SQL
			Call run_from_GitHub(script_repository & "notes/eligibility-summary.vbs")		'Button to run Eligibility Summary
		End If
	Loop until err_msg = ""
Else
	'This dialog is ONLY for viewing SVES - there was no UNEA found so we do not have a CASE/NOTE option
	dlg_len = 80
	dlg_len = dlg_len + (UBound(COLA_NAME_ARRAY)+1) * 20

	Do
		Do
			BeginDialog Dialog1, 0, 0, 300, dlg_len, "Review SVES for Members"
				EditBox 500, 600, 10, 10, dummy_box
				Text 10, 10, 100, 10, "* * * SCRIPT OPERATION * * *"
				Text 15, 25, 305, 10,  "The script can pull up SVES responses for the following members:"

				ButtonGroup ButtonPressed
					y_pos = 35
					For i = 0 to UBound(COLA_NAME_ARRAY)
						PushButton 25, y_pos, 95, 13, COLA_NAME_ARRAY(i), PERS_BUTTON_ARRAY(i)
						y_pos = y_pos + 15
					Next
					y_pos = y_pos + 10
					Text 15, y_pos, 270, 20,  "This script cannot display any additional information or actions because no UNEA is listed on the case for selected household members."
					PushButton 220, y_pos+20, 75, 15, "Finish", finish_script_btn
			EndDialog

			err_msg = ""

			dialog Dialog1
			If ButtonPressed = finish_script_btn Then ButtonPressed = 0
			cancel_without_confirmation

			err_msg = "LOOP"										'this is a little weird, we are just always looping

			'need to check password first
			CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
		Loop Until are_we_passworded_out = false					'loops until user passwords back in

		'Here we figure out what button was pushe
		For i = 0 to UBound(COLA_NAME_ARRAY)							'looking for a MEMBER button - to go to SVES for that member
			If ButtonPressed = PERS_BUTTON_ARRAY(i) Then
				Call nav_in_DAIL("I")
				EMWriteScreen SSN_ARRAY(i), 3, 63
				EMWaitReady 0,0
				' MsgBox "SSN there?"
				Call write_value_and_transmit("SVES", 20, 71)
				EMWaitReady 0,0
				Call write_value_and_transmit("TPQY", 20, 70)
				STATS_manualtime = STATS_manualtime + 30
				call collect_script_usage_data(name_of_script & " - INFC-SVES Viewed ", "SVES for " & COLA_NAME_ARRAY(i), functionality_time)		'record script FUNCTIONALITY usage in SQL
				functionality_time = timer

				'checking for NON-DISCLOSURE AGREEMENT REQUIRED FOR ACCESS TO IEVS FUNCTIONS'
				EMReadScreen agreement_check, 9, 2, 24
				IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")
			End If
		Next
	Loop until err_msg = ""
End If

script_end_procedure_with_error_report(end_msg)

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------12/13/23
'--Tab orders reviewed & confirmed----------------------------------------------12/13/23
'--Mandatory fields all present & Reviewed--------------------------------------12/13/23
'--All variables in dialog match mandatory fields-------------------------------N/A							There are no mandatory fields that are entered by the worker
'Review dialog names for content and content fit in dialog----------------------12/13/23
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------12/13/23
'--CASE:NOTE Header doesn't look funky------------------------------------------12/13/23
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A							Script operates on a loop - showing the dialog again after the NOTE
'--write_variable_in_CASE_NOTE function:
' confirm that proper punctuation is used --------------------------------------12/13/23
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------N/A							DAIL supports should handle for this
'--MAXIS_background_check reviewed (if applicable)------------------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------N/A							DAIL supports should handle for this
'--Out-of-County handling reviewed----------------------------------------------N/A							DAIL supports should handle for this
'--script_end_procedures (w/ or w/o error messaging)----------------------------12/13/23
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------12/13/23
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------12/13/23
'--Incrementors reviewed (if necessary)-----------------------------------------N/A
'--Denomination reviewed -------------------------------------------------------12/13/23
'--Script name reviewed---------------------------------------------------------12/13/23
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------12/13/23
'--comment Code-----------------------------------------------------------------12/13/23
'--Update Changelog for release/update------------------------------------------12/13/23
'--Remove testing message boxes-------------------------------------------------12/13/23
'--Remove testing code/unnecessary code-----------------------------------------12/13/23
'--Review/update SharePoint instructions----------------------------------------12/13/23
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------N/A
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------N/A
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------N/A