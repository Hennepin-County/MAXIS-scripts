'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "UTILITIES - BUDGET ESTIMATOR.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 320         'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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
call changelog_update("03/21/2019", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'CONSTANTS==================================================================================================================

'There is going to be a lot of array work - will need constants
'Maybe classes will happen here - then need to define the classes
'Defining some constants to make array life easier
'Main Array constants
const clt_name 					= 0
const clt_ref  					= 1
const include_snap 				= 2
const include_family_cash		= 3
const include_adult_cash		= 4
const clt_age					= 5
const clt_a_c 					= 6
const clt_asset_total 			= 7
const clt_sav_acct 				= 8
const clt_chk_acct 				= 9
const clt_asset_other_type 		= 10
const clt_asset_other_bal 		= 11
const clt_ei_gross 				= 12
const clt_ssi_income 			= 13
const clt_rsdi_income 			= 14
const clt_other_unea_1_type 	= 15
const clt_other_unea_1_amt 		= 16
const clt_other_unea_2_type 	= 17
const clt_other_unea_2_amt 		= 18

'Array for jobs constants
const employee        = 0
const employer        = 1
const job_retro_gross = 2
const job_prosp_gross = 3
const job_pic_gross   = 4
const job_pay_freq    = 5
const how_many_chck   = 6
const check_1_date    = 7
const check_1_gross   = 8
const check_2_date    = 9
const check_2_gross   = 10
const check_3_date    = 11
const check_3_gross   = 12
const check_4_date    = 13
const check_4_gross   = 14
const check_5_date    = 15
const check_5_gross   = 16
const pic_hrs_wk      = 17
const pic_rate_pay    = 18

const vehicle_type  = 0
const vehicle_year  = 1
const vehicle_make  = 2
const vehicle_model = 3
const vehicle_value = 4

const security_type			= 0
const security_description  = 1
const security_value		= 2
const security_withdrawl	= 3

Const account_holder 	= 0
Const account_type 		= 1
Const account_balance 	= 2

Const unea_person = 0
Const rsdi_amt = 1
Const ssi_amt = 2
Const other_1_type = 3
Const other_1_amt = 4
Const other_2_type = 5
Const other_2_amt = 6

'FUNCTIONS==================================================================================================================

'All the buttons will need Functions

'SELECT MEMBERS BUTTON
'This function calls the HH Member Function on queue so the selection can be changed.
function MEMB_NUMBER_BUTTON_PRESSED
	MEMB_function
	HH_member_array = ""
	FOR i = 0 to total_clients
		IF all_clients_array(i, 0) <> "" THEN 						'creates the final array to be used by other scripts.
			IF all_clients_array(i, 1) = 1 THEN						'if the person/string has been checked on the dialog then the reference number portion (left 2) will be added to new HH_member_array
				'msgbox all_clients_
				HH_member_array = Right(all_clients_array(i, 0), len(all_clients_array(i, 0))   ) & ", " & HH_member_array
			END IF
		END IF
	NEXT
	hh_size_split = Len(HH_member_array) - Len(Replace(HH_member_array,",",""))
	hh_size = CStr(hh_size_split)
end function

'this function creates the hh member dynamic dialog
function MEMB_function

	Do
		err_msg = ""

        Dialog1 = ""
		BeginDialog Dialog1, 0,  0, 355, (40 + (UBound(CASE_INFO_ARRAY, 2) + 1) * 15), "HH Member Dialog"   'Creates the dynamic dialog. The height will change based on the number of clients it finds.
		  Text 10, 5, 145, 10, "Who is applying?:"
		  FOR all_clts = 0 to UBound(CASE_INFO_ARRAY, 2)									'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
		  	  Text 10, (20 + (all_clts * 15)), 150, 10, CASE_INFO_ARRAY(clt_ref, all_clts) & " - " &  CASE_INFO_ARRAY(clt_name, all_clts) & "  " & CASE_INFO_ARRAY(clt_a_c, all_clts)   'Ignores and blank scanned in persons/strings to avoid a blank checkbox
		  	  CheckBox 180, (20 + (all_clts * 15)), 50, 10, "SNAP", CASE_INFO_ARRAY(include_snap, all_clts)
			  CheckBox 240, (20 + (all_clts * 15)), 50, 10, "MFIP", CASE_INFO_ARRAY(include_family_cash, all_clts)
			  CheckBox 300, (20 + (all_clts * 15)), 50, 10, "GA/MSA", CASE_INFO_ARRAY(include_adult_cash, all_clts)
		  NEXT
		  If clients_left <> "" Then Text 10, 20 + ((UBound(CASE_INFO_ARRAY, 2) + 1) * 15), 290, 10, "HH Membs listed in STAT but Removed: " & clients_left
		  ButtonGroup ButtonPressed
		  OkButton 300, 20 + ((UBound(CASE_INFO_ARRAY, 2) + 1) * 15), 50, 15
		EndDialog

		Dialog Dialog1

		cash_hh_size = 0
		FOR all_clts = 0 to UBound(CASE_INFO_ARRAY, 2)
			If CASE_INFO_ARRAY(clt_age, all_clts) < 18 AND CASE_INFO_ARRAY(include_adult_cash, all_clts) = checked Then
				err_msg = "At this time the script does not support minor GA. Everyone requesting adult cash programs must be 18 or older."
				CASE_INFO_ARRAY(include_adult_cash, all_clts) = unchecked
			End If
			If CASE_INFO_ARRAY(include_adult_cash, all_clts) = checked Then cash_hh_size = cash_hh_size + 1
		Next
		If err_msg <> "" Then MsgBox err_msg
	Loop until err_msg = ""

	If cash_hh_size <> 0 AND cash_basis_met_checkbox = unchecked Then
		adult_cash_basis_msg = MsgBox ("There are " & cash_hh_size & "person(s) included in the Adult Cash grant size." &vbnewLine & vbnewLine & "Do all of these person(s) meet the disabilty/age basis of eligibility?", vbquesion + vbyesno)
		If adult_cash_basis_msg = vbYes Then cash_basis_met_checkbox = checked
	End If
	'FIND_FPG_THRIFTY_FOOD
end function

function NEW_CASE_MEMB_FUNCTION
	number_of_adults = number_of_adults & ""
	number_of_children = number_of_children & ""
	snap_hh_size = snap_hh_size & ""
	family_cash_hh_size = family_cash_hh_size & ""
	adult_cash_hh_size = adult_cash_hh_size & ""

	Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 116, 125, "Case Composition"
	  EditBox 10, 25, 15, 15, number_of_adults
	  EditBox 60, 25, 15, 15, number_of_children
	  DropListBox 70, 45, 30, 45, "0"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"5"+chr(9)+"6"+chr(9)+"7"+chr(9)+"8"+chr(9)+"9"+chr(9)+"10"+chr(9)+"11"+chr(9)+"12"+chr(9)+"13"+chr(9)+"14"+chr(9)+"15"+chr(9)+"16"+chr(9)+"17"+chr(9)+"18"+chr(9)+"19"+chr(9)+"20", snap_hh_size
	  DropListBox 70, 65, 30, 45, "0"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"5"+chr(9)+"6"+chr(9)+"7"+chr(9)+"8"+chr(9)+"9"+chr(9)+"10"+chr(9)+"11"+chr(9)+"12"+chr(9)+"13"+chr(9)+"14"+chr(9)+"15"+chr(9)+"16"+chr(9)+"17"+chr(9)+"18"+chr(9)+"19"+chr(9)+"20", family_cash_hh_size
	  DropListBox 70, 85, 30, 45, "0"+chr(9)+"1"+chr(9)+"2", adult_cash_hh_size
	  ButtonGroup ButtonPressed
	    OkButton 60, 105, 50, 15
	  Text 5, 10, 85, 10, "Household Composition:"
	  Text 30, 30, 25, 10, "Adults"
	  Text 80, 30, 35, 10, "Children"
	  Text 10, 50, 50, 10, "SNAP HH Size"
	  Text 10, 70, 55, 10, "Family Cash HH"
	  Text 10, 90, 55, 10, "Adult Cash HH"
	EndDialog
	Do
		hh_comp_err_msg = ""

		Dialog Dialog1

		If IsNumeric(number_of_adults) <> TRUE Then
			hh_comp_err_msg = hh_comp_err_msg & vbnewLine & "The number of adults must be entered as a number."
		Else
			If number_of_adults < 1 Then hh_comp_err_msg = hh_comp_err_msg & vbnewLine & "There must be at least 1 adult in the household."
		End If

		If hh_comp_err_msg <> "" Then MsgBox "Please Resolve to continue:" & vbnewLine & hh_comp_err_msg
	Loop until hh_comp_err_msg = ""

	If number_of_adults = "" Then number_of_adults = 0
	If number_of_children = "" Then number_of_children = 0

	number_of_adults = number_of_adults * 1
	number_of_children = number_of_children * 1
	snap_hh_size = snap_hh_size * 1
	family_cash_hh_size = family_cash_hh_size * 1
	adult_cash_hh_size = adult_cash_hh_size * 1

	If number_of_children <> 0 Then FAMILY_CASE = TRUE

	If adult_cash_hh_size <> 0 AND cash_basis_met_checkbox = unchecked Then
		adult_cash_basis_msg = MsgBox ("There are " & adult_cash_hh_size & "person(s) included in the Adult Cash grant size." &vbnewLine & vbnewLine & "Do all of these person(s) meet the disabilty/age basis of eligibility?", vbquesion + vbyesno)
		If adult_cash_basis_msg = vbYes Then cash_basis_met_checkbox = checked
	End If
end function

'INDICATE RELATIONSHIPS (use what is in proof of relationship script)

'CALCULATE EARNED INCOME (very similar to emer-screen)
function EARNED_INCOME_BUTTON_PRESSED

	For all_clts = 0 to UBOUND(CASE_INFO_ARRAY, 2)
		CASE_INFO_ARRAY(clt_ei_gross, all_clts) = 0
	Next

	For each_job = 0 to UBOUND(EI_ARRAY, 2)
		EI_ARRAY(job_retro_gross, each_job) = EI_ARRAY(job_retro_gross, each_job) & ""
		EI_ARRAY(job_prosp_gross, each_job) = EI_ARRAY(job_prosp_gross, each_job) & ""
		EI_ARRAY(job_pic_gross, each_job) = EI_ARRAY(job_pic_gross, each_job) & ""

		EI_ARRAY(check_1_gross, each_job) = EI_ARRAY(check_1_gross, each_job) & ""
		EI_ARRAY(check_2_gross, each_job) = EI_ARRAY(check_2_gross, each_job) & ""
		EI_ARRAY(check_3_gross, each_job) = EI_ARRAY(check_3_gross, each_job) & ""
		EI_ARRAY(check_4_gross, each_job) = EI_ARRAY(check_4_gross, each_job) & ""
		EI_ARRAY(check_5_gross, each_job) = EI_ARRAY(check_5_gross, each_job) & ""

		EI_ARRAY(pic_rate_pay, each_job) = EI_ARRAY(pic_rate_pay, each_job) & ""
		EI_ARRAY(pic_hrs_wk, each_job) = EI_ARRAY(pic_hrs_wk, each_job) & ""
	Next

	Do
		add_to_len = 0
		For every_one = 0 to UBound(EI_ARRAY, 2)
			add_to_len = add_to_len + 50
	''		MsgBox "~" & EI_ARRAY(how_many_chck, every_one) & "~"
			If EI_ARRAY(how_many_chck, every_one) <> " " then add_to_len = add_to_len + (20 * EI_ARRAY(how_many_chck, every_one))
		Next

		Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 340, 60 + add_to_len, "Earned Income"
			y_pos = 0
			For job_in_case = 0 to UBound(EI_ARRAY, 2)
				If no_case_number_checkbox = unchecked Then DropListBox 5, 20 + y_pos, 105, 45, HH_Memb_DropDown, EI_ARRAY(employee, job_in_case)
				If no_case_number_checkbox = checked Then EditBox 5, 20 + y_pos, 105, 15, EI_ARRAY(employee, job_in_case)
				EditBox 120, 20 + y_pos, 130, 15, EI_ARRAY(employer, job_in_case)
				DropListBox 260, 20 + y_pos, 35, 45, " "+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"5", EI_ARRAY(how_many_chck, job_in_case)
				ButtonGroup ButtonPressed
			  	  PushButton 310, 20 + y_pos, 25, 15, "Enter", job_enter
				Text 5, 45 + y_pos, 50, 10, "Retro Gross:"
				EI_ARRAY(job_retro_gross, job_in_case) = EI_ARRAY(job_retro_gross, job_in_case) & ""
				EditBox 60, 40 + y_pos, 30, 15, EI_ARRAY(job_retro_gross, job_in_case)
				Text 95, 45 + y_pos, 50, 10, "Prosp Gross:"
				EI_ARRAY(job_prosp_gross, job_in_case) = EI_ARRAY(job_prosp_gross, job_in_case) & ""
				EditBox 140, 40 + y_pos, 30, 15, EI_ARRAY(job_prosp_gross, job_in_case)
				Text 180, 45 + y_pos, 50, 10, "PIC Gross:"
				EI_ARRAY(job_prosp_gross, job_in_case) = EI_ARRAY(job_prosp_gross, job_in_case) & ""
				EditBox 235, 40 + y_pos, 30, 15, EI_ARRAY(job_pic_gross, job_in_case)

				Text 5, 65 + y_pos, 50, 10, "Hours/Week:"
				EditBox 60, 60 + y_pos, 30, 15, EI_ARRAY(pic_hrs_wk, job_in_case)
				Text 95, 65 + y_pos, 50, 10, "Rate of Pay:"
				EditBox 140, 60 + y_pos, 30, 15, EI_ARRAY(pic_rate_pay, job_in_case)

				Text 200, 65 + y_pos, 50, 10, "Pay Frequency:"
				DropListBox 260, 60 + y_pos, 75, 45, "Once/Month - 1"+chr(9)+"Twice/Month - 2"+chr(9)+"Biweekly - 3"+chr(9)+"Weekly - 4", EI_ARRAY(job_pay_freq, job_in_case)
				array_counter = 7
				If EI_ARRAY(how_many_chck, job_in_case) <> "" AND EI_ARRAY(how_many_chck, job_in_case) <> " " Then
					'If EI_ARRAY(job_verif, job_in_case) = "" Then EI_ARRAY(job_verif, job_in_case) = "Verifications?"
					Text 35, 80 + y_pos, 20, 10, "Date"
					Text 120, 80 + y_pos, 50, 10, "Gross Amount"
					'Text 200, 60 + y_pos, 50, 10, "Net Amount"
					'DropListBox 260, 40 + y_pos, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", EI_ARRAY(job_verif, job_in_case)
					For checks_to_enter = 1 to EI_ARRAY(how_many_chck, job_in_case)
						EditBox 35, 95 + y_pos, 50, 15, EI_ARRAY(array_counter, job_in_case)
                        EI_ARRAY(array_counter + 1, job_in_case) = EI_ARRAY(array_counter + 1, job_in_case) & ""
                        'EI_ARRAY(array_counter + 2, job_in_case) = EI_ARRAY(array_counter + 2, job_in_case) & ""
						EditBox 120, 95 + y_pos, 50, 15, EI_ARRAY(array_counter + 1, job_in_case)
						'EditBox 200, 55 + y_pos, 50, 15, EI_ARRAY(array_counter + 2, job_in_case)
						array_counter = array_counter + 2
						y_pos = y_pos + 15
					Next
					y_pos = y_pos + 80
				Else
					y_pos = y_pos + 20
				End If
			Next
			ButtonGroup ButtonPressed
			  PushButton 5, 40 + add_to_len, 10, 15, "+", plus_button
			  PushButton 15, 40 + add_to_len, 10, 15, "-", minus_button
			  OkButton 285, 40 + add_to_len, 50, 15
			Text 5, 5, 45, 10, "HH Member"
			Text 120, 5, 40, 10, "Employer"
			Text 260, 5, 40, 10, "# of Checks"
		EndDialog

		Dialog Dialog1

        If ButtonPressed = plus_button Then
		 	add_another = Ubound(EI_ARRAY, 2) + 1
			ReDim Preserve EI_ARRAY (18, add_another)
		End If
		DETERMINE_COUNTED_EI
	Loop Until ButtonPressed = -1

''	DETERMINE_COUNTED_EI
end function

function DETERMINE_COUNTED_EI

	case_ei_gross = 0
	case_ei_net = 0
	SNAP_JOBS_Income = 0
	Adult_Cash_JOBS_Income = 0
	Family_Cash_JOBS_Income = 0
	For each_job = 0 to UBOUND(EI_ARRAY, 2)
		If EI_ARRAY(job_retro_gross, each_job) = "" then EI_ARRAY(job_retro_gross, each_job) = 0
		If EI_ARRAY(job_prosp_gross, each_job) = "" then EI_ARRAY(job_prosp_gross, each_job) = 0
		If EI_ARRAY(job_pic_gross, each_job) = "" then EI_ARRAY(job_pic_gross, each_job) = 0

		If EI_ARRAY(check_1_gross, each_job) = "" then EI_ARRAY(check_1_gross, each_job) = 0
		If EI_ARRAY(check_2_gross, each_job) = "" then EI_ARRAY(check_2_gross, each_job) = 0
		If EI_ARRAY(check_3_gross, each_job) = "" then EI_ARRAY(check_3_gross, each_job) = 0
		If EI_ARRAY(check_4_gross, each_job) = "" then EI_ARRAY(check_4_gross, each_job) = 0
		If EI_ARRAY(check_5_gross, each_job) = "" then EI_ARRAY(check_5_gross, each_job) = 0

		If EI_ARRAY(pic_rate_pay, each_job) = "" Then EI_ARRAY(pic_rate_pay, each_job) = 0
		If EI_ARRAY(pic_hrs_wk, each_job) = "" Then EI_ARRAY(pic_hrs_wk, each_job) = 0

		EI_ARRAY(job_retro_gross, each_job) = EI_ARRAY(job_retro_gross, each_job) * 1
		EI_ARRAY(job_prosp_gross, each_job) = EI_ARRAY(job_prosp_gross, each_job) * 1
		EI_ARRAY(job_pic_gross, each_job) = EI_ARRAY(job_pic_gross, each_job) * 1

		EI_ARRAY(check_1_gross, each_job) = EI_ARRAY(check_1_gross, each_job) * 1
		EI_ARRAY(check_2_gross, each_job) = EI_ARRAY(check_2_gross, each_job) * 1
		EI_ARRAY(check_3_gross, each_job) = EI_ARRAY(check_3_gross, each_job) * 1
		EI_ARRAY(check_4_gross, each_job) = EI_ARRAY(check_4_gross, each_job) * 1
		EI_ARRAY(check_5_gross, each_job) = EI_ARRAY(check_5_gross, each_job) * 1

		EI_ARRAY(pic_rate_pay, each_job) = EI_ARRAY(pic_rate_pay, each_job) * 1
		EI_ARRAY(pic_hrs_wk, each_job) = EI_ARRAY(pic_hrs_wk, each_job) * 1

		If EI_ARRAY(pic_rate_pay, each_job) <> 0 AND EI_ARRAY(pic_hrs_wk, each_job) <> 0 Then
			EI_ARRAY(job_prosp_gross, each_job) = EI_ARRAY(pic_rate_pay, each_job) * EI_ARRAY(pic_hrs_wk, each_job) * 4
			If EI_ARRAY(job_retro_gross, each_job) = 0 Then EI_ARRAY(job_retro_gross, each_job) = EI_ARRAY(pic_rate_pay, each_job) * EI_ARRAY(pic_hrs_wk, each_job) * 4
			If EI_ARRAY(job_pic_gross, each_job) = 0 Then EI_ARRAY(job_pic_gross, each_job) = EI_ARRAY(pic_rate_pay, each_job) * EI_ARRAY(pic_hrs_wk, each_job) * 4.3
		End If

		If EI_ARRAY(job_retro_gross, each_job) = 0 Then
			If EI_ARRAY(job_prosp_gross, each_job) <> 0 Then
				EI_ARRAY(job_retro_gross, each_job) = EI_ARRAY(job_prosp_gross, each_job)
			ElseIf EI_ARRAY(job_pic_gross, each_job) <> 0 Then
				EI_ARRAY(job_retro_gross, each_job) = EI_ARRAY(job_pic_gross, each_job)
			End If
		End If

		If EI_ARRAY(job_prosp_gross, each_job) = 0 Then
			If EI_ARRAY(job_retro_gross, each_job) <> 0 Then
				EI_ARRAY(job_prosp_gross, each_job) = EI_ARRAY(job_retro_gross, each_job)
			ElseIf EI_ARRAY(job_pic_gross, each_job) <> 0 Then
				EI_ARRAY(job_prosp_gross, each_job) = EI_ARRAY(job_pic_gross, each_job)
			End If
		End If

		If EI_ARRAY(job_pic_gross, each_job) = 0 Then
			If EI_ARRAY(job_prosp_gross, each_job) <> 0 Then
				EI_ARRAY(job_pic_gross, each_job) = EI_ARRAY(job_prosp_gross, each_job)
			ElseIf EI_ARRAY(job_retro_gross, each_job) <> 0 Then
				EI_ARRAY(job_pic_gross, each_job) = EI_ARRAY(job_retro_gross, each_job)
			End If
		End If

		case_ei_gross = case_ei_gross + EI_ARRAY(job_prosp_gross, each_job)
		If no_case_number_checkbox = unchecked Then
			For all_clts = 0 to UBOUND(CASE_INFO_ARRAY, 2)
				If Left(EI_ARRAY(employee, each_job), 2) = CASE_INFO_ARRAY(clt_ref, all_clts) Then
					If EI_ARRAY(how_many_chck, each_job) <> " " then
						checks_total = 0
						array_counter = 8
						checks_average = 0

						For entered_check = 1 to EI_ARRAY(how_many_chck, each_job)
							checks_total = checks_total + EI_ARRAY(array_counter, each_job)
							array_counter = array_counter + 2
						Next
						If EI_ARRAY(how_many_chck, each_job) <> 0 Then checks_average = checks_total/EI_ARRAY(how_many_chck, each_job)
						If EI_ARRAY(job_pay_freq, each_job) = "Once/Month - 1"  Then checks_multiplier = 1
						If EI_ARRAY(job_pay_freq, each_job) = "Twice/Month - 2" Then checks_multiplier = 2
						If EI_ARRAY(job_pay_freq, each_job) = "Biweekly - 3"    Then checks_multiplier = 2.15
						If EI_ARRAY(job_pay_freq, each_job) = "Weekly - 4"      Then checks_multiplier = 4.3

						If checks_total <> 0 Then EI_ARRAY(job_retro_gross, each_job) = checks_total

						If CASE_INFO_ARRAY(include_snap, all_clts) = checked AND CASE_INFO_ARRAY(clt_age, all_clts) >= 18 Then SNAP_JOBS_Income = SNAP_JOBS_Income + (checks_average * checks_multiplier)

						EI_ARRAY(job_pic_gross, each_job) = checks_average * checks_multiplier
					Else
						checks_average = EI_ARRAY(pic_rate_pay, each_job) * EI_ARRAY(pic_hrs_wk, each_job)

						If CASE_INFO_ARRAY(include_snap, all_clts) = checked AND CASE_INFO_ARRAY(clt_age, all_clts) >= 18 Then SNAP_JOBS_Income = SNAP_JOBS_Income + checks_average * 4.3

						EI_ARRAY(job_pic_gross, each_job) = checks_average * 4.3
					End If

					If CASE_INFO_ARRAY(include_family_cash, all_clts) = checked AND CASE_INFO_ARRAY(clt_age, all_clts) >= 18 Then
						If EI_ARRAY(job_retro_gross, each_job) <> 0 Then
							Family_Cash_JOBS_Income = Family_Cash_JOBS_Income + EI_ARRAY(job_retro_gross, each_job)
						ElseIf EI_ARRAY(job_prosp_gross, each_job) <> 0 Then
							Family_Cash_JOBS_Income = Family_Cash_JOBS_Income + EI_ARRAY(job_prosp_gross, each_job)
						End If
					End If
					If CASE_INFO_ARRAY(include_adult_cash, all_clts) = checked Then
						If EI_ARRAY(job_retro_gross, each_job) <> 0 Then
							Adult_Cash_JOBS_Income = Adult_Cash_JOBS_Income + EI_ARRAY(job_retro_gross, each_job)
						ElseIf EI_ARRAY(job_prosp_gross, each_job) <> 0 Then
							Adult_Cash_JOBS_Income = Adult_Cash_JOBS_Income + EI_ARRAY(job_prosp_gross, each_job)
						End If
					End If

					'CASE_INFO_ARRAY(clt_ei_gross, all_clts) = CASE_INFO_ARRAY(clt_ei_gross, all_clts) + EI_ARRAY(job_pic_gross, each_job)
				End If
			Next
		Else
			If EI_ARRAY(how_many_chck, each_job) <> " " then
				checks_total = 0
				array_counter = 8
				checks_average = 0

				For entered_check = 1 to EI_ARRAY(how_many_chck, each_job)
					checks_total = checks_total + EI_ARRAY(array_counter, each_job)
					array_counter = array_counter + 2
				Next
				If EI_ARRAY(how_many_chck, each_job) <> 0 Then checks_average = checks_total/EI_ARRAY(how_many_chck, each_job)
				If EI_ARRAY(job_pay_freq, each_job) = "Once/Month - 1"  Then checks_multiplier = 1
				If EI_ARRAY(job_pay_freq, each_job) = "Twice/Month - 2" Then checks_multiplier = 2
				If EI_ARRAY(job_pay_freq, each_job) = "Biweekly - 3"    Then checks_multiplier = 2.15
				If EI_ARRAY(job_pay_freq, each_job) = "Weekly - 4"      Then checks_multiplier = 4.3

				If checks_total <> 0 Then EI_ARRAY(job_retro_gross, each_job) = checks_total

				SNAP_JOBS_Income = SNAP_JOBS_Income + (checks_average * checks_multiplier)
				EI_ARRAY(job_pic_gross, each_job) = checks_average * checks_multiplier
			Else
				checks_average = EI_ARRAY(pic_rate_pay, each_job) * EI_ARRAY(pic_hrs_wk, each_job)

				SNAP_JOBS_Income = SNAP_JOBS_Income + checks_average * 4.3

				EI_ARRAY(job_pic_gross, each_job) = checks_average * 4.3
			End If

			If EI_ARRAY(job_retro_gross, each_job) <> 0 Then
				Family_Cash_JOBS_Income = Family_Cash_JOBS_Income + EI_ARRAY(job_retro_gross, each_job)
			ElseIf EI_ARRAY(job_prosp_gross, each_job) <> 0 Then
				Family_Cash_JOBS_Income = Family_Cash_JOBS_Income + EI_ARRAY(job_prosp_gross, each_job)
			End If

			If EI_ARRAY(job_retro_gross, each_job) <> 0 Then
				Adult_Cash_JOBS_Income = Adult_Cash_JOBS_Income + EI_ARRAY(job_retro_gross, each_job)
			ElseIf EI_ARRAY(job_prosp_gross, each_job) <> 0 Then
				Adult_Cash_JOBS_Income = Adult_Cash_JOBS_Income + EI_ARRAY(job_prosp_gross, each_job)
			End If


		End If
	Next
	case_ei_gross = case_ei_gross * 1
	SNAP_JOBS_Income = SNAP_JOBS_Income * 1
	Family_Cash_JOBS_Income = Family_Cash_JOBS_Income * 1
	Adult_Cash_JOBS_Income = Adult_Cash_JOBS_Income * 1
end function

'CALCULATE UNEARNED INCOME
function UNEA_BUTTON_PRESSED

	If no_case_number_checkbox = unchecked Then
		For all_clts = 0 to UBOUND(CASE_INFO_ARRAY, 2)
			CASE_INFO_ARRAY(clt_rsdi_income, all_clts) = CASE_INFO_ARRAY(clt_rsdi_income, all_clts) & ""
			CASE_INFO_ARRAY(clt_ssi_income, all_clts) = CASE_INFO_ARRAY(clt_ssi_income, all_clts) & ""
			CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) & ""
			CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) & ""
		Next
		total_case_unea = total_case_unea & ""
		Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 556, 60 + (20 * UBOUND(CASE_INFO_ARRAY, 2)), "Unearned Income"
		  Text 5, 5, 50, 10, "Person"
		  Text 170, 5, 35, 10, "RSDI"
		  Text 205, 5, 30, 10, "SSI"
		  Text 275, 5, 50, 10, "Other - 1"
		  Text 415, 5, 50, 10, "Other - 2"
		  For all_clts = 0 to UBOUND(CASE_INFO_ARRAY, 2)
		  	  Text 5, 20 + (20 * all_clts), 150, 10, CASE_INFO_ARRAY(clt_ref, all_clts) & " - " & CASE_INFO_ARRAY(clt_name, all_clts)
			  EditBox 170, 20 + (20 * all_clts), 25, 15, CASE_INFO_ARRAY(clt_rsdi_income, all_clts)
			  EditBox 205, 20 + (20 * all_clts), 25, 15, CASE_INFO_ARRAY(clt_ssi_income, all_clts)
			  ComboBox 275, 20 + (20 * all_clts), 60, 45, ""+chr(9)+"Other"+chr(9)+"Child Support"+chr(9)+"SSI"+chr(9)+"RSDI"+chr(9)+"Non-MN PA"+chr(9)+"VA Disability Benefit"+chr(9)+"VA Pension"+chr(9)+"VA Other"+chr(9)+"VA Aid & Attendance"+chr(9)+"Unemployment Insurance"+chr(9)+"Worker's Comp"+chr(9)+"Railroad Retirement"+chr(9)+"Other Retirement"+chr(9)+"Military Allotment"+chr(9)+"FC Child Requesting FS"+chr(9)+"FC Child Not Req FS"+chr(9)+"FC Adult Requesting FS"+chr(9)+"FC Adult Not Req FS"+chr(9)+"Dividends"+chr(9)+"Interest"+chr(9)+"Cnt Gifts Or Prizes"+chr(9)+"Strike Benefit 27 Contract For Deed"+chr(9)+"Illegal Income"+chr(9)+"Infrequent <30 Not Counted"+chr(9)+"Other FS Only"+chr(9)+"Infreq <= $20 MSA Exclusion"+chr(9)+"Direct Spousal Support"+chr(9)+"Disbursed Spousal Sup"+chr(9)+"Disbursed Spsl Sup Arrears"+chr(9)+"County 88 Gaming", CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts)
			  EditBox 340, 20 + (20 * all_clts), 40, 15, CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts)
			  ComboBox 415, 20 + (20 * all_clts), 60, 45, ""+chr(9)+"Other"+chr(9)+"Child Support"+chr(9)+"SSI"+chr(9)+"RSDI"+chr(9)+"Non-MN PA"+chr(9)+"VA Disability Benefit"+chr(9)+"VA Pension"+chr(9)+"VA Other"+chr(9)+"VA Aid & Attendance"+chr(9)+"Unemployment Insurance"+chr(9)+"Worker's Comp"+chr(9)+"Railroad Retirement"+chr(9)+"Other Retirement"+chr(9)+"Military Allotment"+chr(9)+"FC Child Requesting FS"+chr(9)+"FC Child Not Req FS"+chr(9)+"FC Adult Requesting FS"+chr(9)+"FC Adult Not Req FS"+chr(9)+"Dividends"+chr(9)+"Interest"+chr(9)+"Cnt Gifts Or Prizes"+chr(9)+"Strike Benefit 27 Contract For Deed"+chr(9)+"Illegal Income"+chr(9)+"Infrequent <30 Not Counted"+chr(9)+"Other FS Only"+chr(9)+"Infreq <= $20 MSA Exclusion"+chr(9)+"Direct Spousal Support"+chr(9)+"Disbursed Spousal Sup"+chr(9)+"Disbursed Spsl Sup Arrears"+chr(9)+"County 88 Gaming", CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts)
			  EditBox 480, 20 + (20 * all_clts), 40, 15, CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts)
		  Next
		  ButtonGroup ButtonPressed
		    OkButton 500, 40 + (20 * UBOUND(CASE_INFO_ARRAY, 2)), 50, 15
		EndDialog

		Dialog Dialog1

		total_case_unea = 0
		For all_clts = 0 to UBOUND(CASE_INFO_ARRAY, 2)
			If CASE_INFO_ARRAY(clt_rsdi_income, all_clts)      = "" Then CASE_INFO_ARRAY(clt_rsdi_income, all_clts) = 0
			If CASE_INFO_ARRAY(clt_ssi_income, all_clts)       = "" Then CASE_INFO_ARRAY(clt_ssi_income, all_clts) = 0
			If CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = "" Then CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = 0
			If CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = "" Then CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = 0

			CASE_INFO_ARRAY(clt_rsdi_income, all_clts) = CASE_INFO_ARRAY(clt_rsdi_income, all_clts) * 1
			CASE_INFO_ARRAY(clt_ssi_income, all_clts) = CASE_INFO_ARRAY(clt_ssi_income, all_clts) * 1
			CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) * 1
			CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) * 1

			total_case_unea = total_case_unea + CASE_INFO_ARRAY(clt_rsdi_income, all_clts) + CASE_INFO_ARRAY(clt_ssi_income, all_clts) + CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) + CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts)
		Next
	Else
		total_case_unea = total_case_unea & ""
		For each_unea = 0 to UBOUND(CASE_UNEA_ARRAY, 2)
			CASE_UNEA_ARRAY(rsdi_amt, each_unea)    = CASE_UNEA_ARRAY(rsdi_amt, each_unea) & ""
			CASE_UNEA_ARRAY(ssi_amt, each_unea)     = CASE_UNEA_ARRAY(ssi_amt, each_unea) & ""
			CASE_UNEA_ARRAY(other_1_amt, each_unea) = CASE_UNEA_ARRAY(other_1_amt, each_unea) & ""
			CASE_UNEA_ARRAY(other_2_amt, each_unea) = CASE_UNEA_ARRAY(other_2_amt, each_unea) & ""
		Next

		Do
			Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 555, 60 + (20 * UBOUND(CASE_UNEA_ARRAY, 2)), "Unearned Income"
			  Text 5, 5, 50, 10, "Person"
			  Text 170, 5, 35, 10, "RSDI"
			  Text 220, 5, 30, 10, "SSI"
			  Text 275, 5, 50, 10, "Other - 1"
			  Text 425, 5, 50, 10, "Other - 2"
			  ButtonGroup ButtonPressed
				PushButton 540, 5, 10, 10, "+", plus_button
			  For each_unea = 0 to UBOUND(CASE_UNEA_ARRAY, 2)
				  EditBox 5, 20 + (20 * each_unea), 150, 15, CASE_UNEA_ARRAY(unea_person, each_unea)
				  EditBox 170, 20 + (20 * each_unea), 40, 15, CASE_UNEA_ARRAY(rsdi_amt, each_unea)
				  EditBox 220, 20 + (20 * each_unea), 40, 15, CASE_UNEA_ARRAY(ssi_amt, each_unea)
				  ComboBox 275, 20 + (20 * each_unea), 80, 45, ""+chr(9)+"Other"+chr(9)+"Child Support"+chr(9)+"SSI"+chr(9)+"RSDI"+chr(9)+"Non-MN PA"+chr(9)+"VA Disability Benefit"+chr(9)+"VA Pension"+chr(9)+"VA Other"+chr(9)+"VA Aid & Attendance"+chr(9)+"Unemployment Insurance"+chr(9)+"Worker's Comp"+chr(9)+"Railroad Retirement"+chr(9)+"Other Retirement"+chr(9)+"Military Allotment"+chr(9)+"FC Child Requesting FS"+chr(9)+"FC Child Not Req FS"+chr(9)+"FC Adult Requesting FS"+chr(9)+"FC Adult Not Req FS"+chr(9)+"Dividends"+chr(9)+"Interest"+chr(9)+"Cnt Gifts Or Prizes"+chr(9)+"Strike Benefit 27 Contract For Deed"+chr(9)+"Illegal Income"+chr(9)+"Infrequent <30 Not Counted"+chr(9)+"Other FS Only"+chr(9)+"Infreq <= $20 MSA Exclusion"+chr(9)+"Direct Spousal Support"+chr(9)+"Disbursed Spousal Sup"+chr(9)+"Disbursed Spsl Sup Arrears"+chr(9)+"County 88 Gaming", CASE_UNEA_ARRAY(other_1_type, each_unea)
				  EditBox 360, 20 + (20 * each_unea), 40, 15, CASE_UNEA_ARRAY(other_1_amt, each_unea)
				  ComboBox 425, 20 + (20 * each_unea), 80, 45, ""+chr(9)+"Other"+chr(9)+"Child Support"+chr(9)+"SSI"+chr(9)+"RSDI"+chr(9)+"Non-MN PA"+chr(9)+"VA Disability Benefit"+chr(9)+"VA Pension"+chr(9)+"VA Other"+chr(9)+"VA Aid & Attendance"+chr(9)+"Unemployment Insurance"+chr(9)+"Worker's Comp"+chr(9)+"Railroad Retirement"+chr(9)+"Other Retirement"+chr(9)+"Military Allotment"+chr(9)+"FC Child Requesting FS"+chr(9)+"FC Child Not Req FS"+chr(9)+"FC Adult Requesting FS"+chr(9)+"FC Adult Not Req FS"+chr(9)+"Dividends"+chr(9)+"Interest"+chr(9)+"Cnt Gifts Or Prizes"+chr(9)+"Strike Benefit 27 Contract For Deed"+chr(9)+"Illegal Income"+chr(9)+"Infrequent <30 Not Counted"+chr(9)+"Other FS Only"+chr(9)+"Infreq <= $20 MSA Exclusion"+chr(9)+"Direct Spousal Support"+chr(9)+"Disbursed Spousal Sup"+chr(9)+"Disbursed Spsl Sup Arrears"+chr(9)+"County 88 Gaming", CASE_UNEA_ARRAY(other_2_type, each_unea)
				  EditBox 510, 20 + (20 * each_unea), 40, 15, CASE_UNEA_ARRAY(other_2_amt, each_unea)
			  Next
			  ButtonGroup ButtonPressed
				OkButton 500, 40 + (20 * UBOUND(CASE_UNEA_ARRAY, 2)), 50, 15
			EndDialog

			Dialog Dialog1

			If ButtonPressed = plus_button Then
				add_another = Ubound(CASE_UNEA_ARRAY, 2) + 1
				ReDim Preserve CASE_UNEA_ARRAY (6, add_another)
			End If

		Loop until ButtonPressed = -1

		total_case_unea = 0
		For each_unea = 0 to UBOUND(CASE_UNEA_ARRAY, 2)
			If CASE_UNEA_ARRAY(rsdi_amt, each_unea)      = "" Then CASE_UNEA_ARRAY(rsdi_amt, each_unea) = 0
			If CASE_UNEA_ARRAY(ssi_amt, each_unea)       = "" Then CASE_UNEA_ARRAY(ssi_amt, each_unea) = 0
			If CASE_UNEA_ARRAY(other_1_amt, each_unea) 	 = "" Then CASE_UNEA_ARRAY(other_1_amt, each_unea) = 0
			If CASE_UNEA_ARRAY(other_2_amt, each_unea)   = "" Then CASE_UNEA_ARRAY(other_2_amt, each_unea) = 0

			CASE_UNEA_ARRAY(rsdi_amt, each_unea)    = CASE_UNEA_ARRAY(rsdi_amt, each_unea) * 1
			CASE_UNEA_ARRAY(ssi_amt, each_unea)     = CASE_UNEA_ARRAY(ssi_amt, each_unea) * 1
			CASE_UNEA_ARRAY(other_1_amt, each_unea) = CASE_UNEA_ARRAY(other_1_amt, each_unea) * 1
			CASE_UNEA_ARRAY(other_2_amt, each_unea) = CASE_UNEA_ARRAY(other_2_amt, each_unea) * 1

			total_case_unea = total_case_unea + CASE_UNEA_ARRAY(rsdi_amt, each_unea) + CASE_UNEA_ARRAY(ssi_amt, each_unea) + CASE_UNEA_ARRAY(other_1_amt, each_unea) + CASE_UNEA_ARRAY(other_2_amt, each_unea)
		Next
	End If

end function

'CALCULATE LUMP SUM

'ENTER LIQUID ASSETS
function ASSETS_BUTTON_PRESSED

	If no_case_number_checkbox = unchecked Then
		For all_clts = 0 to UBOUND(CASE_INFO_ARRAY, 2)
			CASE_INFO_ARRAY(clt_chk_acct, all_clts) = CASE_INFO_ARRAY(clt_chk_acct, all_clts) & ""
			CASE_INFO_ARRAY(clt_sav_acct, all_clts) = CASE_INFO_ARRAY(clt_sav_acct, all_clts) & ""
			CASE_INFO_ARRAY(clt_asset_other_bal, all_clts) = CASE_INFO_ARRAY(clt_asset_other_bal, all_clts) & ""

			CASE_INFO_ARRAY(clt_asset_total, all_clts) = CASE_INFO_ARRAY(clt_asset_total, all_clts) & ""
		Next

		total_liquid_assets = total_liquid_assets & ""

		Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 400, 60 + (20 * UBOUND(CASE_INFO_ARRAY, 2)), "Assets"
		  Text 5, 5, 50, 10, "Person"
		  Text 190, 5, 35, 10, "Checking"
		  Text 245, 5, 30, 10, "Savings"
		  Text 300, 5, 50, 10, "Other"
		  For all_clts = 0 to UBOUND(CASE_INFO_ARRAY, 2)
		  	Text 5, 20 + (20 * all_clts), 175, 10, CASE_INFO_ARRAY(clt_ref, all_clts) & " - " & CASE_INFO_ARRAY(clt_name, all_clts)
		  	EditBox 190, 20 + (20 * all_clts), 40, 15, CASE_INFO_ARRAY(clt_chk_acct, all_clts)
		  	EditBox 245, 20 + (20 * all_clts), 40, 15, CASE_INFO_ARRAY(clt_sav_acct, all_clts)
		  	ComboBox 300, 20 + (20 * all_clts), 60, 45, ""+chr(9)+"Debit Card"+chr(9)+"Cash", CASE_INFO_ARRAY(clt_asset_other_type, all_clts)
		  	EditBox 365, 20 + (20 * all_clts), 40, 15, CASE_INFO_ARRAY(clt_asset_other_bal, all_clts)
		  Next
		  ButtonGroup ButtonPressed
		    OkButton 345, 40 + (20 * UBOUND(CASE_INFO_ARRAY, 2)), 50, 15
		EndDialog

		Dialog Dialog1

		total_liquid_assets = 0
		For all_clts = 0 to UBOUND(CASE_INFO_ARRAY, 2)
			If CASE_INFO_ARRAY(clt_chk_acct, all_clts)        = "" Then CASE_INFO_ARRAY(clt_chk_acct, all_clts) = 0
			If CASE_INFO_ARRAY(clt_sav_acct, all_clts)        = "" Then CASE_INFO_ARRAY(clt_sav_acct, all_clts) = 0
			If CASE_INFO_ARRAY(clt_asset_other_bal, all_clts) = "" Then CASE_INFO_ARRAY(clt_asset_other_bal, all_clts) = 0

			CASE_INFO_ARRAY(clt_chk_acct, all_clts) = CASE_INFO_ARRAY(clt_chk_acct, all_clts) * 1
			CASE_INFO_ARRAY(clt_sav_acct, all_clts) = CASE_INFO_ARRAY(clt_sav_acct, all_clts) * 1
			CASE_INFO_ARRAY(clt_asset_other_bal, all_clts) = CASE_INFO_ARRAY(clt_asset_other_bal, all_clts) * 1

			CASE_INFO_ARRAY(clt_asset_total, all_clts) = CASE_INFO_ARRAY(clt_chk_acct, all_clts) + CASE_INFO_ARRAY(clt_sav_acct, all_clts) + CASE_INFO_ARRAY(clt_asset_other_bal, all_clts)
			total_liquid_assets = total_liquid_assets + CASE_INFO_ARRAY(clt_asset_total, all_clts)
		Next
	Else
		total_liquid_assets = total_liquid_assets & ""

		For accounts = 0 to UBOUND(CASE_ACCOUNTS_ARRAY, 2)
			CASE_ACCOUNTS_ARRAY(account_balance, accounts) = CASE_ACCOUNTS_ARRAY(account_balance, accounts) & ""
		Next

		Do
			Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 315, 60 + (20 * UBOUND(CASE_ACCOUNTS_ARRAY, 2)), "Assets"
			  Text 5, 5, 50, 10, "Person"
			  Text 190, 5, 35, 10, "Type"
			  Text 265, 5, 30, 10, "Amount"
			  ButtonGroup ButtonPressed
				PushButton 295, 5, 10, 10, "+", plus_button
			  For account = 0 to UBOUND(CASE_ACCOUNTS_ARRAY, 2)
			  	EditBox 5, 20 + (20 * account), 175, 15, CASE_ACCOUNTS_ARRAY(account_holder, account)
				ComboBox 190, 20 + (20 * account), 60, 75, ""+chr(9)+"Checking"+chr(9)+"Savings"+chr(9)+"Debit Card"+chr(9)+"Cash", CASE_ACCOUNTS_ARRAY(account_type, account)
				EditBox 265, 20 + (20 * account), 40, 15, CASE_ACCOUNTS_ARRAY(account_balance, account)
			  Next
			  ButtonGroup ButtonPressed
				OkButton 260, 40 + (20 * UBOUND(CASE_ACCOUNTS_ARRAY, 2)), 50, 15
			EndDialog

			Dialog Dialog1

			If ButtonPressed = plus_button Then
				add_another = Ubound(CASE_ACCOUNTS_ARRAY, 2) + 1
				ReDim Preserve CASE_ACCOUNTS_ARRAY (2, add_another)
			End If
		Loop until ButtonPressed = -1

		For accounts = 0 to UBOUND(CASE_ACCOUNTS_ARRAY, 2)
			If CASE_ACCOUNTS_ARRAY(account_balance, accounts) = "" Then CASE_ACCOUNTS_ARRAY(account_balance, accounts) = 0
			CASE_ACCOUNTS_ARRAY(account_balance, accounts) = CASE_ACCOUNTS_ARRAY(account_balance, accounts) * 1
		Next

		total_liquid_assets = 0
		For accounts = 0 to UBOUND(CASE_ACCOUNTS_ARRAY, 2)
			total_liquid_assets = total_liquid_assets + CASE_ACCOUNTS_ARRAY(account_balance, accounts)
		Next
	End If

end function

'ENTER OTHER ASSETS
function OTHER_ASSETS_BUTTON_PRESSED
	For security = 0 to  UBOUND(SECURITIES_ARRAY, 2)
		SECURITIES_ARRAY(security_value, security) = SECURITIES_ARRAY(security_value, security) & ""
		SECURITIES_ARRAY(security_withdrawl, security) = SECURITIES_ARRAY(security_withdrawl, security) & ""
	Next

	Do
		vehicle_extend = 20 * UBOUND(VEHICLE_ARRAY, 2)
		security_extend = 20 * UBOUND(SECURITIES_ARRAY, 2)
		dlg_len = 145 + vehicle_extend + security_extend

		Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 290, dlg_len, "Dialog"
		  GroupBox 5, 5, 260, 50 + vehicle_extend, "Vehicles"
		  ButtonGroup ButtonPressed
		    PushButton 270, 10, 15, 15, "+", add_vehicle_button
		  Text 15, 20, 25, 10, "Type"
		  Text 80, 20, 20, 10, "Year"
		  Text 115, 20, 20, 10, "Make"
		  Text 170, 20, 25, 10, "Model"
		  Text 225, 20, 25, 10, "Value"
		  vehicle_add = 0
		  For vehicle = 0 to UBOUND (VEHICLE_ARRAY, 2)
			DropListBox 15, 35 + vehicle_add, 55, 45, "Select One ..."+chr(9)+CARS_type_list, VEHICLE_ARRAY(vehicle_type, vehicle)
			EditBox 80, 35 + vehicle_add, 30, 15, VEHICLE_ARRAY(vehicle_year, vehicle)
		  	EditBox 115, 35 + vehicle_add, 50, 15, VEHICLE_ARRAY(vehicle_make, vehicle)
		  	EditBox 170, 35 + vehicle_add, 50, 15, VEHICLE_ARRAY(vehicle_model, vehicle)
		  	EditBox 225, 35 + vehicle_add, 35, 15, VEHICLE_ARRAY(vehicle_value, vehicle)
			vehicle_add = vehicle_add + 20
		  Next

		  GroupBox 5, 65 + vehicle_extend, 260, 50 + security_extend, "Securities"
		  ButtonGroup ButtonPressed
		    PushButton 270, 70 + vehicle_extend, 15, 15, "+", add_security_button
		  Text 15, 80 + vehicle_extend, 20, 10, "Type"
		  Text 80, 80 + vehicle_extend, 40, 10, "Description"
		  Text 150, 80 + vehicle_extend, 25, 10, "Value"
		  Text 200, 80 + vehicle_extend, 65, 10, "Withdrawl Penalty"
		  security_add = 0
		  For security = 0 to  UBOUND(SECURITIES_ARRAY, 2)
			DropListBox 15, 95 + vehicle_extend + security_add, 55, 45, "Select One ..."+chr(9)+SECU_type_list, SECURITIES_ARRAY(security_type, security)
		  	EditBox 80, 95 + vehicle_extend + security_add, 65, 15, SECURITIES_ARRAY(security_description, security)
		  	EditBox 150, 95 + vehicle_extend + security_add, 40, 15, SECURITIES_ARRAY(security_value, security)
		  	EditBox 200, 95 + vehicle_extend + security_add, 40, 15, SECURITIES_ARRAY(security_withdrawl, security)
			security_add = security_add + 20
		  Next
		  ButtonGroup ButtonPressed
		    OkButton 235, 125 + vehicle_extend + security_extend, 50, 15
		EndDialog

		Dialog Dialog1
		If ButtonPressed = add_vehicle_button Then
			one_more_vehicle = UBOUND(VEHICLE_ARRAY, 2) + 1
			ReDim Preserve VEHICLE_ARRAY(4, one_more_vehicle)
		End If
		If ButtonPressed = add_security_button Then
			one_more_security = UBOUND(SECURITIES_ARRAY, 2) + 1
			ReDim Preserve SECURITIES_ARRAY(3, one_more_security)
		End If
	Loop Until ButtonPressed = -1

	total_other_assets = 0
	For security = 0 to  UBOUND(SECURITIES_ARRAY, 2)
		If SECURITIES_ARRAY(security_value, security) = "" Then SECURITIES_ARRAY(security_value, security) = 0
		If SECURITIES_ARRAY(security_withdrawl, security) = "" Then SECURITIES_ARRAY(security_withdrawl, security) = 0

		SECURITIES_ARRAY(security_value, security) = SECURITIES_ARRAY(security_value, security) * 1
		SECURITIES_ARRAY(security_withdrawl, security) = SECURITIES_ARRAY(security_withdrawl, security) * 1

		total_other_assets = total_other_assets + SECURITIES_ARRAY(security_value, security) - SECURITIES_ARRAY(security_withdrawl, security)
	Next

end function

'SHELTER & UTILITIES TOGETHER
function SHELTER_BUTTON_PRESSED

	rent_expense = rent_expense & ""
	prop_tax_expense = prop_tax_expense & ""
	home_ins_expense = home_ins_expense & ""
	other_expense = other_expense & ""
	actual_utility_expense = actual_utility_expense & ""
	If subsidized_rent = TRUE Then subsidy_checkbox = checked

	Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 281, 85, "Shelter and Utilities Expense"
	  EditBox 75, 5, 50, 15, rent_expense
	  EditBox 75, 25, 50, 15, prop_tax_expense
	  EditBox 75, 45, 50, 15, home_ins_expense
	  EditBox 75, 65, 50, 15, other_expense
	  EditBox 225, 5, 50, 15, actual_utility_expense
	  CheckBox 140, 30, 40, 10, "Heat/AC", heat_ac_checkbox
	  CheckBox 190, 30, 40, 10, "Electric", electric_checkbox
	  CheckBox 240, 30, 35, 10, "Phone", phone_checkbox
	  CheckBox 140, 50, 115, 10, "Check Here if rent is subsidized", subsidy_checkbox
	  ButtonGroup ButtonPressed
	    OkButton 225, 65, 50, 15
	  Text 10, 10, 50, 10, "Rent/Mortgage:"
	  Text 10, 30, 50, 10, "Property Tax:"
	  Text 10, 50, 60, 10, "House Insurance:"
	  Text 10, 70, 50, 10, "Other:"
	  Text 140, 10, 75, 10, "Actual Utilities (DWP):"
	EndDialog

	Dialog Dialog1

	if rent_expense = "" then rent_expense = 0
	if prop_tax_expense = "" then prop_tax_expense = 0
	if home_ins_expense = "" then home_ins_expense = 0
	if other_expense = "" then other_expense = 0
	if actual_utility_expense = "" then actual_utility_expense = 0
	if subsidy_checkbox = checked Then subsidized_rent = TRUE
	if subsidy_checkbox = unchecked Then subsidized_rent = FALSE

	rent_expense = rent_expense * 1
	prop_tax_expense = prop_tax_expense * 1
	home_ins_expense = home_ins_expense * 1
	other_expense = other_expense * 1
	actual_utility_expense = actual_utility_expense * 1

end function

'CHILD CARE EXPENSE/ COURT ORDERED EXPENSES/FMED EXPENSE
function DCEX_COEX_FMED_BUTTON_PRESSED
	monthly_childcare_exp = monthly_childcare_exp & ""
	monthly_adultcare_exp = monthly_adultcare_exp & ""
	child_support_exp = child_support_exp & ""
	alimony_exp = alimony_exp & ""
	monthly_fmed_exp = monthly_fmed_exp & ""

	Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 211, 185, "Monthly Expenses Dialog"
	  GroupBox 5, 5, 200, 50, "Dependent Care Expenses"
	  Text 20, 20, 115, 10, "Total Monthly Child Care Expense"
	  EditBox 140, 15, 50, 15, monthly_childcare_exp
	  Text 20, 40, 115, 10, "Total Monthly Adult Care Expense"
	  EditBox 140, 35, 50, 15, monthly_adultcare_exp
	  GroupBox 5, 65, 200, 50, "Court Ordered Expenses"
	  Text 20, 80, 95, 10, "Monthly Child Support Paid"
	  EditBox 120, 75, 50, 15, child_support_exp
	  Text 20, 100, 95, 10, "Monthly Alimony Paid"
	  EditBox 120, 95, 50, 15, alimony_exp
	  GroupBox 5, 120, 200, 40, "Elderly/Disabled Medical Expenses"
	  Text 20, 140, 90, 10, "Monthly Medical Expense"
	  EditBox 120, 135, 50, 15, monthly_fmed_exp
	  ButtonGroup ButtonPressed
	    OkButton 155, 165, 50, 15
	EndDialog

	Dialog Dialog1

	If monthly_childcare_exp = "" Then monthly_childcare_exp = 0
	If monthly_adultcare_exp = "" Then monthly_adultcare_exp = 0
	If child_support_exp = "" Then child_support_exp = 0
	If alimony_exp = "" Then alimony_exp = 0
	If monthly_fmed_exp = "" Then monthly_fmed_exp = 0

	monthly_childcare_exp = monthly_childcare_exp * 1
	monthly_adultcare_exp = monthly_adultcare_exp * 1
	child_support_exp = child_support_exp * 1
	alimony_exp = alimony_exp * 1
	monthly_fmed_exp = monthly_fmed_exp * 1

end function


'EMPS/TIME/SANC
function PROGRAM_SPECIFIC_BUTTON_PRESSED
	fmed_expenses = fmed_expenses & ""

	Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 155, 180, "Program Information"
	  GroupBox 5, 5, 145, 35, "SNAP"
	  CheckBox 15, 20, 90, 10, "Elderly/Disabled Case", elderly_disabled_checkbox
	  'Text 25, 40, 85, 10, "FMED Expenses/month"
	  'EditBox 110, 35, 30, 15, fmed_expenses
	  GroupBox 5, 45, 145, 60, "MFIP"
	  If no_case_number_checkbox = unchecked Then Text 15, 90, 110, 10, "TIME - HH has used " & case_months & " Monhts"
	  Text 15, 55, 35, 10, "Sanctions"
	  CheckBox 70, 55, 30, 10, "10%", ten_percent_sanc_checkbox
	  CheckBox 70, 70, 50, 10, "30%", thirty_percent_sanc_checkbox
	  GroupBox 5, 110, 145, 45, "MSA/GA"
	  CheckBox 15, 125, 125, 10, "All people in Cash HH meet basis", cash_basis_met_checkbox
	  Text 25, 135, 125, 10, "of eligibility."
	  ButtonGroup ButtonPressed
	    OkButton 100, 160, 50, 15
	EndDialog

	Dialog Dialog1

	If elderly_disabled_checkbox = checked Then elderly_disa_case = True
	If elderly_disabled_checkbox = unchecked Then elderly_disa_case = False

	If fmed_expenses = "" Then fmed_expenses = 0
	fmed_expenses = fmed_expenses * 1
end function


function MSA_SPECIAL_NEEDS

	Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 311, 160, "MSA Special Needs"
	  CheckBox 10, 5, 70, 10, "Rep Payee ($25)", sn_rep_payee_checkbox
	  CheckBox 10, 20, 105, 10, "Guardian/Conservator ($100)", sn_guardian_checkbox
	  CheckBox 150, 5, 90, 10, "Restaraunt Meals ($68)", sn_restaraunt_meals_checkbox
	  CheckBox 150, 20, 105, 10, "Housing Assistance ($194)", sn_housing_assistance_checkbox
	  CheckBox 10, 55, 80, 10, "Anti-Dumping ($29.10)", sn_anti_dumping_checkbox
	  CheckBox 10, 70, 115, 10, "Controlled Protein (40-60) ($194)", sn_control_protien_60_checkbox
	  CheckBox 10, 85, 120, 10, "Controlled Protein (<40) ($242.50)", sn_control_protien_40_checkbox
	  CheckBox 10, 100, 80, 10, "Gluten Free ($48.50)", sn_gluten_free_checkbox
	  CheckBox 10, 115, 80, 10, "High Protein ($48.50)", sn_high_protien_checkbox
	  CheckBox 150, 40, 85, 10, "High Residue ($38.80)", sn_high_residue_checkbox
	  CheckBox 150, 55, 85, 10, "Hypoglycemic ($29.10)", sn_hypoglycemic_checkbox
	  CheckBox 150, 70, 70, 10, "Ketogenic ($48.50)", sn_ketogenic_checkbox
	  CheckBox 150, 85, 85, 10, "Lactose Free ($48.50)", sn_lactose_free_checkbox
	  CheckBox 150, 100, 95, 10, "Low Cholesterol ($48.50)", sn_low_cholesterol_checkbox
	  CheckBox 150, 115, 105, 10, "Pregnancy/Lactation ($67.90)", sn_pregnancy_lactation_checkbox
	  ButtonGroup ButtonPressed
	    OkButton 255, 140, 50, 15
	  GroupBox 5, 35, 300, 95, "Special Diets"
	EndDialog

	dialog Dialog1

end function
'Income Limits and Assistance Standards

function SET_INCOME_LIMITS(HH_size, SNAP_Info, Family_Cash_Info, Adult_Cash_Info, FPG_100_Amt, FPG_130_Amt, FPG_165_Amt, SNAP_assistance_standard, SNAP_Standard_Disregard, MF_Transitional_Standard, MF_Wage_Standard, MF_MF_Standard, MF_FS_Standard, GA_assistance_standard, MSA_assistance_standard)
	'FPG and Thrifty standards'
	Select Case HH_size
	Case 0
		If SNAP_Info = TRUE Then
			FPG_100_Amt = 0
			FPG_130_Amt = 0
			FPG_165_Amt = 0
			SNAP_assistance_standard = 0
			SNAP_Standard_Disregard = 0
		End If
		If Family_Cash_Info = TRUE Then
			MF_Transitional_Standard = 0
			MF_Wage_Standard = 0
			MF_MF_Standard = 0
			MF_FS_Standard = 0
		End If
		If Adult_Cash_Info = TRUE Then
			GA_assistance_standard = 0
			MSA_assistance_standard = 0
		End If
	Case 1
		If SNAP_Info = TRUE Then
			FPG_100_Amt = 990
			FPG_130_Amt = 1287
			FPG_165_Amt = 1634
			SNAP_assistance_standard = 194
			SNAP_Standard_Disregard = 157
		End If
		If Family_Cash_Info = TRUE Then
			MF_Transitional_Standard = 422
			MF_Wage_Standard = 464
			MF_MF_Standard = 250
			MF_FS_Standard = 172
		End If
		If Adult_Cash_Info = TRUE Then
			GA_assistance_standard = 203
			MSA_assistance_standard = 796
		End If
	Case 2
		If SNAP_Info = TRUE Then
			FPG_100_Amt = 1335
			FPG_130_Amt = 1736
			FPG_165_Amt = 2203
			SNAP_assistance_standard = 357
			SNAP_Standard_Disregard = 157
		End If
		If Family_Cash_Info = TRUE Then
			MF_Transitional_Standard = 754
			MF_Wage_Standard = 829
			MF_MF_Standard = 437
			MF_FS_Standard = 317
		End If
		If Adult_Cash_Info = TRUE Then
			GA_assistance_standard = 260
			MSA_assistance_standard = 1194
		End If
	Case 3
		If SNAP_Info = TRUE Then
			FPG_100_Amt = 1680
			FPG_130_Amt = 2184
			FPG_165_Amt = 2772
			SNAP_assistance_standard = 511
			SNAP_Standard_Disregard = 157
		End If
		If Family_Cash_Info = TRUE Then
			MF_Transitional_Standard = 991
			MF_Wage_Standard = 1090
			MF_MF_Standard = 532
			MF_FS_Standard = 459
		End If
	Case 4
		If SNAP_Info = TRUE Then
			FPG_100_Amt = 2025
			FPG_130_Amt = 2633
			FPG_165_Amt = 3342
			SNAP_assistance_standard = 649
			SNAP_Standard_Disregard = 168
		End If
		If Family_Cash_Info = TRUE Then
			MF_Transitional_Standard = 1207
			MF_Wage_Standard = 1328
			MF_MF_Standard = 621
			MF_FS_Standard = 586
		End If
	Case 5
		If SNAP_Info = TRUE Then
			FPG_100_Amt = 2370
			FPG_130_Amt = 3081
			FPG_165_Amt = 3911
			SNAP_assistance_standard = 771
			SNAP_Standard_Disregard = 197
		End If
		If Family_Cash_Info = TRUE Then
			MF_Transitional_Standard = 1395
			MF_Wage_Standard = 1535
			MF_MF_Standard = 697
			MF_FS_Standard = 698
		End If
	Case 6
		If SNAP_Info = TRUE Then
			FPG_100_Amt = 2715
			FPG_130_Amt = 3530
			FPG_165_Amt = 4480
			SNAP_assistance_standard = 925
			SNAP_Standard_Disregard = 226
		End If
		If Family_Cash_Info = TRUE Then
			MF_Transitional_Standard = 1605
			MF_Wage_Standard = 1766
			MF_MF_Standard = 773
			MF_FS_Standard = 832
		End If
	Case 7
		If SNAP_Info = TRUE Then
			FPG_100_Amt = 3061
			FPG_130_Amt = 3980
			FPG_165_Amt = 5051
			SNAP_assistance_standard = 1022
			SNAP_Standard_Disregard = 226
		End If
		If Family_Cash_Info = TRUE Then
			MF_Transitional_Standard = 1748
			MF_Wage_Standard = 1923
			MF_MF_Standard = 850
			MF_FS_Standard = 898
		End If
	Case 8
		If SNAP_Info = TRUE Then
			FPG_100_Amt = 3408
			FPG_130_Amt = 4430
			FPG_165_Amt = 5623
			SNAP_assistance_standard = 1169
			SNAP_Standard_Disregard = 226
		End If
		If Family_Cash_Info = TRUE Then
			MF_Transitional_Standard = 1931
			MF_Wage_Standard = 2124
			MF_MF_Standard = 916
			MF_FS_Standard = 1015
		End If
	Case 9
		If SNAP_Info = TRUE Then
			FPG_100_Amt = 3755
			FPG_130_Amt = 4881
			FPG_165_Amt = 6195
			SNAP_assistance_standard = 1315
			SNAP_Standard_Disregard = 226
		End If
		If Family_Cash_Info = TRUE Then
			MF_Transitional_Standard = 2113
			MF_Wage_Standard = 2324
			MF_MF_Standard = 980
			MF_FS_Standard = 1133
		End If
	Case 10
		If SNAP_Info = TRUE Then
			FPG_100_Amt = 4102
			FPG_130_Amt = 5332
			FPG_165_Amt = 6767
			SNAP_assistance_standard = 1461
			SNAP_Standard_Disregard = 226
		End If
		If Family_Cash_Info = TRUE Then
			MF_Transitional_Standard = 2288
			MF_Wage_Standard = 2517
			MF_MF_Standard = 1035
			MF_FS_Standard = 1253
		End If
	Case 11
		If SNAP_Info = TRUE Then
			FPG_100_Amt = 4449
			FPG_130_Amt = 5783
			FPG_165_Amt = 7339
			SNAP_assistance_standard = 1607
			SNAP_Standard_Disregard = 226
		End If
		If Family_Cash_Info = TRUE Then
			MF_Transitional_Standard = 2462
			MF_Wage_Standard = 2708
			MF_MF_Standard = 1088
			MF_FS_Standard = 1374
		End If
	Case 12
		If SNAP_Info = TRUE Then
			FPG_100_Amt = 4796
			FPG_130_Amt = 6234
			FPG_165_Amt = 7911
			SNAP_assistance_standard = 1753
			SNAP_Standard_Disregard = 226
		End If
		If Family_Cash_Info = TRUE Then
			MF_Transitional_Standard = 2636
			MF_Wage_Standard = 2899
			MF_MF_Standard = 1141
			MF_FS_Standard = 1495
		End If
	Case 13
		If SNAP_Info = TRUE Then
			FPG_100_Amt = 5143
			FPG_130_Amt = 6685
			FPG_165_Amt = 8483
			SNAP_assistance_standard = 1899
			SNAP_Standard_Disregard = 226
		End If
		If Family_Cash_Info = TRUE Then
			MF_Transitional_Standard = 2810
			MF_Wage_Standard = 3090
			MF_MF_Standard = 1194
			MF_FS_Standard = 1616
		End If
	Case 14
		If SNAP_Info = TRUE Then
			FPG_100_Amt = 5490
			FPG_130_Amt = 7136
			FPG_165_Amt = 9055
			SNAP_assistance_standard = 2045
			SNAP_Standard_Disregard = 226
		End If
		If Family_Cash_Info = TRUE Then
			MF_Transitional_Standard = 2984
			MF_Wage_Standard = 3281
			MF_MF_Standard = 1247
			MF_FS_Standard = 1737
		End If
	Case 15
		If SNAP_Info = TRUE Then
			FPG_100_Amt = 5837
			FPG_130_Amt = 7587
			FPG_165_Amt = 9627
			SNAP_assistance_standard = 2191
			SNAP_Standard_Disregard = 226
		End If
		If Family_Cash_Info = TRUE Then
			MF_Transitional_Standard = 3158
			MF_Wage_Standard = 3472
			MF_MF_Standard = 1300
			MF_FS_Standard = 1858
		End If
	Case 16
		If SNAP_Info = TRUE Then
			FPG_100_Amt = 6184
			FPG_130_Amt = 8038
			FPG_165_Amt = 10199
			SNAP_assistance_standard = 2337
			SNAP_Standard_Disregard = 226
		End If
		If Family_Cash_Info = TRUE Then
			MF_Transitional_Standard = 3332
			MF_Wage_Standard = 3663
			MF_MF_Standard = 1353
			MF_FS_Standard = 1979
		End If
	Case 17
		If SNAP_Info = TRUE Then
			FPG_100_Amt = 6531
			FPG_130_Amt = 8489
			FPG_165_Amt = 10771
			SNAP_assistance_standard = 2483
			SNAP_Standard_Disregard = 226
		End If
		If Family_Cash_Info = TRUE Then
			MF_Transitional_Standard = 3506
			MF_Wage_Standard = 3854
			MF_MF_Standard = 1406
			MF_FS_Standard = 2100
		End If
	Case 18
		If SNAP_Info = TRUE Then
			FPG_100_Amt = 6878
			FPG_130_Amt = 8940
			FPG_165_Amt = 11343
			SNAP_assistance_standard = 2629
			SNAP_Standard_Disregard = 226
		End If
		If Family_Cash_Info = TRUE Then
			MF_Transitional_Standard = 3680
			MF_Wage_Standard = 4045
			MF_MF_Standard = 1459
			MF_FS_Standard = 2221
		End If
	Case 19
		If SNAP_Info = TRUE Then
			FPG_100_Amt = 7225
			FPG_130_Amt = 9391
			FPG_165_Amt = 11915
			SNAP_assistance_standard = 2775
			SNAP_Standard_Disregard = 226
		End If
		If Family_Cash_Info = TRUE Then
			MF_Transitional_Standard = 3854
			MF_Wage_Standard = 4236
			MF_MF_Standard = 1512
			MF_FS_Standard = 2342
		End If
	Case 20
		If SNAP_Info = TRUE Then
			FPG_100_Amt = 7572
			FPG_130_Amt = 9842
			FPG_165_Amt = 12487
			SNAP_assistance_standard = 2921
			SNAP_Standard_Disregard = 226
		End If
		If Family_Cash_Info = TRUE Then
			MF_Transitional_Standard = 4028
			MF_Wage_Standard = 4427
			MF_MF_Standard = 1565
			MF_FS_Standard = 2463
		End If
	End Select
end function


'FUNCTION TO CALCULATE PROGRAM ELIGIBILITY - generates amounts and notes for EACH program

function ProgramEstimate
	total_case_assets = total_liquid_assets + total_other_assets
	If total_case_assets = "" Then total_case_assets = 0
	total_case_assets = total_case_assets * 1
	'FAMILY CASH
	If FAMILY_CASE = TRUE Then
		Counted_MFIP_Gross_Income = 0
		Counted_MFIP_UNEA_Income = 0
		Counted_MFIP_JOBS_Income = 0
		Counted_MFIP_Liquid_Assets = 0
		all_child_support = 0
		If no_case_number_checkbox = unchecked Then family_cash_hh_size = 0
		children_in_hh = 0
		Estimated_MF_MF = 0
		Estimated_MF_FS = 0
		Estimated_MF_HG = 0
		family_cash_estimated_benefit = ""
		family_cash_notes = ""
		potentially_family_cash_eligible = TRUE
		child_support_income_on_case = FALSE

		If no_case_number_checkbox = unchecked Then
			For all_clients = 0 to UBOUND(CASE_INFO_ARRAY, 2)
				If CASE_INFO_ARRAY(include_family_cash, all_clients) = checked Then
					family_cash_hh_size = family_cash_hh_size + 1
					If CASE_INFO_ARRAY(clt_age, all_clients) <=18 Then children_in_hh = children_in_hh + 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clients) = "Child Support" Then
					 	all_child_support = all_child_support + CASE_INFO_ARRAY(clt_other_unea_1_amt , all_clients)
						child_support_income_on_case = TRUE
					End If
					If CASE_INFO_ARRAY(clt_other_unea_2_type, all_clients) = "Child Support" Then
						all_child_support = all_child_support + CASE_INFO_ARRAY(clt_other_unea_1_amt , all_clients)
						child_support_income_on_case = TRUE
					End If
					Counted_MFIP_UNEA_Income = Counted_MFIP_UNEA_Income + CASE_INFO_ARRAY(clt_ssi_income , all_clients) + CASE_INFO_ARRAY(clt_rsdi_income , all_clients) + CASE_INFO_ARRAY(clt_other_unea_1_amt , all_clients) + CASE_INFO_ARRAY(clt_other_unea_2_amt , all_clients)
					Counted_MFIP_Liquid_Assets = Counted_MFIP_Liquid_Assets + CASE_INFO_ARRAY(clt_asset_total , all_clients)
				End If
			Next
		End If

		If family_cash_hh_size = 0 Then
			potentially_family_cash_eligible = False
			family_cash_estimated_benefit = "INELIGIBLE"
			family_cash_notes = "No HH members included in Family Cash case."
		End if

		If total_case_assets > 10000 AND potentially_family_cash_eligible = TRUE then
			potentially_family_cash_eligible = False
			family_cash_estimated_benefit = "INELIGIBLE"
			family_cash_notes = "Assets appear to exceed $10,000 - no cash eligibility."
		End If

		Call SET_INCOME_LIMITS(family_cash_hh_size, false, true, false, FPG_100_Amt, FPG_130_Amt, FPG_165_Amt, SNAP_assistance_standard, SNAP_Standard_Disregard, MF_Transitional_Standard, MF_Wage_Standard, MF_MF_Standard, MF_FS_Standard, GA_assistance_standard, MSA_assistance_standard)

		DETERMINE_COUNTED_EI

		Counted_MFIP_JOBS_Income = (Family_Cash_JOBS_Income - 65)/2

		If Counted_MFIP_JOBS_Income < 0 Then Counted_MFIP_JOBS_Income = 0
		Monthly_Need = MF_Wage_Standard - Counted_MFIP_JOBS_Income
		If Monthly_Need <= 0 AND potentially_family_cash_eligible = TRUE Then
			potentially_family_cash_eligible = False
			family_cash_estimated_benefit = "INELIGIBLE"
			family_cash_notes = "Earned income exceeds the Family Wage Level."
		End If

		If potentially_family_cash_eligible = TRUE Then
		 	If MF_Transitional_Standard < Monthly_Need Then Monthly_Need = MF_Transitional_Standard
			Monthly_Need = Monthly_Need - Counted_MFIP_UNEA_Income
			Monthly_Need = Int(Monthly_Need)
			If child_support_income_on_case = TRUE Then
				If children_in_hh = 1 Then child_support_exclusion = 175
				If children_in_hh >= 2 Then child_support_exclusion = 200
				If children_in_hh = 0 Then child_support_exclusion = 0
				If all_child_support < child_support_exclusion Then child_support_exclusion = all_child_support
				Monthly_Need = Monthly_Need + child_support_exclusion
			End If
			If subsidized_rent = TRUE Then Monthly_Need = Monthly_Need - 50
			If Monthly_Need <= 0 Then
				potentially_family_cash_eligible = False
				family_cash_estimated_benefit = "INELIGIBLE"
				family_cash_notes = "Total income exceeds the Transitional Standard."
			End If
		End If

		If potentially_family_cash_eligible = TRUE Then
			If Monthly_Need <= MF_FS_Standard Then
				Estimated_MF_FS = Monthly_Need
				Estimated_MF_MF = 0
				Estimated_MF_HG = 110
			Else
				Estimated_MF_FS = MF_FS_Standard
				Monthly_Need = Monthly_Need - MF_FS_Standard
				Estimated_MF_MF = Monthly_Need
				Estimated_MF_HG = 110
			End If

			If subsidized_rent = TRUE Then
				Estimated_MF_HG = 0
				family_cash_notes = family_cash_notes & " Rent subsidy is reducing grant by $50 & removing housing grant."
			End If

			family_cash_estimated_benefit = "MF-MF - $" & Estimated_MF_MF & " MF-FS - $" & Estimated_MF_FS & " MF-HG - $" & Estimated_MF_HG
		End If
	Else
		family_cash_estimated_benefit = "INELIGIBLE"
		family_cash_notes = "This appears to be an adult case."
	End If

	'ADULT CASH
''	If ADULT_CASE = TRUE Then
		Counted_GA_Gross_Income = 0
		Counted_GA_Liquid_Assets = 0
		If no_case_number_checkbox = unchecked Then adult_cash_hh_size = 0
		adult_cash_estimated_benefit = ""
		adult_cash_notes = ""
		potentially_adult_cash_eligible = TRUE
		GA_benefit = TRUE
		MSA_benefit = FALSE

		If no_case_number_checkbox = unchecked Then
			For all_clients = 0 to UBOUND(CASE_INFO_ARRAY, 2)
				If CASE_INFO_ARRAY(include_adult_cash, all_clients) = checked Then
					adult_cash_hh_size = adult_cash_hh_size + 1
					Counted_GA_UNEA_Income = CASE_INFO_ARRAY(clt_ssi_income , all_clients) + CASE_INFO_ARRAY(clt_rsdi_income , all_clients) + CASE_INFO_ARRAY(clt_other_unea_1_amt , all_clients) + CASE_INFO_ARRAY(clt_other_unea_2_amt , all_clients)
					Counted_GA_Liquid_Assets = Counted_SNAP_Liquid_Assets + CASE_INFO_ARRAY(clt_asset_total , all_clients)
					If CASE_INFO_ARRAY(clt_ssi_income , all_clients) <> 0 Then
						MSA_benefit = TRUE
						GA_benefit = FALSE
					End If
				End If
			Next
		End If

		If adult_cash_hh_size > 2 Then
		 	MsgBox "This script does not support Adult Cash budget with a household size of more than 2. Change the MSA/GA included clients to have the script estimate a budget for MSA or GA"
			adult_cash_hh_size = 0
		End If

		If adult_cash_hh_size = 0 Then
			potentially_adult_cash_eligible = False
			adult_cash_estimated_benefit = "INELIGIBLE"
			adult_cash_notes = "No HH members included in Adult Cash grant."
		End If

		If total_case_assets > 10000 AND potentially_adult_cash_eligible = TRUE then
			potentially_adult_cash_eligible = False
			adult_cash_estimated_benefit = "INELIGIBLE"
			adult_cash_notes = "Assets appear to exceed $10,000 - no cash eligibility."
		End If

		If cash_basis_met_checkbox = unchecked AND potentially_adult_cash_eligible = TRUE Then
			potentially_adult_cash_eligible = False
			adult_cash_estimated_benefit = "INELIGIBLE"
			adult_cash_notes = "It does not appear that anyone in the adult cash unit meet the basis of eligibility. If this is not the case press the DISA button and indicate that basis of eiligibility has been met."
		End If

		If potentially_adult_cash_eligible = TRUE Then
			Call SET_INCOME_LIMITS(adult_cash_hh_size, false, false, true, FPG_100_Amt, FPG_130_Amt, FPG_165_Amt, SNAP_assistance_standard, SNAP_Standard_Disregard, MF_Transitional_Standard, MF_Wage_Standard, MF_MF_Standard, MF_FS_Standard, GA_assistance_standard, MSA_assistance_standard)

			DETERMINE_COUNTED_EI

			If sn_rep_payee_checkbox 			= checked Then
				ten_percent = .1 * (Counted_GA_UNEA_Income + Counted_GA_JOBS_Income)
				If ten_percent < 25 Then
					rep_payee_amt = ten_percent
				Else
					rep_payee_amt = 25
				End If
				MSA_assistance_standard = MSA_assistance_standard + rep_payee_amt
			End If


			If sn_guardian_checkbox 			= checked Then
				five_percent = .05 * (Counted_GA_UNEA_Income + Counted_GA_JOBS_Income)
				If five_percent < 100 Then
					guardian_amt = five_percent
				Else
					guardian_amt = 100
				End If
				MSA_assistance_standard = MSA_assistance_standard + guardian_amt
			End If

			If sn_restaraunt_meals_checkbox		= checked Then MSA_assistance_standard = MSA_assistance_standard + 68
			If sn_housing_assistance_checkbox	= checked Then MSA_assistance_standard = MSA_assistance_standard + 194
			If sn_anti_dumping_checkbox			= checked Then MSA_assistance_standard = MSA_assistance_standard + 29.10
			If sn_control_protien_60_checkbox	= checked Then MSA_assistance_standard = MSA_assistance_standard + 194
			If sn_control_protien_40_checkbox	= checked Then MSA_assistance_standard = MSA_assistance_standard + 242.50
			If sn_gluten_free_checkbox			= checked Then MSA_assistance_standard = MSA_assistance_standard + 48.50
			If sn_high_protien_checkbox			= checked Then MSA_assistance_standard = MSA_assistance_standard + 48.50
			If sn_high_residue_checkbox			= checked Then MSA_assistance_standard = MSA_assistance_standard + 38.80
			If sn_hypoglycemic_checkbox			= checked Then MSA_assistance_standard = MSA_assistance_standard + 29.10
			If sn_ketogenic_checkbox			= checked Then MSA_assistance_standard = MSA_assistance_standard + 48.50
			If sn_lactose_free_checkbox			= checked Then MSA_assistance_standard = MSA_assistance_standard + 48.50
			If sn_low_cholesterol_checkbox		= checked Then MSA_assistance_standard = MSA_assistance_standard + 48.50
			If sn_pregnancy_lactation_checkbox	= checked Then MSA_assistance_standard = MSA_assistance_standard + 67.90

			ssa_income = 0
			For all_clients = 0 to UBOUND(CASE_INFO_ARRAY, 2)
				ssa_income = ssa_income + CASE_INFO_ARRAY(clt_rsdi_income , all_clients) + CASE_INFO_ARRAY(clt_ssi_income , all_clients)
			Next

			If ssa_income <> 0 AND ssa_income < MSA_assistance_standard Then
				MSA_benefit = TRUE
				GA_benefit = FALSE
			Else
				GA_benefit = TRUE
				MSA_benefit = FALSE
			End If

			If GA_benefit = TRUE then
				Counted_GA_JOBS_Income = (Adult_Cash_JOBS_Income - 65)/2

				If Counted_GA_JOBS_Income < 0 Then Counted_GA_JOBS_Income = 0
				Total_GA_Countable_Income = Counted_GA_JOBS_Income + Counted_GA_UNEA_Income
				GA_Subtotal = GA_assistance_standard - Total_GA_Countable_Income

				If GA_Subtotal < 0 Then GA_Subtotal = 0
				If GA_Subtotal = 0 Then
					potentially_adult_cash_eligible = False
					adult_cash_estimated_benefit = "INELIGILE"
					adult_cash_notes = adult_cash_notes & "Income Exceeds the GA Standard."
				End If
			End If

			If MSA_benefit = TRUE Then
				Counted_MSA_UNEA_Income = Counted_GA_UNEA_Income - 20
				If Counted_MSA_UNEA_Income < 0 Then Counted_MSA_UNEA_Income = 0

				Counted_GA_JOBS_Income = (Adult_Cash_JOBS_Income - 65)/2
				If Counted_GA_JOBS_Income < 0 Then Counted_GA_JOBS_Income = 0

				MSA_Net_Income = Counted_MSA_UNEA_Income + Counted_GA_JOBS_Income
				MSA_issuance = MSA_assistance_standard - MSA_Net_Income

				If MSA_issuance < 0 Then MSA_issuance = 0
				If MSA_issuance = 0 Then
					potentially_adult_cash_eligible = False
					adult_cash_estimated_benefit = "INELIGILE"
					adult_cash_notes = adult_cash_notes & "Income Exceeds the MSA Standard of $" & MSA_assistance_standard
				End If
			End If
		End If

		If potentially_adult_cash_eligible = TRUE Then
			If GA_benefit = TRUE Then
				adult_cash_estimated_benefit = "GA $" & GA_Subtotal
			End If

			If MSA_benefit = TRUE Then
				adult_cash_estimated_benefit = "MSA $" & MSA_issuance
			End If

		End If

''	Else
''		adult_cash_estimated_benefit = "INELIGIBLE"
''		adult_cash_notes = "This appears to be a family case."
''	End If

	'SNAP
	Counted_SNAP_Gross_Income = 0
	Counted_SNAP_Liquid_Assets = 0
	Counted_SNAP_Net_Income = 0
	Net_adjusted_SNAP_income = 0
	shelter_expense = 0
	utilities_expenses = 0
	total_shelter_expense = 0
	If no_case_number_checkbox = unchecked Then snap_hh_size = 0
	total_shelter = 0
	total_utilities = 0
	snap_estimated_benefit = ""
	potentially_snap_eligible = TRUE
	snap_elig_msg = ""
	SNAP_notes = ""

	If no_case_number_checkbox = unchecked Then
		For all_clients = 0 to UBOUND(CASE_INFO_ARRAY, 2)
			If CASE_INFO_ARRAY(include_snap, all_clients) = checked Then
				snap_hh_size = snap_hh_size + 1
				'If CASE_INFO_ARRAY(clt_age, all_clients) >=18 Then Counted_SNAP_Gross_Income = Counted_SNAP_Gross_Income + CASE_INFO_ARRAY(clt_ei_gross , all_clients)
				Counted_SNAP_Gross_Income = Counted_SNAP_Gross_Income + CASE_INFO_ARRAY(clt_ssi_income , all_clients) + CASE_INFO_ARRAY(clt_rsdi_income , all_clients) + CASE_INFO_ARRAY(clt_other_unea_1_amt , all_clients) + CASE_INFO_ARRAY(clt_other_unea_2_amt , all_clients)
				Counted_SNAP_Liquid_Assets = Counted_SNAP_Liquid_Assets + CASE_INFO_ARRAY(clt_asset_total , all_clients)
			End If
		Next
	End If
	Counted_SNAP_Gross_Income = Counted_SNAP_Gross_Income + SNAP_JOBS_Income + MSA_issuance + GA_Subtotal

	Call SET_INCOME_LIMITS(snap_hh_size, true, false, false, FPG_100_Amt, FPG_130_Amt, FPG_165_Amt, SNAP_assistance_standard, SNAP_Standard_Disregard, MF_Transitional_Standard, MF_Wage_Standard, MF_MF_Standard, MF_FS_Standard, GA_assistance_standard, MSA_assistance_standard)

	If snap_hh_size = 0 Then
		potentially_snap_eligible = FALSE
		snap_elig_msg = snap_elig_msg & "No HH Member included in SNAP Case; "
	End If

	If Counted_SNAP_Gross_Income > FPG_165_Amt Then
		potentially_snap_eligible = FALSE
		snap_elig_msg = snap_elig_msg & "Gross Income exceeds 165%; "
	End If

	If potentially_snap_eligible = TRUE Then
		If elderly_disabled_checkbox = checked Then elderly_disa_case = True
		If elderly_disabled_checkbox = unchecked Then elderly_disa_case = False
		SNAP_EI_Disregard = .2 * SNAP_JOBS_Income
		shelter_expense = rent_expense + prop_tax_expense + home_ins_expense + other_expense
		'Logic for figuring out utils. The highest priority for the if...then is heat/AC, followed by electric and phone, followed by phone and electric separately.
		If heat_ac_checkbox = checked then
			utilities_expenses = heat_AC_amt
		ElseIf electric_checkbox = checked and phone_checkbox = checked then
			utilities_expenses = phone_amt + electric_amt					'Phone standard plus electric standard.
		ElseIf phone_checkbox = checked and electric_checkbox = unchecked then
			utilities_expenses = phone_amt
		ElseIf electric_checkbox = checked and phone_checkbox = unchecked then
			utilities_expenses = electric_amt
		End if
		total_shelter_expense = shelter_expense + utilities_expenses
		total_shelter_expense = Round(total_shelter_expense)
	''	MsgBox "TOTAL SHEL $" & shelter_expense & vbnewLine & "RENT $" & rent_expense &vbnewLine & "PROP TAX $" & prop_tax_expense & vbnewLine & "HOME INS $" & home_ins_expense & vbnewLine & "OTHER $" & other_expense & vbnewLine & "UTILITIES $" & utilities_expenses


		Counted_SNAP_Net_Income = Counted_SNAP_Gross_Income - SNAP_Standard_Disregard - SNAP_EI_Disregard - monthly_childcare_exp - monthly_adultcare_exp - child_support_exp - alimony_exp - monthly_fmed_exp
		If Counted_SNAP_Net_Income < 0 Then Counted_SNAP_Net_Income = 0
	''		MsgBox "Counted SNAP Income - " & Counted_SNAP_Net_Income
		adjusted_shelter_costs = total_shelter_expense - Round((Counted_SNAP_Net_Income/2))
		If adjusted_shelter_costs < 0 Then adjusted_shelter_costs = 0
	''		MsgBox "Adjusted Shelter Costs - " &adjusted_shelter_costs
		If adjusted_shelter_costs < 517 Then
			counted_shelter_costs = adjusted_shelter_costs
		ElseIf elderly_disa_case = True Then
			counted_shelter_costs = adjusted_shelter_costs
		Else
			counted_shelter_costs = 517
		End If

		Net_adjusted_SNAP_income = Counted_SNAP_Net_Income - counted_shelter_costs
		If Net_adjusted_SNAP_income < 0 Then Net_adjusted_SNAP_income = 0
		If Net_adjusted_SNAP_income > FPG_100_Amt Then
			potentially_snap_eligible = FALSE
			snap_elig_msg = snap_elig_msg & "Net Income exceeds 100%; "
		End If
	End If

	If potentially_snap_eligible = TRUE Then
	''		MsgBox "Assistance Standard - " & SNAP_assistance_standard
	''		MsgBox "30% of Income - " & Net_adjusted_SNAP_income * .3
		standard_reduction = Net_adjusted_SNAP_income * .3
		If standard_reduction < 0 Then standard_reduction = 0
		potential_SNAP_benefit = SNAP_assistance_standard - standard_reduction
		If potential_SNAP_benefit < 0 Then potential_SNAP_benefit = 0
		potential_SNAP_benefit = Int(potential_SNAP_benefit)

		snap_estimated_benefit = "Eligible for : $" & potential_SNAP_benefit

		SNAP_notes = SNAP_notes & "SNAP Gross Income in budget: $" & Counted_SNAP_Gross_Income
		If potentially_adult_cash_eligible = TRUE Then
			If GA_Subtotal <> "" Then SNAP_notes = SNAP_notes & ", Estimated GA benefit included in SNAP estimated budget."
			If MSA_issuance <> "" Then SNAP_notes = SNAP_notes & ", Estimated MSA benefit included in SNAP estimated budget."
		End If

	''	MsgBox "Total SHELTER: $" & total_shelter_expense & vbNewLine & "Counted Shelter costs: $" & counted_shelter_costs & vbNewLine & "TOTAL Income: $" & Counted_SNAP_Net_Income & vbNewLine & "Net Adjusted Income; $" & Net_adjusted_SNAP_income & vbNewLine & "Standard Reduction: $" & standard_reduction
	Else
		snap_estimated_benefit = "INELIGIBLE"
		SNAP_notes = snap_elig_msg
	End If

end function

'SCRIPT======================================================================================================================

'Template Dialog
EMConnect ""

'Declaring our arrays - because life is easier with arrays
Dim CASE_INFO_ARRAY ()
ReDim CASE_INFO_ARRAY (18, 0)
Dim EI_ARRAY ()
ReDim EI_ARRAY (18, 0)
Dim VEHICLE_ARRAY ()
ReDim VEHICLE_ARRAY (4, 0)
Dim SECURITIES_ARRAY ()
ReDim SECURITIES_ARRAY (3, 0)
Dim CASE_ACCOUNTS_ARRAY()
ReDim CASE_ACCOUNTS_ARRAY(2, 0)
Dim CASE_UNEA_ARRAY()
ReDim CASE_UNEA_ARRAY(6, 0)

FPG_size = ""
thrifty_food = ""
total_case_income = ""
total_case_assets = ""
total_liquid_assets = ""
SNAP_active = False
DWP_active = False
MFIP_active = False
GA_active = False
MSA_active = False
subsidized_rent = False

FAMILY_CASE = FALSE
ADULT_CASE = TRUE

SNAP_JOBS_Income = 0
Adult_Cash_JOBS_Income = 0
Family_Cash_JOBS_Income = 0
total_case_unea = 0

monthly_childcare_exp = 0
monthly_adultcare_exp = 0
child_support_exp = 0
alimony_exp = 0
monthly_fmed_exp = 0

snap_hh_size = 0
family_cash_hh_size = 0
adult_cash_hh_size = 0

heat_AC_amt = 493
electric_amt = 126
phone_amt = 47

DISA_Case = False

call check_for_MAXIS(False)	'checking for an active MAXIS session

call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
If MAXIS_footer_month = "" Then MAXIS_footer_month = Right("00" & DatePart("m", date), 2)
If MAXIS_footer_year = "" Then MAXIS_footer_year = Right("00" & DatePart("yyyy", date), 2)

'Defining case number dialog'
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 256, 80, "Budget Estimator"
  EditBox 65, 5, 80, 15, MAXIS_case_number
  EditBox 205, 5, 20, 15, MAXIS_footer_month
  EditBox 230, 5, 20, 15, MAXIS_footer_year
  CheckBox 10, 25, 210, 10, "Check here if estimating on a new case (with no case number)", no_case_number_checkbox
  EditBox 75, 40, 175, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 145, 60, 50, 15
    CancelButton 200, 60, 50, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 155, 10, 45, 10, "Footer Month:"
  Text 5, 45, 65, 10, "Worker's Signature:"
EndDialog

Do
	err_msg = ""
	dialog Dialog1
	If ButtonPressed = 0 Then script_end_procedure("")
	If MAXIS_case_number = "" AND no_case_number_checkbox = unchecked Then err_msg = err_msg & vbnewLine & "- Enter a case number. To run on a situation with no case number, check the box for a new case."
	If MAXIS_footer_month = "" OR MAXIS_footer_year = "" Then err_msg = err_msg & vbnewLine & "- Enter footer month and year."
	If worker_signature = "" Then err_msg = err_msg & vbnewLine & "- Enter your name."
	If err_msg <> "" Then MsgBox "** Please resolve the following to continue:" & vbnewLine & err_msg
Loop until err_msg = ""

If no_case_number_checkbox = unchecked Then

	'Creating the dropdown of HH Members for use in dialogs
	Call Generate_Client_List(HH_Memb_DropDown, "Select One...")

	Call Navigate_to_MAXIS_screen ("CASE", "CURR")

	Dim search_fields_array (5)			'Creates an array of different progeam options to loop through
	search_fields_array(0) = "Case:"
	search_fields_array(1) = "MFIP:"
	search_fields_array(2) = "DWP:"
	search_fields_array(3) = "GA:"
	search_fields_array(4) = "MSA:"
	search_fields_array(5) = "FS:"
	For each program in search_fields_array		'this will now loop through each of the program options and set a boolean based on the information found.
		prog_status = ""						'clearin the variable
		row = 1
		col = 1
		search = program
		EMSearch search, row, col
		If row <> 0 Then 						'If the search finds that program type on case curr - it will read the status associated with it
			EMReadScreen prog_status, 9, row, 9
		End If
		prog_status = trim(prog_status)			'Now it set the case types based on the programs and status
		If program = "Case:" AND prog_status = "INACTIVE" Then Case_inactive = TRUE
		If prog_status = "ACTIVE" Then
			If program = "MFIP:" Then MFIP_active = TRUE
			If program = "FS:" 	 Then SNAP_Active = TRUE
			If program = "DWP:"  Then DWP_active = TRUE
			If program = "GA:"   Then GA_active = TRUE
			If program = "MSA:"  Then MSA_active = TRUE
		End If
		If prog_status = "APP CLOSE" Then
			If program = "MFIP:" Then
				MFIP_active = TRUE
			ElseIf program = "FS:" 	 Then
				SNAP_Active = TRUE
			ElseIf program = "DWP:"  Then
				DWP_active = TRUE
			ElseIf program = "GA:"   Then
				GA_active = TRUE
			ElseIf program = "MSA:"  Then
				MSA_active = TRUE
			End If
		End If
	Next

	CALL Navigate_to_MAXIS_screen("STAT", "REMO")
	clients_left = ""
	memb_row = 5
	Do
		EMReadScreen ref_numb, 2, memb_row, 3
		If ref_numb = "  " Then Exit Do
		EMWriteScreen ref_numb, 20, 76
		transmit
		EMReadScreen left_date, 8, 8, 53
		EMReadScreen expect_date, 8, 14, 53
		EMReadScreen return_date, 8, 16, 53

		left_date = replace(left_date, " ", "/")
		expect_date = replace(expect_date, " ", "/")
		return_date = replace(return_date, " ", "/")

		If left_date <> "__/__/__" Then
			If expect_date = "__/__/__" AND return_date = "__/__/__" then clients_left = clients_left & ", " & ref_numb
			If expect_date <> "__/__/__" Then
				If DateDiff("d", date, expect_date) > 0 Then clients_left = clients_left & ", " & ref_numb
			End if
		End If
		memb_row = memb_row + 1
	Loop until memb_row = 20

	If clients_left <> "" Then clients_left = right(clients_left, len(clients_left) - 2)

	CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.

	memb_row = 5
	people_counter = 0
	Call navigate_to_MAXIS_screen ("STAT", "MEMB")
	Do
		EMReadScreen ref_numb, 2, memb_row, 3
		If ref_numb = "  " Then Exit Do

		EMWriteScreen ref_numb, 20, 76
		transmit
		EMReadScreen first_name, 12, 6, 63
		EMReadScreen last_name, 25, 6, 30
		EMReadscreen client_age, 2, 8, 76
		If InStr(clients_left, ref_numb) = 0 Then
			ReDim Preserve CASE_INFO_ARRAY(18, people_counter)
			client_age = replace(client_age, " ", "")
			if client_age = "" Then client_age = 0
			client_age = client_age * 1
			If client_age >= 20 then
				client_is = "(ADULT)"
			Else
				client_is = "(CHILD)"
			End If

			If client_age >= 60 Then elderly_disabled_checkbox = checked

			If client_age <= 18 Then FAMILY_CASE = TRUE
			If client_age < 18 Then ADULT_CASE = FALSE

			CASE_INFO_ARRAY (clt_name, people_counter) = replace(first_name, "_", "") & " " & replace(last_name, "_", "")
			CASE_INFO_ARRAY (clt_ref, people_counter) = ref_numb
			CASE_INFO_ARRAY (clt_a_c, people_counter) = client_is
			CASE_INFO_ARRAY (clt_age, people_counter) = client_age
			CASE_INFO_ARRAY (include_snap, people_counter) = checked

			people_counter = people_counter + 1
		End if
		memb_row = memb_row + 1
	Loop until memb_row = 20

	'Finding MFIP TIME information
	Call Navigate_to_MAXIS_screen ("STAT", "TIME")
	case_months = 0

	For client = 0 to UBOUND(CASE_INFO_ARRAY, 2)
		EMWriteScreen CASE_INFO_ARRAY(clt_ref, client), 20, 76
		transmit
		EMReadScreen reg_months, 2, 17, 69
		EMReadScreen ext_months, 2, 19, 31
		reg_months = trim(reg_months)
		ext_months = trim(ext_months)
		If ext_months = "__" Then ext_months = 0

		reg_months = reg_months * 1
		ext_months = ext_months * 1

		client_months = reg_months + ext_months

		If client_months > case_months Then case_months = client_months
	Next

	'Gathering DCEX information
	Call Navigate_to_MAXIS_screen("STAT", "DCEX")
	monthly_childcare_exp = 0
	For client = 0 to UBOUND(CASE_INFO_ARRAY, 2)
		EMWriteScreen CASE_INFO_ARRAY(clt_ref, client), 20, 76
		transmit
		mx_row = 11
		Do
			retro_exp = ""
			EMReadScreen prosp_exp, 8, mx_row, 63
			prosp_exp = trim(prosp_exp)

			If prosp_exp = "________" then
				EMReadScreen retro_exp, 8, mx_row, 48
				retro_exp = trim(retro_exp)
				If retro_exp = "________" Then Exit Do
				prosp_exp = 0
				retro_exp = retro_exp * 1
				monthly_childcare_exp = monthly_childcare_exp + retro_exp
			Else
				monthly_childcare_exp = monthly_childcare_exp + prosp_exp
			End If
		Loop until mx_row = 17
	Next

	'Gathering COEX information
	Call Navigate_to_MAXIS_screen("STAT", "COEX")
	child_support_exp = 0
	alimony_exp = 0
	For client = 0 to UBOUND(CASE_INFO_ARRAY, 2)
		EMWriteScreen CASE_INFO_ARRAY(clt_ref, client), 20, 76
		transmit
		EMReadScreen client_support, 8, 10, 63
		EMReadScreen client_alimony, 8, 11, 63

		If client_support = "________" Then EMReadScreen client_support, 8, 10, 45
		If client_alimony = "________" Then EMReadScreen client_alimony, 8, 11, 45

		client_support = trim(client_support)
		client_alimony = trim(client_alimony)

		If client_support = "________" Then client_support = 0
		If client_alimony = "________" Then client_alimony = 0

		client_support = client_support * 1
		client_alimony = client_alimony * 1

		child_support_exp = child_support_exp + client_support
		alimony_exp = alimony_exp + client_alimony
	Next

	'Getting all EARNED Income information
	'JOBS
	array_counter = 0
	For client = 0 to UBOUND(CASE_INFO_ARRAY, 2)
		Call Navigate_to_MAXIS_screen("STAT", "JOBS")
		EMWriteScreen CASE_INFO_ARRAY(clt_ref, client), 20, 76
		transmit
		Do
			pic_check_exists = False
			retro_check_exists = false
			EMReadScreen employer_name, 30, 7, 42
			employer_name = replace(employer_name, "_", "")
			If employer_name = "" Then Exit Do
			EMReadScreen end_date, 8, 9, 49
			end_date = replace(end_date, " ", "/")
			If end_date <> "__/__/__" Then
				If DateDiff("d", end_date, date) > 1 Then Exit Do
			End If
			ReDim Preserve EI_ARRAY(18, array_counter)
			EMReadScreen retro_total, 8, 17, 38

			EMReadScreen retro_check_1_date, 8, 12, 25
			EMReadScreen retro_check_1_amt,  8, 12, 38
			EMReadScreen retro_check_2_date, 8, 13, 25
			EMReadScreen retro_check_2_amt,  8, 13, 38
			EMReadScreen retro_check_3_date, 8, 14, 25
			EMReadScreen retro_check_3_amt,  8, 14, 38
			EMReadScreen retro_check_4_date, 8, 15, 25
			EMReadScreen retro_check_4_amt,  8, 15, 38
			EMReadScreen retro_check_5_date, 8, 16, 25
			EMReadScreen retro_check_5_amt,  8, 16, 38

			retro_total = trim(retro_total)
			If retro_total = "" Then retro_total = 0
			retro_total = retro_total * 1

			retro_check_1_date = replace(retro_check_1_date, " ", "/")
			retro_check_1_amt = trim(retro_check_1_amt)
			if retro_check_1_amt = "________" Then retro_check_1_amt = 0
			retro_check_1_amt = retro_check_1_amt * 1
			if retro_check_1_date <> "__/__/__" Then retro_check_exists = true

			retro_check_2_date = replace(retro_check_2_date, " ", "/")
			retro_check_2_amt = trim(retro_check_2_amt)
			if retro_check_2_amt = "________" Then retro_check_2_amt = 0
			retro_check_2_amt = retro_check_2_amt * 1

			retro_check_3_date = replace(retro_check_3_date, " ", "/")
			retro_check_3_amt = trim(retro_check_3_amt)
			if retro_check_3_amt = "________" Then retro_check_3_amt = 0
			retro_check_3_amt = retro_check_3_amt * 1

			retro_check_4_date = replace(retro_check_4_date, " ", "/")
			retro_check_4_amt = trim(retro_check_4_amt)
			if retro_check_4_amt = "________" Then retro_check_4_amt = 0
			retro_check_4_amt = retro_check_4_amt * 1

			retro_check_5_date = replace(retro_check_5_date, " ", "/")
			retro_check_5_amt = trim(retro_check_5_amt)
			if retro_check_5_amt = "________" Then retro_check_5_amt = 0
			retro_check_5_amt = retro_check_5_amt * 1

			EMReadScreen prosp_total, 8, 17, 67
			prosp_total = trim(prosp_total)
			If prosp_total = "" Then prosp_total = 0
			prosp_total = prosp_total * 1

			EMReadScreen pay_freq, 1, 18, 35

			If SNAP_Active = True Then
				EMWriteScreen "X", 19, 38
				transmit

				EMReadScreen pic_total, 8, 18, 56
				pic_total = trim(pic_total)

				EMReadScreen antic_hour, 6, 8, 64
				EMReadScreen antic_rate, 8, 9, 66
				antic_hour = trim(antic_hour)
				antic_rate = trim(antic_rate)
				If antic_hour = "______" Then antic_hour = 0
				If antic_rate = "________" Then antic_rate = 0

				If antic_hour <> "" AND pay_freq <> "" Then
					EMReadScreen pic_check_1_date, 8, 9,  13
					EMReadScreen pic_check_1_amt,  8, 9,  25
					EMReadScreen pic_check_2_date, 8, 10, 13
					EMReadScreen pic_check_2_amt,  8, 10, 25
					EMReadScreen pic_check_3_date, 8, 11, 13
					EMReadScreen pic_check_3_amt,  8, 11, 25
					EMReadScreen pic_check_4_date, 8, 12, 13
					EMReadScreen pic_check_4_amt,  8, 12, 25
					EMReadScreen pic_check_5_date, 8, 13, 13
					EMReadScreen pic_check_5_amt,  8, 13, 25

					pic_check_1_date = replace(pic_check_1_date, " ", "/")
					pic_check_1_amt = trim(pic_check_1_amt)
					if pic_check_1_amt = "________" Then pic_check_1_amt = 0
					pic_check_1_amt = pic_check_1_amt * 1
					If pic_check_1_date <> "__/__/__" Then pic_check_exists = True

					pic_check_2_date = replace(pic_check_2_date, " ", "/")
					pic_check_2_amt = trim(pic_check_2_amt)
					if pic_check_2_amt = "________" Then pic_check_2_amt = 0
					pic_check_2_amt = pic_check_2_amt * 1

					pic_check_3_date = replace(pic_check_3_date, " ", "/")
					pic_check_3_amt = trim(pic_check_3_amt)
					if pic_check_3_amt = "________" Then pic_check_3_amt = 0
					pic_check_3_amt = pic_check_3_amt * 1

					pic_check_4_date = replace(pic_check_4_date, " ", "/")
					pic_check_4_amt = trim(pic_check_4_amt)
					if pic_check_4_amt = "________" Then pic_check_4_amt = 0
					pic_check_4_amt = pic_check_4_amt * 1

					pic_check_5_date = replace(pic_check_5_date, " ", "/")
					pic_check_5_amt = trim(pic_check_5_amt)
					if pic_check_5_amt = "________" Then pic_check_5_amt = 0
					pic_check_5_amt = pic_check_5_amt * 1

				End If

				PF3
			End If

			EI_ARRAY(employee, array_counter) = CASE_INFO_ARRAY(clt_ref, client) & " - " & CASE_INFO_ARRAY(clt_name, client)
			EI_ARRAY(employer, array_counter) = employer_name
			EI_ARRAY(job_retro_gross, array_counter) = retro_total
			EI_ARRAY(job_prosp_gross, array_counter) = prosp_total
			EI_ARRAY(job_pic_gross, array_counter) = pic_total
			If pay_freq = "1" then EI_ARRAY(job_pay_freq, array_counter) = "Once/Month - 1"
			If pay_freq = "2" then EI_ARRAY(job_pay_freq, array_counter) = "Twice/Month - 2"
			If pay_freq = "3" then EI_ARRAY(job_pay_freq, array_counter) = "Biweekly - 3"
			If pay_freq = "4" then EI_ARRAY(job_pay_freq, array_counter) = "Weekly - 4"

			If retro_check_exists = true AND pic_check_exists = false Then
				EI_ARRAY(check_1_date, array_counter) = retro_check_1_date
				EI_ARRAY(check_1_gross, array_counter) = retro_check_1_amt
				EI_ARRAY(how_many_chck, array_counter) = 1
				If retro_check_2_date <> "__/__/__" Then
					EI_ARRAY(check_2_date, array_counter) = retro_check_2_date
					EI_ARRAY(check_2_gross, array_counter) = retro_check_2_amt
					EI_ARRAY(how_many_chck, array_counter) = 2
				End If
				If retro_check_3_date <> "__/__/__" Then
					EI_ARRAY(check_3_date, array_counter) = retro_check_3_date
					EI_ARRAY(check_3_gross, array_counter) = retro_check_3_amt
					EI_ARRAY(how_many_chck, array_counter) = 3
				End If
				If retro_check_4_date <> "__/__/__" Then
					EI_ARRAY(check_4_date, array_counter) = retro_check_4_date
					EI_ARRAY(check_4_gross, array_counter) = retro_check_4_amt
					EI_ARRAY(how_many_chck, array_counter) = 4
				End If
				If retro_check_5_date <> "__/__/__" Then
					EI_ARRAY(check_5_date, array_counter) = retro_check_5_date
					EI_ARRAY(check_5_gross, array_counter) = retro_check_5_amt
					EI_ARRAY(how_many_chck, array_counter) = 5
				End If
			End If

			If retro_check_exists = false AND pic_check_exists = true Then
				EI_ARRAY(check_1_date, array_counter) = pic_check_1_date
				EI_ARRAY(check_1_gross, array_counter) = pic_check_1_amt
				EI_ARRAY(how_many_chck, array_counter) = 1
				If pic_check_2_date <> "__/__/__" Then
					EI_ARRAY(check_2_date, array_counter) = pic_check_2_date
					EI_ARRAY(check_2_gross, array_counter) = pic_check_2_amt
					EI_ARRAY(how_many_chck, array_counter) = 2
				End If
				If pic_check_3_date <> "__/__/__" Then
					EI_ARRAY(check_3_date, array_counter) = pic_check_3_date
					EI_ARRAY(check_3_gross, array_counter) = pic_check_3_amt
					EI_ARRAY(how_many_chck, array_counter) = 3
				End If
				If pic_check_4_date <> "__/__/__" Then
					EI_ARRAY(check_4_date, array_counter) = pic_check_4_date
					EI_ARRAY(check_4_gross, array_counter) = pic_check_4_amt
					EI_ARRAY(how_many_chck, array_counter) = 4
				End If
				If pic_check_5_date <> "__/__/__" Then
					EI_ARRAY(check_5_date, array_counter) = pic_check_5_date
					EI_ARRAY(check_5_gross, array_counter) = pic_check_5_amt
					EI_ARRAY(how_many_chck, array_counter) = 5
				End If
			End If

			If retro_check_exists = true AND pic_check_exists = true Then
				If GA_active = TRUE OR MFIP_active = TRUE Then
					EI_ARRAY(check_1_date, array_counter) = retro_check_1_date
					EI_ARRAY(check_1_gross, array_counter) = retro_check_1_amt
					EI_ARRAY(how_many_chck, array_counter) = 1
					If retro_check_2_date <> "__/__/__" Then
						EI_ARRAY(check_2_date, array_counter) = retro_check_2_date
						EI_ARRAY(check_2_gross, array_counter) = retro_check_2_amt
						EI_ARRAY(how_many_chck, array_counter) = 2
					End If
					If retro_check_3_date <> "__/__/__" Then
						EI_ARRAY(check_3_date, array_counter) = retro_check_3_date
						EI_ARRAY(check_3_gross, array_counter) = retro_check_3_amt
						EI_ARRAY(how_many_chck, array_counter) = 3
					End If
					If retro_check_4_date <> "__/__/__" Then
						EI_ARRAY(check_4_date, array_counter) = retro_check_4_date
						EI_ARRAY(check_4_gross, array_counter) = retro_check_4_amt
						EI_ARRAY(how_many_chck, array_counter) = 4
					End If
					If retro_check_5_date <> "__/__/__" Then
						EI_ARRAY(check_5_date, array_counter) = retro_check_5_date
						EI_ARRAY(check_5_gross, array_counter) = retro_check_5_amt
						EI_ARRAY(how_many_chck, array_counter) = 5
					End If
				Else
					EI_ARRAY(check_1_date, array_counter) = pic_check_1_date
					EI_ARRAY(check_1_gross, array_counter) = pic_check_1_amt
					EI_ARRAY(how_many_chck, array_counter) = 1
					If pic_check_2_date <> "__/__/__" Then
						EI_ARRAY(check_2_date, array_counter) = pic_check_2_date
						EI_ARRAY(check_2_gross, array_counter) = pic_check_2_amt
						EI_ARRAY(how_many_chck, array_counter) = 2
					End If
					If pic_check_3_date <> "__/__/__" Then
						EI_ARRAY(check_3_date, array_counter) = pic_check_3_date
						EI_ARRAY(check_3_gross, array_counter) = pic_check_3_amt
						EI_ARRAY(how_many_chck, array_counter) = 3
					End If
					If pic_check_4_date <> "__/__/__" Then
						EI_ARRAY(check_4_date, array_counter) = pic_check_4_date
						EI_ARRAY(check_4_gross, array_counter) = pic_check_4_amt
						EI_ARRAY(how_many_chck, array_counter) = 4
					End If
					If pic_check_5_date <> "__/__/__" Then
						EI_ARRAY(check_5_date, array_counter) = pic_check_5_date
						EI_ARRAY(check_5_gross, array_counter) = pic_check_5_amt
						EI_ARRAY(how_many_chck, array_counter) = 5
					End If
				End If
			End If

			EI_ARRAY(pic_hrs_wk, array_counter) = antic_hour
			EI_ARRAY(pic_rate_pay, array_counter) = antic_rate

			array_counter = array_counter + 1
			transmit
			EMReadScreen nav_msg, 5, 24, 2
		Loop Until nav_msg = "ENTER"
	Next

	'Getting all Self Employment Information
	For client = 0 to UBOUND(CASE_INFO_ARRAY, 2)
		Call Navigate_to_MAXIS_screen("STAT", "BUSI")
		EMWriteScreen CASE_INFO_ARRAY(clt_ref, client), 20, 76
		transmit
		Do
			EMReadScreen busi_type, 2, 5, 37
			busi_type = trim(busi_type)
			busi_type = replace(busi_type, "_", "")
			If busi_type = "" Then Exit Do
			ReDim Preserve EI_ARRAY(18, array_counter)

			EMReadScreen cash_retro, 8, 8, 55
			EMReadScreen cash_prosp, 8, 8, 69
			EMReadScreen snap_retro, 8, 10, 55
			EMReadScreen snap_prosp, 8, 10, 69

			cash_retro = trim(cash_retro)
			cash_prosp = trim(cash_prosp)
			snap_retro = trim(snap_retro)
			snap_prosp = trim(snap_prosp)

			EI_ARRAY(employee, array_counter) = CASE_INFO_ARRAY(clt_ref, client) & " - " & CASE_INFO_ARRAY(clt_name, client)
			EI_ARRAY(employer, array_counter) = "SELF EMPLOYMENT"
			EI_ARRAY(job_retro_gross, array_counter) = cash_retro
			EI_ARRAY(job_prosp_gross, array_counter) = cash_prosp
			EI_ARRAY(job_pic_gross, array_counter) = snap_prosp
			EI_ARRAY(job_pay_freq, array_counter) = "1"

			array_counter = array_counter + 1
			transmit
			EMReadScreen nav_msg, 5, 24, 2
		Loop Until nav_msg = "ENTER"
	Next

	'Getting all UNEA information
	For all_clts = 0 to UBOUND(CASE_INFO_ARRAY, 2)
		Call Navigate_to_MAXIS_screen("STAT", "UNEA")
		EMWriteScreen CASE_INFO_ARRAY(clt_ref, all_clts), 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		Do
			EMReadScreen unea_type, 2, 5, 37
			Select Case unea_type

				Case "03"
					EMReadScreen ssi_total, 8, 18, 68
					ssi_total = trim(ssi_total)
					If ssi_total = "" then ssi_total = 0
					ssi_total = ssi_total * 1
					CASE_INFO_ARRAY(clt_ssi_income, all_clts) = ssi_total
				Case "01", "02"
					EMReadScreen rsdi_total, 8, 18, 68
					rsdi_total = trim(rsdi_total)
					If rsdi_total = "" then rsdi_total = 0
					rsdi_total = rsdi_total * 1
					CASE_INFO_ARRAY(clt_rsdi_income, all_clts) = CASE_INFO_ARRAY(clt_rsdi_income, all_clts) + rsdi_total
				Case "06"
					EMReadScreen unea_total, 8, 18, 68
					If unea_total = "" Then unea_total = 0
					unea_total = unea_total * 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "Non-MN PA"
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "Non-MN PA"
					Else
						MsgBox CASE_INFO_ARRAY(clt_name, all_clts) & " has more UNEA panel types than this script can accomodate right now, the information autofilled may be incorrect."
					End If
				Case "11", "12", "13", "38"
					EMReadScreen unea_total, 8, 18, 68
					If unea_total = "" Then unea_total = 0
					unea_total = unea_total * 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "VA Income"
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "VA Income"
					Else
						MsgBox CASE_INFO_ARRAY(clt_name, all_clts) & " has more UNEA panel types than this script can accomodate right now, the information autofilled may be incorrect."
					End If
				Case "14"
					EMReadScreen unea_total, 8, 18, 68
					If unea_total = "" Then unea_total = 0
					unea_total = unea_total * 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "Unemployment Insurance"
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "Unemployment Insurance"
					Else
						MsgBox CASE_INFO_ARRAY(clt_name, all_clts) & " has more UNEA panel types than this script can accomodate right now, the information autofilled may be incorrect."
					End If
				Case "15"
					EMReadScreen unea_total, 8, 18, 68
					If unea_total = "" Then unea_total = 0
					unea_total = unea_total * 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "Worker's Comp"
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "Worker's Comp"
					Else
						MsgBox CASE_INFO_ARRAY(clt_name, all_clts) & " has more UNEA panel types than this script can accomodate right now, the information autofilled may be incorrect."
					End If
				Case "16"
					EMReadScreen unea_total, 8, 18, 68
					If unea_total = "" Then unea_total = 0
					unea_total = unea_total * 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "Railroad Retirement"
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "Railroad Retirement"
					Else
						MsgBox CASE_INFO_ARRAY(clt_name, all_clts) & " has more UNEA panel types than this script can accomodate right now, the information autofilled may be incorrect."
					End If
				Case "17"
					EMReadScreen unea_total, 8, 18, 68
					If unea_total = "" Then unea_total = 0
					unea_total = unea_total * 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "Other Retirement"
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "Other Retirement"
					Else
						MsgBox CASE_INFO_ARRAY(clt_name, all_clts) & " has more UNEA panel types than this script can accomodate right now, the information autofilled may be incorrect."
					End If
				Case "18"
					EMReadScreen unea_total, 8, 18, 68
					If unea_total = "" Then unea_total = 0
					unea_total = unea_total * 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "Military Allotment"
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "Military Allotment"
					Else
						MsgBox CASE_INFO_ARRAY(clt_name, all_clts) & " has more UNEA panel types than this script can accomodate right now, the information autofilled may be incorrect."
					End If
				Case "19"
					EMReadScreen unea_total, 8, 18, 68
					If unea_total = "" Then unea_total = 0
					unea_total = unea_total * 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "FC Child Requesting FS"
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "FC Child Requesting FS"
					Else
						MsgBox CASE_INFO_ARRAY(clt_name, all_clts) & " has more UNEA panel types than this script can accomodate right now, the information autofilled may be incorrect."
					End If
				Case "20"
					EMReadScreen unea_total, 8, 18, 68
					If unea_total = "" Then unea_total = 0
					unea_total = unea_total * 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "FC Child Not Req FS"
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "FC Child Not Req FS"
					Else
						MsgBox CASE_INFO_ARRAY(clt_name, all_clts) & " has more UNEA panel types than this script can accomodate right now, the information autofilled may be incorrect."
					End If
				Case "21"
					EMReadScreen unea_total, 8, 18, 68
					If unea_total = "" Then unea_total = 0
					unea_total = unea_total * 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "FC Adult Requesting FS"
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "FC Adult Requesting FS"
					Else
						MsgBox CASE_INFO_ARRAY(clt_name, all_clts) & " has more UNEA panel types than this script can accomodate right now, the information autofilled may be incorrect."
					End If
				Case "22"
					EMReadScreen unea_total, 8, 18, 68
					If unea_total = "" Then unea_total = 0
					unea_total = unea_total * 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "FC Adult Not Req FS"
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "FC Adult Not Req FS"
					Else
						MsgBox CASE_INFO_ARRAY(clt_name, all_clts) & " has more UNEA panel types than this script can accomodate right now, the information autofilled may be incorrect."
					End If
				Case "23"
					EMReadScreen unea_total, 8, 18, 68
					If unea_total = "" Then unea_total = 0
					unea_total = unea_total * 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "Dividends"
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "Dividends"
					Else
						MsgBox CASE_INFO_ARRAY(clt_name, all_clts) & " has more UNEA panel types than this script can accomodate right now, the information autofilled may be incorrect."
					End If
				Case "24"
					EMReadScreen unea_total, 8, 18, 68
					If unea_total = "" Then unea_total = 0
					unea_total = unea_total * 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "Interest"
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "Interest"
					Else
						MsgBox CASE_INFO_ARRAY(clt_name, all_clts) & " has more UNEA panel types than this script can accomodate right now, the information autofilled may be incorrect."
					End If
				Case "25"
					EMReadScreen unea_total, 8, 18, 68
					If unea_total = "" Then unea_total = 0
					unea_total = unea_total * 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "Cnt Gifts Or Prizes"
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "Cnt Gifts Or Prizes"
					Else
						MsgBox CASE_INFO_ARRAY(clt_name, all_clts) & " has more UNEA panel types than this script can accomodate right now, the information autofilled may be incorrect."
					End If
				Case "26"
					EMReadScreen unea_total, 8, 18, 68
					If unea_total = "" Then unea_total = 0
					unea_total = unea_total * 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "Strike Benefit"
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "Strike Benefit"
					Else
						MsgBox CASE_INFO_ARRAY(clt_name, all_clts) & " has more UNEA panel types than this script can accomodate right now, the information autofilled may be incorrect."
					End If
				Case "27"
					EMReadScreen unea_total, 8, 18, 68
					If unea_total = "" Then unea_total = 0
					unea_total = unea_total * 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "Contract For Deed"
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "Contract For Deed"
					Else
						MsgBox CASE_INFO_ARRAY(clt_name, all_clts) & " has more UNEA panel types than this script can accomodate right now, the information autofilled may be incorrect."
					End If
				Case "28"
					EMReadScreen unea_total, 8, 18, 68
					If unea_total = "" Then unea_total = 0
					unea_total = unea_total * 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "Illegal Income"
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "Illegal Income"
					Else
						MsgBox CASE_INFO_ARRAY(clt_name, all_clts) & " has more UNEA panel types than this script can accomodate right now, the information autofilled may be incorrect."
					End If
				Case "30"
					EMReadScreen unea_total, 8, 18, 68
					If unea_total = "" Then unea_total = 0
					unea_total = unea_total * 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "Infrequent <30 Not Counted"
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "Infrequent <30 Not Counted"
					Else
						MsgBox CASE_INFO_ARRAY(clt_name, all_clts) & " has more UNEA panel types than this script can accomodate right now, the information autofilled may be incorrect."
					End If
				Case "31"
					EMReadScreen unea_total, 8, 18, 68
					If unea_total = "" Then unea_total = 0
					unea_total = unea_total * 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "Other FS Only"
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "Other FS Only"
					Else
						MsgBox CASE_INFO_ARRAY(clt_name, all_clts) & " has more UNEA panel types than this script can accomodate right now, the information autofilled may be incorrect."
					End If
				Case "35", "37", "40"
					EMReadScreen unea_total, 8, 18, 68
					If unea_total = "" Then unea_total = 0
					unea_total = unea_total * 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "Spousal Support"
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "Spousal Support"
					Else
						MsgBox CASE_INFO_ARRAY(clt_name, all_clts) & " has more UNEA panel types than this script can accomodate right now, the information autofilled may be incorrect."
					End If
				Case "46"
					EMReadScreen unea_total, 8, 18, 68
					If unea_total = "" Then unea_total = 0
					unea_total = unea_total * 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "County 88 Gaming"
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "County 88 Gaming"
					Else
						MsgBox CASE_INFO_ARRAY(clt_name, all_clts) & " has more UNEA panel types than this script can accomodate right now, the information autofilled may be incorrect."
					End If
				Case "08", "36", "39"
					EMReadScreen unea_total, 8, 18, 68
					unea_total = trim(unea_total)
					If unea_total = "" Then unea_total = 0
					unea_total = unea_total * 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "Child Support" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) + unea_total
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "Child Support" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) + unea_total
					ElseIf CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "Child Support"
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "Child Support"
					Else
						MsgBox CASE_INFO_ARRAY(clt_name, all_clts) & " has more UNEA panel types than this script can accomodate right now, the information autofilled may be incorrect."
					End If
				Case Else
					EMReadScreen unea_total, 8, 18, 68
					unea_total = trim(unea_total)
					If unea_total = "" Then unea_total = 0
					unea_total = unea_total * 1
					If CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "Other" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) + unea_total
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "Other" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) + unea_total
					ElseIf CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_1_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_1_type, all_clts) = "Other"
					ElseIf CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "" Then
						CASE_INFO_ARRAY(clt_other_unea_2_amt, all_clts) = unea_total
						CASE_INFO_ARRAY(clt_other_unea_2_type, all_clts) = "Other"
					Else
						MsgBox CASE_INFO_ARRAY(clt_name, all_clts) & " has more UNEA panel types than this script can accomodate right now, the information autofilled may be incorrect."
					End If
			End Select

			transmit
			EMReadScreen nav_msg, 5, 24, 2
		Loop Until nav_msg = "ENTER"
	Next

	'Getting all Account information
	For all_clts = 0 to UBOUND(CASE_INFO_ARRAY, 2)
		Call Navigate_to_MAXIS_screen("STAT", "ACCT")
		EMWriteScreen CASE_INFO_ARRAY(clt_ref, all_clts), 20, 76
		transmit

		Do
			EMReadScreen acct_type, 2, 6, 44
			If acct_type = "__" Then Exit Do

			EMReadScreen acct_balance, 8, 10, 46
			EMReadScreen withdraw_penlty, 8, 12, 46
			acct_balance = trim(acct_balance)
			withdraw_penlty = trim(withdraw_penlty)
			If acct_balance = "________" then acct_balance = 0
			If withdraw_penlty = "________" then withdraw_penlty = 0
			acct_balance = acct_balance - withdraw_penlty

			If acct_type = "CK" Then
				CASE_INFO_ARRAY(clt_chk_acct, all_clts) = CASE_INFO_ARRAY(clt_chk_acct, all_clts) + acct_balance
			ElseIf acct_type = "SV" Then
				CASE_INFO_ARRAY(clt_sav_acct, all_clts) = CASE_INFO_ARRAY(clt_sav_acct, all_clts) + acct_balance
			ElseIf acct_type = "DC" Then
				If CASE_INFO_ARRAY(clt_asset_other_type, all_clts) = "" OR CASE_INFO_ARRAY(clt_asset_other_type, all_clts) = "Debit Card" Then
					CASE_INFO_ARRAY(clt_asset_other_bal, all_clts) = acct_balance
					CASE_INFO_ARRAY(clt_asset_other_type, all_clts) = "Debit Card"
				Else
					CASE_INFO_ARRAY(clt_asset_other_bal, all_clts) = CASE_INFO_ARRAY(clt_asset_other_bal, all_clts) + acct_balance
					CASE_INFO_ARRAY(clt_asset_other_type, all_clts) = "MULTIPLE"
				ENd If
			Else
				If CASE_INFO_ARRAY(clt_asset_other_type, all_clts) = "" Then
					CASE_INFO_ARRAY(clt_asset_other_bal, all_clts) = acct_balance
					CASE_INFO_ARRAY(clt_asset_other_type, all_clts) = "Other"
				Else
					CASE_INFO_ARRAY(clt_asset_other_bal, all_clts) = CASE_INFO_ARRAY(clt_asset_other_bal, all_clts) + acct_balance
					CASE_INFO_ARRAY(clt_asset_other_type, all_clts) = "MULTIPLE"
				ENd If
			End If

			transmit
			EMReadScreen nav_msg, 5, 24, 2
		Loop Until nav_msg = "ENTER"
	Next

	subsidized_rent = FALSE
	'Getting shelter and other expense information
	For all_clts = 0 to UBOUND(CASE_INFO_ARRAY, 2)
		Call Navigate_to_MAXIS_screen("STAT", "SHEL")
		EMWriteScreen CASE_INFO_ARRAY(clt_ref, all_clts), 20, 76
		transmit

		EMReadScreen subsidy_check,    1,  6, 46
		EMReadScreen entered_rent,     8, 11, 56
		EMReadScreen entered_lot_rent, 8, 12, 56
		EMReadScreen entered_mortgage, 8, 13, 56
		EMReadScreen entered_ins,      8, 14, 56
		EMReadScreen entered_taxes,    8, 15, 56
		EMReadScreen entered_room,     8, 16, 56

		entered_rent     = trim(entered_rent)
		entered_lot_rent = trim(entered_lot_rent)
		entered_mortgage = trim(entered_mortgage)
		entered_ins      = trim(entered_ins)
		entered_taxes    = trim(entered_taxes)
		entered_room     = trim(entered_room)

		If entered_rent =     "________" Then entered_rent     = 0
		If entered_lot_rent = "________" Then entered_lot_rent = 0
		If entered_mortgage = "________" Then entered_mortgage = 0
		If entered_ins =      "________" Then entered_ins      = 0
		If entered_taxes =    "________" Then entered_taxes    = 0
		If entered_room =     "________" Then entered_room     = 0

		entered_rent     = entered_rent * 1
		entered_lot_rent = entered_lot_rent * 1
		entered_mortgage = entered_mortgage* 1
		entered_ins      = entered_ins * 1
		entered_taxes    = entered_taxes * 1
		entered_room     = entered_room * 1

		If subsidy_check = "Y" Then subsidized_rent = TRUE
		rent_expense = rent_expense + entered_rent + entered_mortgage
		prop_tax_expense = prop_tax_expense + entered_taxes
		home_ins_expense = home_ins_expense + entered_ins
		other_expense = other_expense + entered_lot_rent + entered_room

		Call Navigate_to_MAXIS_screen("STAT", "HEST")
		EMWriteScreen CASE_INFO_ARRAY(clt_ref, all_clts), 20, 76
		transmit

		EMReadScreen heat_ac_yn, 1, 13, 60
		EMReadScreen elec_yn,    1, 14, 60
		EMReadScreen phone_yn,   1, 15, 60

		If heat_ac_yn = "Y" then heat_ac_checkbox  = checked
		If elec_yn = "Y"    then electric_checkbox = checked
		If phone_yn = "Y"   then phone_checkbox    = checked

		Call Navigate_to_MAXIS_screen("STAT", "ACUT")
		EMWriteScreen CASE_INFO_ARRAY(clt_ref, all_clts), 20, 76
		transmit

		EMReadScreen entered_heat,  8, 10, 61
		EMReadScreen entered_air,   8, 11, 61
		EMReadScreen entered_elec,  8, 12, 61
		EMReadScreen entered_fuel,  8, 13, 61
		EMReadScreen entered_garb,  8, 14, 61
		EMReadScreen entered_water, 8, 15, 61
		EMReadScreen entered_sewer, 8, 16, 61
		EMReadScreen entered_othr,  8, 17, 61

		entered_heat  = trim(entered_heat)
		entered_air   = trim(entered_air)
		entered_elec  = trim(entered_elec)
		entered_fuel  = trim(entered_fuel)
		entered_garb  = trim(entered_garb)
		entered_water = trim(entered_water)
		entered_sewer = trim(entered_sewer)
		entered_othr  = trim(entered_othr)

		If entered_heat =  "________" Then entered_heat  = 0
		If entered_air =   "________" Then entered_air   = 0
		If entered_elec =  "________" Then entered_elec  = 0
		If entered_fuel =  "________" Then entered_fuel  = 0
		If entered_garb =  "________" Then entered_garb  = 0
		If entered_water = "________" Then entered_water = 0
		If entered_sewer = "________" Then entered_sewer = 0
		If entered_othr =  "________" Then entered_othr  = 0

		entered_heat  = entered_heat * 1
		entered_air   = entered_air * 1
		entered_elec  = entered_elec* 1
		entered_fuel  = entered_fuel * 1
		entered_garb  = entered_garb * 1
		entered_water = entered_water * 1
		entered_sewer = entered_sewer * 1
		entered_othr  = entered_othr * 1

		actual_utility_expense = actual_utility_expense + entered_heat + entered_air + entered_elec + entered_fuel + entered_garb + entered_water + entered_sewer + entered_othr

	Next

	If SNAP_Active = TRUE Then
		Call Navigate_to_MAXIS_screen ("ELIG", "FS")
		EMWriteScreen "99", 19, 78
		transmit
		mx_row = 7
		Do
			EMReadScreen approval_status, 10, mx_row, 50
			approval_status = trim(approval_status)
			If approval_status = "APPROVED" Then
				EMReadScreen version_number, 2, mx_row, 22
				version_number = trim(version_number)
				EMWriteScreen version_number, 18, 54
				transmit
				Exit Do
			End If
			mx_row = mx_row + 1
		Loop until approval_status = ""

		fs_participants = 0
		For all_clients = 0 to UBOUND(CASE_INFO_ARRAY, 2)
			MX_row = 7
			Do
				EMReadScreen elig_ref, 2, mx_row, 10
				If elig_ref = CASE_INFO_ARRAY(clt_ref, all_clients) Then
					EMReadScreen elig_status, 10, mx_row, 57
					elig_status = trim(elig_status)
					If elig_status = "ELIGIBLE" Then
						CASE_INFO_ARRAY (include_snap, all_clients) = checked
						fs_participants = fs_participants + 1
					Else
						CASE_INFO_ARRAY (include_snap, all_clients) = unchecked
					End If
					Exit Do
				Else
					mx_row = mx_row + 1
				End If
			Loop until elig_ref = "  "
		Next

		EMWriteScreen "FSSM", 19, 70
		transmit

		EMReadScreen fs_grant, 8, 13, 73

		fs_grant = trim(fs_grant)

		snap_current_benefit = "$"  & fs_grant & ". HH Size: " & fs_participants

	End If

	If MFIP_active = TRUE Then
		Call Navigate_to_MAXIS_screen ("ELIG", "MFIP")
		EMWriteScreen "99", 20, 79
		transmit
		mx_row = 7
		Do
			EMReadScreen approval_status, 10, mx_row, 50
			approval_status = trim(approval_status)
			If approval_status = "APPROVED" Then
				EMReadScreen version_number, 2, mx_row, 22
				version_number = trim(version_number)
				EMWriteScreen version_number, 18, 54
				transmit
				Exit Do
			End If
			mx_row = mx_row + 1
		Loop until approval_status = ""

		mf_participants = 0
		For all_clients = 0 to UBOUND(CASE_INFO_ARRAY, 2)
			MX_row = 7
			Do
				EMReadScreen elig_ref, 2, mx_row, 6
				If elig_ref = CASE_INFO_ARRAY(clt_ref, all_clients) Then
					EMReadScreen elig_status, 10, mx_row, 53
					elig_status = trim(elig_status)
					If elig_status = "ELIGIBLE" Then
						CASE_INFO_ARRAY (include_family_cash, all_clients) = checked
						mf_participants = mf_participants + 1
					Else
						CASE_INFO_ARRAY (include_family_cash, all_clients) = unchecked
					End If
					Exit Do
				Else
					mx_row = mx_row + 1
				End If
			Loop until elig_ref = "  "
		Next

		EMWriteScreen "MFSM", 20, 71
		transmit

		EMReadScreen MF_all, 8, 13, 73
		EMReadScreen MF_MF,  8, 14, 73
		EMReadScreen MF_FS,  8, 15, 73
		EMReadScreen MF_HG,  8, 16, 73

		MF_all = trim(MF_all)
		MF_MF  = trim(MF_MF)
		MF_FS  = trim(MF_FS)
		MF_HG  = trim(MF_HG)

		family_cash_current_benefit = family_cash_current_benefit &  "MF-MF - $" & MF_MF & ", MF-FS - $" & MF_FS & ", MF-HG - $" & MF_HG & ". HH Size: " & mf_participants

	End If

	If DWP_active = TRUE Then
		Call Navigate_to_MAXIS_screen ("ELIG", "DWP")
		EMWriteScreen "99", 20, 79
		transmit
		mx_row = 7
		Do
			EMReadScreen approval_status, 10, mx_row, 50
			approval_status = trim(approval_status)
			If approval_status = "APPROVED" Then
				EMReadScreen version_number, 2, mx_row, 22
				version_number = trim(version_number)
				EMWriteScreen version_number, 18, 54
				transmit
				Exit Do
			End If
			mx_row = mx_row + 1
		Loop until approval_status = ""

		dwp_participants = 0
		For all_clients = 0 to UBOUND(CASE_INFO_ARRAY, 2)
			MX_row = 7
			Do
				EMReadScreen elig_ref, 2, mx_row, 5
				If elig_ref = CASE_INFO_ARRAY(clt_ref, all_clients) Then
					EMReadScreen elig_status, 10, mx_row, 57
					elig_status = trim(elig_status)
					If elig_status = "ELIGIBLE" Then
						CASE_INFO_ARRAY (include_family_cash, all_clients) = checked
						dwp_participants = dwp_participants + 1
					Else
						CASE_INFO_ARRAY (include_family_cash, all_clients) = unchecked
					End If
					Exit Do
				Else
					mx_row = mx_row + 1
				End If
			Loop until elig_ref = "  "
		Next

		EMWriteScreen "DWSM", 20, 71
		transmit

		EMReadScreen dwp_all, 8, 12, 73
		EMReadScreen dwp_shel,  8, 13, 73
		EMReadScreen dwp_pers,  8, 14, 73

		dwp_all = trim(dwp_all)
		dwp_shel  = trim(dwp_shel)
		dwp_pers  = trim(dwp_pers)

		family_cash_current_benefit = family_cash_current_benefit & "DWP - Total $" & dwp_all & ": Shelter Benefit - $" & dwp_shel & ", Personal Needs - $" & dwp_pers & ". " & dwp_participants & " people on the grant."

	End If

	If GA_active = TRUE Then
		Call Navigate_to_MAXIS_screen ("ELIG", "GA")
		EMWriteScreen "99", 20, 78
		transmit
		mx_row = 7
		Do
			EMReadScreen approval_status, 10, mx_row, 50
			approval_status = trim(approval_status)
			If approval_status = "APPROVED" Then
				EMReadScreen version_number, 2, mx_row, 22
				version_number = trim(version_number)
				EMWriteScreen version_number, 18, 54
				transmit
				Exit Do
			End If
			mx_row = mx_row + 1
		Loop until approval_status = ""

		ga_participants = 0
		For all_clients = 0 to UBOUND(CASE_INFO_ARRAY, 2)
			MX_row = 8
			Do
				EMReadScreen elig_ref, 2, mx_row, 9
				If elig_ref = CASE_INFO_ARRAY(clt_ref, all_clients) Then
					EMReadScreen elig_status, 4, mx_row, 57
					elig_status = trim(elig_status)
					If elig_status = "ELIG" Then
						CASE_INFO_ARRAY (include_adult_cash, all_clients) = checked
						ga_participants = ga_participants + 1
					Else
						CASE_INFO_ARRAY (include_adult_cash, all_clients) = unchecked
					End If
					Exit Do
				Else
					mx_row = mx_row + 1
				End If
			Loop until elig_ref = "  "
		Next

		EMWriteScreen "GASM", 20, 70
		transmit

		EMReadScreen ga_grant,6, 14, 74

		ga_grant = trim(ga_grant)

		adult_cash_current_benefit = adult_cash_current_benefit & "GA - $" & ga_grant & ". HH Size: " & ga_participants
	End If

	If MSA_active = TRUE Then
		Call Navigate_to_MAXIS_screen ("ELIG", "MSA")

		EMWriteScreen "99", 20, 79
		transmit

		mx_row = 7
		Do
			EMReadScreen approval_status, 10, mx_row, 50
			approval_status = trim(approval_status)
			If approval_status = "APPROVED" Then
				EMReadScreen version_number, 2, mx_row, 22
				version_number = trim(version_number)
				EMWriteScreen version_number, 18, 54
				transmit
				Exit Do
			End If
			mx_row = mx_row + 1
		Loop until approval_status = ""

		msa_participants = 0
		For all_clients = 0 to UBOUND(CASE_INFO_ARRAY, 2)
			MX_row = 7
			Do
				EMReadScreen elig_ref, 2, mx_row, 5
				If elig_ref = CASE_INFO_ARRAY(clt_ref, all_clients) Then
					EMReadScreen elig_status, 10, mx_row, 46
					elig_status = trim(elig_status)
					If elig_status = "ELIGIBLE" Then
						CASE_INFO_ARRAY (include_adult_cash, all_clients) = checked
						msa_participants = msa_participants + 1
					Else
						CASE_INFO_ARRAY (include_adult_cash, all_clients) = unchecked
					End If
					Exit Do
				Else
					mx_row = mx_row + 1
				End If
			Loop until elig_ref = "  "
		Next

		transmit
		transmit

		EMWriteScreen "X", 6, 43
		transmit

		sn_row = 8
		sn_col = 6
		Do
			EMReadScreen sn_type, 2, sn_row, sn_col
			If sn_type = "__" Then Exit Do

			If sn_type = "RP" Then sn_rep_payee_checkbox 			= checked
			If sn_type = "GF" Then sn_guardian_checkbox				= checked
			If sn_type = "RM" Then sn_restaraunt_meals_checkbox		= checked
			If sn_type = "SN" Then sn_housing_assistance_checkbox	= checked
			If sn_type = "09" Then sn_anti_dumping_checkbox			= checked
			If sn_type = "02" Then sn_control_protien_60_checkbox	= checked
			If sn_type = "03" Then sn_control_protien_40_checkbox	= checked
			If sn_type = "07" Then sn_gluten_free_checkbox			= checked
			If sn_type = "01" Then sn_high_protien_checkbox			= checked
			If sn_type = "05" Then sn_high_residue_checkbox			= checked
			If sn_type = "10" Then sn_hypoglycemic_checkbox			= checked
			If sn_type = "11" Then sn_ketogenic_checkbox			= checked
			If sn_type = "08" Then sn_lactose_free_checkbox			= checked
		 	If sn_type = "04" Then sn_low_cholesterol_checkbox		= checked
			If sn_type = "06" Then sn_pregnancy_lactation_checkbox 	= checked

			sn_row = sn_row + 1
			If sn_row = 14 Then
				sn_col = sn_col + 36
				sn_row = 8
			End If
		Loop until sn_col = 78
		PF3

		transmit
		transmit

		EMReadScreen msa_grant, 8, 17, 73

		msa_grant = trim(msa_grant)

		adult_cash_current_benefit = adult_cash_current_benefit & "MSA - $"  & msa_grant & " HH Size: " & msa_participants
	End If

	'Finding if DISA
	Call Navigate_to_MAXIS_screen("STAT", "DISA")
	cash_basis_met_checkbox = checked
	cash_HH_size = 0
	For client = 0 to UBOUND(CASE_INFO_ARRAY, 2)
		If CASE_INFO_ARRAY(include_adult_cash, client) = checked Then cash_HH_size = cash_HH_size + 1
		EMWriteScreen CASE_INFO_ARRAY(clt_ref, client), 20, 76
		transmit
		 EMReadScreen disa_end_date, 10, 6, 69
		 disa_end_date = replace(disa_end_date, " ", "/")
		 If disa_end_date = "__/__/____" Then
			EMReadScreen disa_start_date, 10, 6, 47
			If disa_start_date <> "__ __ ____" Then
				If CASE_INFO_ARRAY(include_snap, client) = checked Then elderly_disabled_checkbox = checked
				DISA_Case = TRUE
			End If
		ElseIf DateDiff("d", disa_end_date, date) < 0 Then
			If CASE_INFO_ARRAY(include_snap, client) = checked Then elderly_disabled_checkbox = checked
			DISA_Case = TRUE
		Else
			If CASE_INFO_ARRAY(clt_age, client) < 65 AND CASE_INFO_ARRAY(include_adult_cash, client) = checked Then cash_basis_met_checkbox = unchecked
		End If
	Next

	If cash_HH_size = 0 Then cash_basis_met_checkbox = unchecked

	If GA_active = TRUE Then cash_basis_met_checkbox = checked
	If MSA_active = TRUE Then cash_basis_met_checkbox = checked

End If

If no_case_number_checkbox = unchecked Then MEMB_function
If no_case_number_checkbox = checked Then NEW_CASE_MEMB_FUNCTION
ProgramEstimate

If SNAP_active = False Then snap_current_benefit = "Not Active"
If MFIP_active = False AND DWP_active = False Then family_cash_current_benefit = "Not Active"
If GA_active = False AND MSA_active = False Then adult_cash_current_benefit = "Not Active"

Do
	main_err_msg = ""

	ProgramEstimate

	total_earned_income = 0
	total_unearned_income = 0
	total_assets = 0

	snap_hh_size = snap_hh_size & ""
	family_cash_hh_size = family_cash_hh_size & ""
	adult_cash_hh_size = adult_cash_hh_size & ""

	total_case_assets = total_liquid_assets + total_other_assets

	For client_in_case = 0 to UBOUND(CASE_INFO_ARRAY, 2)
		'total_earned_income = total_earned_income + CASE_INFO_ARRAY(clt_ei_gross, client_in_case)
		'total_unearned_income = total_unearned_income +  CASE_INFO_ARRAY(clt_ssi_income , client_in_case) + CASE_INFO_ARRAY(clt_rsdi_income , client_in_case) + CASE_INFO_ARRAY(clt_other_unea_1_amt , client_in_case) + CASE_INFO_ARRAY(clt_other_unea_2_amt , client_in_case)
		'total_assets = total_assets + CASE_INFO_ARRAY(clt_asset_total , client_in_case)
	Next

	total_earned_income = case_ei_gross
''	MsgBox FPG_165_Amt
	Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 400, 335, "Budget Estimator"

	  GroupBox 5, 5, 195, 60, "Household Composition"
	  Text 15, 20, 50, 10, "SNAP HH Size"
	  DropListBox 75, 15, 30, 45, "0"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"5"+chr(9)+"6"+chr(9)+"7"+chr(9)+"8"+chr(9)+"9"+chr(9)+"10"+chr(9)+"11"+chr(9)+"12"+chr(9)+"13"+chr(9)+"14"+chr(9)+"15"+chr(9)+"16"+chr(9)+"17"+chr(9)+"18"+chr(9)+"19"+chr(9)+"20", snap_hh_size
	  ButtonGroup ButtonPressed
	    PushButton 165, 15, 25, 10, "MEMB", MEMB_BUTTON
	  Text 15, 35, 55, 10, "Family Cash HH"
	  DropListBox 75, 30, 30, 45, "0"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"5"+chr(9)+"6"+chr(9)+"7"+chr(9)+"8"+chr(9)+"9"+chr(9)+"10"+chr(9)+"11"+chr(9)+"12"+chr(9)+"13"+chr(9)+"14"+chr(9)+"15"+chr(9)+"16"+chr(9)+"17"+chr(9)+"18"+chr(9)+"19"+chr(9)+"20", family_cash_hh_size
	  Text 15, 50, 50, 10, "Adult Cash HH"
	  DropListBox 75, 45, 30, 45, "0"+chr(9)+"1"+chr(9)+"2", adult_cash_hh_size

	  GroupBox 5, 60, 195, 70, "Income"
	  Text 15, 70, 180, 10, "TOTAL EARNED - $" & total_earned_income &  " | TOTAL UNEA - $" & total_case_unea
	  ButtonGroup ButtonPressed
		PushButton 15, 90, 110, 10, "Calculate Earned - JOBS & BUSI", jobs_button
		PushButton 15, 105, 80, 10, "Calculate Unearned", unea_button

	  GroupBox 5, 125, 195, 60, "Assets"
	  Text 15, 140, 170, 10, "TOTAL ASSETS - $" & total_case_assets
	  ButtonGroup ButtonPressed
	    PushButton 15, 155, 85, 10, "Calculate Liquid Assets", assets_button
		PushButton 15, 170, 85, 10, "Calculate Other Assets", other_assets_button

	  GroupBox 5, 185, 195, 70, "Expenses"
	  Text 15, 195, 170, 10, "TOTAL COUNTED EXPENSES - $" & total_case_expenses
	  ButtonGroup ButtonPressed
	    PushButton 15, 210, 110, 10, "Calculate Shelter & Utilities", shelter_button
	    PushButton 15, 235, 65, 10, "FMED", fmed_button
	    PushButton 15, 225, 65, 10, "DCEX", child_care_button
	    PushButton 95, 225, 65, 10, "COEX", child_support_button

	  GroupBox 5, 250, 195, 45, "Program Specific"
	  ButtonGroup ButtonPressed
	  	PushButton 15, 265, 70, 10, "EMPS/TIME/SANC", fam_cash_button
	  	PushButton 15, 275, 65, 10, "WREG", wreg_button
	  	PushButton 95, 275, 65, 10, "DISA", disa_button
		PushButton 95, 265, 65, 10, "Special Needs", msa_sn_button

	  If no_case_number_checkbox = unchecked Then CheckBox 15, 300, 350, 10, "Check here to have the script case note that a program information/eligibility discussion happened", case_note_checkbox
	  CheckBox 15, 315, 300, 10, "Check here to have the script create a Word Doc of the information input.", word_doc_checkbox

	  GroupBox 210, 5, 185, 290, "Program Eligibility"

	  GroupBox 210, 20, 185, 90, "SNAP"
	  Text 220, 35, 170, 10, "Current: " & snap_current_benefit
	  Text 220, 50, 170, 10, "Estimated: " & snap_estimated_benefit
	  Text 220, 65, 170, 35, "Notes: " & SNAP_notes

	  GroupBox 210, 105, 185, 110, "Family Cash"
	  Text 215, 115, 175, 40, "Current:" & family_cash_current_benefit
	  Text 215, 145, 175, 20, "Estimated:" & family_cash_estimated_benefit
	  Text 215, 170, 175, 40, "Notes:" & family_cash_notes

	  GroupBox 210, 210, 185, 85, "Adult Cash"
	  Text 220, 225, 175, 10, "Current:" & adult_cash_current_benefit
	  Text 220, 240, 175, 10, "Estimated:" & adult_cash_estimated_benefit
	  Text 220, 255, 175, 35, "Notes:" & adult_cash_notes

	  ButtonGroup ButtonPressed
	    OkButton 295, 315, 50, 15
		CancelButton 345, 315, 50, 15

	EndDialog

	Dialog Dialog1

	cancel_confirmation

	rent_expense = rent_expense
	prop_tax_expense = prop_tax_expense
	home_ins_expense = home_ins_expense
	other_expense = other_expense
	actual_utility_expense = actual_utility_expense

	heat_ac_checkbox 	= heat_ac_checkbox
	electric_checkbox 	= electric_checkbox
	phone_checkbox 		= phone_checkbox
	subsidy_checkbox	= subsidy_checkbox

	sn_rep_payee_checkbox			= sn_rep_payee_checkbox
	sn_guardian_checkbox			= sn_guardian_checkbox
	sn_restaraunt_meals_checkbox	= sn_restaraunt_meals_checkbox
	sn_housing_assistance_checkbox	= sn_housing_assistance_checkbox
	sn_anti_dumping_checkbox		= sn_anti_dumping_checkbox
	sn_control_protien_60_checkbox	= sn_control_protien_60_checkbox
	sn_control_protien_40_checkbox	= sn_control_protien_40_checkbox
	sn_gluten_free_checkbox			= sn_gluten_free_checkbox
	sn_high_protien_checkbox		= sn_high_protien_checkbox
	sn_high_residue_checkbox		= sn_high_residue_checkbox
	sn_hypoglycemic_checkbox		= sn_hypoglycemic_checkbox
	sn_ketogenic_checkbox			= sn_ketogenic_checkbox
	sn_lactose_free_checkbox		= sn_lactose_free_checkbox
	sn_low_cholesterol_checkbox		= sn_low_cholesterol_checkbox
	sn_pregnancy_lactation_checkbox	=  sn_pregnancy_lactation_checkbox

	snap_hh_size 			= snap_hh_size
	family_cash_hh_size 	= family_cash_hh_size
	adult_cash_hh_size 		= adult_cash_hh_size

	elderly_disabled_checkbox 		= elderly_disabled_checkbox
	ten_percent_sanc_checkbox 		= ten_percent_sanc_checkbox
	thirty_percent_sanc_checkbox 	= thirty_percent_sanc_checkbox

	number_of_adults = number_of_adults * 1
	number_of_children = number_of_children * 1
	snap_hh_size = snap_hh_size * 1
	family_cash_hh_size = family_cash_hh_size * 1
	adult_cash_hh_size = adult_cash_hh_size * 1

	Select Case ButtonPressed
	Case MEMB_BUTTON
		main_err_msg = "LOOP"
		If no_case_number_checkbox = unchecked Then MEMB_function
		If no_case_number_checkbox = checked Then NEW_CASE_MEMB_FUNCTION
	Case jobs_button
		main_err_msg = "LOOP"
		EARNED_INCOME_BUTTON_PRESSED
	Case unea_button
		main_err_msg = "LOOP"
		UNEA_BUTTON_PRESSED
	Case assets_button
		main_err_msg = "LOOP"
		ASSETS_BUTTON_PRESSED
	Case other_assets_button
		main_err_msg = "LOOP"
		OTHER_ASSETS_BUTTON_PRESSED
	Case shelter_button
		main_err_msg = "LOOP"
		SHELTER_BUTTON_PRESSED
	Case fmed_button
		main_err_msg = "LOOP"
		DCEX_COEX_FMED_BUTTON_PRESSED
	Case child_care_button
		main_err_msg = "LOOP"
		DCEX_COEX_FMED_BUTTON_PRESSED
	Case child_support_button
		main_err_msg = "LOOP"
		DCEX_COEX_FMED_BUTTON_PRESSED
	Case fam_cash_button
		main_err_msg = "LOOP"
		PROGRAM_SPECIFIC_BUTTON_PRESSED
	Case wreg_button
		main_err_msg = "LOOP"
		PROGRAM_SPECIFIC_BUTTON_PRESSED
	Case disa_button
		main_err_msg = "LOOP"
		PROGRAM_SPECIFIC_BUTTON_PRESSED
	Case msa_sn_button
		main_err_msg = "LOOP"
		MSA_SPECIAL_NEEDS
	End Select

Loop Until main_err_msg = ""

If word_doc_checkbox = checked Then
	'OPENING WORD - THIS IS HOW THE INFORMATION IS COLLECTED
	Set objWord = CreateObject("Word.Application")

	'Opening the first document - the case information collection
	Set objScreenDoc = objWord.Documents.Add()
	objWord.visible = True

	With objScreenDoc.PageSetup
		.TopMargin 		= objWord.InchesToPoints(.5)
		.BottomMargin 	= objWord.InchesToPoints(.5)
		.LeftMargin 	= objWord.InchesToPoints(.5)
		.RightMargin 	= objWord.InchesToPoints(.5)
	End With

	Set objScreenSelect = objWord.Selection

	objScreenSelect.Font.Name = "calibri"
	objScreenSelect.Font.Size = "13"
	objScreenSelect.ParagraphFormat.SpaceAfter = 0

	objScreenSelect.ParagraphFormat.Alignment = 1
	objScreenSelect.Font.Bold = True
	objScreenSelect.TypeText "Case Budget Estimate Summary"
	objScreenSelect.TypeParagraph()

	objScreenSelect.Font.Bold = false
	objScreenSelect.TypeText "___________________________________________________________________________________" & chr(13)
	objScreenSelect.ParagraphFormat.Alignment = 0

	If no_case_number_checkbox = unchecked Then
		objScreenSelect.Font.Bold = True
		objScreenSelect.TypeText "Case #: " & MAXIS_case_number & chr(9) & chr(9) & chr(9) & chr(9) & "Case Name: " & CASE_INFO_ARRAY(clt_name, 0)
		objScreenSelect.TypeParagraph()
		objScreenSelect.Font.Bold = false

		objScreenSelect.TypeText chr(9) & "Current SNAP: " & snap_current_benefit
		objScreenSelect.TypeParagraph()

		objScreenSelect.TypeText chr(9) & "Current Family Cash: " & family_cash_current_benefit
		objScreenSelect.TypeParagraph()

		objScreenSelect.TypeText chr(9) & "Current Adult Cash: " & adult_cash_current_benefit
		objScreenSelect.TypeParagraph()
	Else
		objScreenSelect.Font.Bold = True
		objScreenSelect.TypeText "Case is not known in MAXIS - all information manually entered"
		objScreenSelect.TypeParagraph()
		objScreenSelect.Font.Bold = false
	End If

	objScreenSelect.TypeText "___________________________________________________________________________________" & chr(13)

	objScreenSelect.Font.Bold = True
	objScreenSelect.TypeText "Income - Total Earned: $" & total_earned_income & " - Total UNEA: $" & total_case_unea
	objScreenSelect.TypeParagraph()
	objScreenSelect.Font.Bold = false
	objScreenSelect.Font.Underline = True
	objScreenSelect.TypeText "Earned Income" & chr(13)
	objScreenSelect.Font.Underline = False
	For each_job = 0 to UBOUND(EI_ARRAY, 2)
		If EI_ARRAY(employee, each_job) <> "" Then
			objScreenSelect.TypeText EI_ARRAY(employee, each_job) & " employed at " & EI_ARRAY(employer, each_job) & chr(13) & chr(9) & "Retro Gross $" & EI_ARRAY(job_retro_gross, each_job) & " - Prosp Gross $" & EI_ARRAY(job_prosp_gross, each_job) & " - PIC Gross $" & EI_ARRAY(job_pic_gross, each_job) & chr(13)
			If EI_ARRAY(pic_hrs_wk, each_job) <> 0 OR EI_ARRAY(pic_rate_pay, each_job) <> 0 Then objScreenSelect.TypeText chr(9) & EI_ARRAY(pic_hrs_wk, each_job) & " Hours/Week" & " at $" & EI_ARRAY(pic_rate_pay, each_job) & "/hour. Paid " & EI_ARRAY(job_pay_freq, each_job) & chr(13)
			If EI_ARRAY(how_many_chck, each_job) <> " " Then
				If EI_ARRAY(check_1_gross, each_job) <> "" Then objScreenSelect.TypeText chr(9) & "Checks Received:" & chr(13)
				For checks = 0 to EI_ARRAY(how_many_chck, each_job)
					If EI_ARRAY(7 + 2 * checks, each_job) <> "" Then
						objScreenSelect.TypeText chr(9) & "On " & EI_ARRAY(7 + 2 * checks, each_job) & " for $" & EI_ARRAY(8 + 2 * checks, each_job) & chr(13)
					End If
				Next
			End If
		End If
	Next
	objScreenSelect.Font.Underline = True
	objScreenSelect.TypeText "Unearned Income" & chr(13)
	objScreenSelect.Font.Underline = False
	If no_case_number_checkbox = unchecked Then
		For clients = 0 to UBOUND(CASE_INFO_ARRAY, 2)
			If CASE_INFO_ARRAY(clt_ssi_income, clients) <> 0 Then objScreenSelect.TypeText "SSI - $" & CASE_INFO_ARRAY(clt_ssi_income, clients) & " for " & CASE_INFO_ARRAY(clt_name, clients) & CASE_INFO_ARRAY(clt_a_c, clients) & chr(13)
			If CASE_INFO_ARRAY(clt_rsdi_income, clients) <> 0 Then objScreenSelect.TypeText "RSDI - $" & CASE_INFO_ARRAY(clt_rsdi_income, clients) & " for " & CASE_INFO_ARRAY(clt_name, clients) & CASE_INFO_ARRAY(clt_a_c, clients) & chr(13)
			If CASE_INFO_ARRAY(clt_other_unea_1_amt, clients) <> 0 Then objScreenSelect.TypeText CASE_INFO_ARRAY(clt_other_unea_1_type, clients) & " - $" & CASE_INFO_ARRAY(clt_other_unea_1_amt, clients) & " for " & CASE_INFO_ARRAY(clt_name, clients) & CASE_INFO_ARRAY(clt_a_c, clients) & chr(13)
			If CASE_INFO_ARRAY(clt_other_unea_2_amt, clients) <> 0 Then objScreenSelect.TypeText CASE_INFO_ARRAY(clt_other_unea_2_type, clients) & " - $" & CASE_INFO_ARRAY(clt_other_unea_2_amt, clients) & " for " & CASE_INFO_ARRAY(clt_name, clients) & CASE_INFO_ARRAY(clt_a_c, clients) & chr(13)
		Next
	Else
		For each_unea = 0 to UBOUND(CASE_UNEA_ARRAY, 2)
			If CASE_UNEA_ARRAY(ssi_amt, each_unea) <> 0 Then objScreenSelect.TypeText "SSI - $" & CASE_UNEA_ARRAY(ssi_amt, each_unea) & " for " & CASE_UNEA_ARRAY(unea_person, each_unea) & chr(13)
			If CASE_UNEA_ARRAY(rsdi_amt, each_unea) <> 0 Then objScreenSelect.TypeText "RSDI - $" & CASE_UNEA_ARRAY(rsdi_amt, each_unea) & " for " & CASE_UNEA_ARRAY(unea_person, each_unea) & chr(13)
			If CASE_UNEA_ARRAY(other_1_amt, each_unea) <> 0 Then objScreenSelect.TypeText CASE_UNEA_ARRAY(other_1_type, each_unea) & " - $" & CASE_UNEA_ARRAY(other_1_amt, each_unea) & " for " & CASE_UNEA_ARRAY(unea_person, each_unea) & chr(13)
			If CASE_INFO_ARRAY(other_2_amt, each_unea) <> 0 Then objScreenSelect.TypeText CASE_UNEA_ARRAY(other_2_type, each_unea) & " - $" & CASE_UNEA_ARRAY(other_2_amt, each_unea) & " for " & CASE_UNEA_ARRAY(unea_person, each_unea) & chr(13)
		Next
	End If

	objScreenSelect.TypeParagraph()
	objScreenSelect.Font.Bold = True
	objScreenSelect.TypeText "Expenses"
	objScreenSelect.TypeParagraph()
	objScreenSelect.Font.Bold = false
	objScreenSelect.TypeText "Total Shelter Expense $" & (rent_expense + prop_tax_expense + home_ins_expense + other_expense) & chr(13)
	objScreenSelect.TypeText "Utilities: "
	If heat_ac_checkbox = checked Then
		objScreenSelect.TypeText "Heat/AC - $" & heat_AC_amt
	ElseIf electric_checkbox = checked Then
		If phone_checkbox = checked Then objScreenSelect.TypeText "Electric = $" & electric_amt & " & Phone - $" & phone_amt
		If phone_checkbox = unchecked Then objScreenSelect.TypeText "Electric = $" & electric_amt
	Else
		If phone_checkbox = checked Then objScreenSelect.TypeText "Phone - $" & phone_amt
		If phone_checkbox = unchecked Then objScreenSelect.TypeText "NONE"
	End If
	objScreenSelect.TypeParagraph()
	objScreenSelect.Font.Underline = True
	objScreenSelect.TypeText "Other Monthly Expenses"
	objScreenSelect.Font.Underline = False
	objScreenSelect.TypeParagraph()
	If monthly_childcare_exp <> 0 Then objScreenSelect.TypeText chr(9) & "Child Care Expense - $" & monthly_childcare_exp & chr(13)
	If monthly_adultcare_exp <> 0 Then objScreenSelect.TypeText chr(9) & "Adult Care Expense - $" & monthly_adultcare_exp & chr(13)
	If child_support_exp <> 0 Then objScreenSelect.TypeText chr(9) & "Child Support Paid - $" & child_support_exp & chr(13)
	If alimony_exp <> 0 Then objScreenSelect.TypeText chr(9) & "Alimony Paid - $" & alimony_exp & chr(13)
	If monthly_fmed_exp <> 0 Then objScreenSelect.TypeText chr(9) & "Medical Expenses (FMED) - $" & monthly_fmed_exp & chr(13)

	objScreenSelect.TypeParagraph()
	objScreenSelect.Font.Bold = True
	objScreenSelect.TypeText "Assets - $" & total_case_assets
	objScreenSelect.TypeParagraph()
	objScreenSelect.Font.Bold = false
	objScreenSelect.Font.Underline = True
	objScreenSelect.TypeText "Accounts - Liquid Assets"
	objScreenSelect.Font.Underline = False
	objScreenSelect.TypeParagraph()
	If no_case_number_checkbox = unchecked Then
		For clients = 0 to UBOUND(CASE_INFO_ARRAY, 2)
			If CASE_INFO_ARRAY(clt_sav_acct, clients) <> 0 Then objScreenSelect.TypeText "Savings - $" & CASE_INFO_ARRAY(clt_sav_acct, clients) & " of " & CASE_INFO_ARRAY(clt_name, clients) & chr(13)
			If CASE_INFO_ARRAY(clt_chk_acct, clients) <> 0 Then objScreenSelect.TypeText "Checking - $" & CASE_INFO_ARRAY(clt_chk_acct, clients) & " of " & CASE_INFO_ARRAY(clt_name, clients) & chr(13)
			If CASE_INFO_ARRAY(clt_asset_other_bal, clients) <> 0 Then objScreenSelect.TypeText CASE_INFO_ARRAY(clt_asset_other_type, clients) & " - $" & CASE_INFO_ARRAY(clt_asset_other_bal, clients) & " of " & CASE_INFO_ARRAY(clt_name, clients) & chr(13)
		Next
	Else
		For account = 0 to UBOUND(CASE_ACCOUNTS_ARRAY, 2)
		  objScreenSelect.TypeText CASE_ACCOUNTS_ARRAY(account_type, account) & " - $" & CASE_ACCOUNTS_ARRAY(account_balance, account) & " of " & CASE_ACCOUNTS_ARRAY(account_holder, account)
		Next
	End If
	objScreenSelect.Font.Underline = True
	objScreenSelect.TypeText "Other Assets"
	objScreenSelect.Font.Underline = False
	objScreenSelect.TypeParagraph()
	For vehicle = 0 to UBOUND (VEHICLE_ARRAY, 2)
		If VEHICLE_ARRAY(vehicle_type, vehicle) <> "Select One ..." AND VEHICLE_ARRAY(vehicle_type, vehicle) <> "" Then objScreenSelect.TypeText right(VEHICLE_ARRAY(vehicle_type, vehicle), len(VEHICLE_ARRAY(vehicle_type, vehicle))-4) & ": " & VEHICLE_ARRAY(vehicle_year, vehicle) & " - " & VEHICLE_ARRAY(vehicle_make, vehicle) & " " & VEHICLE_ARRAY(vehicle_model, vehicle) & " Value ~ $" & VEHICLE_ARRAY(vehicle_value, vehicle) & chr(13)
	Next
	For security = 0 to UBOUND (SECURITIES_ARRAY, 2)
		If SECURITIES_ARRAY(security_type, security) <> "Select One ..." AND SECURITIES_ARRAY(security_type, security) <> "" Then objScreenSelect.TypeText right(SECURITIES_ARRAY(security_type, security), len(SECURITIES_ARRAY(security_type, security))-5) & " - " & SECURITIES_ARRAY(security_description, security) & " Value ~ $" & SECURITIES_ARRAY(security_value, security)
		If SECURITIES_ARRAY(security_withdrawl, security) <> 0 Then objScreenSelect.TypeText ", Withdrawl Penalty $" & SECURITIES_ARRAY(security_withdrawl, security)
		objScreenSelect.TypeParagraph()
	Next

	objScreenSelect.TypeText "___________________________________________________________________________________" & chr(13)

	objScreenSelect.Font.Bold = True
	objScreenSelect.TypeText "Program Estimates"
	objScreenSelect.TypeParagraph()
	objScreenSelect.Font.Bold = false

	objScreenSelect.TypeText "SNAP - unit size: " & snap_hh_size & chr(13)
	objScreenSelect.TypeText chr(9) & "Estimated Benefit: " & snap_estimated_benefit & chr(13)
	objScreenSelect.TypeText chr(9) & "Notes about benefit: " & SNAP_notes & chr(13)

	objScreenSelect.TypeText "Family Cash - unit size: " & family_cash_hh_size & chr(13)
	objScreenSelect.TypeText chr(9) & "Estimated Benefit: " & family_cash_estimated_benefit & chr(13)
	objScreenSelect.TypeText chr(9) & "Notes about benefit: " & family_cash_notes & chr(13)

	objScreenSelect.TypeText "Adult Cash - unit size: " & adult_cash_hh_size & chr(13)
	objScreenSelect.TypeText chr(9) & "Estimated Benefit: " & adult_cash_estimated_benefit & chr(13)
	objScreenSelect.TypeText chr(9) & "Notes about benefit: " & adult_cash_notes & chr(13)

	objScreenSelect.TypeText "___________________________________________________________________________________" & chr(13)
End if

If case_note_checkbox = checked Then


Dialog1 = ""
If SNAP_active = FALSE Then BeginDialog Dialog1, 0, 0, 221, 250, "Case Note Budget Discussion"
If SNAP_active = TRUE Then BeginDialog Dialog1, 0, 0, 221, 145, "Case Note Budget Discussion"
  Text 10, 10, 205, 40, "This case note will indicate that a discussion happened with the client about potential changes to benefit. The case note will not list an estimated benefit as this estimate is not a promise to the client. Client will need to provide verifications/application and STAT updated for the benefit to change. "
  Text 10, 65, 120, 10, "How did this discussion take place?"
  ComboBox 135, 60, 80, 45, " "+chr(9)+"In Person"+chr(9)+"On Phone", contact_type
  If SNAP_active = FALSE Then
	  GroupBox 10, 85, 205, 95, "Poential New SNAP Request"
	  Text 20, 100, 165, 10, "Check everything that was explained to the client."
	  CheckBox 20, 115, 165, 10, "How to apply and submit application explained.", how_to_apply_checkbox
	  CheckBox 20, 130, 90, 10, "An interview is required.", interview_needed_checkbox
	  CheckBox 20, 145, 130, 10, "Importance of CAF1 date/filing date.", caf_date_checkbox
	  CheckBox 20, 160, 125, 10, "About Expedited SNAP Processing.", xfs_possible_checkbox
	  Text 10, 190, 25, 10, "Notes"
	  EditBox 40, 185, 175, 15, other_notes
	  Text 10, 210, 60, 10, "Worker Signature"
	  EditBox 75, 205, 140, 15, worker_signature
	  ButtonGroup ButtonPressed
	    OkButton 110, 230, 50, 15
	    CancelButton 165, 230, 50, 15
	Else
		Text 10, 85, 25, 10, "Notes"
		EditBox 40, 80, 175, 15, other_notes
		Text 10, 105, 60, 10, "Worker Signature"
		EditBox 75, 100, 140, 15, worker_signature
		ButtonGroup ButtonPressed
		  OkButton 110, 125, 50, 15
		  CancelButton 165, 125, 50, 15
	End If
EndDialog

Do
	err_msg = ""
	Dialog Dialog1
	cancel_confirmation

	If contact_type = " " Then err_msg = err_msg & vbnewLine & "Indicate how contact wsa made with client (phone, in person, etc)."
	If worker_signature = "" Then err_msg = err_msg & vbnewLine & "Enter worker name for the case note."

	If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbnewLine & err_msg

Loop until err_msg = ""

Call start_a_blank_CASE_NOTE

Call write_variable_in_CASE_NOTE ("Spoke with client RE: Cash/SNAP Program Information/Eligibility")
Call write_variable_in_CASE_NOTE ("Spoke to clt " & contact_type & " today - " & date)
Call write_variable_in_CASE_NOTE ("Discussed possible cash and food program eligibility with client.")
If SNAP_active = FALSE Then Call write_variable_in_CASE_NOTE ("---")
If SNAP_active = FALSE Then Call write_variable_in_CASE_NOTE ("** Case is NOT currently active SNAP. **")
If how_to_apply_checkbox = checked Then Call write_variable_in_CASE_NOTE ("* Advised client of how to apply and submit an application for SNAP.")
If interview_needed_checkbox = checked Then Call write_variable_in_CASE_NOTE ("* Explained to client that an interview must be completed for SNAP and that this can be completed by phone")
If caf_date_checkbox = checked Then Call write_variable_in_CASE_NOTE ("* Advised client that benefits are started from the date the completed CAF1 is received.")
If xfs_possible_checkbox = checked Then Call write_variable_in_CASE_NOTE ("* Explained about expedited processing.")
Call write_bullet_and_variable_in_CASE_NOTE ("Notes", other_notes)
Call write_variable_in_CASE_NOTE ("---")
Call write_variable_in_CASE_NOTE (worker_signature)

End If

script_end_procedure_with_error_report ("Success! Script has been completed.")
