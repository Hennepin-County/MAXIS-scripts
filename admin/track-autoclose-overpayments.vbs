'Required for statistical purposes==========================================================================================
name_of_script = "ADMIN - OVERPAYMENT ON AUTOCLOSE PAUSE.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 180          'manual run time in seconds
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

function ensure_variable_is_a_number(variable_here, decimal_places)
	If variable_here = "" Then variable_here = 0
	If IsNumeric(variable_here) = False Then variable_here = 0
	variable_here = FormatNumber(variable_here, decimal_places, -1, 0, 0)
	variable_here = variable_here *1
end function

function budget_calculate_income(earned_income_correct_amt, unearned_correct_amt, earned_deduction_correct_amt, total_income_correct_amt, output_type)
	' output_type - "STRING" or "NUMBER"
	Call ensure_variable_is_a_number(earned_income_correct_amt, 2)
	Call ensure_variable_is_a_number(unearned_correct_amt, 2)

	'TODO - To make this global we need to handle for if the earned income deduction is allowed or not
	total_income_correct_amt = earned_income_correct_amt + unearned_correct_amt
	earned_deduction_correct_amt = 0.2 * earned_income_correct_amt

	total_income_correct_amt = FormatNumber(total_income_correct_amt, 2, -1, 0, -1)
	earned_deduction_correct_amt = FormatNumber(earned_deduction_correct_amt, 2, -1, 0, -1)

	If UCase(output_type) = "STRING" Then
		earned_income_correct_amt = earned_income_correct_amt & ""
		unearned_correct_amt = unearned_correct_amt & ""
		total_income_correct_amt = total_income_correct_amt & ""
		earned_deduction_correct_amt = earned_deduction_correct_amt & ""
	End If
end function

function budget_calculate_household(correct_hh_size, disa_household, cat_elig, standard_deduction_correct_amt, max_shelter_cost_correct_amt, max_gross_income_correct_amt, max_net_adj_income_correct_amt, max_snap_benefit, output_type)
	' disa_household - True/False
	' cat_elig - True/False
	' output_type - "STRING" or "NUMBER"
	Call ensure_variable_is_a_number(correct_hh_size, 0)

	'TODO - To make this global we need a variable to get the right month for the trifty food plan
	'TODO - To make this global we need to handle for proration
	'THRIFTY FOOD PLAN - CM 22.12.01
	If correct_hh_size = 0 Then
		max_snap_benefit = 0
		standard_deduction_correct_amt = 0
		max_net_adj_income_correct_amt = 0
		max_gross_income_correct_amt = 0
	Elseif correct_hh_size = 1 Then
		max_snap_benefit = 250
		standard_deduction_correct_amt = 177
		max_net_adj_income_correct_amt = 1074
		If cat_elig = True Then max_gross_income_correct_amt = 1771
		If cat_elig = False Then max_gross_income_correct_amt = 1396
	ElseIf correct_hh_size = 2 Then
		max_snap_benefit = 459
		standard_deduction_correct_amt = 177
		max_net_adj_income_correct_amt = 1452
		If cat_elig = True Then max_gross_income_correct_amt = 2396
		If cat_elig = False Then max_gross_income_correct_amt = 1888
	ElseIf correct_hh_size = 3 Then
		max_snap_benefit = 658
		standard_deduction_correct_amt = 177
		max_net_adj_income_correct_amt = 1830
		If cat_elig = True Then max_gross_income_correct_amt = 3020
		If cat_elig = False Then max_gross_income_correct_amt = 2379
	ElseIf correct_hh_size = 4 Then
		max_snap_benefit = 835
		standard_deduction_correct_amt = 184
		max_net_adj_income_correct_amt = 2209
		If cat_elig = True Then max_gross_income_correct_amt = 3644
		If cat_elig = False Then max_gross_income_correct_amt = 2871
	ElseIf correct_hh_size = 5 Then
		max_snap_benefit = 992
		standard_deduction_correct_amt = 215
		max_net_adj_income_correct_amt = 2587
		If cat_elig = True Then max_gross_income_correct_amt = 4268
		If cat_elig = False Then max_gross_income_correct_amt = 3363
	ElseIf correct_hh_size = 6 Then
		max_snap_benefit = 1190
		standard_deduction_correct_amt = 246
		max_net_adj_income_correct_amt = 2965
		If cat_elig = True Then max_gross_income_correct_amt = 4893
		If cat_elig = False Then max_gross_income_correct_amt = 3855
	ElseIf correct_hh_size = 7 Then
		max_snap_benefit = 1316
		standard_deduction_correct_amt = 246
		max_net_adj_income_correct_amt = 3344
		If cat_elig = True Then max_gross_income_correct_amt = 5517
		If cat_elig = False Then max_gross_income_correct_amt = 4347
	ElseIf correct_hh_size = 8 Then
		max_snap_benefit = 1504
		standard_deduction_correct_amt = 246
		max_net_adj_income_correct_amt = 3722
		If cat_elig = True Then max_gross_income_correct_amt = 6141
		If cat_elig = False Then max_gross_income_correct_amt = 4839
	Else
		max_snap_benefit = 1504 + ((correct_hh_size-8) * 188)
		max_net_adj_income_correct_amt = 3722 + ((correct_hh_size-8) * 379)
		standard_deduction_correct_amt = 246
		If cat_elig = True Then max_gross_income_correct_amt = 6141 + ((correct_hh_size-8) * 625)
		If cat_elig = False Then max_gross_income_correct_amt = 4839 + ((correct_hh_size-8) * 492)
	End If

	If disa_household = True Then max_shelter_cost_correct_amt = 0
	If disa_household = False Then max_shelter_cost_correct_amt = 597


	standard_deduction_correct_amt = FormatNumber(standard_deduction_correct_amt, 2, -1, 0, -1)
	max_shelter_cost_correct_amt = FormatNumber(max_shelter_cost_correct_amt, 2, -1, 0, -1)
	max_gross_income_correct_amt = FormatNumber(max_gross_income_correct_amt, 2, -1, 0, -1)
	max_net_adj_income_correct_amt = FormatNumber(max_net_adj_income_correct_amt, 2, -1, 0, -1)
	max_snap_benefit = FormatNumber(max_snap_benefit, 2, -1, 0, -1)

	If UCase(output_type) = "STRING" Then
		correct_hh_size = correct_hh_size & ""
		standard_deduction_correct_amt = standard_deduction_correct_amt & ""
		max_shelter_cost_correct_amt = max_shelter_cost_correct_amt & ""
		max_gross_income_correct_amt = max_gross_income_correct_amt & ""
		max_net_adj_income_correct_amt = max_net_adj_income_correct_amt & ""
		max_snap_benefit = max_snap_benefit & ""
	End If
end function

' function budget_calculate_deductions()
' function budget_calculate_deductions(earned_deduction_correct_amt, dependent_care_deduction_correct_amt, child_support_deduction_correct_amt, standard_deduction_correct_amt, total_deduction_correct_amt, total_income_correct_amt, net_income_correct_amt, fifty_perc_net_income_correct_amt, output_type)
function budget_calculate_deductions(earned_deduction_correct_amt, medical_deduction_correct_amt, dependent_care_deduction_correct_amt, child_support_deduction_correct_amt, standard_deduction_correct_amt, total_deduction_correct_amt, total_income_correct_amt, net_income_correct_amt, fifty_perc_net_income_correct_amt, output_type)
	' output_type - "STRING" or "NUMBER"
	Call ensure_variable_is_a_number(earned_deduction_correct_amt, 2)
	Call ensure_variable_is_a_number(medical_deduction_correct_amt, 2)
	Call ensure_variable_is_a_number(dependent_care_deduction_correct_amt, 2)
	Call ensure_variable_is_a_number(child_support_deduction_correct_amt, 2)
	Call ensure_variable_is_a_number(standard_deduction_correct_amt, 2)
	Call ensure_variable_is_a_number(total_income_correct_amt, 2)

	total_deduction_correct_amt = earned_deduction_correct_amt + medical_deduction_correct_amt + dependent_care_deduction_correct_amt + child_support_deduction_correct_amt + standard_deduction_correct_amt
	net_income_correct_amt = total_income_correct_amt - total_deduction_correct_amt
	fifty_perc_net_income_correct_amt = net_income_correct_amt * 0.5
	Call ensure_variable_is_a_number(fifty_perc_net_income_correct_amt, 2)

	If UCase(output_type) = "STRING" Then
		total_deduction_correct_amt = total_deduction_correct_amt & ""
		earned_deduction_correct_amt = earned_deduction_correct_amt & ""
		medical_deduction_correct_amt = medical_deduction_correct_amt & ""
		dependent_care_deduction_correct_amt = dependent_care_deduction_correct_amt & ""
		child_support_deduction_correct_amt = child_support_deduction_correct_amt & ""
		standard_deduction_correct_amt = standard_deduction_correct_amt & ""

		net_income_correct_amt = net_income_correct_amt & ""
		fifty_perc_net_income_correct_amt = fifty_perc_net_income_correct_amt & ""
	End If
end function

function budget_calculate_shelter_costs(rent_mortgage_correct_amt, tax_correct_amt, insurance_correct_amt, other_cost_correct_amt, utilities_correct_amt, total_shelter_cost_correct_amt, adj_shelter_cost_correct_amt, max_shelter_cost_correct_amt, counted_shelter_cost_correct_amt, fifty_perc_net_income_correct_amt, net_income_correct_amt, net_adj_income_correct_amt, output_type)
	Call ensure_variable_is_a_number(rent_mortgage_correct_amt, 2)
	Call ensure_variable_is_a_number(tax_correct_amt, 2)
	Call ensure_variable_is_a_number(insurance_correct_amt, 2)
	Call ensure_variable_is_a_number(other_cost_correct_amt, 2)
	Call ensure_variable_is_a_number(utilities_correct_amt, 2)
	Call ensure_variable_is_a_number(fifty_perc_net_income_correct_amt, 2)
	Call ensure_variable_is_a_number(max_shelter_cost_correct_amt, 2)
	Call ensure_variable_is_a_number(net_income_correct_amt, 2)

	total_shelter_cost_correct_amt = rent_mortgage_correct_amt + tax_correct_amt + insurance_correct_amt + other_cost_correct_amt + utilities_correct_amt
	adj_shelter_cost_correct_amt = total_shelter_cost_correct_amt - fifty_perc_net_income_correct_amt
	If max_shelter_cost_correct_amt = 0 Then max_shelter_cost_correct_amt = adj_shelter_cost_correct_amt
	If adj_shelter_cost_correct_amt > max_shelter_cost_correct_amt Then
		counted_shelter_cost_correct_amt = max_shelter_cost_correct_amt
	Else
		counted_shelter_cost_correct_amt = adj_shelter_cost_correct_amt
	End If
	net_adj_income_correct_amt = net_income_correct_amt - counted_shelter_cost_correct_amt
	If net_adj_income_correct_amt < 0 Then net_adj_income_correct_amt = 0

	If UCase(output_type) = "STRING" Then
		rent_mortgage_correct_amt = rent_mortgage_correct_amt & ""
		tax_correct_amt = tax_correct_amt & ""
		insurance_correct_amt = insurance_correct_amt & ""
		other_cost_correct_amt = other_cost_correct_amt & ""
		utilities_correct_amt = utilities_correct_amt & ""
		total_shelter_cost_correct_amt = total_shelter_cost_correct_amt & ""

		fifty_perc_net_income_correct_amt = fifty_perc_net_income_correct_amt & ""
		adj_shelter_cost_correct_amt = adj_shelter_cost_correct_amt & ""
		counted_shelter_cost_correct_amt = counted_shelter_cost_correct_amt & ""
		max_shelter_cost_correct_amt = max_shelter_cost_correct_amt & ""
		net_income_correct_amt = net_income_correct_amt & ""
		net_adj_income_correct_amt = net_adj_income_correct_amt & ""
	End If
end function

function budget_calculate_benefit_details(cat_elig, total_income_correct_amt, net_adj_income_correct_amt, max_net_adj_income_correct_amt, max_gross_income_correct_amt, max_snap_benefit, monthly_snap_benefit_correct_amt, snap_issued_amt, overpayment_exists, supplement_exists, snap_overpayment_amt, snap_supplement_amt, output_type)
	' cat_elig - True/Fals

	Call ensure_variable_is_a_number(total_income_correct_amt, 2)
	Call ensure_variable_is_a_number(net_adj_income_correct_amt, 2)
	Call ensure_variable_is_a_number(max_net_adj_income_correct_amt, 2)
	Call ensure_variable_is_a_number(max_gross_income_correct_amt, 2)
	Call ensure_variable_is_a_number(max_snap_benefit, 2)
	Call ensure_variable_is_a_number(snap_issued_amt, 2)

	overpayment_exists = False
	supplement_exists = False
	income_exceeded = False
	snap_overpayment_amt = 0
	snap_supplement_amt = 0

	If cat_elig = True Then
		If total_income_correct_amt > max_gross_income_correct_amt Then income_exceeded = True
	Else
		If net_adj_income_correct_amt > max_net_adj_income_correct_amt Then income_exceeded = True
	End If

	If income_exceeded = False Then
		thirty_perc_of_net_income = 0.3 * net_adj_income_correct_amt
		monthly_snap_benefit_correct_amt = max_snap_benefit - thirty_perc_of_net_income
		monthly_snap_benefit_correct_amt = Int(monthly_snap_benefit_correct_amt)
		Call ensure_variable_is_a_number(monthly_snap_benefit_correct_amt, 2)
	End If
	If monthly_snap_benefit_correct_amt > snap_issued_amt Then
		supplement_exists = True
		snap_supplement_amt = monthly_snap_benefit_correct_amt - snap_issued_amt
	End If
	If monthly_snap_benefit_correct_amt < snap_issued_amt Then
		overpayment_exists = True
		snap_overpayment_amt = snap_issued_amt - monthly_snap_benefit_correct_amt
	End If

	If UCase(output_type) = "STRING" Then
		total_income_correct_amt = total_income_correct_amt & ""
		net_adj_income_correct_amt = net_adj_income_correct_amt & ""
		max_net_adj_income_correct_amt = max_net_adj_income_correct_amt & ""
		max_gross_income_correct_amt = max_gross_income_correct_amt & ""
		max_snap_benefit = max_snap_benefit & ""
		snap_issued_amt = snap_issued_amt & ""
		rent_mortgage_correct_amt = rent_mortgage_correct_amt & ""
		monthly_snap_benefit_correct_amt = monthly_snap_benefit_correct_amt & ""
		snap_overpayment_amt = snap_overpayment_amt & ""
		snap_supplement_amt = snap_supplement_amt & ""
	End if

end function


'Connecting to MAXIS
EMConnect ""
'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
MAXIS_footer_month = "02"
MAXIS_footer_year = "22"

calc_btn = 500
snap_claculation_done_btn = 501

cat_elig = True
disa_household = False

'Grabbing the case number
call MAXIS_case_number_finder(MAXIS_case_number)
back_to_self 'to ensure we are not in edit mode'

'case number dialog
Do
	err_msg = ""
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 166, 100, "Case Number Dialog"
	  EditBox 90, 10, 70, 15, MAXIS_case_number
	  ButtonGroup ButtonPressed
	    OkButton 55, 80, 50, 15
	    CancelButton 110, 80, 50, 15
	  Text 10, 15, 80, 10, "Enter the Case Number:"
	  Text 10, 35, 150, 45, "This script is specific to the detailed review of the cases impacted by the Autoclose Pause that happened in 02/22 and does not take any MAXIS action or create CASE/NOTEs as this process is handled external from MAXIS."
	EndDialog

	dialog Dialog1
	cancel_without_confirmation

	Call validate_MAXIS_case_number(err_msg, "*")

	If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbCr & err_msg
 Loop until err_msg = ""

'read for MFIP/SNAP in 02/22
SNAP_active = False
MFIP_active = False
Call navigate_to_MAXIS_screen("STAT", "PROG")
EMReadScreen cash_1_prog, 2, 6, 67
EMReadScreen cash_1_stat, 4, 6, 74
EMReadScreen cash_2_prog, 2, 7, 67
EMReadScreen cash_2_stat, 4, 7, 74
EMReadScreen snap_prog, 2, 10, 67
EMReadScreen snap_stat, 4, 10, 74

If cash_1_stat = "ACTV" and cash_1_prog = "MF" Then MFIP_active = True
If cash_2_stat = "ACTV" and cash_2_prog = "MF" Then MFIP_active = True
If snap_stat = "ACTV" Then SNAP_active = True
call back_to_self

'Read ELIG for 02/22
If SNAP_active = True Then
	Call navigate_to_MAXIS_screen("ELIG", "FS__")
	EMWriteScreen "99", 19, 78
	transmit
	elig_row = 7
	version_numb = ""
	Do
		EMReadScreen approval_status, 8, elig_row, 50
		If approval_status = "APPROVED" Then
			EMReadScreen version_numb, 2, elig_row, 22
			version_numb = trim(version_numb)
			version_numb = right("00" & version_numb, 2)
			Exit Do
		End If
		elig_row = elig_row + 1
	Loop until approval_status = "        "
	transmit
	EMWriteScreen version_numb, 19, 78
	transmit

	function read_amount_from_MAXIS(variable_here, length, row, col)
		EMReadScreen variable_here, length, row, col
		variable_here = trim(variable_here)
		If variable_here = "" Then variable_here = 0
		If IsNumeric(variable_here) = False Then variable_here = 0
		variable_here = FormatNumber(variable_here, decimal_places, -1, 0, 0)
		' variable_here = variable_here *1
	end function

	write_value_and_transmit "FSCR", 19, 70


	write_value_and_transmit "FSB1", 19, 70
	Call read_amount_from_MAXIS(earned_income_budgeted_amt, 10, 8, 31)
	Call read_amount_from_MAXIS(unearned_budgeted_amt, 10, 18, 31)
	Call read_amount_from_MAXIS(total_income_budgeted_amt, 10, 7, 71)

	Call read_amount_from_MAXIS(standard_deduction_budgeted_amt, 10, 10, 71)
	Call read_amount_from_MAXIS(earned_deduction_budgeted_amt, 10, 11, 71)
	Call read_amount_from_MAXIS(medical_deduction_budgeted_amt, 10, 12, 71)
	Call read_amount_from_MAXIS(dependent_care_deduction_budgeted_amt, 10, 13, 71)
	Call read_amount_from_MAXIS(child_support_deduction_budgeted_amt, 10, 14, 71)
	Call read_amount_from_MAXIS(total_deduction_budgeted_amt, 10, 16, 71)



	write_value_and_transmit "FSB2", 19, 70

	Call read_amount_from_MAXIS(rent_mortgage_budgeted_amt, 10, 5, 27)
	Call read_amount_from_MAXIS(tax_budgeted_amt, 10, 6, 27)
	Call read_amount_from_MAXIS(insurance_budgeted_amt, 10, 7, 27)
	Call read_amount_from_MAXIS(other_cost_budgeted_amt, 10, 12, 27)

	Call read_amount_from_MAXIS(electricity, 10, 8, 27)
	Call read_amount_from_MAXIS(heat_air, 10, 9, 27)
	Call read_amount_from_MAXIS(phone, 10, 11, 27)
	electricity = electricity*1
	heat_air = heat_air*1
	phone = phone*1
	utilities_budgeted_amt = heat_air + electricity + phone
	utilities_budgeted_amt = utilities_budgeted_amt & ""
	Call read_amount_from_MAXIS(total_shelter_cost_budgeted_amt, 10, 14, 27)

	Call read_amount_from_MAXIS(snap_issued_amt, 10, 10, 71)
	Call read_amount_from_MAXIS(fed_benefit_amt, 10, 17, 71)
	Call read_amount_from_MAXIS(state_benefit_amt, 10, 18, 71)

	write_value_and_transmit "FSSM", 20, 70

	EMReadScreen budgeted_hh_size, 2, 13, 31
	budgeted_hh_size = trim(budgeted_hh_size)

	earned_income_correct_amt = earned_income_budgeted_amt
	unearned_correct_amt = unearned_budgeted_amt
	correct_hh_size = budgeted_hh_size
	medical_deduction_correct_amt = medical_deduction_budgeted_amt
	dependent_care_deduction_correct_amt = dependent_care_deduction_budgeted_amt
	child_support_deduction_correct_amt = child_support_deduction_budgeted_amt
	rent_mortgage_correct_amt = rent_mortgage_budgeted_amt
	tax_correct_amt = tax_budgeted_amt
	insurance_correct_amt = insurance_budgeted_amt
	other_cost_correct_amt = other_cost_budgeted_amt
	utilities_correct_amt = utilities_budgeted_amt

	Call budget_calculate_income(earned_income_correct_amt, unearned_correct_amt, earned_deduction_correct_amt, total_income_correct_amt, "STRING")
	Call budget_calculate_household(correct_hh_size, disa_household, cat_elig, standard_deduction_correct_amt, max_shelter_cost_correct_amt, max_gross_income_correct_amt, max_net_adj_income_correct_amt, max_snap_benefit, "STRING")
	Call budget_calculate_deductions(earned_deduction_correct_amt, medical_deduction_correct_amt, dependent_care_deduction_correct_amt, child_support_deduction_correct_amt, standard_deduction_correct_amt, total_deduction_correct_amt, total_income_correct_amt, net_income_correct_amt, fifty_perc_net_income_correct_amt, "STRING")
	Call budget_calculate_shelter_costs(rent_mortgage_correct_amt, tax_correct_amt, insurance_correct_amt, other_cost_correct_amt, utilities_correct_amt, total_shelter_cost_correct_amt, adj_shelter_cost_correct_amt, max_shelter_cost_correct_amt, counted_shelter_cost_correct_amt, fifty_perc_net_income_correct_amt, net_income_correct_amt, net_adj_income_correct_amt, "STRING")
	Call budget_calculate_benefit_details(cat_elig, total_income_correct_amt, net_adj_income_correct_amt, max_net_adj_income_correct_amt, max_gross_income_correct_amt, max_snap_benefit, monthly_snap_benefit_correct_amt, snap_issued_amt, overpayment_exists, supplement_exists, snap_overpayment_amt, snap_supplement_amt, "STRING")


	' 978321
	'
	' snap_issued_amt = 1316
	call back_to_self
End If

If MFIP_active = True Then
	Call navigate_to_MAXIS_screen("ELIG", "MFIP")

	call back_to_self
End If

' START A LOOP HERE
recalculation_confirmed = False
overpayment_exists = False
supplement_exists = False
calculation_needed = True
Do
	'Determine what happened with the review/mont process by dialog
	Do
		err_msg = ""
		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 316, 105, "02/22 Report Process Information"
		  DropListBox 180, 10, 60, 45, "Select One..."+chr(9)+"REVW"+chr(9)+"MONT", feb_process
		  DropListBox 260, 25, 50, 45, "Select One..."+chr(9)+"Yes"+chr(9)+"No", process_complete
		  DropListBox 65, 45, 90, 45, "Select One..."+chr(9)+"None Received"+chr(9)+"CAF"+chr(9)+"HRF"+chr(9)+"HUF"+chr(9)+"MNBenefits"+chr(9)+"Combined AR", form_received
		  EditBox 260, 45, 50, 15, form_received_date
		  DropListBox 65, 65, 90, 45, "Select One..."+chr(9)+"Not Required"+chr(9)+"Completed"+chr(9)+"Incomplete", interview_information
		  EditBox 260, 65, 50, 15, interview_date
		  DropListBox 65, 85, 60, 45, "Select One..."+chr(9)+"None Needed"+chr(9)+"Partial"+chr(9)+"Complete", verifs_received
		  ButtonGroup ButtonPressed
		    OkButton 205, 85, 50, 15
		    CancelButton 260, 85, 50, 15
		  Text 15, 15, 50, 10, "Case Number:"
		  Text 70, 15, 50, 10, MAXIS_case_number
		  Text 125, 15, 55, 10, "02/22 Process:"
		  Text 15, 30, 245, 10, "Was the MONT/REVW completed, with all required forms and verifications?"
		  Text 10, 50, 55, 10, "Form Received:"
		  Text 185, 50, 70, 10, "Date Form Recieved:"
		  Text 30, 70, 35, 10, "Interview:"
		  Text 205, 70, 50, 10, "Interview Date:"
		  Text 20, 90, 45, 10, "Verifications:"
		EndDialog

		dialog Dialog1
		cancel_confirmation

		If feb_process = "Select One..." Then err_msg = err_msg & vbCr & "* Select the process that was due for 02/22."
		If process_complete = "Select One..." Then err_msg = err_msg & vbCr & "* Indicate if the process was completed and case would have been able to be processedd and 'APP'd with the everything on file."
		If form_received = "Select One..." Then err_msg = err_msg & vbCr & "* Select which form was submitted or indicate that no form was received."
		If form_received <> "Select One..." and form_received <> "None Received" and IsDate(form_received_date) = False Then  err_msg = err_msg & vbCr & "* Since a form was received, enter a valid date for the date the form was received."
		If interview_information = "Select One..." Then err_msg = err_msg & vbCr & "* Indicate hwat happened with the interview process."
		If interview_information = "Completed" and IsDate(interview_date) = False Then  err_msg = err_msg & vbCr & "* Since the interview was completed, enter a valid date for the date the interview was completed."
		If verifs_received = "Select One..." Then err_msg = err_msg & vbCr & "* Indicate the status of the verifications for this case in the 02/22 report process."
		If process_complete = "Yes" and form_received = "None Received" Then err_msg = err_msg & vbCr & "* If the process is complete, The form received should not be 'None Received' - enter the form name."
		If process_complete = "Yes" and interview_information = "Incomplete" Then err_msg = err_msg & vbCr & "* If the process is complete, the interview should not be listed as 'Incomplete' - it should either be 'Not Required' or 'Completed'."
		If process_complete = "Yes" and verifs_received = "Partial" Then err_msg = err_msg & vbCr & "* If the process is complete, verifications received should not be 'Partial' - they should either be 'Complete' or 'None Needed'."

		If err_msg <> "" then MsgBox "Please resolve to continue:" & vbCr & err_msg
	Loop until err_msg = ""

	If process_complete = "No" Then calculation_needed = False

	'dialog for OP calculation
	If calculation_needed = True Then
		If MFIP_active = True Then
			Do
				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 556, 385, "Dialog"
				  'ISSUANCE
				  GroupBox 10, 5, 200, 35, "Benefit Issued for 02/22"
				  Text 15, 15, 30, 10, "MF-Cash:"
				  Text 50, 15, 30, 10, "$ XXXX"
				  Text 80, 15, 30, 10, "MF-Food:"
				  Text 110, 15, 30, 10, "$ XXXX"
				  Text 145, 15, 25, 10, "MF-HG:"
				  Text 170, 15, 30, 10, "$ XXXX"
				  Text 25, 25, 25, 10, "SNAP:"
				  Text 50, 25, 30, 10, "$ XXXX"
				  'Earned Income
				  'Unearned Income
				  'HH Comp
				  'Budget
				  GroupBox 340, 5, 210, 355, "Corrected Budget"
				  Text 375, 15, 55, 10, "Earned Income:"
				  Text 435, 15, 30, 10, "$ XXXX"
				  Text 355, 25, 75, 10, "Earned Inc Disregard:"
				  Text 445, 25, 30, 10, "- $ XXXX"
				  Text 365, 35, 65, 10, "Child Support Ded:"
				  Text 445, 35, 30, 10, "- $ XXXX"
				  Text 365, 45, 65, 10, "Net Earned Income:"
				  Text 445, 45, 30, 10, "$ XXXX"
				  Text 360, 60, 65, 10, "Family Wage Level:"
				  Text 435, 60, 30, 10, "$ XXXX"
				  Text 365, 70, 65, 10, "Net Earned Income:"
				  Text 445, 70, 30, 10, "- $ XXXX"
				  Text 390, 80, 40, 10, "Difference:"
				  Text 445, 80, 30, 10, "$ XXXX"
				  Text 355, 90, 75, 10, "Transitional Standard:"
				  Text 435, 90, 30, 10, "$ XXXX"
				  GroupBox 340, 95, 210, 20, ""
				  Text 355, 105, 120, 10, "Difference or Transitional Standard:"
				  Text 475, 105, 30, 10, "$ XXXX"
				  Text 370, 120, 60, 10, "Unearned Income:"
				  Text 370, 130, 65, 10, "Child Support Ded:"
				  Text 350, 140, 80, 10, "Child Support Exclusion:"
				  Text 375, 150, 55, 10, "Deemed Income:"
				  Text 350, 165, 125, 10, "Net Difference Transitional Standard:"
				  Text 425, 175, 50, 10, " Cash Portion:"
				  Text 425, 185, 45, 10, "Food Portion:"
				  Text 380, 200, 50, 10, "Subsidy/Tribal:"
				  Text 375, 210, 60, 10, "Net Cash Portion:"
				  Text 355, 220, 80, 10, "Tribal Counted Income:"
				  Text 375, 230, 60, 10, "Net Food Portion:"
				  Text 380, 240, 95, 10, "Total Cash and Food Portion:"
				  Text 375, 250, 60, 10, "Sanction Amount:"
				  Text 370, 265, 65, 10, "Correct MFIP Grant:"
				  Text 360, 275, 75, 10, "Correct Housing Grant:"
				  Text 400, 290, 75, 10, "MFIP Grant Received:"
				  Text 480, 290, 30, 10, "$ XXXX"
				  Text 425, 300, 50, 10, "HG Received:"
				  Text 455, 315, 50, 10, "Overpayment:"
				  Text 455, 325, 45, 10, "Cash Portion:"
				  Text 455, 335, 50, 10, " Food Portion:"
				  Text 465, 345, 40, 10, "HG Portion:"
				  ButtonGroup ButtonPressed
				    PushButton 385, 365, 165, 15, "MFIP Budget is Complete", mfip_claculation_done_btn

				EndDialog

				dialog Dialog1
				cancel_confirmation

				' If ButtonPressed = -1 Then ButtonPressed = calc_btn
				output_type = "STRING"
				If ButtonPressed = mfip_claculation_done_btn Then output_type = "NUMBER"


			Loop until ButtonPressed = mfip_claculation_done_btn
		End If

		If SNAP_active = True Then
			Do
				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 556, 385, "02/22 SNAP Incorrect Payment Calculation"

				  GroupBox 10, 5, 200, 35, "Benefit Issued for 02/22"
				  Text 15, 15, 30, 10, "MF-Cash:"
				  Text 50, 15, 30, 10, "$ " & mf_cash_issued_amt
				  Text 80, 15, 30, 10, "MF-Food:"
				  Text 110, 15, 30, 10, "$ " & mf_food_issued_amt
				  Text 145, 15, 25, 10, "MF-HG:"
				  Text 170, 15, 30, 10, "$ " & mf_hg_issued_amt
				  Text 25, 25, 25, 10, "SNAP:"
				  Text 50, 25, 30, 10, "$ " & snap_issued_amt

				  GroupBox 10, 40, 200, 75, "Income"
				  Text 90, 40, 35, 10, "Budgeted"
				  Text 140, 40, 35, 10, "Correct"
				  Text 30, 60, 55, 10, "Earned Income:"
				  Text 90, 60, 30, 10, "$ " & earned_income_budgeted_amt
				  EditBox 140, 55, 50, 15, earned_income_correct_amt
				  Text 20, 80, 65, 10, "Unearned Income:"
				  Text 90, 80, 30, 10, "$ " & unearned_budgeted_amt
				  EditBox 140, 75, 50, 15, unearned_correct_amt
				  Text 60, 100, 20, 10, "Total:"
				  Text 90, 100, 30, 10, "$ " & total_income_budgeted_amt
				  Text 140, 100, 30, 10, "$ " & total_income_correct_amt
				  ButtonGroup ButtonPressed
				    PushButton 175, 100, 30, 10, "CALC", calc_btn

				  GroupBox 10, 120, 200, 115, "Deductions"
				  Text 90, 120, 35, 10, "Budgeted"
				  Text 140, 120, 35, 10, "Correct"
				  Text 35, 140, 50, 10, "Standard Ded:"
				  Text 90, 140, 30, 10, "$ " & standard_deduction_budgeted_amt
				  ' EditBox 140, 135, 50, 15, earned_deduction_correct_amt
				  Text 140, 140, 30, 15, "$ " & standard_deduction_correct_amt

				  Text 15, 150, 70, 10, "Earned Income Ded:"
				  Text 90, 150, 30, 10, "$ " & earned_deduction_budgeted_amt
				  ' EditBox 150, 135, 50, 15, earned_deduction_correct_amt
				  Text 140, 150, 30, 15, "$ " & earned_deduction_correct_amt
				  Text 40, 165, 50, 10, "Medical Ded:"
				  Text 90, 165, 30, 10, "$ " & medical_deduction_budgeted_amt
				  EditBox 140, 160, 50, 15, medical_deduction_correct_amt
				  Text 15, 185, 70, 10, "Dependent Care Ded:"
				  Text 90, 185, 30, 10, "$ " & dependent_care_deduction_budgeted_amt
				  EditBox 140, 180, 50, 15, dependent_care_deduction_correct_amt
				  Text 37, 205, 50, 10, "Child Support:"
				  Text 90, 205, 30, 10, "$ " & child_support_deduction_budgeted_amt
				  EditBox 140, 200, 50, 15, child_support_deduction_correct_amt
				  Text 35, 220, 20, 10, "Total:"
				  Text 90, 220, 30, 10, "$ " & total_deduction_budgeted_amt
				  Text 140, 220, 30, 10, "$ " & total_deduction_correct_amt
				  ButtonGroup ButtonPressed
				    PushButton 175, 220, 30, 10, "CALC", calc_btn

				  'SHELTER '
				  GroupBox 10, 240, 200, 140, "Shelter Costs"
				  Text 90, 240, 35, 10, "Budgeted"
				  Text 140, 240, 35, 10, "Correct"
				  Text 30, 260, 55, 10, "Rent/Mortgage:"
				  Text 90, 260, 30, 10, "$ " & rent_mortgage_budgeted_amt
				  EditBox 140, 255, 50, 15, rent_mortgage_correct_amt
				  Text 35, 280, 45, 10, "Property Tax:"
				  Text 90, 280, 30, 10, "$ " & tax_budgeted_amt
				  EditBox 140, 275, 50, 15, tax_correct_amt
				  Text 25, 300, 60, 10, "Home Insurance:"
				  Text 90, 300, 30, 10, "$ " & insurance_budgeted_amt
				  EditBox 140, 295, 50, 15, insurance_correct_amt
				  Text 15, 320, 20, 10, "Other:"
				  EditBox 40, 315, 45, 15, other_cost_detail
				  Text 90, 320, 30, 10, "$ " & other_cost_budgeted_amt
				  EditBox 140, 315, 50, 15, other_cost_correct_amt
				  Text 55, 340, 30, 10, "Utilities:"
				  Text 90, 340, 30, 10, "$ " & utilities_budgeted_amt
				  DropListBox 140, 335, 50, 15, ""+chr(9)+"488"+chr(9)+"205"+chr(9)+"149"+chr(9)+"56"+chr(9)+"0", utilities_correct_amt
				  Text 60, 360, 20, 10, "Total:"
				  Text 90, 360, 30, 10, "$ " & total_shelter_cost_budgeted_amt
				  Text 140, 360, 30, 10, "$ " & total_shelter_cost_correct_amt
				  ButtonGroup ButtonPressed
				    PushButton 175, 360, 30, 10, "CALC", calc_btn

				  GroupBox 215, 5, 120, 115, "HH Composition"
				  Text 230, 20, 65, 10, "Budgeted HH Size:"
				  Text 300, 20, 15, 10, budgeted_hh_size
				  Text 240, 40, 55, 10, "Correct HH Size:"
				  EditBox 300, 35, 25, 15, correct_hh_size
				  ButtonGroup ButtonPressed
				    PushButton 295, 55, 30, 10, "CALC", calc_btn
				  Text 230, 70, 75, 10, "Standard Deduction:"
				  Text 260, 85, 35, 10, "Budgeted:"
				  Text 300, 85, 25, 10, "$ " & standard_deduction_budgeted_amt
				  Text 270, 100, 30, 10, "Correct:"
				  Text 300, 100, 25, 10, "$ " & standard_deduction_correct_amt
				  'BUTTON
				  GroupBox 340, 5, 210, 345, "Corrected Budget"
				  Text 360, 20, 55, 10, " Earned Income:"
				  Text 425, 20, 30, 10, "$ " & earned_income_correct_amt
				  Text 355, 30, 60, 10, "Unearned Income:"
				  Text 425, 30, 30, 10, "$ " & unearned_correct_amt
				  Text 385, 40, 50, 10, " Total Income:"
				  Text 440, 40, 30, 10, "$ " & total_income_correct_amt

				  Text 350, 55, 70, 10, " Earned Income Ded:"
				  Text 425, 55, 30, 10, "$ " & earned_deduction_correct_amt
				  Text 370, 65, 50, 10, " Standard Ded:"
				  Text 425, 65, 30, 10, "$ " & standard_deduction_correct_amt
				  Text 375, 75, 45, 10, "Medical Ded:"
				  Text 425, 75, 30, 10, "$ " & medical_deduction_correct_amt
				  Text 350, 85, 70, 10, "Dependent Care Ded:"
				  Text 425, 85, 30, 10, "$ " & dependent_care_deduction_correct_amt
				  Text 375, 95, 50, 10, "Child Support:"
				  Text 425, 95, 30, 10, "$ " & child_support_deduction_correct_amt
				  Text 375, 105, 60, 10, " Total Deductions:"
				  Text 440, 105, 30, 10, "$ " & total_deduction_correct_amt

				  Text 390, 120, 40, 10, "Net Income:"
				  Text 440, 120, 30, 10, "$ " & net_income_correct_amt

				  Text 370, 135, 50, 10, "Rent/Mortgage:"
				  Text 425, 135, 30, 10, "$ " & rent_mortgage_correct_amt
				  Text 375, 145, 45, 10, "Property Tax:"
				  Text 425, 145, 30, 10, "$ " & tax_correct_amt
				  Text 360, 155, 60, 10, " House Insurance:"
				  Text 425, 155, 30, 10, "$ " & insurance_correct_amt
				  Text 390, 165, 30, 10, " Utilities:"
				  Text 425, 165, 30, 10, "$ " & utilities_correct_amt
				  Text 355, 175, 70, 10, "Other (" & other_cost_detail & "):"
				  Text 425, 175, 30, 10, "$ " & other_cost_correct_amt
				  Text 365, 185, 70, 10, " Total Shelter Costs:"
				  Text 440, 185, 30, 10, "$ " & total_shelter_cost_correct_amt

				  Text 360, 200, 65, 10, "50% of Net Income:"
				  Text 425, 200, 30, 10, "$ " & fifty_perc_net_income_correct_amt
				  Text 345, 210, 80, 10, "Adjusted Shelter Costs:"
				  Text 425, 210, 30, 10, "$ " & adj_shelter_cost_correct_amt
				  Text 360, 220, 65, 10, " Max Allow Shelter:"
				  Text 425, 220, 30, 10, "$ " & max_shelter_cost_correct_amt
				  Text 345, 230, 90, 10, " Counted Shelter Expense:"
				  Text 440, 230, 30, 10, "$ " & counted_shelter_cost_correct_amt
				  Text 350, 245, 70, 10, "Net Adjusted Income:"
				  Text 425, 245, 30, 10, "$ " & net_adj_income_correct_amt
				  Text 365, 255, 55, 10, "Household Size:"
				  Text 425, 255, 30, 10, correct_hh_size
				  Text 350, 265, 70, 10, " Max Net Adj Income:"
				  Text 425, 265, 30, 10, "$ " & max_net_adj_income_correct_amt
				  Text 360, 275, 75, 10, "Monthly SNAP benefit:"
				  Text 440, 275, 30, 10, "$ " & monthly_snap_benefit_correct_amt
				  Text 360, 285, 75, 10, "Sanction/Recoupment:"
				  Text 440, 285, 30, 10, "$ " & sanction_rcoupment_correct_amt
				  Text 405, 305, 100, 10, "Correct SNAP Benefit Amount:"
				  Text 510, 305, 30, 10, "$ " & snap_correct_amt
				  Text 425, 320, 80, 10, "Benefit amount issued:"
				  Text 510, 320, 30, 10, "$ " & snap_issued_amt
				  If overpayment_exists = True Then
					  Text 455, 335, 50, 10, "Overpayment:"
					  Text 510, 335, 30, 10, "$ " & snap_overpayment_amt
				  End If
				  If supplement_exists = True Then
					  Text 455, 335, 50, 10, "Supplement:"
					  Text 510, 335, 30, 10, "$ " & snap_supplement_amt
				  End If
				  If overpayment_exists = False And supplement_exists = False Then
				  	  Text 400, 335, 100, 10, "02/22 Issuance was Correct"
				  End If
				  ButtonGroup ButtonPressed
				    PushButton 480, 10, 65, 12, "CALCULATE", calc_btn
					PushButton 385, 365, 165, 15, "SNAP Budget is Complete", snap_claculation_done_btn

				EndDialog

				dialog Dialog1
				cancel_confirmation

				' If ButtonPressed = -1 Then ButtonPressed = calc_btn
				output_type = "STRING"
				If ButtonPressed = snap_claculation_done_btn Then output_type = "NUMBER"

				Call budget_calculate_income(earned_income_correct_amt, unearned_correct_amt, earned_deduction_correct_amt, total_income_correct_amt, output_type)
				Call budget_calculate_household(correct_hh_size, disa_household, cat_elig, standard_deduction_correct_amt, max_shelter_cost_correct_amt, max_gross_income_correct_amt, max_net_adj_income_correct_amt, max_snap_benefit, output_type)
				Call budget_calculate_deductions(earned_deduction_correct_amt, medical_deduction_correct_amt, dependent_care_deduction_correct_amt, child_support_deduction_correct_amt, standard_deduction_correct_amt, total_deduction_correct_amt, total_income_correct_amt, net_income_correct_amt, fifty_perc_net_income_correct_amt, output_type)
				Call budget_calculate_shelter_costs(rent_mortgage_correct_amt, tax_correct_amt, insurance_correct_amt, other_cost_correct_amt, utilities_correct_amt, total_shelter_cost_correct_amt, adj_shelter_cost_correct_amt, max_shelter_cost_correct_amt, counted_shelter_cost_correct_amt, fifty_perc_net_income_correct_amt, net_income_correct_amt, net_adj_income_correct_amt, output_type)
				Call budget_calculate_benefit_details(cat_elig, total_income_correct_amt, net_adj_income_correct_amt, max_net_adj_income_correct_amt, max_gross_income_correct_amt, max_snap_benefit, monthly_snap_benefit_correct_amt, snap_issued_amt, overpayment_exists, supplement_exists, snap_overpayment_amt, snap_supplement_amt, output_type)


			Loop until ButtonPressed = snap_claculation_done_btn

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 166, 100, "TEST DIALOG"
			  EditBox 90, 10, 70, 15, MAXIS_case_number
			  ButtonGroup ButtonPressed
			    OkButton 55, 80, 50, 15
			    CancelButton 110, 80, 50, 15
			  Text 10, 15, 80, 10, "Enter the Case Number:"
			  Text 10, 35, 150, 45, "This script is specific to the detailed review of the cases impacted by the Autoclose Pause that happened in 02/22 and does not take any MAXIS action or create CASE/NOTEs as this process is handled external from MAXIS."
			EndDialog

			dialog Dialog1
		End If
	End If

	MsgBox "DONE"

	'dialog with calculation and ready for confirmation

	' LOOP UNTIL THIS IS CONFIRMED
Loop until recalculation_confirmed = True

'Create PDF and save

'Open Excel and add information to the excel
